import os
import pyodbc
import pandas as pd
from datetime import datetime
from sqlalchemy import create_engine
import warnings
from sqlalchemy.exc import SAWarning

# 屏蔽SQLAlchemy版本警告
warnings.filterwarnings("ignore", category=SAWarning)

# ========== 自动适配本机所有ODBC驱动（终极兼容，不再锁17） ==========
def get_best_sql_driver():
    for d in pyodbc.drivers():
        if "SQL Server" in d:
            return d
    raise Exception("未检测到任何SQL Server ODBC驱动，请安装驱动！")

BEST_DRIVER = get_best_sql_driver()

# 固定输出文件夹：桌面的 db_sync_output
OUTPUT_FOLDER = os.path.join(os.path.expanduser("~"), "Desktop", "db_sync_output")
if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)


class DataSyncChecker:
    def __init__(self, source_conn, target_conn):
        self.source_conn = source_conn
        self.target_conn = target_conn

    def get_table_data(self, conn_str, table_name, select_fields, primary_key, query_timeout=30):
        col_str = ", ".join([f"[{c}]" for c in select_fields])
        sql = f"SELECT {col_str} FROM [{table_name}]"

        engine = create_engine(
            f"mssql+pyodbc:///?odbc_connect={conn_str}",
            connect_args={"timeout": query_timeout}
        )
        with engine.connect() as conn:
            df = pd.read_sql(sql, conn)

        print(f"【{table_name}】读取字段列表：{df.columns.tolist()}")
        if df.empty:
            print(f"【警告】表[{table_name}]查询结果为空！")

        # 前置校验：主键字段必须存在
        if primary_key not in df.columns:
            raise ValueError(f"校验失败：查询结果缺少主键字段「{primary_key}」，检查select_fields配置！")

        # 重复主键检测
        dup_count = df.duplicated(subset=[primary_key]).sum()
        if dup_count > 0:
            raise ValueError(f"【严重错误】表[{table_name}]存在 {dup_count} 条重复主键，无法正常比对，请先清理脏数据！")

        # 核心优化：保留原始Id列 + 设置索引提速
        df["tmp_index"] = df[primary_key]
        df = df.set_index("tmp_index")

        return df

    @staticmethod
    def format_sql_value(val):
        if pd.isna(val):
            return "NULL"
        if isinstance(val, (int, float)):
            return str(val)
        if isinstance(val, datetime):
            return f"'{val.strftime('%Y-%m-%d %H:%M:%S')}'"
        # 单引号转义，防止SQL语法错误
        s = str(val).replace("'", "''")
        return f"'{s}'"

    def compare_table(self, table_name, primary_key, select_fields, compare_fields, insert_primary_key=True):
        print(f"\n========== 开始比对数据表：{table_name} ==========")
        try:
            source_df = self.get_table_data(self.source_conn, table_name, select_fields, primary_key)
            target_df = self.get_table_data(self.target_conn, table_name, select_fields, primary_key)
        except Exception as e:
            print(f"读取表异常:{e}")
            import traceback
            traceback.print_exc()
            return

        # 利用索引快速获取主键集合（高性能、哈希匹配）
        source_keys = set(source_df.index)
        target_keys = set(target_df.index)

        need_add_keys = source_keys - target_keys
        need_del_keys = target_keys - source_keys
        common_keys = source_keys & target_keys

        # 字段差异比对
        diff_list = []
        for key in common_keys:
            src_row = source_df.loc[key]
            tgt_row = target_df.loc[key]
            for col in compare_fields:
                src_val = src_row[col]
                tgt_val = tgt_row[col]
                if pd.isna(src_val) and pd.isna(tgt_val):
                    continue
                if str(src_val).strip() == str(tgt_val).strip():
                    continue
                diff_list.append({
                    "主键": key,
                    "字段": col,
                    "源库值": src_val,
                    "目标库值": tgt_val
                })

        # ========== 生成SQL脚本 ==========
        sql_scripts = []
        sql_scripts.append("BEGIN TRANSACTION; -- 执行完确认无误，取消下面COMMIT注释")
        sql_scripts.append("")

        # 新增语句（完美支持主键插入）
        for key in need_add_keys:
            row = source_df.loc[key]
            if insert_primary_key:
                insert_cols = [primary_key] + compare_fields
            else:
                insert_cols = compare_fields
            cols = ", ".join([f"[{c}]" for c in insert_cols])
            vals = ", ".join([self.format_sql_value(row[c]) for c in insert_cols])
            sql_scripts.append(f"INSERT INTO [{table_name}] ({cols}) VALUES ({vals});")

        sql_scripts.append("")
        # 删除语句
        for key in need_del_keys:
            sql_scripts.append(f"DELETE FROM [{table_name}] WHERE [{primary_key}] = {self.format_sql_value(key)};")

        sql_scripts.append("")
        # 更新语句合并
        update_map = {}
        for item in diff_list:
            pk = item["主键"]
            col = item["字段"]
            val = self.format_sql_value(item["源库值"])
            if pk not in update_map:
                update_map[pk] = []
            update_map[pk].append(f"[{col}] = {val}")
        for pk, set_expr in update_map.items():
            sql_scripts.append(f"UPDATE [{table_name}] SET {', '.join(set_expr)} WHERE [{primary_key}] = {self.format_sql_value(pk)};")

        sql_scripts.append("")
        sql_scripts.append("--COMMIT;")
        sql_scripts.append("--ROLLBACK;")

        # ========== 修复：时间不带冒号，彻底解决Windows报错 ==========
        now = datetime.now().strftime("%Y%m%d_%H_%M_%S")
        excel_name = os.path.join(OUTPUT_FOLDER, f"完整数据差异报告_{table_name}_{now}.xlsx")
        sql_name = os.path.join(OUTPUT_FOLDER, f"同步脚本_{table_name}_{now}.sql")

        try:
            with pd.ExcelWriter(excel_name, engine="openpyxl") as writer:
                # 1.字段差异明细
                pd.DataFrame(diff_list).to_excel(writer, sheet_name="字段差异明细", index=False)

                # 2.待新增完整数据
                add_full_data = source_df.loc[list(need_add_keys)].reset_index(drop=True)
                add_full_data.to_excel(writer, sheet_name="待新增完整数据", index=False)

                # 3.待删除完整数据
                del_full_data = target_df.loc[list(need_del_keys)].reset_index(drop=True)
                del_full_data.to_excel(writer, sheet_name="待删除完整数据", index=False)

                # 4.比对汇总统计
                summary_data = [{
                    "待新增数据条数": len(need_add_keys),
                    "待删除数据条数": len(need_del_keys),
                    "字段差异总数": len(diff_list),
                    "比对时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }]
                pd.DataFrame(summary_data).to_excel(writer, sheet_name="比对汇总统计", index=False)

            with open(sql_name, "w", encoding="utf-8") as f:
                f.write("\n".join(sql_scripts))

        except Exception as e:
            print(f"写入文件失败：{e}")

        print(f"\n✅========== 比对完成（V3.3 高性能最终版） ==========✅")
        print(f"📌 待新增数据：{len(need_add_keys)} 条")
        print(f"📌 待删除数据：{len(need_del_keys)} 条")
        print(f"📌 字段不一致：{len(diff_list)} 处")
        print(f"📁 完整Excel报告：{excel_name}")
        print(f"📁 同步SQL脚本：{sql_name}\n")


if __name__ == "__main__":
    # 动态驱动，全机器兼容
    source_db = (
        f"DRIVER={{{BEST_DRIVER}}};"
        "SERVER=.\\SQLEXPRESS;"
        "DATABASE=ClubDB_Dev;"
        "Trusted_Connection=yes;"
    )
    target_db = (
        f"DRIVER={{{BEST_DRIVER}}};"
        "SERVER=.\\SQLEXPRESS;"
        "DATABASE=ClubDB_Test;"
        "Trusted_Connection=yes;"
    )

    checker = DataSyncChecker(source_db, target_db)

    select_fields = ["Id", "UserName", "RealName", "Phone", "Email", "CreateTime"]
    compare_fields = ["UserName", "RealName", "Phone", "Email", "CreateTime"]

    checker.compare_table(table_name="User",
                          primary_key="Id",
                          select_fields=select_fields,
                          compare_fields=compare_fields,
                          insert_primary_key=True)
