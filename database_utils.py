# database_utils.py (最终版)

import pymysql.cursors
from db_config import DB_CONFIG, TABLE_NAME

def create_connection():
    """ 创建并返回一个数据库连接 (使用 PyMySQL) """
    connection = None
    try:
        connection = pymysql.connect(**DB_CONFIG,
                                     cursorclass=pymysql.cursors.DictCursor)
        print("成功连接到MySQL数据库 (使用 PyMySQL)。")
        return connection
    except pymysql.MySQLError as e:
        print(f"连接MySQL时发生错误: {e}")
        return None

def create_table_if_not_exists(connection):
    """ 在数据库中创建日志表 (包含12个数据列) """
    with connection.cursor() as cursor:
        create_table_query = f"""
        CREATE TABLE IF NOT EXISTS `{TABLE_NAME}` (
            `id` INT AUTO_INCREMENT PRIMARY KEY,
            `printer_source`  VARCHAR(255),
            `log_date`        DATE,
            `end_time`        TIME,
            `user`            VARCHAR(255),
            `output_location` VARCHAR(255),
            `job_info`        TEXT,
            `page_info`       VARCHAR(255),
            `pages`           INT,
            `sheets`          INT,
            `job_status`      VARCHAR(255),
            `job_level`       VARCHAR(255),
            `job_id`          VARCHAR(255),
            UNIQUE KEY `unique_job_entry` (`printer_source`, `log_date`, `end_time`, `job_id`)
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
        """
        try:
            cursor.execute(create_table_query)
            print(f"数据表 '{TABLE_NAME}' 已成功创建或已存在。")
        except pymysql.MySQLError as e:
            print(f"创建数据表时发生错误: {e}")

def insert_log_data(connection, data_list):
    """ 将日志数据列表批量插入到数据库中 (需要12个元素的元组) """
    if not data_list:
        return 0
    
    with connection.cursor() as cursor:
        # 这个SQL语句需要12个 %s 占位符
        sql = f"""
        INSERT INTO `{TABLE_NAME}` 
        (printer_source, log_date, end_time, user, output_location, job_info, page_info, pages, sheets, job_status, job_level, job_id) 
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        ON DUPLICATE KEY UPDATE sheets=VALUES(sheets);
        """
        try:
            rowcount = cursor.executemany(sql, data_list)
            connection.commit()
            print(f"成功向数据库同步 {rowcount} 条记录。")
            return rowcount
        except pymysql.MySQLError as e:
            print(f"插入数据时发生错误: {e}")
            connection.rollback()
            return 0