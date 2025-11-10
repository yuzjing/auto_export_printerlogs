import time
import os
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from datetime import date, datetime
import calendar
from dateutil.relativedelta import relativedelta
import pandas as pd
import glob 

# from data_processor import process_and_insert_data # 您的数据处理模块

# --- [新增] 导入 webdriver-manager 和 pyvirtualdisplay 相关的模块 ---
from pyvirtualdisplay import Display
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager

# --- [新增] 初始化 display 和 driver 变量，以便在 finally 中使用 ---
display = None
driver = None
download_successful = False

# --- [新增] 使用一个全局的 try...finally 结构来确保所有外部资源都能被正确关闭 ---
try:
    # --- [新增] 启动虚拟显示器 ---
    print("正在启动虚拟显示器...")
    display = Display(visible=0, size=(1920, 1080))
    display.start()
    print("虚拟显示器已启动。")

    # --- 您的原始代码 (完全不变) ---
    today = date.today()
    previous_month_date = today - relativedelta(months=1)
    EXPORT_YEAR = previous_month_date.year
    EXPORT_MONTH = previous_month_date.month

    RETRY_COUNT = 5 # 设置每天的最大重试次数
    RETRY_DELAY = 10 # 设置每次重试的间隔时间（秒）

    # --- 1. 配置信息 (完全不变) ---
    USERNAME = '11111'
    PASSWORD = 'x-admin'
    IP_ADDRESS = '192.168.12.250'
    AUTH_LOGIN_URL = f"http://{USERNAME}:{PASSWORD}@{IP_ADDRESS}"
    CHROME_BINARY_PATH = "/usr/bin/google-chrome-unstable"
    HOME_DIR = os.path.expanduser('~')
    ARCHIVE_DIR = os.path.join(HOME_DIR, 'PrinterReportsArchive/floor2')
    FINAL_REPORT_DIR = os.path.join(HOME_DIR, 'PrinterReportsFinal/floor2')

    run_date_str = today.strftime('%Y-%m-%d')
    run_archive_dir = os.path.join(ARCHIVE_DIR, f"Run_{run_date_str}")
    DOWNLOAD_DIR = run_archive_dir

    os.makedirs(run_archive_dir, exist_ok=True)
    os.makedirs(FINAL_REPORT_DIR, exist_ok=True)

    if os.path.exists(DOWNLOAD_DIR):
        for filename in os.listdir(DOWNLOAD_DIR):
            file_path = os.path.join(DOWNLOAD_DIR, filename)
            try:
                if os.path.isfile(file_path):
                    os.remove(file_path)
                    print(f"已删除旧文件: {filename}")
            except Exception as e:
                print(f"删除文件 {filename} 失败: {e}")
    else:
        print(f"目录 {DOWNLOAD_DIR} 不存在，无需清理")

    # --- 2. 初始化浏览器 ---
    options = Options()
    options.binary_location = CHROME_BINARY_PATH
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--ignore-ssl-errors')
    # 【新增】为后台环境优化，强制指定窗口大小和语言，确保环境一致性
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--lang=zh-CN")
    #options.add_argument('--headless')
    prefs = {'download.default_directory': DOWNLOAD_DIR}
    options.add_experimental_option('prefs', prefs)

    if not os.path.exists(ARCHIVE_DIR):
        os.makedirs(ARCHIVE_DIR)

    # --- [核心修改] 浏览器初始化方式 ---
    print("正在自动配置并启动浏览器...")
    service = ChromeService(executable_path=ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    print("浏览器启动成功！")

    wait = WebDriverWait(driver, 10)

    # --- 步骤 1: 登录 ---
    print("正在使用认证方式访问...")
    driver.get(AUTH_LOGIN_URL)
    print("已发送登录请求。")

    # --- 步骤 2: 处理 JavaScript 警告框 ---
    print("等待并处理安全警告框...")
    try:
        alert1 = wait.until(EC.alert_is_present())
        alert1.accept()
        print("第一个警告框已处理。")
        alert2 = wait.until(EC.alert_is_present())
        alert2.accept()
        print("第二个警告框已处理。")
    except:
        print("在10秒内未检测到警告框，将继续执行。")
    
    print("成功进入主页面！")
    try:
        print("正在等待并切换到名为 'TF' 的框架...")
        wait.until(EC.frame_to_be_available_and_switch_to_it("TF"))
        print("成功切换到 'TF' 框架！")
    except:
        print("错误：找不到名为 'TF' 的框架。请检查框架的 name 或 id。")
        raise 

    print("正在点击 'Properties' 链接...")
    properties_link_locator = "a[href='prop.htm']"
    properties_link = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, properties_link_locator)))
    properties_link.click()
    print("已成功点击 'Properties'。")
    
    print("正在从 'TF' 框架切换回顶层...")
    driver.switch_to.default_content()
    print("已切换回顶层。WebDriver 正在重新评估页面结构...")
    
    print("正在智能等待并切换到新出现的 'NF' 框架...")
    wait.until(EC.frame_to_be_available_and_switch_to_it("NF"))
    print("成功切换到 'NF' 框架！")

    print("正在点击 'General Setup' 链接...")
    general_setup_locator = "//a[contains(@href, \"flip('nvTree',2)\")]"
    general_setup_link = wait.until(EC.element_to_be_clickable((By.XPATH, general_setup_locator)))
    general_setup_link.click()
    print("已成功点击 'General Setup'。")

    print("正在点击 'Job Management' 链接...")
    job_management_locator = "//a[contains(@href, \"flip('genItms',1)\")]"
    job_management_link = wait.until(EC.element_to_be_clickable((By.XPATH, job_management_locator)))
    job_management_link.click()
    print("已成功点击 'Job Management'。")

    print("正在点击 'Export Job History' 链接...")
    export_history_locator = "//a[contains(@href, \"dclick('jobItms',1)\")]"
    export_history_link = wait.until(EC.element_to_be_clickable((By.XPATH, export_history_locator)))
    export_history_link.click()
    print("已成功点击 'Export Job History'，进入最终页面！")
    print("正在从 'NF' 框架切换回顶层...")
    driver.switch_to.default_content()

    print("正在填写日期...")
    dateform_frame_name = "RF"
    print("正在从 'NF' 框架切换到 'RF' 框架...")
    wait.until(EC.frame_to_be_available_and_switch_to_it(dateform_frame_name))
    print("已切换到 'RF' 框架。")

    print("正在填写开始日期...")

    _, num_days = calendar.monthrange(EXPORT_YEAR, EXPORT_MONTH)
    
    for day in range(1, num_days + 1):
        day_str = f"{day:02d}"
        month_str = f"{EXPORT_MONTH:02d}"
        current_date_str = f"{EXPORT_YEAR}-{month_str}-{day_str}"
        
        for attempt in range(RETRY_COUNT):
            try:
                print(f"\n--- [尝试 {attempt + 1}/{RETRY_COUNT}] 开始处理日期: {current_date_str} ---")

                start_year_input = wait.until(EC.presence_of_element_located((By.NAME, 'YEAR1')))
                start_year_input.clear()
                start_year_input.send_keys(str(EXPORT_YEAR))
                start_month_input = driver.find_element(By.NAME, 'MONTH1')
                start_month_input.clear()
                start_month_input.send_keys(str(EXPORT_MONTH))
                start_day_input = driver.find_element(By.NAME, 'DAY1')
                start_day_input.clear()
                start_day_input.send_keys(day_str)
                start_hour_input = driver.find_element(By.NAME, 'HOUR1')
                start_hour_input.clear()
                start_hour_input.send_keys("00")
                start_miniute_input = driver.find_element(By.NAME, 'MIN1')
                start_miniute_input.clear()
                start_miniute_input.send_keys("00")
                
                end_year_input = driver.find_element(By.NAME, 'YEAR2')
                end_year_input.clear()
                end_year_input.send_keys(str(EXPORT_YEAR))
                end_month_input = driver.find_element(By.NAME, 'MONTH2')
                end_month_input.clear()
                end_month_input.send_keys(str(EXPORT_MONTH))
                end_day_input = driver.find_element(By.NAME, 'DAY2')
                end_day_input.clear()
                end_day_input.send_keys(day_str)
                end_hour_input = driver.find_element(By.NAME, 'HOUR2')
                end_hour_input.clear()
                end_hour_input.send_keys("23")
                end_miniute_input = driver.find_element(By.NAME, 'MIN2')
                end_miniute_input.clear()
                end_miniute_input.send_keys("59")
                
                print("日期已填写。")

                export_link_locator = "//a[contains(text(), 'Export file in')]"
                wait.until(EC.element_to_be_clickable((By.XPATH, export_link_locator))).click()
                print("已点击导出链接。")

                time.sleep(3)
                error_locator = "//*[contains(text(), 'Unable to export') or contains(text(), 'There is no job') or contains(text(), 'Your request was not')]"
                try:
                    error_element = driver.find_element(By.XPATH, error_locator)
                    error_text = error_element.text
                    
                    is_final_feedback = False
                    
                    if "Unable to export" in error_text:
                        print(f"!! [最终反馈] 日期 {current_date_str} 数据过多。")
                        is_final_feedback = True
                    elif "no job" in error_text:
                        print(f"   [最终反馈] 日期 {current_date_str} 当日无数据。")
                        is_final_feedback = True
                    elif "Your request was not" in error_text:
                        print(f"!! [瞬时错误] 打印机正忙 (尝试 {attempt + 1}/{RETRY_COUNT})。")
                    
                    driver.find_element(By.XPATH, "//input[@value='Back']").click()
                    
                    if is_final_feedback:
                        break
                    else:
                        raise Exception("Printer is busy or unknown error, will retry...")
                except NoSuchElementException:
                    print(f"[成功] 日期 {current_date_str} 导出成功，开始处理文件...")
                    default_filename = "jobhist.csv"
                    original_filepath = os.path.join(DOWNLOAD_DIR, default_filename)
                    download_wait_time = 120
                    time_elapsed = 0
                    while not os.path.exists(original_filepath) and time_elapsed < download_wait_time:
                        time.sleep(1)
                        time_elapsed += 1
                        print(f"\r正在等待下载... {time_elapsed}s / {download_wait_time}s", end="")
                    print()

                    if os.path.exists(original_filepath):
                        print(f"文件 {default_filename} 已成功下载！")
                        time.sleep(2)
                        new_filename = f"jobhist_{current_date_str}.csv"
                        new_filepath = os.path.join(run_archive_dir, new_filename)
                        try:
                            shutil.move(original_filepath, new_filepath)
                            print(f"文件处理完毕: {new_filepath}")
                        except Exception as move_error:
                            print(f"!! 严重错误: 文件移动失败! Error: {move_error}")
                    else:
                        print(f"!! [下载失败] 等待 {download_wait_time} 秒后，未找到下载文件。可能当日无数据。")
                    break
            except Exception as day_error:
                print(f"!! [错误] 处理 {current_date_str} 时发生错误 (尝试 {attempt + 1}/{RETRY_COUNT})")
                print(f"   错误详情: {day_error}")
                if attempt == RETRY_COUNT - 1:
                    print(f"!!!!!! [最终失败] 日期 {current_date_str} 在重试 {RETRY_COUNT} 次后彻底失败，将跳过此天。 !!!!!!")
                else:
                    print(f"   将在 {RETRY_DELAY} 秒后重试...")
                    time.sleep(RETRY_DELAY)
        else:
            pass
            
    print(f"\n[任务完成] {EXPORT_YEAR}年{EXPORT_MONTH}月的所有数据已处理完毕！")
    download_successful = True
    
except Exception as e:
    print(f"!!!!!! [阶段一 失败] 在浏览器自动化过程中发生严重错误: {e} !!!!!!")
    download_successful = False

# --- [修改] finally 块负责关闭所有外部资源 ---
finally:
    if driver:
        driver.quit()
        print("浏览器已关闭。")
    if display:
        display.stop()
        print("虚拟显示器已关闭。")

# --- [阶段二] (保持不变) ---
if download_successful:
    print("\n--- [阶段二 开始] 处理已下载文件并写入数据库 ---")
    try:
        # process_and_insert_data(directory_path=run_archive_dir, printer_type='floor2')
        print("模拟调用 process_and_insert_data 函数。")
    except Exception as e:
        print(f"!!!!!! [阶段二 失败] 在数据处理或数据库插入过程中发生错误: {e} !!!!!!")
else:
    print("\n[任务中止] 由于阶段一（文件下载）未成功，阶段二（数据处理）被跳过。")

print("\n--- [任务总流程结束] ---")