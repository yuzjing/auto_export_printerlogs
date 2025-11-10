import time
import os
import shutil
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from datetime import date
import calendar
from dateutil.relativedelta import relativedelta
import pandas as pd
import glob 
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# --- [新增] 导入 webdriver-manager 和 pyvirtualdisplay 相关的模块 ---
from pyvirtualdisplay import Display
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager

from data_processor import process_and_insert_data # 您的数据处理模块

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

    RETRY_COUNT = 5
    RETRY_DELAY = 10

    # --- 1. 配置信息 (完全不变) ---
    USERNAME = '11111'
    PASSWORD = 'x-admin'
    IP_ADDRESS = '192.168.14.250'
    AUTH_LOGIN_URL = f"http://{IP_ADDRESS}"
    CHROME_BINARY_PATH = "/usr/bin/google-chrome-unstable"
    HOME_DIR = os.path.expanduser('~')
    ARCHIVE_DIR = os.path.join(HOME_DIR, 'PrinterReportsArchive/floor4')
    FINAL_REPORT_DIR = os.path.join(HOME_DIR, 'PrinterReportsFinal/floor4')

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

    # --- 2. 初始化浏览器 ---
    options = Options()
    options.binary_location = CHROME_BINARY_PATH
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--ignore-ssl-errors')
    options.add_argument('--headless')
    prefs = {'download.default_directory': DOWNLOAD_DIR}
    options.add_experimental_option('prefs', prefs)

    if not os.path.exists(ARCHIVE_DIR):
        os.makedirs(ARCHIVE_DIR)

    # --- 浏览器初始化方式 ---
    print("正在自动配置并启动浏览器...")
    service = ChromeService(executable_path=ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    print("浏览器启动成功！")

    wait = WebDriverWait(driver, 10)

    # --- 步骤 1: 登录 (您的原始逻辑) ---
    print("正在打开页面...")
    driver.get(AUTH_LOGIN_URL)
    print("已发送页面请求。")

    print("等待页面加载完毕...")
    time.sleep(3)

    login_in_locator = "a[tabindex='0']"
    login_in_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, login_in_locator)))
    login_in_button.click()

    print("正在填写用户名密码...")
    username_input = wait.until(EC.presence_of_element_located((By.ID, 'loginName')))
    username_input.send_keys(USERNAME)
    time.sleep(1)
    password_input = wait.until(EC.presence_of_element_located((By.ID, 'loginPsw')))
    password_input.send_keys(PASSWORD)
    login_in_button = wait.until(EC.element_to_be_clickable((By.ID, 'loginButton')))
    login_in_button.click()

    try:
        print("正在处理警告框...")
        secure_button1 = wait.until(EC.element_to_be_clickable((By.ID, 'securityAlertConfirmKoDefault')))
        secure_button1.click()
        secure_button2 = wait.until(EC.element_to_be_clickable((By.ID, 'securityAlertConfirmSnmpDefault')))
        secure_button2.click()
        print("已处理警告框。")
    except Exception as e:
        print(f"处理警告框时出错或未找到: {e}")

    print("登录成功，正在等待页面加载蒙版消失...")
    loading_overlay_locator = (By.XPATH, "//div[contains(@class, 'xux-modal')]")
    long_wait = WebDriverWait(driver, 30)
    long_wait.until(EC.invisibility_of_element_located(loading_overlay_locator))
    print("加载蒙版已消失！页面已准备好交互。")
    
    hamberger_button_locator = "//button[@title='Toggle Navigation']"
    hamberger_button = wait.until(EC.element_to_be_clickable((By.XPATH, hamberger_button_locator)))
    hamberger_button.click()
    time.sleep(1)

    print("正在查找并点击 'Jobs' 链接...")
    jobs_link_locator = (By.XPATH, "//li[@id='globalnavJobs']//a")
    jobs_link_element = wait.until(EC.presence_of_element_located(jobs_link_locator))
    jobs_link_element.click()
    print("已成功点击 'Jobs' 链接！")

    # --- 按天循环导出 (您的原始逻辑) ---
    _, num_days = calendar.monthrange(EXPORT_YEAR, EXPORT_MONTH)
    
    for day in range(1, num_days + 1):
        day_str = f"{day:02d}"
        month_str = f"{EXPORT_MONTH:02d}"
        current_date_str = f"{EXPORT_YEAR}-{month_str}-{day_str}"
        
        for attempt in range(RETRY_COUNT):
            try:
                print(f"\n--- [尝试 {attempt + 1}/{RETRY_COUNT}] 开始处理日期: {current_date_str} ---")
                try:
                    cancel_button_locator = (By.XPATH, "//button/span[text()='Cancel']")
                    cancel_button = driver.find_element(*cancel_button_locator)
                    if cancel_button.is_displayed():
                        print("检测到上一次的模态框还开着，正在关闭...")
                        cancel_button.click()
                        wait.until(EC.invisibility_of_element_located(cancel_button_locator))
                        print("模态框已关闭。")
                except:
                    print("页面状态正常，模态框已关闭。")
                
                print("正在打开导出设置模态框...")
                export_history_button_locator = (By.ID, "openJobHistoryExportModalWindow")
                wait.until(EC.element_to_be_clickable(export_history_button_locator)).click()
                
                start_area_locator = (By.ID, "startDateTimeSection")
                wait.until(EC.presence_of_element_located(start_area_locator))
                print("模态框已成功打开并加载。")

                start_area = driver.find_element(By.ID, "startDateTimeSection")
                start_date_input = start_area.find_element(By.CSS_SELECTOR, "input.hasDatepicker")
                start_date_string = f"{EXPORT_YEAR}/{EXPORT_MONTH:02d}/{day_str}"
                start_date_input.clear(); start_date_input.send_keys(start_date_string); time.sleep(0.3)
                # start_hour_visible_input = start_area.find_element(By.XPATH, ".//input[@role='spinbutton' and @aria-label='Hour']")
                # start_minute_visible_input = start_area.find_element(By.XPATH, ".//input[@role='spinbutton' and @aria-label='Minute']")
                # ActionChains(driver).click(start_hour_visible_input).pause(0.2).send_keys(Keys.CONTROL + "a").send_keys("00").perform()
                # ActionChains(driver).click(start_minute_visible_input).pause(0.2).send_keys(Keys.CONTROL + "a").send_keys("00").perform()

                end_area = driver.find_element(By.ID, "endDateTimeSection")
                end_date_input = end_area.find_element(By.CSS_SELECTOR, "input.hasDatepicker")
                end_date_string = f"{EXPORT_YEAR}/{EXPORT_MONTH:02d}/{day_str}"
                end_date_input.clear(); end_date_input.send_keys(end_date_string); time.sleep(0.3)
                # end_hour_visible_input = end_area.find_element(By.XPATH, ".//input[@role='spinbutton' and @aria-label='Hour']")
                # end_minute_visible_input = end_area.find_element(By.XPATH, ".//input[@role='spinbutton' and @aria-label='Minute']")
                # ActionChains(driver).click(end_hour_visible_input).pause(0.2).send_keys(Keys.CONTROL + "a").send_keys("23").perform()
                # ActionChains(driver).click(end_minute_visible_input).pause(0.2).send_keys(Keys.CONTROL + "a").send_keys("59").perform()
                # print("所有日期已填写。")

                # --- 定位小时和分钟输入框 ---
                
                start_hour_input = start_area.find_element(By.XPATH, ".//input[@role='spinbutton' and @aria-label='Hour']")
                start_minute_input = start_area.find_element(By.XPATH, ".//input[@role='spinbutton' and @aria-label='Minute']")
                end_hour_input = end_area.find_element(By.XPATH, ".//input[@role='spinbutton' and @aria-label='Hour']")
                end_minute_input = end_area.find_element(By.XPATH, ".//input[@role='spinbutton' and @aria-label='Minute']")

                # --- 使用 JavaScript 清空并用 send_keys 填写的“双保险”模式 ---
                print("正在清空并填写开始时间...")
                # 步骤1: 用 JS 强制清空
                driver.execute_script("arguments[0].value = '';", start_hour_input)
                # 步骤2: 用 send_keys 输入，这会触发 onchange 事件
                start_hour_input.send_keys("00")

                driver.execute_script("arguments[0].value = '';", start_minute_input)
                start_minute_input.send_keys("00")
                print("开始时间已填写。")

                print("正在清空并填写结束时间...")
                driver.execute_script("arguments[0].value = '';", end_hour_input)
                end_hour_input.send_keys("23")

                driver.execute_script("arguments[0].value = '';", end_minute_input)
                end_minute_input.send_keys("59")
                print("结束时间已填写。")
                
                export_button = wait.until(EC.element_to_be_clickable((By.ID, 'jobHistoryExportExecute')))
                export_button.click()
                print("已点击导出链接。")

                # --- [核心修改] 采用新的、更全面的成功/失败判断逻辑 ---
                default_filename = "jobhist.csv"
                original_filepath = os.path.join(DOWNLOAD_DIR, default_filename)
                error_locator = "//*[contains(normalize-space(.), 'Unable to export') or normalize-space(.)='There was no Job History in the specified period.' or contains(normalize-space(.), 'A job is in progress')]"
                wait_time_for_reaction = 120
                is_successful_day = "pending" # 使用字符串来标记状态: pending, success, final_error
                
                print("等待服务器响应（最多120秒），可能是报错或开始下载...")
                for i in range(wait_time_for_reaction):
                    # 每一秒，我们都检查两件事：
                    
                    # 1. 是否出现了错误信息？
                    try:
                        error_element = driver.find_element(By.XPATH, error_locator)
                        error_text = error_element.text
                        
                        if "A job is in progress" in error_text:
                            print(f"!! [瞬时错误] 打印机正忙 (尝试 {attempt + 1}/{RETRY_COUNT})。")
                            wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Close')]"))).click()
                            raise Exception("Printer is busy, will retry...") # 抛出异常，触发外层的重试
                        else: # 否则，是最终反馈
                            if "Unable to export" in error_text:
                                print(f"!! [最终反馈] 日期 {current_date_str} 数据过多。")
                            elif "There was no" in error_text:
                                print(f"   [最终反馈] 日期 {current_date_str} 当日无数据。")
                            wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Close')]"))).click()
                            is_successful_day = "final_error" # 标记为已知错误，跳出等待
                            break
                    except NoSuchElementException:
                        # 没找到错误元素，很好，继续检查第二件事
                        pass

                    # 2. 文件是否已经下载完成？ (检查大小稳定)
                    if os.path.exists(original_filepath):
                        # 等待1秒看文件大小是否增长
                        time.sleep(1) 
                        if os.path.exists(original_filepath) and os.path.getsize(original_filepath) > 0:
                            is_successful_day = "success"
                            break # 成功，跳出等待
                    
                    time.sleep(1) # 等待1秒继续检查
                    print(f"\r等待中... {i+1}s / {wait_time_for_reaction}s", end="")

                print() # 换行

                # --- 根据等待结果，进行后续处理 ---
                if is_successful_day == "success":
                    print(f"[成功] 日期 {current_date_str} 文件已下载！正在处理...")
                    time.sleep(2) # 等待文件I/O完成
                    new_filename = f"jobhist_{current_date_str}.csv"
                    new_filepath = os.path.join(run_archive_dir, new_filename)
                    try:
                        shutil.move(original_filepath, new_filepath)
                        print(f"文件处理完毕: {new_filepath}")
                    except Exception as move_error:
                        print(f"!! 严重错误: 文件移动失败! Error: {move_error}")
                    break # 成功处理完一天，跳出重试循环

                elif is_successful_day == "final_error":
                    # 是已知的最终错误，跳出重试循环，处理下一天
                    break
                
                else: # 循环走完（超时）
                    print(f"!! [超时失败] 等待 {wait_time_for_reaction} 秒后，既未看到错误也未下载文件。")
                    raise TimeoutException(f"No reaction from server after {wait_time_for_reaction} seconds.")

            except Exception as day_error:
                print(f"!! [错误] 处理 {current_date_str} 时发生错误 (尝试 {attempt + 1}/{RETRY_COUNT})")
                print(f"   错误详情: {day_error}")
                if attempt == RETRY_COUNT - 1:
                    print(f"!!!!!! [最终失败] 日期 {current_date_str} 在重试 {RETRY_COUNT} 次后彻底失败，将跳过此天。 !!!!!!")
                else:
                    print(f"   将在 {RETRY_DELAY} 秒后重试...")
                    time.sleep(RETRY_DELAY)
                    driver.refresh()
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
        process_and_insert_data(directory_path=run_archive_dir, printer_type='floor4')
        print("调用 process_and_insert_data 函数。")
    except Exception as e:
        print(f"!!!!!! [阶段二 失败] 在数据处理或数据库插入过程中发生错误: {e} !!!!!!")
else:
    print("\n[任务中止] 由于阶段一（文件下载）未成功，阶段二（数据处理）被跳过。")

print("\n--- [任务总流程结束] ---")