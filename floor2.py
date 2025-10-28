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
from data_processor import process_and_insert_data



# 脚本设计为每月1号运行，自动获取上个月数据
today = date.today()
previous_month_date = today - relativedelta(months=1)
EXPORT_YEAR = previous_month_date.year
EXPORT_MONTH = previous_month_date.month

RETRY_COUNT = 5 # 设置每天的最大重试次数
RETRY_DELAY = 10 # 设置每次重试的间隔时间（秒）

# --- 1. 配置信息 (保持不变) ---
USERNAME = '11111'
PASSWORD = 'x-admin'
IP_ADDRESS = '192.168.12.250'
AUTH_LOGIN_URL = f"http://{USERNAME}:{PASSWORD}@{IP_ADDRESS}"
CHROME_BINARY_PATH = "/usr/sbin/google-chrome-unstable"
HOME_DIR = os.path.expanduser('~')
# 主存档目录，所有运行的原始数据都会存放在这里
ARCHIVE_DIR = os.path.join(HOME_DIR, 'PrinterReportsArchive/floor2')
# 最终合并报告的存放目录
FINAL_REPORT_DIR = os.path.join(HOME_DIR, 'PrinterReportsFinal/floor2')

# 根据当前运行日期，创建一个本次任务专属的子文件夹
run_date_str = today.strftime('%Y-%m-%d')
run_archive_dir = os.path.join(ARCHIVE_DIR, f"Run_{run_date_str}")
DOWNLOAD_DIR = run_archive_dir

os.makedirs(run_archive_dir, exist_ok=True)
os.makedirs(FINAL_REPORT_DIR, exist_ok=True)

# 清理下载目录中的旧文件
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


# --- 2. 初始化浏览器 (保持不变) ---
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

print("正在自动配置并启动浏览器...")
driver = webdriver.Chrome(options=options)
print("浏览器启动成功！")

wait = WebDriverWait(driver, 10) # 设置10秒的智能等待


try:
    # --- 步骤 1: 登录 ---
    print("正在使用认证方式访问...")
    driver.get(AUTH_LOGIN_URL)
    print("已发送登录请求。")

    # --- 步骤 2: 处理 JavaScript 警告框 ---
    print("等待并处理安全警告框...")
    try:
        # 我们把处理警告框也放在一个小的 try-except 里，这样即使不弹框也不会报错
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
        # 如果找不到框架，直接抛出异常，触发调试暂停
        raise 

    print("正在点击 'Properties' 链接...")
    properties_link_locator = "a[href='prop.htm']"
    properties_link = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, properties_link_locator)))
    properties_link.click()
    print("已成功点击 'Properties'。")


    # **步骤 A: 先从当前框架（TF）切回顶层**
    # 这是关键的第一步，让 Selenium 的焦点回到页面的最外层
    print("正在从 'TF' 框架切换回顶层...")
    driver.switch_to.default_content()
    print("已切换回顶层。WebDriver 正在重新评估页面结构...")
    
    # **步骤 B: 现在，在新结构下，等待并切换到 'NF' 框架**
    # 因为我们已经回到了顶层，Selenium 现在应该能“看”到 JS 画出来的新框架了
    print("正在智能等待并切换到新出现的 'NF' 框架...")
    wait.until(EC.frame_to_be_available_and_switch_to_it("NF"))
    print("成功切换到 'NF' 框架！")



    print("正在点击 'General Setup' 链接...")
    general_setup_locator = "//a[contains(@href, \"flip('nvTree',2)\")]"
    general_setup_link = wait.until(EC.element_to_be_clickable((By.XPATH, general_setup_locator)))
    general_setup_link.click()
    print("已成功点击 'General Setup'。")

    # --- 点击 'Job Management' 链接 ---
    print("正在点击 'Job Management' 链接...")
    job_management_locator = "//a[contains(@href, \"flip('genItms',1)\")]"
    job_management_link = wait.until(EC.element_to_be_clickable((By.XPATH, job_management_locator)))
    job_management_link.click()
    print("已成功点击 'Job Management'。")

    # --- 点击 'Export Job History' 链接 ---
    # 这是我们最终的目标页面
    print("正在点击 'Export Job History' 链接...")
    export_history_locator = "//a[contains(@href, \"dclick('jobItms',1)\")]"
    export_history_link = wait.until(EC.element_to_be_clickable((By.XPATH, export_history_locator)))
    export_history_link.click()
    print("已成功点击 'Export Job History'，进入最终页面！")
    print("正在从 'NF' 框架切换回顶层...")
    driver.switch_to.default_content()

    # --- 步骤 7: 在最终页面上进行操作 ---
    # 现在您已经到达了“Export Job History”页面，可以开始填写日期、点击查询了
    print("正在填写日期...")
    dateform_frame_name = "RF"
    print("正在从 'NF' 框架切换到 'RF' 框架...")
    wait.until(EC.frame_to_be_available_and_switch_to_it(dateform_frame_name))
    print("已切换到 'RF' 框架。")

    print("正在填写开始日期...")




    # --- 按天循环导出 ---
    _, num_days = calendar.monthrange(EXPORT_YEAR, EXPORT_MONTH)
    
    for day in range(1, num_days + 1):
        day_str = f"{day:02d}"
        month_str = f"{EXPORT_MONTH:02d}"
        current_date_str = f"{EXPORT_YEAR}-{month_str}-{day_str}"
        
        # 【【【新增：重试循环】】】
        for attempt in range(RETRY_COUNT):
            try:
                print(f"\n--- [尝试 {attempt + 1}/{RETRY_COUNT}] 开始处理日期: {current_date_str} ---")

                # -- 填写日期 -- (等待页面元素稳定)
                start_year_input = wait.until(EC.presence_of_element_located((By.NAME, 'YEAR1')))
                start_year_input.clear()
                start_year_input.send_keys(EXPORT_YEAR)
                start_month_input = driver.find_element(By.NAME, 'MONTH1')
                start_month_input.clear()
                start_month_input.send_keys(EXPORT_MONTH)
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
                end_year_input.send_keys(EXPORT_YEAR)
                end_month_input = driver.find_element(By.NAME, 'MONTH2')
                end_month_input.clear()
                end_month_input.send_keys(EXPORT_MONTH)
                end_day_input = driver.find_element(By.NAME, 'DAY2')
                end_day_input.clear()
                end_day_input.send_keys(day_str)
                end_hour_input = driver.find_element(By.NAME, 'HOUR2')
                end_hour_input.clear()
                end_hour_input.send_keys("23")
                end_miniute_input = driver.find_element(By.NAME, 'MIN2')
                end_miniute_input.clear()
                end_miniute_input.send_keys("59")
                # ... (填写开始和结束日期的代码) ...
                print("日期已填写。")

                # -- 点击导出 --
                export_link_locator = "//a[contains(text(), 'Export file in')]"
                wait.until(EC.element_to_be_clickable((By.XPATH, export_link_locator))).click()
                print("已点击导出链接。")

                # -- 判断结果 --
                time.sleep(3)
                error_locator = "//*[contains(text(), 'Unable to export') or contains(text(), 'There is no job') or contains(text(), 'Your request was not')]"
                try:
                    # 尝试寻找失败标志
                    error_element = driver.find_element(By.XPATH, error_locator)
                    error_text = error_element.text
                    
                    
                    # 如果能找到，我们就判断一下是哪种错误
                    # 定义一个标志位，用于判断是否需要 break
                    is_final_feedback = False
                    
                    # --- 开始判断错误类型 ---
                    
                    if "Unable to export" in error_text:
                        print(f"!! [最终反馈] 日期 {current_date_str} 数据过多。")
                        is_final_feedback = True
                    
                    elif "no job" in error_text:
                        print(f"   [最终反馈] 日期 {current_date_str} 当日无数据。")
                        is_final_feedback = True
                    
                    elif "Your request was not" in error_text:
                        print(f"!! [瞬时错误] 打印机正忙 (尝试 {attempt + 1}/{RETRY_COUNT})。")
                        # 瞬时错误不需要设置 is_final_feedback，因为它要重试
                    
                    # 无论哪种业务反馈，我们都需要点击 "Back" 按钮
                    driver.find_element(By.XPATH, "//input[@value='Back']").click()
                    
                    # --- 根据错误类型，决定下一步行动 ---
                    
                    if is_final_feedback:
                        # 如果是“最终反馈”，我们就 break 跳出重试循环
                        break
                    else:
                        # 否则（就是瞬时错误），我们就抛出异常以触发重试
                        raise Exception("Printer is busy or unknown error, will retry...")
                except NoSuchElementException:
                    print(f"[成功] 日期 {current_date_str} 导出成功，开始处理文件...")
                    # 1. 定义文件名和超时时间
                    default_filename = "jobhist.csv"
                    original_filepath = os.path.join(DOWNLOAD_DIR, default_filename)
                    # 添加调试信息
                    #print(f"期望的文件路径: {original_filepath}")
                    #print(f"当前工作目录: {os.getcwd()}")
                    #print(f"下载目录内容: {os.listdir(DOWNLOAD_DIR) if os.path.exists(DOWNLOAD_DIR) else '目录不存在'}")
                    download_wait_time = 120 # 最长等待120秒
                    time_elapsed = 0

                    # 3. 轮询等待新文件下载完成
                    while not os.path.exists(original_filepath) and time_elapsed < download_wait_time:
                        time.sleep(1)
                        time_elapsed += 1
                        print(f"\r正在等待下载... {time_elapsed}s / {download_wait_time}s", end="")
                    
                    print() # 换行

                    # 4. 检查等待结果并处理文件
                    if os.path.exists(original_filepath):
                        print(f"文件 {default_filename} 已成功下载！")
                        time.sleep(2) # 等待文件I/O完成
                        
                        new_filename = f"jobhist_{current_date_str}.csv"
                        # 将文件移动到本次运行的专属文件夹内
                        new_filepath = os.path.join(run_archive_dir, new_filename)
                        
                        try:
                            shutil.move(original_filepath, new_filepath)
                            print(f"文件处理完毕: {new_filepath}")
                        except Exception as move_error:
                            print(f"!! 严重错误: 文件移动失败! Error: {move_error}")
                    else:
                        print(f"!! [下载失败] 等待 {download_wait_time} 秒后，未找到下载文件。可能当日无数据。")

                    print(f"文件处理完毕: jobhist_{current_date_str}.csv")
                    
                    # 【关键】既然成功了，就用 break 跳出重试循环
                    break

            except Exception as day_error:
                print(f"!! [错误] 处理 {current_date_str} 时发生错误 (尝试 {attempt + 1}/{RETRY_COUNT})")
                print(f"   错误详情: {day_error}")
                
                # 如果这是最后一次尝试，记录最终失败
                if attempt == RETRY_COUNT - 1:
                    print(f"!!!!!! [最终失败] 日期 {current_date_str} 在重试 {RETRY_COUNT} 次后彻底失败，将跳过此天。 !!!!!!")
                else:
                    # 否则，等待后进行下一次重试
                    print(f"   将在 {RETRY_DELAY} 秒后重试...")
                    time.sleep(RETRY_DELAY)
                    # 刷新页面，回到一个干净的状态，准备重试
                    # driver.refresh() 
        else:
            # 这个 else 块属于 for 循环，只有在 for 循环正常结束（即 break 没有被执行）时才会运行
            # 在我们的场景里，如果成功了就会 break，所以这里可以留空或用于记录未曾成功 break 的情况
            pass
            
    # --- 所有循环都成功后 ---
    print(f"\n[任务完成] {EXPORT_YEAR}年{EXPORT_MONTH}月的所有数据已处理完毕！")
    

    download_successful = True # 如果程序能顺利跑到这里，说明下载成功
    
except Exception as e:
    print(f"!!!!!! [阶段一 失败] 在浏览器自动化过程中发生严重错误: {e} !!!!!!")
    # 异常发生，download_successful 依然是 False

finally:
    if driver:
        driver.quit()
        print("浏览器已关闭。")

# --- [阶段二] 开始: 处理数据并写入数据库 ---
# 只有在第一阶段成功完成后，才执行第二阶段
if download_successful:
    print("\n--- [阶段二 开始] 处理已下载文件并写入数据库 ---")
    try:
        # 直接调用处理函数，传入刚刚保存文件的目录和打印机类型
        process_and_insert_data(directory_path=run_archive_dir, printer_type='floor2')
    except Exception as e:
        print(f"!!!!!! [阶段二 失败] 在数据处理或数据库插入过程中发生错误: {e} !!!!!!")
else:
    print("\n[任务中止] 由于阶段一（文件下载）未成功，阶段二（数据处理）被跳过。")

print("\n--- [任务总流程结束] ---")
    