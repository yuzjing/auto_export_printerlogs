import time
import os
import shutil
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from datetime import date, datetime
import calendar
from dateutil.relativedelta import relativedelta
import pandas as pd
import glob 
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException


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
IP_ADDRESS = '192.168.14.250'
AUTH_LOGIN_URL = f"http://{IP_ADDRESS}"
CHROME_BINARY_PATH = "/usr/sbin/google-chrome-unstable"
HOME_DIR = os.path.expanduser('~')
# 主存档目录，所有运行的原始数据都会存放在这里
ARCHIVE_DIR = os.path.join(HOME_DIR, 'PrinterReportsArchive/floor4')
# 最终合并报告的存放目录
FINAL_REPORT_DIR = os.path.join(HOME_DIR, 'PrinterReportsArchive/floor4')

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


# --- 2. 初始化浏览器  ---
options = Options()
options.binary_location = CHROME_BINARY_PATH
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--ignore-certificate-errors')
options.add_argument('--ignore-ssl-errors')
# options.add_argument('--headless')
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
        print(f"处理警告框时出错: {e}")

    print("登录成功，正在等待页面加载蒙版消失...")
    
    # 我们用 XPath 来定位这个作为“拦路虎”的蒙版
    # 它的 class 包含 'xux-modal'，并且它是一个 div
    loading_overlay_locator = (By.XPATH, "//div[contains(@class, 'xux-modal')]")
    
    long_wait = WebDriverWait(driver, 30)
    
    # 等待这个元素变为“不可见”。
    # 如果蒙版一开始就不存在，这个等待会立刻成功，非常安全。
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
                # 1. 检查模态框是否还开着 (通过查找一个只有它开着时才存在的元素，比如 'Start' 按钮)
                try:
                    cancel_button_locator = (By.XPATH, "//button/span[text()='Cancel']")
                    cancel_button = driver.find_element(*cancel_button_locator)
                    if cancel_button.is_displayed():
                        print("检测到上一次的模态框还开着，正在关闭...")
                        cancel_button.click()
                        # 等待它消失
                        wait.until(EC.invisibility_of_element_located(cancel_button_locator))
                        print("模态框已关闭。")
                except:
                    # 如果找不到 "Cancel" 按钮，说明模态框本来就是关着的，很好
                    print("页面状态正常，模态框已关闭。")
                
                # --- 现在，我们处于一个已知的、干净的起点，开始标准流程 ---
                
                # 2. 点击 'Export Job History' 按钮，打开模态框
                print("正在打开导出设置模态框...")
                export_history_button_locator = (By.ID, "openJobHistoryExportModalWindow")
                wait.until(EC.element_to_be_clickable(export_history_button_locator)).click()
                
                # 3. 等待模态框加载完成
                start_area_locator = (By.ID, "startDateTimeSection")
                start_area = wait.until(EC.presence_of_element_located(start_area_locator))
                print("模态框已成功打开并加载。")

                
                start_area = driver.find_element(By.ID, "startDateTimeSection")
                start_date_input = start_area.find_element(By.CSS_SELECTOR, "input.hasDatepicker")
                
                
                start_date_string = f"{EXPORT_YEAR}/{EXPORT_MONTH:02d}/{day_str}"
                start_date_input.clear()
                start_date_input.send_keys(start_date_string)
                time.sleep(0.3)
                print(f"已填写开始日期: {start_date_string}")
                start_hour_visible_input = start_area.find_element(By.XPATH, ".//input[@role='spinbutton' and @aria-label='Hour']")
                start_minute_visible_input = start_area.find_element(By.XPATH, ".//input[@role='spinbutton' and @aria-label='Minute']")
                print(f"正在通过 ActionChains 模拟键盘输入开始日期: {start_date_string}")
                ActionChains(driver).click(start_hour_visible_input).perform()
                time.sleep(0.3)
                driver.execute_script("arguments[0].value = '';", start_hour_visible_input)
                # 步骤 C: 【关键】加入一个极短的延迟，等待 JS 响应
                time.sleep(0.2)
    
                ActionChains(driver)\
                    .send_keys("00")\
                    .perform()

                # **分钟** (重复相同的稳健逻辑)
                ActionChains(driver).click(start_minute_visible_input).perform()
                time.sleep(0.3)
                driver.execute_script("arguments[0].value = '';", start_minute_visible_input)
                # 步骤 C: 【关键】加入一个极短的延迟，等待 JS 响应
                time.sleep(0.2)
                ActionChains(driver)\
                    .send_keys("00")\
                    .perform()

                print("开始时间已填写。")
        

                end_area = driver.find_element(By.ID, "endDateTimeSection")
                end_date_input = end_area.find_element(By.CSS_SELECTOR, "input.hasDatepicker")

                # 2. 在“结束”区域内部，查找对应的输入框
                end_date_string = f"{EXPORT_YEAR}/{EXPORT_MONTH:02d}/{day_str}"
                end_date_input.clear()
                end_date_input.send_keys(end_date_string)

                print(f"已填写结束日期: {end_date_string}")
                time.sleep(0.3)
                end_hour_visible_input = end_area.find_element(By.XPATH, ".//input[@role='spinbutton' and @aria-label='Hour']")
                end_minute_visible_input = end_area.find_element(By.XPATH, ".//input[@role='spinbutton' and @aria-label='Minute']")
                print(f"正在通过 ActionChains 模拟键盘输入结束日期: {end_date_string}")
    
                # 创建一个 ActionChains 对象
                actions = ActionChains(driver)
                # 3. 填写开始时间
                print("正在模拟键盘输入结束时间...")
                
                ActionChains(driver).click(end_hour_visible_input).perform()
                time.sleep(0.3)
                driver.execute_script("arguments[0].value = '';", end_hour_visible_input)
                # 步骤 C: 【关键】加入一个极短的延迟，等待 JS 响应
                time.sleep(0.2)
    
                ActionChains(driver)\
                    .send_keys("23")\
                    .perform()

                # **分钟** (重复相同的稳健逻辑)
                ActionChains(driver).click(end_minute_visible_input).perform()
                time.sleep(0.3)
                driver.execute_script("arguments[0].value = '';", end_minute_visible_input)
                # 步骤 C: 【关键】加入一个极短的延迟，等待 JS 响应
                time.sleep(0.2)
                ActionChains(driver)\
                    .send_keys("59")\
                    .perform()

                
                
            
                
                
                            
                
                
               
                print("所有日期已填写。")
                export_button = wait.until(EC.element_to_be_clickable((By.ID, 'jobHistoryExportExecute')))
                export_button.click()
                print("已点击导出链接。")

                

               

                # -- 判断结果 --
                time.sleep(1)
                error_locator = "//*[contains(text(), 'Unable to export') or contains(text(), 'There was no') or contains(text(), 'Unable to enter Administrator Mode') or contains(text(), 'A job is in progress')]"
                try:
                    
                    
                   # 尝试寻找任何一种已知的业务反馈信息
                    error_element = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, error_locator)))
                    error_text = error_element.text
                    
                    # 如果能执行到这里，说明页面上【一定】有某种错误信息
                    
                    # 1. 判断是否为“瞬时错误”
                    if "A job is in progress" in error_text:
                        print(f"!! [瞬时错误] 打印机正忙 (尝试 {attempt + 1}/{RETRY_COUNT})。")
                        # 点击关闭/返回按钮
                        close_button_locator = (By.XPATH, "//button[contains(., 'Close')]")
                        wait.until(EC.element_to_be_clickable(close_button_locator)).click()
                        # 主动抛出异常，触发【外层】的重试
                        raise Exception("Printer is busy, will retry...")
                    
                    # 2. 否则，就是“最终反馈”
                    else:
                        if "Unable to export" in error_text:
                            print(f"!! [最终反馈] 日期 {current_date_str} 数据过多。")
                        elif "There was no" in error_text:
                            print(f"   [最终反馈] 日期 {current_date_str} 当日无数据。")
                        
                        # 点击关闭/返回按钮
                        close_button_locator = (By.XPATH, "//button[contains(., 'Close')]")
                        wait.until(EC.element_to_be_clickable(close_button_locator)).click()
                        # 跳出【内层】的重试循环，处理下一天
                        break

                
                except TimeoutException:
                    print(f"[成功] 日期 {current_date_str} 导出成功，开始处理文件...")
                    # 1. 定义文件名和超时时间
                    default_filename = "jobhist.csv"
                    original_filepath = os.path.join(DOWNLOAD_DIR, default_filename)
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
                    driver.refresh() 
        else:
            # 这个 else 块属于 for 循环，只有在 for 循环正常结束（即 break 没有被执行）时才会运行
            # 在我们的场景里，如果成功了就会 break，所以这里可以留空或用于记录未曾成功 break 的情况
            pass
            
    # --- 所有循环都成功后 ---
    print(f"\n[任务完成] {EXPORT_YEAR}年{EXPORT_MONTH}月的所有数据已处理完毕！")
    driver.quit()

    print("\n--- [阶段二开始] 正在合并所有下载的 CSV 文件... ---")
    
    # 1. 找到本次运行下载的所有 CSV 文件
    csv_files = glob.glob(os.path.join(run_archive_dir, "*.csv"))

    csv_files.sort()
    
    if not csv_files:
        print("!! 警告: 在存档目录中未找到任何 CSV 文件，跳过合并步骤。")
    else:
        print(f"找到 {len(csv_files)} 个每日 CSV 文件，准备合并...")
        
        # 2. 逐个读取文件，并存放到一个列表中
        df_list = []
        for file in csv_files:
            try:
                # 尝试用 'gbk' 解码，因为很多中文Windows导出的CSV是这个格式
                df = pd.read_csv(file, encoding='gbk')
                # 提取文件名中的日期（如 jobhist_2025-03-01.csv → 2025-03-01）
                filename = os.path.basename(file)
                date_str = filename.replace('jobhist_', '').replace('.csv', '')
                # 补充日期列
                df['日期'] = date_str
                df_list.append(df)
            except UnicodeDecodeError:
                try:
                    # 如果 gbk 失败，再尝试 'utf-8'
                    df = pd.read_csv(file, encoding='utf-8')
                    filename = os.path.basename(file)
                    date_str = filename.replace('jobhist_', '').replace('.csv', '')
                    # 补充日期列
                    df['日期'] = date_str
                    df_list.append(df)
                except Exception as e:
                    print(f"!! 错误: 读取文件 {file} 失败: {e}")

        # 3. 如果列表不为空，就合并所有 DataFrame
        if df_list:
            # 使用 concat 一次性合并所有文件
            merged_df = pd.concat(df_list, ignore_index=True)

            if '日期' in merged_df.columns and '结束时间' in merged_df.columns:
                print("正在按 '日期' 和 '结束时间' 对数据进行排序...")
        
                # **方案A
                # 直接按两个字符串列进行排序
                merged_df.sort_values(by=['日期', '结束时间'], inplace=True)

                # **方案B
                # # 1. 将 '日期' 和 '结束时间' 合并成一个真正的时间戳列
                # merged_df['完整时间'] = pd.to_datetime(merged_df['日期'] + ' ' + merged_df['结束时间'])
                # # 2. 按这个新列排序
                # merged_df.sort_values(by='完整时间', inplace=True)
                # # 3. (可选) 删除临时的辅助列
                # merged_df.drop(columns=['完整时间'], inplace=True)

                print("数据排序完成！")
            else:
                print("!! 警告: 找不到 '日期' 或 '结束时间' 列，跳过排序步骤。")
            
            # 4. 定义最终合并文件的名称和路径
            final_filename = f"Monthly_Report_{EXPORT_YEAR}-{EXPORT_MONTH:02d}.csv"
            final_filepath = os.path.join(FINAL_REPORT_DIR, final_filename)
            
            # 5. 保存合并后的文件
            # 使用 'utf-8-sig' 编码可以确保 Excel 正确打开包含中文的 CSV 文件
            merged_df.to_csv(final_filepath, index=False, encoding='utf-8-sig')
            
            print(f"\n[成功] 所有 CSV 文件已合并！")
            print(f"最终报表已保存至: {final_filepath}")
            print(f"共包含 {len(merged_df)} 条记录。")



except Exception as e:
    # --- 只要上面 try 块中任何一步出错了，程序就会立刻跳到这里 ---
    
    print("\n!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
    print("【脚本执行失败】")
    print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
    
    # 1. 打印清晰的错误信息
    print(f"错误类型: {type(e).__name__}")
    # 打印具体的错误日志，这里包含了非常有用的信息
    print(f"错误详情:\n{e}")

   

    # 3. 使用 input() 暂停程序，等待你来排查
    print("\n【程序已暂停】浏览器保持在出错页面，请进行排查。")
    print("排查完毕后，请在此终端窗口按 Enter 键以结束程序。")
    # input("----------------------------------------------------")