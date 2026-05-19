import logging
import os
import re
import threading
import time
import urllib.parse
import concurrent.futures
import queue
import random
from datetime import datetime


import pandas as pd
from lxml import etree
from selenium import webdriver
from selenium.common import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
# from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait


input_xlsx = '20250915报刊整理.xlsx'

driver_path = "/Users/admin/Desktop/taobao/chromedriver-mac-arm64/chromedriver"
# 定义全局Cookie变量
login_page = 'https://dclibrary.idm.oclc.org/login?url=https://infoweb.newsbank.com/signin/DistrictofColumbiaPublicLibrary/AWNB'

# 账号配置 - 每个账号对应一个驱动和线程
account_configs = [
    {
        'username': '41172493532509',
        'password': '1986',
        'driver': None,
        'thread_id': 1
    },
    {
        'username': '41172636532510', 
        'password': '1973',
        'driver': None,
        'thread_id': 2
    },
    {
        'username': '41172229532518',
        'password': '1966', 
        'driver': None,
        'thread_id': 3
    },
    {
        'username': '41172311532521',
        'password': '1999',
        'driver': None,
        'thread_id': 4
    },
    {
        'username': '41172898532522',
        'password': '1999',
        'driver': None,
        'thread_id': 5
    },
    {
        'username': '41172786532523',
        'password': '1999',
        'driver': None,
        'thread_id': 6
    },
    {
        'username': '41172986532526',
        'password': '1999',
        'driver': None,
        'thread_id': 7
    },
    {
        'username': '41172578532527',
        'password': '1999',
        'driver': None,
        'thread_id': 8
    },
    {
        'username': '41172433532528',
        'password': '1999',
        'driver': None,
        'thread_id': 9
    },
    {
        'username': '41172891532529',
        'password': '1999',
        'driver': None,
        'thread_id': 10
    },
    {
        'username': '41172103532530',
        'password': '1999',
        'driver': None,
        'thread_id': 11
    },
    {
        'username': '41172164532535',
        'password': '1999',
        'driver': None,
        'thread_id': 12
    },
    {
        'username': '41172990532537',
        'password': '1999',
        'driver': None,
        'thread_id': 13
    },
    {
        'username': '41172227532539',
        'password': '1999',
        'driver': None,
        'thread_id': 14
    },
    {
        'username': '41172808532543',
        'password': '1999',
        'driver': None,
        'thread_id': 15
    },
    {
        'username': '41172513532546',
        'password': '1999',
        'driver': None,
        'thread_id': 16
    },
    # {
    #     'username': 'PACREG534168',
    #     'password': '1999',
    #     'driver': None,
    #     'thread_id': 17
    # },
    # {
    #     'username': '41172319534169',
    #     'password': '1999',
    #     'driver': None,
    #     'thread_id': 18
    # },
    # {
    #     'username': '41172793534170',
    #     'password': '1999',
    #     'driver': None,
    #     'thread_id': 19
    # },
    # {
    #     'username': '41172485534171',
    #     'password': '1999',
    #     'driver': None,
    #     'thread_id': 20
    # },
    # {
    #     'username': '41172111534172',
    #     'password': '1999',
    #     'driver': None,
    #     'thread_id': 21
    # },
    # {
    #     'username': '41172121534173',
    #     'password': '1999',
    #     'driver': None,
    #     'thread_id': 22
    # },
    # {
    #     'username': '41172527534174',
    #     'password': '1999',
    #     'driver': None,
    #     'thread_id': 23
    # },
    # {
    #     'username': '41172847534175',
    #     'password': '1999',
    #     'driver': None,
    #     'thread_id': 24
    # },
    # {
    #     'username': '41172307534176',
    #     'password': '1999',
    #     'driver': None,
    #     'thread_id': 25
    # },
    # {
    #     'username': '41172974534177',
    #     'password': '1999',
    #     'driver': None,
    #     'thread_id': 26
    # },
    # {
    #     'username': '22022100794609',
    #     'password': '1999',
    #     'driver': None,
    #     'thread_id': 27
    # },
    # {
    #     'username': '22022100794617',
    #     'password': '1999',
    #     'driver': None,
    #     'thread_id': 28
    # },
    # {
    #     'username': '22022100794625',
    #     'password': '1999',
    #     'driver': None,
    #     'thread_id': 29
    # },
    # {
    #     'username': '22022100793999',
    #     'password': '1999',
    #     'driver': None,
    #     'thread_id': 30
    # },
    # {
    #     'username': '22022100794005',
    #     'password': '1999',
    #     'driver': None,
    #     'thread_id': 31
    # },
    # {
    #     'username': '22022100794021',
    #     'password': '1999',
    #     'driver': None,
    #     'thread_id': 32
    # },
    # {
    #     'username': '22022100794039',
    #     'password': '1999',
    #     'driver': None,
    #     'thread_id': 33
    # },
    # {
    #     'username': '22022100794088',
    #     'password': '1999',
    #     'driver': None,
    #     'thread_id': 34
    # },
    # {
    #     'username': '22022100794096',
    #     'password': '1999',
    #     'driver': None,
    #     'thread_id': 35
    # },
]

domain_url = 'https://infoweb-newsbank-com.dclibrary.idm.oclc.org'
country_list_file = "country_small.txt"
journal_name_to_code = {}
no_journal_name = []

# 保持向后兼容
driver = None

# 重试机制配置
# 当二组合搜索失败(-1)或处理异常(-2)时，会自动触发重试机制
# 重试过程中会：
# 1. 使用递增延迟避免频繁请求
# 2. 尝试重建浏览器实例恢复连接
# 3. 清理失败记录并更新为新的结果
# 4. 自动识别并重新处理历史失败的组合
RETRY_CONFIG = {
    'max_retries': 3,                    # 最大重试次数
    'base_retry_delay': 5,               # 基础重试延迟（秒）
    'max_retry_delay': 15,               # 最大重试延迟（秒）
    'exception_retry_delay_multiplier': 2,  # 异常重试延迟倍数
    'enable_browser_recreate': True      # 是否启用浏览器重建
}


def get_driver_for_thread(thread_id):
    """根据线程ID获取对应的浏览器实例"""
    for config in account_configs:
        if config['thread_id'] == thread_id:
            return config['driver']
    return None

def get_account_for_thread(thread_id):
    """根据线程ID获取对应的账号信息"""
    for config in account_configs:
        if config['thread_id'] == thread_id:
            return config['username'], config['password']
    return None, None

def recreate_driver_for_config(config, logger_thread):
    """为指定配置重新创建浏览器实例"""
    try:
        # 先尝试关闭旧的driver
        if config['driver']:
            try:
                config['driver'].quit()
                logger_thread.info(f"[浏览器{config['thread_id']}|账号{config['username']}] 🔄 旧浏览器实例已关闭")
            except:
                pass  # 忽略关闭失败的错误
        
        # 创建新的浏览器实例

        service = Service(driver_path)
        options = webdriver.ChromeOptions()
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-plugins")
        options.add_argument("--disable-images")
        options.add_argument("--incognito")
        # 禁用 Selenium 自动化标识（关键反爬处理）
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option("useAutomationExtension", False)
        
        # 动态生成User-Agent
        def generate_user_agent(thread_id):
            """根据线程ID动态生成User-Agent"""
            
            # Chrome版本号范围 (最新的稳定版本)
            chrome_versions = [
                "120.0.0.0", "121.0.0.0", "122.0.0.0", "123.0.0.0", "124.0.0.0"
            ]
            
            # 操作系统版本
            windows_versions = [
                "Windows NT 10.0; Win64; x64",
                "Windows NT 11.0; Win64; x64"
            ]
            
            macos_versions = [
                "Macintosh; Intel Mac OS X 10_15_7",
                "Macintosh; Intel Mac OS X 13_6_0", 
                "Macintosh; Intel Mac OS X 14_1_0",
                "Macintosh; Intel Mac OS X 14_2_0"
            ]
            
            # 根据线程ID确定浏览器类型和操作系统
            browser_type = thread_id % 2  # 0=Chrome Windows, 1=Chrome macOS
            
            if browser_type == 0:  # Chrome Windows
                os_info = windows_versions[thread_id % len(windows_versions)]
                chrome_ver = chrome_versions[thread_id % len(chrome_versions)]
                webkit_ver = f"537.{36 + (thread_id % 10)}"
                return f"Mozilla/5.0 ({os_info}) AppleWebKit/{webkit_ver} (KHTML, like Gecko) Chrome/{chrome_ver} Safari/{webkit_ver}"
            else:  # Chrome macOS
                os_info = macos_versions[thread_id % len(macos_versions)]
                chrome_ver = chrome_versions[thread_id % len(chrome_versions)]
                webkit_ver = f"537.{36 + (thread_id % 10)}"
                return f"Mozilla/5.0 ({os_info}) AppleWebKit/{webkit_ver} (KHTML, like Gecko) Chrome/{chrome_ver} Safari/{webkit_ver}"
        
        # 为每个线程ID生成唯一的User-Agent
        selected_ua = generate_user_agent(config['thread_id'])
        options.add_argument(f"user-agent={selected_ua}")
        
        # 为每个线程ID分配不同的视口大小
        window_sizes = [
            (1366, 768), (1920, 1080), (1440, 900), (1536, 864), (1600, 900),
            (1280, 800), (1680, 1050), (2560, 1440), (1280, 1024), (1024, 768)
        ]
        size_index = (config['thread_id'] - 1) % len(window_sizes)
        width, height = window_sizes[size_index]
        options.add_argument(f"--window-size={width},{height}")
        
        # 创建新的driver
        config['driver'] = webdriver.Chrome(service=service, options=options)
        logger_thread.info(f"[浏览器{config['thread_id']}|账号{config['username']}] ✅ 浏览器重建成功！UA: {selected_ua[:60]}... 视口: {width}x{height}")
        
        # 重新登录
        login_success = get_cookie_for_thread(config['thread_id'], logger_thread, config['username'], config['password'])
        if login_success:
            logger_thread.info(f"[浏览器{config['thread_id']}|账号{config['username']}] ✅ 重建后登录成功")
            return True
        else:
            logger_thread.error(f"❌ [浏览器{config['thread_id']}|账号{config['username']}]  重建后登录失败")
            return False
            
    except Exception as e:
        logger_thread.error(f"❌ [浏览器{config['thread_id']}|账号{config['username']}] ❌ 浏览器重建失败: {e}")
        return False

def check_driver_health(driver, logger_thread, context=""):
    """检查浏览器驱动是否健康"""
    try:
        # 尝试获取当前URL来测试连接
        current_url = driver.current_url
        # 尝试获取页面标题
        title = driver.title
        # logger_thread.debug(f"{context} 🔍 浏览器健康检查通过: {title[:30]}...")
        return True
    except Exception as e:
        logger_thread.warning(f"⚠️ {context} 浏览器健康检查失败: {e}")
        return False

def init_driver():
    """初始化浏览器driver（封装配置，便于重建）"""
    global driver
    try:
        # 使用本地 ChromeDriver
        service = Service(driver_path)
        options = webdriver.ChromeOptions()
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        
        driver = webdriver.Chrome(service=service, options=options)
    except Exception as e:
        print(f"driver初始化失败: {e}")
        return None


def read_file(file_path, mode='r', encoding='utf-8'):
    try:
        # 使用with语句打开文件，自动处理资源释放
        with open(file_path, mode, encoding=encoding) as file:
            # 读取全部内容
            content = file.read()
            decoded_content = content.encode('latin-1').decode('unicode_escape')
            logger_global.info(f"成功读取文件: {file_path}")
            return decoded_content
    except Exception as e:
        logger_global.info(f"读取文件时发生错误: {str(e)}")
        raise SystemExit(1)


def find_pattern_country(name, country_content):
    pattern = re.escape(name) + r'.*?",\s*("pubname:.*?")'
    match = re.search(pattern, country_content)
    if match:
        return match.group(1).strip('"') # 提取第一个捕获组（pubname:和!之间的内容）
    else:
        raise SystemExit(f"存在异常杂志社：{name}")


def get_total(html):
    # 检测403错误
    if "403 Forbidden" in html or "<title>403 Forbidden</title>" in html:
        logger_global.error("🚨 检测到403 Forbidden错误，网站可能已封禁访问")
        raise Exception("403 Forbidden - 访问被拒绝")
    
    tree = etree.HTML(html)
    tree_layout_content = tree.xpath('//*[@class="layout__content-above layout__content-above--results"]')
    if len(tree_layout_content) == 0:
        logger_global.info(f"访问网页失败!!! ")
        raise Exception(f"访问网页失败！！！")
    total_list = tree_layout_content[0].xpath('//*[@class="search-hits__meta--total_hits"]/text()')
    if len(total_list) == 0:
        return 0
    total = total_list[0]
    a = re.sub('\\n| |Results|,', '', total)
    return a

RESULT_SEPARATOR = "---QUERY_RESULT_END---\n"
def append_to_file(content, mode='a',filename="results.txt"):
    """将查询得到的字符串按行追加到txt文件"""
    with open(filename, mode=mode, encoding='utf-8') as f:
        f.write(content)
        if mode == 'a':
            if filename == "condition.txt":
                f.write('\n')  # condition.txt 使用换行符分隔
            else:
                f.write(RESULT_SEPARATOR)  # results.txt 使用自定义分隔符


def read_as_set(filename="results.txt"):
    """从文件读取所有查询结果，转换为去重的set"""
    if not os.path.exists(filename):
        return set()
    with open(filename, 'r', encoding='utf-8') as f:
        all_content = f.read()
        # 根据文件类型使用不同的分隔符
        if filename == "condition.txt":
            results = [res.strip() for res in all_content.split('\n') if res.strip()]
        else:
            results = [res for res in all_content.split(RESULT_SEPARATOR) if res.strip()]
    return set(results)


def write_csv_detail(name, all_map, filter_map):
    sheet_one_value = list()
    for year in all_map.keys():
        all_value = all_map.get(year, 0)
        filter_value = len(filter_map.get(year, set()))
        sheet_one_value.append((year, str(all_value), str(filter_value)))
    df_sheet_one = pd.DataFrame(sheet_one_value)
    # 创建一个 Excel 写入器对象
    with pd.ExcelWriter(f'search_results/{name}', mode='w') as writer:
        # 写入第一个数据框到第一个 sheet，设置 sheet 名称和标题行
        df_sheet_one.to_excel(writer, sheet_name='汇总', index=False,
                              header=['year', 'all_value', 'filter_value'])
        # 写入每年的标题
        for year in all_map.keys():
            year_set = filter_map.get(year, set())
            if len(year_set) > 0:
                df_sheet_other = pd.DataFrame((content.split('|||') for content in year_set))
                df_sheet_other.to_excel(writer, sheet_name=year, index=False,
                                        header=['date', 'title', 'perview'])


def check_file_existence_with_pathlib(filename):
    """使用pathlib模块检查文件是否存在（Python 3.4+）"""
    # 创建Path对象
    file_path = "search_results/" + filename
    # 检查文件是否存在
    if os.path.exists(file_path):
        logger_global.info(f"文件 '{file_path}' 在 search_results/ 目录中存在")
        return True
    else:
        logger_global.info(f"文件 '{file_path}' 在 search_results/ 目录中不存在")
        return False


def request_and_map(filter_html):
    tree = etree.HTML(filter_html)
    # 获取详情存入filter_map
    filter_root_list = tree.xpath('//*[@class="search-hits__hit__inner"]')
    for filter_root in filter_root_list:
        title = filter_root.xpath('.//*[@class="search-hits__hit__title search-hits__title"]/a/text()')[0]
        date = filter_root.xpath(
            './/*[@class="search-hits__hit__meta__item search-hits__hit__meta__item--display-date"]/text()')[0].strip()
        preview = filter_root.xpath('.//*[@class="preview-first-paragraph"]/text()')[0]
        # 添加新值到集合中
        str = '|||'.join([date, title, preview])
        # 追加查询结果到临时文件
        append_to_file(str)


def concat_url(index, base_url, keyword, relation):
    str2 = f'val-base-{index}={urllib.parse.quote(keyword)}'
    str3 = f'fld-base-{index}=alltext'
    if index != 0:
        str1 = f'bln-base-{index}={urllib.parse.quote(relation)}'
        return '&'.join([base_url, str1, str2, str3])
    return '&'.join([base_url, str2, str3])


def is_on_the_hour():
    """判断当前时间是否为整点（分钟和秒数都为0）"""
    now = datetime.now()
    # 检查分钟和秒数是否都为0
    return now.minute == 0 or now.second == 30




def get_cookie_for_thread(thread_id, logger_thread, username, password):
    """为指定线程获取Cookie - 已修改为无密码登录"""
    driver = get_driver_for_thread(thread_id)
    if not driver:
        logger_thread.error(f"❌ [线程{thread_id}|账号{username}] 无法获取对应的浏览器驱动")
        return False
    
    try:
        logger_thread.info(f"[线程{thread_id}|账号{username}] 开始登录，访问登录页面...")
        driver.get(login_page)
        wait = WebDriverWait(driver, 100, 0.5)

        # 1. 等待用户名输入框可点击
        username_input = wait.until(
            EC.element_to_be_clickable((By.NAME, "user"))
        )
        driver.execute_script("arguments[0].scrollIntoView();", username_input)
        username_input.clear()
        username_input.send_keys(username)

        # 2. 密码输入步骤已移除 - 无需密码登录

        # 3. 点击登录按钮
        login_button = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button"))
        )
        driver.execute_script("arguments[0].scrollIntoView();", login_button)
        try:
            login_button.click()
        except:
            driver.execute_script("arguments[0].click();", login_button)

        # 4. 等待跳转完成
        wait.until(EC.title_contains("Access World News"))
        logger_thread.info(f"[线程{thread_id}|账号{username}] ✅ 登录成功（无密码模式）")
        return True

    except TimeoutException as e:
        logger_thread.error(f"❌ [线程{thread_id}|账号{username}] 登录元素超时未加载: {e}")
        return False
    except Exception as e:
        logger_thread.error(f"❌ [线程{thread_id}|账号{username}] 登录异常: {e}")
        return False


# 全局变量记录周期开始时间
cycle_start_time = None
# @retry(stop_max_attempt_number=3, wait_random_min=1000, wait_random_max=3000)
def request_url(url, logger_thread):
    # global cycle_start_time
    # if cycle_start_time is None:
    #     cycle_start_time = datetime.now()
    #     logger_thread.info(f"程序开始时间: {cycle_start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    # # 检查是否已经过了1小时
    # current_time = datetime.now()
    # elapsed_seconds = (current_time - cycle_start_time).total_seconds()
    # if elapsed_seconds >= 3 * 60 * 60:  # 3600秒 = 1小时
    #     time.sleep(30 * 60)
    #     # 重置周期开始时间
    #     cycle_start_time = datetime.now()
    #     logger_thread.info(f"休息结束时间: {cycle_start_time.strftime('%H:%M:%S')}")
    #     logger_thread.info("开始新的周期...")
    driver.get(url)
    time.sleep(13)
    wait = WebDriverWait(driver, 100, 0.5)
    try:
        # 等待页面主体内容加载（超时10秒）
        wait.until(EC.title_contains("Access World News"))
    except TimeoutException:
        logger_thread.error("❌ 页面无响应，加载超时，尝试刷新页面")
        driver.refresh()
        time.sleep(1)
        try:
            # 刷新后再次等待页面加载
            wait.until(EC.title_contains("Access World News"))
            logger_thread.info("页面刷新后加载成功")
        except TimeoutException:
            logger_thread.error("❌页面刷新后仍然超时")
    return driver.page_source


def click_search_for_thread(conditions, base_url, logger_thread, thread_driver, thread_id=None, username=None):
    """多线程版本的click_search函数，使用指定的driver"""
    global cycle_start_time
    if cycle_start_time is None:
        cycle_start_time = datetime.now()
        logger_thread.info(f"程序开始时间: {cycle_start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    # 检查是否已经过了1小时 - 已禁用休息机制
    # current_time = datetime.now()
    # elapsed_seconds = (current_time - cycle_start_time).total_seconds()
    # if elapsed_seconds >= 40 * 60:  # 3600秒 = 1小时
    #     # 随机休息10到20分钟
    #     sleep_minutes = random.randint(5, 10)
    #     sleep_seconds = sleep_minutes * 60
    #     logger_thread.info(f"开始休息 {sleep_minutes} 分钟...")
    #     time.sleep(sleep_seconds)
    #     # 重置周期开始时间
    #     cycle_start_time = datetime.now()
    #     logger_thread.info(f"休息结束时间: {cycle_start_time.strftime('%H:%M:%S')}")
    #     logger_thread.info("开始新的周期...")
    try:
        add_condition = WebDriverWait(thread_driver, 100).until(
            EC.element_to_be_clickable((By.ID, "addTriplet"))
        )
        add_condition.click()
        # 等待下拉框元素加载完成并可交互
        for i in range(3):
            select_element = WebDriverWait(thread_driver, 100).until(
                EC.element_to_be_clickable((By.NAME, f"fld-base-{str(i)}"))
            )
            select = Select(select_element)
            select.select_by_value("alltext")
            condition = WebDriverWait(thread_driver, 100).until(
                EC.element_to_be_clickable((By.NAME, f"val-base-{str(i)}"))
            )
            condition.clear()
            if i < len(conditions):
                condition.send_keys(conditions[i])
        search_button = WebDriverWait(thread_driver, 100).until(
            EC.element_to_be_clickable((By.ID, "edit-submit"))
        )
        search_button.click()
        time.sleep(13)
        try:
            WebDriverWait(thread_driver, 100).until(
                EC.presence_of_element_located((By.ID, "main-content"))
            )
        except TimeoutException:
            context = f"[线程{thread_id}|账号{username}]" if thread_id and username else ""
            logger_thread.error(f"❌ {context} 等待main-content元素超时，尝试刷新页面")
            thread_driver.refresh()
            time.sleep(1)
            try:
                WebDriverWait(thread_driver, 100).until(
                    EC.presence_of_element_located((By.ID, "main-content"))
                )
                logger_thread.info(f"❌ {context} 页面刷新后main-content元素加载成功")
            except TimeoutException:
                logger_thread.error(f"❌ {context} 页面刷新后main-content元素仍然超时")
    except Exception as e:
        context = f"[线程{thread_id}|账号{username}]" if thread_id and username else ""
        logger_thread.error(f"❌ {context} 页面无响应，加载超时")
        # 尝试刷新页面并重新执行搜索
        try:
            thread_driver.refresh()
            time.sleep(1)
            logger_thread.info(f"{context} 异常后尝试刷新页面")
            
            # 刷新后需要重新设置搜索条件并执行搜索
            add_condition = WebDriverWait(thread_driver, 100).until(
                EC.element_to_be_clickable((By.ID, "addTriplet"))
            )
            add_condition.click()
            
            # 重新设置搜索条件
            for i in range(3):
                select_element = WebDriverWait(thread_driver, 100).until(
                    EC.element_to_be_clickable((By.NAME, f"fld-base-{str(i)}"))
                )
                select = Select(select_element)
                select.select_by_value("alltext")
                condition = WebDriverWait(thread_driver, 100).until(
                    EC.element_to_be_clickable((By.NAME, f"val-base-{str(i)}"))
                )
                condition.clear()
                if i < len(conditions):
                    condition.send_keys(conditions[i])
            
            # 重新点击搜索按钮
            search_button = WebDriverWait(thread_driver, 100).until(
                EC.element_to_be_clickable((By.ID, "edit-submit"))
            )
            search_button.click()
            time.sleep(13)
            
            WebDriverWait(thread_driver, 100).until(
                EC.presence_of_element_located((By.ID, "main-content"))
            )
            logger_thread.info(f"{context} 刷新后重新搜索成功")
        except Exception as refresh_e:
            logger_thread.error(f"❌{context} 刷新页面后重新搜索失败: {refresh_e}")
            
            # 底层重试机制：连续尝试多次深度重置
            retry_refresh_count = 0
            max_refresh_retries = 3
            reset_success = False
            
            while retry_refresh_count < max_refresh_retries:
                retry_refresh_count += 1
                try:
                    logger_thread.warning(f"⚠️ {context} 第{retry_refresh_count}次尝试深度重置")
                    
                    # 深度重置步骤1：刷新页面
                    thread_driver.refresh()
                    time.sleep(1)  # 更长的等待时间
                    
                    # 深度重置步骤2：等待页面完全加载
                    WebDriverWait(thread_driver, 100).until(
                        EC.presence_of_element_located((By.TAG_NAME, "body"))
                    )
                    
                    # 深度重置步骤3：检查页面是否正常
                    page_source = thread_driver.page_source
                    if "403 Forbidden" in page_source or "error" in page_source.lower():
                        logger_thread.warning(f"❌{context} 第{retry_refresh_count}次重置后页面仍有问题")
                        continue
                    
                    # 深度重置步骤4：重新设置搜索条件
                    logger_thread.info(f"{context} 第{retry_refresh_count}次深度重置页面正常")
                    # 验证搜索结果页面加载完成
                    WebDriverWait(thread_driver, 100).until(
                        EC.presence_of_element_located((By.ID, "main-content"))
                    )
                    
                    logger_thread.info(f"{context} 第{retry_refresh_count}次深度重置成功")
                    reset_success = True
                    break
                        
                except Exception as retry_e:
                    logger_thread.warning(f"❌ {context} 第{retry_refresh_count}次深度重置失败: {retry_e}")
                    if retry_refresh_count >= max_refresh_retries:
                        logger_thread.error(f"❌ {context} 所有深度重置尝试都失败")
                        break
                    time.sleep(1)  # 重试间隔
            
            # 检查深度重置是否成功
            if not reset_success:
                logger_thread.error(f"❌ {context} 深度重置完全失败，无法恢复正常搜索")
                return "DEEP_RESET_FAILED"
            else:
                logger_thread.info(f"✅ {context} 深度重置成功，返回搜索结果页面")
                return thread_driver.page_source
    
    return thread_driver.page_source


def click_search(conditions, url, logger_thread):
    global cycle_start_time
    if cycle_start_time is None:
        cycle_start_time = datetime.now()
        logger_thread.info(f"程序开始时间: {cycle_start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    # 检查是否已经过了1小时 - 已禁用休息机制
    # current_time = datetime.now()
    # elapsed_seconds = (current_time - cycle_start_time).total_seconds()
    # if elapsed_seconds >= 40 * 60:  # 3600秒 = 1小时
    #     # 随机休息10到20分钟
    #     sleep_minutes = random.randint(5, 10)
    #     sleep_seconds = sleep_minutes * 60
    #     logger_thread.info(f"开始休息 {sleep_minutes} 分钟...")
    #     time.sleep(sleep_seconds)
    #     # 重置周期开始时间
    #     cycle_start_time = datetime.now()
    #     logger_thread.info(f"休息结束时间: {cycle_start_time.strftime('%H:%M:%S')}")
    #     logger_thread.info("开始新的周期...")
    try:
        add_condition = WebDriverWait(driver, 100).until(
            EC.element_to_be_clickable((By.ID, "addTriplet"))
        )
        add_condition.click()
        # 等待下拉框元素加载完成并可交互
        for i in range(3):
            select_element = WebDriverWait(driver, 100).until(
                EC.element_to_be_clickable((By.NAME, f"fld-base-{str(i)}"))
            )
            select = Select(select_element)
            select.select_by_value("alltext")
            condition = WebDriverWait(driver, 100).until(
                EC.element_to_be_clickable((By.NAME, f"val-base-{str(i)}"))
            )
            condition.clear()
            if i < len(conditions):
                condition.send_keys(conditions[i])
        search_button = WebDriverWait(driver, 100).until(
            EC.element_to_be_clickable((By.ID, "edit-submit"))
        )
        search_button.click()
        time.sleep(13)
        try:
            WebDriverWait(driver, 100).until(
                EC.presence_of_element_located((By.ID, "main-content"))
            )
        except TimeoutException:
            logger_thread.error("❌ 等待main-content元素超时，尝试刷新页面")
            driver.refresh()
            time.sleep(1)
            try:
                WebDriverWait(driver, 100).until(
                    EC.presence_of_element_located((By.ID, "main-content"))
                )
                logger_thread.info("页面刷新后main-content元素加载成功")
            except TimeoutException:
                logger_thread.error("❌ 页面刷新后main-content元素仍然超时")
    except Exception as e:
        logger_thread.error(f"❌ 页面无响应，加载超时")
        # 尝试刷新页面并重新执行搜索
        try:
            driver.refresh()
            time.sleep(1)
            logger_thread.info("异常后尝试刷新页面")
            
            # 刷新后需要重新设置搜索条件并执行搜索
            add_condition = WebDriverWait(driver, 100).until(
                EC.element_to_be_clickable((By.ID, "addTriplet"))
            )
            add_condition.click()
            
            # 重新设置搜索条件
            for i in range(3):
                select_element = WebDriverWait(driver, 100).until(
                    EC.element_to_be_clickable((By.NAME, f"fld-base-{str(i)}"))
                )
                select = Select(select_element)
                select.select_by_value("alltext")
                condition = WebDriverWait(driver, 100).until(
                    EC.element_to_be_clickable((By.NAME, f"val-base-{str(i)}"))
                )
                condition.clear()
                if i < len(conditions):
                    condition.send_keys(conditions[i])
            
            # 重新点击搜索按钮
            search_button = WebDriverWait(driver, 100).until(
                EC.element_to_be_clickable((By.ID, "edit-submit"))
            )
            search_button.click()
            time.sleep(13)
            
            WebDriverWait(driver, 100).until(
                EC.presence_of_element_located((By.ID, "main-content"))
            )
            logger_thread.info("刷新后重新搜索成功")
        except Exception as refresh_e:
            logger_thread.error(f"❌ 刷新页面并重新搜索失败: {refresh_e}")
    return driver.page_source


def is_element_exist(html_source, xpath_expr):
    """
    判断元素是否存在
    :param html_source: 页面HTML源码字符串
    :param xpath_expr: 用于定位元素的XPath表达式
    :return: 存在返回True，不存在返回False
    """
    try:
        # 解析HTML源码
        tree = etree.HTML(html_source)

        # 使用XPath查找元素，返回所有匹配的元素列表
        elements = tree.xpath(xpath_expr)

        # 如果列表长度大于0，说明元素存在
        return len(elements) > 0

    except Exception as e:
        print(f"解析HTML或执行XPath时出错: {e}")
        return False


def crawl_data(file_name, pubname, date_range, climate, policy, uncertainty, logger_thread):
    # 更换通道登录时更换域名、 登录过期时更换cookie
    
    relation = 'and'
    max_results = 60
    base_url = f'{domain_url}/apps/news/results?sort=YMD_date%3AD&p=AWNB&hide_duplicates=2&t={urllib.parse.quote(pubname)}&maxresults={max_results}&f=advanced'
    
    # 让所有config中的driver都访问到2025对应的year_url页面进行初始化
    year_url = '&'.join([base_url, f'fld-nav-0=YMD_date', f'val-nav-0={urllib.parse.quote("-2025.9")}'])
    logger_thread.info(f'开始初始化所有driver到2025年页面: {year_url}')
    
    # 并行访问所有driver到2025年页面
    def init_driver_to_2025(config):
        """初始化单个driver到2025年页面"""
        if config['driver']:
            try:
                current_driver = config['driver']
                current_driver.get(year_url)
                time.sleep(13)  # 等待页面加载
                logger_thread.info(f'账号 {config["username"]} 的driver成功访问2025年页面')
                # # 清空日期范围
                # date_from_element = WebDriverWait(current_driver, 100).until(
                #     EC.element_to_be_clickable((By.ID, "edit-date-from--2"))
                # )
                # date_from_element.clear()
                # date_to_element = WebDriverWait(current_driver, 100).until(
                #     EC.element_to_be_clickable((By.ID, "edit-date-to--2"))
                # )
                # date_to_element.clear()
                # search_button = WebDriverWait(current_driver, 100).until(
                #     EC.element_to_be_clickable((By.ID, "search_nav-submit"))
                # )
                # search_button.click()
                # time.sleep(random.randint(5, 10))
                # logger_thread.info(f'账号 {config["username"]} 的driver成功访问搜索页面')
                return True
            except Exception as e:
                logger_thread.error(f'❌ 账号 {config["username"]} 的driver访问2025年页面失败: {e}')
                return False
        return False
    
    # 使用线程池并行初始化所有driver
    with concurrent.futures.ThreadPoolExecutor(max_workers=len(account_configs)) as executor:
        future_to_config = {executor.submit(init_driver_to_2025, config): config for config in account_configs}
        
        # 等待所有初始化完成
        for future in concurrent.futures.as_completed(future_to_config):
            config = future_to_config[future]
            try:
                result = future.result()
                if result:
                    logger_thread.info(f"✅ 账号 {config['username']} 的driver初始化成功")
                else:
                    logger_thread.error(f"❌ 账号 {config['username']} 的driver初始化失败")
            except Exception as exc:
                logger_thread.error(f"❌ 账号 {config['username']} 的driver初始化异常: {exc}")
    
    logger_thread.info('所有driver初始化完成')
    
    # urllib3.disable_warnings()
    start_time = datetime.now()

    climate_list = climate.split('/')
    policy_list = policy.split('/')
    uncertainty_list = uncertainty.split('/')
    
    # 第一步：处理三种二组合类型
    # 1. climate + policy 组合
    climate_policy_combinations = []
    for climate_i in climate_list:
        for policy_i in policy_list:
            climate_policy_combinations.append(('climate_policy', climate_i, policy_i))
    
    # 2. climate + uncertainty 组合
    climate_uncertainty_combinations = []
    for climate_i in climate_list:
        for uncertainty_i in uncertainty_list:
            climate_uncertainty_combinations.append(('climate_uncertainty', climate_i, uncertainty_i))
    
    # 3. policy + uncertainty 组合
    policy_uncertainty_combinations = []
    for policy_i in policy_list:
        for uncertainty_i in uncertainty_list:
            policy_uncertainty_combinations.append(('policy_uncertainty', policy_i, uncertainty_i))
    
    # 合并所有二组合
    all_two_combinations = climate_policy_combinations + climate_uncertainty_combinations + policy_uncertainty_combinations
    
    logger_thread.info(f'开始处理 {len(all_two_combinations)} 个二组合:')
    logger_thread.info(f'  - climate+policy: {len(climate_policy_combinations)} 个')
    logger_thread.info(f'  - climate+uncertainty: {len(climate_uncertainty_combinations)} 个')
    logger_thread.info(f'  - policy+uncertainty: {len(policy_uncertainty_combinations)} 个')
    
    # 检查已处理的二组合 - 支持三种类型
    two_combo_results_file = f'two_combo_results_{file_name.replace(".xlsx", ".txt")}'
    processed_two_combos = set()
    valid_two_combos = set()
    failed_two_combos = set()  # 新增：追踪失败的二组合
    
    if os.path.exists(two_combo_results_file):
        with open(two_combo_results_file, 'r', encoding='utf-8') as f:
            for line in f:
                if line.strip():
                    parts = [part.strip() for part in line.strip().split('|||')]
                    if len(parts) == 4:  # 新格式：type|||param1|||param2|||result_count
                        combo_type, param1, param2, result_count = parts
                        combo_key = (combo_type, param1, param2)
                        processed_two_combos.add(combo_key)
                        if int(result_count) > 0:
                            valid_two_combos.add(combo_key)
                        elif int(result_count) < 0:  # -1 或 -2 的失败情况
                            failed_two_combos.add(combo_key)
                    elif len(parts) == 3:  # 兼容旧格式：param1|||param2|||result_count (默认为climate_policy)
                        param1, param2, result_count = parts
                        combo_key = ('climate_policy', param1, param2)
                        processed_two_combos.add(combo_key)
                        if int(result_count) > 0:
                            valid_two_combos.add(combo_key)
                        elif int(result_count) < 0:
                            failed_two_combos.add(combo_key)
    
    # 处理未完成的二组合和需要重试的失败二组合
    remaining_two_combos = [combo for combo in all_two_combinations if combo not in processed_two_combos]
    retry_failed_combos = list(failed_two_combos)  # 失败的组合需要重试
    
    # 合并需要处理的组合
    all_pending_combos = remaining_two_combos + retry_failed_combos
    
    logger_thread.info(f'需要处理 {len(remaining_two_combos)} 个新的二组合')
    if retry_failed_combos:
        logger_thread.info(f'需要重试 {len(retry_failed_combos)} 个失败的二组合')
    logger_thread.info(f'总计需要处理 {len(all_pending_combos)} 个二组合')
    
    if all_pending_combos:
        
        # 创建一个队列存放可用的driver
        driver_queue = queue.Queue()
        for config in account_configs:
            if config['driver']:
                driver_queue.put(config['driver'])
        
        max_workers = min(driver_queue.qsize(), len(all_pending_combos))
        logger_thread.info(f'使用 {max_workers} 个线程处理二组合')
        
        # 创建线程安全的文件写入锁
        file_write_lock = threading.Lock()
        
        def process_two_combo(combo_data):
            """处理单个二组合的函数"""
            combo_type, param1, param2 = combo_data
            thread_id = threading.current_thread().ident
            
            # 检查是否是重试的失败组合
            is_retry = combo_data in retry_failed_combos
            
            # 从队列中获取一个driver
            try:
                current_driver = driver_queue.get(timeout=60)
            except queue.Empty:
                logger_thread.error(f"❌ [线程{thread_id}]  无法获取可用的driver处理二组合")
                return False
            try:
                # 根据获取到的driver找到对应的账号信息
                username = "未知账号"
                for config in account_configs:
                    if config['driver'] == current_driver:
                        username = config['username']
                        break

                # 重试机制配置
                max_retries = RETRY_CONFIG['max_retries']
                retry_count = 0
                success = False

                while retry_count < max_retries and not success:
                    try:
                        retry_info = f"(尝试 {retry_count + 1}/{max_retries})" if retry_count > 0 else ""
                        status_info = "重试失败" if is_retry else "新"
                        
                        # 根据组合类型构建搜索条件
                        if combo_type == 'climate_policy':
                            two_conditions = [param1, param2]  # climate, policy
                            combo_desc = f"{param1} + {param2}"
                        elif combo_type == 'climate_uncertainty':
                            two_conditions = [param1, param2]  # climate, uncertainty
                            combo_desc = f"{param1} + {param2}"
                        elif combo_type == 'policy_uncertainty':
                            two_conditions = [param1, param2]  # policy, uncertainty
                            combo_desc = f"{param1} + {param2}"
                        
                        logger_thread.info(f'[线程{thread_id}|账号{username}] 🔍 开始处理{status_info}二组合({combo_type})：{combo_desc} {retry_info}')
                        search_data = click_search_for_thread(two_conditions, base_url, logger_thread, current_driver, thread_id, username)

                        if search_data != "DEEP_RESET_FAILED":
                            result_count = int(get_total(search_data))
                            # 获取当前搜索URL
                            search_url = current_driver.current_url
                            result_line = f"{combo_type}|||{param1}|||{param2}|||{result_count}\n"

                            # 线程安全的文件写入
                            with file_write_lock:
                                # 如果是重试的组合，需要先清理旧的失败记录
                                if is_retry:
                                    # 读取现有文件内容
                                    existing_lines = []
                                    if os.path.exists(two_combo_results_file):
                                        with open(two_combo_results_file, 'r', encoding='utf-8') as f:
                                            existing_lines = f.readlines()

                                    # 过滤掉当前组合的旧记录
                                    filtered_lines = []
                                    for line in existing_lines:
                                        if line.strip():
                                            parts = [part.strip() for part in line.strip().split('|||')]
                                            if len(parts) == 4:  # 新格式
                                                old_type, old_param1, old_param2, old_result = parts
                                                if not (old_type == combo_type and old_param1 == param1 and old_param2 == param2):
                                                    filtered_lines.append(line)
                                            elif len(parts) == 3:  # 兼容旧格式
                                                old_param1, old_param2, old_result = parts
                                                if not (combo_type == 'climate_policy' and old_param1 == param1 and old_param2 == param2):
                                                    filtered_lines.append(line)

                                    # 重写文件，去掉旧的失败记录
                                    with open(two_combo_results_file, 'w', encoding='utf-8') as result_file:
                                        result_file.writelines(filtered_lines)
                                        result_file.write(result_line)
                                        result_file.flush()

                                    logger_thread.info(f'[线程{thread_id}|账号{username}] 🔄 已清理旧的失败记录并更新结果')
                                else:
                                    # 新组合直接追加
                                    with open(two_combo_results_file, 'a', encoding='utf-8') as result_file:
                                        result_file.write(result_line)
                                        result_file.flush()

                            if result_count > 0:
                                valid_two_combos.add(combo_data)
                                logger_thread.info(f'[线程{thread_id}|账号{username}] ✅ 二组合({combo_type}) {combo_desc} 有 {result_count} 条结果 📍 URL: {search_url}')
                            else:
                                logger_thread.info(f'[线程{thread_id}|账号{username}] ❌ 二组合({combo_type}) {combo_desc} 无结果 📍 URL: {search_url}')
                            
                            # 记录二组合搜索摘要到专门的摘要文件夹
                            if combo_type == 'climate_policy':
                                log_search_summary(param1, param2, "", result_count, search_url, file_name)
                            elif combo_type == 'climate_uncertainty':
                                log_search_summary(param1, "", param2, result_count, search_url, file_name)
                            elif combo_type == 'policy_uncertainty':
                                log_search_summary("", param1, param2, result_count, search_url, file_name)

                            success = True
                            return True
                        else:
                            # 搜索失败，进入重试逻辑
                            retry_count += 1
                            if retry_count < max_retries:
                                # 使用配置的重试延迟
                                retry_delay = random.randint(
                                    RETRY_CONFIG['base_retry_delay'],
                                    RETRY_CONFIG['max_retry_delay']
                                ) * retry_count  # 递增延迟
                                logger_thread.warning(f'⚠️ [线程{thread_id}|账号{username}]  二组合({combo_type}) {combo_desc} 搜索失败，{retry_delay}秒后重试 ({retry_count}/{max_retries})')
                                time.sleep(1)

                                # 尝试重建浏览器
                                if RETRY_CONFIG['enable_browser_recreate']:
                                    target_config = None
                                    for config in account_configs:
                                        if config['driver'] == current_driver:
                                            target_config = config
                                            break
                                    if target_config:
                                        recreate_success = recreate_driver_for_config(target_config, logger_thread)
                                        if recreate_success:
                                            current_driver = target_config['driver']
                                            logger_thread.info(
                                                f'[线程{thread_id}|账号{username}] ✅ 异常后浏览器重建成功')

                                            # 重建后需要重新访问基础页面
                                            year_url = '&'.join([base_url, f'fld-nav-0=YMD_date',
                                                                 f'val-nav-0={urllib.parse.quote("-2025.9")}'])
                                            current_driver.get(year_url)
                                            time.sleep(13)
                                        else:
                                            logger_thread.warning(
                                                f'⚠️ [线程{thread_id}|账号{username}] 异常后浏览器重建失败')
                                            raise Exception("浏览器重建失败")
                                    else:
                                        logger_thread.error(
                                            f'❌[线程{thread_id}|账号{username}]  无法找到对应的config，跳过当前组合')
                                        raise Exception("无法找到对应的config")
                            else:
                                # 达到最大重试次数，记录为-1
                                result_line = f"{combo_type}|||{param1}|||{param2}|||-1\n"
                                with file_write_lock:
                                    with open(two_combo_results_file, 'a', encoding='utf-8') as result_file:
                                        result_file.write(result_line)
                                        result_file.flush()
                                logger_thread.error(f' ❌ [线程{thread_id}|账号{username}] 二组合({combo_type}) {combo_desc} 重试{max_retries}次后仍然失败')
                                return False

                    except Exception as e:
                        retry_count += 1
                        if retry_count < max_retries:
                            # 异常重试延迟更长
                            retry_delay = random.randint(
                                RETRY_CONFIG['base_retry_delay'] * RETRY_CONFIG['exception_retry_delay_multiplier'],
                                RETRY_CONFIG['max_retry_delay'] * RETRY_CONFIG['exception_retry_delay_multiplier']
                            ) * retry_count
                            logger_thread.error(f'  ❌ [线程{thread_id}|账号{username}] 二组合({combo_type}) {combo_desc} 处理异常: {e}，{retry_delay}秒后重试 ({retry_count}/{max_retries})')
                            time.sleep(1)

                            # 异常后尝试重建浏览器
                            if RETRY_CONFIG['enable_browser_recreate']:
                                target_config = None
                                for config in account_configs:
                                    if config['driver'] == current_driver:
                                        target_config = config
                                        break

                                if target_config:
                                    recreate_success = recreate_driver_for_config(target_config, logger_thread)
                                    if recreate_success:
                                        current_driver = target_config['driver']
                                        logger_thread.info(f'[线程{thread_id}|账号{username}] ✅ 异常后浏览器重建成功')

                                        # 重建后需要重新访问基础页面
                                        year_url = '&'.join([base_url, f'fld-nav-0=YMD_date',
                                                             f'val-nav-0={urllib.parse.quote("-2025.9")}'])
                                        current_driver.get(year_url)
                                        time.sleep(13)
                                    else:
                                        logger_thread.warning(f'⚠️  [线程{thread_id}|账号{username}] 异常后浏览器重建失败')
                                        raise Exception("浏览器重建失败")
                                else:
                                    logger_thread.error(
                                        f'❌ [线程{thread_id}|账号{username}]  无法找到对应的config，跳过当前组合')
                                    raise Exception("无法找到对应的config")
                        else:
                            # 达到最大重试次数，记录为-2
                            result_line = f"{combo_type}|||{param1}|||{param2}|||-2\n"
                            with file_write_lock:
                                with open(two_combo_results_file, 'a', encoding='utf-8') as result_file:
                                    result_file.write(result_line)
                                    result_file.flush()
                            logger_thread.error(f'❌ [线程{thread_id}|账号{username}]  二组合({combo_type}) {combo_desc} 异常重试{max_retries}次后仍然失败: {e}')
                            return False

            # 将driver放回队列供其他任务使用
            except:
                logger_thread.warning(f'⚠️ [线程{thread_id}|账号{username}]  无法将driver放回队列')
                return False
            finally:
                driver_queue.put(current_driver)
        
        # 使用线程池处理所有二组合
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
            # 提交所有二组合任务
            future_to_combo = {
                executor.submit(process_two_combo, combo): combo 
                for combo in all_pending_combos
            }
            
            # 等待所有任务完成
            completed_count = 0
            for future in concurrent.futures.as_completed(future_to_combo):
                combo = future_to_combo[future]
                try:
                    result = future.result()
                    completed_count += 1
                    if result:
                        logger_thread.info(f"✅ 二组合 {combo} 处理成功 ({completed_count}/{len(all_pending_combos)})")
                    else:
                        logger_thread.error(f"❌ 二组合 {combo} 处理失败 ({completed_count}/{len(all_pending_combos)})")
                except Exception as exc:
                    completed_count += 1
                    logger_thread.error(f"❌ 二组合 {combo} 处理异常: {exc} ({completed_count}/{len(all_pending_combos)})")
        
        logger_thread.info(f'多线程二组合处理完成，共处理 {len(all_pending_combos)} 个二组合')
        
        # 重新读取文件获取最新的有效二组合
        if os.path.exists(two_combo_results_file):
            with open(two_combo_results_file, 'r', encoding='utf-8') as f:
                for line in f:
                    if line.strip():
                        parts = [part.strip() for part in line.strip().split('|||')]
                        if len(parts) == 4:  # 新格式
                            combo_type, param1, param2, result_count = parts
                            if int(result_count) > 0:
                                valid_two_combos.add((combo_type, param1, param2))
                        elif len(parts) == 3:  # 兼容旧格式
                            param1, param2, result_count = parts
                            if int(result_count) > 0:
                                valid_two_combos.add(('climate_policy', param1, param2))
    else:
        logger_thread.info('所有二组合已处理完成')
    
    logger_thread.info(f'二组合处理完成，有效组合数: {len(valid_two_combos)}')
    
    # 第二步：基于有效的二组合生成三组合
    all_combinations = []
    
    # 统计各类型的有效二组合
    climate_policy_valid = set()
    climate_uncertainty_valid = set()
    policy_uncertainty_valid = set()
    
    for combo_data in valid_two_combos:
        combo_type, param1, param2 = combo_data
        if combo_type == 'climate_policy':
            climate_policy_valid.add((param1, param2))
        elif combo_type == 'climate_uncertainty':
            climate_uncertainty_valid.add((param1, param2))
        elif combo_type == 'policy_uncertainty':
            policy_uncertainty_valid.add((param1, param2))
    
    logger_thread.info(f'有效二组合统计:')
    logger_thread.info(f'  - climate+policy: {len(climate_policy_valid)} 个')
    logger_thread.info(f'  - climate+uncertainty: {len(climate_uncertainty_valid)} 个')
    logger_thread.info(f'  - policy+uncertainty: {len(policy_uncertainty_valid)} 个')
    
    # 生成所有可能的三组合，但只保留至少有一个有效二组合的三组合
    all_combinations = []
    
    for climate_i in climate_list:
        for policy_i in policy_list:
            for uncertainty_i in uncertainty_list:
                # 检查这个三组合的三个二组合是否都有效
                all_two_combos_valid = (
                    (climate_i, policy_i) in climate_policy_valid and
                    (climate_i, uncertainty_i) in climate_uncertainty_valid and
                    (policy_i, uncertainty_i) in policy_uncertainty_valid
                )
                
                if all_two_combos_valid:
                    all_combinations.append((climate_i, policy_i, uncertainty_i))
    
    logger_thread.info(f'基于有效二组合筛选，生成了 {len(all_combinations)} 个三组合')
    
    rounds = len(all_combinations)
    logger_thread.info(f'基于有效二组合，总共需要处理 {rounds} 个三组合')
    
    # 检查已完成的三组合条件
    last_conditions = read_as_set(filename='condition.txt')
    completed_combinations = set()
    for last_condition in last_conditions:
        parts = last_condition.split('|||')
        if len(parts) == 3:
            completed_combinations.add(tuple(parts))
    
    # 过滤掉已完成的三组合
    remaining_combinations = [combo for combo in all_combinations if combo not in completed_combinations]
    logger_thread.info(f'剩余需要处理 {len(remaining_combinations)} 个三组合')
    
    # 如果没有剩余组合需要处理，直接进行年份统计
    if len(remaining_combinations) == 0:
        logger_thread.info('所有组合已完成，直接进行年份统计')
    else:
        # 使用线程池处理所有组合，为每个线程分配专用的driver
        # 创建一个队列存放可用的driver
        driver_queue = queue.Queue()
        for config in account_configs:
            if config['driver']:
                driver_queue.put(config['driver'])
    
        max_workers = min(driver_queue.qsize(), len(remaining_combinations))
        logger_thread.info(f'使用 {max_workers} 个线程处理组合')
        
        def process_combination_with_driver(combination_data):
            """使用专用driver处理单个组合的函数"""
            climate_i, policy_i, uncertainty_i = combination_data
            thread_id = threading.current_thread().ident
            
            # 从队列中获取一个driver
            try:
                current_driver = driver_queue.get(timeout=60)
            except queue.Empty:
                logger_thread.error(f"❌ [线程{thread_id}|未知账号]  无法获取可用的driver")
                return False

            # 根据获取到的driver找到对应的账号信息
            username = "未知账号"
            for config in account_configs:
                if config['driver'] == current_driver:
                    username = config['username']
                    break
            # 线程间随机延迟，避免同时请求
            thread_delay = random.randint(1, 5)

            try:
                logger_thread.info(f'[线程{thread_id}|账号{username}] 🚀 获取driver延迟{thread_delay}秒 🔍 开始处理组合：{climate_i} + {policy_i} + {uncertainty_i}')
                # 构建搜索条件，添加重试机制
                conditions = [climate_i, policy_i, uncertainty_i]
                retry_count = 0
                max_retries = 3
                
                while retry_count < max_retries:
                    try:
                        # 重连机制：先检查driver健康状态
                        if not check_driver_health(current_driver, logger_thread, f"[线程{thread_id}|账号{username}]"):
                            logger_thread.warning(f'⚠️ [线程{thread_id}|账号{username}] 🔧 检测到浏览器不健康，尝试重建浏览器')
                            
                            # 找到对应的config并重建driver
                            target_config = None
                            for config in account_configs:
                                if config['driver'] == current_driver:
                                    target_config = config
                                    break
                            
                            if target_config:
                                recreate_success = recreate_driver_for_config(target_config, logger_thread)
                                if recreate_success:
                                    current_driver = target_config['driver']
                                    logger_thread.info(f'[线程{thread_id}|账号{username}] ✅ 浏览器重建成功，继续处理组合：{climate_i}+{policy_i}+{uncertainty_i}')
                                    
                                    # 重建后需要重新访问基础页面
                                    year_url = '&'.join([base_url, f'fld-nav-0=YMD_date', f'val-nav-0={urllib.parse.quote("-2025.9")}'])
                                    current_driver.get(year_url)
                                    time.sleep(13)
                                else:
                                    logger_thread.error(f'❌ [线程{thread_id}|账号{username}]  浏览器重建失败，跳过当前组合')
                                    raise Exception("浏览器重建失败")
                            else:
                                logger_thread.error(f'❌ [线程{thread_id}|账号{username}]  无法找到对应的config，跳过当前组合')
                                raise Exception("无法找到对应的config")
                        
                        uncertainty_data = click_search_for_thread(conditions, base_url, logger_thread, current_driver, thread_id, username)
                        # 检查是否返回深度重置失败标记
                        if uncertainty_data == "DEEP_RESET_FAILED":
                            logger_thread.error(f'❌ [线程{thread_id}|账号{username}]  深度重置失败，尝试重建浏览器')
                            
                            # 找到对应的config并重建driver（最后一次尝试）
                            target_config = None
                            for config in account_configs:
                                if config['driver'] == current_driver:
                                    target_config = config
                                    break
                            
                            if target_config:
                                recreate_success = recreate_driver_for_config(target_config, logger_thread)
                                if recreate_success:
                                    current_driver = target_config['driver']
                                    logger_thread.info(f'[线程{thread_id}|账号{username}] ✅ 深度重置失败后浏览器重建成功，最后一次尝试组合：{climate_i}+{policy_i}+{uncertainty_i}')
                                    
                                    # 重建后需要重新访问基础页面
                                    year_url = '&'.join([base_url, f'fld-nav-0=YMD_date', f'val-nav-0={urllib.parse.quote("-2025.9")}'])
                                    current_driver.get(year_url)
                                    time.sleep(13)
                                    # 重新尝试搜索
                                    uncertainty_data = click_search_for_thread(conditions, base_url, logger_thread, current_driver, thread_id, username)
                                    if uncertainty_data == "DEEP_RESET_FAILED":
                                        logger_thread.error(f'❌ [线程{thread_id}|账号{username}] 重建后仍然失败，记录失败条件')
                                        failed_condition = '|||'.join([climate_i, policy_i, uncertainty_i])
                                        append_to_file(failed_condition, mode='a', filename='failed_conditions.txt')
                                        logger_thread.error(f'❌ [线程{thread_id}|账号{username}] 已记录失败条件到failed_conditions.txt: {failed_condition}')
                                        return False
                                else:
                                    logger_thread.error(f'❌ [线程{thread_id}|账号{username}] ❌ 深度重置失败后浏览器重建也失败，记录失败条件')
                                    failed_condition = '|||'.join([climate_i, policy_i, uncertainty_i])
                                    append_to_file(failed_condition, mode='a', filename='failed_conditions.txt')
                                    logger_thread.error(f'❌ [线程{thread_id}|账号{username}] 已记录失败条件到failed_conditions.txt: {failed_condition}')
                                    return False
                            else:
                                logger_thread.error(f'❌ [线程{thread_id}|账号{username}] ❌ 深度重置失败后无法找到config，记录失败条件')
                                failed_condition = '|||'.join([climate_i, policy_i, uncertainty_i])
                                append_to_file(failed_condition, mode='a', filename='failed_conditions.txt')
                                logger_thread.error(f'❌ [线程{thread_id}|账号{username}] 已记录失败条件到failed_conditions.txt: {failed_condition}')
                                return False
                        
                        uncertainty_total = int(get_total(uncertainty_data))
                        # 记录搜索后的URL
                        search_url = current_driver.current_url
                        break  # 成功则跳出重试循环
                    except Exception as e:
                        retry_count += 1
                        logger_thread.warning(f'⚠️ [线程{thread_id}|账号{username}]  组合：{climate_i}+{policy_i}+{uncertainty_i} 第{retry_count}次重试失败: {e}')
                        
                        # 检查是否是连接错误，如果是则尝试重建浏览器
                        if "Connection refused" in str(e) or "HTTPConnectionPool" in str(e):
                            logger_thread.warning(f'⚠️ [线程{thread_id}|账号{username}] 🔧 检测到连接错误，尝试重建浏览器')
                            
                            # 找到对应的config并重建driver
                            target_config = None
                            for config in account_configs:
                                if config['driver'] == current_driver:
                                    target_config = config
                                    break
                            
                            if target_config:
                                recreate_success = recreate_driver_for_config(target_config, logger_thread)
                                if recreate_success:
                                    current_driver = target_config['driver']
                                    logger_thread.info(f'[线程{thread_id}|账号{username}] ✅ 连接错误后浏览器重建成功，继续重试')
                                    
                                    # 重建后需要重新访问基础页面
                                    year_url = '&'.join([base_url, f'fld-nav-0=YMD_date', f'val-nav-0={urllib.parse.quote("-2025.9")}'])
                                    current_driver.get(year_url)
                                    time.sleep(13)
                                    
                                    # 重建成功后，减少重试计数，给重建后的浏览器一次机会
                                    retry_count -= 1
                                else:
                                    logger_thread.error(f'❌ [线程{thread_id}|账号{username}] ❌ 连接错误后浏览器重建失败')
                        else:
                            # 对于非连接错误，尝试刷新页面
                            try:
                                current_driver.refresh()
                                logger_thread.info(f'[线程{thread_id}|账号{username}] 🔄 页面刷新完成')
                            except:
                                logger_thread.warning(f' ⚠️ [线程{thread_id}|账号{username}]  页面刷新失败')
                        
                        if retry_count >= max_retries:
                            logger_thread.error(f'❌ [线程{thread_id}|账号{username}]  达到最大重试次数，放弃处理该组合')
                            raise e  # 达到最大重试次数，抛出异常
                        time.sleep(1)  # 递增等待时间
                
                # 合并搜索完成和开始处理的日志
                total_pages = (int(uncertainty_total) - 1) // 60 + 1 if uncertainty_total > 0 else 0
                
                if uncertainty_total == 0:
                    logger_thread.info(f'[线程{thread_id}|账号{username}] 📊 搜索完成 - {climate_i}+{policy_i}+{uncertainty_i} = {uncertainty_total} 条 📝 无结果已记录 📍 URL: {search_url}')
                    log_search_summary(climate_i, policy_i, uncertainty_i, uncertainty_total, search_url, file_name)
                    condition = '|||'.join([climate_i, policy_i, uncertainty_i])
                    append_to_file(condition, mode='a', filename='condition.txt')
                    return True
                else:
                    logger_thread.info(f'[线程{thread_id}|账号{username}] 📊 搜索完成 - {climate_i}+{policy_i}+{uncertainty_i} = {uncertainty_total} 条 📄 开始处理数据(预计{total_pages}页) 🔍 处理第1页 📍 URL: {search_url}')
                    log_search_summary(climate_i, policy_i, uncertainty_i, uncertainty_total, search_url, file_name)
                    
                # 处理第一页数据
                request_and_map(uncertainty_data)
                
                # 处理分页数据
                max_results = 60
                page = 0
                while (page + 1) * max_results < int(uncertainty_total):
                    page += 1
                    try:
                        # 检查driver健康状态
                        if not check_driver_health(current_driver, logger_thread, f"[线程{thread_id}|账号{username}]"):
                            logger_thread.warning(f'⚠️ [线程{thread_id}|账号{username}] 🔧 分页处理时检测到浏览器不健康，尝试重建浏览器')
                            
                            # 找到对应的config并重建driver
                            target_config = None
                            for config in account_configs:
                                if config['driver'] == current_driver:
                                    target_config = config
                                    break
                            
                            if target_config:
                                recreate_success = recreate_driver_for_config(target_config, logger_thread)
                                if recreate_success:
                                    current_driver = target_config['driver']
                                    logger_thread.info(f'[线程{thread_id}|账号{username}] ✅ 分页处理时浏览器重建成功')
                                    
                                    # 重建后需要重新执行搜索并导航到当前页
                                    year_url = '&'.join([base_url, f'fld-nav-0=YMD_date', f'val-nav-0={urllib.parse.quote("-2025.9")}'])
                                    current_driver.get(year_url)
                                    time.sleep(13)
                                    # 重新执行搜索
                                    uncertainty_data = click_search_for_thread(conditions, base_url, logger_thread, current_driver, thread_id, username)
                                    
                                    # 如果需要，导航到指定页面
                                    for nav_page in range(1, page + 1):
                                        if nav_page > 1:
                                            page_data = current_driver.page_source
                                            is_page_exist = is_element_exist(page_data, f'//a[@title="Go to page {nav_page}"]')
                                            if is_page_exist:
                                                page_link = WebDriverWait(current_driver, 100).until(
                                                    EC.element_to_be_clickable((By.XPATH, f'//a[@title="Go to page {nav_page}"]'))
                                                )
                                                current_driver.execute_script("arguments[0].scrollIntoView();", page_link)
                                                page_link.click()
                                                time.sleep(13)
                                                WebDriverWait(current_driver, 100).until(
                                                    EC.presence_of_element_located((By.ID, "main-content"))
                                                )
                                    logger_thread.info(f'[线程{thread_id}|账号{username}] ✅ 重建后已导航到第{page}页')
                                else:
                                    logger_thread.error(f'❌ [线程{thread_id}|账号{username}]  分页处理时浏览器重建失败，结束分页处理')
                                    break
                            else:
                                logger_thread.error(f'❌ [线程{thread_id}|账号{username}]  分页处理时无法找到对应的config，结束分页处理')
                                break
                        
                        page_data = current_driver.page_source
                        is_page_exist = is_element_exist(page_data, f'//a[@title="Go to page {page + 1}"]')
                        if is_page_exist:
                            page_link = WebDriverWait(current_driver, 100).until(
                                EC.element_to_be_clickable((By.XPATH, f'//a[@title="Go to page {page + 1}"]'))
                            )
                            current_driver.execute_script("arguments[0].scrollIntoView();", page_link)
                            page_link.click()
                            time.sleep(13)
                            WebDriverWait(current_driver, 100).until(
                                EC.presence_of_element_located((By.ID, "main-content"))
                            )
                            page_data = current_driver.page_source
                            request_and_map(page_data)
                            logger_thread.info(f'[线程{thread_id}|账号{username}] 🔍 第{page + 1}页/{total_pages}页 ✅ 处理完成')
                        else:
                            logger_thread.warning(f'⚠️ [线程{thread_id}|账号{username}] ⚠️ 第{page + 1}页不存在，结束分页')
                            break
                    except Exception as page_e:
                        logger_thread.error(f'❌ [线程{thread_id}|账号{username}]  处理第{page + 1}页时发生错误: {page_e}')
                        
                        # 检查是否是连接错误
                        if "Connection refused" in str(page_e) or "HTTPConnectionPool" in str(page_e):
                            logger_thread.warning(f'⚠️ [线程{thread_id}|账号{username}] 🔧 分页处理连接错误，尝试重建浏览器')
                            
                            # 找到对应的config并重建driver
                            target_config = None
                            for config in account_configs:
                                if config['driver'] == current_driver:
                                    target_config = config
                                    break
                            
                            if target_config and recreate_driver_for_config(target_config, logger_thread):
                                current_driver = target_config['driver']
                                logger_thread.info(f'[线程{thread_id}|账号{username}] ✅ 分页连接错误后浏览器重建成功，跳过当前页')
                                continue
                        
                        logger_thread.warning(f'⚠️ [线程{thread_id}|账号{username}]  跳过第{page + 1}页，继续处理下一页')
                        continue
                
                # 标记组合已完成
                condition = '|||'.join([climate_i, policy_i, uncertainty_i])
                append_to_file(condition, mode='a', filename='condition.txt')
                
                logger_thread.info(f'[线程{thread_id}|账号{username}] ✅ 完成：{climate_i}+{policy_i}+{uncertainty_i}')
                return True
                
            except Exception as e:
                logger_thread.error(f'❌  [线程{thread_id}|账号{username}] 组合处理失败：{climate_i}+{policy_i}+{uncertainty_i} - 错误: {e}')
                return False
            finally:
                # 将driver放回队列供其他任务使用
                driver_queue.put(current_driver)
        
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
            # 提交所有组合任务
            future_to_combination = {
                executor.submit(process_combination_with_driver, combo): combo 
                for combo in remaining_combinations
            }
            
            # 等待所有任务完成
            completed_count = 0
            for future in concurrent.futures.as_completed(future_to_combination):
                combination = future_to_combination[future]
                try:
                    result = future.result()
                    completed_count += 1
                    if result:
                        logger_thread.info(f"✅ 组合 {combination} 处理成功 ({completed_count}/{len(remaining_combinations)})")
                    else:
                        logger_thread.error(f"❌ 组合 {combination} 处理失败 ({completed_count}/{len(remaining_combinations)})")
                except Exception as exc:
                    completed_count += 1
                    logger_thread.error(f"❌ 组合 {combination} 处理异常: {exc} ({completed_count}/{len(remaining_combinations)})")
        
        logger_thread.info(f'所有组合处理完成，共处理 {len(remaining_combinations)} 个组合')

    # 处理失败的组合
    failed_conditions_file = 'failed_conditions.txt'
    if os.path.exists(failed_conditions_file):
        failed_conditions = read_as_set(filename=failed_conditions_file)
        if failed_conditions:
            logger_thread.info(f'🔄 发现 {len(failed_conditions)} 个失败的三组合条件，开始重新处理...')
            
            # 解析失败条件并重新处理
            failed_combinations = []
            for failed_condition in failed_conditions:
                if failed_condition.strip():
                    # 移除分隔符，解析组合条件
                    clean_condition = failed_condition.replace('---QUERY_RESULT_END---', '').strip()
                    parts = clean_condition.split('|||')
                    if len(parts) == 3:
                        climate_i, policy_i, uncertainty_i = parts
                        failed_combinations.append((climate_i.strip(), policy_i.strip(), uncertainty_i.strip()))
                        logger_thread.info(f'🔄 准备重新处理失败组合: {climate_i} + {policy_i} + {uncertainty_i}')
            
            if failed_combinations:
                logger_thread.info(f'🔄 开始重新处理 {len(failed_combinations)} 个失败组合')
                
                # 创建driver队列用于处理失败组合
                driver_queue_failed = queue.Queue()
                for config in account_configs:
                    if config['driver']:
                        driver_queue_failed.put(config['driver'])
                
                max_workers_failed = min(driver_queue_failed.qsize(), len(failed_combinations))
                logger_thread.info(f'🔄 使用 {max_workers_failed} 个线程重新处理失败组合')
                
                def process_failed_combination(combination_data):
                    """处理单个失败组合的函数"""
                    climate_i, policy_i, uncertainty_i = combination_data
                    thread_id = threading.current_thread().ident
                    
                    # 从队列中获取一个driver
                    try:
                        current_driver = driver_queue_failed.get(timeout=60)
                    except queue.Empty:
                        logger_thread.error(f"❌ [线程{thread_id}|失败重试] 无法获取可用的driver")
                        return False

                    # 根据获取到的driver找到对应的账号信息
                    username = "未知账号"
                    for config in account_configs:
                        if config['driver'] == current_driver:
                            username = config['username']
                            break
                    
                    try:
                        logger_thread.info(f'[线程{thread_id}|账号{username}] 🔄 重新处理失败组合：{climate_i} + {policy_i} + {uncertainty_i}')
                        
                        # 构建搜索条件
                        conditions = [climate_i, policy_i, uncertainty_i]
                        uncertainty_data = click_search_for_thread(conditions, base_url, logger_thread, current_driver, thread_id, username)
                        
                        if uncertainty_data != "DEEP_RESET_FAILED":
                            uncertainty_total = int(get_total(uncertainty_data))
                            search_url = current_driver.current_url
                            
                            if uncertainty_total == 0:
                                logger_thread.info(f'[线程{thread_id}|账号{username}] 📊 失败组合重试 - {climate_i}+{policy_i}+{uncertainty_i} = {uncertainty_total} 条 📝 无结果')
                                log_search_summary(climate_i, policy_i, uncertainty_i, uncertainty_total, search_url, file_name)
                                condition = '|||'.join([climate_i, policy_i, uncertainty_i])
                                append_to_file(condition, mode='a', filename='condition.txt')
                                return True
                            else:
                                total_pages = (int(uncertainty_total) - 1) // 60 + 1 if uncertainty_total > 0 else 0
                                logger_thread.info(f'[线程{thread_id}|账号{username}] 📊 失败组合重试 - {climate_i}+{policy_i}+{uncertainty_i} = {uncertainty_total} 条 📄 开始处理数据(预计{total_pages}页)')
                                log_search_summary(climate_i, policy_i, uncertainty_i, uncertainty_total, search_url, file_name)
                                
                                # 处理第一页数据
                                request_and_map(uncertainty_data)
                                
                                # 处理分页数据
                                max_results = 60
                                page = 0
                                while (page + 1) * max_results < int(uncertainty_total):
                                    page += 1
                                    try:
                                        page_data = current_driver.page_source
                                        is_page_exist = is_element_exist(page_data, f'//a[@title="Go to page {page + 1}"]')
                                        if is_page_exist:
                                            page_link = WebDriverWait(current_driver, 100).until(
                                                EC.element_to_be_clickable((By.XPATH, f'//a[@title="Go to page {page + 1}"]'))
                                            )
                                            current_driver.execute_script("arguments[0].scrollIntoView();", page_link)
                                            page_link.click()
                                            time.sleep(13)
                                            WebDriverWait(current_driver, 100).until(
                                                EC.presence_of_element_located((By.ID, "main-content"))
                                            )
                                            page_data = current_driver.page_source
                                            request_and_map(page_data)
                                            logger_thread.info(f'[线程{thread_id}|账号{username}] 🔄 失败组合重试 - 第{page + 1}页/{total_pages}页 ✅ 处理完成')
                                        else:
                                            logger_thread.warning(f'⚠️ [线程{thread_id}|账号{username}] 失败组合重试 - 第{page + 1}页不存在，结束分页')
                                            break
                                    except Exception as page_e:
                                        logger_thread.error(f'❌ [线程{thread_id}|账号{username}] 失败组合重试 - 处理第{page + 1}页时发生错误: {page_e}')
                                        continue
                                
                                # 标记组合已完成
                                condition = '|||'.join([climate_i, policy_i, uncertainty_i])
                                append_to_file(condition, mode='a', filename='condition.txt')
                                
                                logger_thread.info(f'[线程{thread_id}|账号{username}] ✅ 失败组合重试完成：{climate_i}+{policy_i}+{uncertainty_i}')
                                return True
                        else:
                            logger_thread.error(f'❌ [线程{thread_id}|账号{username}] 失败组合重试 - 深度重置失败：{climate_i}+{policy_i}+{uncertainty_i}')
                            return False
                            
                    except Exception as e:
                        logger_thread.error(f'❌ [线程{thread_id}|账号{username}] 失败组合重试异常：{climate_i}+{policy_i}+{uncertainty_i} - 错误: {e}')
                        return False
                    finally:
                        # 将driver放回队列供其他任务使用
                        driver_queue_failed.put(current_driver)
                
                # 使用线程池处理所有失败组合
                with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers_failed) as executor:
                    # 提交所有失败组合任务
                    future_to_failed = {
                        executor.submit(process_failed_combination, combo): combo 
                        for combo in failed_combinations
                    }
                    
                    # 等待所有任务完成
                    completed_failed_count = 0
                    for future in concurrent.futures.as_completed(future_to_failed):
                        combination = future_to_failed[future]
                        try:
                            result = future.result()
                            completed_failed_count += 1
                            if result:
                                logger_thread.info(f"✅ 失败组合重试 {combination} 处理成功 ({completed_failed_count}/{len(failed_combinations)})")
                            else:
                                logger_thread.error(f"❌ 失败组合重试 {combination} 处理失败 ({completed_failed_count}/{len(failed_combinations)})")
                        except Exception as exc:
                            completed_failed_count += 1
                            logger_thread.error(f"❌ 失败组合重试 {combination} 处理异常: {exc} ({completed_failed_count}/{len(failed_combinations)})")
                
                logger_thread.info(f'🔄 失败组合重试完成，共处理 {len(failed_combinations)} 个失败组合')
                
                # 处理完成后清空失败条件文件
                try:
                    with open(failed_conditions_file, 'w', encoding='utf-8') as f:
                        f.write('')  # 清空文件
                    logger_thread.info(f'✅ 已清空失败条件文件: {failed_conditions_file}')
                except Exception as e:
                    logger_thread.error(f'❌ 清空失败条件文件失败: {e}')
        else:
            logger_thread.info('✅ 没有需要重新处理的失败组合')
    else:
        logger_thread.info('✅ 没有失败条件文件，无需重新处理')

    # 年份统计部分 - 使用多线程并行处理
    all_map = dict()
    start_year, end_year = map(int, date_range.split('-'))
    
    # 获取可用的drivers
    available_drivers = []
    for config in account_configs:
        if config['driver']:
            available_drivers.append(config['driver'])
    
    if not available_drivers:
        logger_thread.error("❌ 没有可用的driver进行年份统计")
        return
    
    # 创建driver队列用于多线程
    driver_queue_stats = queue.Queue()
    for driver in available_drivers:
        driver_queue_stats.put(driver)
    
    def process_year_stats(year):
        """处理单个年份的统计"""
        current_driver = None
        try:
            # 从队列获取driver
            current_driver = driver_queue_stats.get(timeout=30)
            # 通过id定位元素设置年份
            date_from_element = WebDriverWait(current_driver, 100).until(
                EC.element_to_be_clickable((By.ID, "edit-date-from--2"))
            )
            date_from_element.clear()
            date_from_element.send_keys(str(year))
            date_to_element = WebDriverWait(current_driver, 100).until(
                EC.element_to_be_clickable((By.ID, "edit-date-to--2"))
            )
            date_to_element.clear()
            # 如果是2025年，设置为2025.9，否则使用年份
            if year == 2025:
                date_to_element.send_keys("2025.9")
            else:
                date_to_element.send_keys(str(year))
            search_button = WebDriverWait(current_driver, 100).until(
                EC.element_to_be_clickable((By.ID, "search_nav-submit"))
            )
            search_button.click()
            time.sleep(13)
            current_url = current_driver.current_url
            logger_thread.info(f'[统计线程|年份统计] 📍 {year}年统计URL: {current_url}')
            WebDriverWait(current_driver, 100).until(
                EC.presence_of_element_located((By.ID, "main-content"))
            )
            year_data = current_driver.page_source
            year_total = get_total(year_data)
            logger_thread.info(f'[统计线程|年份统计] 📈 {year}年统计完成，总数: {year_total} 条')
            return year, year_total
        except Exception as e:
            logger_thread.error(f'❌ [统计线程|年份统计]  {year}年统计失败: {e}')
            return year, "0"
        finally:
            # 将driver放回队列
            if current_driver:
                driver_queue_stats.put(current_driver)

    # 使用线程池并行初始化所有driver
    with concurrent.futures.ThreadPoolExecutor(max_workers=len(available_drivers)) as executor:
        future_to_config = {executor.submit(init_driver_to_2025, config): config for config in account_configs}
        # 等待所有初始化完成
        for future in concurrent.futures.as_completed(future_to_config):
            config = future_to_config[future]
            try:
                result = future.result()
                if result:
                    logger_thread.info(f"✅ 账号 {config['username']} 查询年的driver初始化成功")
                else:
                    logger_thread.error(f"❌ 账号 {config['username']} 查询年的driver初始化失败")
            except Exception as exc:
                logger_thread.error(f"❌ 账号 {config['username']} 查询年的driver初始化异常: {exc}")

    # 使用多线程处理年份统计
    with concurrent.futures.ThreadPoolExecutor(max_workers=min(len(available_drivers), end_year - start_year + 1)) as executor:
        # 提交所有年份统计任务
        year_futures = {
            executor.submit(process_year_stats, year): year 
            for year in range(start_year, end_year + 1)
        }
        
        # 等待所有年份统计完成
        for future in concurrent.futures.as_completed(year_futures):
            year = year_futures[future]
            try:
                result_year, result_total = future.result()
                all_map[str(result_year)] = str(result_total)
                logger_thread.info(f'[统计线程|年份统计] ✅ {result_year}年统计任务完成')
            except Exception as exc:
                logger_thread.error(f'❌ [统计线程|年份统计]  {year}年统计任务异常: {exc}')
                all_map[str(year)] = "0"
    
    logger_thread.info(f'[统计线程|年份统计] 🎉 所有年份统计完成: {all_map}')

    # 将map写入表格
    filter_map = dict()
    filter_data_set = read_as_set(f"task_data/{file_name.replace('.xlsx', '')}/results.txt")
    for filter_data in filter_data_set:
        parts = filter_data.split('|||')
        if len(parts) == 3:
            date, title, preview = parts
            date_year = [part for part in date.split(' ') if part.isdigit()][0]
            if date_year not in filter_map:
                filter_map[date_year] = set()
            filter_map[date_year].add(filter_data)
    write_csv_detail(file_name, all_map, filter_map)
    append_to_file('', mode='w')
    append_to_file('', mode='w', filename='condition.txt')
    
    # 检查是否有失败的条件需要重新处理
    failed_conditions_file = 'failed_conditions.txt'
    if os.path.exists(failed_conditions_file):
        failed_conditions = read_as_set(filename=failed_conditions_file)
        if failed_conditions:
            logger_thread.warning(f'⚠️ 发现 {len(failed_conditions)} 个失败的三组合条件，建议稍后重新处理: {failed_conditions_file}')
        else:
            logger_thread.info('所有搜索条件都已成功处理完成')
    else:
        logger_thread.info('所有搜索条件都已成功处理完成')
    
    # 统计二组合结果
    two_combo_results_file = f'two_combo_results_{file_name.replace(".xlsx", ".txt")}'
    if os.path.exists(two_combo_results_file):
        total_two_combos = 0
        valid_two_combos_count = 0
        failed_two_combos = 0
        with open(two_combo_results_file, 'r', encoding='utf-8') as f:
            for line in f:
                if line.strip():
                    parts = [part.strip() for part in line.strip().split('|||')]
                    if len(parts) == 3:
                        total_two_combos += 1
                        result_count = int(parts[2])
                        if result_count > 0:
                            valid_two_combos_count += 1
                        elif result_count < 0:
                            failed_two_combos += 1
        
        logger_thread.info(f'📊 二组合统计 - 总数: {total_two_combos}, 有效: {valid_two_combos_count}, 无结果: {total_two_combos - valid_two_combos_count - failed_two_combos}, 失败: {failed_two_combos}')
        logger_thread.info(f'📂 二组合结果文件: {two_combo_results_file}')
    
    logger_thread.info(f'线程id:{threading.current_thread().ident}+{file_name}== {file_name}== 开始时间:{start_time} - 结束时间: {datetime.now()}')


def log_search_summary(climate, policy, uncertainty, total_count, search_url, task_filename):
    """记录简化的搜索摘要到专门的摘要文件夹"""
    try:
        # 生成摘要文件名
        summary_filename = task_filename.replace('.xlsx', '_summary.txt')
        summary_path = f'search_summary_logs/{summary_filename}'

        # 创建简化的搜索记录
        if uncertainty:  # 三组合
            search_condition = f"{climate}+{policy}+{uncertainty}"
        else:  # 二组合
            search_condition = f"{climate}+{policy}"
        summary_line = f"条件: {search_condition} | 结果: {total_count} 条 | URL: {search_url}\n"

        # 追加到摘要文件
        with open(summary_path, 'a', encoding='utf-8') as f:
            f.write(summary_line)
    except Exception as e:
        print(f"记录搜索摘要失败: {e}")


def create_task_logger(task_filename):
    """为特定任务创建日志记录器"""
    # 从Excel文件名生成日志文件名（去掉.xlsx后缀，加上.txt）
    log_filename = task_filename.replace('.xlsx', '.txt')
    logger_name = f'task_{log_filename}'
    
    # 创建日志记录器
    logger = logging.getLogger(logger_name)
    logger.setLevel(logging.DEBUG)
    logger.handlers.clear()  # 清除现有处理器
    
    # 创建文件处理器
    file_handler = logging.FileHandler(f'log/{log_filename}', encoding='utf-8')
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)
    
    # 添加处理器
    logger.addHandler(file_handler)
    
    return logger

def read_df(index, row, logger_thread, sheet_name):
    logger_thread.info(f'第 {index} 行数据：')
    logger_thread.info(row)
    # logger_thread.info(f'第 {index} 行数据：', row)
    country_chinese = row.iloc[1]
    date_range = row.iloc[4].replace("至今", "2025")
    # date_range = '2024-2024'
    climate = row.iloc[5]
    policy = row.iloc[6]
    uncertainty = row.iloc[7]
    magazine_col = row.iloc[3]
    if pd.isna(country_chinese) or country_chinese.strip() == '中国':
        return
    if pd.notna(magazine_col) and magazine_col.strip() != '':
        magazines = re.split(r'[;；]', magazine_col)
        for magazine in magazines:
        # if magazines:
        #     magazine = magazines[0].strip()  # 去除首尾空格
            if magazine:  # 确保杂志名称不为空
                pubname = journal_name_to_code.get(magazine)
                if pubname:  # 确保能找到对应的代码
                    # 检查是否已有结果文件
                    filename = sheet_name + country_chinese + magazine + '.xlsx'
                    if not check_file_existence_with_pathlib(filename):
                        # 为这个任务创建专用的日志记录器
                        task_logger = create_task_logger(filename)
                        task_logger.info(f'🚀 开始处理任务: {filename}')
                        
                        # print(f'sheet:{sheet_name}, country_chinese:{country_chinese}, magazine:{magazine}, filename:{filename}')
                        crawl_data(filename, pubname, date_range, climate, policy, uncertainty, task_logger)
                        
                        task_logger.info(f'✅ 任务完成: {filename}')
                        # 清理日志处理器
                        for handler in task_logger.handlers[:]:
                            handler.close()
                            task_logger.removeHandler(handler)
                else:
                    # print(f'未找到期刊 {country_chinese},{magazine} 的代码映射')
                    logger_thread.warning(f"⚠️ 未找到期刊 '{country_chinese}','{magazine}' 的代码映射")
    else:
        logger_thread.warning(f"⚠️ 第 {index} 行的期刊列（第三列）为空")


def thread_function(index, row, sheet_name):
    # 创建一个以线程ID命名的日志记录器
    logger_thread = logging.getLogger(f'thread_{threading.current_thread().ident}')
    logger_thread.setLevel(logging.DEBUG)
    logger_thread.handlers.clear()

    # 创建一个文件处理器，将日志写入以线程ID命名的txt文件
    file_handler = logging.FileHandler(f'log/thread_{threading.current_thread().ident}.txt', encoding='utf-8')
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)

    # 将文件处理器添加到日志记录器
    logger_thread.addHandler(file_handler)

    try:
        logger_thread.info(f'线程 {threading.current_thread().ident} 开始执行')
        # 这里可以添加线程的具体逻辑
        read_df(index, row, logger_thread, sheet_name)
        logger_thread.info(f'线程 {threading.current_thread().ident} 执行结束')
        return 1
    finally:
        # 移除处理器，避免logger被缓存后重复使用
        logger_thread.removeHandler(file_handler)
        # 关闭文件处理器（释放文件句柄）
        file_handler.close()



def validate_magazine_group(sheet_names):
    country_content = read_file(country_list_file)
    # 遍历所有 sheet
    for sheet_name in sheet_names:
        # 获取当前 sheet 的数据
        df = excel_file.parse(sheet_name)
        # 查看数据的基本信息
        logger_global.info(f'sheet表名为{sheet_name}的基本信息：')
        df.info()
        # 查看数据集行数和列数
        rows, columns = df.shape
        if rows == 0:
            logger_global.info(f'sheet表名为{sheet_name}的没有数据')
            continue
        for index, row in df.iterrows():
            country_chinese = row.iloc[1]
            if pd.isna(country_chinese) or country_chinese.strip() == '中国':
                continue
            magazine_col = row.iloc[3]
            if pd.notna(magazine_col) and magazine_col.strip() != '':
                magazines = re.split(r'[;；]', magazine_col)
                for magazine in magazines:
                    magazine = magazine.strip()  # 去除首尾空格
                    if magazine:  # 确保杂志名称不为空
                        pubname = find_pattern_country(magazine, country_content)
                        if not pubname:
                            logger_global.warning(f"⚠️ 未找到期刊 '{country_chinese}','{magazine}' 的代码映射")
                            raise SystemExit("存在无法识别的期刊，程序退出")
                        journal_name_to_code[magazine] = pubname
            else:
                logger_global.warning(f"⚠️ 第 {index} 行的期刊列（第三列）为空")
    logger_global.info(journal_name_to_code)


if __name__ == '__main__':
    # time.sleep(8 * 60 * 60)
    # 创建一个以线程ID命名的日志记录器
    logger_global = logging.getLogger(f'thread_{threading.current_thread().ident}')
    logger_global.setLevel(logging.DEBUG)
    # 创建一个文件处理器，将日志写入以线程ID命名的txt文件
    file_handler_global = logging.FileHandler(f'log/global.txt', encoding='utf-8')
    formatter_global = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler_global.setFormatter(formatter_global)
    # 将文件处理器添加到日志记录器
    logger_global.addHandler(file_handler_global)
    logger_global.info(f'线程 {threading.current_thread().ident} 开始执行')
    
    # 第一阶段：使用线程池并行创建所有浏览器
    logger_global.info(f"开始创建{len(account_configs)}个浏览器实例...")
    
    def create_browser(config):
        """创建单个浏览器实例"""
        try:
            # 使用本地 ChromeDriver

            service = Service(driver_path)
            options = webdriver.ChromeOptions()
            options.add_argument("--disable-blink-features=AutomationControlled")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--disable-gpu")
            options.add_argument("--disable-extensions")
            options.add_argument("--disable-plugins")
            options.add_argument("--disable-images")
            options.add_argument("--incognito")
            # 禁用 Selenium 自动化标识（关键反爬处理）
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_experimental_option("useAutomationExtension", False)
            # 动态生成User-Agent
            def generate_user_agent(thread_id):
                """根据线程ID动态生成User-Agent"""
                
                # Chrome版本号范围 (最新的稳定版本)
                chrome_versions = [
                    "120.0.0.0", "121.0.0.0", "122.0.0.0", "123.0.0.0", "124.0.0.0"
                ]
                
                # Firefox版本号范围
                firefox_versions = [
                    "120.0", "121.0", "122.0", "123.0", "124.0"
                ]
                
                # Safari版本号范围
                safari_versions = [
                    "17.1", "17.2", "17.3", "17.4"
                ]
                
                # 操作系统版本
                windows_versions = [
                    "Windows NT 10.0; Win64; x64",
                    "Windows NT 11.0; Win64; x64"
                ]
                
                macos_versions = [
                    "Macintosh; Intel Mac OS X 10_15_7",
                    "Macintosh; Intel Mac OS X 13_6_0", 
                    "Macintosh; Intel Mac OS X 14_1_0",
                    "Macintosh; Intel Mac OS X 14_2_0"
                ]
                
                linux_versions = [
                    "X11; Linux x86_64",
                    "X11; Ubuntu; Linux x86_64"
                ]
                
                # 根据线程ID确定浏览器类型和操作系统，确保每个线程有唯一组合
                browser_type = thread_id % 4  # 0=Chrome Windows, 1=Chrome macOS, 2=Firefox, 3=Safari
                
                if browser_type == 0:  # Chrome Windows
                    os_info = windows_versions[thread_id % len(windows_versions)]
                    chrome_ver = chrome_versions[thread_id % len(chrome_versions)]
                    webkit_ver = f"537.{36 + (thread_id % 10)}"
                    return f"Mozilla/5.0 ({os_info}) AppleWebKit/{webkit_ver} (KHTML, like Gecko) Chrome/{chrome_ver} Safari/{webkit_ver}"
                
                elif browser_type == 1:  # Chrome macOS
                    os_info = macos_versions[thread_id % len(macos_versions)]
                    chrome_ver = chrome_versions[thread_id % len(chrome_versions)]
                    webkit_ver = f"537.{36 + (thread_id % 10)}"
                    return f"Mozilla/5.0 ({os_info}) AppleWebKit/{webkit_ver} (KHTML, like Gecko) Chrome/{chrome_ver} Safari/{webkit_ver}"
                
                elif browser_type == 2:  # Firefox
                    if thread_id % 2 == 0:
                        os_info = windows_versions[thread_id % len(windows_versions)]
                        firefox_ver = firefox_versions[thread_id % len(firefox_versions)]
                        rv_ver = f"109.{thread_id % 10}"
                        return f"Mozilla/5.0 ({os_info}; rv:{rv_ver}) Gecko/20100101 Firefox/{firefox_ver}"
                    else:
                        os_info = macos_versions[thread_id % len(macos_versions)]
                        firefox_ver = firefox_versions[thread_id % len(firefox_versions)]
                        rv_ver = f"109.{thread_id % 10}"
                        return f"Mozilla/5.0 ({os_info}; rv:{rv_ver}) Gecko/20100101 Firefox/{firefox_ver}"
                
                else:  # Safari macOS
                    os_info = macos_versions[thread_id % len(macos_versions)]
                    safari_ver = safari_versions[thread_id % len(safari_versions)]
                    webkit_ver = f"605.1.{15 + (thread_id % 10)}"
                    return f"Mozilla/5.0 ({os_info}) AppleWebKit/{webkit_ver} (KHTML, like Gecko) Version/{safari_ver} Safari/{webkit_ver}"
            
            # 为每个线程ID生成唯一的User-Agent
            selected_ua = generate_user_agent(config['thread_id'])
            options.add_argument(f"user-agent={selected_ua}")
            
            # 为每个线程ID分配不同的视口大小
            window_sizes = [
                (1366, 768),   # 最常见的桌面分辨率
                (1920, 1080),  # Full HD
                (1440, 900),   # MacBook Pro 13"
                (1536, 864),   # Surface Pro
                (1600, 900),   # 16:9 ratio
                (1280, 800),   # MacBook Air
                (1680, 1050),  # 16:10 ratio
                (2560, 1440),  # 2K
                (1280, 1024),  # 5:4 ratio
                (1024, 768),   # 4:3 ratio
                (1152, 864),   # 4:3 ratio
                (1280, 720),   # HD
                (1360, 768),   # Common laptop
                (1600, 1200),  # 4:3 ratio
                (1920, 1200),  # 16:10 ratio
                (2048, 1152),  # 16:9 ratio
                (2560, 1600),  # 16:10 ratio
                (3840, 2160),  # 4K
            ]
            size_index = (config['thread_id'] - 1) % len(window_sizes)
            width, height = window_sizes[size_index]
            options.add_argument(f"--window-size={width},{height}")
            
            logger_global.info(f"[浏览器{config['thread_id']}] 动态生成UA: {selected_ua[:60]}... 视口: {width}x{height}")
            
            config['driver'] = webdriver.Chrome(service=service, options=options)
            logger_global.info(f"[浏览器{config['thread_id']}|账号{config['username']}] ✅ 浏览器启动成功！")
            return True
        except Exception as e:
            logger_global.error(f"❌  [浏览器{config['thread_id']}|账号{config['username']}]  浏览器启动失败: {e}")
            return False
    
    with concurrent.futures.ThreadPoolExecutor(max_workers=len(account_configs)) as executor:
        # 提交所有浏览器创建任务
        future_to_config = {executor.submit(create_browser, config): config for config in account_configs}
        
        # 等待所有任务完成并收集结果
        for future in concurrent.futures.as_completed(future_to_config):
            config = future_to_config[future]
            try:
                result = future.result()
                if result:
                    logger_global.info(f"✅ 账号 {config['username']} 浏览器创建成功")
                else:
                    logger_global.error(f"❌ 账号 {config['username']} 浏览器创建失败")
            except Exception as exc:
                logger_global.error(f"❌ 账号 {config['username']} 浏览器创建异常: {exc}")
    
    # 第二阶段：为每个浏览器实例分配账号并并行登录
    logger_global.info(f"开始为{len(account_configs)}个浏览器并行登录...")
    
    def login_account(config):
        """单个账号登录函数"""
        if config['driver']:
            return get_cookie_for_thread(config['thread_id'], logger_global, config['username'], config['password'])
        else:
            logger_global.error(f"❌ [浏览器{config['thread_id']}|账号{config['username']}]  无可用浏览器实例")
            return False
    
    # 使用线程池并行登录所有账号
    with concurrent.futures.ThreadPoolExecutor(max_workers=len(account_configs)) as executor:
        # 提交所有登录任务
        future_to_config = {executor.submit(login_account, config): config for config in account_configs}
        
        # 等待所有任务完成并收集结果
        for future in concurrent.futures.as_completed(future_to_config):
            config = future_to_config[future]
            try:
                result = future.result()
                if result:
                    logger_global.info(f"✅ 账号 {config['username']} 登录成功")
                else:
                    logger_global.error(f"❌ 账号 {config['username']} 登录失败")
            except Exception as exc:
                logger_global.error(f"❌ 账号 {config['username']} 登录异常: {exc}")
    
    # 保持向后兼容
    driver = account_configs[0]['driver'] if account_configs[0]['driver'] else None
    
    logger_global.info(f'线程 {threading.current_thread().ident} 执行结束')
    # 读取 Excel 文件
    excel_file = pd.ExcelFile(input_xlsx)
    sheet_names = excel_file.sheet_names
    validate_magazine_group(sheet_names)

    # get_cookie(logger_global)
    # 等五分钟，输入账号密码
    # driver.get(login_page)
    # time.sleep(5 * 60)
    # 遍历所有 sheet
    for sheet_name in sheet_names:
        # 获取当前 sheet 的数据
        df = excel_file.parse(sheet_name)
        # 查看数据的基本信息
        logger_global.info(f'sheet表名为{sheet_name}的基本信息：{df.info()}')
        # 查看数据集行数和列数
        rows, columns = df.shape
        if rows == 0:
            logger_global.info(f'sheet表名为{sheet_name}的没有数据')
            continue
        for index, row in df.iterrows():
            # 逐个提交任务，并将Future对象添加到列表
            read_df(index, row, logger_global, sheet_name)
        #     future = executor.submit(thread_function, index, row, sheet_name)
        #     futures.append(future)
