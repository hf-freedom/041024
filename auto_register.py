import random
import string
import os
import time
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from playwright.sync_api import sync_playwright
import openpyxl
from openpyxl import Workbook

SCREENSHOT_DIR = r"C:\Users\12824\Desktop\dataLabel\0319\p13\login_picture"
EXCEL_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "register_data.xlsx")

FIRST_NAMES = ["张", "王", "李", "赵", "刘", "陈", "杨", "黄", "周", "吴", "徐", "孙", "马", "朱", "胡", "郭", "何", "高", "林", "罗"]
LAST_NAMES = ["伟", "芳", "娜", "秀英", "敏", "静", "丽", "强", "磊", "军", "洋", "勇", "艳", "杰", "娟", "涛", "明", "超", "秀兰", "霞"]

excel_lock = threading.Lock()

def generate_random_string(length=8):
    return ''.join(random.choices(string.ascii_lowercase + string.digits, k=length))

def generate_random_email():
    username = generate_random_string(10)
    return f"{username}@163.com"

def generate_random_password():
    return generate_random_string(12) + random.choice(string.ascii_uppercase) + random.choice(string.digits)

def generate_random_name():
    first_name = random.choice(FIRST_NAMES)
    last_name = random.choice(LAST_NAMES)
    return first_name + last_name

def generate_random_age():
    return str(random.randint(18, 60))

def generate_random_phone():
    prefixes = ["130", "131", "132", "133", "134", "135", "136", "137", "138", "139",
                "150", "151", "152", "153", "155", "156", "157", "158", "159",
                "170", "176", "177", "178",
                "180", "181", "182", "183", "184", "185", "186", "187", "188", "189"]
    prefix = random.choice(prefixes)
    suffix = ''.join(random.choices(string.digits, k=8))
    return prefix + suffix

def ensure_screenshot_dir():
    if not os.path.exists(SCREENSHOT_DIR):
        os.makedirs(SCREENSHOT_DIR)
        print(f"创建截图目录: {SCREENSHOT_DIR}")

def init_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = "注册数据"
        headers = ["序号", "用户名", "密码", "邮箱", "姓名", "年龄", "手机号", "注册开始时间", "注册结束时间", "注册耗时(秒)", "登录开始时间", "登录结束时间", "登录耗时(秒)", "总耗时(秒)", "注册状态", "登录状态", "验证状态"]
        ws.append(headers)
        wb.save(EXCEL_FILE)
        print(f"创建Excel文件: {EXCEL_FILE}")
    return EXCEL_FILE

def save_to_excel(data):
    with excel_lock:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb.active
        ws.append(data)
        wb.save(EXCEL_FILE)

def find_input(page, selectors, field_name):
    for selector in selectors:
        try:
            element = page.query_selector(selector)
            if element:
                return element, selector
        except:
            continue
    return None, None

def perform_register(page, user_data, task_id):
    username = user_data["username"]
    password = user_data["password"]
    email = user_data["email"]
    name = user_data["name"]
    age = user_data["age"]
    phone = user_data["phone"]
    
    register_start_time = datetime.now()
    
    username_selectors = [
        'input[name="username"]',
        'input[name="user"]',
        'input[name="userName"]',
        'input[placeholder*="用户名"]',
        'input[placeholder*="账号"]',
        '#username',
        '#userName',
        'input[type="text"]:first-of-type'
    ]
    
    password_selectors = [
        'input[name="password"]',
        'input[name="pwd"]',
        'input[name="userPassword"]',
        'input[placeholder*="密码"]',
        '#password',
        '#pwd',
        'input[type="password"]'
    ]
    
    email_selectors = [
        'input[name="email"]',
        'input[name="mail"]',
        'input[placeholder*="邮箱"]',
        'input[placeholder*="Email"]',
        '#email',
        'input[type="email"]'
    ]
    
    name_selectors = [
        'input[name="name"]',
        'input[name="realName"]',
        'input[name="realname"]',
        'input[name="userName"]',
        'input[placeholder*="姓名"]',
        'input[placeholder*="真实姓名"]',
        '#name',
        '#realName',
        '#userName'
    ]
    
    age_selectors = [
        'input[name="age"]',
        'input[placeholder*="年龄"]',
        '#age',
        'input[type="number"]'
    ]
    
    phone_selectors = [
        'input[name="phone"]',
        'input[name="mobile"]',
        'input[name="tel"]',
        'input[name="phoneNumber"]',
        'input[placeholder*="手机"]',
        'input[placeholder*="电话"]',
        'input[placeholder*="手机号"]',
        '#phone',
        '#mobile',
        '#tel'
    ]
    
    print(f"[任务{task_id}] 查找注册表单字段...")
    
    username_input, _ = find_input(page, username_selectors, "用户名")
    password_input, _ = find_input(page, password_selectors, "密码")
    email_input, _ = find_input(page, email_selectors, "邮箱")
    name_input, _ = find_input(page, name_selectors, "姓名")
    age_input, _ = find_input(page, age_selectors, "年龄")
    phone_input, _ = find_input(page, phone_selectors, "手机号")
    
    print(f"[任务{task_id}] 填写注册信息...")
    
    if username_input:
        username_input.fill(username)
    if password_input:
        password_input.fill(password)
    if email_input:
        email_input.fill(email)
    if name_input:
        name_input.fill(name)
    if age_input:
        age_input.fill(age)
    if phone_input:
        phone_input.fill(phone)
    
    page.wait_for_timeout(500)
    
    submit_selectors = [
        'button:has-text("注册")',
        'button:has-text("提交")',
        'button:has-text("确定")',
        'input[type="submit"]',
        'input[value="注册"]',
        'input[value="提交"]',
        '.register-btn',
        '#register-btn',
        'button[type="submit"]'
    ]
    
    submit_btn = None
    for selector in submit_selectors:
        try:
            submit_btn = page.query_selector(selector)
            if submit_btn:
                break
        except:
            continue
    
    register_status = "失败"
    if submit_btn:
        print(f"[任务{task_id}] 点击注册按钮...")
        submit_btn.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(2000)
        register_status = "成功"
    
    register_end_time = datetime.now()
    register_duration = (register_end_time - register_start_time).total_seconds()
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    screenshot_path = os.path.join(SCREENSHOT_DIR, f"task{task_id}_register_{timestamp}.png")
    page.screenshot(path=screenshot_path)
    
    return {
        "register_start_time": register_start_time,
        "register_end_time": register_end_time,
        "register_duration": register_duration,
        "register_status": register_status
    }

def perform_login(page, user_data, task_id):
    username = user_data["username"]
    password = user_data["password"]
    
    login_start_time = datetime.now()
    
    print(f"[任务{task_id}] 跳转到登录页面...")
    
    login_link_selectors = [
        'a:has-text("登录")',
        'text=登录',
        'a[href*="login"]',
        '.login-link',
        '#login-link'
    ]
    
    login_link = None
    for selector in login_link_selectors:
        try:
            login_link = page.query_selector(selector)
            if login_link:
                break
        except:
            continue
    
    if login_link:
        login_link.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1000)
    else:
        page.goto("http://39.107.109.8:8082/", timeout=30000)
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1000)
    
    username_selectors = [
        'input[name="username"]',
        'input[name="user"]',
        'input[name="userName"]',
        'input[placeholder*="用户名"]',
        'input[placeholder*="账号"]',
        '#username',
        '#userName',
        'input[type="text"]:first-of-type'
    ]
    
    password_selectors = [
        'input[name="password"]',
        'input[name="pwd"]',
        'input[name="userPassword"]',
        'input[placeholder*="密码"]',
        '#password',
        '#pwd',
        'input[type="password"]'
    ]
    
    username_input, _ = find_input(page, username_selectors, "用户名")
    password_input, _ = find_input(page, password_selectors, "密码")
    
    print(f"[任务{task_id}] 填写登录信息...")
    
    if username_input:
        username_input.fill(username)
    if password_input:
        password_input.fill(password)
    
    page.wait_for_timeout(500)
    
    login_btn_selectors = [
        'button:has-text("登录")',
        'button:has-text("Login")',
        'input[type="submit"]',
        'input[value="登录"]',
        '.login-btn',
        '#login-btn',
        'button[type="submit"]'
    ]
    
    login_btn = None
    for selector in login_btn_selectors:
        try:
            login_btn = page.query_selector(selector)
            if login_btn:
                break
        except:
            continue
    
    login_status = "失败"
    verify_status = "未验证"
    
    if login_btn:
        print(f"[任务{task_id}] 点击登录按钮...")
        login_btn.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(2000)
        login_status = "成功"
        
        print(f"[任务{task_id}] 验证用户信息...")
        
        page_content = page.content()
        verify_keywords = [user_data["username"], user_data["name"], "个人信息", "欢迎", "我的", "用户中心"]
        
        for keyword in verify_keywords:
            if keyword in page_content:
                verify_status = "验证成功"
                break
        
        if verify_status != "验证成功":
            verify_status = "验证失败"
    
    login_end_time = datetime.now()
    login_duration = (login_end_time - login_start_time).total_seconds()
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    screenshot_path = os.path.join(SCREENSHOT_DIR, f"task{task_id}_login_{timestamp}.png")
    page.screenshot(path=screenshot_path)
    
    return {
        "login_start_time": login_start_time,
        "login_end_time": login_end_time,
        "login_duration": login_duration,
        "login_status": login_status,
        "verify_status": verify_status
    }

def single_task(task_id, user_data):
    print(f"\n[任务{task_id}] 开始执行...")
    print(f"[任务{task_id}] 用户名: {user_data['username']}")
    
    task_start_time = datetime.now()
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()
        
        try:
            print(f"[任务{task_id}] 正在打开网站...")
            page.goto("http://39.107.109.8:8082/", timeout=30000)
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(2000)
            
            print(f"[任务{task_id}] 查找注册链接...")
            register_link = page.query_selector('text=注册') or page.query_selector('text=立即注册') or page.query_selector('a:has-text("注册")')
            
            if register_link:
                print(f"[任务{task_id}] 点击注册链接...")
                register_link.click()
                page.wait_for_load_state("networkidle")
                page.wait_for_timeout(1000)
            
            register_result = perform_register(page, user_data, task_id)
            
            login_result = perform_login(page, user_data, task_id)
            
            task_end_time = datetime.now()
            total_duration = (task_end_time - task_start_time).total_seconds()
            
            excel_data = [
                task_id,
                user_data["username"],
                user_data["password"],
                user_data["email"],
                user_data["name"],
                user_data["age"],
                user_data["phone"],
                register_result["register_start_time"].strftime("%Y-%m-%d %H:%M:%S"),
                register_result["register_end_time"].strftime("%Y-%m-%d %H:%M:%S"),
                round(register_result["register_duration"], 2),
                login_result["login_start_time"].strftime("%Y-%m-%d %H:%M:%S"),
                login_result["login_end_time"].strftime("%Y-%m-%d %H:%M:%S"),
                round(login_result["login_duration"], 2),
                round(total_duration, 2),
                register_result["register_status"],
                login_result["login_status"],
                login_result["verify_status"]
            ]
            
            save_to_excel(excel_data)
            
            print(f"\n[任务{task_id}] ========== 执行完成 ==========")
            print(f"[任务{task_id}] 注册状态: {register_result['register_status']}")
            print(f"[任务{task_id}] 注册耗时: {register_result['register_duration']:.2f}秒")
            print(f"[任务{task_id}] 登录状态: {login_result['login_status']}")
            print(f"[任务{task_id}] 登录耗时: {login_result['login_duration']:.2f}秒")
            print(f"[任务{task_id}] 验证状态: {login_result['verify_status']}")
            print(f"[任务{task_id}] 总耗时: {total_duration:.2f}秒")
            print(f"[任务{task_id}] ==============================\n")
            
            return {
                "task_id": task_id,
                "status": "成功",
                "total_duration": total_duration,
                "register_status": register_result["register_status"],
                "login_status": login_result["login_status"],
                "verify_status": login_result["verify_status"]
            }
            
        except Exception as e:
            print(f"[任务{task_id}] 发生错误: {e}")
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            error_screenshot = os.path.join(SCREENSHOT_DIR, f"task{task_id}_error_{timestamp}.png")
            try:
                page.screenshot(path=error_screenshot)
            except:
                pass
            
            return {
                "task_id": task_id,
                "status": "失败",
                "error": str(e)
            }
        finally:
            browser.close()

def generate_user_data():
    return {
        "username": "user_" + generate_random_string(6),
        "password": generate_random_password(),
        "email": generate_random_email(),
        "name": generate_random_name(),
        "age": generate_random_age(),
        "phone": generate_random_phone()
    }

def run_parallel_register(num_tasks=5):
    ensure_screenshot_dir()
    init_excel()
    
    print("=" * 60)
    print(f"开始并行执行 {num_tasks} 个注册任务")
    print("=" * 60)
    
    overall_start_time = datetime.now()
    
    users_data = [generate_user_data() for _ in range(num_tasks)]
    
    print("\n生成的用户信息:")
    for i, user in enumerate(users_data, 1):
        print(f"  任务{i}: {user['username']}")
    
    results = []
    
    with ThreadPoolExecutor(max_workers=num_tasks) as executor:
        futures = {executor.submit(single_task, i+1, user): i+1 for i, user in enumerate(users_data)}
        
        for future in as_completed(futures):
            task_id = futures[future]
            try:
                result = future.result()
                results.append(result)
            except Exception as e:
                print(f"任务{task_id}执行异常: {e}")
                results.append({"task_id": task_id, "status": "异常", "error": str(e)})
    
    overall_end_time = datetime.now()
    overall_duration = (overall_end_time - overall_start_time).total_seconds()
    
    print("\n" + "=" * 60)
    print("所有任务执行完成!")
    print("=" * 60)
    
    success_count = sum(1 for r in results if r.get("status") == "成功")
    fail_count = num_tasks - success_count
    
    print(f"\n执行统计:")
    print(f"  总任务数: {num_tasks}")
    print(f"  成功: {success_count}")
    print(f"  失败: {fail_count}")
    print(f"  总耗时: {overall_duration:.2f}秒")
    print(f"  平均耗时: {overall_duration/num_tasks:.2f}秒/任务")
    
    print(f"\n数据已保存到: {EXCEL_FILE}")
    print(f"截图已保存到: {SCREENSHOT_DIR}")
    
    return results

if __name__ == "__main__":
    results = run_parallel_register(5)
