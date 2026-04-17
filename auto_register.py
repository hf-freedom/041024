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

def init_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = "字段校验结果"
        headers = ["序号", "账号", "密码", "测试字段", "原字段值", "目标修改值", "接口返回结果", "测试时间", "注册状态", "登录状态"]
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

def generate_test_values():
    special_chars = "!@#$%^&*()_+-=[]{}|;':\",./<>?`~"
    return {
        "name": [
            ("空字段", ""),
            ("100字符长字段", "张" * 100),
            ("数字", "12345"),
            ("特殊符号", "!@#$%^&*()"),
            ("混合字符", "张三123!@#")
        ],
        "age": [
            ("空字段", ""),
            ("100字符长字段", "1" * 100),
            ("非数字", "abcde"),
            ("特殊符号", "!@#$%"),
            ("负数", "-25"),
            ("超大数字", "999999"),
            ("小数", "25.5")
        ],
        "phone": [
            ("空字段", ""),
            ("100字符长字段", "1" * 100),
            ("非数字", "abcdefghijk"),
            ("特殊符号", "!@#$%^&*()"),
            ("短号码", "123"),
            ("含字母", "1381234567a")
        ],
        "email": [
            ("空字段", ""),
            ("100字符长字段", "a" * 100 + "@test.com"),
            ("无@符号", "testtest.com"),
            ("特殊符号", "!@#$%@test.com"),
            ("无域名", "test@"),
            ("多个@", "test@test@test.com")
        ]
    }

def perform_register(page, user_data, task_id):
    username = user_data["username"]
    password = user_data["password"]
    email = user_data["email"]
    name = user_data["name"]
    age = user_data["age"]
    phone = user_data["phone"]
    
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
    
    return register_status

def perform_login(page, user_data, task_id):
    username = user_data["username"]
    password = user_data["password"]
    
    print(f"[任务{task_id}] 跳转到登录页面...")
    
    page.goto("http://39.107.109.8:8082/login", timeout=30000)
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)
    
    username_selectors = [
        'input[placeholder*="账号"]',
        'input[placeholder*="用户名"]',
        'input[name="username"]',
        'input[name="user"]',
        '#username',
        'input[type="text"]'
    ]
    
    password_selectors = [
        'input[placeholder*="密码"]',
        'input[name="password"]',
        '#password',
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
    
    if login_btn:
        print(f"[任务{task_id}] 点击登录按钮...")
        login_btn.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(3000)
        
        print(f"[任务{task_id}] 尝试访问profile页面验证登录状态...")
        page.goto("http://39.107.109.8:8082/profile", timeout=30000)
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(2000)
        
        if "/profile" in page.url:
            edit_btn = page.query_selector('button:has-text("编辑")')
            if edit_btn:
                login_status = "成功"
                print(f"[任务{task_id}] 登录验证成功")
    
    return login_status

def click_edit_button(page, task_id):
    print(f"[任务{task_id}] 查找并点击编辑按钮...")
    
    edit_selectors = [
        'button:has-text("编辑")',
        'button.btn-edit',
        '.btn-edit',
        'button:has-text("修改")',
        'button:has-text("编辑资料")'
    ]
    
    for attempt in range(3):
        for selector in edit_selectors:
            try:
                edit_btn = page.query_selector(selector)
                if edit_btn and edit_btn.is_visible():
                    print(f"[任务{task_id}] 点击编辑按钮: {selector}")
                    edit_btn.click()
                    page.wait_for_timeout(1500)
                    return True
            except:
                continue
        
        if attempt < 2:
            print(f"[任务{task_id}] 第{attempt+1}次尝试未找到编辑按钮，等待后重试...")
            page.wait_for_timeout(1000)
    
    print(f"[任务{task_id}] 未找到编辑按钮")
    return False

def get_field_value(page, field_selectors):
    for selector in field_selectors:
        try:
            element = page.query_selector(selector)
            if element:
                value = element.input_value()
                return value
        except:
            continue
    return ""

def test_field_validation(page, field_name, test_value, original_value, user_data, task_id, test_index):
    field_selectors_map = {
        "name": [
            'input[placeholder*="请输入姓名"]',
            'input.edit-input[type="text"]',
            'input[name="name"]',
            'input[placeholder*="姓名"]'
        ],
        "age": [
            'input[placeholder*="请输入年龄"]',
            'input.edit-input[type="number"]',
            'input[name="age"]',
            'input[placeholder*="年龄"]'
        ],
        "phone": [
            'input[placeholder*="请输入手机号"]',
            'input.edit-input[type="tel"]',
            'input[name="phone"]',
            'input[name="mobile"]',
            'input[placeholder*="手机"]'
        ],
        "email": [
            'input[placeholder*="请输入邮箱"]',
            'input.edit-input[type="email"]',
            'input[name="email"]',
            'input[placeholder*="邮箱"]'
        ]
    }
    
    selectors = field_selectors_map.get(field_name, [])
    field_input, _ = find_input(page, selectors, field_name)
    
    if not field_input:
        return None, "未找到字段输入框"
    
    try:
        if field_name == "age":
            field_input.evaluate('el => el.value = "' + str(test_value).replace('"', '\\"') + '"')
            field_input.dispatch_event('input')
            field_input.dispatch_event('change')
        else:
            field_input.fill("")
            page.wait_for_timeout(100)
            field_input.fill(test_value)
        page.wait_for_timeout(300)
    except Exception as e:
        return test_value, f"填写字段失败: {str(e)[:50]}"
    
    save_selectors = [
        'button.btn-save',
        'button:has-text("保存")',
        '.btn-save',
        'button:has-text("提交")',
        'button[type="submit"]'
    ]
    
    save_btn = None
    for selector in save_selectors:
        try:
            save_btn = page.query_selector(selector)
            if save_btn:
                break
        except:
            continue
    
    if save_btn:
        save_btn.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1500)
        
        error_selectors = [
            '.error',
            '.error-message',
            '.alert-error',
            '.alert-danger',
            '.message-error',
            '[class*="error"]',
            '.toast-error',
            '.notification-error'
        ]
        
        error_text = ""
        for selector in error_selectors:
            try:
                error_element = page.query_selector(selector)
                if error_element:
                    error_text = error_element.inner_text()
                    if error_text:
                        break
            except:
                continue
        
        if not error_text:
            try:
                page_content = page.content()
                error_keywords = ["错误", "失败", "无效", "不正确", "格式", "必填", "不能为空", "error", "invalid", "required", "failed"]
                for keyword in error_keywords:
                    if keyword.lower() in page_content.lower():
                        error_text = f"检测到错误提示(含'{keyword}')"
                        break
            except:
                pass
        
        if error_text:
            result = f"校验失败: {error_text[:100]}"
        else:
            current_value = get_field_value(page, selectors)
            if current_value == test_value:
                result = "修改成功(无校验)"
            else:
                result = "修改失败(值未变更)"
        
        return test_value, result
    else:
        return test_value, "未找到保存按钮"

def perform_field_validation_tests(page, user_data, task_id):
    print(f"[任务{task_id}] 开始字段校验测试...")
    
    test_values = generate_test_values()
    test_results = []
    test_index = 0
    
    for field_name, test_cases in test_values.items():
        print(f"[任务{task_id}] 测试字段: {field_name}")
        
        field_selectors_map = {
            "name": [
                'input[placeholder*="请输入姓名"]',
                'input.edit-input[type="text"]',
                'input[name="name"]',
                'input[placeholder*="姓名"]'
            ],
            "age": [
                'input[placeholder*="请输入年龄"]',
                'input.edit-input[type="number"]',
                'input[name="age"]',
                'input[placeholder*="年龄"]'
            ],
            "phone": [
                'input[placeholder*="请输入手机号"]',
                'input.edit-input[type="tel"]',
                'input[name="phone"]',
                'input[name="mobile"]',
                'input[placeholder*="手机"]'
            ],
            "email": [
                'input[placeholder*="请输入邮箱"]',
                'input.edit-input[type="email"]',
                'input[name="email"]',
                'input[placeholder*="邮箱"]'
            ]
        }
        
        original_value = get_field_value(page, field_selectors_map.get(field_name, []))
        
        for test_name, test_value in test_cases:
            test_index += 1
            print(f"[任务{task_id}]   - 测试用例: {test_name}")
            
            page.goto("http://39.107.109.8:8082/profile", timeout=30000)
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(1000)
            
            click_edit_button(page, task_id)
            
            modified_value, result = test_field_validation(
                page, field_name, test_value, original_value, 
                user_data, task_id, test_index
            )
            
            if modified_value is not None:
                test_results.append({
                    "test_index": test_index,
                    "account": user_data["username"],
                    "password": user_data["password"],
                    "field_name": field_name,
                    "original_value": original_value,
                    "test_value": test_value,
                    "test_name": test_name,
                    "result": result,
                    "test_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                })
                
                excel_data = [
                    test_index,
                    user_data["username"],
                    user_data["password"],
                    f"{field_name}({test_name})",
                    original_value,
                    test_value,
                    result,
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "成功",
                    "成功"
                ]
                save_to_excel(excel_data)
            
            page.wait_for_timeout(500)
    
    return test_results

def single_task(task_id, user_data):
    print(f"\n[任务{task_id}] 开始执行...")
    print(f"[任务{task_id}] 用户名: {user_data['username']}")
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
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
            
            register_status = perform_register(page, user_data, task_id)
            print(f"[任务{task_id}] 注册状态: {register_status}")
            
            login_status = perform_login(page, user_data, task_id)
            print(f"[任务{task_id}] 登录状态: {login_status}")
            
            if login_status == "成功":
                click_edit_button(page, task_id)
                
                test_results = perform_field_validation_tests(page, user_data, task_id)
                
                print(f"\n[任务{task_id}] ========== 测试结果 ==========")
                for result in test_results:
                    print(f"[任务{task_id}] 字段: {result['field_name']}, 测试: {result['test_name']}, 结果: {result['result']}")
                print(f"[任务{task_id}] ==============================\n")
            
            return {
                "task_id": task_id,
                "status": "成功",
                "register_status": register_status,
                "login_status": login_status,
                "test_count": len(test_results) if 'test_results' in dir() else 0
            }
            
        except Exception as e:
            print(f"[任务{task_id}] 发生错误: {e}")
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

def run_parallel_register(num_tasks=1):
    init_excel()
    
    print("=" * 60)
    print(f"开始执行 {num_tasks} 个注册和字段校验任务")
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
    
    print(f"\n数据已保存到: {EXCEL_FILE}")
    
    return results

if __name__ == "__main__":
    results = run_parallel_register(1)
