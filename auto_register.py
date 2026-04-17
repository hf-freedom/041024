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

EXCEL_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "validation_results.xlsx")
BASE_URL = "http://39.107.109.8:8082"

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
        headers = ["序号", "测试时间", "账号", "要修改字段", "原字段值", "目标修改值", "测试类型", "接口返回结果", "校验是否通过"]
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


def find_input(page, selectors):
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
    
    username_input, _ = find_input(page, username_selectors)
    password_input, _ = find_input(page, password_selectors)
    email_input, _ = find_input(page, email_selectors)
    name_input, _ = find_input(page, name_selectors)
    age_input, _ = find_input(page, age_selectors)
    phone_input, _ = find_input(page, phone_selectors)
    
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
    
    return register_status == "成功"


def perform_login(page, user_data, task_id):
    username = user_data["username"]
    password = user_data["password"]
    
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
        page.goto(f"{BASE_URL}/", timeout=30000)
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
    
    username_input, _ = find_input(page, username_selectors)
    password_input, _ = find_input(page, password_selectors)
    
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
        page.wait_for_timeout(2000)
        login_status = "成功"
    
    return login_status == "成功"


def navigate_to_profile(page, task_id):
    print(f"[任务{task_id}] 查找个人信息页面入口...")
    
    # 首先检查当前页面是否已经是个人信息页面
    page_content = page.content()
    profile_keywords = ["个人信息", "个人中心", "编辑资料", "修改资料", "资料修改"]
    if any(keyword in page_content for keyword in profile_keywords):
        print(f"[任务{task_id}] 当前页面可能已是个人信息页面")
        # 检查是否有可编辑的字段
        edit_selectors = [
            'input[name="name"]', 'input[name="email"]', 'input[name="age"]', 'input[name="phone"]',
            '#name', '#email', '#age', '#phone',
            'input[placeholder*="姓名"]', 'input[placeholder*="邮箱"]', 'input[placeholder*="年龄"]', 'input[placeholder*="手机"]'
        ]
        for selector in edit_selectors:
            try:
                element = page.query_selector(selector)
                if element:
                    print(f"[任务{task_id}] 找到可编辑字段，确认是个人信息页面")
                    return True
            except:
                continue
    
    profile_selectors = [
        'a:has-text("个人信息")',
        'a:has-text("个人中心")',
        'a:has-text("我的")',
        'a:has-text("用户中心")',
        'a:has-text("编辑资料")',
        'a:has-text("修改资料")',
        'text=个人信息',
        'text=个人中心',
        'text=编辑资料',
        '.profile-link',
        '#profile-link',
        '.user-center',
        '#user-center',
        'a[href*="profile"]',
        'a[href*="user"]',
        'a[href*="edit"]',
        'a[href*="setting"]'
    ]
    
    profile_link = None
    for selector in profile_selectors:
        try:
            profile_link = page.query_selector(selector)
            if profile_link:
                print(f"[任务{task_id}] 找到个人信息入口: {selector}")
                break
        except:
            continue
    
    if profile_link:
        print(f"[任务{task_id}] 点击进入个人信息页面...")
        profile_link.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(2000)
        return True
    else:
        # 尝试多种可能的URL
        possible_urls = [
            f"{BASE_URL}/profile",
            f"{BASE_URL}/user/profile",
            f"{BASE_URL}/user/edit",
            f"{BASE_URL}/settings",
            f"{BASE_URL}/user/settings",
            f"{BASE_URL}/account",
            f"{BASE_URL}/member/profile",
            f"{BASE_URL}/home"
        ]
        
        for url in possible_urls:
            try:
                print(f"[任务{task_id}] 尝试访问: {url}")
                page.goto(url, timeout=15000)
                page.wait_for_load_state("networkidle")
                page.wait_for_timeout(1500)
                
                # 检查是否有可编辑字段
                edit_selectors = [
                    'input[name="name"]', 'input[name="email"]', 'input[name="age"]', 'input[name="phone"]',
                    '#name', '#email', '#age', '#phone'
                ]
                for selector in edit_selectors:
                    try:
                        element = page.query_selector(selector)
                        if element:
                            print(f"[任务{task_id}] 在 {url} 找到可编辑字段")
                            return True
                    except:
                        continue
            except Exception as e:
                continue
        
        print(f"[任务{task_id}] 未能找到个人信息编辑页面")
        return False


def get_current_field_values(page):
    """获取当前个人信息页面的字段值"""
    field_values = {}
    
    field_selectors = {
        "username": ['input[name="username"]', '#username', 'input[placeholder*="用户名"]'],
        "email": ['input[name="email"]', '#email', 'input[type="email"]', 'input[placeholder*="邮箱"]'],
        "name": ['input[name="name"]', '#name', 'input[placeholder*="姓名"]', 'input[placeholder*="真实姓名"]'],
        "age": ['input[name="age"]', '#age', 'input[placeholder*="年龄"]'],
        "phone": ['input[name="phone"]', '#phone', 'input[placeholder*="手机"]', 'input[placeholder*="电话"]']
    }
    
    for field_name, selectors in field_selectors.items():
        for selector in selectors:
            try:
                element = page.query_selector(selector)
                if element:
                    value = element.input_value()
                    if value:
                        field_values[field_name] = value
                        break
            except:
                continue
    
    return field_values


def generate_test_cases():
    """生成各种测试用例"""
    test_cases = []
    
    test_cases.append({"field": "name", "value": "", "type": "空字段"})
    test_cases.append({"field": "name", "value": "a" * 100, "type": "超长字段(100字符)"})
    test_cases.append({"field": "name", "value": "123456", "type": "纯数字"})
    test_cases.append({"field": "name", "value": "!@#$%^&*()", "type": "特殊符号"})
    test_cases.append({"field": "name", "value": "<script>alert('xss')</script>", "type": "XSS攻击"})
    test_cases.append({"field": "name", "value": "张三' OR '1'='1", "type": "SQL注入"})
    
    test_cases.append({"field": "email", "value": "", "type": "空字段"})
    test_cases.append({"field": "email", "value": "invalid_email", "type": "无效格式"})
    test_cases.append({"field": "email", "value": "test@", "type": "不完整邮箱"})
    test_cases.append({"field": "email", "value": "test@test.com'", "type": "SQL注入"})
    test_cases.append({"field": "email", "value": "a" * 100 + "@test.com", "type": "超长邮箱"})
    
    test_cases.append({"field": "age", "value": "", "type": "空字段"})
    test_cases.append({"field": "age", "value": "abc", "type": "非数字"})
    test_cases.append({"field": "age", "value": "-1", "type": "负数"})
    test_cases.append({"field": "age", "value": "999", "type": "超大数字"})
    test_cases.append({"field": "age", "value": "3.14", "type": "小数"})
    test_cases.append({"field": "age", "value": "25'", "type": "SQL注入"})
    
    test_cases.append({"field": "phone", "value": "", "type": "空字段"})
    test_cases.append({"field": "phone", "value": "abc123", "type": "非数字"})
    test_cases.append({"field": "phone", "value": "123", "type": "位数不足"})
    test_cases.append({"field": "phone", "value": "1380013800013800", "type": "超长号码"})
    test_cases.append({"field": "phone", "value": "13800138000'", "type": "SQL注入"})
    test_cases.append({"field": "phone", "value": "<script>", "type": "XSS攻击"})
    
    return test_cases


def perform_field_validation(page, user_data, task_id, seq_id):
    """执行字段校验测试"""
    test_cases = generate_test_cases()
    results = []
    
    field_selectors = {
        "name": ['input[name="name"]', '#name', 'input[placeholder*="姓名"]', 'input[placeholder*="真实姓名"]'],
        "email": ['input[name="email"]', '#email', 'input[type="email"]', 'input[placeholder*="邮箱"]'],
        "age": ['input[name="age"]', '#age', 'input[placeholder*="年龄"]'],
        "phone": ['input[name="phone"]', '#phone', 'input[placeholder*="手机"]', 'input[placeholder*="电话"]']
    }
    
    save_btn_selectors = [
        'button:has-text("保存")',
        'button:has-text("提交")',
        'button:has-text("修改")',
        'input[value="保存"]',
        'input[value="提交"]',
        '.save-btn',
        '#save-btn',
        'button[type="submit"]'
    ]
    
    current_values = get_current_field_values(page)
    
    for test_case in test_cases:
        field_name = test_case["field"]
        test_value = test_case["value"]
        test_type = test_case["type"]
        
        original_value = current_values.get(field_name, "")
        
        print(f"[任务{task_id}] 测试字段: {field_name}, 类型: {test_type}, 值: {test_value[:30] if len(test_value) > 30 else test_value}")
        
        selectors = field_selectors.get(field_name, [])
        field_input, _ = find_input(page, selectors)
        
        if not field_input:
            print(f"[任务{task_id}] 未找到字段 {field_name} 的输入框，跳过")
            continue
        
        try:
            field_input.fill("")
            page.wait_for_timeout(200)
            field_input.fill(test_value)
            page.wait_for_timeout(300)
        except Exception as fill_error:
            # 如果填充失败（如number类型不能填非数字），记录错误并继续
            print(f"[任务{task_id}] 字段 {field_name} 填充值 '{test_value[:20]}...' 失败: {fill_error}")
            api_response = f"前端阻止输入: {str(fill_error)}"
            validation_passed = "是"  # 前端阻止了非法输入，说明有校验
            
            test_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            excel_row = [
                seq_id,
                test_time,
                user_data["username"],
                field_name,
                original_value,
                test_value if len(test_value) <= 100 else test_value[:100] + "...",
                test_type,
                api_response[:500] if len(api_response) > 500 else api_response,
                validation_passed
            ]
            
            save_to_excel(excel_row)
            results.append({
                "field": field_name,
                "type": test_type,
                "value": test_value,
                "response": api_response,
                "validation": validation_passed
            })
            
            seq_id += 1
            continue
        
        save_btn = None
        for selector in save_btn_selectors:
            try:
                save_btn = page.query_selector(selector)
                if save_btn:
                    break
            except:
                continue
        
        api_response = "未获取到响应"
        validation_passed = "未知"
        
        if save_btn:
            # 记录点击前的URL
            before_url = page.url
            
            try:
                # 点击保存按钮
                save_btn.click()
                
                # 等待页面加载完成（传统表单提交会跳转页面）
                page.wait_for_load_state("networkidle")
                page.wait_for_timeout(2000)
                
                # 获取提交后的页面信息
                after_url = page.url
                page_content = page.content()
                page_text = page_content.lower()
                
                # 分析后端返回结果
                # 1. 检查URL变化
                url_changed = before_url != after_url
                is_success_url = "success" in after_url.lower() or "welcome" in after_url.lower()
                is_error_url = "error" in after_url.lower()
                
                # 2. 从页面中提取错误/成功消息（通常是后端渲染的）
                error_messages = []
                success_messages = []
                
                # 常见的后端错误消息选择器
                error_selectors = [
                    '.error-message', '.alert-error', '.text-danger', '.error',
                    '[class*="error"]', '[class*="danger"]', 
                    '.el-message--error', '.ant-message-error',
                    '#error-msg', '.tips-error', '.form-error',
                    '.invalid-feedback', '.help-block'
                ]
                
                # 常见的后端成功消息选择器
                success_selectors = [
                    '.success-message', '.alert-success', '.text-success', '.success',
                    '[class*="success"]', '.el-message--success', '.ant-message-success',
                    '#success-msg', '.tips-success', '.form-success',
                    '.valid-feedback'
                ]
                
                for selector in error_selectors:
                    try:
                        elements = page.query_selector_all(selector)
                        for el in elements:
                            text = el.inner_text()
                            if text and text.strip():
                                error_messages.append(text.strip())
                    except:
                        pass
                
                for selector in success_selectors:
                    try:
                        elements = page.query_selector_all(selector)
                        for el in elements:
                            text = el.inner_text()
                            if text and text.strip():
                                success_messages.append(text.strip())
                    except:
                        pass
                
                # 3. 检查页面文本中的关键词
                error_keywords = ["错误", "失败", "invalid", "error", "fail", "不能为空", "格式不正确", 
                                 "非法", "不合法", "请输入", "格式错误", "不正确", "有误", "不符合",
                                 "already exists", "已存在", "重复"]
                success_keywords = ["成功", "success", "保存成功", "修改成功", "注册成功", "提交成功", 
                                   "已完成", "操作成功", "welcome", "欢迎"]
                
                found_error_keywords = [kw for kw in error_keywords if kw in page_text]
                found_success_keywords = [kw for kw in success_keywords if kw in page_text]
                
                # 构建后端返回结果描述
                result_parts = []
                
                if error_messages:
                    result_parts.append(f"后端错误提示: {'; '.join(error_messages[:2])}")
                if success_messages:
                    result_parts.append(f"后端成功提示: {'; '.join(success_messages[:2])}")
                if found_error_keywords:
                    result_parts.append(f"页面含错误关键词: {', '.join(found_error_keywords[:5])}")
                if found_success_keywords:
                    result_parts.append(f"页面含成功关键词: {', '.join(found_success_keywords[:5])}")
                
                # 添加URL信息
                if url_changed:
                    result_parts.append(f"URL变化: {before_url} -> {after_url}")
                else:
                    result_parts.append("URL未变化(仍在原页面)")
                
                if result_parts:
                    api_response = " | ".join(result_parts)
                else:
                    api_response = f"页面正常加载 | URL: {after_url}"
                
                # 判断校验是否通过
                has_error_msg = len(error_messages) > 0 or len(found_error_keywords) > 0
                has_success_msg = len(success_messages) > 0 or len(found_success_keywords) > 0
                
                if has_error_msg:
                    validation_passed = "是"  # 后端返回了错误信息，说明有校验
                elif has_success_msg or is_success_url:
                    validation_passed = "否"  # 后端返回成功或跳转到成功页面，说明没有拦截
                elif not url_changed:
                    validation_passed = "是"  # 还在原页面，可能是被阻止了
                elif is_error_url:
                    validation_passed = "是"
                else:
                    validation_passed = "待确认"
                    
            except Exception as e:
                api_response = f"请求异常: {str(e)}"
                validation_passed = "异常"
                page.wait_for_timeout(1000)
        
        test_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        excel_row = [
            seq_id,
            test_time,
            user_data["username"],
            field_name,
            original_value,
            test_value if len(test_value) <= 100 else test_value[:100] + "...",
            test_type,
            api_response if len(api_response) <= 500 else api_response[:500] + "...",
            validation_passed
        ]
        
        save_to_excel(excel_row)
        results.append({
            "field": field_name,
            "type": test_type,
            "value": test_value,
            "response": api_response,
            "validation": validation_passed
        })
        
        seq_id += 1
        
        if original_value:
            try:
                field_input.fill("")
                page.wait_for_timeout(200)
                field_input.fill(original_value)
                page.wait_for_timeout(300)
            except:
                pass
    
    return results, seq_id


def single_task(task_id, user_data, start_seq_id):
    print(f"\n[任务{task_id}] 开始执行...")
    print(f"[任务{task_id}] 用户名: {user_data['username']}")
    
    seq_id = start_seq_id
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        page = context.new_page()
        
        try:
            print(f"[任务{task_id}] 正在打开网站...")
            page.goto(f"{BASE_URL}/", timeout=30000)
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(2000)
            
            print(f"[任务{task_id}] 查找注册链接...")
            register_link = page.query_selector('text=注册') or page.query_selector('text=立即注册') or page.query_selector('a:has-text("注册")')
            
            if register_link:
                print(f"[任务{task_id}] 点击注册链接...")
                register_link.click()
                page.wait_for_load_state("networkidle")
                page.wait_for_timeout(1000)
            
            register_success = perform_register(page, user_data, task_id)
            
            if not register_success:
                print(f"[任务{task_id}] 注册失败，跳过后续操作")
                return {"task_id": task_id, "status": "注册失败", "seq_id": seq_id}
            
            print(f"[任务{task_id}] 注册成功，开始登录...")
            login_success = perform_login(page, user_data, task_id)
            
            if not login_success:
                print(f"[任务{task_id}] 登录失败，跳过后续操作")
                return {"task_id": task_id, "status": "登录失败", "seq_id": seq_id}
            
            print(f"[任务{task_id}] 登录成功，查找可编辑字段...")
            
            # 首先检查当前页面是否有可编辑字段
            edit_selectors = [
                'input[name="name"]', 'input[name="email"]', 'input[name="age"]', 'input[name="phone"]',
                '#name', '#email', '#age', '#phone',
                'input[placeholder*="姓名"]', 'input[placeholder*="邮箱"]', 'input[placeholder*="年龄"]', 'input[placeholder*="手机"]'
            ]
            has_editable_field = False
            for selector in edit_selectors:
                try:
                    element = page.query_selector(selector)
                    if element:
                        has_editable_field = True
                        print(f"[任务{task_id}] 当前页面有可编辑字段: {selector}")
                        break
                except:
                    continue
            
            # 如果当前页面没有可编辑字段，尝试导航到个人信息页面
            if not has_editable_field:
                profile_success = navigate_to_profile(page, task_id)
                if not profile_success:
                    print(f"[任务{task_id}] 未能找到个人信息编辑页面，返回注册页面进行字段校验测试")
                    # 返回注册页面进行字段校验测试
                    try:
                        page.goto(f"{BASE_URL}/", timeout=30000)
                        page.wait_for_load_state("networkidle")
                        page.wait_for_timeout(2000)
                        register_link = page.query_selector('text=注册') or page.query_selector('text=立即注册') or page.query_selector('a:has-text("注册")')
                        if register_link:
                            register_link.click()
                            page.wait_for_load_state("networkidle")
                            page.wait_for_timeout(1000)
                        print(f"[任务{task_id}] 已返回注册页面进行字段校验测试")
                    except Exception as e:
                        print(f"[任务{task_id}] 返回注册页面失败: {e}")
            
            print(f"[任务{task_id}] 开始字段校验测试...")
            validation_results, seq_id = perform_field_validation(page, user_data, task_id, seq_id)
            
            print(f"\n[任务{task_id}] ========== 执行完成 ==========")
            print(f"[任务{task_id}] 注册状态: 成功")
            print(f"[任务{task_id}] 登录状态: 成功")
            print(f"[任务{task_id}] 校验测试数: {len(validation_results)}")
            print(f"[任务{task_id}] ==============================\n")
            
            return {
                "task_id": task_id,
                "status": "成功",
                "validation_count": len(validation_results),
                "seq_id": seq_id
            }
            
        except Exception as e:
            print(f"[任务{task_id}] 发生错误: {e}")
            return {
                "task_id": task_id,
                "status": "失败",
                "error": str(e),
                "seq_id": seq_id
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


def run_validation_test(num_tasks=3):
    init_excel()
    
    print("=" * 60)
    print(f"开始执行 {num_tasks} 个字段校验测试任务")
    print("=" * 60)
    
    overall_start_time = datetime.now()
    
    users_data = [generate_user_data() for _ in range(num_tasks)]
    
    print("\n生成的用户信息:")
    for i, user in enumerate(users_data, 1):
        print(f"  任务{i}: {user['username']}")
    
    results = []
    seq_id = 1
    
    for i, user in enumerate(users_data):
        task_id = i + 1
        try:
            result = single_task(task_id, user, seq_id)
            results.append(result)
            seq_id = result.get("seq_id", seq_id)
        except Exception as e:
            print(f"任务{task_id}执行异常: {e}")
            results.append({"task_id": task_id, "status": "异常", "error": str(e), "seq_id": seq_id})
    
    overall_end_time = datetime.now()
    overall_duration = (overall_end_time - overall_start_time).total_seconds()
    
    print("\n" + "=" * 60)
    print("所有任务执行完成!")
    print("=" * 60)
    
    success_count = sum(1 for r in results if r.get("status") == "成功")
    fail_count = num_tasks - success_count
    total_validation = sum(r.get("validation_count", 0) for r in results)
    
    print(f"\n执行统计:")
    print(f"  总任务数: {num_tasks}")
    print(f"  成功: {success_count}")
    print(f"  失败: {fail_count}")
    print(f"  总校验测试数: {total_validation}")
    print(f"  总耗时: {overall_duration:.2f}秒")
    
    print(f"\n数据已保存到: {EXCEL_FILE}")
    
    return results


if __name__ == "__main__":
    results = run_validation_test(3)
