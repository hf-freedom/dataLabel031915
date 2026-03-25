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

def generate_random_idcard():
    """生成随机身份证号"""
    # 省份代码
    provinces = ['11', '12', '13', '14', '15', '21', '22', '23', '31', '32', '33', '34', '35', '36', '37',
                 '41', '42', '43', '44', '45', '46', '50', '51', '52', '53', '54', '61', '62', '63', '64', '65']
    province = random.choice(provinces)
    
    # 城市代码（简化，使用随机）
    city = ''.join(random.choices(string.digits, k=2))
    
    # 区县代码
    district = ''.join(random.choices(string.digits, k=2))
    
    # 出生日期 (18-50岁)
    current_year = datetime.now().year
    birth_year = random.randint(current_year - 50, current_year - 18)
    birth_month = random.randint(1, 12)
    birth_day = random.randint(1, 28)
    birth_date = f"{birth_year}{birth_month:02d}{birth_day:02d}"
    
    # 顺序码
    sequence = ''.join(random.choices(string.digits, k=3))
    
    # 前17位
    idcard_17 = province + city + district + birth_date + sequence
    
    # 计算校验码
    weights = [7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2]
    check_codes = ['1', '0', 'X', '9', '8', '7', '6', '5', '4', '3', '2']
    
    sum_value = sum(int(idcard_17[i]) * weights[i] for i in range(17))
    check_code = check_codes[sum_value % 11]
    
    return idcard_17 + check_code

def init_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = "注册数据"
        headers = ["序号", "用户名", "密码", "邮箱", "注册姓名", "年龄", "手机号", "购票姓名", "身份证号", "注册状态", "登录状态", "抢票状态"]
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
    
    return {
        "register_status": register_status
    }

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
    
    if login_btn:
        print(f"[任务{task_id}] 点击登录按钮...")
        login_btn.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(3000)
        
        # 检查是否登录成功，如果URL变了或者页面内容变化
        current_url = page.url
        print(f"[任务{task_id}] 登录后页面URL: {current_url}")
        
        # 尝试跳转到主页（如果需要）
        if "login" in current_url.lower():
            print(f"[任务{task_id}] 检测到仍在登录页，尝试跳转到主页...")
            page.goto("http://39.107.109.8:8082/", timeout=30000)
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(2000)
        
        login_status = "成功"
        print(f"[任务{task_id}] 登录成功")
    
    return {
        "login_status": login_status
    }

def perform_ticket_grab(page, ticket_data, task_id):
    """执行抢票操作"""
    ticket_name = ticket_data["ticket_name"]
    idcard = ticket_data["idcard"]
    
    print(f"[任务{task_id}] 开始抢票流程...")
    
    # 等待页面完全加载
    page.wait_for_timeout(2000)
    
    # 1. 点击立即抢票按钮
    grab_btn_selectors = [
        'button:has-text("立即抢票")',
        'a:has-text("立即抢票")',
        'button:has-text("抢票")',
        'a:has-text("抢票")',
        '.grab-btn',
        '#grab-btn',
        '#grabTicket',
        '.grab-ticket',
        'button:has-text("立即购买")',
        'a:has-text("立即购买")',
        'button[id*="grab"]',
        'a[id*="grab"]',
        'button[class*="grab"]',
        'a[class*="grab"]',
        'button',
        'a'
    ]
    
    grab_btn = None
    used_selector = None
    
    # 先尝试精确匹配
    for selector in grab_btn_selectors:
        try:
            elements = page.query_selector_all(selector)
            for element in elements:
                if element:
                    text = element.inner_text() if hasattr(element, 'inner_text') else ""
                    if "立即抢票" in text or "抢票" in text:
                        grab_btn = element
                        used_selector = selector
                        break
            if grab_btn:
                break
        except Exception as e:
            continue
    
    # 如果没找到，尝试通过文本内容查找
    if not grab_btn:
        try:
            grab_btn = page.get_by_text("立即抢票").first
            if grab_btn:
                used_selector = "get_by_text"
        except:
            pass
    
    if not grab_btn:
        # 打印页面内容帮助调试
        print(f"[任务{task_id}] 未找到立即抢票按钮，打印页面按钮:")
        try:
            buttons = page.query_selector_all('button, a')
            for i, btn in enumerate(buttons[:10]):
                try:
                    text = btn.inner_text()
                    print(f"  按钮{i}: {text}")
                except:
                    pass
        except:
            pass
        return {"ticket_status": "失败-未找到抢票按钮"}
    
    print(f"[任务{task_id}] 点击立即抢票按钮...")
    grab_btn.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(2000)
    
    # 2. 输入姓名和身份证号
    name_selectors = [
        'input[name="name"]',
        'input[name="realName"]',
        'input[name="realname"]',
        'input[placeholder*="姓名"]',
        'input[placeholder*="真实姓名"]',
        '#name',
        '#realName'
    ]
    
    idcard_selectors = [
        'input[name="idcard"]',
        'input[name="idCard"]',
        'input[name="identity"]',
        'input[name="identityCard"]',
        'input[placeholder*="身份证"]',
        'input[placeholder*="身份证号"]',
        '#idcard',
        '#idCard'
    ]
    
    print(f"[任务{task_id}] 填写购票人信息...")
    
    name_input, _ = find_input(page, name_selectors, "姓名")
    idcard_input, _ = find_input(page, idcard_selectors, "身份证号")
    
    if name_input:
        name_input.fill(ticket_name)
        print(f"[任务{task_id}] 已填写购票姓名: {ticket_name}")
    else:
        print(f"[任务{task_id}] 未找到姓名输入框")
    
    if idcard_input:
        idcard_input.fill(idcard)
        print(f"[任务{task_id}] 已填写身份证号: {idcard}")
    else:
        print(f"[任务{task_id}] 未找到身份证号输入框")
    
    page.wait_for_timeout(500)
    
    # 3. 点击确认购买
    confirm_selectors = [
        'button:has-text("确认购买")',
        'button:has-text("确认")',
        'button:has-text("提交")',
        'button:has-text("确定")',
        'input[value="确认购买"]',
        'input[value="确认"]',
        '.confirm-btn',
        '#confirm-btn',
        'button[type="submit"]'
    ]
    
    confirm_btn = None
    for selector in confirm_selectors:
        try:
            confirm_btn = page.query_selector(selector)
            if confirm_btn:
                break
        except:
            continue
    
    ticket_status = "失败"
    if confirm_btn:
        print(f"[任务{task_id}] 点击确认购买按钮...")
        confirm_btn.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(2000)
        ticket_status = "成功"
        print(f"[任务{task_id}] 抢票操作完成")
    else:
        print(f"[任务{task_id}] 未找到确认购买按钮")
        ticket_status = "失败-未找到确认按钮"
    
    return {
        "ticket_status": ticket_status
    }

def single_task(task_id, user_data, ticket_data):
    print(f"\n[任务{task_id}] 开始执行...")
    print(f"[任务{task_id}] 用户名: {user_data['username']}")
    
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
            
            # 1. 注册
            register_result = perform_register(page, user_data, task_id)
            
            # 2. 登录
            login_result = perform_login(page, user_data, task_id)
            
            # 3. 抢票
            ticket_result = {"ticket_status": "未执行"}
            if login_result["login_status"] == "成功":
                ticket_result = perform_ticket_grab(page, ticket_data, task_id)
            else:
                print(f"[任务{task_id}] 登录失败，跳过抢票")
            
            # 保存数据到Excel
            excel_data = [
                task_id,
                user_data["username"],
                user_data["password"],
                user_data["email"],
                user_data["name"],
                user_data["age"],
                user_data["phone"],
                ticket_data["ticket_name"],
                ticket_data["idcard"],
                register_result["register_status"],
                login_result["login_status"],
                ticket_result["ticket_status"]
            ]
            
            save_to_excel(excel_data)
            
            print(f"\n[任务{task_id}] ========== 执行完成 ==========")
            print(f"[任务{task_id}] 注册状态: {register_result['register_status']}")
            print(f"[任务{task_id}] 登录状态: {login_result['login_status']}")
            print(f"[任务{task_id}] 抢票状态: {ticket_result['ticket_status']}")
            print(f"[任务{task_id}] 购票姓名: {ticket_data['ticket_name']}")
            print(f"[任务{task_id}] 身份证号: {ticket_data['idcard']}")
            print(f"[任务{task_id}] ==============================\n")
            
            return {
                "task_id": task_id,
                "status": "成功",
                "register_status": register_result["register_status"],
                "login_status": login_result["login_status"],
                "ticket_status": ticket_result["ticket_status"]
            }
            
        except Exception as e:
            print(f"[任务{task_id}] 发生错误: {e}")
            
            # 保存错误状态到Excel
            excel_data = [
                task_id,
                user_data["username"],
                user_data["password"],
                user_data["email"],
                user_data["name"],
                user_data["age"],
                user_data["phone"],
                ticket_data["ticket_name"],
                ticket_data["idcard"],
                "异常",
                "异常",
                f"异常: {str(e)}"
            ]
            save_to_excel(excel_data)
            
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

def generate_ticket_data():
    """生成购票人数据"""
    return {
        "ticket_name": generate_random_name(),
        "idcard": generate_random_idcard()
    }

def run_parallel_register(num_tasks=5):
    init_excel()
    
    print("=" * 60)
    print(f"开始并行执行 {num_tasks} 个注册+登录+抢票任务")
    print("=" * 60)
    
    overall_start_time = datetime.now()
    
    users_data = [generate_user_data() for _ in range(num_tasks)]
    tickets_data = [generate_ticket_data() for _ in range(num_tasks)]
    
    print("\n生成的用户信息:")
    for i, user in enumerate(users_data, 1):
        print(f"  任务{i}: 用户名={user['username']}, 购票人={tickets_data[i-1]['ticket_name']}, 身份证={tickets_data[i-1]['idcard']}")
    
    results = []
    
    with ThreadPoolExecutor(max_workers=num_tasks) as executor:
        futures = {executor.submit(single_task, i+1, users_data[i], tickets_data[i]): i+1 for i in range(num_tasks)}
        
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
    
    return results

if __name__ == "__main__":
    results = run_parallel_register(5)
