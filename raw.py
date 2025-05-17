from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains

from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options


import time
import os
import sys
import re
import json
import msvcrt





script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
os.chdir(script_dir)


# 中文-代码映射
mapping = {
    "训练机型": {"单发陆地": "SL_FLY", "单发陆地-FTD/FFS": "SL_FTD", "多发陆地": "ML_FLY", "多发陆地-FTD/FFS": "ML_FTD", "高性能": "HPA_FLY", "高性能-FTD/FFS": "HPA_FTD"},
    "机型": {"TB-10": "AM_01", "TB-20": "AM_02", "CAP-10": "AM_03", "Beech58": "AM_04", "Beech76": "AM_05", "Decathlon 8KCAB": "AM_06", "C-152": "AM_07", "C-172": "AM_08", "172S": "AM_08",
           "PA-34": "AM_09", "PA-44": "AM_10", "DA-42": "AM_11", "DA-20": "AM_12", "DA-40": "AM_13", "Kingair C-90": "AM_14", "C-525 CJ1+": "AM_15", "SR-20": "AM_16",
           "TECNAM": "AM_17", "MA600": "AM_18", "C-525 M2": "AM_19", "其他机型": "AM_20"},
    "进近方式": {"目视": "目视（visual）", "目视Visual": "目视（visual）" },  # 这里实际选择用点击实现
    "仪表类型": {"真实仪表": "R", "模拟仪表": "S", "无": "无"},
    "训练类型": {"日常训练": "T", "训练": "T", "检查/考试": "E", "检查": "E"},
    "起飞场站": {"ZULP": "重庆梁平机场"},
    "着陆场站": {"ZULP": "重庆梁平机场"},

    # 次数类，直接填数字字符串
    "日间着陆次数": None,       # 对应 name="dayLandTimes"
    "夜间着陆次数": None,       # 对应 name="nightLandTimes"

    # 时间类，格式 'YYYY-MM-DD hh:ii'
    "出发时刻": None,           # 对应 name="takeoffTime"
    "到达时刻": None,           # 对应 name="landingTime"
    
   

    # 时分下拉框：直接从 data 读取 "hh:mm" 再拆分
    "转场时间": None,           # id="transitionsTimeDiv"
    "特技时间": None,           # id="spicTimeDiv"
    "仪表时间": None,           # id="meterTimeDiv"
    "单飞时间": None,           # id="soloFlyTimeDiv"
    "带飞时间": None,           # id="instrTimeDiv"
    "昼间训练时间": None,       # id="dayTrainTimeDiv"

    "夜间训练时间": None,       # id="nightTrainTimeDiv"

     # 日期类，格式 'YYYY-MM-DD'
    "训练日期": None,           # 对应 name="trainingDate"
}



def get_resource_path(relative_path):
    """获取资源文件的绝对路径，兼容PyInstaller打包"""
    # if hasattr(sys, '_MEIPASS'):
    #     # 运行时，从临时解压目录获取资源
    #     return os.path.join(sys._MEIPASS, relative_path)
    # else:
    #     # 未打包时，直接从源代码目录获取
    #     return os.path.join(os.path.abspath("."), relative_path)

    """获取资源文件的绝对路径（exe 同目录）"""
    if getattr(sys, 'frozen', False):  # 如果是打包后的 exe
        base_path = os.path.dirname(sys.executable)
    else:  # 没打包时
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)



    return os.path.join(os.path.abspath("."), relative_path)

def get_driver():
    from selenium import webdriver
    from selenium.common.exceptions import WebDriverException

    
   
    options = webdriver.ChromeOptions()
    options.add_argument('--log-level=3')  # 只显示严重错误
    options.add_argument("--remote-debugging-port=0")  # 禁用调试端口
    options.debugger_address = "127.0.0.1:9222"

    try:
        
        chromedriver_path = get_resource_path("tools/chromedriver.exe")
        # 输出驱动目录
        print("chromedriver.exe 路径：", chromedriver_path)


       

        # 检查 chromedriver 是否存在
        if not os.path.exists(chromedriver_path):
            print(f"找不到 chromedriver.exe，请确认路径：{chromedriver_path}")
            return None
        else:
            print(f"找到 chromedriver.exe，路径：{chromedriver_path}")

        service = Service(executable_path=chromedriver_path)
        driver = webdriver.Chrome(service=service, options=options)
        print("已成功连接到调试中的 Chrome 浏览器。")
        return driver
    except WebDriverException as e:
        print(" 无法连接到正在调试的 Chrome 实例，请确认是否已正确启动 Chrome：")
        print("示例命令：")
        print('chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\\selenum\\profile"')
        print(f"错误详情：{e}")
        return None





def read_iFly(driver):
    # 输出网页标题
    print("网页标题:", driver.title)
    wait = WebDriverWait(driver, 5)
    
    def wait_for_detail_page():
        while True:
            try:
                wait.until(lambda d: d.title == "飞行记录 - 详情 - iFly.Top")
                print("成功跳转到飞行记录详情页")
                break  # 成功跳转，跳出循环
            except Exception as e:
                print("跳转到飞行记录详情页失败（页面标题不符）。")
                print("请手动打开详情页，完成后按回车继续，或按 ESC 键退出。")
                time.sleep(0.2)  # 防止 getch 自动触发
                while True:
                    key = msvcrt.getch()
                    if key == b'\r':  # 回车
                        break  # 再次尝试 wait
                    elif key == b'\x1b':  # ESC
                        print("\n程序已退出。")
                        sys.exit(0)

    # 初始提示
    print("打开需要录入的飞行记录，页面加载完成后按回车继续...（按 ESC 键退出）")
    while True:
        key = msvcrt.getch()
        if key == b'\r':
            break
        elif key == b'\x1b':
            print("\n程序已退出。")
            sys.exit(0)

    # 等待详情页加载（带失败重试机制）
    wait_for_detail_page()


    # 等待详情页 div.details 出现
    details_div = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.details")))

    # 解析 div.details 下所有 <dd> 标签
    dd_elements = details_div.find_elements(By.TAG_NAME, "dd")

    data = {}
    for dd in dd_elements:
        try:
            title_span = dd.find_element(By.CSS_SELECTOR, "span.title")
            title = title_span.text.strip().rstrip("：")
            # 获取 title_span 后的文本节点内容
            # Selenium 里无法直接拿 next_sibling，需要拿整个 dd 文本减去标题
            full_text = dd.text.strip()
            value = full_text.replace(title_span.text, '').strip()
            data[title] = value
        except Exception as e:
            #print(f"读取某字段失败: {e}")
            continue

    # print("\n飞行记录详情：")
    # for k, v in data.items():
    #     print(f"{k}：{v}")

    # TODO: 这里需要根据实际的映射关系进行调整
    # 训练机型固定写死
    training_plane = "单发陆地"

    # 从网页数据中提取机型和登记号，机型格式示例："C-172 / B-106U"
    plane_raw = ""
    for key in data.keys():
        if "机型" in key and "机号" in key:
            plane_raw = data[key]
            break

    plane_type, plane_number = ("", "")
    if plane_raw:
        if '／' in plane_raw:  # 中文斜杠
            parts = plane_raw.split('／')
        elif '/' in plane_raw:  # 英文斜杠
            parts = plane_raw.split('/')
        else:
            parts = [plane_raw]

        if len(parts) >= 2:
            plane_type = parts[0].strip()
            plane_number = parts[1].strip()
        else:
            plane_type = plane_raw.strip()

    # 教员姓名 (仅仅保留中文字符)
    instructor_name = data.get("教员", "")
    instructor_name = "".join(re.findall(r"[\u4e00-\u9fa5]", instructor_name))

    # 进近方式
    approach = data.get("进近方式", "")

    # 仪表类型，网页里字段名可能为“模拟/真实仪表”
    instrument_type = data.get("模拟/真实仪表", "")


    def time_str_to_minutes(time_str):
        try:
            hours, minutes = map(int, time_str.split(':'))
            return hours * 60 + minutes
        except:
            return 0

    def minutes_to_str(total_minutes):
        h = total_minutes // 60
        m = total_minutes % 60
        return f"{h:02d}:{m:02d}"

    # 昼夜训练时间，网页字段“昼 / 夜时长”，格式示例："01:20 / 00:00"
    duration_raw = data.get("昼 / 夜时长", "")
    day_duration, night_duration = ("00:00", "00:00")
    if '/' in duration_raw:
        parts = duration_raw.split('/')
        day_duration = parts[0].strip()
        night_duration = parts[1].strip()

    # 总飞行时长 （昼夜训练时间之和）
    total_minutes = time_str_to_minutes(day_duration) + time_str_to_minutes(night_duration)
    total_time_str = minutes_to_str(total_minutes)


    # 训练类型，网页字段可能是“飞行类型”，只取第一个词
    training_type_raw = data.get("飞行类型", "")
    training_type = training_type_raw.split('/')[0].strip() if training_type_raw else ""

    is_solo = "单飞" in training_type_raw
    is_instructor = "带飞" in training_type_raw

    # 出发和到达时间，示例："2025-05-01 08:30 ~ 2025-05-01 09:45"
    time_raw = data.get("出发~到达时间", "")
    start_time, end_time = ("", "")
    total_minutes = 0
    if '~' in time_raw:
        parts = time_raw.split('~')
        start_time = parts[0].strip()
        end_time = parts[1].strip()
        

    # 分配单飞和带飞时长
    def minutes_to_str(minutes):
        h = minutes // 60
        m = minutes % 60
        return f"{h:02}:{m:02}"

    if is_solo and not is_instructor:
        solo_time = total_time_str
        instructor_time = "00:00"
    elif is_instructor and not is_solo:
        solo_time = "00:00"
        instructor_time = total_time_str
    else:
        solo_time = "00:00"
        instructor_time = "00:00"


    # 起飞和着陆场站，网页字段示例“起飞~降落机场”： "重庆梁平机场 ~ 重庆梁平机场"
    airport_raw = data.get("起飞~降落机场", "")
    airport_start, airport_end = ("", "")
    if '~' in airport_raw:
        parts = airport_raw.split('~')
        airport_start = parts[0].strip()
        airport_end = parts[1].strip()

    # 备注
    #remark = data.get("备注", "")

    # 日间夜间着陆次数，网页字段示例 "昼 / 夜起落": "2 / 1"
    landing_raw = data.get("昼 / 夜起落", "")
    day_landings, night_landings = ("0", "0")
    if '/' in landing_raw:
        parts = landing_raw.split('/')
        day_landings = parts[0].strip()
        night_landings = parts[1].strip()


    # 训练日期
    training_date = data.get("飞行日期", "")

    # 转场时间，仪表时间，特技时间(默认00:00)
    # 网页字段“仪表/转场/螺旋”，格式示例："00:05 / 00:10 / 00:00"
    time_parts_raw = data.get("仪表/转场/螺旋", "")
    instrument_time, transit_time, stunt_time = ("00:00", "00:00", "00:00")
    if time_parts_raw:
        parts = [x.strip() for x in time_parts_raw.split('/')]
        if len(parts) == 3:
            instrument_time, transit_time, stunt_time = parts


    # 组装输出
    output_text = f"""训练机型：{training_plane}
    机型：{plane_type}
    教员姓名：{instructor_name}
    航空器登记号：{plane_number}
    进近方式：{approach}
    仪表类型：{instrument_type}
    训练类型：{training_type}
    起飞场站：{airport_start}
    着陆场站：{airport_end}
    备注：{"正常训练"}

    日间着陆次数：{day_landings}
    夜间着陆次数：{night_landings}

    出发时刻：{start_time}
    到达时刻：{end_time}

    训练日期：{training_date}

    转场时间：{transit_time}
    特技时间：{stunt_time}
    仪表时间：{instrument_time}
    单飞时间：{solo_time}
    带飞时间：{instructor_time}
    昼间训练时间：{day_duration}
    夜间训练时间：{night_duration}
    """

    print("\n整理后的飞行记录数据：\n")
    print(output_text)
    # print("当前工作目录是：", os.getcwd())

    try:
        with open("data.txt", "w", encoding="utf-8") as f:
            f.write(output_text)
        print("数据已成功写入 data.txt 文件。")
    except Exception as e:
        print(f"写入文件时发生错误：{e}")

    driver.quit()



def read_data(filename):
    data = {}
    with open(filename, encoding='utf-8') as f:
        for line in f:
            if "：" in line:
                key, value = line.strip().split("：", 1)
                data[key] = value

    # print("读取的数据：")
    # for key, value in data.items():
    #     print(f"{key}: {value}")

    return data



def fill_form(data):
    
    options = Options()
    options.add_argument('--log-level=3')  # 只显示严重错误
    options.add_argument("--remote-debugging-port=0")  # 禁用调试端口
    # # 启动浏览器
    driver = webdriver.Chrome(options=options)

    config_path = get_resource_path("config.json")

    # 检查配置文件是否存在
    if not os.path.exists(config_path):
        print(f"配置文件 config.json 不存在，请确认路径：{config_path}")
        driver.quit()
        return

    with open(config_path, "r", encoding="utf-8") as f:
        config = json.load(f)

    url = config.get("url", "")
    driver.get(url)

    #driver.get("https://pilot.ceair.com/flyingnet-web-portal/portal_students/students_dailyPaper.do?idCode=MDMxMDk3Ng==")

    # 检查title是否为 训练任务上报
    try:
        WebDriverWait(driver, 10).until(EC.title_contains("训练任务上报"))
        print("成功跳转到训练任务上报页面")
    except Exception as e:
        print("跳转到训练任务上报页面失败:", e)
        driver.quit()
        return
    
    # 等待页面加载
    #time.sleep(3)

    # 点击加号按钮
    button = driver.find_element(By.XPATH, "//button[@class='contTextCls' and contains(@onclick, 'newAndEdit')]")
    button.click()
    wait = WebDriverWait(driver, 10)

    
    # 训练日期 trainingDate
    try:
        val = data.get("训练日期", "") 
        training_input = driver.find_element(By.NAME, "trainingDate")
        driver.execute_script("arguments[0].removeAttribute('readonly')", training_input)
        training_input.clear()
        training_input.send_keys(val)
    except Exception as e:
        print(f"训练日期输入失败: {e}")


    # 训练机型 select，页面名称可能是trainModel或trainType，注意确认
    train_model_name = "trainModel" if driver.find_elements(By.NAME, "trainModel") else "trainType"
    try:
        train_model_select = Select(wait.until(EC.presence_of_element_located((By.NAME, train_model_name))))
        train_model_value = mapping["训练机型"].get(data.get("训练机型"), "")
        if train_model_value:
            train_model_select.select_by_value(train_model_value)
    except Exception as e:
        print(f"训练机型选择失败: {e}")

    # 机型 select
    try:
        ac_type_select = Select(wait.until(EC.presence_of_element_located((By.NAME, "acTypeCode"))))
        ac_type_value = mapping["机型"].get(data.get("机型"), "")
        if ac_type_value:
            ac_type_select.select_by_value(ac_type_value)
    except Exception as e:
        print(f"机型选择失败: {e}")

    # 教员姓名 input
    try:
        instructor_input = wait.until(EC.presence_of_element_located((By.NAME, "instructorName")))
        instructor_input.clear()
        instructor_input.send_keys(data.get("教员姓名", ""))
    except Exception as e:
        print(f"教员姓名输入失败: {e}")

    # 航空器登记号 input
    try:
        sn_fly_input = wait.until(EC.presence_of_element_located((By.NAME, "snFly")))
        sn_fly_input.clear()
        sn_fly_input.send_keys(data.get("航空器登记号", ""))
    except Exception as e:
        print(f"航空器登记号输入失败: {e}")

    # 进近方式 - 点击下拉展开，再选择对应项
    try:
        dropdown_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-toggle='dropdown']")))
        dropdown_button.click()
        time.sleep(0.5)
        #flying_mode = data.get("进近方式", "")

        raw_value = data.get("进近方式", "").strip()


        # TODO: 这里需要根据实际的映射关系进行调整
        mapped_value = mapping.get("进近方式", {}).get(raw_value, raw_value)  
        # 如果映射值不是“目视”或“目视（Visual）”，则设置为“其他”

        if mapped_value != "目视（visual）":
            mapped_value = "其他（other）"

        option_xpath = f"//span[@class='text' and contains(text(), '{mapped_value}')]"
        option_element = wait.until(EC.element_to_be_clickable((By.XPATH, option_xpath)))
        option_element.click()
    except Exception as e:
        print(f"进近方式选择失败: {e}")

    # 仪表类型 select
    try:
        meter_type_select = Select(wait.until(EC.presence_of_element_located((By.NAME, "meterType"))))
        meter_value = mapping["仪表类型"].get(data.get("仪表类型"), "")
        if meter_value:
            meter_type_select.select_by_value(meter_value)
    except Exception as e:
        print(f"仪表类型选择失败: {e}")

    # 训练类型 select
    try:
        train_type_select = Select(wait.until(EC.presence_of_element_located((By.NAME, "trainType"))))
        train_type_value = mapping["训练类型"].get(data.get("训练类型"), "")
        if train_type_value:
            train_type_select.select_by_value(train_type_value)
    except Exception as e:
        print(f"训练类型选择失败: {e}")

    # 起飞场站 input
    try:
        takeoff_input = wait.until(EC.presence_of_element_located((By.NAME, "takeoffStation")))
        takeoff_input.clear()
        raw_value = data.get("起飞场站", "").strip()
        mapped_value = mapping.get("起飞场站", {}).get(raw_value, raw_value)  
        takeoff_input.send_keys(mapped_value)
    except Exception as e:
        print(f"起飞场站输入失败: {e}")

    # 着陆场站 input
    try:
        landing_input = wait.until(EC.presence_of_element_located((By.NAME, "landingStation")))
        landing_input.clear()
        raw_value = data.get("着陆场站", "").strip()
        mapped_value = mapping.get("着陆场站", {}).get(raw_value, raw_value)  
        landing_input.send_keys(mapped_value)
    except Exception as e:
        print(f"着陆场站输入失败: {e}")



     # 日间着陆次数 dayLandTimes
    try:
        val = data.get("日间着陆次数", "0")
        day_land_input = driver.find_element(By.NAME, "dayLandTimes")
        day_land_input.clear()
        day_land_input.send_keys(val)
    except Exception as e:
        print(f"日间着陆次数输入失败: {e}")

    # 夜间着陆次数 nightLandTimes
    try:
        val = data.get("夜间着陆次数", "0")
        night_land_input = driver.find_element(By.NAME, "nightLandTimes")
        night_land_input.clear()
        night_land_input.send_keys(val)
    except Exception as e:
        print(f"夜间着陆次数输入失败: {e}")

    # 出发时刻 takeoffTime
    try:
        val = data.get("出发时刻", "")
        takeoff_input = driver.find_element(By.NAME, "takeoffTime")
        driver.execute_script("arguments[0].removeAttribute('readonly')", takeoff_input)
        takeoff_input.clear()
        takeoff_input.send_keys(val)
    except Exception as e:
        print(f"出发时刻输入失败: {e}")

    # 到达时刻 landingTime
    try:
        val = data.get("到达时刻", "")
        landing_input = driver.find_element(By.NAME, "landingTime")
        driver.execute_script("arguments[0].removeAttribute('readonly')", landing_input)
        landing_input.clear()
        landing_input.send_keys(val)
    except Exception as e:
        print(f"到达时刻输入失败: {e}")

    # 点击空白区域关闭可能的弹窗或下拉
    body = driver.find_element(By.TAG_NAME, "body")
    ActionChains(driver).move_to_element_with_offset(body, 0, 0).click().perform()


    # 找到页面 body 元素
    body = driver.find_element(By.TAG_NAME, "body")
    # 在 (0,0) 位置点击一次，或根据需要调整偏移量
    ActionChains(driver).move_to_element_with_offset(body, 0, 0).click().perform()

    
    # ***************** 时间 ***********************
    
    # 通用函数：从 data 读取 HH:MM 并拆分
    def fill_time(div_id, field_name):
        try:
            val = data.get(field_name, "00:00")
            hh, mm = val.split(":", 1)
            container = driver.find_element(By.ID, div_id)
            Select(container.find_element(By.CLASS_NAME, "timeHours")).select_by_visible_text(hh)
            Select(container.find_element(By.CLASS_NAME, "timeMinute")).select_by_visible_text(mm)
        except Exception as e:
            print(f"{field_name} 设置失败: {e}")

    # 转场时间 transitionsTimeDiv
    fill_time("transitionsTimeDiv", "转场时间")
    # 特技时间 spicTimeDiv
    fill_time("spicTimeDiv", "特技时间")
    # 仪表时间 meterTimeDiv
    fill_time("meterTimeDiv", "仪表时间")
    # 单飞时间 soloFlyTimeDiv
    fill_time("soloFlyTimeDiv", "单飞时间")
    # 带飞时间 instrTimeDiv
    fill_time("instrTimeDiv", "带飞时间")
    # 昼间训练时间 dayTrainTimeDiv
    fill_time("dayTrainTimeDiv", "昼间训练时间")
    # 夜间训练时间 nightTrainTimeDiv
    fill_time("nightTrainTimeDiv", "夜间训练时间")
    
    # 备注 textarea
    try:
        remark_input = wait.until(EC.presence_of_element_located((By.NAME, "remark")))
        remark_input.clear()
        remark_input.send_keys(data.get("备注", ""))
    except Exception as e:
        print(f"评语输入失败: {e}")

    #driver.quit()

    # 填写完成
    print("填写完成,请检查数据是否正确。")
    print("关闭东航雏鹰浏览器窗口，在录入下一条记录之前")
    print("请稍后")


def main():

    os.environ["TF_CPP_MIN_LOG_LEVEL"] = "3"  # 只保留FATAL日志
    
    while True:
        driver = get_driver()
        if driver is None:
            print("终止程序。")
            break

        try:
            read_iFly(driver)

            data_file = get_resource_path("data.txt")
            data = read_data(data_file)

            fill_form(data)

        except Exception as e:
            print(f"[错误] 程序运行过程中出错：{e}")

        # finally:
        #     driver.quit()

        print("\n按回车键继续录入下一条飞行记录，按 Esc 退出程序...")
        while True:
            key = msvcrt.getch()
            if key == b'\r':  # 回车键
                driver.quit()
                break  # 继续录入下一条
            elif key == b'\x1b':  # ESC 键
                driver.quit()
                print("\n程序已退出。")
                sys.exit(0)



if __name__ == "__main__":
    main()








