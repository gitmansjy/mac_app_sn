import os
import re
import shutil
import base64
import requests
from datetime import datetime
from openpyxl import Workbook
import tkinter as tk
from tkinter import filedialog, messagebox

# ---------- 百度 OCR 配置 ----------
BAIDU_API_KEY = "HdKtbeKb7WcotsyumdfpvJUQ"
BAIDU_SECRET_KEY = "3KsAG9bpG99mHJdyCHqOeb0WSY630kaw"

# ---------- 飞书表格配置 ----------
FEISHU_APP_ID = "cli_a8591d0b32cd500e"
FEISHU_APP_SECRET = "vRSzhgzSIVk2jNKl5ixRabOtQhhiqmrv"
SPREADSHEET_TOKEN = "TX1ystYw1hZ90jtaVKxcBcOmnic"
SHEET_ID_SN = "1llBlN"
SHEET_ID_REG = "0QDKnP"

# ---------- 目录配置 ----------
# 默认目录：macOS 下设为 ~/Documents/SN识别；Windows 下仍可保留 D:\SN识别，但用户会通过弹窗选择
DEFAULT_DIR = os.path.expanduser("~/Documents/SN识别")  # macOS 用户文档目录

# 支持的图片扩展名
IMAGE_EXT = (".png", ".jpg", ".jpeg", ".bmp")
OCR_API_STANDARD = "general_basic"
OCR_API_ACCURATE = "accurate_basic"


# ---------- 函数定义 ----------
def select_directory():
    """弹出目录选择对话框，返回用户选择的路径"""
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    root.attributes("-topmost", True)  # 置顶
    dir_path = filedialog.askdirectory(
        title="请选择包含SN图片的文件夹", initialdir=DEFAULT_DIR
    )
    root.destroy()
    return dir_path


def get_baidu_access_token():
    url = "https://aip.baidubce.com/oauth/2.0/token"
    params = {
        "grant_type": "client_credentials",
        "client_id": BAIDU_API_KEY,
        "client_secret": BAIDU_SECRET_KEY,
    }
    return requests.post(url, params=params).json().get("access_token")


def get_feishu_tenant_access_token(app_id, app_secret):
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    headers = {"Content-Type": "application/json; charset=utf-8"}
    payload = {"app_id": app_id, "app_secret": app_secret}
    resp = requests.post(url, headers=headers, json=payload).json()
    return resp.get("tenant_access_token")


def read_feishu_sheet(access_token, spreadsheet_token, sheet_id, range_str):
    url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/values/{sheet_id}!{range_str}"
    headers = {"Authorization": f"Bearer {access_token}"}
    resp = requests.get(url, headers=headers).json()
    if resp.get("code") != 0:
        print(f"读取飞书失败: {resp.get('msg')} (code {resp.get('code')})")
        return []
    value_range = resp.get("data", {}).get("valueRange", {})
    return value_range.get("values", [])


def extract_sn(ocr_json):
    words_list = [item["words"] for item in ocr_json.get("words_result", [])]
    for i, w in enumerate(words_list):
        if "S/N:" in w or "SN:" in w:
            if ":" in w:
                parts = w.split(":", 1)
                if len(parts) > 1 and parts[1].strip():
                    sn_candidate = parts[1].strip().replace(" ", "")
                    if sn_candidate:
                        return sn_candidate
            for j in range(i + 1, min(i + 10, len(words_list))):
                next_word = words_list[j].strip()
                skip_words = {
                    "净",
                    "毛",
                    "重",
                    "产品名称",
                    "包装尺寸",
                    "生产者",
                    "生产地址",
                    "执行标准",
                    "企业网址",
                    "合格证",
                    "S/N",
                    "SN",
                }
                if next_word in skip_words:
                    continue
                if re.match(r"^[A-Z0-9\s]+$", next_word, re.I) and len(next_word) >= 8:
                    sn_candidate = next_word.replace(" ", "")
                    return sn_candidate
            break
    pattern = r"[A-Z0-9]{10,20}"
    for w in words_list:
        match = re.search(pattern, w.replace(" ", ""))
        if match:
            return match.group()
    return None


def ocr_image(baidu_token, image_base64, api_type):
    url = f"https://aip.baidubce.com/rest/2.0/ocr/v1/{api_type}?access_token={baidu_token}"
    payload = {
        "image": image_base64,
        "detect_direction": "false",
        "detect_language": "false",
        "paragraph": "false",
        "probability": "false",
    }
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    try:
        resp = requests.post(url, headers=headers, data=payload, timeout=30)
        return resp.json()
    except Exception as e:
        print(f"OCR请求异常 ({api_type}): {e}")
        return None


def build_sn_to_shipment(sheet_data):
    mapping = {}
    for row in sheet_data:
        if len(row) >= 2 and row[0] and row[1]:
            shipment = str(row[0]).strip()
            sn = str(row[1]).strip()
            mapping[sn] = shipment
    return mapping


def build_shipment_to_platform(sheet_data):
    mapping = {}
    for i, row in enumerate(sheet_data):
        if i == 0:
            continue
        if len(row) >= 3 and row[1] and row[2]:
            platform = str(row[1]).strip()
            shipment = str(row[2]).strip()
            mapping[shipment] = platform
    return mapping


def process_image(
    image_path, sn_to_ship, ship_to_platform, baidu_token, target_date_dir, fail_records
):
    filename = os.path.basename(image_path)
    ext = os.path.splitext(image_path)[1].lower()
    try:
        with open(image_path, "rb") as f:
            img_data = f.read()
        image_base64 = base64.b64encode(img_data).decode("utf-8")

        print(f"正在处理: {filename} (标准版)")
        ocr_result = ocr_image(baidu_token, image_base64, OCR_API_STANDARD)
        sn = extract_sn(ocr_result) if ocr_result else None

        if not sn:
            print(f"标准版未识别到SN，切换至高精版重试: {filename}")
            ocr_result = ocr_image(baidu_token, image_base64, OCR_API_ACCURATE)
            sn = extract_sn(ocr_result) if ocr_result else None

        if not sn:
            fail_records.append(
                [filename, "未识别到SN", datetime.now().strftime("%Y-%m-%d %H:%M:%S")]
            )
            return False

        if sn not in sn_to_ship:
            fail_records.append(
                [
                    filename,
                    f"SN [{sn}] 未在表2中找到对应发运单号",
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                ]
            )
            return False
        shipment_no = sn_to_ship[sn]

        if shipment_no not in ship_to_platform:
            fail_records.append(
                [
                    filename,
                    f"发运单号 [{shipment_no}] 未在表1中找到对应平台单号",
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                ]
            )
            return False
        platform_no = ship_to_platform[shipment_no]

        os.makedirs(target_date_dir, exist_ok=True)
        new_name = platform_no + ext
        new_path = os.path.join(target_date_dir, new_name)
        shutil.move(image_path, new_path)
        print(f"✅ 成功: {filename} -> {new_path}")
        return True

    except Exception as e:
        fail_records.append(
            [
                filename,
                f"处理异常: {str(e)}",
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            ]
        )
        return False


def main():
    print("=== SN 图片批量处理工具 ===")
    # 选择文件夹
    base_dir = select_directory()
    if not base_dir:
        print("未选择文件夹，程序退出")
        messagebox.showwarning("提示", "未选择文件夹，程序将退出")
        return

    # 定义成功目录和失败日志路径（基于所选文件夹）
    success_dir = os.path.join(base_dir, "成功")
    fail_log_path = os.path.join(base_dir, "失败日志.xlsx")

    # 确保基础目录存在
    os.makedirs(success_dir, exist_ok=True)

    # 获取 token
    print("获取百度OCR token...")
    baidu_token = get_baidu_access_token()
    if not baidu_token:
        print("获取百度 token 失败")
        messagebox.showerror("错误", "获取百度 token 失败")
        return

    print("获取飞书 token...")
    feishu_token = get_feishu_tenant_access_token(FEISHU_APP_ID, FEISHU_APP_SECRET)
    if not feishu_token:
        print("获取飞书 token 失败")
        messagebox.showerror("错误", "获取飞书 token 失败")
        return

    print("读取表2（SN数据）...")
    sn_sheet = read_feishu_sheet(feishu_token, SPREADSHEET_TOKEN, SHEET_ID_SN, "A:B")
    if not sn_sheet:
        print("表2无数据")
        messagebox.showerror("错误", "飞书表格表2无数据")
        return
    sn_to_ship = build_sn_to_shipment(sn_sheet)
    print(f"表2中共有 {len(sn_to_ship)} 条 SN 映射记录")

    print("读取表1（SN登记表）...")
    reg_sheet = read_feishu_sheet(feishu_token, SPREADSHEET_TOKEN, SHEET_ID_REG, "A:E")
    if not reg_sheet:
        print("表1无数据")
        messagebox.showerror("错误", "飞书表格表1无数据")
        return
    ship_to_platform = build_shipment_to_platform(reg_sheet)
    print(f"表1中共有 {len(ship_to_platform)} 条发运单号映射记录")

    date_str = datetime.now().strftime("%Y-%m-%d")
    target_date_dir = os.path.join(success_dir, date_str)

    # 收集图片文件
    image_files = []
    for file in os.listdir(base_dir):
        if file.lower().endswith(IMAGE_EXT):
            full_path = os.path.join(base_dir, file)
            if os.path.isfile(full_path):
                image_files.append(full_path)

    print(f"共找到 {len(image_files)} 张图片，开始处理...")
    fail_records = []
    for img_path in image_files:
        process_image(
            img_path,
            sn_to_ship,
            ship_to_platform,
            baidu_token,
            target_date_dir,
            fail_records,
        )

    if fail_records:
        wb = Workbook()
        ws = wb.active
        ws.title = "失败日志"
        ws.append(["文件名", "失败原因", "处理时间"])
        for row in fail_records:
            ws.append(row)
        wb.save(fail_log_path)
        print(f"失败日志已保存至: {fail_log_path} (共 {len(fail_records)} 条)")
        messagebox.showinfo(
            "完成",
            f"处理完成！成功: {len(image_files)-len(fail_records)}，失败: {len(fail_records)}\n失败日志保存至: {fail_log_path}",
        )
    else:
        print("所有图片处理成功，无失败记录")
        messagebox.showinfo("完成", "所有图片处理成功！")


if __name__ == "__main__":
    main()
