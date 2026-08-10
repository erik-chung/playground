#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
QQ群成员导出工具 - Web版
连接到 NapCatQQ 的 HTTP API (OneBot v11)
"""
import os
import requests
from datetime import datetime
from pathlib import Path

from flask import Flask, jsonify, send_file
from flask_cors import CORS

try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False


app = Flask(__name__)
CORS(app)

EXPORT_DIR = Path("exports")
EXPORT_DIR.mkdir(exist_ok=True)

ONEBOT_URL = os.environ.get("ONEBOT_URL", "http://127.0.0.1:8081")


def export_to_excel(group_id, group_name, members):
    """导出群成员到Excel - 按入群时间正序，列：QQ号/昵称/群名片/角色/性别/入群时间/最后发言"""
    if not OPENPYXL_AVAILABLE:
        return None, "openpyxl 未安装，请运行：pip install openpyxl"

    members_sorted = sorted(members, key=lambda m: m.get("join_time", 0))

    wb = openpyxl.Workbook()
    ws = wb.active
    wb.remove(ws)
    ws = wb.create_sheet("成员名单")

    header_font = Font(name="微软雅黑", size=11, bold=True)
    header_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000"),
    )

    headers = ["QQ号", "昵称", "群名片", "角色", "入群时间", "最后发言"]
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num, value=header)
        cell.font = header_font
        cell.alignment = header_alignment
        cell.fill = header_fill
        cell.border = thin_border

    data_font = Font(name="微软雅黑", size=10)
    data_alignment = Alignment(horizontal="left", vertical="center")

    role_map = {"owner": "群主", "admin": "管理员", "member": "成员"}

    for row_num, member in enumerate(members_sorted, 2):
        user_id = member.get("user_id", "")
        nickname = member.get("nickname", "")
        card = member.get("card", "")
        display_name = card if card else nickname
        role = role_map.get(member.get("role", "member"), "成员")

        ws.cell(row=row_num, column=1, value=str(user_id))
        ws.cell(row=row_num, column=2, value=nickname)
        ws.cell(row=row_num, column=3, value=display_name)
        ws.cell(row=row_num, column=4, value=role)

        join_time = member.get("join_time")
        if join_time:
            ws.cell(row=row_num, column=5, value=datetime.fromtimestamp(join_time).strftime("%Y-%m-%d %H:%M:%S"))

        last_sent = member.get("last_sent_time")
        if last_sent:
            ws.cell(row=row_num, column=6, value=datetime.fromtimestamp(last_sent).strftime("%Y-%m-%d %H:%M:%S"))

        for col_num in range(1, 7):
            cell = ws.cell(row=row_num, column=col_num)
            cell.font = data_font
            cell.alignment = data_alignment
            cell.border = thin_border

    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['B'].width = 16
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 8
    ws.column_dimensions['E'].width = 20
    ws.column_dimensions['F'].width = 20

    safe_name = "".join(c for c in group_name if c not in r'\/:*?"<>|')
    filename = EXPORT_DIR / f"{safe_name}-{group_id}-{datetime.now().strftime('%Y%m%d-%H%M%S')}.xlsx"

    wb.save(filename)
    return filename, None


def call_onebot_api(action, params=None):
    """调用 OneBot API"""
    try:
        resp = requests.post(
            f"{ONEBOT_URL}/{action}",
            json=params or {},
            timeout=60
        )
        result = resp.json()
        if result.get("retcode") != 0:
            return None, result.get("msg", "API调用失败")
        return result.get("data"), None
    except requests.exceptions.ConnectionError:
        return None, "无法连接到 OneBot API，请确认 NapCatQQ 正在运行"
    except Exception as e:
        return None, str(e)


def check_login():
    data, error = call_onebot_api("get_login_info")
    return error is None


def get_groups():
    """获取群列表（按群主入群时间即群创建时间正序，带群主信息）"""
    data, error = call_onebot_api("get_group_list")
    if error:
        return [], error

    groups = []
    for group in data:
        gid = group.get("group_id")
        owner = None
        create_time = 0
        members_data, err = call_onebot_api("get_group_member_list", {"group_id": gid})
        if not err and isinstance(members_data, list):
            for m in members_data:
                if m.get("role") == "owner":
                    owner = m
                    create_time = m.get("join_time", 0)
                    break

        groups.append({
            "group_id": gid,
            "group_name": group.get("group_name"),
            "member_count": group.get("member_count"),
            "max_member_count": group.get("max_member_count"),
            "avatar_url": f"https://p.qlogo.cn/gh/{gid}/{gid}/100",
            "owner_id": owner.get("user_id") if owner else None,
            "owner_nickname": owner.get("nickname") if owner else "",
            "owner_avatar": f"https://q1.qlogo.cn/g?b=qq&nk={owner.get('user_id')}&s=100" if owner else "",
            "create_time": create_time,
        })

    groups.sort(key=lambda g: g["create_time"] or 0, reverse=True)
    return groups, None


def get_group_members(group_id):
    data, error = call_onebot_api(
        "get_group_member_list",
        {"group_id": group_id, "no_cache": True}
    )
    if error:
        return [], error
    return data, None


@app.route("/")
def index():
    return send_file("templates/index.html")


@app.route("/api/status")
def get_status():
    logged_in = check_login()
    return jsonify({
        "logged_in": logged_in,
        "login_status": "已连接" if logged_in else "未连接，请先启动 NapCatQQ",
    })


@app.route("/api/groups")
def get_groups_api():
    if not check_login():
        return jsonify({"error": "未连接到 OneBot API"}), 401

    groups, error = get_groups()
    if error:
        return jsonify({"error": error}), 500

    return jsonify({"groups": groups, "total": len(groups)})


@app.route("/api/export/<int:group_id>", methods=["POST"])
def export_group_api(group_id):
    if not check_login():
        return jsonify({"error": "未连接到 OneBot API"}), 401

    groups, error = get_groups()
    if error:
        return jsonify({"error": error}), 500

    group_name = "群"
    for g in groups:
        if g.get("group_id") == group_id:
            group_name = g.get("group_name", "群")
            break

    members, error = get_group_members(group_id)
    if error:
        return jsonify({"error": error}), 500

    filename, error = export_to_excel(group_id, group_name, members)
    if error:
        return jsonify({"error": error}), 500

    return send_file(filename, as_attachment=True)


def main():
    print("=" * 66)
    print("                 QQ群成员导出工具")
    print("=" * 66)
    print()
    print("使用步骤：")
    print("1. 启动 NapCatQQ 并确保 QQ 已登录")
    print(f"2. 确保 HTTP API 服务运行在 {ONEBOT_URL}")
    print("3. 打开浏览器访问：http://127.0.0.1:8080")
    print()
    print("=" * 66)

    app.run(host="0.0.0.0", port=8080, debug=True)


if __name__ == "__main__":
    main()
