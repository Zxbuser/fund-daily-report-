import datetime as dt
import html
import json
import os
import re
import smtplib
import subprocess
import sys
import traceback
import winreg
from email.message import EmailMessage
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


ROOT = Path(__file__).resolve().parent
REPORT_ROOT = ROOT / "reports"
LOG_DIR = REPORT_ROOT / "logs"
FUND_ONE = ["027052", "021528", "021485", "022365", "026376", "026733", "968044", "008971", "021143", "000218", "040046"]
FUND_TWO = ["021528", "022365", "014915", "025209", "011452", "024975", "005359", "011892", "018957"]
RECIPIENTS = {
    "基金一": "1569227264@qq.com",
    "基金二": "zxb991213@qq.com",
}
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587


def beijing_now():
    return dt.datetime.now(dt.timezone.utc).replace(tzinfo=None) + dt.timedelta(hours=8)


RUN_DATE = beijing_now().date()
ANALYSIS_DATE = RUN_DATE - dt.timedelta(days=1)
RUN_DIR = REPORT_ROOT / RUN_DATE.isoformat()
LOG_FILE = LOG_DIR / f"fund_daily_report_{RUN_DATE.isoformat()}.log"


def log(message):
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    line = f"[{beijing_now().strftime('%Y-%m-%d %H:%M:%S')}] {message}"
    print(line)
    with LOG_FILE.open("a", encoding="utf-8") as fh:
        fh.write(line + "\n")


def user_env(name):
    value = os.environ.get(name)
    if value:
        return value
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment") as key:
            value, _ = winreg.QueryValueEx(key, name)
            return value
    except OSError:
        return ""


def curl_text(url):
    raw = subprocess.check_output(
        ["curl.exe", "-s", "-L", "--max-time", "30", url],
        stderr=subprocess.DEVNULL,
    )
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        return raw.decode("gb18030", "replace")


def parse_var(text, name):
    match = re.search(rf'var\s+{re.escape(name)}\s*=\s*"(.*?)";', text)
    return match.group(1) if match else ""


def parse_assignment_json(text, var_name):
    match = re.search(rf"var\s+{re.escape(var_name)}\s*=\s*(.*?);/\*", text, re.S)
    if not match:
        match = re.search(rf"var\s+{re.escape(var_name)}\s*=\s*(.*?);", text, re.S)
    if not match:
        return None
    try:
        return json.loads(match.group(1))
    except Exception:
        return None


def fund_latest_networth(text):
    data = parse_assignment_json(text, "Data_netWorthTrend")
    if not data:
        return None
    latest = None
    for item in data:
        day = (dt.datetime.fromtimestamp(item["x"] / 1000, dt.timezone.utc).replace(tzinfo=None) + dt.timedelta(hours=8)).date()
        if day <= ANALYSIS_DATE:
            latest = {
                "date": day.isoformat(),
                "nav": item.get("y"),
                "day_return": item.get("equityReturn"),
            }
    return latest


def manager_names(text):
    data = parse_assignment_json(text, "Data_currentFundManager")
    if not data:
        return "--"
    names = [item.get("name") for item in data if item.get("name")]
    return "、".join(names) if names else "--"


def latest_scale(text):
    data = parse_assignment_json(text, "Data_fluctuationScale")
    if not data:
        return "--"
    categories = data.get("categories") or []
    series = data.get("series") or []
    if not categories or not series:
        return "--"
    item = series[-1]
    value = item.get("y")
    mom = item.get("mom")
    if value is None:
        return "--"
    return f"{categories[-1]}：{value}亿" + (f"，环比{mom}" if mom else "")


def latest_allocation(text):
    data = parse_assignment_json(text, "Data_assetAllocation")
    if not data:
        return "--"
    parts = []
    for item in data.get("series", []):
        values = item.get("data") or []
        if values:
            parts.append(f"{item.get('name', '')}{values[-1]}%")
    return "；".join(parts[:4]) if parts else "--"


def fetch_fund(code):
    url = f"https://fund.eastmoney.com/pingzhongdata/{code}.js?v={int(dt.datetime.now(dt.timezone.utc).timestamp())}"
    try:
        text = curl_text(url)
        name = parse_var(text, "fS_name")
        if not name:
            raise ValueError("fund name not found")
        latest = fund_latest_networth(text)
        return {
            "code": code,
            "name": name,
            "latest": latest,
            "m1": parse_var(text, "syl_1y") or "--",
            "m3": parse_var(text, "syl_3y") or "--",
            "m6": parse_var(text, "syl_6y") or "--",
            "y1": parse_var(text, "syl_1n") or "--",
            "manager": manager_names(text),
            "scale": latest_scale(text),
            "allocation": latest_allocation(text),
            "source": f"https://fund.eastmoney.com/{code}.html",
            "error": "",
        }
    except Exception as exc:
        return {
            "code": code,
            "name": "未检索到",
            "latest": None,
            "m1": "--",
            "m3": "--",
            "m6": "--",
            "y1": "--",
            "manager": "--",
            "scale": "--",
            "allocation": "--",
            "source": f"https://fund.eastmoney.com/{code}.html",
            "error": str(exc),
        }


def fetch_index(name, secid):
    end = ANALYSIS_DATE.strftime("%Y%m%d")
    beg = (ANALYSIS_DATE - dt.timedelta(days=10)).strftime("%Y%m%d")
    url = (
        "https://push2his.eastmoney.com/api/qt/stock/kline/get?"
        f"secid={secid}&fields1=f1,f2,f3,f4,f5,f6&"
        "fields2=f51,f52,f53,f54,f55,f56,f57,f58,f59,f60,f61&"
        f"klt=101&fqt=1&beg={beg}&end={end}"
    )
    try:
        data = json.loads(curl_text(url)).get("data") or {}
        klines = data.get("klines") or []
        if not klines:
            raise ValueError("no index data")
        parts = klines[-1].split(",")
        return {
            "name": name,
            "date": parts[0],
            "open": parts[1],
            "close": parts[2],
            "high": parts[3],
            "low": parts[4],
            "amount": parts[6],
            "amp": parts[7],
            "pct": parts[8],
            "chg": parts[9],
            "source": "东方财富行情接口",
        }
    except Exception as exc:
        return {"name": name, "date": "--", "close": "--", "pct": "--", "chg": "--", "low": "--", "high": "--", "source": str(exc)}


def to_float(value):
    try:
        return float(value)
    except Exception:
        return None


def pct_text(value):
    number = to_float(value)
    if number is None:
        return "--"
    return f"{number:+.2f}%"


def group_stats(codes, funds):
    values = []
    for code in codes:
        latest = funds[code].get("latest") or {}
        number = to_float(latest.get("day_return"))
        if number is not None:
            values.append(number)
    if not values:
        return {"count": 0, "avg": "--", "up": 0, "down": 0}
    return {
        "count": len(values),
        "avg": f"{sum(values) / len(values):+.2f}%",
        "up": sum(1 for value in values if value > 0),
        "down": sum(1 for value in values if value < 0),
    }


def analyze_fund(fund):
    """基于多周期收益数据，对单只基金进行趋势、动能和仓位建议分析"""
    m1 = to_float(fund["m1"])
    m3 = to_float(fund["m3"])
    m6 = to_float(fund["m6"])
    y1 = to_float(fund["y1"])
    latest = fund.get("latest") or {}
    day_ret = to_float(latest.get("day_return"))

    periods = []
    if m1 is not None:
        periods.append(("近1月", m1))
    if m3 is not None:
        periods.append(("近3月", m3))
    if m6 is not None:
        periods.append(("近6月", m6))
    if y1 is not None:
        periods.append(("近1年", y1))

    if len(periods) < 1:
        return {
            "trend": "数据不足",
            "momentum": "数据不足",
            "suggestion": "⚪ 暂不评估",
            "reason": "有效收益数据不足，无法生成分析",
            "daily_signal": "无当日数据" if day_ret is None else f"日涨跌 {day_ret:+.2f}%",
            "risk": "--",
        }

    pos_count = sum(1 for _, v in periods if v > 0)
    neg_count = sum(1 for _, v in periods if v < 0)
    all_pos = pos_count == len(periods)
    all_neg = neg_count == len(periods)

    # ---- 趋势判断 ----
    if all_pos:
        trend = "全面上行 ✅"
    elif all_neg:
        trend = "全面下行 ❌"
    elif pos_count > neg_count:
        trend = "偏强震荡 ↗"
    elif neg_count > pos_count:
        trend = "偏弱震荡 ↘"
    else:
        trend = "多空均衡 ➡"

    # ---- 动能分析 ----
    if m1 is not None and m3 is not None and m6 is not None:
        if m1 > 0 and m3 > 0 and m1 > m3 > m6:
            momentum = "加速上攻 🚀"
        elif m1 > 0 and m3 > 0 and m1 < m3:
            momentum = "涨势放缓 ⏸"
        elif m1 > 0 and m3 < 0:
            momentum = "短期反弹 🔄"
        elif m1 < 0 and m3 > 0:
            momentum = "短期回调 📉"
        elif m1 < 0 and m3 < 0 and m1 < m3:
            momentum = "加速下跌 ⚠"
        else:
            momentum = "趋势延续 ➡"
    elif m1 is not None and m3 is not None:
        if m1 > 0 and m3 > 0:
            momentum = "短中期向好"
        elif m1 < 0 and m3 < 0:
            momentum = "短中期偏弱"
        elif m1 > 0 > m3:
            momentum = "短期反弹中"
        else:
            momentum = "短期回调中"
    else:
        momentum = "数据有限"

    # ---- 单日信号 ----
    if day_ret is not None:
        if day_ret > 3:
            daily_signal = f"单日大涨 {day_ret:+.2f}%⚠ 追高风险"
        elif day_ret > 1.5:
            daily_signal = f"单日活跃 {day_ret:+.2f}%↑"
        elif day_ret < -3:
            daily_signal = f"单日大跌 {day_ret:+.2f}%‼ 注意风险"
        elif day_ret < -1.5:
            daily_signal = f"单日承压 {day_ret:+.2f}%↓"
        else:
            daily_signal = f"波动正常 {day_ret:+.2f}%"
    else:
        daily_signal = "净值未更新"

    # ---- 风险等级 ----
    if day_ret is not None and abs(day_ret) > 3:
        risk = "高"
    elif all_neg:
        risk = "高"
    elif m1 is not None and m1 < -5:
        risk = "较高"
    elif all_pos:
        risk = "较低"
    else:
        risk = "中等"

    # ---- 仓位建议 ----
    if fund.get("error"):
        suggestion = "⚪ 数据缺失"
        reason = "基金数据抓取失败，建议手动核实"
    elif all_pos:
        if m1 is not None and m3 is not None and m1 > m3 * 1.5 and m1 > 8:
            suggestion = "🟠 分批止盈"
            reason = f"近1月涨幅{m1:+.1f}%远超近3月{m3:+.1f}%，短期过热，建议分批锁定利润"
        else:
            suggestion = "🟢 继续持有"
            reason = "各周期趋势一致向好，维持仓位，无需频繁操作"
    elif pos_count > neg_count and m1 is not None and m1 < 0:
        suggestion = "🟡 逢低关注"
        reason = f"中长期偏强但近1月回调{m1:+.1f}%，等待企稳后可能是加仓窗口"
    elif pos_count > neg_count:
        suggestion = "🟢 持有观望"
        reason = "整体偏强，短期表现尚可，按纪律持有"
    elif all_neg:
        if m1 is not None and m1 < -8:
            suggestion = "🔴 建议减仓"
            reason = f"近1月跌幅{m1:+.1f}%较大，各周期全面下行，建议降低仓位控制风险"
        else:
            suggestion = "🟡 观望等待"
            reason = "各周期偏弱但跌幅尚可控，观察是否出现企稳信号"
    elif m1 is not None and m1 > 0 and m3 is not None and m3 < 0:
        suggestion = "🟡 轻仓试探"
        reason = "短期反弹但中期仍弱，可小仓试探但不宜重仓"
    else:
        suggestion = "🟡 观望为主"
        reason = "信号不明确，建议等待方向确认后再做决策"

    return {
        "trend": trend,
        "momentum": momentum,
        "suggestion": suggestion,
        "reason": reason,
        "daily_signal": daily_signal,
        "risk": risk,
    }


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), fill)
    tc_pr.append(shading)


def set_cell_text(cell, text, bold=False):
    cell.text = str(text)
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.name = "Microsoft YaHei"
            run._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
            run.font.size = Pt(8.5)
            run.bold = bold


def style_table(table):
    table.style = "Table Grid"
    for row_index, row in enumerate(table.rows):
        for cell in row.cells:
            if row_index == 0:
                set_cell_shading(cell, "D9EAF7")
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.space_after = Pt(0)


def add_table(document, headers, rows):
    table = document.add_table(rows=1, cols=len(headers))
    for idx, header in enumerate(headers):
        set_cell_text(table.rows[0].cells[idx], header, bold=True)
    for row in rows:
        cells = table.add_row().cells
        for idx, value in enumerate(row):
            set_cell_text(cells[idx], value)
    style_table(table)
    document.add_paragraph()
    return table


def setup_document(title, subtitle):
    document = Document()
    section = document.sections[0]
    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.55)
    section.right_margin = Inches(0.55)
    styles = document.styles
    normal = styles["Normal"]
    normal.font.name = "Microsoft YaHei"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    normal.font.size = Pt(9.5)
    title_paragraph = document.add_paragraph()
    title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_paragraph.add_run(title)
    run.bold = True
    run.font.size = Pt(18)
    run.font.color.rgb = RGBColor(31, 78, 121)
    subtitle_paragraph = document.add_paragraph()
    subtitle_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle_paragraph.add_run(subtitle)
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(102, 112, 133)
    return document


def fund_rows(codes, funds):
    rows = []
    for code in codes:
        fund = funds[code]
        latest = fund.get("latest") or {}
        rows.append([
            code,
            fund["name"],
            latest.get("date", "--"),
            latest.get("nav", "--"),
            pct_text(latest.get("day_return")),
            pct_text(fund["m1"]),
            pct_text(fund["m3"]),
            pct_text(fund["m6"]),
            pct_text(fund["y1"]),
            fund["manager"],
            fund["scale"],
            fund["allocation"],
        ])
    return rows


def make_docx(label, codes, funds, indices, output_path):
    stats = group_stats(codes, funds)
    doc = setup_document(
        f"{label}每日复盘与仓位建议",
        f"运行日期：{RUN_DATE.isoformat()}    分析日期：{ANALYSIS_DATE.isoformat()}    数据源：东方财富/天天基金公开数据",
    )

    doc.add_heading("一、摘要", level=1)
    add_table(doc, ["组合", "有数据基金数", "平均日涨跌", "上涨只数", "下跌只数"], [[label, stats["count"], stats["avg"], stats["up"], stats["down"]]])
    doc.add_paragraph("结论：昨日市场偏向成长和高弹性资产。若组合已明显盈利，不建议因单日上涨追高；更适合检查集中度、重复暴露和单只基金权重。")

    doc.add_heading("二、市场环境", level=1)
    add_table(
        doc,
        ["指数", "日期", "收盘", "涨跌幅", "涨跌额", "日内区间"],
        [[item["name"], item["date"], item["close"], pct_text(item["pct"]), item["chg"], f"{item['low']} - {item['high']}"] for item in indices],
    )

    doc.add_heading("三、基金明细", level=1)
    add_table(
        doc,
        ["代码", "基金名称", "净值日期", "单位净值", "日涨跌", "近1月", "近3月", "近6月", "近1年", "基金经理", "规模", "资产配置"],
        fund_rows(codes, funds),
    )

    doc.add_heading("四、逐只基金分析与仓位建议", level=1)
    analysis_headers = ["代码", "名称", "日涨跌", "趋势判断", "动能状态", "风险等级", "操作建议", "建议理由"]
    analysis_rows = []
    for code in codes:
        fund = funds[code]
        a = analyze_fund(fund)
        latest = fund.get("latest") or {}
        analysis_rows.append([
            code,
            fund["name"],
            pct_text(latest.get("day_return")),
            a["trend"],
            a["momentum"],
            a["risk"],
            a["suggestion"],
            a["reason"],
        ])
    add_table(doc, analysis_headers, analysis_rows)

    doc.add_heading("五、组合整体评估", level=1)
    # 计算组合层面统计
    risk_high = sum(1 for r in analysis_rows if "高" in str(r[5]) and "较高" not in str(r[5]))
    risk_elevated = sum(1 for r in analysis_rows if "较高" in str(r[5]))
    hold_green = sum(1 for r in analysis_rows if "继续持有" in str(r[6]) or "持有观望" in str(r[6]))
    watch_yellow = sum(1 for r in analysis_rows if "逢低关注" in str(r[6]) or "观望" in str(r[6]) or "轻仓" in str(r[6]))
    reduce_red = sum(1 for r in analysis_rows if "减仓" in str(r[6]) or "止盈" in str(r[6]))

    combo_lines = [
        f"· 组合整体表现：{stats['avg']}，{stats['up']}只上涨 / {stats['down']}只下跌",
        f"· 建议持有/观望：{hold_green}只 | 关注/等待：{watch_yellow}只 | 减仓/止盈关注：{reduce_red}只",
        f"· 高风险基金：{risk_high}只 | 较高风险：{risk_elevated}只",
    ]
    for line in combo_lines:
        doc.add_paragraph(line)

    doc.add_heading("六、跨基金配置提示", level=1)
    doc.add_paragraph("1. 成长/科技相关基金在昨日市场环境中更容易受益，但短期涨幅越集中，后续波动和回撤风险也越高。")
    doc.add_paragraph("2. 两个组合中存在重复代码 021528 和 022365，需要在总账户层面合并计算暴露，避免实际权重高于预期。")
    doc.add_paragraph("3. QDII、港股、黄金等资产净值可能滞后一个交易日，判断组合表现时应区分披露日期。")

    doc.add_heading("七、风险提示与数据来源", level=1)
    doc.add_paragraph("本报告基于公开数据自动生成，不构成个性化投资顾问意见，不承诺收益。最终操作应结合风险承受能力、资金期限和实际持仓成本。")
    add_table(doc, ["基金代码", "来源链接", "异常"], [[code, funds[code]["source"], funds[code]["error"] or "--"] for code in codes])
    doc.save(output_path)


def html_pct(value):
    number = to_float(value)
    if number is None:
        return "--"
    color = "#b42318" if number > 0 else ("#067647" if number < 0 else "#475467")
    return f'<span style="color:{color};font-weight:600;">{number:+.2f}%</span>'


def html_table(headers, rows):
    header_html = "".join(f"<th>{html.escape(str(header))}</th>" for header in headers)
    row_html = []
    for row in rows:
        row_html.append("<tr>" + "".join(f"<td>{cell}</td>" for cell in row) + "</tr>")
    return f"<table><thead><tr>{header_html}</tr></thead><tbody>{''.join(row_html)}</tbody></table>"


def email_body(label, codes, funds, indices):
    stats = group_stats(codes, funds)
    index_rows = [
        [
            html.escape(item["name"]),
            html.escape(item["date"]),
            html.escape(str(item["close"])),
            html_pct(item["pct"]),
            html.escape(str(item["chg"])),
            html.escape(f"{item['low']} - {item['high']}"),
        ]
        for item in indices
    ]
    fund_rows_html = []
    for code in codes:
        fund = funds[code]
        latest = fund.get("latest") or {}
        fund_rows_html.append([
            html.escape(code),
            html.escape(fund["name"]),
            html.escape(str(latest.get("date", "--"))),
            html.escape(str(latest.get("nav", "--"))),
            html_pct(latest.get("day_return")),
            html_pct(fund["m1"]),
            html_pct(fund["m3"]),
            html_pct(fund["m6"]),
            html_pct(fund["y1"]),
        ])

    # ---- 逐只分析与仓位建议 ----
    analysis_rows_html = []
    for code in codes:
        fund = funds[code]
        a = analyze_fund(fund)
        latest = fund.get("latest") or {}
        suggestion_color = "#b42318" if "减仓" in a["suggestion"] or "止盈" in a["suggestion"] else ("#e67e22" if "观望" in a["suggestion"] or "关注" in a["suggestion"] or "轻仓" in a["suggestion"] else ("#067647" if "持有" in a["suggestion"] else "#475467"))
        analysis_rows_html.append([
            html.escape(code),
            html.escape(fund["name"]),
            html_pct(latest.get("day_return")),
            html.escape(a["trend"]),
            html.escape(a["momentum"]),
            html.escape(str(a["risk"])),
            f'<span style="color:{suggestion_color};font-weight:700;">{html.escape(a["suggestion"])}</span>',
            html.escape(a["reason"]),
        ])

    # ---- 组合统计 ----
    risk_high = sum(1 for r in analysis_rows_html if "高" in str(r[5]) and "较高" not in str(r[5]))
    hold_green = sum(1 for r in analysis_rows_html if "继续持有" in str(r[6]) or "持有观望" in str(r[6]))
    watch_yellow = sum(1 for r in analysis_rows_html if "逢低关注" in str(r[6]) or "观望" in str(r[6]) or "轻仓" in str(r[6]))
    reduce_red = sum(1 for r in analysis_rows_html if "减仓" in str(r[6]) or "止盈" in str(r[6]))

    css = """
    body { font-family: Arial, 'Microsoft YaHei', sans-serif; color:#101828; line-height:1.55; }
    h1 { font-size:20px; margin:0 0 6px; }
    h2 { font-size:16px; margin:22px 0 8px; border-left:4px solid #2e90fa; padding-left:8px; }
    .meta,.note { color:#667085; font-size:13px; }
    table { border-collapse:collapse; width:100%; font-size:13px; margin:8px 0 14px; }
    th { background:#f2f4f7; border:1px solid #d0d5dd; padding:7px 8px; text-align:left; white-space:nowrap; }
    td { border:1px solid #d0d5dd; padding:7px 8px; vertical-align:top; }
    .combo { background:#f9fafb; border-left:4px solid #2e90fa; padding:10px 14px; margin:12px 0; font-size:13px; line-height:1.8; }
    """
    return f"""<!doctype html><html><head><meta charset="utf-8"><style>{css}</style></head><body>
    <h1>{label}每日复盘与仓位建议</h1>
    <div class="meta">运行日期：{RUN_DATE.isoformat()} ｜ 分析日期：{ANALYSIS_DATE.isoformat()} ｜ 数据源：东方财富/天天基金公开数据</div>
    <h2>组合表现</h2>
    {html_table(["组合", "有数据基金数", "平均日涨跌", "上涨只数", "下跌只数"], [[label, stats["count"], stats["avg"], stats["up"], stats["down"]]])}
    <div class="combo">📊 持有/观望：{hold_green}只 | ⚠ 关注/等待：{watch_yellow}只 | 🔴 减仓/止盈关注：{reduce_red}只 | ⚡ 高风险：{risk_high}只</div>
    <h2>市场概况</h2>
    {html_table(["指数", "日期", "收盘", "涨跌幅", "涨跌额", "日内区间"], index_rows)}
    <h2>基金明细</h2>
    {html_table(["代码", "基金名称", "净值日期", "单位净值", "日涨跌", "近1月", "近3月", "近6月", "近1年"], fund_rows_html)}
    <h2>📋 逐只分析与仓位建议</h2>
    {html_table(["代码", "名称", "日涨跌", "趋势判断", "动能状态", "风险", "操作建议", "建议理由"], analysis_rows_html)}
    <p class="note">本邮件为公开数据复盘和仓位管理框架，不构成个性化投资建议或收益承诺。完整表格见附件 Word 文档。</p>
    </body></html>"""


def send_email(label, recipient, docx_path, codes, funds, indices):
    user = user_env("FUND_REPORT_SMTP_USER") or "zxb991213@gmail.com"
    password = user_env("FUND_REPORT_SMTP_PASS")
    if not password:
        raise RuntimeError("FUND_REPORT_SMTP_PASS is missing")
    msg = EmailMessage()
    msg["From"] = user
    msg["To"] = recipient
    msg["Subject"] = f"{label}每日复盘与仓位建议 - {ANALYSIS_DATE.isoformat()}"
    msg.set_content(f"{label}每日复盘与仓位建议已生成。请查看附件 Word 文档。")
    msg.add_alternative(email_body(label, codes, funds, indices), subtype="html")
    data = docx_path.read_bytes()
    msg.add_attachment(
        data,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=docx_path.name,
    )
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=45) as smtp:
        smtp.ehlo()
        smtp.starttls()
        smtp.ehlo()
        smtp.login(user, password)
        smtp.send_message(msg)


def run():
    RUN_DIR.mkdir(parents=True, exist_ok=True)
    log(f"starting report run; run_date={RUN_DATE.isoformat()} analysis_date={ANALYSIS_DATE.isoformat()}")
    indices = [
        fetch_index("上证指数", "1.000001"),
        fetch_index("深证成指", "0.399001"),
        fetch_index("创业板指", "0.399006"),
        fetch_index("沪深300", "1.000300"),
        fetch_index("恒生指数", "100.HSI"),
    ]
    funds = {}
    for code in sorted(set(FUND_ONE + FUND_TWO)):
        funds[code] = fetch_fund(code)
        log(f"fetched fund {code}: {funds[code]['name']} error={funds[code]['error'] or '-'}")

    doc1 = RUN_DIR / f"基金一_每日复盘_{ANALYSIS_DATE.isoformat()}.docx"
    doc2 = RUN_DIR / f"基金二_每日复盘_{ANALYSIS_DATE.isoformat()}.docx"
    make_docx("基金一", FUND_ONE, funds, indices, doc1)
    make_docx("基金二", FUND_TWO, funds, indices, doc2)
    log(f"created docx: {doc1}")
    log(f"created docx: {doc2}")

    send_email("基金一", RECIPIENTS["基金一"], doc1, FUND_ONE, funds, indices)
    log(f"sent email for 基金一 to {RECIPIENTS['基金一']}")
    send_email("基金二", RECIPIENTS["基金二"], doc2, FUND_TWO, funds, indices)
    log(f"sent email for 基金二 to {RECIPIENTS['基金二']}")
    log("report run completed")


if __name__ == "__main__":
    try:
        run()
    except Exception:
        log("FAILED")
        log(traceback.format_exc())
        sys.exit(1)
