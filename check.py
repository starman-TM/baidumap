import pandas as pd
import re
import requests
import urllib.parse
import hashlib
import time
import os
import sys
import logging
import subprocess
from difflib import SequenceMatcher
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


# ==================== macOS 专属路径与工作目录处理 ====================
def get_target_work_dir():
    """获取正确的工作目录：
    如果是双击打包后的 .app 运行，返回 .app 所在的文件夹；
    如果是直接运行源码，返回源码所在文件夹。
    """
    if getattr(sys, "frozen", False):
        exec_path = os.path.abspath(sys.executable)
        # PyInstaller 打包的 macOS .app 中，可执行文件在 Contents/MacOS/ 内
        if ".app/Contents/MacOS/" in exec_path:
            app_path = exec_path.split(".app/Contents/MacOS/")[0] + ".app"
            return os.path.dirname(app_path)
        return os.path.dirname(exec_path)
    return os.path.dirname(os.path.abspath(__file__))


try:
    # 强制切换工作目录，保证相对路径的文件读写正确
    os.chdir(get_target_work_dir())
except OSError:
    pass

# ==================== 日志与弹窗配置 ====================
LOG_FILE = "run.log"
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    encoding="utf-8"
)


def alert(msg, title="网点核对"):
    """Mac 原生弹窗提示，解决无终端时看不到报错的问题"""
    try:
        safe_msg = str(msg).replace('"', '\\"').replace("'", "\\'")
        subprocess.run([
            "osascript", "-e",
            f'display dialog "{safe_msg}" with title "{title}" buttons {{"好"}} default button "好"'
        ], check=False, capture_output=True)
    except Exception as e:
        logging.error(f"弹窗失败: {e}")


# ==================== 配置项 ====================
# 建议：如果有安全要求，可以把 AK/SK 写入 config.json 并在代码中读取
AK = "WHS4VddJBLbPtyvt7g86c3VooHMHsQ6g"
SK = "swDJnphOonQDZdjVgfX8hHJCzV17I9OG"

# 固定在同一目录下读写
SOURCE_FILE = "原始网点信息2.xlsx"
OUTPUT_FILE = "百度地图核对结果-同福中路支行.xlsx"

# 源数据列名映射
COL_SOURCE_UID = "uid"
COL_SOURCE_NAME = "机构全称"
COL_SOURCE_ADDR = "地址"
COL_SOURCE_PHONE = "电话"
COL_SOURCE_HOURS_WEEKDAY = "营业时间（周一至周五）"
COL_SOURCE_HOURS_SAT = "营业时间（周六）"
COL_SOURCE_HOURS_SUN = "营业时间（周天）"

# 爬取结果列名
COL_CRAWL_NAME = "name"
COL_CRAWL_ADDR = "address"
COL_CRAWL_PHONE = "phone"
COL_CRAWL_HOURS_WEEKDAY = "weekday_hours"
COL_CRAWL_HOURS_SAT = "saturday_hours"
COL_CRAWL_HOURS_SUN = "sunday_hours"


# ==================== 营业时间解析 ====================
def _fmt_hm(h, m):
    try:
        return f"{int(h):02d}:{int(m):02d}"
    except (TypeError, ValueError):
        return ""


def parse_periods_to_day_hours(periods):
    """将 periods 解析为 {day: 'HH:MM-HH:MM'} (day: 0=周日, 1-5=周一至周五, 6=周六)"""
    day_hours = {}
    if not isinstance(periods, list):
        return day_hours
    for p in periods:
        if not isinstance(p, dict):
            continue
        try:
            o = p["open"]
            c = p["close"]
            day = int(o["day"])
            tr = f"{_fmt_hm(o['hour'], o['minute'])}-{_fmt_hm(c['hour'], c['minute'])}"
        except (KeyError, TypeError, ValueError):
            continue
        if day in day_hours:
            day_hours[day] += "," + tr
        else:
            day_hours[day] = tr
    return day_hours


def merge_days(day_hours, days):
    """合并多天的时间段为单一字符串用于展示"""
    if not day_hours:
        return ""
    vals = [day_hours.get(d, "休息") for d in days]
    if len(set(vals)) == 1:
        return vals[0]
    return " / ".join(vals)


# ==================== 爬取与核对 ====================
def get_requests_session():
    """带重试机制的请求 Session"""
    session = requests.Session()
    retry = Retry(connect=3, backoff_factor=0.5, status_forcelist=[500, 502, 503, 504])
    adapter = HTTPAdapter(max_retries=retry)
    session.mount('http://', adapter)
    session.mount('https://', adapter)
    return session


def fetch_poi_detail(session, uid, ak=AK, sk=SK):
    """根据 uid 获取 POI 详情"""
    params = {"uid": uid, "output": "json", "scope": "2", "ak": ak}
    params_arr = [f"{k}={params[k]}" for k in params]
    query_str = f"/place/v2/detail?{'&'.join(params_arr)}"
    encoded_str = urllib.parse.quote(query_str, safe="/:=&?#+!$,;'@()*[]")
    raw_str = encoded_str + sk
    sn = hashlib.md5(urllib.parse.quote_plus(raw_str).encode("utf8")).hexdigest()
    url = "https://api.map.baidu.com" + query_str + f"&sn={sn}"

    try:
        resp = session.get(url, timeout=10)
        if resp.status_code == 200:
            data = resp.json()
            if data.get("status") == 0:
                result = data.get("result", {}) or {}
                detail = result.get("detail_info", {}) or {}
                regular = detail.get("regular_open_hour", {}) or {}
                periods = regular.get("periods", []) if isinstance(regular, dict) else []

                day_hours = parse_periods_to_day_hours(periods)

                return {
                    "uid": uid,
                    "name": result.get("name", ""),
                    "address": result.get("address", ""),
                    "phone": detail.get("phone", result.get("telephone", "")),
                    "weekday_hours": merge_days(day_hours, [1, 2, 3, 4, 5]),
                    "saturday_hours": merge_days(day_hours, [6]),
                    "sunday_hours": merge_days(day_hours, [0]),
                }
            else:
                logging.warning(f"API返回错误 uid={uid}: {data.get('message')}")
    except Exception as e:
        logging.error(f"请求异常 uid={uid}: {e}")
    return None


def fetch_all_pois(source_df, delay=0.5):
    """根据 uid 批量爬取 POI 详情"""
    if "uid" not in source_df.columns:
        logging.error("源数据缺少 uid 列，无法爬取！")
        return pd.DataFrame()

    cache = {}
    records = []
    session = get_requests_session()

    total = len(source_df)
    for idx, (_, row) in enumerate(source_df.iterrows(), 1):
        uid = row["uid"]
        if pd.isna(uid): continue

        if uid in cache:
            detail = cache[uid]
        else:
            logging.info(f"正在爬取 ({idx}/{total}): {row.get(COL_SOURCE_NAME)} (uid={uid})")
            detail = fetch_poi_detail(session, uid)
            cache[uid] = detail
            time.sleep(delay)

        if detail:
            records.append(detail)
    return pd.DataFrame(records)


# ==================== 数据清洗与比较 ====================
def clean_address(addr):
    if not isinstance(addr, str): return addr
    addr = re.sub(r'^广东省', '', addr).strip()
    if addr.startswith('海珠区'): return '广州市' + addr
    match = re.search(r'(海珠区.*)', addr)
    if match: return '广州市' + match.group(1)
    if '广州' in addr: addr = re.sub(r'^广州', '广州市', addr)
    if not addr.startswith('广州市'): addr = '广州市' + addr
    return addr


def normalize_address(addr):
    if not isinstance(addr, str): return ""
    addr = addr.strip().replace("（", "(").replace("）", ")")
    addr = addr.replace("＃", "#").replace("～", "~")
    addr = re.sub(r'\([^)]*\)', '', addr)
    addr = re.sub(r'首层|二层|三层|四层|首、二层', '', addr)
    addr = re.sub(r'号铺|商铺|铺位|号(?!\d)', '', addr)
    return re.sub(r'\s+', ' ', addr).strip()


def extract_house_number(addr):
    if not isinstance(addr, str): return None
    numbers = re.findall(r'\d+', addr)
    return max(numbers, key=len) if numbers else None


def address_match(addr1, addr2, threshold=0.85):
    if not addr1 or not addr2: return False
    norm1, norm2 = normalize_address(addr1), normalize_address(addr2)
    num1, num2 = extract_house_number(norm1), extract_house_number(norm2)
    if num1 and num2 and num1 == num2: return True
    return SequenceMatcher(None, norm1, norm2).ratio() >= threshold


def compare_fields(val1, val2):
    def norm(s):
        if not isinstance(s, str): return ""
        s = s.strip().replace("（", "(").replace("）", ")").replace("，", ",").replace("：", ":")
        return re.sub(r'\s+', ' ', s)

    return norm(str(val1)) == norm(str(val2))


def normalize_hours(s):
    if not isinstance(s, str): return ""
    s = s.strip()
    if s in ("休息", "关闭", "闭店"): return "休息"
    s = (s.replace("~", "-").replace("～", "-").replace("—", "-")
         .replace("–", "-").replace("－", "-").replace("至", "-").replace("：", ":"))
    s = re.sub(r'(\d{1,2}):(\d{2})', lambda m: f"{int(m.group(1)):02d}:{m.group(2)}", s)
    return re.sub(r'\s+', '', s)


def compare_hours(src, crawl):
    s, c = normalize_hours(src), normalize_hours(crawl)
    if not s and not c: return False
    return s == c


def check_data(df_source, df_crawl, **kwargs):
    df_crawl = df_crawl.copy()
    if COL_CRAWL_ADDR in df_crawl.columns:
        df_crawl[COL_CRAWL_ADDR] = df_crawl[COL_CRAWL_ADDR].apply(clean_address)

    crawl_dict = {}
    for _, row in df_crawl.iterrows():
        key = row[COL_CRAWL_NAME].strip() if pd.notna(row[COL_CRAWL_NAME]) else ""
        if key and key not in crawl_dict:
            crawl_dict[key] = row

    result_rows = []
    for _, src_row in df_source.iterrows():
        src_name = src_row[kwargs['src_name_col']].strip() if pd.notna(src_row[kwargs['src_name_col']]) else ""
        src_addr = src_row[kwargs['src_addr_col']] if pd.notna(src_row[kwargs['src_addr_col']]) else ""
        src_phone = src_row[kwargs['src_phone_col']] if pd.notna(src_row[kwargs['src_phone_col']]) else ""
        src_wd = src_row[kwargs['src_hours_weekday_col']] if pd.notna(src_row[kwargs['src_hours_weekday_col']]) else ""
        src_sat = src_row[kwargs['src_hours_sat_col']] if pd.notna(src_row[kwargs['src_hours_sat_col']]) else ""
        src_sun = src_row[kwargs['src_hours_sun_col']] if pd.notna(src_row[kwargs['src_hours_sun_col']]) else ""

        match_status = "未找到"
        crawl_addr = crawl_phone = ""
        crawl_wd = crawl_sat = crawl_sun = ""
        addr_ok = phone_ok = wd_ok = sat_ok = sun_ok = False

        if src_name in crawl_dict:
            crawl_row = crawl_dict[src_name]
            crawl_addr = crawl_row.get(COL_CRAWL_ADDR, "")
            crawl_phone = crawl_row.get(COL_CRAWL_PHONE, "")
            crawl_wd = crawl_row.get(COL_CRAWL_HOURS_WEEKDAY, "")
            crawl_sat = crawl_row.get(COL_CRAWL_HOURS_SAT, "")
            crawl_sun = crawl_row.get(COL_CRAWL_HOURS_SUN, "")

            addr_ok = address_match(src_addr, crawl_addr)
            phone_ok = compare_fields(src_phone, crawl_phone)
            wd_ok = compare_hours(src_wd, crawl_wd)
            sat_ok = compare_hours(src_sat, crawl_sat)
            sun_ok = compare_hours(src_sun, crawl_sun)

            if addr_ok and phone_ok and wd_ok and sat_ok and sun_ok:
                match_status = "完全一致"
            else:
                match_status = "有差异"

        def mark(ok):
            return "" if match_status == "未找到" else ("✅" if ok else "❌")

        result_rows.append({
            "机构全称": src_name,
            "已有地址": src_addr,
            "爬取地址(清洗后)": crawl_addr,
            "地址一致": mark(addr_ok),
            "已有电话": src_phone,
            "爬取电话": crawl_phone,
            "电话一致": mark(phone_ok),
            "已有营业时间(周一至周五)": src_wd,
            "爬取营业时间(周一至周五)": crawl_wd,
            "营业时间(周一至周五)一致": mark(wd_ok),
            "已有营业时间(周六)": src_sat,
            "爬取营业时间(周六)": crawl_sat,
            "营业时间(周六)一致": mark(sat_ok),
            "已有营业时间(周天)": src_sun,
            "爬取营业时间(周天)": crawl_sun,
            "营业时间(周天)一致": mark(sun_ok),
            "核对状态": match_status,
        })
    return pd.DataFrame(result_rows)


# ==================== 主流程 ====================
def _main_impl():
    if not os.path.exists(SOURCE_FILE):
        raise FileNotFoundError(f"未找到输入文件: {SOURCE_FILE}\n请确保它和程序在同一个文件夹内。")

    logging.info(f"开始读取源数据: {SOURCE_FILE}")
    df_source = pd.read_excel(SOURCE_FILE)
    logging.info(f"已读取源数据，共 {len(df_source)} 条记录")

    df_crawl = fetch_all_pois(df_source)
    if df_crawl.empty:
        raise ValueError("爬取失败，未获取到任何数据。请检查网络或 AK/SK 是否正确。")
    logging.info(f"爬取完成，共 {len(df_crawl)} 条记录")

    df_result = check_data(
        df_source, df_crawl,
        src_name_col=COL_SOURCE_NAME, src_addr_col=COL_SOURCE_ADDR, src_phone_col=COL_SOURCE_PHONE,
        src_hours_weekday_col=COL_SOURCE_HOURS_WEEKDAY, src_hours_sat_col=COL_SOURCE_HOURS_SAT,
        src_hours_sun_col=COL_SOURCE_HOURS_SUN
    )
    logging.info(f"核对完成，共 {len(df_result)} 条记录")

    logging.info(f"正在保存结果至: {OUTPUT_FILE}")
    try:
        with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
            df_crawl.to_excel(writer, sheet_name="原始爬取数据", index=False)
            df_result.to_excel(writer, sheet_name="核对结果", index=False)
    except PermissionError:
        raise PermissionError(f"无法保存文件。请确保没有其他程序（如 Excel）正在打开【{OUTPUT_FILE}】。")

    logging.info("全部流程执行完毕。")


def main():
    logging.info("========== 程序启动 ==========")
    try:
        _main_impl()
        alert(f"核对成功！\n结果已保存至:\n{OUTPUT_FILE}")
    except Exception as e:
        error_msg = str(e)
        logging.error(f"运行出错: {error_msg}")
        logging.error(traceback.format_exc())
        alert(f"运行失败：\n{error_msg}\n\n详细错误日志请查看同目录下的 {LOG_FILE}")
    finally:
        logging.info("========== 程序退出 ==========")


if __name__ == "__main__":
    import traceback

    main()