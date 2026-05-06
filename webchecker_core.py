import asyncio
from typing import List, Optional, Set, Tuple
import sys

# 供除錯／版本確認：與 Streamlit 側欄「檢核核心版本」應一致
SCAN_ENGINE_BUILD = "2026-05-06-p24"
import random
import requests
import urllib3
from bs4 import BeautifulSoup
import re
import pandas as pd
import json
import os
import uuid
import urllib.parse
import time
from io import BytesIO
from urllib.parse import urljoin, urlparse, urlunparse, quote, unquote
from html import unescape as html_unescape
import html as html_stdlib
from datetime import datetime, timedelta
from pathlib import PurePosixPath
from playwright.async_api import async_playwright

# recycle.eri.com.tw utmap 店家詳情：首屏 HTML 常僅殼層，「官方網址」等欄位晚數百 ms 才由 API 注入；併發掃描時若略早取
# `page.content()`，會漏掉外站 href，Excel 即無該外站列（但站內列「符合」仍在）。
_UTMAP_STATION_DETAIL_PATH_RE=re.compile(
    r"/utmap/stations/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\b",
    re.I,
)

# 與側邊選單頁面切換一致（勿與按鈕文案脫鉤）
APP_MODE_SCAN="🔍執行全站掃描"
APP_MODE_FAV="⭐常用網站設定"
APP_MODE_EXCLUDE="🛡️排除規則設定"
APP_MODE_INDICATOR_HELP="📖檢核指標說明"

# ==========================================
# 關鍵修正：解決Windows環境與事件迴圈問題
# ==========================================
if sys.platform=='win32':
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

CONFIG_FILE="config.json"
FAVORITES_FILE="favorites.json"
# 常用網站 selectbox：選項列用 (id, name, url)，顯示「名稱｜網址」
_SCAN_FAV_NONE=("__none__","","（請選擇）")

def _format_fav_select_row(row):
    if row[0]=="__none__":return row[2]
    n,u=row[1]or"",row[2]or""
    return f"{n} ｜ {u}" if n else u

def ordered_visited_urls_for_export(visited_urls:Set[str],visited_order:Optional[List[str]]):
    vs,vo=visited_urls,visited_order or[]
    seen=set()
    out=[]
    for u in vo:
        if u in vs and u not in seen:
            seen.add(u);out.append(u)
    for u in sorted(vs):
        if u not in seen:
            seen.add(u);out.append(u)
    return out

def _fav_select_rows(favs):
    return[_SCAN_FAV_NONE]+[(x.get("id",""),x.get("name",""),x.get("url","")) for x in favs]

# Excel 實務上過長之 file://／外部 URL 嵌入 hyperlink 仍可能觸發舊版相容警告；過長者保留為可換行之純文字。
_EXCEL_HYPERLINK_URL_MAX_LEN = 2048

def _excel_cell_url_hyperlink_writable(u:str)->bool:
    u=(u or"").strip()
    if not u or u=="—":
        return False
    if len(u)>_EXCEL_HYPERLINK_URL_MAX_LEN:
        return False
    lo=u.lower()
    return lo.startswith(("http://","https://"))


def visited_urls_to_excel_bytes(
    url_list, external_probed_list=None, link_invalid_urls:Optional[Set[str]]=None,
    external_source_map:Optional[dict]=None,
):
    """產生含站內掃描與外站之網址清單。

    external_probed_list：彙整後**全部外站 URL**；每一筆皆已做 HEAD/GET 連線驗證（見 probe_external_links_unreachable），
    「類型」欄為外站（已連線驗證）。

    link_invalid_urls：列入「5.有效連結」不符合清單的網址集合（站內載入失敗或外站探測失敗）；用於「有效連結」欄。
    若為 None，該欄一律填「—」（相容舊呼叫端）。

    external_source_map：{外站 URL: [出現該連結之站內頁面 URL, ...]}；用於「來源站內頁面」欄。
    站內列固定填「—」；外站列以換行串接所有來源頁，配合 wrap_text 顯示多行。

    輸出之 xlsx：第一列標題啟用自動篩選；**「有效連結」為「不符合」之列排在清單最上方**（其餘列
    維持原本站內→外站之相對順序）；序號於排序後重新連續編號。

    「掃描網址」及單筆之「來源站內頁面」若為長度在
    _EXCEL_HYPERLINK_URL_MAX_LEN 內之 http(s) URL，儲存為藍字底線可點擊超連結；更長者仍為
    純文字以降低 Excel 開檔相容性問題。
    """
    v=url_list or []
    e=[]if external_probed_list is None else list(external_probed_list)
    bad=link_invalid_urls if link_invalid_urls is not None else None
    src=external_source_map if isinstance(external_source_map,dict)else{}
    rows=[]
    for i,u in enumerate(v,1):
        lc="—"if bad is None else("不符合"if u in bad else"符合")
        rows.append({"序號":i,"掃描網址":u,"類型":"站內掃描","有效連結":lc,"來源站內頁面":"—"})
    off=len(rows)
    for idx,u in enumerate(e):
        lc="—"if bad is None else("不符合"if u in bad else"符合")
        srcs=src.get(u)or[]
        # 多列來源以換行串接；空集合（極少見：經其他途徑加入而無紀錄）顯示「—」
        srcs_text="\n".join(srcs)if srcs else"—"
        rows.append({"序號":off+idx+1,"掃描網址":u,"類型":"外站（已連線驗證）","有效連結":lc,"來源站內頁面":srcs_text})
    # 「不符合」列置頂，其餘維持既有相對順序（stable sort）；序號重新 1..N
    def _excel_lc_sort_key(r:dict)->tuple:
        lc=(r.get("有效連結")or"").strip()
        if lc=="不符合":
            return(0,)
        if lc=="符合":
            return(1,)
        return(2,)
    rows=sorted(rows,key=_excel_lc_sort_key)
    for i,r in enumerate(rows,1):
        r["序號"]=i
    buf=BytesIO()
    cols=["序號","掃描網址","類型","有效連結","來源站內頁面"]
    df=pd.DataFrame(rows,columns=cols)if rows else pd.DataFrame({c:[]for c in cols})
    # 關閉 xlsxwriter 之 strings_to_urls 自動轉超連結：改於寫入後對「長度安全」之 URL 明確 write_url，
    # 避免批量自動轉換時混入超長網址導致 Excel 開檔警告（見 _EXCEL_HYPERLINK_URL_MAX_LEN）。
    engine_kw={"options":{"strings_to_urls":False}}
    with pd.ExcelWriter(buf,engine="xlsxwriter",engine_kwargs=engine_kw) as writer:
        df.to_excel(writer,index=False,sheet_name="掃描清單")
        wb=writer.book;ws=writer.sheets["掃描清單"]
        wrap_fmt=wb.add_format({"text_wrap":True,"valign":"top"})
        plain_fmt=wb.add_format({"valign":"top"})
        link_wrap_fmt=wb.add_format({"text_wrap":True,"valign":"top","font_color":"blue","underline":True})
        n=len(df)
        ws.set_column(0,0,6,plain_fmt)         # 序號
        ws.set_column(1,1,72,wrap_fmt)         # 掃描網址
        ws.set_column(2,2,20,plain_fmt)        # 類型
        ws.set_column(3,3,10,plain_fmt)        # 有效連結
        ws.set_column(4,4,72,wrap_fmt)         # 來源站內頁面
        # 標題列自動篩選
        if n>=0:
            last_row=n if n>0 else 0
            ws.autofilter(0,0,last_row,4)
        # 可安全嵌入者改為超連結（掃描網址＝第2欄 index 1）
        for r in range(1,n+1):
            cell_u=str(df.iloc[r-1].get("掃描網址")or"")
            if _excel_cell_url_hyperlink_writable(cell_u):
                try:
                    ws.write_url(r,1,cell_u,link_wrap_fmt,cell_u)
                except Exception:
                    pass
            src_raw=str(df.iloc[r-1].get("來源站內頁面")or"")
            if "\n" not in src_raw and _excel_cell_url_hyperlink_writable(src_raw):
                try:
                    ws.write_url(r,4,src_raw,link_wrap_fmt,src_raw)
                except Exception:
                    pass
    buf.seek(0)
    return buf.getvalue()

# ==========================================
# 檔案與設定管理
# ==========================================
def load_favorites():
    if os.path.exists(FAVORITES_FILE):
        try:
            with open(FAVORITES_FILE,"r",encoding="utf-8") as f:
                data=json.load(f)
            return data if isinstance(data,list) else []
        except:pass
    return []

def save_favorites(items):
    with open(FAVORITES_FILE,"w",encoding="utf-8") as f:
        json.dump(items,f,ensure_ascii=False,indent=4)

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE,"r",encoding="utf-8") as f:
                return json.load(f)
        except:return {}
    return {}

def save_config(config_data):
    with open(CONFIG_FILE,"w",encoding="utf-8") as f:
        json.dump(config_data,f,ensure_ascii=False,indent=4)

def normalize_url_input(url:str)->str:
    u=(url or "").strip()
    if not u:return ""
    if "://" not in u:u="https://"+u
    return u

def get_clean_domain(url):
    if not url:return ""
    parsed=urlparse(url if "://" in url else "http://"+url)
    return parsed.netloc.lower()

def path_scope_base_parts(scope_url:str):
    """依起始掃描網址解析路徑前綴。回傳 (host_lower, path_prefix)；path_prefix 為 None 表示僅網域根（整站）。

    常見站台以「目錄內某一頁」為入口（如 …/Sale/Login），同層另有 …/Sale/WebInfo；此時若仍以完整路徑為前綴會漏掃。
    規則：路徑以 / 結尾者視為目錄，前綴為該目錄；否則若最後一段像檔名（含 .）或路徑深度 ≥3，則前綴改為上一層目錄。
    若要「僅限某子路徑底下」且最後一段無副檔名，請在起始網址末端加上 / 表示目錄（例如 …/Sale/）。
    """
    if not(scope_url or"").strip():return"",None
    p=urlparse(scope_url.strip())
    host=(p.netloc or"").lower()
    raw=(p.path or"")
    if raw in("","/"):return host,None
    is_dir_url=len(raw)>1 and raw.endswith("/")
    base=raw.rstrip("/")
    if not base.startswith("/"):base="/"+base
    segs=[s for s in base.split("/")if s]
    last=segs[-1]if segs else""
    if not is_dir_url:
        if("."in last)or(len(segs)>=3):
            par=str(PurePosixPath(base).parent)
            if par in("/","."):
                return host,None
            base=par
    return host,base

def _random_read_delay_ms()->int:
    """模擬人類閱讀間隔（毫秒）。"""
    return random.randint(1000,3000)

def _site_root_url(url:str)->str:
    """該網域根目錄（含 scheme），供 Referer 等使用。"""
    p=urlparse(normalize_url_input(url)if url else"")
    scheme=(p.scheme or"https").lower()
    netloc=(p.netloc or"").lower()
    if not netloc:return normalize_url_input(url)if url else""
    return urlunparse((scheme,netloc,"/","","",""))

def _pick_user_agent()->str:
    """每次新頁面自池內隨機選用。"""
    return random.choice(_CRAWLER_USER_AGENT_POOL)

def _extra_browser_headers(referer_url:str)->dict:
    """Accept-Language 與 Referer（子資源請求一併帶入）。"""
    return{"Accept-Language":"zh-TW,zh;q=0.9","Referer":referer_url or""}

# 不同 OS／瀏覽器版本之 User-Agent 池（實際仍由 Chromium 發請求，選用主流 Chrome 系字串以降低指紋衝突）
_CRAWLER_USER_AGENT_POOL=[
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36 Edg/120.0.0.0",
    "Mozilla/5.0 (Windows NT 11.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 14_2_1) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.2 Safari/605.1.15",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 13_6_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.6099.199 Safari/537.36",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
    "Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:122.0) Gecko/20100101 Firefox/122.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36 OPR/105.0.0.0",
]

def url_in_scan_scope(link_url:str,scope_url:str)->bool:
    """同網域且 path 落在 path_scope_base_parts 所決定之前綴範圍內（見該函式說明）。

    另：同網域之 /download/、/dowload/ 靜態檔常與應用子路徑分離，仍一併納入掃描。
    """
    if not(scope_url or"").strip():return True
    host,base=path_scope_base_parts(scope_url)
    ll=urlparse((link_url or"").strip())
    if(ll.netloc or"").lower()!=host:return False
    if _is_public_download_storage_url(link_url):
        return True
    if base is None:return True
    lp=ll.path or"/"
    if not lp.startswith("/"):lp="/"+lp
    return lp==base or lp.startswith(base+"/")

# ==========================================
# 智慧型URL編碼校正 (徹底解決中文路徑與重複編碼問題)
# ==========================================
def make_safe_url(url):
    try:
        # 1.先徹底還原成原始中文，避免原本已編碼的連結被重複編碼(%25)
        raw_url=unquote(url)
        parsed=urlparse(raw_url)
        # 2.僅對Path(路徑)進行編碼，且保留/符號與避免破壞中文
        encoded_path=quote(parsed.path,safe='/')
        # 3.重新組合
        return urlunparse(parsed._replace(path=encoded_path))
    except:
        return url

# 掃描／下載連結副檔名（含 OpenDocument；連結可見文字以此結尾者一律納入檢核）
_SCAN_FILE_EXTENSIONS=(".pdf",".doc",".docx",".xls",".xlsx",".ppt",".pptx",".rar",".zip",".odt",".ods",".odp")
# 常見下載目錄內之圖檔（原排除之 .png 等，在此路徑下仍須掃描與驗證連結）
_DOWNLOAD_ASSET_IMAGE_EXTS=(".png",".jpg",".jpeg",".gif",".webp",".svg",".ico",".bmp",".tif",".tiff")

def _is_public_download_storage_url(url:str)->bool:
    """路徑含 /download/ 或常見拼錯 /dowload/（不分大小寫、支援中文路徑解碼後比對）。"""
    try:
        raw=(url or"").strip()
        if not raw:return False
        p=urlparse(unquote(raw))
        pl=(p.path or"/").lower()
        return"/download/"in pl or"/dowload/"in pl
    except Exception:
        return False

def _should_skip_plain_image_url_for_crawl(url:str)->bool:
    """圖檔連結依 **path** 判斷（忽略 ?query／#fragment），不納入站內逐頁 Playwright；/download/ 目錄例外。

    避免 `favicon.png?v=token` 因 endswith 比對不到 .png 而被當成 HTML 掃描，誤判 5.有效連結。
    """
    try:
        pl=(urlparse((url or"").strip()).path or"").lower()
        if not any(pl.endswith(e)for e in _DOWNLOAD_ASSET_IMAGE_EXTS):
            return False
        return not _is_public_download_storage_url(url)
    except Exception:
        return False

# onclick 屬性中常見的頁面跳轉模式（如 location.href='...'、window.open('...')）
_ONCLICK_URL_RE = re.compile(
    r"""(?:(?:window\.|self\.|top\.)?location(?:\.href)?\s*=\s*"""
    r"""|(?:window\.location|location)\.(?:replace|assign)\s*\(\s*"""
    r"""|window\.open\s*\(\s*"""
    r"""|(?:navigate|go|goto|redirect)\s*\(\s*)"""
    r"""['"]([^'"]{3,4096})['"]""",
    re.I,
)

# 看起來像 CSS 選擇器（用於 scrollTo / 錨點切換）而非實際 URL：開頭為 . 或 #，且僅含
# 識別字元／空白／點號／井字。常見於台灣政府網站之 onclick 內 `goto('.div_xxx')` 樣式，
# urljoin 會把它當相對路徑接到目前頁面後造成 404 連結（誤入「3.語系編碼」名單）。
_CSS_SELECTOR_LIKE_RE = re.compile(r"^[.#][A-Za-z][\w\s.#-]*$")

def _looks_like_css_selector(s: str) -> bool:
    if not s:
        return False
    s = s.strip()
    if not s:
        return False
    # 真正的 URL 都會有以下任一：scheme(://) / 路徑分隔(/) / 查詢(?) / 副檔名 .xxx 結尾或路徑中含 /xxx.
    if "://" in s or "/" in s or "?" in s:
        return False
    return bool(_CSS_SELECTOR_LIKE_RE.match(s))

# 常見自訂 data-* URL 屬性名稱
_DATA_URL_ATTRS = ('data-href', 'data-url', 'data-link', 'data-go', 'data-goto', 'data-nav', 'data-target-url')

def _trim_url_tail(u: str) -> str:
    """移除 URL 末端常見尾贅標點，但保留路徑內成對的 ()。

    舊版正則排除路徑中之 ) 並用 rstrip 清掉尾贅，遇到 PDF 檔名含 () 之中文檔（如
    `…(含反詐騙宣導)…顧問).pdf`）會被截斷，致連結被誤判為 5.有效連結／3.語系編碼。
    新版：允許 ) 在 URL 內，僅在「結尾 ) 數量多於 (」時逐一移除。"""
    if not u:
        return u
    while u:
        last = u[-1]
        if last in ".,;:\"'":
            u = u[:-1]
            continue
        if last == ")":
            # 計算 URL 內括號是否平衡：若右括號比左括號多，去掉一個尾端 )。
            if u.count(")") > u.count("("):
                u = u[:-1]
                continue
            break
        break
    return u

# 下載處理器常見 binary MIME 前綴（URL 無副檔名但回傳文件內容）
_BINARY_DOC_MIME_PREFIXES = (
    'application/pdf',
    'application/msword',
    'application/vnd.ms-excel',
    'application/vnd.ms-powerpoint',
    'application/vnd.openxmlformats-officedocument',
    'application/vnd.oasis.opendocument',
    'application/x-rar',
    'application/vnd.rar',
    'application/zip',
    'application/x-zip',
)
# Office/壓縮類型 → 指標 10；PDF 僅列入清單不報錯
_OFFICE_MIME_PARTS = ('msword', 'ms-excel', 'ms-powerpoint', 'openxmlformats', 'opendocument', 'x-rar', 'vnd.rar', '/zip', 'x-zip')

def _document_link_label_regex(file_exts):
    inner="|".join(re.escape(e.lower().lstrip("."))for e in file_exts)
    img="|".join(re.escape(e.lower().lstrip("."))for e in _DOWNLOAD_ASSET_IMAGE_EXTS)
    return re.compile(rf"\.(?:{inner}|{img})\s*$",re.I)

def _soup_anchor_document_label_blob(tag)->str:
    """a／area 上可見檔名：文字、title、alt、區塊圖 alt。"""
    parts=[]
    if getattr(tag,"name",None)=="area"and tag.get("alt"):
        parts.append(str(tag.get("alt")))
    tx=tag.get_text(strip=True)
    if tx:parts.append(tx)
    for attr in("title","aria-label"):
        v=tag.get(attr)
        if v:parts.append(str(v))
    im=tag.find("img")
    if im and im.get("alt"):
        parts.append(str(im.get("alt")))
    return" ".join(parts).strip()

def extract_same_domain_links(html:str,page_url:str,target_domain:str,file_exts=None):
    """從已渲染HTML擷取同網域連結：標籤解析 + 內嵌字串(含script/JSON)備援，避免ASP.NET僅在JS帶出下載網址時漏掃。

    另：若 a／area 之可見文字結尾為常見文件副檔名（如 .PDF、.ODT），一律納入同網域 href（供後續實際開啟檢核）。
    """
    exts=file_exts or _SCAN_FILE_EXTENSIONS
    lab_re=_document_link_label_regex(exts)
    extracted=set()
    td=target_domain.lower()
    soup=BeautifulSoup(html,'html.parser')
    tag_sources=[]
    for tag in soup.find_all(['a','area'],href=True):
        tag_sources.append(tag.get('href'))
    for tag in soup.find_all(['iframe','frame','embed'],src=True):
        tag_sources.append(tag.get('src'))
    for tag in soup.find_all('object'):
        if tag.get('data'):tag_sources.append(tag.get('data'))
    for tag in soup.find_all('link',href=True):
        rel=(tag.get('rel') or [])
        if isinstance(rel,str):rel=[rel]
        rel_l=[x.lower() for x in rel]
        if 'stylesheet' in rel_l or 'preconnect' in rel_l or 'dns-prefetch' in rel_l:continue
        tag_sources.append(tag.get('href'))
    for raw_href in tag_sources:
        if not raw_href:continue
        rh=raw_href.strip()
        if rh.lower().startswith(('javascript:','mailto:','tel:','#','data:')):continue
        full=urljoin(page_url,rh).split('#')[0]
        if urlparse(full).netloc.lower()!=td:continue
        if _should_skip_plain_image_url_for_crawl(full):continue
        extracted.add(full)
    try:
        flat=html_unescape(html)
        host=re.escape(td)
        # 路徑字元集不再排除 )；尾贅標點以 _trim_url_tail（會平衡括號）統一處理，
        # 避免將 PDF 中文檔名含 () 之同網域連結截成假連結後送進待掃佇列。
        for m in re.finditer(rf'https?://{host}(?:/[^\s\"\'<>]*)?',flat,re.I):
            u=_trim_url_tail(m.group(0))
            if not u:continue
            if u.lower().startswith(('javascript:','mailto:','tel:','data:')):continue
            full=urljoin(page_url,u).split('#')[0]
            if urlparse(full).netloc.lower()!=td:continue
            if _should_skip_plain_image_url_for_crawl(full):continue
            extracted.add(full)
    except Exception:pass
    for tag in soup.find_all(["a","area"],href=True):
        blob=_soup_anchor_document_label_blob(tag)
        if not blob or not lab_re.search(blob):
            continue
        rh=(tag.get("href")or"").strip()
        if rh.lower().startswith(("javascript:","mailto:","tel:","#","data:")):
            continue
        full=urljoin(page_url,rh).split("#")[0]
        if urlparse(full).netloc.lower()!=td:
            continue
        if _should_skip_plain_image_url_for_crawl(full):
            continue
        extracted.add(full)
    # ── onclick 屬性中的跳轉 URL（ASP.NET / 舊式台灣政府網站常見以 onclick 取代 href）──
    def _add_if_same_domain(rh:str):
        if not rh or rh.lower().startswith(("javascript:","mailto:","tel:","#","data:")):
            return
        # CSS 選擇器（如 `.div_xxx`、`#header`）多半是 scroll/anchor 目標，urljoin 後會
        # 變成 `…/EN/.div_xxx` 之類無效路徑（IIS 回 404+big5 而被誤判 3.語系編碼）。
        if _looks_like_css_selector(rh):
            return
        full=urljoin(page_url,rh).split("#")[0]
        if urlparse(full).netloc.lower()!=td:
            return
        if _should_skip_plain_image_url_for_crawl(full):
            return
        extracted.add(full)
    for tag in soup.find_all(True,onclick=True):
        onclick_val=tag.get("onclick")or""
        for m in _ONCLICK_URL_RE.finditer(onclick_val):
            _add_if_same_domain(m.group(1).strip())
    # ── data-href / data-url 等自訂屬性（Vue / React / 舊式 JS 框架常見）──
    for tag in soup.find_all(True):
        for attr in _DATA_URL_ATTRS:
            val=(tag.get(attr)or"").strip()
            if val:
                _add_if_same_domain(val)
    return extracted

# 政府／大型站台在**高併發 HEAD/GET** 下易短暫逾時或重連；過短的連線逾時 + 單次探測會誤判為不可達。
# 略放寬逾時，並在 `_probe_external_url_sync` 內對失敗結果做少量退避重試（見 `_EXTERNAL_PROBE_ATTEMPTS`）。
_PROBE_REQ_TIMEOUT=(12, 30)
# 同一批 Playwright 並行頁面**共用**此外站探測併發；過高易觸發對端限速／中斷，反增假陰性
_BATCH_EXTERNAL_PROBE_CONCURRENCY=8
# 外站探測：同一 URL 組（含 http→https、尾端 / 變體）全滅後再重試之回合數（間隔遞增 sleep）
_EXTERNAL_PROBE_ATTEMPTS=3
# 單頁含「開頁＋擷取連結＋**全部**外站實測」總上限（外站多時必須夠長）
_PAGE_SCAN_TIMEOUT_S=900
_urllib3_insecure_warn_disabled=False

def extract_external_http_links(html:str,page_url:str,site_host:str)->set:
    """擷取與 site_host 不同網域之 http(s) 連結（a／area 的 href），供連線有效性檢查。"""
    out=set()
    sh=(site_host or"").lower()
    soup=BeautifulSoup(html,"html.parser")
    for tag in soup.find_all(["a","area"],href=True):
        rh=(tag.get("href")or"").strip()
        if not rh or rh.lower().startswith(("javascript:","mailto:","tel:","#","data:")):
            continue
        full=urljoin(page_url,rh).split("#")[0]
        try:
            p=urlparse(full)
        except Exception:
            continue
        if(p.scheme or"").lower()not in("http","https"):
            continue
        host=(p.netloc or"").lower()
        if not host or host==sh:
            continue
        if len(full)>4096:
            continue
        out.add(full)
    return out

async def _external_http_hrefs_from_item_live_dom(page, page_url: str, site_host: str) -> set:
    """從**目前 DOM** 的 #Item 讀取 a／area 的**已解析 href**（等同開發者工具內 a.href），補
    Vue/SPA 在少數情況下與 `page.content()`+BeautifulSoup 序列化之落差，避免國外分頁列表漏列。"""
    sh = (site_host or "").lower()
    if not sh:
        return set()
    _ = page_url  # 與 extract_external 參數簽名對齊；瀏覽器內以 baseURI 解析
    try:
        hrefs = await page.evaluate(
            """(sh) => {
          const root = document.querySelector("#Item");
          if (!root) return [];
          const out = [];
          const seen = new Set();
          for (const el of root.querySelectorAll("a[href], area[href]")) {
            const raw = (el.getAttribute("href") || "").trim();
            if (!raw) continue;
            const l = raw.toLowerCase();
            if (l.startsWith("javascript:") || l.startsWith("mailto:") || l.startsWith("tel:") || l.startsWith("#")
                || l.startsWith("data:")) continue;
            let u;
            try { u = new URL(el.href, document.baseURI).href; } catch (e) { continue; }
            u = u.split("#")[0];
            if (!u || u.length > 4096) continue;
            let p; try { p = new URL(u); } catch (e) { continue; }
            if (p.protocol !== "http:" && p.protocol !== "https:") continue;
            if (!p.hostname) continue;
            if (p.hostname.toLowerCase() === sh) continue;
            if (seen.has(u)) continue;
            seen.add(u);
            out.push(u);
          }
          return out;
        }""",
            sh,
        )
    except Exception:
        return set()
    return {u for u in (hrefs or []) if u}

def _http_status_reachable_for_external_check(status:int)->bool:
    if status is None:
        return False
    if status<400:
        return True
    if status in(401,403,407,429):
        return True
    if status==503:
        return True
    return False

def _build_external_probe_headers(referer:str)->dict:
    """盡量貼近瀏覽器，降低主機不支援 HEAD、CDN/社群阻擋造成的誤判。"""
    h={**_extra_browser_headers(referer),"User-Agent":_pick_user_agent()}
    h["Accept"]="text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
    h.setdefault("Accept-Language","zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7")
    h["Accept-Encoding"]="identity"
    h["DNT"]="1"
    h["Upgrade-Insecure-Requests"]="1"
    return h

# 常見「縮網址服務」host：對方 HTTP 仍 200，但內文寫『轉址失敗／找不到該網址』即代表
# 原縮址已失效，瀏覽器點擊後使用者看到的是錯誤頁。需於 probe 階段一併偵測為失效。
_URL_SHORTENER_HOSTS = {
    "reurl.cc","lihi.cc","lihi1.cc","lihi2.cc","lihi3.cc","lihi.io","lihi.com",
    "pse.is","psee.io","bit.ly","goo.gl","tinyurl.com","t.co",
    "ppt.cc","0rz.tw","is.gd","buff.ly","s.id","cutt.ly",
}

# 縮網址「找不到該短址」之高鑑別度提示語：須確切點名「不存在／找不到／已過期／已失效」，
# 避免短址服務暫時忙碌（出現『系統忙碌中』之類）時誤殺有效連結。
_SHORTENER_NOT_FOUND_TEXTS = (
    "找不到該網址","短網址不存在","短網址已失效","網址不存在","網址已失效","已過期",
    "Short URL not found","URL not found","Link not found","Link is invalid",
    "Link expired","does not exist","not been found",
)


def _host_is_url_shortener(host:str)->bool:
    if not host:
        return False
    h=host.lower().lstrip(".")
    if h.startswith("www."):h=h[4:]
    if h in _URL_SHORTENER_HOSTS:
        return True
    return any(h.endswith("."+d) for d in _URL_SHORTENER_HOSTS)


def _looks_like_shortener_soft_404(host:str,body_text:str)->bool:
    """縮網址後台說「找不到」的軟性 404；HTTP 仍 200 但內容明示連結失效。

    僅以「找不到／不存在／已過期」之精準字串為判定依據，避免短網址服務暫時忙碌時誤判。
    """
    if not body_text or not _host_is_url_shortener(host):
        return False
    return any(p in body_text for p in _SHORTENER_NOT_FOUND_TEXTS)


# 純 JS 轉址停車頁特徵：body 極短、僅含 `window.location` 跳轉，無人類可讀內容。
# 例（yilannews.xyz/?p=…）：
#   <!DOCTYPE html><html><head><script>window.onload=function(){window.location.href="/lander?…"}</script></head></html>
# 此類網址原文章已下線，網域被當成停車／導流頁，視同失效；應自 5.有效連結 視角列入失敗。
_JS_REDIRECT_HINTS=("window.location","location.replace","location.assign","location.href")


def _looks_like_js_redirect_parking(body_text:str)->bool:
    if not body_text:
        return False
    snippet=body_text[:2048]
    if len(snippet)>1500:
        return False
    low=snippet.lower()
    if not any(h in low for h in _JS_REDIRECT_HINTS):
        if "<meta" in low and "http-equiv" in low and "refresh" in low:
            pass # meta refresh 也算
        else:
            return False
    # 去掉 <script>...</script> 與所有標籤後內容應極短：典型停車頁清掉後僅剩空白
    text_only=re.sub(r"<script[\s\S]*?</script>","",snippet,flags=re.I)
    text_only=re.sub(r"<style[\s\S]*?</style>","",text_only,flags=re.I)
    text_only=re.sub(r"<[^>]+>","",text_only).strip()
    return len(text_only)<30


def _read_response_text_snippet(r,max_bytes:int=8192)->str:
    """從 streaming Response 讀最多 max_bytes 解碼成字串；非 HTML/text 直接回空。"""
    ct=(r.headers.get("content-type")or"").lower()
    if ct and ("text/" not in ct and "html" not in ct and "xml" not in ct and "json" not in ct):
        return ""
    try:
        chunks,total=[],0
        for chunk in r.iter_content(chunk_size=2048,decode_unicode=False):
            if chunk is None:break
            chunks.append(chunk)
            total+=len(chunk)
            if total>=max_bytes:break
        raw=b"".join(chunks)
    except Exception:
        return ""
    # encoding：依 r.apparent_encoding 或 utf-8 fallback
    enc=(r.encoding or"").lower()
    if not enc or enc in("iso-8859-1",):
        # iso-8859-1 是 requests 對未指定 charset 時之預設，對中文站台幾乎都不對；改 utf-8
        enc="utf-8"
    try:
        return raw.decode(enc,errors="ignore")
    except Exception:
        try:
            return raw.decode("utf-8",errors="ignore")
        except Exception:
            return ""


def _get_probe_get_result(url:str,headers:dict)->Tuple[int,str,str]:
    """送 GET，回傳 (status_code, content_type, body_snippet[≤8KB])。
    body_snippet 用於『soft-404／JS 停車頁』檢查；非 HTML 內容回空字串以省頻寬。
    """
    r=requests.get(url,headers=headers,timeout=_PROBE_REQ_TIMEOUT,allow_redirects=True,verify=False,stream=True)
    try:
        code=r.status_code
        ct=(r.headers.get("content-type")or"").lower()
        # 部分社群（如 Facebook sharer）對程式請求回 400，但 Content-Type 仍為 HTML／或 charset 帶引號
        if code==400 and("text/html"in ct or"/html"in ct or"html"in ct):
            code=200
        body=""
        if code<400 and ("text/html"in ct or"html"in ct or"text/"in ct or not ct):
            body=_read_response_text_snippet(r)
        return code,ct,body
    finally:
        try:
            r.close()
        except Exception:
            pass

def _probe_url_variants_trailing_slash(u:str):
    """path 末端多一個 / 時，部分站台（如 moenv /page/{id}/）會 404；補試去掉尾端 / 的網址。"""
    yield u
    try:
        p=urlparse(u)
        path=p.path or""
        if len(path)>1 and path.endswith("/"):
            alt=urlunparse((p.scheme,p.netloc,path.rstrip("/"),p.params,p.query,""))
            if alt!=u:
                yield alt
    except Exception:
        pass

def _try_probe_single_url(url:str,headers:dict)->bool:
    """對單一 URL 先 HEAD；若未判定為可連線（含少見 4xx）一律再試 GET 複核。

    部分主機對 HEAD 回 404/405/501 或**非列表內之 4xx**，但 GET 仍正常；舊版僅對固定幾種
    HEAD 狀態改試 GET，其餘直接 False，易與瀏覽器不一致。

    社群／分享頁（Facebook sharer 等）對 HEAD 也會回 400+text/html，瀏覽器仍可正常開啟；
    本函式在 HEAD 端同樣套用「400 且回 HTML 視為可達」之放寬規則。

    額外：若 HTTP 回 200 但內文為下列兩種情況之一，仍視為失效：
      (a) 縮網址服務（reurl.cc、lihi.cc、pse.is、bit.ly…）顯示「轉址失敗／找不到該網址」
          → 該短網址已不存在或已被刪除。
      (b) Body < 1.5KB 且只含 `window.location` 類 JS 跳轉（停車／導流頁特徵）
          → 原網址早已下線、域名被改作其他用途。
    這兩種情況 HEAD 帶不出蛛絲馬跡，必須讀取 GET 內文才能辨識。"""
    def ok_status(code:Optional[int])->bool:
        return _http_status_reachable_for_external_check(code)
    try:
        host=urlparse(url).netloc
    except Exception:
        host=""
    is_shortener=_host_is_url_shortener(host)

    def _get_and_validate()->bool:
        try:
            code,_ct,body=_get_probe_get_result(url,headers)
        except Exception:
            return False
        if not ok_status(code):
            return False
        if body:
            if is_shortener and _looks_like_shortener_soft_404(host,body):
                return False
            if _looks_like_js_redirect_parking(body):
                return False
        return True

    try:
        r=requests.head(url,headers=headers,timeout=_PROBE_REQ_TIMEOUT,allow_redirects=True,verify=False)
        sc=r.status_code
        ct=(r.headers.get("content-type")or"").lower()
        head_ok=ok_status(sc) or (sc==400 and("text/html"in ct or"/html"in ct or"html"in ct))
        if head_ok:
            # 縮網址的 HEAD 也會 200（後端服務本身存在），需 GET 拿內文確認；
            # 同樣，極短停車頁可能 HEAD 200，body 才看得出 JS-only 跳轉，故統一以 GET 複核。
            if is_shortener:
                return _get_and_validate()
            # 一般情形：HEAD 200 直接視為可達；停車頁多半 HEAD 405 或 GET 才看得到，下方 except 路徑覆蓋。
            return True
        # 非 ok 狀態 → 試 GET（含 body 檢查）
        return _get_and_validate()
    except Exception:
        pass
    return _get_and_validate()

def _probe_external_url_sync(url:str,referer:str)->bool:
    """以 HEAD 先詢、必要時再 GET 粗查外部網址（verify=False）。

    許多主機未實作 HEAD 或回 404，實際 GET 仍 200，僅以 HEAD 會誤列「不符合」；另若連線/SSL
    在 HEAD 失敗也會再試 GET。

    **http→https 優先**：若 href 為 `http://`，先試同 host／path／query 的 `https://`，再試原 URL。
    避免「http 只 302 到已換網域之中繼站」導致 DNS/SSL 失敗，而瀏覽器直接開 https 首頁卻正常
    （例如 epamail.moenv.gov.tw）。

    **尾端 /**：若 href 為 `…/page/ID/`，先試原網址再試去掉尾端 `/`（與 moenv 等站台行為一致）。

    對**同一組**候選網址若皆失敗，會以短暫遞增間隔最多重試 `_EXTERNAL_PROBE_ATTEMPTS` 次，以降低
    高併發下對方節流／瞬斷／TLS 握手逾時造成之假陰性（如 moenv 各子域同批驗證）。
    """
    global _urllib3_insecure_warn_disabled
    if not _urllib3_insecure_warn_disabled:
        try:
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        except Exception:
            pass
        _urllib3_insecure_warn_disabled=True
    headers=_build_external_probe_headers(referer)
    raw=(url or"").strip()
    if not raw:
        return False
    candidates=[raw]
    if raw.lower().startswith("http://"):
        try:
            p=urlparse(raw)
            if p.netloc:
                https_u=urlunparse(("https",p.netloc,p.path or"",p.params,p.query,""))
                if https_u and https_u!=raw:
                    candidates=[https_u,raw]
        except Exception:
            pass
    for attempt in range(max(1,_EXTERNAL_PROBE_ATTEMPTS)):
        seen=set()
        for u in candidates:
            u=(u or"").strip()
            if not u:
                continue
            for v in _probe_url_variants_trailing_slash(u):
                if not v or v in seen:
                    continue
                seen.add(v)
                if _try_probe_single_url(v,headers):
                    return True
        if attempt+1<max(1,_EXTERNAL_PROBE_ATTEMPTS):
            try:
                time.sleep(0.7*(attempt+1))
            except Exception:
                pass
    return False

def _probe_cache_get(probe_cache: dict, url: str, *, fail_ttl_s: int = 600) -> Optional[bool]:
    """外站探測快取讀取。

    目的：避免「暫時性失敗」被永久快取為 False，導致正常連結整批掃描都被誤判失效。
    - True：永遠視為命中（不再探測）。
    - False：僅在 fail_ttl_s 秒內視為命中；超過則允許重新探測一次。

    相容舊格式：probe_cache[url] 可能是 bool；新版會寫入 {'ok': bool, 'ts': float}。
    """
    if not probe_cache or not url:
        return None
    v = probe_cache.get(url)
    if v is None:
        return None
    if isinstance(v, bool):
        # 舊版：False 可能是暫時性錯誤；不要永遠相信它。
        return True if v else None
    if isinstance(v, dict):
        ok = v.get("ok")
        if ok is True:
            return True
        if ok is False:
            try:
                ts = float(v.get("ts") or 0.0)
            except Exception:
                ts = 0.0
            if ts > 0 and (time.time() - ts) < float(fail_ttl_s):
                return False
            return None
    return None


def _probe_cache_set(probe_cache: dict, url: str, ok: bool) -> None:
    if not isinstance(probe_cache, dict) or not url:
        return
    probe_cache[url] = {"ok": bool(ok), "ts": float(time.time())}


async def probe_external_links_unreachable(
    urls:set,referer:str,probe_cache:dict,probe_sem:Optional[asyncio.Semaphore]=None
)->Tuple[List[str],List[str]]:
    """(無法連線或 HTTP 判定失效, 掃到之**全部**外站 URL 字母序清單)。

    **每一筆**外站 URL 皆會實際送出 HEAD/GET 驗證（已快取者不重複請求）。單頁外站極多時掃描會變慢，
    屬預期行為。probe_cache 鍵為完整 URL，值為是否可連。
    probe_sem：同一批並行掃描頁面共用，限制全批外站 HTTP 總併發。
    """
    if not urls:
        return [], []
    all_sorted=sorted(urls)
    sem=probe_sem if probe_sem is not None else asyncio.Semaphore(_BATCH_EXTERNAL_PROBE_CONCURRENCY)
    failed=[]

    async def one(u:str):
        cached = _probe_cache_get(probe_cache, u, fail_ttl_s=600)
        if cached is True:
            return None
        if cached is False:
            return u
        async with sem:
            ok=await asyncio.to_thread(_probe_external_url_sync,u,referer)
        _probe_cache_set(probe_cache, u, ok)
        return u if not ok else None

    _chunk=128
    for i in range(0,len(all_sorted),_chunk):
        part=all_sorted[i:i+_chunk]
        for u in await asyncio.gather(*[one(x)for x in part]):
            if u:
                failed.append(u)
    return failed, all_sorted

def _url_matches_download_candidate(u:str,td:str,file_exts):
    try:
        if urlparse(u).netloc.lower()!=td:return False
        low=u.lower().split("?",1)[0]
        if any(low.endswith(ext) for ext in file_exts):return True
        if _is_public_download_storage_url(u):
            return True
        if "/elearning"in low:
            return True
        return False
    except Exception:
        return False

async def _gather_file_button_locators(frame,max_n):
    """ASP.NET常見：LinkButton外層是<a>內只有圖、文字在title/alt，或包在<span>內；用多組定位並以座標去重。"""
    seen_xy,cands=set(),[]
    groups=[
        frame.get_by_text("檔案",exact=True),
        frame.locator('[title*="檔案"]'),
        frame.locator('[aria-label*="檔案"]'),
        frame.locator('img[alt*="檔案"]'),
        frame.locator('a, button, input[type="button"], input[type="submit"], [role="button"]').filter(has_text=re.compile(r"檔案")),
    ]
    for grp in groups:
        try:
            n=await grp.count()
        except Exception:
            continue
        for i in range(n):
            if len(cands)>=max_n:break
            el=grp.nth(i)
            try:
                if not await el.is_visible():continue
                box=await el.bounding_box()
                if not box:continue
                key=(round(box["x"],0),round(box["y"],0))
                if key in seen_xy:continue
                tag=(await el.evaluate("e => e.tagName")).upper()
                if tag=="INPUT":
                    val=await el.get_attribute("value")or""
                    if "檔案"not in val:continue
                    lab=val
                else:
                    lab=(await el.inner_text()).strip()
                    tit=await el.get_attribute("title")or""
                    al=await el.get_attribute("aria-label")or""
                    if tag=="IMG":
                        lab=await el.get_attribute("alt")or""
                    blob=lab+tit+al
                    if "檔案"not in blob:continue
                    if "請先登入"in blob or"檔案內容"in blob:continue
                    if tag!="IMG" and len(lab)>80:continue
                seen_xy.add(key);cands.append(el)
            except Exception:
                continue
        if len(cands)>=max_n:break
    try:
        imgs=frame.locator("table").filter(has_text="檔案內容").locator('input[type="image"]')
        ni=await imgs.count()
        for i in range(ni):
            if len(cands)>=max_n:break
            el=imgs.nth(i)
            try:
                if not await el.is_visible():continue
                td=el.locator("xpath=ancestor::td[1]")
                if await td.count()==0:continue
                tdt=(await td.first.inner_text()).strip()
                if "請先登入"in tdt:continue
                box=await el.bounding_box()
                if not box:continue
                key=(round(box["x"],0),round(box["y"],0))
                if key in seen_xy:continue
                seen_xy.add(key);cands.append(el)
            except Exception:
                continue
    except Exception:
        pass
    return cands

_CLICK_DISCOVER_BUDGET_S = 40  # 點擊探索函式整體時間上限（秒）；避免公告等頁面卡住整批掃描

async def discover_download_links_from_file_buttons(page,browser_context,target_domain,file_exts,max_clicks=20):
    """部分站台(如課程列表)的PDF僅在點擊「檔案」後才發請求；以自動點擊+監聽request/response/download/同頁網址/新分頁補齊連結。"""
    td=target_domain.lower()
    out=set()
    restore=page.url
    t_start=time.monotonic()
    frames=[]
    try:
        seen_f=set()
        for f in [page.main_frame]+[x for x in page.frames if x is not page.main_frame]:
            if id(f) in seen_f:continue
            seen_f.add(id(f))
            fu=f.url or""
            if f is page.main_frame:
                frames.append(f);continue
            if fu.lower().startswith("about:")or urlparse(fu).netloc.lower()==td:
                frames.append(f)
    except Exception:
        frames=[page.main_frame]
    per_frame=max(10,max_clicks//max(1,len(frames)))
    for fr in frames:
        cands=await _gather_file_button_locators(fr,per_frame)
        for el in cands[:per_frame]:
            if time.monotonic()-t_start>_CLICK_DISCOVER_BUDGET_S:
                break
            seen,captured,dl_urls,resp_urls=set(),[],[],[]
            def on_request(request):
                try:
                    u=request.url
                    if not _url_matches_download_candidate(u,td,file_exts):return
                    b=u.split("#")[0]
                    if b not in seen:
                        seen.add(b);captured.append(b)
                except Exception:pass
            def on_response(response):
                try:
                    u=response.url
                    if td not in u.lower():return
                    if not _url_matches_download_candidate(u,td,file_exts):return
                    ct=(response.headers.get("content-type")or"").lower()
                    if any(u.lower().split("?")[0].endswith(ext) for ext in file_exts)or"pdf"in ct or"octet-stream"in ct or"msword"in ct or"officedocument"in ct or"opendocument"in ct:
                        b=u.split("#")[0]
                        if b not in seen:
                            seen.add(b);resp_urls.append(b)
                except Exception:pass
            def on_download(d):
                try:
                    u=d.url
                    if u and urlparse(u).netloc.lower()==td:dl_urls.append(u.split("#")[0])
                except Exception:pass
            page.on("request",on_request)
            page.on("response",on_response)
            page.on("download",on_download)
            pages_before=len(browser_context.pages)
            try:
                await el.scroll_into_view_if_needed()
                await el.click(timeout=20000)
            except Exception:
                try:page.remove_listener("request",on_request)
                except Exception:pass
                try:page.remove_listener("response",on_response)
                except Exception:pass
                try:page.remove_listener("download",on_download)
                except Exception:pass
                continue
            await page.wait_for_timeout(2000)
            try:page.remove_listener("request",on_request)
            except Exception:pass
            try:page.remove_listener("response",on_response)
            except Exception:pass
            try:page.remove_listener("download",on_download)
            except Exception:pass
            for u in captured+dl_urls+resp_urls:out.add(u)
            try:
                cur=page.url.split("#")[0]
                if _url_matches_download_candidate(cur,td,file_exts):out.add(cur)
            except Exception:pass
            try:
                for p in browser_context.pages[pages_before:]:
                    try:
                        u=p.url
                        if urlparse(u).netloc.lower()==td:out.add(u.split("#")[0])
                    except Exception:pass
                    try:await p.close()
                    except Exception:pass
            except Exception:pass
            try:
                if page.url.split("#")[0]!=restore.split("#")[0]:
                    await page.goto(restore,wait_until="domcontentloaded",timeout=45000)
                    await page.wait_for_timeout(1500)
            except Exception:
                break
        try:
            if page.url.split("#")[0]!=restore.split("#")[0]:
                await page.goto(restore,wait_until="domcontentloaded",timeout=45000)
                await page.wait_for_timeout(1000)
        except Exception:pass
    try:
        if page.url.split("#")[0]!=restore.split("#")[0]:
            await page.goto(restore,wait_until="domcontentloaded",timeout=45000)
            await page.wait_for_timeout(1000)
    except Exception:pass
    return out

async def discover_document_labeled_anchor_urls(page,browser_context,target_domain,file_exts,max_clicks=20):
    """連結可見文字結尾為 .PDF、.ODT 等檔名者：納入待掃；若 href 非直接檔案網址（如 PostBack、.aspx），則點擊並監聽網路以取得實際下載網址。"""
    label_re=_document_link_label_regex(file_exts)
    td=target_domain.lower()
    out=set()
    restore=page.url
    frames=[]
    try:
        seen_f=set()
        for f in [page.main_frame]+[x for x in page.frames if x is not page.main_frame]:
            if id(f)in seen_f:continue
            seen_f.add(id(f))
            fu=f.url or""
            if f is page.main_frame:
                frames.append(f);continue
            if fu.lower().startswith("about:")or urlparse(fu).netloc.lower()==td:
                frames.append(f)
    except Exception:
        frames=[page.main_frame]
    clicks=0
    t_start=time.monotonic()
    for fr in frames:
        if time.monotonic()-t_start>_CLICK_DISCOVER_BUDGET_S:
            break
        base_u=fr.url if(fr.url and not fr.url.lower().startswith("about:"))else page.url
        try:
            loc=fr.locator("a[href], area[href]")
            ni=await loc.count()
        except Exception:
            continue
        for i in range(min(ni,300)):
            el=loc.nth(i)
            try:
                if not await el.is_visible():
                    continue
                label=await el.evaluate(r"""e=>{
                  const t=(e.innerText||'').replace(/\s+/g,' ').trim();
                  if(t)return t;
                  const g=(e.getAttribute('title')||'')+' '+(e.getAttribute('aria-label')||'');
                  const im=e.querySelector&&e.querySelector('img');
                  const a=im?im.getAttribute('alt'):'';
                  if(e.tagName==='AREA')return((e.getAttribute('alt')||'')+' '+g+a).trim();
                  return(g+a).trim();
                }""")
                if not label or not label_re.search(label.strip()):
                    continue
                href=await el.get_attribute("href")
                if not href:
                    continue
                rh=href.strip()
                full=urljoin(base_u,rh).split("#")[0]
                if urlparse(full).netloc.lower()!=td:
                    continue
                low_path=full.lower().split("?",1)[0]
                href_has_doc_ext=any(low_path.endswith(ext)for ext in file_exts)
                rhl=rh.lower()
                is_special_href=rhl.startswith(("javascript:","mailto:","tel:","data:"))or rhl in("#","javascript:void(0)","javascript:void(0);")or(rh.strip()=="#")
                if href_has_doc_ext and not is_special_href:
                    out.add(full)
                    continue
                if clicks>=max_clicks or time.monotonic()-t_start>_CLICK_DISCOVER_BUDGET_S:
                    continue
                seen,captured,dl_urls,resp_urls=set(),[],[],[]
                def on_request(request):
                    try:
                        u=request.url
                        if not _url_matches_download_candidate(u,td,file_exts):return
                        b=u.split("#")[0]
                        if b not in seen:
                            seen.add(b);captured.append(b)
                    except Exception:pass
                def on_response(response):
                    try:
                        u=response.url
                        if td not in u.lower():return
                        if not _url_matches_download_candidate(u,td,file_exts):return
                        ct=(response.headers.get("content-type")or"").lower()
                        if any(u.lower().split("?")[0].endswith(ext)for ext in file_exts)or"pdf"in ct or"octet-stream"in ct or"msword"in ct or"officedocument"in ct or"opendocument"in ct:
                            b=u.split("#")[0]
                            if b not in seen:
                                seen.add(b);resp_urls.append(b)
                    except Exception:pass
                def on_download(d):
                    try:
                        u=d.url
                        if u and urlparse(u).netloc.lower()==td:dl_urls.append(u.split("#")[0])
                    except Exception:pass
                page.on("request",on_request)
                page.on("response",on_response)
                page.on("download",on_download)
                pages_before=len(browser_context.pages)
                try:
                    await el.scroll_into_view_if_needed()
                    await el.click(timeout=20000)
                except Exception:
                    try:page.remove_listener("request",on_request)
                    except Exception:pass
                    try:page.remove_listener("response",on_response)
                    except Exception:pass
                    try:page.remove_listener("download",on_download)
                    except Exception:pass
                    continue
                clicks+=1
                await page.wait_for_timeout(2000)
                try:page.remove_listener("request",on_request)
                except Exception:pass
                try:page.remove_listener("response",on_response)
                except Exception:pass
                try:page.remove_listener("download",on_download)
                except Exception:pass
                for u in captured+dl_urls+resp_urls:out.add(u)
                try:
                    cur=page.url.split("#")[0]
                    if _url_matches_download_candidate(cur,td,file_exts):out.add(cur)
                except Exception:pass
                try:
                    for p in browser_context.pages[pages_before:]:
                        try:
                            u=p.url
                            if urlparse(u).netloc.lower()==td:out.add(u.split("#")[0])
                        except Exception:pass
                        try:await p.close()
                        except Exception:pass
                except Exception:pass
                try:
                    if page.url.split("#")[0]!=restore.split("#")[0]:
                        await page.goto(restore,wait_until="domcontentloaded",timeout=45000)
                        await page.wait_for_timeout(1500)
                except Exception:
                    break
            except Exception:
                continue
        try:
            if page.url.split("#")[0]!=restore.split("#")[0]:
                await page.goto(restore,wait_until="domcontentloaded",timeout=45000)
                await page.wait_for_timeout(1000)
        except Exception:pass
    try:
        if page.url.split("#")[0]!=restore.split("#")[0]:
            await page.goto(restore,wait_until="domcontentloaded",timeout=45000)
            await page.wait_for_timeout(1000)
    except Exception:pass
    return out

# 課程總覽等「網址不變、AJAX/PostBack 分頁」頁面：僅掃第一頁會漏掉內頁 PDF
PAGINATION_MAX_STEPS=100

async def _dom_fingerprint_for_pagination(page):
    """偵測分頁是否已切換（同 URL 下比對列表內文）。

    若整頁第一個 table 為導覽、版面用小型 table，內文不變→會誤判『永遠沒翻頁』。改從
    main／#content 內**較大**的表格或主區內文取紋，才會隨清單列內容變化。
    """
    try:
        return await page.evaluate(r"""() => {
          const M = (document.querySelector("main, [role=main], #content, article") || document.body);
          if (!M) return "";
          let s = "";
          const item = document.querySelector("#Item, #Item ol, main section .content ol, main ol");
          if (item) {
            const w = item.innerText || "";
            if (w.length > 80) { s = w; }
          }
          if (!s) {
            const tables = Array.from(M.querySelectorAll("table"));
            for (const t of tables) {
              const w = t.innerText || "";
              if (w.length > 200) { s = w; break; }
            }
          }
          if (!s) s = M.innerText || "";
          return s.slice(0, 10000);
        }""")
    except Exception:
        return ""

async def _item_http_href_fingerprint(page) -> str:
    """#Item 內外連／列表連結的穩定指紋。Vue/SPA 翻頁只換列表時，內文 slice 有時不變，導致誤判未翻頁。"""
    try:
        return await page.evaluate(r"""() => {
          const r = document.querySelector("#Item");
          if (!r) return "";
          const out = new Set();
          for (const a of r.querySelectorAll("a[href]")) {
            let u;
            try { u = new URL(a.href, document.baseURI).href; } catch (e) { continue; }
            const low = (u || "").toLowerCase();
            if (low.startsWith("javascript:") || low.startsWith("mailto:")) continue;
            out.add(u.split("#")[0]);
          }
          return Array.from(out).sort().join("\n");
        }""")
    except Exception:
        return ""

async def _pagination_state_key(page) -> str:
    """分頁前後比對用：優先 #Item 連結集合，否則用 main 內文指紋。"""
    h = await _item_http_href_fingerprint(page)
    if len(h) >= 8:
        return h
    d = await _dom_fingerprint_for_pagination(page)
    return h + "\n--\n" + d if h else d

async def _wait_pagination_state_change(page, st0: str, max_ms: int = 48000) -> bool:
    """翻頁後等列表更新（多數為客端路由，不觸發 navigation）。慢機／Flask 執行緒下 22s 常不足導致誤判無下一頁。"""
    t0 = time.monotonic()
    while (time.monotonic() - t0) * 1000 < max_ms:
        if await _pagination_state_key(page) != st0:
            return True
        await page.wait_for_timeout(220)
    return False

_PAGER_CLICK_JS=r"""
() => {
  const M = document.querySelector("main, [role=main], #content, article, body");
  if (!M) return "";
  const isPagerRow = (tr) => {
    const t = (tr.textContent || "").replace(/\\s+/g, " ").trim();
    if (t.length > 500) return false;
    if (/[年月日時分秒]|https?:\\/\\/|表單|說明|瀏覽|下載|單位|地址|電話/.test(t)) return false;
    const d = t.replace(/[^0-9]/g, "");
    if (d.length < 2) return false;
    if (!/1/.test(t) || !/2/.test(t)) return false;
    return true;
  };
    for (const tr of M.querySelectorAll("tr")) {
    if (!isPagerRow(tr)) continue;
    for (const el of tr.querySelectorAll("a, button, [role=button], input[type=button], input[type=image]")) {
      if (!el.offsetParent) continue;
      const s = (el.textContent || "").replace(/\\s+/g, "").trim();
      if (s === ">" || s === "›" || s === "»" || s === "＞" || s === "》") {
        el.click();
        return "next";
      }
    }
  }
  for (const tr of M.querySelectorAll("tr")) {
    if (!isPagerRow(tr)) continue;
    const tds = Array.from(tr.querySelectorAll("td, th"));
    if (tds.length < 2) continue;
    let cur = 0;
    tds.forEach((td) => {
      const c = (td.className || "");
      if (/\b(active|current|selected)\b/i.test(c) || (td.getAttribute("aria-selected") || "") === "true" || (td.getAttribute("aria-current") || "") === "page") {
        const raw = (td.textContent || "").replace(/\\s+/g, "").match(/^(\\d{1,2})/);
        if (raw) { cur = parseInt(raw[1], 10); }
      }
    });
    const links = Array.from(tr.querySelectorAll("a, button, [role=button], input[type=button]"));
    const byNum = {};
    links.forEach((el) => {
      if (!el.offsetParent) return;
      const s = (el.textContent || "").replace(/\\s+/g, "").match(/^(\\d{1,2})$/);
      if (s) {
        const n = parseInt(s[1], 10);
        if (n > 0 && n <= 100) { byNum[n] = el; }
      }
    });
    if (Object.keys(byNum).length < 2) continue;
    if (cur < 1) {
      tds.forEach((td) => {
        if (td.querySelector("a, button")) return;
        const w = (td.textContent || "").replace(/\\s+/g, "").match(/^(\\d{1,2})$/);
        if (w) { cur = parseInt(w[1], 10); }
      });
    }
    if (cur < 1) cur = 1;
    const w = String(cur + 1);
    if (byNum[cur + 1]) {
      byNum[cur + 1].click();
      return "n";
    }
  }
  /* Vue / SPA：div.pages 內常為 <ol class="page">，下一頁是 SVG 箭頭，title=往後一頁，無 ">" 字元 */
  for (const box of M.querySelectorAll("div.pages, .content .pages, #Item .pages, main .pages, ol.page")) {
    const t = (box.textContent || "").replace(/\\s+/g, " ").trim();
    if (t.length > 500) continue;
    if (!/[12]/.test(t)) continue;
    const nextA = box.querySelector('a[title="往後一頁"]') || box.querySelector('a[title="下一頁"]') || box.querySelector("li.PageRight a, .PageRight a");
    if (nextA && nextA.offsetParent) {
      nextA.click();
      return "vnextT";
    }
    for (const el of box.querySelectorAll("a, button, [role=button]")) {
      if (!el.offsetParent) continue;
      const s = (el.textContent || "").replace(/\\s+/g, "").trim();
      if (s === ">" || s === "›" || s === "»" || s === "＞" || s === "》") {
        el.click();
        return "vnext";
      }
    }
  }
  for (const box of M.querySelectorAll("div.pages, .content .pages, #Item .pages, main .pages, ol.page")) {
    const t = (box.textContent || "").replace(/\\s+/g, " ").trim();
    if (t.length > 500) continue;
    const links = Array.from(box.querySelectorAll("a, button, [role=button]"));
    if (links.length < 2) continue;
    const byNum = {};
    links.forEach((el) => {
      if (!el.offsetParent) return;
      const m = (el.textContent || "").replace(/\\s+/g, "").match(/^(\\d{1,2})$/);
      if (m) {
        const n = parseInt(m[1], 10);
        if (n > 0 && n < 200) { byNum[n] = el; }
      }
    });
    if (Object.keys(byNum).length < 2) continue;
    let cur = 0;
    links.forEach((el) => {
      const c = (el.className || "");
      if (/\b(active|current|selected|PageFocus|on|sel|router-link-active)\b/i.test(c) || (el.getAttribute("aria-current") || "") === "page") {
        const m = (el.textContent || "").replace(/\\s+/g, "").match(/^(\\d{1,2})/);
        if (m) { cur = parseInt(m[1], 10); }
      }
    });
    if (cur < 1) cur = 1;
    if (byNum[cur + 1]) {
      byNum[cur + 1].click();
      return "vnum";
    }
  }
  return "";
}
"""

async def _try_advance_pagination_aspnetish_table(page, st0) -> bool:
    """同 URL 下：table tr 分頁、**Vue/SPA 的 div.pages**（如 #Item 旁 `ol`+`.pages`）等，箭頭不在 class=Pager 時仍翻頁。"""
    try:
        kind=await page.evaluate(_PAGER_CLICK_JS.strip())
    except Exception:
        return False
    if not kind:
        return False
    await page.wait_for_timeout(2200)
    try:await page.wait_for_load_state("domcontentloaded",timeout=15000)
    except Exception:pass
    if await _wait_pagination_state_change(page, st0):
        return True
    return await _pagination_state_key(page) != st0

async def gather_extracted_links_for_visible_page(
    page,browser_context,final_url,target_domain,file_exts,ext_http_seen:Optional[set]=None
):
    html_content=await page.content()
    if ext_http_seen is not None:
        ext_http_seen|=extract_external_http_links(html_content,final_url,target_domain)
        ext_http_seen|=await _external_http_hrefs_from_item_live_dom(
            page, final_url, target_domain
        )
    links=extract_same_domain_links(html_content,final_url,target_domain,file_exts)
    links|=await discover_download_links_from_file_buttons(page,browser_context,target_domain,file_exts)
    links|=await discover_document_labeled_anchor_urls(page,browser_context,target_domain,file_exts)
    return links

async def _try_advance_pagination_vue_div_pages_playwright(page, st0) -> bool:
    """Vue 的 `div.pages` / `ol.page`：以 Playwright 實體點擊，避免在 evaluate 用 el.click 未觸發 @click 導致列表未換頁。
    一律**優先**在 `#Item` 可見的 `.pages` 內點，避免多個全站/隱藏區塊的 `div.pages` 擋住『下一頁』點到錯節點。
    """
    try:
        ev_ok = await page.evaluate(
            r"""() => {
          const r = document.querySelector("#Item");
          if (!r) return false;
          const a = r.querySelector('a[title="往後一頁"]') || r.querySelector("li.PageRight a");
          if (!a) return false;
          const li = a.closest("li");
          if (li) {
            const cls = (li.className || "").toLowerCase();
            if (cls.includes("disabled") || li.getAttribute("aria-disabled") === "true") return false;
          }
          try { a.scrollIntoView({ block: "center", inline: "nearest" }); } catch (e) {}
          a.click();
          return true;
        }"""
        )
    except Exception:
        ev_ok = False
    if ev_ok:
        try:
            await page.wait_for_load_state("networkidle", timeout=12000)
        except Exception:
            pass
        if await _wait_pagination_state_change(page, st0):
            return True
    cands = [
        "#Item .pages a[title='往後一頁']",
        "#Item div.pages a[title='往後一頁']",
        "#Item a[title='往後一頁']",
        "main .pages a[title='往後一頁']",
        "main div.pages a[title='往後一頁']",
        "div.pages a[title='往後一頁']",
        "#Item ol.page li.PageRight a, #Item li.PageRight a",
        "ol.page li.PageRight a, div.pages li.PageRight a",
    ]
    for sel in cands:
        loc=page.locator(sel)
        try:nc=await loc.count()
        except Exception:nc=0
        if nc<1:continue
        e=loc.first
        try:await e.scroll_into_view_if_needed()
        except Exception:pass
        try:await e.click(timeout=20000, force=True)
        except Exception:continue
        try:
            await page.wait_for_load_state("networkidle", timeout=12000)
        except Exception:
            pass
        if await _wait_pagination_state_change(page, st0):
            return True
    for box_sel in ("#Item div.pages, #Item .pages", "main div.pages", "div.pages"):
        box = page.locator(box_sel)
        try:nb = await box.count()
        except Exception:nb=0
        for bi in range(min(nb, 4)):
            b0 = box.nth(bi)
            try:cur=b0.locator("a.PageFocus, a[class*='PageFocus' i]")
            except Exception:continue
            try:nc2=await cur.count()
            except Exception:nc2=0
            if nc2<1:continue
            try:t=(await cur.first.inner_text()or"").strip()
            except Exception:continue
            if not t.isdigit():
                continue
            n1=int(t)+1
            if n1<2 or n1>500:continue
            nxt = b0.get_by_text(str(n1), exact=True)
            try:nn=await nxt.count()
            except Exception:nn=0
            if nn<1:continue
            el=nxt.first
            try:await el.scroll_into_view_if_needed()
            except Exception:pass
            try:await el.click(timeout=20000, force=True)
            except Exception:continue
            try:
                await page.wait_for_load_state("networkidle", timeout=12000)
            except Exception:
                pass
            if await _wait_pagination_state_change(page, st0):
                return True
    return False

async def _try_advance_pagination_numeric_buttons(page, st0) -> bool:
    """Vue／Tailwind 風格的『純數字 <button> 分頁』，無 .pagination class 提示。

    例：/utmap/stations 用 `<button>1</button>...<button>5</button>` 排列頁碼，現行 fallback
    只篩 `<a>` 數字節點 → 抓不到。本函式以「同列、相鄰、連續整數、有高亮樣式」為共同特徵
    辨識分頁列，並前進到當前頁碼+1。為降低誤觸（如評分 1～5 之類數字按鈕），要求：
      ① 同一 Y 列至少 3 顆數字按鈕
      ② 數字 X 座標由小至大且皆 >0（同列）
      ③ 必須能找到目前 active 按鈕（class 含 active/selected 或 bg-* 高亮、aria-current=page）
    """
    info = await page.evaluate("""() => {
      const all = Array.from(document.querySelectorAll('button'));
      const nums = [];
      for (const b of all) {
        const t = (b.innerText || '').trim();
        if (!/^\\d+$/.test(t)) continue;
        const r = b.getBoundingClientRect();
        if (r.width <= 0 || r.height <= 0) continue;
        const cls = (b.className || '').toLowerCase();
        const aria = (b.getAttribute('aria-current') || '').toLowerCase();
        // 高亮樣式：原生 active/selected/current，或常見 Tailwind 主色配白字（bg-* + text-white）
        const tailwindActive = /text-white/.test(cls) && /bg-(?:slate|gray|primary|blue|indigo|cyan|sky|teal|emerald|red|rose|orange|amber|yellow|green|purple|pink)/.test(cls);
        const active = aria === 'page' || /\\b(active|selected|current)\\b/.test(cls) || tailwindActive;
        nums.push({n: parseInt(t,10), x: r.x, y: r.y, active});
      }
      if (nums.length < 3) return null;
      // 以 Y 座標 ±10px 分群，挑最大群（最像分頁列）
      const groups = [];
      for (const it of nums) {
        let g = groups.find(g => Math.abs(g.y - it.y) < 10);
        if (!g) { g = {y: it.y, items: []}; groups.push(g); }
        g.items.push(it);
      }
      groups.sort((a,b) => b.items.length - a.items.length);
      const best = groups[0];
      if (!best || best.items.length < 3) return null;
      best.items.sort((a,b) => a.x - b.x);
      // 連續性：相鄰 items 之 n 差距須 ≤ 1（允許起始為 1 或更高）
      const ns = best.items.map(it => it.n);
      for (let i = 1; i < ns.length; i++) {
        if (ns[i] - ns[i-1] !== 1) return null;
      }
      const cur = best.items.find(it => it.active);
      if (!cur) return null;
      return { curN: cur.n, ns: ns };
    }""")
    if not info:
        return False
    cur=info.get("curN")
    ns=info.get("ns") or []
    if not isinstance(cur,int) or not isinstance(ns,list) or len(ns)<3:
        return False
    target=cur+1 if (cur+1) in ns else None
    if target is None:
        return False
    btn=page.get_by_role("button",name=re.compile(rf"^\s*{target}\s*$"))
    try:
        bn=await btn.count()
    except Exception:
        bn=0
    for i in range(min(bn,10)):
        el=btn.nth(i)
        if not await _try_advance_pagination__clickable_guess(el):
            continue
        try:
            await el.scroll_into_view_if_needed()
            await el.click(timeout=15000)
        except Exception:
            continue
        await page.wait_for_timeout(2200)
        try:
            await page.wait_for_load_state("domcontentloaded",timeout=8000)
        except Exception:
            pass
        if await _pagination_state_key(page) != st0:
            return True
    return False


async def try_advance_pagination(page)->bool:
    """點擊下一頁／下頁／pager 箭頭；若列表內文有變化則視為成功。"""
    st0=await _pagination_state_key(page)
    if await _try_advance_pagination_vue_div_pages_playwright(page, st0):
        return True
    if await _try_advance_pagination_numeric_buttons(page, st0):
        return True

    async def _clickable(el)->bool:
        try:
            if not await el.is_visible():
                return False
            if await el.is_disabled():
                return False
        except Exception:
            return False
        try:
            bad=el.locator("xpath=ancestor::*[@aria-disabled='true' or contains(@class,'disabled')][1]")
            if await bad.count()>0:
                return False
        except Exception:
            pass
        try:
            cls=(await el.get_attribute("class")or"").lower()
            if "disabled" in cls:
                return False
        except Exception:
            pass
        return True

    groups=[]
    for pat in(r"^\s*下一頁\s*$",r"^\s*下頁\s*$",r"^\s*Next\s*$"):
        try:
            groups.append(page.get_by_role("link",name=re.compile(pat,re.I)))
        except Exception:
            pass
    try:
        groups.append(page.get_by_role("button",name=re.compile(r"下一頁|下頁|Next",re.I)))
    except Exception:
        pass
    groups.append(page.locator("a").filter(has_text=re.compile(r"^\s*下一頁\s*$|^\s*下頁\s*$")))
    groups.append(
        page.locator('[class*="Pager" i],[class*="pager" i],[class*="pagination" i],[id*="Pager" i],[class*="datagrid" i]')
        .locator("a,button,input[type='submit'],input[type='button'],input[type='image']")
        .filter(has_text=re.compile(r"^\s*>\s*$|^\s*›\s*$|^\s*»\s*$|下一頁|下頁|Next",re.I))
    )

    for grp in groups:
        try:
            n=await grp.count()
        except Exception:
            continue
        for i in range(min(n,24)):
            el=grp.nth(i)
            if not await _clickable(el):
                continue
            try:
                await el.scroll_into_view_if_needed()
                await el.click(timeout=15000)
            except Exception:
                continue
            await page.wait_for_timeout(2200)
            try:
                await page.wait_for_load_state("domcontentloaded",timeout=8000)
            except Exception:
                pass
            fp1=await _pagination_state_key(page)
            if fp1 != st0:
                return True
    if await _try_advance_pagination_fallback(page, st0):
        return True
    return await _try_advance_pagination_aspnetish_table(page, st0)

async def _try_advance_pagination_fallback(page, st0) -> bool:
    """分頁後援：在常見 Pager／pagination 區內尋找「>›»」、英語 Next、或遞增數字。"""
    glist=[
        page.locator("main, [role=main], #content, article, body")
        .locator("a, button, [role=button], input[type=button], input[type=image]")
        .filter(has_text=re.compile(r"^\s*>\s*$|^\s*›\s*$|^\s*»\s*$|^\s*》\s*$|^\s*＞\s*$")),
        page.locator('nav[aria-label*="分頁" i], nav[aria-label*="pagination" i], [class*="Pager" i], [class*="paging" i], [class*="pagination" i], .Pager, .pagination, .paging, .aspNetPager, .GridPager')
        .locator("a,button,[role=button],input[type=image],input[type=button]"),
    ]
    for g in glist:
        for pat in (r"^\s*>\s*$",r"^›$",r"^»$",r"^\s*▶\s*$",r"^Next$",r"^next$"):
            try:
                c=g.filter(has_text=re.compile(pat))
                m=await c.count()
            except Exception:
                continue
            for i in range(min(m,12)):
                el=c.nth(i)
                if not await _try_advance_pagination__clickable_guess(el):
                    continue
                try:
                    await el.scroll_into_view_if_needed()
                    await el.click(timeout=15000)
                except Exception:
                    continue
                await page.wait_for_timeout(2200)
                try:await page.wait_for_load_state("domcontentloaded",timeout=8000)
                except Exception:pass
                if await _pagination_state_key(page) != st0:
                    return True
    try:
        cands=page.locator(
            '[class*="Pager" i] a, [class*="paging" i] a, .GridPager a, [class*=pagination] a, nav a'
        ).filter(has_text=re.compile(r"^(?:[2-9]|\d{2,})$",re.A))
        n2=await cands.count()
    except Exception:
        n2=0
    for i in range(min(n2,6)):
        el=cands.nth(i)
        if not await _try_advance_pagination__clickable_guess(el):
            continue
        try:txt=((await el.inner_text())or"").strip()
        except Exception:continue
        if not txt.isdigit()or int(txt)<=1:continue
        try:
            await el.scroll_into_view_if_needed()
            await el.click(timeout=15000)
        except Exception:
            continue
        await page.wait_for_timeout(2200)
        try:await page.wait_for_load_state("domcontentloaded",timeout=8000)
        except Exception:pass
        if await _pagination_state_key(page) != st0:
            return True
    return False

async def _try_advance_pagination__clickable_guess(el):
    try:
        if not await el.is_visible():return False
        if await el.is_disabled():return False
    except Exception:return False
    try:
        c=(await el.get_attribute("class")or"").lower()
        if"disabled"in c:return False
    except Exception:pass
    return True
async def _collect_links_pagination_on_current_view(
    page,browser_context,final_url,target_domain,file_exts,ext_http_seen:Optional[set]
):
    """單一『視圖』下（一個分頁狀態）：第一頁＋盡量往後翻，合併連結。"""
    out=await gather_extracted_links_for_visible_page(
        page,browser_context,final_url,target_domain,file_exts,ext_http_seen=ext_http_seen
    )
    seen_fp = {hash(await _pagination_state_key(page))}
    for _ in range(PAGINATION_MAX_STEPS):
        if not await try_advance_pagination(page):
            break
        h = hash(await _pagination_state_key(page))
        if h in seen_fp:
            break
        seen_fp.add(h)
        out|=await gather_extracted_links_for_visible_page(
            page,browser_context,final_url,target_domain,file_exts,ext_http_seen=ext_http_seen
        )
    return out

async def _tab_locator_aria(page):
    tmain=page.locator("main, [role=main], .main-content, #content, .container, article")
    t=tmain.get_by_role("tab")
    try:tc=await t.count()
    except Exception:tc=0
    if 2 <= tc <= 20:
        return t
    t2=page.get_by_role("tab")
    try:tc2=await t2.count()
    except Exception:tc2=0
    if 2 <= tc2 <= 20:
        return t2
    try:
        t3=page.get_by_role("tab").filter(has_text=re.compile(r"國內|國外|資收|價|網站",re.I))
        tc3=await t3.count()
        if 2 <= tc3 <= 20:
            return t3
    except Exception:
        pass
    return None

async def _recycle_eri_tab_anchor_ids(page) -> List[str]:
    """回傳 Tab 列上實際存在的錨點 id（Domestic→Foreign→ValueReference 順序）。
    不以 is_visible 過濾：在 Streamlit／不同視窗下**誤判不可見**會整段略過 `a#Foreign`，變成只掃 國內→資收 而漏掉 國外 全部分頁。"""
    out=[]
    bar=page.locator("ul#Button, ul.website-title").first
    try:
        if await bar.count()<1:
            return []
    except Exception:
        return []
    for _id in("Domestic","Foreign","ValueReference"):
        loc=bar.locator(f"a#{_id}")
        try:nc=await loc.count()
        except Exception:nc=0
        if nc>=1:
            out.append(_id)
    return out if len(out)>=2 else[]

async def _click_recycle_tab_by_anchor_id(page, anchor_id:str)->None:
    """Vue 分頁常為 `href=javascript:void(0)`，在部分環境需於瀏覽器內直接 `a.click()` 才會觸發 @click。"""
    ok=False
    try:
        ok=bool(
            await page.evaluate(
                """(id) => {
          const root = document.querySelector("ul#Button, ul.website-title");
          if (!root) return false;
          const a = root.querySelector("a#" + id) || document.querySelector("ul#Button a#" + id + ", ul.website-title a#" + id);
          if (!a) return false;
          try { a.scrollIntoView({block:"center", inline:"nearest"}); } catch (e) {}
          a.click();
          return true;
        }""",
                anchor_id,
            )
        )
    except Exception:
        ok=False
    if ok:
        return
    loc=page.locator(f"ul#Button a#{anchor_id}, ul.website-title a#{anchor_id}").first
    try:
        await loc.scroll_into_view_if_needed()
    except Exception:pass
    try:
        await loc.click(timeout=20000, force=True)
    except Exception:
        try:
            await loc.evaluate("e => e && e.click && e.click()")
        except Exception:pass

async def _zh_gov_service_tab_clicks(page):
    """政府／ASP.NET／**Vue** 常見的『國內網站、國外網站、資收物…』分條（如 `ul#Button.website-title`），**無** role=tab 時以固定順序點擊掃描。"""
    root=page.locator("main, [role=main], .main-content, #content, .container, article, body")
    bar=page.locator("ul#Button, ul.website-title")
    out=[]
    for want in("國內網站","國外網站"):
        loc=bar.get_by_text(want, exact=True)
        try:
            n=await loc.count()
        except Exception:
            n=0
        if n<1:
            loc=root.get_by_text(want, exact=True)
            try:
                n=await loc.count()
            except Exception:
                n=0
        if n<1:return[]
        picked=None
        for j in range(min(n,6)):
            el=loc.nth(j)
            try:
                if not await el.is_visible():
                    continue
            except Exception:
                continue
            p=(await el.text_content()or"").replace("\n","").replace("\r","").strip()
            if p!=want:
                continue
            picked=el;break
        if picked is None:return[]
        out.append(picked)
    t3=bar.get_by_text("資收物回收價參考網站", exact=True)
    try:tn3=await t3.count()
    except Exception:tn3=0
    use=t3
    if tn3<1:
        use=root.get_by_text("資收物回收價參考網站", exact=True)
        try:tn3=await use.count()
        except Exception:tn3=0
    if tn3<1:
        use=root.get_by_text(re.compile(r"資收物.{0,24}回收(價|參|考)",re.S|re.I))
    try:tn=await use.count()
    except Exception:tn=0
    if tn<1:
        use=root.get_by_text(re.compile(r"資收物.{0,12}回收",re.S|re.I))
        try:tn=await use.count()
        except Exception:tn=0
    if tn>=1:
        for j in range(min(tn,8)):
            el=use.nth(j)
            try:
                if not await el.is_visible():
                    continue
            except Exception:
                continue
            p=(await el.text_content()or"").replace("\n","").replace("\r","").strip()
            if len(p)<=64 and"資收"in p and ("回收"in p or"價"in p or"參"in p):
                out.append(el)
                break
    return out if len(out)>=2 else[]

async def _tab_locator_candidates(page):
    return await _tab_locator_aria(page)

async def _wait_item_tab_list_ready(page):
    """`ul#Button` 等切到國內/國外後，等 #Item 出現連結，避免在 Vue 還在換內文時就照 HTML 採集而漏站。"""
    t0=time.time()
    while (time.time()-t0) < 18.0:
        n=0
        try:
            n=await page.locator("#Item a[href]").count()
        except Exception:
            n=0
        if n>0:
            try:await page.locator("#Item").first.scroll_into_view_if_needed()
            except Exception:pass
            await page.wait_for_timeout(500)
            return
        try:
            await page.wait_for_timeout(220)
        except Exception:return
    try:await page.wait_for_timeout(1000)
    except Exception:pass

async def collect_links_with_pagination(
    page,browser_context,final_url,target_domain,file_exts,ext_http_seen:Optional[set]=None
):
    """第一頁＋盡量翻頁；若頁面有 2+ 個 role=tab 或關鍵字 Tab，會依序點每個分頁各掃一輪內層分頁。

    多數『相關連結』頁為 Tab 內再套 GridView 分頁；單一 HTML 內就含全部 href 者無須點 Tab 也會從
    `extract_same_domain_links` 擷到，本段補上『必須切 Tab 或 PostBack 才出現的列表』遺漏。
    """
    out=set()
    tabs=await _tab_locator_candidates(page)
    tcount=0
    if tabs is not None:
        try:tcount=await tabs.count()
        except Exception:tcount=0
    # 有 `ul#Button` 國內／國外等時**優先依序掃**；若同頁還有其它 role=tab 會誤觸 ARIA 路徑導致漏切 Tab。
    tab_ids=await _recycle_eri_tab_anchor_ids(page)
    if len(tab_ids)>=2:
        for aid in tab_ids:
            await _click_recycle_tab_by_anchor_id(page, aid)
            try:await page.wait_for_load_state("domcontentloaded",timeout=20000)
            except Exception:pass
            try:await page.wait_for_load_state("networkidle",timeout=10000)
            except Exception:pass
            await page.wait_for_timeout(2000)
            await _wait_item_tab_list_ready(page)
            await page.wait_for_timeout(2000)
            out|=await _collect_links_pagination_on_current_view(
                page,browser_context,final_url,target_domain,file_exts,ext_http_seen=ext_http_seen
            )
        return out
    zh=await _zh_gov_service_tab_clicks(page)
    if len(zh) >= 2:
        for el in zh:
            try:
                if not await el.is_visible():
                    try:await el.scroll_into_view_if_needed()
                    except Exception:pass
                await el.click(timeout=20000, force=True)
            except Exception:
                try:
                    await el.evaluate("e => e && e.click && e.click()")
                except Exception:
                    continue
            try:await page.wait_for_load_state("domcontentloaded",timeout=20000)
            except Exception:pass
            try:await page.wait_for_load_state("networkidle",timeout=10000)
            except Exception:pass
            await page.wait_for_timeout(2000)
            await _wait_item_tab_list_ready(page)
            await page.wait_for_timeout(2000)
            out|=await _collect_links_pagination_on_current_view(
                page,browser_context,final_url,target_domain,file_exts,ext_http_seen=ext_http_seen
            )
        return out
    if tcount < 2:
        return await _collect_links_pagination_on_current_view(
            page,browser_context,final_url,target_domain,file_exts,ext_http_seen=ext_http_seen
        )
    n=tcount
    for i in range(min(n,20)):
        try:
            el=tabs.nth(i)
            if not await el.is_enabled():continue
        except Exception:continue
        try:
            if not await el.is_visible():
                try:await el.scroll_into_view_if_needed()
                except Exception:pass
            await el.click(timeout=15000)
        except Exception:
            continue
        try:
            await page.wait_for_load_state("domcontentloaded",timeout=15000)
        except Exception:pass
        await page.wait_for_timeout(2000)
        out|=await _collect_links_pagination_on_current_view(
            page,browser_context,final_url,target_domain,file_exts,ext_http_seen=ext_http_seen
        )
    return out

def _normalize_charset_token(raw:str)->str:
    if not raw:return""
    s=raw.strip().strip('"').strip("'").lower().replace(" ","")
    if s=="utf8":return"utf-8"
    return s

def _is_charset_check_html_document(response,html_content:str)->bool:
    """僅對以文字編碼撰寫之 HTML／XHTML 網頁做 UTF-8 檢核；圖影音字型、PDF 等非網頁本文略過。"""
    raw=(response.headers.get("content-type")or"")if response else""
    ct0=raw.split(";",1)[0].strip().lower()
    if ct0.startswith(("image/","video/","audio/","font/","model/")):
        return False
    if ct0=="application/pdf":
        return False
    if ct0 in("text/html","application/xhtml+xml"):
        return True
    head=(html_content or"")[:8192].lstrip().lower()
    if"<!doctype html"in head or re.search(r"<\s*html[\s>]",head):
        return True
    return False

async def check_page_utf8(response,soup)->bool:
    """語系編碼：優先採信 HTML 內之 charset／http-equiv 宣告（與「檢視原始碼」一致），再比對 HTTP Content-Type，最後才以本文嚴格 UTF-8 解碼判定。"""
    meta=soup.find("meta",attrs={"charset":True})
    if meta:
        ch=_normalize_charset_token(meta.get("charset")or"")
        if ch=="utf-8":return True
        if ch:return False
    for meta in soup.find_all("meta"):
        he=(meta.get("http-equiv")or"")
        if he.lower()!="content-type":continue
        content=(meta.get("content")or"").lower()
        m2=re.search(r"charset\s*=\s*([^;\s]+)",content)
        if not m2:continue
        ch=_normalize_charset_token(m2.group(1))
        if ch=="utf-8":return True
        return False
    ct=(response.headers.get("content-type")or"").lower()if response else""
    m=re.search(r"charset\s*=\s*([^;\s]+)",ct)
    if m:
        ch=_normalize_charset_token(m.group(1))
        if ch=="utf-8":return True
        return False
    b=b""
    if response:
        try:
            b=await response.body()
        except Exception:
            b=b""
    if b.startswith(b"\xef\xbb\xbf"):
        b=b[3:]
    try:
        b.decode("utf-8")
        return True
    except Exception:
        return False

def check_flash_or_legacy_ria(html:str,soup)->bool:
    """回傳 True 表示發現 Flash 或常見 RIA／外掛內容（不符合開放標準／應避免之技術）。"""
    h=(html or"").lower()
    if re.search(r"\.swf(?:\?|#|\"|'|>|\\s|$)",h):return True
    if"application/x-shockwave-flash"in h or"application/futuresplash"in h:return True
    if"d27cdb6e-ae6d-11cf-96b8-444553540000"in h:return True
    if"<applet"in h or"application/x-java-applet"in h or"application/x-java-vm"in h:return True
    if"application/x-java-archive"in h:return True
    if"application/x-silverlight"in h or"application/x-silverlight-app"in h:return True
    if".xap"in h and("silverlight"in h or"data:application/x-silverlight"in h):return True
    for tag in soup.find_all(["embed","object"]):
        typ=(tag.get("type")or"").lower()
        if any(x in typ for x in["shockwave-flash","x-silverlight","java-applet","java-vm"]):return True
        for attr in("data","src","classid","codebase"):
            val=(tag.get(attr)or"").lower()
            if".swf"in val:return True
            if".xap"in val:return True
            if"java"in val and("applet"in val or".class"in val):return True
    return False

# ==========================================
# PageSpeed API測速邏輯
# ==========================================
# Google PageSpeed Insights v5 API；金鑰直接內嵌（依使用者設定）。
PAGESPEED_API_BASE = "https://www.googleapis.com/pagespeedonline/v5/runPagespeed"
PAGESPEED_API_KEY = "AIzaSyCLV-l5EioNaDOm7ytU399DByBFiJ_jjNs"
# Lighthouse 偶發暫時性失敗（500 + "Something went wrong" / "FAILED_DOCUMENT_REQUEST" / 502/503/504）：
# 屬 Google 後端非穩定情境，重試常能恢復；針對下列 HTTP 與訊息特徵自動重試。
_PSI_RETRYABLE_HTTP = (429, 500, 502, 503, 504)
_PSI_RETRYABLE_MSG_FRAGMENTS = (
    "lighthouse returned error",
    "something went wrong",
    "failed_document_request",
    "no_fcp",
    "no_lcp",
    "internal error",
    "deadline",
    "rate limit",
)

def _pagespeed_call_once(url:str, strategy:str, api_key:str):
    """單次 PSI 呼叫；回傳 (status_code, json_or_None, friendly_error_msg)。"""
    qs = f"url={requests.utils.quote(url, safe='')}&strategy={strategy}"
    if api_key:
        qs += f"&key={api_key}"
    full = f"{PAGESPEED_API_BASE}?{qs}"
    try:
        r = requests.get(full, timeout=120)
    except Exception as e:
        return 0, None, f"PSI 連線失敗：{type(e).__name__}：{e}"
    try:
        j = r.json()
    except Exception:
        j = None
    if r.status_code == 200 and j:
        return r.status_code, j, ""
    msg = ""
    if isinstance(j, dict):
        err = j.get("error") or {}
        em = (err.get("message") or "").strip()
        reason = ""
        for d in (err.get("details") or []):
            rr = (d.get("reason") or "")
            if rr:
                reason = rr; break
        if em:
            msg = f"PSI {r.status_code}：{em}"
            if reason:
                msg += f"（{reason}）"
    if not msg:
        msg = f"PSI HTTP {r.status_code}"
    return r.status_code, j, msg

def _pagespeed_is_retryable(status:int, msg:str)->bool:
    if status in _PSI_RETRYABLE_HTTP:
        return True
    low = (msg or "").lower()
    return any(frag in low for frag in _PSI_RETRYABLE_MSG_FRAGMENTS)

def _pagespeed_call(url:str, strategy:str, api_key:str, max_attempts:int=3):
    """PSI 呼叫含暫時性失敗自動重試（Lighthouse 偶發 500、502/503/504、限流等）。"""
    last = (0, None, "")
    for attempt in range(1, max_attempts + 1):
        sc, j, msg = _pagespeed_call_once(url, strategy, api_key)
        if sc == 200 and j:
            return sc, j, msg
        last = (sc, j, msg)
        if attempt >= max_attempts or not _pagespeed_is_retryable(sc, msg):
            break
        # 退避：依次 5s / 10s，第三次直接結束。
        time.sleep(5 * attempt)
    return last

def get_pagespeed_score(url):
    """回傳 (success, avg, mobile, desktop, msg)。msg 為失敗時的細節訊息（含 API 端錯誤理由）。"""
    api_key = PAGESPEED_API_KEY
    sc_m, j_m, err_m = _pagespeed_call(url, "mobile", api_key)
    sc_d, j_d, err_d = _pagespeed_call(url, "desktop", api_key)
    def _score(j):
        try:
            return float(j["lighthouseResult"]["categories"]["performance"]["score"]) * 100.0
        except Exception:
            return 0.0
    # 兩端皆成功：取平均
    if sc_m == 200 and sc_d == 200 and j_m and j_d:
        m = _score(j_m); d = _score(j_d)
        return True, (m + d) / 2.0, m, d, "測試成功"
    # 單端成功：仍回傳該端分數，平均以該端值代之；訊息註明對側失敗。
    if sc_m == 200 and j_m:
        m = _score(j_m); d = 0.0
        return True, m, m, d, f"電腦端失敗（{err_d or 'PSI 暫時無法分析'}）；以行動端分數為準"
    if sc_d == 200 and j_d:
        d = _score(j_d); m = 0.0
        return True, d, m, d, f"行動端失敗（{err_m or 'PSI 暫時無法分析'}）；以電腦端分數為準"
    parts = []
    if err_m:
        parts.append(f"行動：{err_m}")
    if err_d:
        parts.append(f"電腦：{err_d}")
    msg = "；".join(parts) or "API 錯誤"
    # 屬 Google 端的暫時性失敗 → 提示稍後再試
    if any(frag in msg.lower() for frag in _PSI_RETRYABLE_MSG_FRAGMENTS) or "PSI 500" in msg or "PSI 502" in msg or "PSI 503" in msg or "PSI 504" in msg:
        msg += "（Google PSI 後端暫時性錯誤，建議稍後再試）"
    return False, 0.0, 0.0, 0.0, msg

# ==========================================
# 核心引擎：精準校驗版Playwright
# ==========================================
async def check_single_page(
    browser_context,url,target_domain,exclusions,scope_root_url=None,referer_url=None,
    external_probe_cache=None,page_exceptions:Optional[List[str]]=None,
    external_probe_sem:Optional[asyncio.Semaphore]=None,
):
    page_errors,detail_found=[],{k:False for k in ["favicon","privacy","security","phone","address","open_data","accessibility","nav","lang_ver","search","opinion","rwd","stats","date_info"]}
    extracted_links,html_content,final_url=set(),"",url # 預初始化避免UnboundLocalError
    external_failed,external_probed=[],[]
    probe_cache=external_probe_cache if external_probe_cache is not None else {}
    
    # 網址標準化處理
    safe_url=make_safe_url(url)
    file_exts=list(_SCAN_FILE_EXTENSIONS)
    is_requested_as_file=any(url.lower().endswith(ext) for ext in file_exts)

    page=await browser_context.new_page()
    try:
        ref=referer_url or _site_root_url(safe_url)
        await page.set_extra_http_headers(_extra_browser_headers(ref))
        # PDF 等仍以 commit 快取首包；一般 HTML 用 domcontentloaded 取得較穩定的 navigation response，減少誤判 5.有效連結
        _nav_wait="commit"if is_requested_as_file else"domcontentloaded"
        try:
            response=await page.goto(safe_url,wait_until=_nav_wait,timeout=45000,referer=ref)
        except Exception as nav_err:
            # Playwright 在 Content-Disposition: attachment 時會丟「Download is starting」例外，
            # 對 .pdf／.docx 等檔案連結而言**屬下載成功**，連結是有效的；不要捕到後置回失敗。
            if is_requested_as_file and "download" in (str(nav_err) or "").lower():
                await page.close()
                return page_errors,"",detail_found,set(),url,[],[]
            raise
        nav_recovered=False
        content_type=""
        st=getattr(response,"status",None)if response else None
        ok_first=bool(response and st is not None and st<400)

        if ok_first:
            final_url=page.url
            content_type=(response.headers.get("content-type")or"").lower()
            if is_requested_as_file:
                final_is_file=any(final_url.lower().endswith(ext)for ext in file_exts)
                bad_file=not final_is_file and("text/html"in content_type)
                if url.lower().endswith(".pdf")and"text/html"in content_type:
                    bad_file=True
                if bad_file:
                    await page.close()
                    return["5.有效連結"],"",detail_found,set(),url,[],[]
        else:
            # 登入／挑戰頁等：首包可能無 response 或 status≥400，但主框架已導向同站 HTML
            if is_requested_as_file:
                await page.close()
                return["5.有效連結"],"",detail_found,set(),url,[],[]
            # 真正破損的 4xx/5xx（如 IIS `/EN/.div_xxx` 走 404+big5 預設頁）必須列入 5.有效連結；
            # 否則程式會「復原」進入 charset 檢查，把錯誤碼網址誤判成「3.語系編碼」。
            # 僅 401/403/407/429（需登入／節流等）與 503（暫時不可用）視為可恢復。
            if st is not None and not _http_status_reachable_for_external_check(st):
                await page.close()
                return["5.有效連結"],"",detail_found,set(),url,[],[]
            try:
                await page.wait_for_load_state("domcontentloaded",timeout=20000)
            except Exception:
                pass
            final_url=(page.url or"").strip()
            try:
                dom_ok=(
                    bool(final_url)
                    and urlparse(final_url).netloc.lower()==target_domain.lower()
                    and not final_url.lower().startswith("about:")
                )
            except Exception:
                dom_ok=False
            if not dom_ok:
                await page.close()
                return["5.有效連結"],"",detail_found,set(),url,[],[]
            nav_recovered=True
            content_type="text/html"

        # 純圖影音字型：不跑 HTML 指標與擷取連結（含 favicon.png?v= 類網址）
        if not is_requested_as_file and response and not nav_recovered:
            _ct_img=(content_type or"").lower().split(";")[0].strip()
            if _ct_img.startswith(("image/","audio/","video/","font/")):
                await page.close()
                return page_errors,"",detail_found,set(),url,[],[]

        # ── 非副檔名 URL 但回傳文件型 MIME（.ashx / DownFile.aspx 等下載處理器） ──
        # 立即返回，避免等待 domcontentloaded 逾時或解析 binary 資料（復原導向後勿信首包 headers）
        if not is_requested_as_file and not nav_recovered and response:
            _ct0=(content_type or"").lower().split(";")[0].strip()
            _is_binary_doc=any(_ct0.startswith(p)for p in _BINARY_DOC_MIME_PREFIXES)
            if not _is_binary_doc and _ct0=="application/octet-stream":
                cd_hdr=(response.headers.get("content-disposition")or"").lower()
                _is_binary_doc="attachment"in cd_hdr
            if _is_binary_doc:
                if any(p in _ct0 for p in _OFFICE_MIME_PARTS):
                    page_errors.append("10.文件格式")
                await page.close()
                return page_errors,"",detail_found,set(),url,[],[]

        # 指標10：文件格式檢核
        if any(url.lower().endswith(ext)for ext in[".doc",".docx",".xls",".xlsx",".ppt",".pptx",".rar",".odt",".ods",".odp"]):
            page_errors.append("10.文件格式")

        # 若確診為檔案，直接返回，不再進行JS渲染
        if is_requested_as_file:
            await page.close()
            return page_errors,"",detail_found,set(),url,[],[]

        # 正常網頁解析
        try:
            await page.wait_for_load_state("domcontentloaded",timeout=20000)
        except Exception:
            pass
        try:
            await page.wait_for_load_state("networkidle",timeout=1500)
        except Exception:
            pass
        await page.wait_for_timeout(1200)
        try:
            if _UTMAP_STATION_DETAIL_PATH_RE.search(urlparse(safe_url).path or ""):
                try:
                    await page.wait_for_load_state("networkidle",timeout=5000)
                except Exception:
                    pass
                await page.wait_for_timeout(800)
        except Exception:
            pass
        final_url=page.url
        if urlparse(final_url).netloc.lower()!=target_domain:
            await page.close()
            return page_errors,"",detail_found,set(),final_url,[],[]
        if scope_root_url and not url_in_scan_scope(final_url,scope_root_url):
            await page.close()
            return page_errors,"",detail_found,set(),final_url,[],[]

        page_text=await page.evaluate("document.body.innerText")
        html_content=await page.content()
        soup=BeautifulSoup(html_content,'html.parser')

        resp_meta=None if nav_recovered else response
        if _is_charset_check_html_document(resp_meta,html_content)and not await check_page_utf8(resp_meta,soup):
            page_errors.append("3.語系編碼")
        if check_flash_or_legacy_ria(html_content,soup):
            page_errors.append("9.動畫格式")

        # 16項指標判定
        if soup.find('link',rel=lambda x:x and 'icon' in x.lower()):detail_found["favicon"]=True
        if "隱私權" in page_text:detail_found["privacy"]=True
        if any(k in page_text for k in ["資訊安全","資安","安全"]):detail_found["security"]=True
        if any(k in page_text for k in ["開放資料","資料開放"]):detail_found["open_data"]=True
        if re.search(r'(?:\+886|0800|0\d{1,3}[\s\-)]?\d{2,4}[\s\-]?\d{4})',page_text):detail_found["phone"]=True
        if re.search(r'(?:北市|新北市|桃園|台中|台南|高雄|縣|市).{2,15}\d+[號樓Ff]',page_text):detail_found["address"]=True
        if '無障礙' in page_text or soup.find(lambda t:t.has_attr('alt') and '無障礙' in t['alt']):detail_found["accessibility"]=True
        if any(k in page_text for k in ["網站導覽","Sitemap"]):detail_found["nav"]=True
        if any(k in page_text.upper() for k in ["EN","ENGLISH","英文"]):detail_found["lang_ver"]=True
        if any(k in page_text for k in ["2024","2025","2026"]):detail_found["date_info"]=True
        if any(k in page_text for k in ["信箱","聯絡我們"]):detail_found["opinion"]=True
        if soup.find('input',{'type':['search','text']}):detail_found["search"]=True
        if soup.find('meta',attrs={'name':'viewport'}):detail_found["rwd"]=True
        if any(kw in html_content.lower() for kw in ['google-analytics','gtag']):detail_found["stats"]=True

        ext_http_seen=set()
        extracted_links=await collect_links_with_pagination(
            page,browser_context,final_url,target_domain,file_exts,ext_http_seen=ext_http_seen
        )
        if scope_root_url:
            extracted_links={u for u in extracted_links if url_in_scan_scope(u,scope_root_url)}
        html_content=await page.content()
        ext_urls=ext_http_seen|extract_external_http_links(html_content,final_url,target_domain)
        ext_urls|=await _external_http_hrefs_from_item_live_dom(
            page, final_url, target_domain
        )
        ref_probe=referer_url or _site_root_url(final_url)
        external_failed,external_probed=await probe_external_links_unreachable(
            ext_urls,ref_probe,probe_cache,probe_sem=external_probe_sem
        )
    except Exception as e:
        if page_exceptions is not None:
            page_exceptions.append(f"{url}: {type(e).__name__}: {e}")
    finally:await page.close()
    return page_errors,html_content,detail_found,extracted_links,final_url,external_failed,external_probed

async def _run_scan_parallel_batch(targets,url_norm,scope_root,referer_url,host_concurrency:int=4,external_probe_cache:Optional[dict]=None,page_exceptions:Optional[List[str]]=None):
    """同一網域下以 3～5 上限之併發數掃描一批網址（各頁獨立 Context 與 UA）。"""
    if not targets:
        return []
    td=urlparse(url_norm).netloc.lower()
    cap=max(1,int(host_concurrency or 4))
    sem=asyncio.Semaphore(cap)
    if external_probe_cache is None:external_probe_cache={}
    probe_cache=external_probe_cache
    batch_ext_sem=asyncio.Semaphore(_BATCH_EXTERNAL_PROBE_CONCURRENCY)
    async with async_playwright() as p:
        browser=await p.chromium.launch(headless=True)
        try:
            async def work(t):
                async with sem:
                    ua=_pick_user_agent()
                    ctx=await browser.new_context(user_agent=ua,ignore_https_errors=True)
                    try:
                        return await asyncio.wait_for(
                            check_single_page(
                                ctx,t,td,[],scope_root,
                                referer_url=referer_url,
                                external_probe_cache=probe_cache,
                                page_exceptions=page_exceptions,
                                external_probe_sem=batch_ext_sem,
                            ),
                            timeout=_PAGE_SCAN_TIMEOUT_S,
                        )
                    except asyncio.TimeoutError as e:
                        if page_exceptions is not None:
                            page_exceptions.append(f"{t}: TimeoutError: page scan exceeded {_PAGE_SCAN_TIMEOUT_S}s")
                        return [],"",{k:False for k in ["favicon","privacy","security","phone","address","open_data","accessibility","nav","lang_ver","search","opinion","rwd","stats","date_info"]},set(),t,[],[]
                    finally:
                        await ctx.close()
            return await asyncio.gather(*[work(x)for x in targets],return_exceptions=True)
        finally:
            await browser.close()
