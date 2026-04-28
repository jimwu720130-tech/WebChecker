import streamlit as st
import asyncio
import random
import html as html_stdlib
import uuid
from datetime import datetime
from webchecker_core import (
    APP_MODE_SCAN, APP_MODE_FAV, APP_MODE_EXCLUDE, APP_MODE_INDICATOR_HELP,
    SCAN_ENGINE_BUILD,
    _SCAN_FAV_NONE, _format_fav_select_row, _fav_select_rows,
    ordered_visited_urls_for_export, visited_urls_to_excel_bytes,
    load_favorites, save_favorites, load_config, save_config,
    normalize_url_input, get_clean_domain, url_in_scan_scope, _site_root_url,
    get_pagespeed_score, _run_scan_parallel_batch,
)

def _ordered_visited_urls_for_export():
    return ordered_visited_urls_for_export(
        st.session_state.visited_urls, st.session_state.get("visited_order") or []
    )

def render_wrapped_external_links(urls):
    _lc="#6ea8ff"if st.session_state.get("ui_theme")=="dark"else"#0b57d0"
    st.markdown(
        f"<style>.wc-url-row{{word-break:break-all;overflow-wrap:anywhere;max-width:100%;font-size:0.88rem;"
        f"line-height:1.45;margin:0 0 0.65em 0;}}.wc-url-row a{{color:{_lc};text-decoration:underline;}}</style>",
        unsafe_allow_html=True,
    )
    for raw in urls:
        esc = html_stdlib.escape(raw, quote=True)
        st.markdown(
            f'<p class="wc-url-row"><a href="{esc}" target="_blank" rel="noopener noreferrer">{esc}</a></p>',
            unsafe_allow_html=True,
        )

def _streamlit_theme_inject():
    """依 session 外觀注入淺/深樣式（與側欄切換鈕連動）。"""
    t=st.session_state.get("ui_theme","light")
    if t=="dark":
        return"""<style>
.stApp,[data-testid="stAppViewContainer"]{background:#0f1419!important;color:#e6edf3!important;}
section.main .block-container,.main .block-container{color:#e6edf3!important;}
section[data-testid="stSidebar"]{background:#1a2332!important;border-right:1px solid #2d3a4d!important;}
.stSidebar [data-testid="stMarkdownContainer"] p,.stSidebar label,
.stSidebar .stText,.stSidebar p,.stSidebar span{color:#e6edf3!important;}
h1,h2,h3,.stMarkdown h1,.stMarkdown h2,.stMarkdown h3,.stMarkdown p,label,p,.stText{color:#e6edf3!important;}
.stTextInput label,.stTextArea label,.stSelectbox label,.stCheckbox label{color:#b8c5d6!important;}
.stTextInput input,.stTextArea textarea,[data-baseweb="input"] input,[data-baseweb="textarea"] textarea{
  background:#0d1117!important;color:#e6edf3!important;border-color:#2d3a4d!important;
}
[data-baseweb="select"]>div{background-color:#0d1117!important;border-color:#2d3a4d!important;}
[data-baseweb="select"] div[role="combobox"]{color:#e6edf3!important;}
.stButton>button[kind="secondary"]{background:#2d3a4d!important;color:#e6edf3!important;}
.stButton>button[kind="primary"]{background:#2563eb!important;color:#fff!important;}
div[data-testid="stExpander"] details summary,div[data-testid="stExpander"] details summary p,
div[data-testid="stExpander"] details summary span{color:#e6edf3!important;}
div[data-testid="stExpander"] .streamlit-expanderContent{background:#161b22!important;}
.stCodeBlock,pre,.stCodeBlock code{background:#0d1117!important;color:#d4d4d4!important;}
.stDownloadButton button,.stDownloadButton>button{background:#2563eb!important;color:#fff!important;}
div[data-testid="column"]{color:#e6edf3!important;}
hr{border-color:#2d3a4d!important;}
</style>"""
    return"""<style>
.stApp,[data-testid="stAppViewContainer"]{background:#ffffff!important;color:#262730!important;}
section.main .block-container,.main .block-container{color:#262730!important;}
section[data-testid="stSidebar"]{background:#f0f2f6!important;}
h1,h2,h3,.stMarkdown p,label,p,.stText{color:#262730!important;}
.stTextInput input,.stTextArea textarea,[data-baseweb="input"] input{
  background:#fff!important;color:#262730!important;border-color:#d3dae4!important;
}
[data-baseweb="select"]>div{background:#fff!important;color:#262730!important;}
.stButton>button[kind="secondary"]{background:#fff!important;color:#262730!important;border:1px solid #d3dae4!important;}
</style>"""

# ==========================================
# UI介面與進度顯示
# ==========================================
st.set_page_config(page_title="網站檢核工具",page_icon="🔍",layout="centered")
if "ui_theme" not in st.session_state:st.session_state.ui_theme="light"
st.markdown(_streamlit_theme_inject(),unsafe_allow_html=True)
if "app_mode" not in st.session_state:st.session_state.app_mode=APP_MODE_SCAN
if "visited_urls" not in st.session_state:st.session_state.visited_urls,st.session_state.to_visit_urls=set(),[]
if "failure_report" not in st.session_state:st.session_state.failure_report={}
if "is_scanning" not in st.session_state:st.session_state.is_scanning,st.session_state.psi_done=False,False
if "psi_result" not in st.session_state:st.session_state.psi_result={}
if "global_features" not in st.session_state:
    st.session_state.global_features={k:False for k in ["favicon","privacy","security","phone","address","open_data","accessibility","nav","lang_ver","search","opinion","rwd","stats","date_info"]}
if "run_psi" not in st.session_state:st.session_state.run_psi=True
if "scan_url_field"not in st.session_state:st.session_state.scan_url_field=""
if "fav_pick"not in st.session_state:st.session_state.fav_pick=_SCAN_FAV_NONE
if isinstance(st.session_state.get("fav_pick"),str):st.session_state.fav_pick=_SCAN_FAV_NONE
if "visited_order"not in st.session_state:st.session_state.visited_order=[]
if "external_probed_order"not in st.session_state:st.session_state.external_probed_order=[]
if "external_probed_seen"not in st.session_state:st.session_state.external_probed_seen=set()

st.sidebar.title("🗂️系統選單")
scan_btn=st.sidebar.button(APP_MODE_SCAN,use_container_width=True,type="primary" if st.session_state.app_mode==APP_MODE_SCAN else "secondary")
fav_btn=st.sidebar.button(APP_MODE_FAV,use_container_width=True,type="primary" if st.session_state.app_mode==APP_MODE_FAV else "secondary")
exclude_btn=st.sidebar.button(APP_MODE_EXCLUDE,use_container_width=True,type="primary" if st.session_state.app_mode==APP_MODE_EXCLUDE else "secondary")
indicator_help_btn=st.sidebar.button(APP_MODE_INDICATOR_HELP,use_container_width=True,type="primary" if st.session_state.app_mode==APP_MODE_INDICATOR_HELP else "secondary")
if scan_btn:st.session_state.app_mode=APP_MODE_SCAN;st.rerun()
if fav_btn:st.session_state.app_mode=APP_MODE_FAV;st.rerun()
if exclude_btn:st.session_state.app_mode=APP_MODE_EXCLUDE;st.rerun()
if indicator_help_btn:st.session_state.app_mode=APP_MODE_INDICATOR_HELP;st.rerun()
st.sidebar.divider()
st.sidebar.caption(f"檢核核心版本：{SCAN_ENGINE_BUILD}")
st.sidebar.caption("外觀")
if st.sidebar.button("🌙 切換為深色"if st.session_state.get("ui_theme")=="light"else"☀️ 切換為淺色",use_container_width=True,key="wc_ui_theme_btn"):
    st.session_state.ui_theme="dark"if st.session_state.get("ui_theme")=="light"else"light"
    st.rerun()

# ==========================================
# 頁面1：掃描作業
# ==========================================
if st.session_state.app_mode==APP_MODE_SCAN:
    st.title("🔍網站檢核工具")
    if st.session_state.pop("_wc_pending_clear_scan",False):
        st.session_state.visited_urls,st.session_state.to_visit_urls,st.session_state.failure_report=set(),[],{}
        st.session_state.is_scanning,st.session_state.psi_done=False,False
        st.session_state.global_features={k:False for k in st.session_state.global_features}
        st.session_state.visited_order=[]
        st.session_state.external_probed_order=[]
        st.session_state.external_probed_seen=set()
        if "scan_scope_root" in st.session_state:del st.session_state["scan_scope_root"]
        if "_wc_scan_host_concurrency" in st.session_state:del st.session_state["_wc_scan_host_concurrency"]
        if "_scan_page_exceptions" in st.session_state:del st.session_state["_scan_page_exceptions"]
        if "_wc_external_url_probe_cache" in st.session_state:del st.session_state["_wc_external_url_probe_cache"]
        for _wk in("scan_url_field","fav_pick"):
            if _wk in st.session_state:del st.session_state[_wk]
    _favs=load_favorites()
    _fav_rows=_fav_select_rows(_favs)
    if st.session_state.get("fav_pick")not in _fav_rows:st.session_state.fav_pick=_SCAN_FAV_NONE
    with st.form("scan_start_form",clear_on_submit=False):
        st.selectbox(
            "⭐常用網站",
            _fav_rows,
            key="fav_pick",
            format_func=_format_fav_select_row,
            help="僅在下方「自訂網站」為空時，才會使用此處選到的網址。",
        )
        st.text_input(
            "🌐自訂網站（請貼上完整網址）",
            key="scan_url_field",
            placeholder="https://…",
            help="優先使用此欄位；有填寫時將忽略上方常用網站。貼上後請按下方「開始掃描」。"
            "若入口為某頁（如 …/Sale/Login）而同層另有 …/Sale/其他頁，程式會自動以 …/Sale 為掃描範圍；"
            "若您要**只掃某一層子目錄**且該層名稱無副檔名，請在網址末端加 /（例如 …/Sale/）。",
        )
        st.checkbox("⚡同時執行載入速度檢測(PageSpeedAPI)",key="run_psi")
        start_scan=st.form_submit_button("🚀開始掃描",use_container_width=True,type="primary")
    raw_custom=(st.session_state.get("scan_url_field")or"").strip()
    custom_first=raw_custom
    pick=st.session_state.get("fav_pick",_SCAN_FAV_NONE)
    if isinstance(pick,str):pick=_SCAN_FAV_NONE
    if custom_first:
        input_url=custom_first
    elif pick[0]!="__none__":
        input_url=(pick[2]or"").strip()
    else:
        input_url=""
    url_norm=normalize_url_input(input_url)
    run_psi=st.session_state.get("run_psi",True)
    if start_scan:
        if url_norm:
            # 每次「開始掃描」視為新一輪：避免 URL 已在 visited 時 pop 後佇列被掏空、畫面卡在 0 頁
            st.session_state.to_visit_urls=[url_norm]
            st.session_state.visited_urls=set()
            st.session_state.visited_order=[]
            st.session_state.external_probed_order=[]
            st.session_state.external_probed_seen=set()
            st.session_state.failure_report={}
            st.session_state["_wc_external_url_probe_cache"]={}
            st.session_state.global_features={k:False for k in ["favicon","privacy","security","phone","address","open_data","accessibility","nav","lang_ver","search","opinion","rwd","stats","date_info"]}
            st.session_state.psi_done=False
            st.session_state.is_scanning=True
            st.session_state.scan_scope_root=url_norm
            st.session_state._wc_scan_host_concurrency=random.randint(3,5)
            for _k in("_scan_page_exception","_scan_last_exception","_scan_page_exceptions"):
                if _k in st.session_state:del st.session_state[_k]
        else:st.warning("請在「自訂網站」貼上網址，或於「常用網站」選擇一筆清單。")

    col2,col3=st.columns(2)
    if col2.button("⏸️暫停掃描",use_container_width=True):st.session_state.is_scanning=False
    if col3.button("🗑️清除重置",use_container_width=True):
        st.session_state._wc_pending_clear_scan=True
        st.rerun()

    status_box,progress_area=st.empty(),st.empty()
    if st.session_state.is_scanning and st.session_state.to_visit_urls:
        if run_psi and not st.session_state.psi_done:
            status_box.info("⏳正在執行PageSpeedAPI載入速度檢測...")
            with st.spinner("⏳ PageSpeed 測速中（約 30～180 秒，請勿關閉視窗）…"):
                success,avg,m,d,msg=get_pagespeed_score(url_norm)
            st.session_state.psi_result={"success":success,"avg":avg,"m":m,"d":d,"msg":msg}
            st.session_state.psi_done=True
        _scope=st.session_state.get("scan_scope_root")or url_norm
        if "_wc_scan_host_concurrency" not in st.session_state:
            st.session_state._wc_scan_host_concurrency=random.randint(3,5)
        _cap=max(1,int(st.session_state._wc_scan_host_concurrency))
        batch=[]
        while len(batch)<_cap and st.session_state.to_visit_urls and st.session_state.is_scanning:
            cand=st.session_state.to_visit_urls.pop(0)
            if cand in st.session_state.visited_urls:continue
            if not url_in_scan_scope(cand,_scope):continue
            batch.append(cand)
        if not batch:
            st.session_state.is_scanning=False
            if st.session_state.visited_urls:
                progress_area.markdown(f"📊**掃描進度**：已完成**{len(st.session_state.visited_urls)}**頁｜待處理**0**筆")
                status_box.success("🎉全站掃描完成！")
            else:
                progress_area.markdown("📊**掃描進度**：佇列已無待掃描網址。")
                status_box.warning("未處理任何頁面。若剛按過「開始掃描」，請確認網址正確，或按「清除重置」後再試。")
        else:
            _shown="`、`".join(x.replace("`","'")for x in batch)
            progress_area.markdown(
                f"🔎**目前掃描**（本輪 **{len(batch)}** 頁／併發上限 **{_cap}**）： `{_shown}`  \n"
                f"📊已完成 **{len(st.session_state.visited_urls)}** 頁｜佇列待處理 **{len(st.session_state.to_visit_urls)}** 筆"
            )
            _sc=st.session_state.get("scan_scope_root")or url_norm
            _ref=(st.session_state.visited_order[-1]if st.session_state.visited_order else _site_root_url(url_norm))
            with st.spinner("⏳ Playwright 掃描中（含分頁與檔案按鈕時可能需數分鐘，畫面會暫停更新屬正常）…"):
                try:
                    results=asyncio.run(_run_scan_parallel_batch(
                        batch,url_norm,_sc,_ref,
                        host_concurrency=st.session_state.get("_wc_scan_host_concurrency") or 4,
                        external_probe_cache=st.session_state.setdefault("_wc_external_url_probe_cache", {}),
                        page_exceptions=st.session_state.setdefault("_scan_page_exceptions", []),
                    ))
                except Exception as e:
                    st.session_state.is_scanning=False
                    st.error(f"掃描發生錯誤：{e}")
                    st.session_state.to_visit_urls[:0]=batch
                    results=[]
            for t,res in zip(batch,results):
                if isinstance(res,BaseException):
                    st.session_state.to_visit_urls.insert(0,t)
                    st.session_state.setdefault("_scan_page_exceptions",[]).append(f"{t}: {type(res).__name__}: {res}")
                    continue
                errs,html,found,links,final_url,ext_bad,ext_probed=res
                st.session_state.visited_urls.add(t)
                st.session_state.visited_order.append(t)
                for u in ext_probed:
                    if u not in st.session_state.external_probed_seen:
                        st.session_state.external_probed_seen.add(u)
                        st.session_state.external_probed_order.append(u)
                for k,v in found.items():
                    if v:st.session_state.global_features[k]=True
                for e in errs:
                    st.session_state.failure_report.setdefault(e,[]).append(t)
                fr5=st.session_state.failure_report.setdefault("5.有效連結",[])
                for u in ext_bad:
                    if u not in fr5:
                        fr5.append(u)
                for l in links:
                    if not url_in_scan_scope(l,_sc):continue
                    if l not in st.session_state.visited_urls:
                        st.session_state.to_visit_urls.append(l)
            if st.session_state.get("_scan_page_exceptions"):
                with st.expander("頁面解析時曾發生例外（該頁結果可能不完整）",expanded=False):
                    st.code("\n".join(st.session_state["_scan_page_exceptions"]))
                del st.session_state["_scan_page_exceptions"]
            if st.session_state.is_scanning and st.session_state.to_visit_urls:
                st.rerun()
            if not st.session_state.to_visit_urls:
                st.session_state.is_scanning=False
                progress_area.markdown(f"📊**掃描進度**：已完成**{len(st.session_state.visited_urls)}**頁｜待處理**0**筆")
                status_box.success("🎉全站掃描完成！")

    if st.session_state.visited_urls:
        st.markdown("---")
        _ord=_ordered_visited_urls_for_export()
        _ext=st.session_state.get("external_probed_order")or[]
        _fr5=set((st.session_state.failure_report or{}).get("5.有效連結")or[])
        _xlsx=visited_urls_to_excel_bytes(_ord,_ext,_fr5)
        _fn=f"掃描網址清單_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        st.download_button(
            "📥 下載掃描網址清單（Excel）",
            data=_xlsx,
            file_name=_fn,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_scan_urls_xlsx",
        )
        st.caption(
            f"共 **{len(_ord)}** 筆站內、**{len(_ext)}** 筆外站。試算表「掃描清單」中，**前段為站內網址**，"
            f"**接續同表後段為外站**（「類型」欄可區分）。"
        )
        st.markdown("### 📊完整檢核指標報告")
        indicators=[
            ("1.基本要件",st.session_state.global_features['privacy'] and st.session_state.global_features['security'] and st.session_state.global_features['address']),
            ("2.導覽功能",st.session_state.global_features['nav']),
            ("3.語系編碼","3.語系編碼"),("4.語言版本",st.session_state.global_features['lang_ver']),
            ("5.有效連結","5.有效連結"),("6.資料正確","manual"),
            ("7.資料即時",st.session_state.global_features['date_info']),("8.傳輸協議","8.傳輸協議"),
            ("9.動畫格式","9.動畫格式"),("10.文件格式","10.文件格式"),
            ("11.意見信箱",st.session_state.global_features['opinion']),("12.搜尋服務",st.session_state.global_features['search']),
            ("13.跨瀏覽器","manual"),("14.跨裝置",st.session_state.global_features['rwd']),
            ("15.載入速度","psi"),("16.流量統計",st.session_state.global_features['stats'])
        ]
        for name,status in indicators:
            if status=="manual":st.write(f"⚠️**{name}**：待人工搭配外部工具檢查")
            elif status=="psi":
                psi=st.session_state.psi_result
                if psi and psi.get('success'):st.write(f"{'✅' if psi['avg']>=50 else '❌'}**15.載入速度**：{psi['avg']:.1f}分（手機 {psi['m']:.1f}、電腦 {psi['d']:.1f}）")
                else:st.write(f"⚠️**15.載入速度**：測速失敗")
            elif isinstance(status,str) and status in ["3.語系編碼","5.有效連結","8.傳輸協議","9.動畫格式","10.文件格式"]:
                fails=st.session_state.failure_report.get(status,[])
                if not fails:st.write(f"✅**{name}**：符合")
                else:
                    with st.expander(f"❌**{name}**：不符合(共{len(fails)}處異常)",expanded=False):
                        render_wrapped_external_links(fails)
            else:st.write(f"{'✅' if status else '❌'}**{name}**：{'符合' if status else '不符合'}")

# ==========================================
# 頁面：檢核指標說明
# ==========================================
elif st.session_state.app_mode==APP_MODE_INDICATOR_HELP:
    st.title(APP_MODE_INDICATOR_HELP)
    st.markdown(
        "以下每一項皆分為 **1. 檢測內容**（驗什麼）與 **2. 技術邏輯**（程式怎麼判）。多數為**輔助參考**，上線或稽核前仍建議人工複核或搭配專用工具。"
    )
    st.divider()
    st.subheader("1. 基本要件")
    st.markdown(
        "1. **檢測內容**：全站是否曾出現過「隱私權」說明、資安相關用語，以及看起來像聯絡地址的敘述（縣市、路名、門牌等線索）。三者須在**整次掃描中各自至少出現過一次**（可分散在不同頁面），報表「基本要件」才顯示符合。\n\n"
        "2. **技術邏輯**：以 Playwright 取得 `document.body.innerText` 後做字串比對——「隱私權」；資安關鍵字為「資訊安全」「資安」「安全」擇一；地址為正則比對（含縣市關鍵字、路名片段、數字與號／樓／F 等）。任一頁面曾命中則將對應旗標寫入 `session_state.global_features`，報表以三旗標**同時為真**為通過。"
    )
    st.subheader("2. 導覽功能")
    st.markdown(
        "1. **檢測內容**：全站是否曾出現「網站導覽」或「Sitemap」字樣（作為導覽資訊之粗略指標）。\n\n"
        "2. **技術邏輯**：於 `innerText` 中搜尋子字串「網站導覽」或「Sitemap」（大小寫不拘）。任一模組頁面命中則 `global_features['nav']` 為真。"
    )
    st.subheader("3. 語系編碼")
    st.markdown(
        "1. **檢測內容**：被視為一般 HTML／XHTML 的頁面，是否以 **UTF-8** 宣告或可解讀為有效 UTF-8 本文。\n\n"
        "2. **技術邏輯**：先依 `_is_charset_check_html_document` 判斷是否屬應檢核之網頁（`Content-Type` 為 html/xhtml，或本文開頭像 HTML；圖影音、PDF 等略過）。再以 `check_page_utf8`：優先 `meta charset`；其次 `http-equiv=\"Content-Type\"` 內之 charset；再讀回應標頭 `Content-Type` 的 charset；皆無則取回應 body 並以嚴格 UTF-8 解碼（含略過 BOM）。未通過之頁面會列入「3.語系編碼」失敗清單。"
    )
    st.subheader("4. 語言版本")
    st.markdown(
        "1. **檢測內容**：全站是否曾出現多語／英文版常見字樣（如 EN、ENGLISH、英文）。\n\n"
        "2. **技術邏輯**：將 `innerText` 轉成大寫後，檢查是否含「EN」「ENGLISH」或一般大小寫之「英文」子字串。任一頁命中則 `global_features['lang_ver']` 為真。"
    )
    st.subheader("5. 有效連結")
    st.markdown(
        "1. **檢測內容**：佇列中每個待掃網址是否能成功載入；是否出現 HTTP 錯誤；「網址指向檔案（如 PDF）卻回傳 HTML 錯誤頁」等偽裝有效之連結；以及**頁面內對其他網域的 http(s) 超連結**是否仍可連線。\n\n"
        "2. **技術邏輯**：同網址以 `page.goto(..., wait_until=\"commit\")` 讀取 `response.status` 與 `Content-Type` 判斷（含 PDF 偽連結規則如上）。另於每個成功解析之 HTML 頁，以 `extract_external_http_links` 自**整頁**（含頁首、主內容、頁尾）之 `a`／`area` 的 **href** 收集**外網** http(s) 連結，並以 `probe_external_links_unreachable` 對**掃到的每一筆**外站做 HEAD／GET（`verify=False`、整輪掃描共用快取與併發上限）；外站極多時單頁耗時會變長。另以 `_external_http_hrefs_from_item_live_dom` 補抓 **#Item** 內即時 DOM 與序列化 HTML 不一致時的外站（頁首／頁尾不在此補抓範圍，但一般已由前述整頁 BeautifulSoup 掃到）。**站內**待掃佇列由 `extract_same_domain_links`（另含 onclick／data-href 等）與分頁、檔案按鈕補齊。**介面只顯示不符合的連結**，通過者不會出現在展開清單中。"
    )
    st.subheader("6. 資料正確")
    st.markdown(
        "1. **檢測內容**：網頁敘述、數字、法規名稱等是否與正式來源一致。\n\n"
        "2. **技術邏輯**：**本工具不實作自動比對**。報表固定顯示為待人工與外部權威資料核對。"
    )
    st.subheader("7. 資料即時")
    st.markdown(
        "1. **檢測內容**：頁面上是否出現近年份字樣，作為「可能有更新痕跡」之極粗指標。\n\n"
        "2. **技術邏輯**：於 `innerText` 搜尋「2024」「2025」「2026」任一字串。命中則 `global_features['date_info']` 為真；**不代表**內容確實最新或正確。"
    )
    st.subheader("8. 傳輸協議")
    st.markdown(
        "1. **檢測內容**：全站是否全面採 **HTTPS**、憑證與轉址設定是否正確。\n\n"
        "2. **技術邏輯**：**目前程式未對 HTTPS 做自動檢測**，亦不會寫入本項失敗清單；報表列此項僅為提醒，需以瀏覽器、主機設定或資安掃描工具另行確認。"
    )
    st.subheader("9. 動畫格式")
    st.markdown(
        "1. **檢測內容**：網頁是否仍使用 Flash、Silverlight、Java Applet 等需外掛之舊式 RIA。\n\n"
        "2. **技術邏輯**：`check_flash_or_legacy_ria` 比對 HTML 小寫字串與 BeautifulSoup 之 `embed`／`object`（含 type、data、src、classid、codebase）：偵測 `.swf`、`application/x-shockwave-flash`、Silverlight／`.xap`、`<applet`、Java 相關 MIME 等。若命中則該頁加入「9.動畫格式」失敗清單。一般 CSS／Canvas／影片標籤**未必**會觸發。"
    )
    st.subheader("10. 文件格式")
    st.markdown(
        "1. **檢測內容**：待掃網址是否為 **Word／Excel／PowerPoint、OpenDocument、RAR** 等檔案連結（實際是否允許依各機關規範）。\n\n"
        "2. **技術邏輯**：若 URL 路徑結尾為 `.doc`、`.docx`、`.xls`、`.xlsx`、`.ppt`、`.pptx`、`.rar`、`.odt`、`.ods`、`.odp` 之一，於該筆結果附加「10.文件格式」並列入該項清單（屬**格式提醒**，與 HTTP 成敗或 PDF 偽連結之「5.有效連結」不同）。"
    )
    st.subheader("11. 意見信箱")
    st.markdown(
        "1. **檢測內容**：全站是否曾出現「信箱」或「聯絡我們」等聯絡管道字樣。\n\n"
        "2. **技術邏輯**：於 `innerText` 搜尋上列子字串。任一頁命中則 `global_features['opinion']` 為真。"
    )
    st.subheader("12. 搜尋服務")
    st.markdown(
        "1. **檢測內容**：頁面是否具備可輸入之查詢欄位（粗略對應「搜尋」功能）。\n\n"
        "2. **技術邏輯**：以 BeautifulSoup 尋找 `input`，且 `type` 為 `search` 或 `text` 即視為命中。可能與一般表單欄位混淆，僅供參考。"
    )
    st.subheader("13. 跨瀏覽器")
    st.markdown(
        "1. **檢測內容**：版面與功能在 Chrome、Edge、Safari、Firefox 等瀏覽器是否皆可用。\n\n"
        "2. **技術邏輯**：**本工具僅以 Playwright Chromium 自動化**，無多瀏覽器矩陣。報表顯示為待人工或專用相容性測試。"
    )
    st.subheader("14. 跨裝置")
    st.markdown(
        "1. **檢測內容**：是否有利於手機／平板顯示的常見設定（響應式設計之必要條件之一）。\n\n"
        "2. **技術邏輯**：解析後之 HTML 中是否存在 `meta name=\"viewport\"`。有則 `global_features['rwd']` 為真；**不保證**實際版面於小螢幕上可用。"
    )
    st.subheader("15. 載入速度")
    st.markdown(
        "1. **檢測內容**：以第三方對外公開之效能分數，粗估首頁（起始 URL）載入體感。\n\n"
        "2. **技術邏輯**：若掃描時勾選測速，呼叫 Google PageSpeed Insights v5 API，分別以 `strategy=mobile` 與 `strategy=desktop` 各取 Lighthouse 之 performance score，換算為 0～100 分後**取平均**。介面以**平均 ≥50** 顯示綠燈、否則紅燈；失敗時顯示測速錯誤。"
    )
    st.subheader("16. 流量統計")
    st.markdown(
        "1. **檢測內容**：原始碼中是否出現常見 **Google Analytics** 相關片段。\n\n"
        "2. **技術邏輯**：於完整 HTML 字串（小寫）搜尋子字串 `google-analytics` 或 `gtag`。命中則 `global_features['stats']` 為真；未命中**不表示**未使用其他分析工具。"
    )

# ==========================================
# 頁面：常用網站設定
# ==========================================
elif st.session_state.app_mode==APP_MODE_FAV:
    st.title("⭐常用網站設定")
    with st.form("add_fav",clear_on_submit=True):
        st.subheader("➕新增網站")
        n=st.text_input("網站名稱",key="add_fav_name")
        u=st.text_area("網站網址（可貼上完整網址）",key="add_fav_url",height=100,placeholder="https://…")
        add_sub=st.form_submit_button("儲存至清單")
    if add_sub and n and u:
        u_first=(u.strip().splitlines()[0]if u.strip()else"").strip()
        f_list=load_favorites()
        f_list.append({"id":str(uuid.uuid4()),"name":n.strip(),"url":normalize_url_input(u_first)})
        save_favorites(f_list);st.success("已新增！");st.rerun()
    current_favs=load_favorites()
    for i,item in enumerate(current_favs):
        with st.expander(f"📌{item['name']}({item['url']})"):
            new_n=st.text_input("修改名稱",value=item['name'],key=f"edit_n_{item['id']}")
            new_u=st.text_area("修改網址（可貼上）",value=item['url'],key=f"edit_u_{item['id']}",height=96)
            c_save,c_del=st.columns(2)
            if c_save.button("💾儲存變更",key=f"save_{item['id']}",use_container_width=True):
                u_fix=(new_u.strip().splitlines()[0]if new_u.strip()else"").strip()
                current_favs[i]={"id":item['id'],"name":new_n.strip(),"url":normalize_url_input(u_fix)}
                save_favorites(current_favs);st.success("已更新！");st.rerun()
            if c_del.button("🗑️刪除網站",key=f"del_{item['id']}",use_container_width=True):
                current_favs.pop(i);save_favorites(current_favs);st.rerun()

# ==========================================
# 頁面：排除規則設定
# ==========================================
elif st.session_state.app_mode==APP_MODE_EXCLUDE:
    st.title("🛡️排除規則設定")
    in_url=st.text_input("請輸入網域(例如:recycle.eri.com.tw)")
    key=get_clean_domain(in_url)
    if key:
        _cfg=load_config();rules=st.text_area("排除規則(每行一個關鍵字)",value=_cfg.get(key,""))
        if st.button("💾儲存規則"):
            _cfg[key]=rules;save_config(_cfg);st.success("規則已儲存！")