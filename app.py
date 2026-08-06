"""
直播记录系统 v5.4
=====================
- 环节自动交替 + 下拉手动可改
- 第1轮干货手动开始时间+时长；之后环节开始/结束时间自动取（可改），时长=结束-开始
- 平均流速 = 累计总进房 ÷ 本轮总时长
- 一轮 = 干货 + 售卖 两行；仅「环节/时间/时长/出单/退款」5 列分两行，
- 出单/退款：干货填本环节值，售卖填增量差值
- 成交/成交率：一轮一行，成交=售卖上传总出单，成交率=成交÷总加购
- 曝光/进房/商品曝光/点击/加购：第一轮=总累计，之后=跨轮增量差值
- 率列统一加 % 后缀
- 颜色规则：可填写输入框=白底黑字；按钮/上传区/图表=深底白字
- 添加主播：内联输入框+按钮，持久化到 hosts
- Excel 表：Handsontable 真网格（增删行列 / 合并单元格 / 标黄 / 主键锁定 / 回写）
- 往期查询：结束本场写本地 sessions/ 文件，可查看+下载；图表可下载 PNG
- 右上角：撤回 + 导出；自动保存：每次操作入栈，session 内可 Undo
- 百分比显示 30%（整数百分号）
- 不上传图也可纯手填数据；切换环节自动清空已上传图片
- OCR 暂未接入（Python 3.13 兼容问题，保持手动录入）
"""
import streamlit as st
import pandas as pd
import os
from datetime import datetime, time as dtime
import re
import os
import json
import time
from io import BytesIO
import streamlit.components.v1 as components

# Supabase：优先 Streamlit secrets，再回退到 env vars，最后回退到 hardcode
try:
    import os
    _sec = {}
    try:
        _sec = st.secrets
    except Exception:
        pass
    SUPABASE_URL = _sec.get("SUPABASE_URL", os.environ.get('SUPABASE_URL', "https://zetaijjtdabbwqoomtpm.supabase.co"))
    SUPABASE_KEY = _sec.get("SUPABASE_KEY", os.environ.get('SUPABASE_KEY', "sb_publishable_gfEkoBv9YA2yPJGbr-tkQg_gVfmEyPr"))
    db = create_client(SUPABASE_URL, SUPABASE_KEY)
    HAS_DB = bool(SUPABASE_URL and SUPABASE_KEY)
except Exception as e:
    print(f'DB init err: {e}')
    HAS_DB = False
    db = None

# ============================================================
HAS_OCR = False  # OCR 已禁用

# ============================================================
try:
    import plotly.graph_objects as go
    HAS_PLOTLY = True
except Exception:
    pass

# 图表导出（kaleido）
HAS_KALEIDO = False
try:
    import kaleido  # noqa
    HAS_KALEIDO = True
except Exception:
    pass

# ============================================================
# 组件：Handsontable 真·Excel 网格（本地打包，离线可用）
# ============================================================
_HOTGRID_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "components", "hotgrid")
try:
    hotgrid = components.declare_component("hotgrid", path=_HOTGRID_PATH)
except Exception:
    hotgrid = None

def ensure_sessions_dir():
    d = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'sessions')
    os.makedirs(d, exist_ok=True)
    return d

def auto_save_path():
    """自动保存文件路径：sessions/_autosave.json"""
    return os.path.join(ensure_sessions_dir(), '_autosave.json')

def auto_save():
    """持久化当前 data_rows（1s 防抖）"""
    now = time.time()
    last = st.session_state.get('_last_autosave_time', 0)
    if now - last < 1.0: return
    st.session_state._last_autosave_time = now
    try:
        snap = {
            'session': st.session_state.current_session,
            'data_rows': [dict(r) for r in st.session_state.data_rows],
            'current_round': st.session_state.current_round,
            'current_phase': st.session_state.current_phase,
            'first_round': st.session_state.first_round,
            'first_dry_start': st.session_state.first_dry_start.strftime('%H:%M') if hasattr(st.session_state.first_dry_start, 'strftime') else str(st.session_state.first_dry_start),
            'first_dry_duration': st.session_state.first_dry_duration,
            'last_end_time': st.session_state.last_end_time.isoformat() if st.session_state.last_end_time and hasattr(st.session_state.last_end_time, 'isoformat') else None,
            'last_cumulative': st.session_state.last_cumulative,
            'session_round_totals': st.session_state.session_round_totals,
            'prev_round_totals': st.session_state.prev_round_totals,
            'round_label_override': st.session_state.round_label_override,
        }
        # 本地缓存（作为备份）
        try:
            with open(auto_save_path(), 'w', encoding='utf-8') as f:
                json.dump(snap, f, ensure_ascii=False, default=str)
        except: pass
        # 云端保存
        db_save("autosave", snap)
    except Exception:
        pass

def auto_load():
    """启动时恢复数据：先云端，再本地"""
    if st.session_state.get('_session_ended'):
        return False
    snap = db_load("autosave") if HAS_DB else None
    if not snap:
        path = auto_save_path()
        if not os.path.exists(path): return False
        try:
            with open(path, 'r', encoding='utf-8') as f:
                snap = json.load(f)
        except Exception:
            return False
    if not snap: return False
    try:
        st.session_state.current_session = snap.get('session', st.session_state.current_session)
        st.session_state.data_rows = snap.get('data_rows', [])
        st.session_state.current_round = snap.get('current_round', 1)
        st.session_state.current_phase = snap.get('current_phase', '干货')
        st.session_state.first_round = snap.get('first_round', True)
        st.session_state.first_dry_duration = snap.get('first_dry_duration', 40)
        st.session_state.last_cumulative = snap.get('last_cumulative')
        st.session_state.session_round_totals = snap.get('session_round_totals', {})
        st.session_state.prev_round_totals = snap.get('prev_round_totals', {})
        st.session_state.round_label_override = snap.get('round_label_override', {})
        fd = snap.get('first_dry_start', '7:00')
        try:
            h, m = map(int, str(fd).split(':')[:2])
            st.session_state.first_dry_start = dtime(h, m)
        except: pass
        let = snap.get('last_end_time')
        if let:
            try: st.session_state.last_end_time = datetime.fromisoformat(let)
            except: pass
        return True
    except Exception:
        return False

def plotly_download_button(fig, fname):
    """图表下载：优先 PNG(kaleido)，失败回退 HTML。"""
    import io
    try:
        if HAS_KALEIDO:
            buf = io.BytesIO()
            fig.write_image(buf, format='png', engine='kaleido', scale=2)
            st.download_button('📥 下载图表(PNG)', data=buf.getvalue(),
                              file_name=fname + '.png', mime='image/png',
                              key='dl_' + fname, use_container_width=True)
            return
    except Exception:
        pass
    html = fig.to_html(include_plotlyjs='cdn')
    st.download_button('📥 下载图表(HTML)', data=html,
                      file_name=fname + '.html', mime='text/html',
                      key='dl_' + fname, use_container_width=True)

# ============================================================
# 配置
# ============================================================
st.set_page_config(page_title="🎯 直播记录系统", page_icon="🎯", layout="wide", initial_sidebar_state="collapsed")

# ============================================================
# 登录页（仅当启用 DB 时）
# ============================================================
if HAS_DB:
    if 'user' not in st.session_state:
        st.markdown('<div style="text-align:center;padding:40px 0"><h2>🎯 直播记录系统</h2></div>', unsafe_allow_html=True)
        _, c1, c2, _ = st.columns([1, 2, 2, 1])
        with c1:
            st.markdown('<div style="text-align:right;padding-top:8px;color:rgba(255,255,255,0.6)">用户名：</div>', unsafe_allow_html=True)
        with c2:
            un = st.text_input('用户名', key='login_un', label_visibility='collapsed', placeholder='必填')
        _, c1, c2, _ = st.columns([1, 2, 2, 1])
        with c1:
            st.markdown('<div style="text-align:right;padding-top:8px;color:rgba(255,255,255,0.6)">团队码：</div>', unsafe_allow_html=True)
        with c2:
            tc = st.text_input('团队码', key='login_tc', label_visibility='collapsed', placeholder='可空')
        _, c1, c2, _ = st.columns([1, 2, 2, 1])
        with c2:
            share = st.checkbox('📤 共享我的数据给团队', key='login_share', value=False)
        _, c1, c2, _ = st.columns([1, 2, 2, 1])
        with c2:
            if st.button('进入', type='primary', use_container_width=True):
                if not un.strip():
                    st.error('请输入用户名')
                    st.stop()
                st.session_state.user = un.strip()
                st.session_state.team = tc.strip()
                st.session_state.share = share
                st.rerun()
        st.stop()

# 助手：DB 读写
def db_save(data_type, payload):
    if not HAS_DB: return
    try:
        db.table("data_store").upsert({
            "id": f"{st.session_state.user}_{data_type}",
            "user_name": st.session_state.user,
            "team_code": st.session_state.get('team', '') if st.session_state.get('share', False) else '',
            "data_type": data_type,
            "payload": payload
        }).execute()
    except Exception as e: print(f'db_save err: {e}')

def db_load(data_type):
    if not HAS_DB: return None
    try:
        r = db.table("data_store").select("*").eq("user_name", st.session_state.user).eq("data_type", data_type).execute()
        return r.data[0]['payload'] if r.data else None
    except Exception as e: print(f'db_load err: {e}'); return None

def db_load_sessions(user_only=False):
    """返回该用户的所有 session 数据；user_only=True 只看自己的"""
    if not HAS_DB: return []
    try:
        q = db.table("data_store").select("*").eq("data_type", "session")
        if user_only:
            r = q.eq("user_name", st.session_state.user).execute()
        else:
            team = st.session_state.get('team', '')
            if not team: return []
            r = q.eq("team_code", team).execute()
        return r.data
    except: return []

COLUMNS_ORDER = [
    '主播','轮次','环节','时间','时长','总时长','出单','退款',
    '总曝光','总进房','曝光','进房','平均流速',
    '总商品曝光','总点击','总加购','总点击率','总加购率',
    '商品曝光','点击','加购','点击率','加购率',
    '成交','成交率'
]
OCR_FIELDS = ['总曝光','总进房','总商品曝光','总点击','总加购','退款','出单']
PCT_COLS = {'总点击率','总加购率','点击率','加购率','成交率'}
CALC_COLS = PCT_COLS | {'平均流速','总时长'}
READONLY_COLS = {'主播','轮次','环节','时间'}

# ============================================================
# CSS：深底白字 / 白底黑字
# ============================================================
st.markdown("""
<style>
    .stApp{background:linear-gradient(135deg,#1A1B3E 0%,#2D2A5C 50%,#1F1E40 100%)}
    .main>div{background:transparent}
    /* 通用深色背景组件 - 白字 */
    .stApp,.stApp h1,.stApp h2,.stApp h3,.stApp h4,.stApp h5,.stApp h6,
    .stApp p,.stApp span,.stApp div,.stApp label,.stApp li,
    .stMarkdown,.stMarkdown *{color:rgba(255,255,255,0.92)}
    .main-title{font-size:1.9rem;font-weight:700;background:linear-gradient(90deg,#A29BFE,#FDCB6E);-webkit-background-clip:text;-webkit-text-fill-color:transparent;letter-spacing:1px;color:transparent!important}
    .sub-title{color:rgba(255,255,255,0.4)!important;font-size:0.78rem;letter-spacing:2px}
    .glass{background:rgba(255,255,255,0.05);backdrop-filter:blur(16px);-webkit-backdrop-filter:blur(16px);border-radius:16px;border:1px solid rgba(255,255,255,0.08);padding:18px 20px;margin-bottom:14px;box-shadow:0 8px 32px rgba(0,0,0,0.35)}
    .glass-tight{padding:12px 16px}
    .section-title{color:rgba(255,255,255,0.78)!important;font-weight:600;font-size:0.9rem;margin-bottom:8px}
    .badge{display:inline-block;padding:4px 12px;border-radius:20px;font-size:0.75rem;font-weight:600}
    .badge-purple{background:rgba(124,123,255,0.18);color:#A29BFE!important;border:1px solid rgba(124,123,255,0.2)}
    .badge-orange{background:rgba(253,203,110,0.18);color:#FDCB6E!important;border:1px solid rgba(253,203,110,0.2)}
    /* 普通按钮：不透明深底 + 白字（descendant 选择器，兼容 help 提示框包裹） */
    .stButton button{border-radius:10px!important;font-weight:600!important;transition:all 0.3s!important;border:1px solid rgba(255,255,255,0.12)!important;padding:0.45rem 1.2rem!important;color:#fff!important;background:rgba(124,123,255,0.22)!important}
    .stButton button:active{transform:scale(0.96)}
    .stButton button[kind="primary"]{background:linear-gradient(135deg,#7C7BFF,#A29BFE)!important;color:#fff!important;border:none!important;box-shadow:0 4px 20px rgba(124,123,255,0.3)!important}
    .stButton button[kind="primary"]:hover{box-shadow:0 4px 30px rgba(124,123,255,0.5)!important;transform:translateY(-1px)}
    .stButton button:hover{background:rgba(124,123,255,0.40)!important}
    .end-btn button{background:linear-gradient(135deg,#ff6b6b,#ee5a6f)!important;color:#fff!important;border:none!important;font-weight:700!important;box-shadow:0 4px 20px rgba(255,107,107,0.3)!important}
    /* 添加主播自定义组件无需额外样式 */
    /* 输入框样式：白底 + 黑字 + 内嵌步骤按钮 */
    .stNumberInput>div>div>div input{font-size:0.78rem!important;padding:0 8px!important;min-height:0!important;height:26px!important;background:#fff!important;color:#111!important}
    .stNumberInput button{background:rgba(0,0,0,0.04)!important;border:none!important;color:rgba(0,0,0,0.5)!important;font-size:0.72rem!important;padding:0 6px!important;height:26px!important;min-height:0!important}
    .stNumberInput button:hover{background:rgba(0,0,0,0.10)!important;color:#111!important}
    .stNumberInput>div{border:1px solid rgba(0,0,0,0.12)!important;background:#fff!important;border-radius:8px!important;overflow:hidden;box-shadow:none!important}
    .stNumberInput>div:focus-within{border-color:rgba(124,123,255,0.5)!important;box-shadow:0 0 0 2px rgba(124,123,255,0.2)!important}
    .stTextInput input{font-size:0.78rem!important;padding:0 8px!important;height:26px!important;background:#fff!important;color:#111!important}
    .stTextInput>div{border:1px solid rgba(0,0,0,0.12)!important;background:#fff!important;border-radius:8px!important;box-shadow:none!important}
    .stTextInput>div:focus-within{border-color:rgba(124,123,255,0.5)!important;box-shadow:0 0 0 2px rgba(124,123,255,0.2)!important}
    div[data-testid="stNumberInput"]{gap:0!important}
    /* selectbox 同样白底 */
    div[data-baseweb="select"]>div{background:#fff!important;color:#111!important}
    div[data-baseweb="select"] span{color:#111!important}
    /* 强制 hotgrid 容器及所有父级都允许溢出 */
    iframe, iframe *, div:has(> iframe){overflow:visible!important}
    [data-testid="stCustomComponent"], [data-testid="stCustomComponent"] *{overflow:visible!important}
    /* dataframe 左上角/右上角弹出框图标改成黑色 */
    div[data-testid="stDataFrame"] button{color:#1A1B3E!important}
    div[data-testid="stDataFrame"] button svg{fill:#1A1B3E!important}
    div[data-testid="stDataFrame"] [data-testid="stDataFrameToolbar"] button{color:#1A1B3E!important}
    .upload-zone{border:2px dashed rgba(124,123,255,0.3);border-radius:14px;padding:18px 14px;text-align:center;background:rgba(124,123,255,0.04);color:rgba(255,255,255,0.7)!important}
    .divider{height:1px;background:linear-gradient(90deg,transparent,rgba(124,123,255,0.18),transparent);margin:14px 0}
    ::-webkit-scrollbar{width:5px;height:5px}
    ::-webkit-scrollbar-track{background:rgba(255,255,255,0.02)}
    ::-webkit-scrollbar-thumb{background:rgba(124,123,255,0.25);border-radius:8px}
    .stTabs [data-baseweb="tab-list"]{background:transparent!important}
    .stTabs [data-baseweb="tab-list"] button{color:rgba(255,255,255,0.55)!important;background:rgba(255,255,255,0.04)!important;border-radius:8px 8px 0 0!important;padding:8px 14px!important}
    .stTabs [data-baseweb="tab-list"] button[aria-selected="true"]{background:linear-gradient(135deg,rgba(124,123,255,0.3),rgba(124,123,255,0.15))!important;color:#A29BFE!important;font-weight:700!important}
    /* 所有可写输入框：强制白底黑字（加固命中，避免白底白字） */
    input,textarea{color:#111!important}
    div[data-testid="stNumberInput"] input,
    div[data-testid="stTimeInput"] input,
    div[data-testid="stTextInput"] input,
    div[data-testid="stSelectbox"] input,
    div[data-testid="stMultiSelect"] input{background:#fff!important;color:#111!important;font-weight:600}
    /* 数字输入 */
    div[data-testid="stNumberInput"] input{background:rgba(255,255,255,0.95)!important;border:1px solid rgba(0,0,0,0.25)!important;border-radius:8px!important}
    div[data-testid="stNumberInput"] label{color:rgba(255,255,255,0.7)!important;font-size:0.72rem}
    /* 时间输入：白底黑字（含内部文字元素） */
    div[data-testid="stTimeInput"]>div{background:#fff!important;border:1px solid rgba(0,0,0,0.25)!important;border-radius:8px!important}
    div[data-testid="stTimeInput"] input,div[data-testid="stTimeInput"] *{color:#111!important;background:#fff!important;border-color:transparent!important}
    div[data-testid="stTimeInput"] label{color:rgba(255,255,255,0.7)!important;font-size:0.72rem}
    /* text input：白底黑字 */
    div[data-testid="stTextInput"]>div{background:#fff!important;border:1px solid rgba(0,0,0,0.25)!important;border-radius:8px!important}
    div[data-testid="stTextInput"] input{color:#111!important;background:#fff!important;border:none!important;font-weight:600}
    div[data-testid="stTextInput"] label{color:rgba(255,255,255,0.78)!important}
    /* file uploader：深底白字 */
    div[data-testid="stFileUploader"]{color:#fff!important}
    div[data-testid="stFileUploader"] section{background:rgba(124,123,255,0.10)!important;border:2px dashed rgba(124,123,255,0.4)!important;border-radius:14px!important;color:rgba(255,255,255,0.85)!important}
    div[data-testid="stFileUploader"] section *{color:rgba(255,255,255,0.85)!important}
    div[data-testid="stFileUploader"] section button{background:linear-gradient(135deg,#7C7BFF,#A29BFE)!important;color:#fff!important;border:none!important}
    div[data-testid="stFileUploader"] section small{color:rgba(255,255,255,0.55)!important}
    /* 上传区中文化：替换拖拽提示文字 */
    div[data-testid="stFileUploadDropzone"]>span{font-size:0!important}
    div[data-testid="stFileUploadDropzone"]>span::after{content:"拖拽图片到此处，或点击下方按钮选择文件";font-size:0.88rem!important;color:rgba(255,255,255,0.85)!important}
    div[data-testid="stFileUploader"] button{font-size:0!important}
    div[data-testid="stFileUploader"] button::after{content:"📁 选择文件";font-size:0.82rem!important;color:#fff!important}
    div[data-testid="stFileUploadDropzone"] small{font-size:0.68rem!important}
    /* selectbox：白底黑字（可填写）- 加强命中 baseweb 内部所有文字节点 */
    div[data-testid="stSelectbox"] label{color:rgba(255,255,255,0.78)!important;font-size:0.72rem}
    div[data-testid="stSelectbox"] div[data-baseweb="select"],
    div[data-testid="stSelectbox"] div[data-baseweb="select"]>div,
    div[data-testid="stSelectbox"] div[data-baseweb="select"]>div>div,
    div[data-testid="stSelectbox"] div[data-baseweb="select"]>div>div>div,
    div[data-testid="stSelectbox"] div[data-baseweb="select"]>div>div>div>div{color:#111!important;background:#fff!important;border-color:rgba(0,0,0,0.2)!important}
    div[data-testid="stSelectbox"] div[data-baseweb="select"] [class*="SingleValue"],
    div[data-testid="stSelectbox"] div[data-baseweb="select"] [class*="Placeholder"],
    div[data-testid="stSelectbox"] div[data-baseweb="select"] [class*="ValueContainer"],
    div[data-testid="stSelectbox"] div[data-baseweb="select"] [class*="Input"]{color:#111!important;background:transparent!important}
    div[data-testid="stSelectbox"] div[data-baseweb="select"] svg{color:#111!important;fill:#111!important}
    div[data-testid="stSelectbox"] div[role="listbox"] div{color:#111!important;background:rgba(255,255,255,0.95)!important}
    /* text input：白底黑字 - 加强命中 */
    div[data-testid="stTextInput"] input{color:#111!important;background:#fff!important;border:none!important;font-weight:600}
    div[data-testid="stTextInput"] input::placeholder{color:rgba(0,0,0,0.45)!important;font-weight:400}
    div[data-testid="stTextInput"]>div>div>div{color:#111!important;background:#fff!important}
    /* multiselect：白底黑字（可填写） */
    div[data-testid="stMultiSelect"] label{color:rgba(255,255,255,0.78)!important;font-size:0.72rem}
    div[data-testid="stMultiSelect"] div[data-baseweb="select"]>div{color:#111!important;background:#fff!important;border-color:rgba(0,0,0,0.2)!important}
    div[data-testid="stMultiSelect"] div[data-baseweb="select"] svg{color:#111!important;fill:#111!important}
    /* radio：选项文字白字（深色背景上可见） */
    div[data-testid="stRadio"] label, div[data-testid="stRadio"] label span{color:rgba(255,255,255,0.88)!important}
    /* caption：图表说明白字 */
    .stCaption, div[data-testid="stCaption"], div[data-testid="stCaptionContainer"], .stCaption *, div[data-testid="stCaption"] *, div[data-testid="stCaptionContainer"] *{color:rgba(255,255,255,0.65)!important}
    /* dataframe / data_editor：深底白字标题，编辑区白底黑字 */
    div[data-testid="stDataFrame"] th{background:rgba(124,123,255,0.15)!important;color:#A29BFE!important;font-weight:600!important}
    div[data-testid="stDataFrame"] td{color:#fff!important;background:rgba(255,255,255,0.03)!important}
    div[data-testid="stDataEditor"] input{color:#111!important;background:rgba(255,255,255,0.95)!important;font-weight:600}
    .finished-tag{background:rgba(72,219,128,0.12);color:#48db80!important;border:1px solid rgba(72,219,128,0.2);padding:3px 10px;border-radius:12px;font-size:0.7rem;font-weight:600}
    .round-tag{display:inline-block;background:linear-gradient(135deg,#7C7BFF,#A29BFE);color:#fff!important;padding:6px 16px;border-radius:14px;font-weight:700;font-size:0.85rem;letter-spacing:1px;box-shadow:0 4px 14px rgba(124,123,255,0.3)}
    .undo-btn button{background:rgba(253,203,110,0.12)!important;color:#FDCB6E!important;border:1px solid rgba(253,203,110,0.3)!important;font-weight:600!important}
    .undo-btn button:hover{background:rgba(253,203,110,0.2)!important;border-color:rgba(253,203,110,0.5)!important}
    .highlight-card{background:linear-gradient(135deg,rgba(124,123,255,0.15),rgba(162,155,254,0.08));border:1px solid rgba(124,123,255,0.3);border-radius:14px;padding:14px 18px;margin-bottom:12px}
</style>
""", unsafe_allow_html=True)

# ============================================================
# Session State
# ============================================================
DEFAULTS = {
    'data_rows': [],
    'current_session': '',
    'current_round': 1,
    'current_phase': '干货',
    'first_round': True,
    'first_dry_start': dtime(7, 0),
    'first_dry_duration': 40,
    'phase_start_time': None,
    'phase_end_time': None,
    'last_end_time': None,
    'last_cumulative': None,
    'uploaded_images': [],
    'phase_data': {},
    'edit_data': {},
    'hosts': ['邹志俐', '田丽丽', '沈晓书'],
    'current_host': '邹志俐',
    'finished_sessions': [],
    'session_round_totals': {},
    'round_label_override': {},  # 用户手动改的「第N场·第M轮」标签
    'prev_round_totals': {},      # 上轮累计值，用于曝光等增量计算
    'row_id_counter': 1000,
    'history': [],  # 操作栈：每步存 data_rows + finished 的快照
    'auto_save': True,
}
for k, v in DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v

# 自动恢复：刷新不丢数据
if not st.session_state.get('_loaded'):
    st.session_state._loaded = True
    auto_load()

# ============================================================
# 历史/Undo
# ============================================================
def push_history():
    """每次操作前调用，记录当前状态"""
    snap = {
        'data_rows': [dict(r) for r in st.session_state.data_rows],
        'finished_sessions': [dict(s) for s in st.session_state.finished_sessions],
        'current_round': st.session_state.current_round,
        'current_phase': st.session_state.current_phase,
        'first_round': st.session_state.first_round,
        'last_end_time': st.session_state.last_end_time,
        'last_cumulative': dict(st.session_state.last_cumulative) if st.session_state.last_cumulative else None,
        'session_round_totals': {k: dict(v) for k, v in st.session_state.session_round_totals.items()},
    }
    st.session_state.history.append(snap)
    # 只保留最近 20 步
    if len(st.session_state.history) > 20:
        st.session_state.history = st.session_state.history[-20:]

def do_undo():
    if not st.session_state.history:
        return False, '无可撤回的操作'
    snap = st.session_state.history.pop()
    st.session_state.data_rows = snap['data_rows']
    st.session_state.finished_sessions = snap['finished_sessions']
    st.session_state.current_round = snap['current_round']
    st.session_state.current_phase = snap['current_phase']
    st.session_state.first_round = snap['first_round']
    st.session_state.last_end_time = snap['last_end_time']
    st.session_state.last_cumulative = snap['last_cumulative']
    st.session_state.session_round_totals = snap['session_round_totals']
    return True, '已撤回'

# ============================================================

def merge_images_data(images_data):
    merged, conflicts = {}, {}
    for img_data in images_data:
        for k, v in img_data.items():
            if k.startswith('_'): continue
            if k in merged:
                if merged[k] != v:
                    conflicts.setdefault(k, set()).add(merged[k])
                    conflicts[k].add(v)
                    merged[k] = v
            else:
                merged[k] = v
    return merged, conflicts

def check_anomalies(rows):
    """检查每行数据的异常，返回 [(行标签, 异常描述)]"""
    warns = []
    for r in rows:
        tag = f"R{r.get('轮次','?')} {r.get('环节','?')}"
        s, t, m, z, j, d, c, cr = 0, 0, 0, 0, 0, 0, 0, 0
        try: s = int(r.get('总曝光', 0) or 0)
        except: pass
        try: t = int(r.get('总进房', 0) or 0)
        except: pass
        try: m = int(r.get('总商品曝光', 0) or 0)
        except: pass
        try: z = int(r.get('总点击', 0) or 0)
        except: pass
        try: j = int(r.get('总加购', 0) or 0)
        except: pass
        try: d = int(r.get('出单', 0) or 0)
        except: pass
        try: c = int(r.get('成交', 0) or 0)
        except: pass
        try: cr = int(r.get('成交率', 0) or 0)
        except: pass
        # 异常规则
        if t > s > 0: warns.append((tag, f"总进房({t}) > 总曝光({s}) — 异常"))
        if z > m > 0: warns.append((tag, f"总点击({z}) > 总商品曝光({m}) — 异常"))
        if j > z > 0 and j > m: warns.append((tag, f"总加购({j}) > 总点击({z}) — 异常"))
        if cr > 100: warns.append((tag, f"成交率({cr}%) > 100% — 请检查"))
        if cr < 0: warns.append((tag, f"成交率为负数({cr}%) — 数据异常"))
        if c > j > 0: warns.append((tag, f"成交({c}) > 总加购({j}) — 成交率可能>100%"))
    return warns

def calc_rates(row):
    """计算比率类指标 + 累计平均流速。成交率 = 成交÷加购。
    干货行：成交=出单（占位联动）；售卖行：成交独立。"""
    out = dict(row)
    def pct(a, b):
        try: av, bv = float(a), float(b)
        except: return 0
        return round(av/bv*100) if bv>0 else 0
    # 干货行 成交 = 出单（占位联动）
    if out.get('环节') == '干货':
        out['成交'] = out.get('出单', 0)
    out['总点击率'] = pct(out.get('总点击',0), out.get('总商品曝光',0))
    out['总加购率'] = pct(out.get('总加购',0), out.get('总点击',0))
    out['点击率'] = pct(out.get('点击',0), out.get('商品曝光',0))
    out['加购率'] = pct(out.get('加购',0), out.get('点击',0))
    out['成交率'] = pct(out.get('成交',0), out.get('加购',0))  # v5.4: 成交÷加购
    return out

def sync_meta_to_last_row():
    """左侧控件变化时自动同步到上一行数据。
    仅当当前轮次+环节与上一行相同时才联动，否则是新数据。"""
    if not st.session_state.data_rows: return
    if st.session_state.get('_editing_new_row'): return
    last = st.session_state.data_rows[-1]
    # 轮次或环节变了 → 用户开始新一轮/环节 → 不联动
    if str(last.get('轮次')) != str(st.session_state.current_round): return
    if last.get('环节') != st.session_state.current_phase: return
    changed = False
    # 场次
    if last.get('场次') != st.session_state.current_session:
        last['场次'] = st.session_state.current_session; changed = True
    # 主播
    if last.get('主播') != st.session_state.current_host:
        last['主播'] = st.session_state.current_host; changed = True
    # 轮次
    if str(last.get('轮次')) != str(st.session_state.current_round):
        last['轮次'] = st.session_state.current_round; changed = True
    # 时间
    ps = st.session_state.phase_start_time
    pe = st.session_state.phase_end_time
    phase = st.session_state.current_phase
    if ps and pe:
        t_str = f'{fmt_time(ps)}-{fmt_time(pe)}'
        dur = diff_minutes(ps, pe)
    elif st.session_state.first_round and phase == '干货':
        fd = st.session_state.first_dry_start
        t_str = fd.strftime('%H:%M') if hasattr(fd, 'strftime') else ''
        dur = st.session_state.first_dry_duration
    else:
        t_str = last.get('时间', '')
        dur = last.get('时长', 0)
    if t_str and last.get('时间') != t_str:
        last['时间'] = t_str; changed = True
    if dur > 0 and last.get('时长') != dur:
        last['时长'] = dur; changed = True
    # 总时长 = 时长（同值）
    if last.get('总时长') != dur:
        last['总时长'] = dur; changed = True
    if changed:
        st.session_state.data_rows[-1] = calc_rates(last)
        auto_save()

def fmt_time(t):
    if t is None: return ''
    if isinstance(t, datetime): return t.strftime('%H:%M')
    if isinstance(t, dtime): return t.strftime('%H:%M')
    return str(t)

def fmt_range(s, e):
    return f"{fmt_time(s)}-{fmt_time(e)}" if s and e else ''

def diff_minutes(t1, t2):
    if t1 is None or t2 is None: return 0
    delta = (t2 - t1).total_seconds()/60
    if delta < 0: delta += 24*60
    return round(delta, 1)

def combine_dt(t, base=None):
    if t is None: return None
    if isinstance(t, datetime): return t
    return datetime.combine(base or datetime.now().date(), t)

# ============================================================
# 保存
# ============================================================
def save_current_phase():
    st.session_state._session_ended = False  # 新数据写入，标记未结束
    data = st.session_state.edit_data or st.session_state.phase_data
    if not data and not (st.session_state.first_round and st.session_state.current_phase=='干货'):
        return False, '请先填数据或上传截图'
    phase = st.session_state.current_phase
    rnd = st.session_state.current_round
    host = st.session_state.current_host
    # 首次保存自动设场次名 = 本场第一轮第一环节开始时间
    if not st.session_state.current_session:
        fd = st.session_state.first_dry_start
        now = datetime.now()
        if hasattr(fd, 'hour'):
            t = fd.strftime('%H:%M')
        else:
            t = now.strftime('%H:%M')
        st.session_state.current_session = f'{now.month}/{now.day} {t}'
    session = st.session_state.current_session
    today = datetime.now().date()

    if st.session_state.first_round and phase=='干货':
        start_dt = combine_dt(st.session_state.first_dry_start, today)
        duration = float(st.session_state.first_dry_duration)
        end_dt = start_dt + pd.Timedelta(minutes=duration).to_pytimedelta()
    else:
        start_dt = st.session_state.phase_start_time or st.session_state.last_end_time or datetime.now()
        end_dt = st.session_state.phase_end_time or datetime.now()
        duration = diff_minutes(start_dt, end_dt)

    last = st.session_state.last_cumulative or {}
    def gv(k): return int(data.get(k,0))
    def diffv(k): return gv(k) - int(last.get(k, 0))

    is_first = st.session_state.first_round and phase=='干货'
    is_sell = (phase == '售卖')

    # 整轮累计列（带"总"的列）：始终取累计值 gv
    总曝光 = gv('总曝光'); 总进房 = gv('总进房')
    总商品曝光 = gv('总商品曝光'); 总点击 = gv('总点击'); 总加购 = gv('总加购')

    # v5.4: 曝光/进房/商品曝光/点击/加购（不带"总"）= 跨轮增量差值
    # 第1轮 = 本轮累计；第2+轮 = 本轮累计 - 上轮累计
    if is_first or not st.session_state.prev_round_totals:
        曝光 = 总曝光; 进房 = 总进房
        商品曝光 = 总商品曝光; 点击 = 总点击; 加购 = 总加购
    else:
        p = st.session_state.prev_round_totals
        曝光 = 总曝光 - p.get('总曝光', 0)
        进房 = 总进房 - p.get('总进房', 0)
        商品曝光 = 总商品曝光 - p.get('总商品曝光', 0)
        点击 = 总点击 - p.get('总点击', 0)
        加购 = 总加购 - p.get('总加购', 0)

    # v5.4: 出单/退款 = 同级计算——仅第1轮干货用原始值，之后全用增量差值
    if is_first:
        出单 = gv('出单')
        退款 = gv('退款')
    else:
        出单 = diffv('出单')
        退款 = diffv('退款')

    # 成交：一轮一行。干货先填干货出单占位，售卖填总出单(售卖上传原始值)
    成交 = gv('出单')

    row = {
        '_row_id': st.session_state.row_id_counter,
        '场次': session, '主播': host, '轮次': rnd, '环节': phase,
        '时间': fmt_range(start_dt, end_dt), '时长': duration,
        '总时长': 0, '出单': 出单, '退款': 退款,
        '总曝光': 总曝光, '总进房': 总进房,
        '曝光': 曝光, '进房': 进房, '平均流速': 0,
        '总商品曝光': 总商品曝光, '总点击': 总点击, '总加购': 总加购,
        '总点击率': 0, '总加购率': 0,
        '商品曝光': 商品曝光, '点击': 点击, '加购': 加购,
        '点击率': 0, '加购率': 0,
        '成交': 成交, '成交率': 0,
    }
    row = calc_rates(row)
    st.session_state.row_id_counter += 1

    # 累计本轮时长（用于整轮 总时长 / 平均流速）
    rd = st.session_state.session_round_totals
    rd.setdefault(rnd, {'时长':0})
    rd[rnd]['时长'] += duration
    total_dur = rd[rnd]['时长']
    # 平均流速 = 整轮总进房(row['总进房']) / 本轮总时长
    avg_speed = round(row['总进房'] / total_dur, 1) if total_dur>0 else 0
    row['总时长'] = total_dur
    row['平均流速'] = avg_speed

    push_history()
    st.session_state.data_rows.append(row)

    # 用录入的累计值（gv）记录 last_cumulative（不能用 row，避免环节增量干扰）
    st.session_state.last_cumulative = {
        '总曝光': 总曝光, '总进房': 总进房,
        '总商品曝光': 总商品曝光, '总点击': 总点击,
        '总加购': 总加购, '退款': 退款, '出单': 出单
    }
    st.session_state.last_end_time = end_dt

    # 售卖保存成功后：把整轮累计列（保留 5 个分两行列）回填到同轮干货行
    if is_sell:
        MERGE_COLS = ['总曝光','总进房','曝光','进房','总商品曝光','总点击','总加购',
                      '总点击率','总加购率','商品曝光','点击','加购','点击率','加购率',
                      '成交率','总时长','平均流速','成交']  # v5.4: 加成交
        for idx, r in enumerate(st.session_state.data_rows):
            if r['场次']==session and r['轮次']==rnd and r['环节']=='干货':
                for c in MERGE_COLS:
                    r[c] = row[c]
                st.session_state.data_rows[idx] = r
                break
        # v5.4: 更新跨轮累计基准
        st.session_state.prev_round_totals = {
            '总曝光': 总曝光, '总进房': 总进房,
            '总商品曝光': 总商品曝光, '总点击': 总点击, '总加购': 总加购
        }

    if phase == '售卖':
        st.session_state.current_round += 1
        st.session_state.current_phase = '干货'
        st.session_state.first_round = False
    else:
        st.session_state.current_phase = '售卖'

    # 切换环节：自动清空上传图
    st.session_state.uploaded_images = []
    st.session_state.phase_data = {}
    st.session_state.edit_data = {}
    st.session_state.phase_start_time = None
    st.session_state.phase_end_time = None
    auto_save()
    st.session_state._editing_new_row = False  # 已保存，允许联动
    st.session_state._edit_data_snapshot = dict(st.session_state.edit_data)  # 保存快照
    return True, '已保存'

def end_current_session():
    if not st.session_state.data_rows: return False, '没有数据'
    snap = {
        'name': st.session_state.current_session,
        'host': st.session_state.current_host,
        'end_time': datetime.now(),
        'rows': [dict(r) for r in st.session_state.data_rows]
    }
    found = False
    for i, s in enumerate(st.session_state.finished_sessions):
        if s['name'] == snap['name']:
            st.session_state.finished_sessions[i] = snap; found = True; break
    if not found: st.session_state.finished_sessions.append(snap)
    # 本地文件
    try:
        d = ensure_sessions_dir()
        ts = datetime.now().strftime('%Y-%m-%d_%H-%M')
        safe_name = re.sub(r'[\\/:*?"<>|]', '_', str(snap['name']))
        path = os.path.join(d, f"{safe_name}_{ts}.xlsx")
        with open(path, 'wb') as f:
            f.write(to_excel_bytes(snap['rows']))
    except Exception:
        pass
    # 云端：转为 JSON 安全的副本
    try:
        cloud_snap = {
            'name': str(snap['name']),
            'host': str(snap.get('host', '')),
            'end_time': snap['end_time'].isoformat() if hasattr(snap['end_time'], 'isoformat') else str(snap['end_time']),
            'rows': [dict(r) for r in snap['rows']]
        }
        db_save(f"session_{snap['name']}_{datetime.now().strftime('%Y%m%d%H%M')}", cloud_snap)
    except Exception as e:
        print(f'cloud save err: {e}')
    st.session_state.data_rows = []
    st.session_state.current_round = 1
    st.session_state.current_phase = '干货'
    st.session_state.first_round = True
    st.session_state.last_end_time = None
    st.session_state.last_cumulative = None
    st.session_state.uploaded_images = []
    st.session_state.phase_data = {}
    st.session_state.edit_data = {}
    st.session_state.phase_start_time = None
    st.session_state.phase_end_time = None
    st.session_state.session_round_totals = {}
    st.session_state.prev_round_totals = {}
    n = len(st.session_state.finished_sessions) + 1
    st.session_state.current_session = ''
    st.session_state._session_ended = True  # 防止刷新后恢复旧数据
    # 清除自动保存文件
    try:
        asp = auto_save_path()
        if os.path.exists(asp): os.remove(asp)
    except Exception: pass
    return True, snap['name']

def new_session():
    st.session_state.current_session = ''
    for k in ['data_rows','last_end_time','last_cumulative','uploaded_images','phase_data','edit_data','phase_start_time','phase_end_time','session_round_totals']:
        st.session_state[k] = None if k != 'data_rows' else []
    st.session_state.data_rows = []
    st.session_state.current_round = 1
    st.session_state.current_phase = '干货'
    st.session_state.first_round = True
    st.session_state.session_round_totals = {}
    st.session_state.prev_round_totals = {}

@st.cache_data
def to_excel_bytes(rows):
    df = pd.DataFrame(rows)
    cs = ['场次'] + COLUMNS_ORDER
    cs = [c for c in cs if c in df.columns]
    df = df[cs]
    out = BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as w: df.to_excel(w, sheet_name='直播数据', index=False)
    return out.getvalue()

def sync_editor_to_rows(edited_df, display_keys, current_rows):
    """用复合主键 (场次,轮次,环节) 定位行，无需 ID 列。"""
    edited_keys = set(zip(edited_df['场次'].astype(str), edited_df['轮次'].astype(str), edited_df['环节'].astype(str)))
    deleted = set(display_keys) - edited_keys
    # 更新：按主键匹配
    for _, erow in edited_df.iterrows():
        key = (str(erow['场次']), str(erow['轮次']), str(erow['环节']))
        if key in deleted:
            continue
        for r in current_rows:
            if (str(r.get('场次')), str(r.get('轮次')), str(r.get('环节'))) == key:
                for c in EDITABLE_COLS:
                    if c in erow and c in r and pd.notna(erow[c]):
                        r[c] = erow[c]
                updated = calc_rates(r)
                for k in CALC_COLS:
                    if k in updated: r[k] = updated[k]
                break
    # 删除：主键在原显示中、编辑后消失的
    if deleted:
        current_rows = [r for r in current_rows
                        if (str(r.get('场次')), str(r.get('轮次')), str(r.get('环节'))) not in deleted]
    return current_rows

EDITABLE_COLS = [c for c in COLUMNS_ORDER if c not in CALC_COLS and c not in READONLY_COLS]

def sync_grid_to_rows(grid_rows, current_rows, columns):
    """把 Handsontable 回写的数据同步回 data_rows（复合主键定位；支持增/删/改）。"""
    try:
        return _sync_grid_to_rows(grid_rows, current_rows, columns)
    except Exception as e:
        st.warning(f'数据同步出错（已保留原数据）: {e}')
        return current_rows

def _sync_grid_to_rows(grid_rows, current_rows, columns):
    key_of = lambda r: (str(r.get('场次')), str(r.get('轮次')), str(r.get('环节')))
    # 率列去% + 数字字段强制转int
    rate_cols = [c for c in PCT_COLS if c in columns]
    num_cols = [c for c in columns if c not in READONLY_COLS and c not in PCT_COLS and c not in ['场次','主播','时间']]
    for rec in grid_rows:
        for c in rate_cols:
            v = rec.get(c, '')
            if isinstance(v, str) and v.endswith('%'): rec[c] = int(v.replace('%', ''))
            else:
                try: rec[c] = int(float(v))
                except: pass
        for c in num_cols:
            v = rec.get(c, '')
            try: rec[c] = int(float(v)) if v != '' and v is not None else 0
            except: pass
    # ---
    existing = {key_of(r): r for r in current_rows}
    new_keys = [key_of(r) for r in grid_rows]
    result = []
    for rec in grid_rows:
        k = key_of(rec)
        if k in existing:
            for c in EDITABLE_COLS:
                if c in rec and c in existing[k] and pd.notna(rec[c]):
                    existing[k][c] = rec[c]
            existing[k] = calc_rates(existing[k])
            result.append(existing[k])
        else:
            nr = {c: (rec.get(c, 0) if c in rec else 0) for c in columns}
            nr = calc_rates(nr)
            result.append(nr)
    # 删除：原 current_rows 中存在、但 grid 中消失的关键字行
    for k in [key_of(r) for r in current_rows]:
        if k not in new_keys:
            continue  # 跳过 -> 不加入 result，即删除
    return result

# ============================================================
# 头部
# ============================================================
st.markdown("""
<div style="text-align:center;padding:6px 0 0 0">
    <span class="main-title">🎯 直播记录系统</span>
</div>
<div class="divider"></div>
""", unsafe_allow_html=True)


# ============================================================
# 顶部控制栏（原左侧内容 → 标题下方横排）
# ============================================================

# 用户信息 + 共享开关（仅DB模式）
if HAS_DB:
    uc1, uc2, uc3 = st.columns([2, 1, 1])
    with uc1:
        st.markdown(f'<span style="color:rgba(255,255,255,0.5);font-size:0.7rem">👤 {st.session_state.user}</span>', unsafe_allow_html=True)
    with uc2:
        if st.session_state.get('team', ''):
            st.markdown(f'<span style="color:rgba(124,123,255,0.6);font-size:0.7rem">🏷 团队 {st.session_state.team}</span>', unsafe_allow_html=True)
    with uc3:
        was_share = st.session_state.get('share', False)
        new_share = st.checkbox('📤 共享我的数据', value=was_share, key='ui_share')
        if new_share != was_share:
            st.session_state.share = new_share
            # 重新保存当前数据以反映共享状态
            if st.session_state.data_rows:
                auto_save()
            st.rerun()

# Row 1: 轮次(含场次) | 环节 | 主播 | 添加主播
r1c1, r1c2, r1c3, r1c4 = st.columns([1.6, 1.0, 1.6, 1.6])
with r1c1:
    auto_label = f"{st.session_state.current_session or '·'} · 第{st.session_state.current_round}轮"
    label_key = f"{st.session_state.current_session}_{st.session_state.current_round}"
    default_lbl = st.session_state.round_label_override.get(label_key, auto_label)
    new_lbl = st.text_input('轮次', value=default_lbl, key=f'rl_{label_key}', label_visibility='collapsed')
    if new_lbl != default_lbl: st.session_state.round_label_override[label_key] = new_lbl

with r1c2:
    st.markdown(f'<div class="highlight-card" style="margin:0;padding:4px 8px;font-size:0.7rem"><span class="badge badge-{"orange" if st.session_state.current_phase=="售卖" else "purple"}">{"📦" if st.session_state.current_phase=="售卖" else "📚"} {st.session_state.current_phase}</span></div>', unsafe_allow_html=True)
    phase_opts = ['干货', '售卖']
    new_phase = st.radio('', phase_opts, index=phase_opts.index(st.session_state.current_phase), key='phase_sel', horizontal=True, label_visibility='collapsed')
    if new_phase != st.session_state.current_phase:
        st.session_state.uploaded_images = []; st.session_state.phase_data = {}; st.session_state.edit_data = {}
        st.session_state.phase_start_time = None; st.session_state.phase_end_time = None
        st.session_state.current_phase = new_phase

with r1c3:
    sel = st.selectbox('主播', st.session_state.hosts,
        index=st.session_state.hosts.index(st.session_state.current_host) if st.session_state.current_host in st.session_state.hosts else 0,
        key='hs', label_visibility='collapsed')
    st.session_state.current_host = sel

with r1c4:
    # 添加主播：Streamlit 输入 + 按钮，CSS 一体白框
    nh = st.text_input('新主播名', key='nhp', label_visibility='collapsed', placeholder='新主播')
    if st.button('添加', key='ahb'):
        name = (nh or '').strip()
        if name and name not in st.session_state.hosts:
            st.session_state.hosts.append(name)
        st.session_state.current_host = name if name else st.session_state.current_host
        st.rerun()

# Row 2: 时间
r2c1, r2c2 = st.columns([1.5, 2.5])
today = datetime.now().date()
with r2c1:
    st.markdown('<div style="color:rgba(255,255,255,0.6);font-size:0.68rem;margin:4px 0 2px">⏱ 时间</div>', unsafe_allow_html=True)
    if st.session_state.first_round and st.session_state.current_phase == '干货':
        r2a, r2b = st.columns([1, 1])
        with r2a:
            h = st.number_input('时', min_value=0, max_value=23, value=int(st.session_state.first_dry_start.hour), step=1, key='fd_h', label_visibility='collapsed')
            st.session_state.first_dry_start = dtime(int(h), int(st.session_state.first_dry_start.minute))
        with r2b:
            m = st.number_input('分', min_value=0, max_value=59, value=int(st.session_state.first_dry_start.minute), step=1, key='fd_m', label_visibility='collapsed')
            st.session_state.first_dry_start = dtime(int(h), int(m))
        r2c, r2d = st.columns([1, 1])
        with r2c:
            d = st.number_input('时长(分)', min_value=1, max_value=600, value=int(st.session_state.first_dry_duration), step=1, key='fd_d', label_visibility='collapsed')
            st.session_state.first_dry_duration = int(d)
    else:
        def_start = st.session_state.phase_start_time or st.session_state.last_end_time or datetime.now()
        stv = def_start.time() if isinstance(def_start, datetime) else def_start
        r2a, r2b = st.columns([1, 1])
        with r2a: psh = st.number_input('开始·时', min_value=0, max_value=23, value=int(stv.hour) if stv else 0, step=1, key='ps_h', label_visibility='collapsed')
        with r2b: psm = st.number_input('开始·分', min_value=0, max_value=59, value=int(stv.minute) if stv else 0, step=1, key='ps_m', label_visibility='collapsed')
        st.session_state.phase_start_time = combine_dt(dtime(int(psh), int(psm)), today)
        def_end = st.session_state.phase_end_time or datetime.now()
        etv = def_end.time() if isinstance(def_end, datetime) else def_end
        r2c, r2d = st.columns([1, 1])
        with r2c: peh = st.number_input('结束·时', min_value=0, max_value=23, value=int(etv.hour) if etv else 0, step=1, key='pe_h', label_visibility='collapsed')
        with r2d: pem = st.number_input('结束·分', min_value=0, max_value=59, value=int(etv.minute) if etv else 0, step=1, key='pe_m', label_visibility='collapsed')
        st.session_state.phase_end_time = combine_dt(dtime(int(peh), int(pem)), today)

# Row 3: 数据确认 7字段一行
st.markdown('<div style="color:rgba(255,255,255,0.6);font-size:0.68rem;margin:4px 0 2px">📝 数据确认</div>', unsafe_allow_html=True)
merged = st.session_state.phase_data
ed = {}
all_fields = OCR_FIELDS
cols = st.columns(7)
for i, field in enumerate(all_fields):
    with cols[i]:
        default = merged.get(field) if merged and field in merged else st.session_state.edit_data.get(field, 0)
        ed[field] = st.number_input(field, value=int(default) if default else 0, min_value=0, step=1,
            key=f'e_{field}_{st.session_state.current_round}_{st.session_state.current_phase}', format='%d', label_visibility='visible')
st.session_state.edit_data = ed
# 检测是否开始填写新数据：和前一次保存的快照比较
snap = st.session_state.get('_edit_data_snapshot', {})
if any(ed.get(k, 0) != snap.get(k, 0) for k in OCR_FIELDS if ed.get(k, 0)):
    st.session_state._editing_new_row = True
if st.session_state.uploaded_images:
    m2, conflicts = merge_images_data(st.session_state.uploaded_images)
    if conflicts:
        with st.expander(f'⚠ {len(conflicts)} 冲突', expanded=False):
            for key, vals in conflicts.items():
                opts = sorted({int(v) for v in vals if isinstance(v,(int,float))})
                choice = st.radio(key, options=opts, key=f'cf_{key}_{st.session_state.current_round}', horizontal=True, label_visibility='collapsed')
                if choice is not None: m2[key] = int(choice); st.session_state.phase_data[key] = int(choice)

# Row 5: 保存 / 清空 / 结束本场
st.markdown('<div style="margin-top:4px"></div>', unsafe_allow_html=True)
b1, b2, b3 = st.columns([1, 1, 1])
with b1:
    if st.button('✅ 保存', type='primary', use_container_width=True, key='sb'):
        ok, msg = save_current_phase()
        if ok: st.success(f'✅ {msg}'); st.rerun()
        else: st.error(f'❌ {msg}')
with b2:
    if st.button('🔄 清空', use_container_width=True, key='cb'):
        st.session_state.uploaded_images = []; st.session_state.phase_data = {}; st.session_state.edit_data = {}
        st.session_state.phase_start_time = None; st.session_state.phase_end_time = None; st.rerun()
with b3:
    st.markdown('<div class="end-btn">', unsafe_allow_html=True)
    if st.button('🏁 结束本场', use_container_width=True, key='esb'):
        ok, info = end_current_session()
        if ok: st.success(f'✅ {info} 已存档'); st.rerun()
        else: st.warning(f'⚠️ {info}')
    st.markdown('</div>', unsafe_allow_html=True)

if st.session_state.finished_sessions:
    with st.expander(f'📚 已结束 ({len(st.session_state.finished_sessions)} 场)', expanded=False):
        for i, fs in enumerate(st.session_state.finished_sessions):
            c1, c2 = st.columns([8, 1])
            with c1: st.markdown(f'<span style="color:rgba(255,255,255,0.75);font-size:0.78rem">{fs["name"]} · {len(fs["rows"])} 行</span>', unsafe_allow_html=True)
            with c2:
                if st.button('🗑', key=f'df_{i}', help='删除'): st.session_state.finished_sessions.pop(i); st.rerun()

st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

# 联动：左侧控件变化 → 自动更新数据行
sync_meta_to_last_row()

# ============================================================
# 主内容区：数据表格 / 数据对比分析 二选一
# ============================================================
main_tab = st.radio('', ['📋 数据表格', '📊 数据对比分析'], key='main_tab', horizontal=True, label_visibility='collapsed')

# 拼接全量数据（两个模式都要用，提前构建）
all_sources = []
for fs in st.session_state.finished_sessions:
    for r in fs['rows']: all_sources.append(dict(r, 场次=fs['name']))
for r in st.session_state.data_rows: all_sources.append(dict(r))
df_all = pd.DataFrame(all_sources) if all_sources else pd.DataFrame(columns=['场次']+COLUMNS_ORDER)

if '数据表格' in main_tab:
    
    # ---- 数据表格区域 ----
    # 三列：当前数据 | 下载按钮 | 往期查询（下载按钮与其他按钮同高同宽）
    if 'vmode' not in st.session_state: st.session_state.vmode = '📝 当前数据'
    view_mode = st.session_state.vmode
    vcol1, vcol2, vcol3 = st.columns([3, 1.2, 3])
    with vcol1:
        if st.button('📝 当前数据', key='vmode_curr', use_container_width=True,
                     type='primary' if view_mode == '📝 当前数据' else 'secondary'):
            st.session_state.vmode = '📝 当前数据'; st.rerun()
    with vcol2:
        if not df_all.empty or st.session_state.data_rows:
            bts = to_excel_bytes(st.session_state.data_rows)
            st.download_button('⬇ 下载', data=bts, file_name=f'直播数据_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', key='dl_cur',
                type='primary', use_container_width=True)
    with vcol3:
        if st.button('📚 往期查询', key='vmode_hist', use_container_width=True,
                     type='primary' if view_mode == '📚 往期查询' else 'secondary'):
            st.session_state.vmode = '📚 往期查询'; st.rerun()
    
    if view_mode == '📚 往期查询':
        d = ensure_sessions_dir()
        files = sorted([f for f in os.listdir(d) if f.endswith('.xlsx')], reverse=True)
        if not files:
            st.markdown('<div style="text-align:center;padding:36px;color:rgba(255,255,255,0.3)">📭 暂无往期数据，先「结束本场」存档</div>', unsafe_allow_html=True)
        else:
            sel_file = st.selectbox('选择往期文件', files, key='pf', label_visibility='collapsed')
            fpath = os.path.join(d, sel_file)
            pc1, pc2 = st.columns([7, 1])
            with pc1:
                if st.button('👁 查看', key='pv', use_container_width=True):
                    try:
                        dfp = pd.read_excel(fpath)
                        st.dataframe(dfp, use_container_width=True, hide_index=True)
                    except Exception as e:
                        st.error(str(e))
            with pc2:
                try:
                    with open(fpath, 'rb') as f: bts = f.read()
                    st.download_button('📥', data=bts, file_name=sel_file,
                                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                       key='pdl')
                except Exception:
                    pass
    else:
        # 当前数据：Handsontable 真网格（增删行列 / 合并 / 标黄 / 主键锁定 / 回写）
        if not st.session_state.data_rows:
            st.markdown('<div style="text-align:center;padding:60px 20px;color:rgba(255,255,255,0.15)"><div style="font-size:3.5rem;margin-bottom:10px">📭</div><div>暂无当前数据，先在左侧录入</div></div>', unsafe_allow_html=True)
        else:
            show = pd.DataFrame(st.session_state.data_rows)
            show['_s'] = show.get('环节', pd.Series()).map({'干货':0,'售卖':1}).fillna(2)
            show = show.sort_values(['场次','轮次','_s']).drop('_s', axis=1)
            disp_cols = [c for c in ['场次']+COLUMNS_ORDER if c in show.columns]
            show_disp = show[disp_cols].copy()
            # v5.4: 率列加 % 后缀，供 hotgrid 和原生编辑器显示
            rate_cols = [c for c in PCT_COLS if c in show_disp.columns]
            for c in rate_cols:
                show_disp[c] = show_disp[c].apply(lambda x: f"{int(x)}%" if pd.notna(x) else '')
            records = show_disp.to_dict('records')
            SPLIT_COLS = ['环节','时间','时长','出单','退款']  # v5.4: 5列，成交移出、一轮一行
            READ_ONLY = ['场次','轮次','环节','时间']
            # v5.4: 本场数据汇总（单独显示在网格下方）
            rows_raw = st.session_state.data_rows
            summary_row = None
            if rows_raw:
                summary_row = {c: '/' for c in disp_cols}
                summary_row['场次'] = '📋 本场数据汇总'
                t_first = str(rows_raw[0].get('时间', '')).split('-')
                t_last = str(rows_raw[-1].get('时间', '')).split('-')
                t_start = t_first[0].strip() if t_first else ''
                t_end = t_last[-1].strip() if len(t_last) > 1 else (t_last[0].strip() if t_last else '')
                def tm(s):
                    try: p=s.strip().split(':'); return int(p[0])*60+int(p[1])
                    except: return 0
                dur = max(0, tm(t_end)-tm(t_start))
                summary_row['时间'] = f'{t_start}-{t_end}' if t_start else '/'
                summary_row['时长'] = dur if dur>0 else '/'
                summary_row['总时长'] = dur if dur>0 else '/'
                summary_row['主播'] = st.session_state.current_host
                # 出单/退款：汇总 = 干货出单 + 售卖出单（来自 OCR 录入值）
                summary_row['出单'] = sum(int(r.get('出单',0) or 0) for r in rows_raw)
                summary_row['退款'] = sum(int(r.get('退款',0) or 0) for r in rows_raw)
                # 成交：先强制按 calc_rates 重算每行后求和（确保干货占位被刷新）
                for r in rows_raw:
                    r2 = calc_rates(r)
                    for k in r2: r[k] = r2[k]
                summary_row['成交'] = sum(int(r.get('成交',0) or 0) for r in rows_raw)
                last_r = rows_raw[-1]
                for ck in ['总曝光','总进房','总商品曝光','总点击','总加购']:
                    summary_row[ck] = int(last_r.get(ck,0) or 0)
                summary_row['平均流速'] = round(summary_row['总进房']/dur, 1) if dur>0 and isinstance(summary_row.get('总进房'),(int,float)) else '/'
                if isinstance(summary_row.get('成交'),(int,float)) and isinstance(summary_row.get('总加购'),(int,float)) and summary_row['总加购']>0:
                    summary_row['成交率'] = f"{round(summary_row['成交']/summary_row['总加购']*100)}%"
                # 率列加 %（除了已被强制 "/" 的列）
                for c in rate_cols:
                    if summary_row.get(c) != '/':
                        summary_row[c] = f"{int(last_r.get(c,0) or 0)}%"
                # 汇总行不需要的列全部置 /
                for ck_remove in ['曝光','进房','商品曝光','点击','加购','点击率','加购率','总点击率','总加购率']:
                    summary_row[ck_remove] = '/'
    
            # v5.4: 右键菜单操作（增删行列/合并/标黄/回写）
            if hotgrid is not None:
                res = hotgrid(data=records, columns=disp_cols, read_only_cols=READ_ONLY,
                              split_cols=SPLIT_COLS, key='hotgrid_main')
                if res and isinstance(res, dict):
                    sig = json.dumps(res.get('data'), default=str) + '|' + str(res.get('action'))
                    if st.session_state.get('_hot_sig') != sig:
                        st.session_state._hot_sig = sig
                        if res.get('action') in ('writeback', 'edit'):
                            st.session_state.data_rows = sync_grid_to_rows(
                                res.get('data', []), st.session_state.data_rows, disp_cols)
                            if res.get('action') == 'writeback':
                                st.success('✅ 已回写')
                            auto_save()
                            st.rerun()
            else:
                # 回退：原生编辑器
                col_cfg = {}
                for c in disp_cols:
                    if c in READONLY_COLS:
                        col_cfg[c] = st.column_config.Column(disabled=True, label=c)
                    elif c in PCT_COLS:
                        col_cfg[c] = st.column_config.Column(disabled=True, label=c)  # v5.4: % 已在数据中
                    elif c in CALC_COLS:
                        col_cfg[c] = st.column_config.Column(disabled=True, label=c)
                    elif c in EDITABLE_COLS:
                        col_cfg[c] = st.column_config.NumberColumn(label=c, format='%d' if c not in ['时长','总时长','平均流速'] else '%.1f')
                dkeys = list(zip(show['场次'].astype(str), show['轮次'].astype(str), show['环节'].astype(str)))
                edited_df = st.data_editor(show_disp, use_container_width=True, height=340,
                                           column_config=col_cfg, num_rows="dynamic", key='data_editor_fb', hide_index=True)
                if st.button('💾 回写保存', type='primary', use_container_width=True, key='sync_fb'):
                    push_history()
                    st.session_state.data_rows = sync_editor_to_rows(edited_df, dkeys, st.session_state.data_rows)
                    st.success('✅ 已回写'); st.rerun()
    # v5.4: 本场数据汇总行（单独显示在网格下方，紧贴）
    if summary_row:
        sdf = pd.DataFrame([summary_row])
        st.dataframe(sdf, use_container_width=True, hide_index=True)
    # 数据异常提示
    if st.session_state.data_rows:
        warns = check_anomalies(st.session_state.data_rows)
        if warns:
            cards = ''.join([f'<div style="background:rgba(255,107,107,0.10);border:1px solid rgba(255,107,107,0.30);border-radius:8px;padding:6px 10px;font-size:0.72rem;margin:3px 0"><b style="color:#ff6b6b">⚠ {tag}</b>　<span style="color:rgba(255,255,255,0.7)">{msg}</span></div>' for tag, msg in warns])
            st.markdown(f'<div style="margin-top:8px"><div style="font-size:0.78rem;color:#ff6b6b;font-weight:700;margin-bottom:4px">⚠ 数据异常（{len(warns)} 条）</div>{cards}</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)
    
else:
    # ========== 块 1：对比设置 + 数据分析表 ==========
    st.markdown("""
    <div style="background:linear-gradient(135deg,rgba(124,123,255,0.25),rgba(253,203,110,0.15));border:2px solid rgba(124,123,255,0.4);border-radius:16px;padding:14px 18px;margin:16px 0 10px 0;text-align:center">
        <div style="font-size:1.2rem;font-weight:800;color:#FDCB6E;letter-spacing:2px">📊 数据对比设置</div>
        <div style="font-size:0.7rem;color:rgba(255,255,255,0.55);margin-top:4px">分别选择两个对比对象（场次·轮次·环节自由组合），下面两张卡分别呈现：📋分析表 / 📈图表</div>
    </div>
    """, unsafe_allow_html=True)
    
    if not df_all.empty:
        # v5.4: 选择对比范围
        cmode = st.radio('📐 对比范围', ['🏟 按场次', '🔄 按轮次', '📍 按环节'], key='cmode', horizontal=True, label_visibility='collapsed')
        all_sessions = list(dict.fromkeys(df_all['场次'].tolist()))
        all_rounds = sorted([int(x) for x in df_all['轮次'].unique() if str(x).isdigit()])
        all_phases = ['干货','售卖']
    
        if '按场次' in cmode:
            cA1, cA2 = st.columns(2)
            with cA1:
                cmpA_sess = st.selectbox('🔵 对比 A · 场次', all_sessions, key='ca_ss', label_visibility='visible')
            with cA2:
                other_s = [s for s in all_sessions if s != cmpA_sess] or all_sessions
                cmpB_sess = st.selectbox('🟠 对比 B · 场次', other_s, key='cb_ss', label_visibility='visible')
            cmpA = df_all[df_all['场次']==cmpA_sess]
            cmpB = df_all[df_all['场次']==cmpB_sess]
            labelA = cmpA_sess; labelB = cmpB_sess
            # 合计各场全部数据
            cmpA = cmpA.groupby('场次', as_index=False).sum(numeric_only=True)
            cmpB = cmpB.groupby('场次', as_index=False).sum(numeric_only=True)
    
        elif '按轮次' in cmode:
            cR0, cR1, cR2 = st.columns(3)
            with cR0:
                cmp_sess = st.selectbox('场次', all_sessions, key='cr_s', label_visibility='visible')
            with cR1:
                cmpA_rnd = st.selectbox('🔵 对比 A · 轮次', all_rounds, key='cr_ar', label_visibility='visible')
            with cR2:
                other_r = [r for r in all_rounds if r != cmpA_rnd] or all_rounds
                cmpB_rnd = st.selectbox('🟠 对比 B · 轮次', other_r, key='cr_br', label_visibility='visible')
            cmpA = df_all[(df_all['场次']==cmp_sess)&(df_all['轮次']==cmpA_rnd)]
            cmpB = df_all[(df_all['场次']==cmp_sess)&(df_all['轮次']==cmpB_rnd)]
            labelA = f'{cmp_sess} · 第{cmpA_rnd}轮'; labelB = f'{cmp_sess} · 第{cmpB_rnd}轮'
    
        else:  # 按环节
            ca1, ca2, ca3 = st.columns(3)
            with ca1:
                cmpA_sess = st.selectbox('🔵 对比 A · 场次', all_sessions, key='ca_s', label_visibility='visible')
            with ca2:
                cmpA_rnd = st.selectbox('🔵 对比 A · 轮次', all_rounds, key='ca_r', label_visibility='visible')
            with ca3:
                cmpA_ph = st.selectbox('🔵 对比 A · 环节', all_phases, key='ca_p', label_visibility='visible')
    
            cb1, cb2, cb3 = st.columns(3)
            with cb1:
                cmpB_sess = st.selectbox('🟠 对比 B · 场次', all_sessions, key='cb_s', label_visibility='visible')
            with cb2:
                cmpB_rnd = st.selectbox('🟠 对比 B · 轮次', all_rounds, key='cb_r', label_visibility='visible')
            with cb3:
                cmpB_ph = st.selectbox('🟠 对比 B · 环节', all_phases, key='cb_p', label_visibility='visible')
    
            cmpA = df_all[(df_all['场次']==cmpA_sess)&(df_all['轮次']==cmpA_rnd)&(df_all['环节']==cmpA_ph)]
            cmpB = df_all[(df_all['场次']==cmpB_sess)&(df_all['轮次']==cmpB_rnd)&(df_all['环节']==cmpB_ph)]
            labelA = f"{cmpA_sess} · 第{cmpA_rnd}轮 · {cmpA_ph}"
            labelB = f"{cmpB_sess} · 第{cmpB_rnd}轮 · {cmpB_ph}"
    
        # 共享：指标选择 + 分析表/图表（三种模式共用）
        met_opts = [m for m in ['总曝光','总进房','曝光','进房','平均流速','总商品曝光','总点击','总加购','总点击率','总加购率','商品曝光','点击','加购','点击率','加购率','成交','成交率','退款','出单','时长'] if m in df_all.columns]
        def_m = ['总曝光','总进房','总点击'][:min(3,len(met_opts))]
        sel_m = st.multiselect('指标', met_opts, default=def_m, key='smi', label_visibility='collapsed')
        st.markdown(f"""
        <div class="glass" style="margin-top:14px;border:2px solid rgba(162,155,254,0.3)">
            <div style="display:flex;align-items:center;gap:8px;margin-bottom:10px">
                <span style="font-size:1.1rem;font-weight:800;color:#A29BFE">📋 数据分析表</span>
                <span style="font-size:0.7rem;color:rgba(255,255,255,0.5)">🔵 {labelA}  vs  🟠 {labelB}</span>
            </div>
        """, unsafe_allow_html=True)
        if not sel_m:
            st.markdown('<div style="color:rgba(255,255,255,0.6);font-size:0.78rem;padding:8px 12px;background:rgba(124,123,255,0.10);border-radius:8px">👆 请选择指标</div>', unsafe_allow_html=True)
        else:
            # 构建对比值
            comp_vals = {}; diffs = []
            for m in sel_m:
                vA = cmpA[m].sum() if m in cmpA.columns and not cmpA.empty else 0
                vB = cmpB[m].sum() if m in cmpB.columns and not cmpB.empty else 0
                delta = vB - vA
                pct_chg = round(delta/vA*100) if vA != 0 else (100 if vB > 0 else 0)
                is_rate = m in PCT_COLS
                comp_vals[m] = (vA, vB, delta, pct_chg, is_rate)
                diffs.append((m, abs(delta), delta, pct_chg, is_rate))
    
            # ── 紧凑摘要行 ──
            summary_parts = []
            for m, _, d, p, is_rate in sorted(diffs, key=lambda x: abs(x[1]), reverse=True)[:4]:
                arrow = '↑' if d > 0 else ('↓' if d < 0 else '→')
                color = '#48db80' if d > 0 else ('#ff6b6b' if d < 0 else '#ccc')
                if is_rate:
                    summary_parts.append(f'{m}: <b style="color:{color}">{arrow}{p}pp</b>')
                else:
                    summary_parts.append(f'{m}: <b style="color:{color}">{arrow}{abs(d):.0f}({p}%)</b>')
            if summary_parts:
                st.markdown(f'<div style="padding:10px 14px;background:rgba(124,123,255,0.08);border-radius:10px;margin:8px 0;font-size:0.78rem;color:rgba(255,255,255,0.85);line-height:1.6">📊 <b>🟠 {labelB}</b> vs <b>🔵 {labelA}</b> | {" ｜ ".join(summary_parts)}</div>', unsafe_allow_html=True)
    
            # ── 关键发现卡片 ──
            top_diffs = sorted(diffs, key=lambda x: abs(x[1]), reverse=True)[:3]
            if top_diffs:
                cards = []
                emoji_map = {True: '📈', False: '📉'}
                for i, (m, ad, d, p, is_rate) in enumerate(top_diffs):
                    color = '#48db80' if d > 0 else '#ff6b6b'
                    arrow = '上升' if d > 0 else ('下降' if d < 0 else '持平')
                    unit = 'pp' if is_rate else ''
                    cards.append(f'<div style="background:rgba(255,255,255,0.05);border-radius:10px;padding:8px 12px;text-align:center;flex:1;min-width:120px"><div style="font-size:0.68rem;color:rgba(255,255,255,0.55)">{m}</div><div style="font-size:1.1rem;font-weight:700;color:{color}">{arrow} {abs(d):.0f}{unit} ({p}%)</div></div>')
                st.markdown(f'<div style="display:flex;gap:10px;margin:8px 0">{ "".join(cards) }</div>', unsafe_allow_html=True)
    
            # ── 数据分析表（含 Δ 差异行）──
            rows_data = {}
            rows_data['项目'] = [labelA, labelB, 'Δ (B−A)', '变化%']
            for m in sel_m:
                vA, vB, delta, pct_chg, is_rate = comp_vals[m]
                if is_rate:
                    rows_data[m] = [f'{int(vA)}%', f'{int(vB)}%', f'{delta:+d}pp', f'{pct_chg:+d}%']
                else:
                    fmt = lambda v: f'{v:.1f}' if isinstance(v, float) else str(v)
                    rows_data[m] = [fmt(vA), fmt(vB), f'{delta:+.0f}', f'{pct_chg:+d}%']
            cmp_df = pd.DataFrame(rows_data)
            st.dataframe(cmp_df, use_container_width=True, hide_index=True)
    
            # 更新 chart_vals 供图表使用
            chart_vals = {}
            for m in sel_m:
                vA, vB, delta, pct_chg, is_rate = comp_vals[m]
                chart_vals[m] = [vA, vB, delta, pct_chg, is_rate]
    
            chart_tabs = st.tabs(['📊 柱状图', '📈 折线图', '🥧 饼图'])
            pal = ['#7C7BFF','#FDCB6E','#48db80','#6CCAFF','#ff6b6b','#A29BFE','#fd79a8','#e17055','#00cec9','#fab1a0']
            x_labels = [labelA, labelB]
    
            with chart_tabs[0]:
                fig = go.Figure()
                for j, m in enumerate(sel_m):
                    vals = chart_vals[m]
                    is_rate = vals[4]
                    text_vals = [f'{int(vals[0])}%' if is_rate else str(vals[0]),
                                 f'{int(vals[1])}%' if is_rate else str(vals[1])]
                    # 柱上图标注变化率
                    annotations = [
                        dict(x=x_labels[0], y=vals[0], text=text_vals[0], showarrow=False, yshift=14, font=dict(color='#fff', size=9)),
                        dict(x=x_labels[1], y=vals[1], text=f'{text_vals[1]} ({vals[3]:+d}%)', showarrow=False, yshift=14, font=dict(color='#48db80' if vals[2]>0 else '#ff6b6b' if vals[2]<0 else '#fff', size=9))
                    ]
                    fig.add_trace(go.Bar(x=x_labels, y=[vals[0], vals[1]], name=m, marker_color=pal[j%len(pal)],
                        text=text_vals, textposition='auto', textfont=dict(color='#fff')))
                fig.update_layout(barmode='group', template='plotly_dark',
                    paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                    font=dict(color='rgba(255,255,255,0.85)', size=10),
                    xaxis=dict(gridcolor='rgba(255,255,255,0.05)'),
                    yaxis=dict(gridcolor='rgba(255,255,255,0.05)'),
                    legend=dict(font=dict(color='rgba(255,255,255,0.95)', size=10), bgcolor='rgba(0,0,0,0.35)'),
                    margin=dict(l=10,r=10,t=10,b=10), height=320)
                st.plotly_chart(fig, use_container_width=True)
                plotly_download_button(fig, 'bar')
    
            with chart_tabs[1]:
                fig = go.Figure()
                for j, m in enumerate(sel_m):
                    vals = chart_vals[m]
                    is_rate = vals[4]
                    fig.add_trace(go.Scatter(x=x_labels, y=[vals[0], vals[1]], mode='lines+markers',
                        name=f'{m} ({"↑" if vals[2]>0 else "↓" if vals[2]<0 else "→"}{abs(vals[3])}%)',
                        line=dict(color=pal[j%len(pal)], width=2), marker=dict(size=8)))
                fig.update_layout(template='plotly_dark',
                    paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                    font=dict(color='rgba(255,255,255,0.85)', size=10),
                    xaxis=dict(gridcolor='rgba(255,255,255,0.05)'),
                    yaxis=dict(gridcolor='rgba(255,255,255,0.05)'),
                    legend=dict(font=dict(color='rgba(255,255,255,0.95)', size=10), bgcolor='rgba(0,0,0,0.35)'),
                    margin=dict(l=10,r=10,t=10,b=10), height=320)
                st.plotly_chart(fig, use_container_width=True)
                plotly_download_button(fig, 'line')
    
            with chart_tabs[2]:
                cols = st.columns(min(2, len(sel_m)))
                for j, m in enumerate(sel_m):
                    with cols[j % 2]:
                        fig = go.Figure()
                        vals = chart_vals[m]
                        labels = [labelA, labelB]
                        values = [vals[0], vals[1]]
                        if any(v > 0 for v in values):
                            fig.add_trace(go.Pie(labels=labels, values=values, textinfo='label+percent',
                                textfont=dict(color='#ffffff'), marker=dict(colors=pal[:2]), hole=0.35, name=m))
                            arrow = '↑' if vals[2]>0 else ('↓' if vals[2]<0 else '→')
                            fig.update_layout(title=f'{m} ({arrow}{vals[3]:+d}%)',
                                title_font_color='rgba(255,255,255,0.85)',
                                template='plotly_dark', paper_bgcolor='rgba(0,0,0,0)',
                                font=dict(color='rgba(255,255,255,0.95)', size=10),
                                legend=dict(font=dict(color='rgba(255,255,255,0.95)', size=10), bgcolor='rgba(0,0,0,0.35)'),
                                margin=dict(l=5,r=5,t=30,b=5), height=240)
                            st.plotly_chart(fig, use_container_width=True)
                            plotly_download_button(fig, 'pie_' + re.sub(r'\W+', '', str(m)))
                        else:
                            st.caption(f'{m}: 无数据')
        st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.markdown("""
        <div class="glass" style="text-align:center;padding:40px;color:rgba(255,255,255,0.15)">
            <div style="font-size:2.5rem">📭</div>
            <div>暂无数据可对比，先在左侧录入数据吧</div>
        </div>
        """, unsafe_allow_html=True)
    
# ============================================================
# 页脚
# ============================================================
ocr_status = '✅ easyocr' if (_ocr_reader and _ocr_reader != 'rapid') else ('✅ rapidocr' if _ocr_reader == 'rapid' else '❌ OCR 未装')
st.markdown(f"""
<div class="divider"></div>
<div style="display:flex;justify-content:space-between;padding:4px 6px">
    <span style="color:rgba(255,255,255,0.18);font-size:0.65rem">💡 Ctrl+V 粘贴 · 数据自动保存 · OCR: {ocr_status} · 内存保存已结束场次</span>
    <span style="color:rgba(255,255,255,0.1);font-size:0.65rem">v5.4</span>
</div>
""", unsafe_allow_html=True)