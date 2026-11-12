import streamlit as st
import streamlit.components.v1 as components
import pymupdf as fitz
from pptx import Presentation
import requests
import time
import re
import base64
import io
import os
import xml.etree.ElementTree as ET
from streamlit_drawable_canvas import st_canvas
from PIL import Image

st.set_page_config(page_title="GeekReader Pro", page_icon="🚀", layout="wide")

# ================= 获取云端 API Key =================
# 优先从云端 secrets 读取，如果本地跑，就给个空字符串
try:
    api_key = st.secrets["DEEPSEEK_API_KEY"]
except:
    api_key = ""

# ================= 状态初始化 =================
if 'app_mode' not in st.session_state: st.session_state.app_mode = "🔬 研究模式"
if 'current_file' not in st.session_state: st.session_state.current_file = None
if 'full_text' not in st.session_state: st.session_state.full_text = ""
if 'translation_cache' not in st.session_state: st.session_state.translation_cache = {}
if 'view_mode' not in st.session_state: st.session_state.view_mode = "📖 沉浸式双语"
if 'font_size' not in st.session_state: st.session_state.font_size = 18

# 研究模式数据
if 'med_glossary' not in st.session_state: st.session_state.med_glossary = []
# 学习模式数据 (课程隔离)
if 'course_glossary' not in st.session_state: st.session_state.course_glossary = {"默认课程": []}
if 'current_course' not in st.session_state: st.session_state.current_course = "默认课程"

# 画板数据
if 'canvas_drawings' not in st.session_state: st.session_state.canvas_drawings = {}

# ================= 核心函数 =================
def ask_tutor(text, key, mode="explain"):
    if not key:
        return "⚠️ 请先在左侧配置 API Key！"
    url = "https://api.deepseek.com/chat/completions"
    headers = {"Authorization": f"Bearer {key}", "Content-Type": "application/json"}
    if st.session_state.app_mode == "🔬 研究模式":
        if mode == "explain": prompt = f"你是一个医学导师。分点解析这段内容：\n\n{text}"
        else: prompt = f"请翻译并解析以下医学词汇。格式：\n【翻译】中文\n【解析】概念\n\n{text}"
    else:
        if mode == "explain": prompt = f"你是一位学科老师。请对以下内容拆解并通俗解释：\n\n{text}"
        else: prompt = f"请给出词义、词性及常见考点。格式：\n【翻译】词义\n【考点】考点\n\n{text}"
    try:
        res = requests.post(url, headers=headers, json={"model": "deepseek-chat", "messages": [{"role": "user", "content": prompt}], "temperature": 0.3}, timeout=60)
        return res.json()['choices'][0]['message']['content'].strip()
    except: return "❌ 网络错误或 Key 失效"

def translate_and_summarize(text, key):
    if not key: return "⚠️ 缺少Key", "无法翻译"
    url = "https://api.deepseek.com/chat/completions"
    headers = {"Authorization": f"Bearer {key}", "Content-Type": "application/json"}
    prompt = f"严格按格式输出：\n【翻译】\n(中文翻译)\n【总结】\n(20字以内概括)\n原文：\n{text}"
    try:
        res = requests.post(url, headers=headers, json={"model": "deepseek-chat", "messages": [{"role": "user", "content": prompt}]}, timeout=90)
        res_text = res.json()['choices'][0]['message']['content'].strip()
        if "【翻译】" in res_text and "【总结】" in res_text:
            parts = res_text.split("【总结】")
            return parts[0].replace("【翻译】", "").strip(), parts[1].strip()
    except: pass
    return "⚠️ 翻译失败", "总结失败"

def split_text(text, max_chars=1200):
    text = text.replace('\n', ' ')
    sentences = re.split(r'(?<=[.!?])\s+', text)
    chunks, current = [], ""
    for s in sentences:
        if len(current) + len(s) < max_chars: current += s + " "
        else:
            chunks.append(current.strip()); current = s + " "
    if current: chunks.append(current.strip())
    return chunks

def parse_res(res_text, type_name):
    if "【翻译】" in res_text and f"【{type_name}】" in res_text:
        parts = res_text.split(f"【{type_name}】")
        return parts[0].replace("【翻译】", "").strip(), parts[1].strip()
    return res_text, ""

# ================= 注入划词 JS & CSS =================
st.markdown("""
<script>
document.addEventListener('mouseup', function() {
    var t = window.getSelection().toString().trim();
    var b = document.getElementById('floating-btn');
    if (b) b.remove();
    if (t.length > 0 && t.length < 100) {
        var btn = document.createElement('button');
        btn.id = 'floating-btn'; btn.innerText = '🔍 解析';
        btn.style = 'position:absolute; left:'+window.event.pageX+'px; top:'+(window.event.pageY-40)+'px; z-index:9999; padding:5px 10px; background:#e67e22; color:white; border:none; border-radius:5px; cursor:pointer;';
        btn.onclick = function() {
            var ta = window.parent.document.querySelectorAll('textarea')[0];
            if(ta) {
                var setter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, "value").set;
                setter.call(ta, t); ta.dispatchEvent(new Event('input', { bubbles: true })); btn.remove();
            }
        };
        document.body.appendChild(btn);
    }
});
</script>
""", unsafe_allow_html=True)

st.markdown(f"""
    <style>
        #MainMenu {{visibility: hidden;}} footer {{visibility: hidden;}}
        .en-text {{ font-family: 'Georgia', serif; font-size: {st.session_state.font_size}px; color: #222; line-height: 1.8; }}
        .zh-text {{ font-family: 'Microsoft YaHei', sans-serif; font-size: {st.session_state.font_size - 2}px; color: #2e7d32; line-height: 1.8; }}
        .term-card {{ background-color: #fff; border: 1px solid #eee; padding: 15px; border-radius: 8px; margin-bottom: 15px; border-left: 4px solid #e67e22; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }}
        .term-title {{ color: #e67e22; font-weight: bold; font-size: 18px; margin-bottom: 5px; }}
    </style>
""", unsafe_allow_html=True)

# ================= 侧边栏 =================
with st.sidebar:
    st.title("🚀 GeekReader Pro")
    
    # 模式切换
    st.session_state.app_mode = st.radio("模式", ["🔬 研究模式", "🎓 学习模式"], horizontal=True)
    st.markdown("---")
    
    nav_choice = st.radio("导航", ["📖 阅读台", "🎨 画板", "📝 知识本"], label_visibility="collapsed")
    
    with st.expander("🔑 API 密钥 (云端部署需填)", expanded=not api_key):
        if not api_key:
            api_key = st.text_input("DeepSeek Key", type="password")

    if nav_choice == "📖 阅读台":
        st.radio("视图", ["📖 沉浸式双语", "🇬🇧 纯原文"], key="view_mode")
        st.session_state.font_size = st.slider("字号", 14, 28, st.session_state.font_size)
        
        with st.expander("⚡ 划词/长难句解析", expanded=True):
            paste = st.text_area("粘贴/选中文字", height=68, key="paste_box")
            if paste and len(paste) < 200 and st.session_state.get("last_p") != paste:
                st.session_state["last_p"] = paste
                with st.spinner("解析中..."):
                    if st.session_state.app_mode == "🔬 研究模式":
                        res = ask_tutor(paste, api_key, "translate")
                        trans, exp = parse_res(res, "解析")
                        st.session_state.med_glossary.append({"term": paste, "trans": trans, "explain": exp})
                        st.success(res)
                    else:
                        res = ask_tutor(paste, api_key, "translate")
                        trans, exp = parse_res(res, "考点")
                        st.session_state.course_glossary[st.session_state.current_course].append({"term": paste, "trans": trans, "explain": exp})
                        st.success(res)

# ================= 页面 1：阅读台 =================
if nav_choice == "📖 阅读台":
    if st.session_state.app_mode == "🔬 研究模式":
        st.title("🔬 严谨科研工作台")
    else:
        st.title("🎓 沉浸式课程学习")
        # 课程管理
        c1, c2 = st.columns([2, 1])
        with c1:
            courses = list(st.session_state.course_glossary.keys())
            st.session_state.current_course = st.selectbox("当前课程", courses)
        with c2:
            new_c = st.text_input("新课程", placeholder="四六级")
            if st.button("➕ 创建") and new_c:
                st.session_state.course_glossary[new_c] = []; st.session_state.current_course = new_c; st.rerun()

    # 文件上传
    with st.expander("📄 上传 PDF/PPT", expanded=False):
        file = st.file_uploader("上传", type=["pdf", "pptx"], label_visibility="collapsed")
        if file and st.session_state.current_file != file.name:
            st.session_state.current_file = file.name; st.session_state.translation_cache = {}
            if file.name.endswith('.pdf'):
                doc = fitz.open(stream=file.read(), filetype="pdf")
                st.session_state.full_text = "".join([doc.load_page(i).get_text() + "\n" for i in range(len(doc))])
            st.rerun()

    # 阅读区
    if st.session_state.current_file and st.session_state.full_text:
        chunks = split_text(st.session_state.full_text)
        c1, c2 = st.columns([8, 2])
        with c1: st.info(f"📄 {st.session_state.current_file}")
        with c2:
            if st.button("🚀 一键翻译", use_container_width=True):
                prog = st.progress(0)
                for i in range(len(chunks)):
                    if i not in st.session_state.translation_cache:
                        zh, _ = translate_and_summarize(chunks[i], api_key); st.session_state.translation_cache[i] = zh
                    prog.progress((i+1)/len(chunks))
                st.rerun()
                
        for i in range(len(chunks)):
            if st.session_state.view_mode == "📖 沉浸式双语" and i in st.session_state.translation_cache:
                ce, cz = st.columns(2)
                ce.markdown(f"<div class='en-text'>{chunks[i]}</div>", unsafe_allow_html=True)
                cz.markdown(f"<div class='zh-text'>{st.session_state.translation_cache[i]}</div>", unsafe_allow_html=True)
            else:
                st.markdown(f"<div class='en-text'>{chunks[i]}</div>", unsafe_allow_html=True)
            
            if st.button("🧑‍🏫 解析这段", key=f"btn_{i}"):
                st.info(ask_tutor(chunks[i], api_key, "explain"))
            st.markdown("---")

# ================= 页面 2：画板 =================
elif nav_choice == "🎨 画板":
    st.title("🎨 随记/默写板")
    with st.form("canvas"):
        res = st_canvas(stroke_width=2, stroke_color="#ff0000", background_color="#fff", height=400, width=800, key="draw")
        if st.form_submit_button("保存画板"):
            if res.image_data is not None:
                img = Image.fromarray(res.image_data.astype('uint8'), 'RGBA')
                buf = io.BytesIO(); img.save(buf, format="PNG")
                st.session_state.canvas_drawings[str(time.time())] = base64.b64encode(buf.getvalue()).decode()
                st.success("保存成功")

# ================= 页面 3：知识本 =================
elif nav_choice == "📝 知识本":
    st.title("📝 专属知识本")
    if st.session_state.app_mode == "🔬 研究模式":
        items = st.session_state.med_glossary
        if not items: st.caption("无记录")
        for item in items: st.markdown(f"<div class='term-card'><div class='term-title'>{item['term']}</div><div><b>翻译：</b>{item['trans']}</div><div>{item['explain']}</div></div>", unsafe_allow_html=True)
    else:
        tabs = st.tabs(list(st.session_state.course_glossary.keys()))
        for i, tab in enumerate(tabs):
            c_name = list(st.session_state.course_glossary.keys())[i]
            with tab:
                items = st.session_state.course_glossary[c_name]
                if not items: st.caption("该课程暂无记录")
                for item in items: st.markdown(f"<div class='term-card'><div class='term-title'>{item['term']}</div><div><b>释义：</b>{item['trans']}</div><div>{item['explain']}</div></div>", unsafe_allow_html=True)