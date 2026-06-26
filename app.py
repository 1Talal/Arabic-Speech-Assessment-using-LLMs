"""
Arabic Speech Assessment - Streamlit Interface
واجهة تقييم اللغة العربية المنطوقة
"""

import os
import json
import tempfile
import subprocess
import concurrent.futures
from io import BytesIO

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from openai import OpenAI

# =====================================================
# Page Configuration
# =====================================================
st.set_page_config(
    page_title="تقييم النطق العربي",
    page_icon="🎙️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =====================================================
# RTL + Custom CSS
# =====================================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@300;400;600;700;900&display=swap');

    * {
        font-family: 'Cairo', sans-serif !important;
    }

    .main, .stApp {
        direction: rtl;
        text-align: right;
    }

    .stTextInput input, .stTextArea textarea {
        direction: rtl;
        text-align: right;
    }

    [data-testid="stSidebar"] {
        direction: rtl;
        text-align: right;
    }

    .hero-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 2rem;
        border-radius: 16px;
        color: white;
        text-align: center;
        margin-bottom: 2rem;
        box-shadow: 0 8px 24px rgba(102, 126, 234, 0.3);
    }

    .hero-header h1 {
        font-size: 2.5rem;
        font-weight: 900;
        margin: 0;
    }

    .hero-header p {
        font-size: 1.1rem;
        opacity: 0.95;
        margin-top: 0.5rem;
    }

    .score-card {
        background: white;
        padding: 1.5rem;
        border-radius: 12px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
        border-right: 4px solid #667eea;
        margin-bottom: 1rem;
    }

    .model-card {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        padding: 1.5rem;
        border-radius: 12px;
        margin-bottom: 1rem;
    }

    .best-model {
        background: linear-gradient(135deg, #ffd700 0%, #ffb347 100%);
        padding: 1.5rem;
        border-radius: 12px;
        color: #333;
        font-weight: 700;
    }

    .cefr-badge {
        display: inline-block;
        padding: 0.3rem 0.8rem;
        border-radius: 20px;
        font-weight: 700;
        font-size: 0.9rem;
        color: white;
    }
    .cefr-A1, .cefr-A2 { background: #ef4444; }
    .cefr-B1, .cefr-B2 { background: #f59e0b; }
    .cefr-C1, .cefr-C2 { background: #10b981; }

    .stButton button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.6rem 2rem;
        font-weight: 700;
        font-size: 1rem;
    }

    .stButton button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 16px rgba(102, 126, 234, 0.4);
    }

    .strength-item {
        background: #d1fae5;
        padding: 0.5rem 1rem;
        border-radius: 8px;
        margin: 0.3rem 0;
        border-right: 4px solid #10b981;
    }

    .weakness-item {
        background: #fee2e2;
        padding: 0.5rem 1rem;
        border-radius: 8px;
        margin: 0.3rem 0;
        border-right: 4px solid #ef4444;
    }

    .recommendation-box {
        background: linear-gradient(135deg, #fef3c7 0%, #fde68a 100%);
        padding: 1rem;
        border-radius: 8px;
        border-right: 4px solid #f59e0b;
        margin: 0.5rem 0;
    }
</style>
""", unsafe_allow_html=True)

# =====================================================
# Constants
# =====================================================
DEFAULT_MODELS = [
    {"id": "openai/gpt-4o-mini",            "name": "GPT-4o Mini",        "provider": "OpenAI"},
    {"id": "anthropic/claude-3.5-sonnet",   "name": "Claude 3.5 Sonnet",  "provider": "Anthropic"},
    {"id": "meta-llama/llama-3.3-70b-instruct", "name": "Llama 3.3 70B",  "provider": "Meta"},
    {"id": "google/gemini-flash-1.5",       "name": "Gemini 1.5 Flash",   "provider": "Google"},
    {"id": "mistralai/mistral-large",       "name": "Mistral Large",      "provider": "Mistral"},
    {"id": "qwen/qwen-2.5-72b-instruct",    "name": "Qwen 2.5 72B",       "provider": "Alibaba"},
]

SYSTEM_PROMPT = """أنت خبير في تقييم مستوى اللغة العربية المنطوقة.
قيّم النص العربي الناتج من كلام الطالب.
أرجع JSON فقط بدون أي كلام إضافي.

استخدم هذا الشكل بالضبط:
{
  "pronunciation_score": 7,
  "fluency_score": 6,
  "grammar_score": 8,
  "vocabulary_score": 7,
  "overall_score": 7,
  "cefr_level": "A2",
  "strengths": ["نقطة قوة 1", "نقطة قوة 2"],
  "weaknesses": ["نقطة ضعف 1", "نقطة ضعف 2"],
  "recommendation": "توصية مختصرة وعملية",
  "detailed_feedback": "تقييم تفصيلي بالعربية يشرح مستوى الطالب"
}

ملاحظات مهمة:
- الدرجات من 0 إلى 10
- مستوى CEFR: A1, A2, B1, B2, C1, C2
- كن عادلاً ودقيقاً في التقييم
- قدّم ملاحظات تفصيلية حقيقية"""

USER_TEMPLATE = """قيّم النص العربي التالي:

النص: {text}

أرجع JSON فقط."""

# =====================================================
# Session State Init
# =====================================================
if "results" not in st.session_state:
    st.session_state.results = None
if "transcript" not in st.session_state:
    st.session_state.transcript = ""
if "audio_bytes" not in st.session_state:
    st.session_state.audio_bytes = None

# =====================================================
# Helper Functions
# =====================================================
def get_cefr_color(level: str) -> str:
    level = (level or "").upper().strip()
    if level in ("A1", "A2"):
        return "#ef4444"
    if level in ("B1", "B2"):
        return "#f59e0b"
    if level in ("C1", "C2"):
        return "#10b981"
    return "#6b7280"


def transcribe_with_groq(audio_path: str, groq_key: str) -> dict:
    """Transcribe Arabic audio using Groq Whisper."""
    try:
        from groq import Groq
        client = Groq(api_key=groq_key)
        with open(audio_path, "rb") as f:
            result = client.audio.transcriptions.create(
                file=(os.path.basename(audio_path), f),
                model="whisper-large-v3",
                language="ar",
                response_format="json",
                temperature=0.0,
            )
        return {"text": result.text.strip(), "success": True, "error": None}
    except Exception as e:
        return {"text": "", "success": False, "error": str(e)}


def transcribe_with_whisper_local(audio_path: str, model_size: str = "small") -> dict:
    """Transcribe using local Whisper (slow but free)."""
    try:
        import whisper
        model = whisper.load_model(model_size)
        result = model.transcribe(
            audio_path,
            language="ar",
            task="transcribe",
            fp16=False,
            temperature=0,
            condition_on_previous_text=False,
        )
        return {"text": result["text"].strip(), "success": True, "error": None}
    except Exception as e:
        return {"text": "", "success": False, "error": str(e)}


def convert_audio_to_wav(input_path: str) -> str:
    """Convert any audio file to 16kHz mono WAV using ffmpeg."""
    output_path = input_path.rsplit(".", 1)[0] + "_converted.wav"
    subprocess.run(
        ["ffmpeg", "-y", "-i", input_path, "-ar", "16000", "-ac", "1", output_path],
        check=True,
        capture_output=True,
    )
    return output_path


def assess_with_model(client: OpenAI, model_id: str, model_name: str, text: str) -> dict:
    """Send text to a single OpenRouter model and parse the JSON result."""
    try:
        response = client.chat.completions.create(
            model=model_id,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user",   "content": USER_TEMPLATE.format(text=text)},
            ],
            temperature=0.1,
            max_tokens=1500,
        )
        raw = response.choices[0].message.content or ""
        raw = raw.replace("```json", "").replace("```", "").strip()

        start, end = raw.find("{"), raw.rfind("}") + 1
        if start == -1 or end == 0:
            raise ValueError(f"لم يتم العثور على JSON في الرد:\n{raw[:200]}")

        result = json.loads(raw[start:end])

        return {
            "Model": model_name,
            "ModelID": model_id,
            "Success": True,
            "Pronunciation":  int(result.get("pronunciation_score", 0)),
            "Fluency":        int(result.get("fluency_score", 0)),
            "Grammar":        int(result.get("grammar_score", 0)),
            "Vocabulary":     int(result.get("vocabulary_score", 0)),
            "Overall":        int(result.get("overall_score", 0)),
            "CEFR":           result.get("cefr_level", "N/A"),
            "Strengths":      result.get("strengths", []),
            "Weaknesses":     result.get("weaknesses", []),
            "Recommendation": result.get("recommendation", ""),
            "Feedback":       result.get("detailed_feedback", ""),
            "Error": None,
        }

    except Exception as e:
        return {
            "Model": model_name,
            "ModelID": model_id,
            "Success": False,
            "Pronunciation": 0, "Fluency": 0, "Grammar": 0,
            "Vocabulary": 0,   "Overall": 0,
            "CEFR": "N/A",
            "Strengths": [], "Weaknesses": [],
            "Recommendation": "",
            "Feedback": str(e),
            "Error": str(e),
        }


def run_all_assessments(api_key: str, models: list, text: str) -> list:
    """Run all selected models in parallel."""
    client = OpenAI(base_url="https://openrouter.ai/api/v1", api_key=api_key)
    results = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=len(models)) as ex:
        futures = {
            ex.submit(assess_with_model, client, m["id"], m["name"], text): m
            for m in models
        }
        for fut in concurrent.futures.as_completed(futures):
            results.append(fut.result())
    return results


def render_radar_chart(results: list):
    """Render plotly radar chart comparing models."""
    fig = go.Figure()
    categories = ["النطق", "الطلاقة", "النحو", "المفردات"]
    for r in results:
        if not r["Success"]:
            continue
        fig.add_trace(go.Scatterpolar(
            r=[r["Pronunciation"], r["Fluency"], r["Grammar"], r["Vocabulary"]],
            theta=categories,
            fill="toself",
            name=r["Model"],
        ))
    fig.update_layout(
        polar=dict(radialaxis=dict(visible=True, range=[0, 10])),
        showlegend=True,
        height=500,
        font=dict(family="Cairo", size=14),
    )
    return fig


def render_bar_chart(results: list):
    """Bar chart of overall scores."""
    successful = [r for r in results if r["Success"]]
    if not successful:
        return None
    fig = go.Figure(go.Bar(
        x=[r["Model"] for r in successful],
        y=[r["Overall"] for r in successful],
        text=[r["Overall"] for r in successful],
        textposition="outside",
        marker=dict(
            color=[r["Overall"] for r in successful],
            colorscale="Viridis",
            showscale=False,
        ),
    ))
    fig.update_layout(
        title="مقارنة الدرجة العامة بين النماذج",
        yaxis=dict(range=[0, 10], title="الدرجة"),
        xaxis=dict(title="النموذج"),
        height=400,
        font=dict(family="Cairo", size=14),
    )
    return fig


# =====================================================
# Sidebar - Settings
# =====================================================
with st.sidebar:
    st.markdown("### الإعدادات")

    openrouter_key = st.text_input(
        "مفتاح OpenRouter API",
        value=os.getenv("OPENROUTER_API_KEY", ""),
        type="password",
        help="احصل على المفتاح من openrouter.ai",
    )

    st.markdown("---")
    st.markdown("### طريقة التفريغ الصوتي")

    transcribe_method = st.radio(
        "اختر الطريقة",
        ["Groq Whisper (سريع)", "Whisper محلي (بطيء)", "إدخال نص يدوي"],
        index=2,
    )

    groq_key = ""
    if "Groq" in transcribe_method:
        groq_key = st.text_input(
            "مفتاح Groq API",
            value=os.getenv("GROQ_API_KEY", ""),
            type="password",
        )

    st.markdown("---")
    st.markdown("### النماذج المستخدمة")

    selected_models = []
    for i, m in enumerate(DEFAULT_MODELS):
        if st.checkbox(f"{m['name']} ({m['provider']})", value=(i < 3), key=f"model_{i}"):
            selected_models.append(m)

    st.markdown("---")
    st.caption("صُنع بحب لتعليم اللغة العربية")


# =====================================================
# Main Content
# =====================================================
st.markdown("""
<div class="hero-header">
    <h1>🎙️ تقييم النطق العربي بالذكاء الاصطناعي</h1>
    <p>قارن بين أفضل نماذج LLM لتقييم مستواك في اللغة العربية</p>
</div>
""", unsafe_allow_html=True)

# Input section
tab1, tab2 = st.tabs(["📥 الإدخال", "📊 النتائج"])

with tab1:
    if "نص" in transcribe_method:
        st.markdown("### أدخل النص العربي للتقييم")
        text_input = st.text_area(
            "اكتب أو الصق النص هنا",
            height=200,
            placeholder="مثال: السلام عليكم، اسمي محمد، أحب أن أتعلم اللغة العربية لأنها لغة جميلة وغنية.",
        )
        st.session_state.transcript = text_input
    else:
        st.markdown("### ارفع ملف صوتي بالعربية")
        audio_file = st.file_uploader(
            "اختر ملف صوتي (mp3, wav, m4a, ogg)",
            type=["mp3", "wav", "m4a", "ogg", "flac"],
        )

        if audio_file is not None:
            st.session_state.audio_bytes = audio_file.read()
            st.session_state.audio_name = audio_file.name
            st.audio(st.session_state.audio_bytes)

            if st.button("🎯 تفريغ الصوت إلى نص", use_container_width=True):
                with st.spinner("جاري تفريغ الصوت..."):
                    # استخراج الامتداد الأصلي من اسم الملف
                    original_ext = os.path.splitext(audio_file.name)[1].lower() or ".mp3"
                    if original_ext not in [".flac", ".mp3", ".mp4", ".mpeg", ".mpga",
                                            ".m4a", ".ogg", ".opus", ".wav", ".webm"]:
                        original_ext = ".mp3"

                    with tempfile.NamedTemporaryFile(delete=False, suffix=original_ext) as tmp:
                        tmp.write(st.session_state.audio_bytes)
                        tmp_path = tmp.name

                    try:
                        wav_path = convert_audio_to_wav(tmp_path)
                    except Exception as e:
                        st.warning(f"تعذرت معالجة الصوت بـ ffmpeg، سيستخدم الملف الأصلي. ({e})")
                        wav_path = tmp_path

                    if "Groq" in transcribe_method:
                        if not groq_key:
                            st.error("الرجاء إدخال مفتاح Groq API")
                        else:
                            result = transcribe_with_groq(wav_path, groq_key)
                            if result["success"]:
                                st.session_state.transcript = result["text"]
                                st.success("تم التفريغ بنجاح!")
                            else:
                                st.error(f"خطأ: {result['error']}")
                    else:
                        result = transcribe_with_whisper_local(wav_path)
                        if result["success"]:
                            st.session_state.transcript = result["text"]
                            st.success("تم التفريغ بنجاح!")
                        else:
                            st.error(f"خطأ: {result['error']}")

        if st.session_state.transcript:
            st.markdown("#### النص المستخرج:")
            st.session_state.transcript = st.text_area(
                "يمكنك التعديل قبل التقييم",
                value=st.session_state.transcript,
                height=150,
            )

    st.markdown("---")

    col_a, col_b = st.columns([3, 1])
    with col_a:
        if st.button("🚀 ابدأ التقييم بكل النماذج", use_container_width=True, type="primary"):
            if not openrouter_key:
                st.error("الرجاء إدخال مفتاح OpenRouter API في الشريط الجانبي")
            elif not st.session_state.transcript or not st.session_state.transcript.strip():
                st.error("الرجاء إدخال نص أو تفريغ ملف صوتي أولاً")
            elif not selected_models:
                st.error("اختر نموذجاً واحداً على الأقل من الشريط الجانبي")
            else:
                with st.spinner(f"جاري التقييم باستخدام {len(selected_models)} نماذج..."):
                    st.session_state.results = run_all_assessments(
                        openrouter_key, selected_models, st.session_state.transcript
                    )
                st.success("اكتمل التقييم! انتقل إلى تبويب النتائج.")

with tab2:
    if st.session_state.results is None:
        st.info("لم يتم التقييم بعد. ابدأ من تبويب الإدخال.")
    else:
        results = st.session_state.results
        successful = [r for r in results if r["Success"]]

        if not successful:
            st.error("فشلت جميع النماذج. تحقق من المفتاح والاتصال.")
            for r in results:
                st.warning(f"{r['Model']}: {r['Error']}")
        else:
            # --- Best model ---
            best = max(successful, key=lambda r: r["Overall"])
            avg_overall = sum(r["Overall"] for r in successful) / len(successful)

            col1, col2, col3, col4 = st.columns(4)
            col1.metric("النص المُقيَّم", f"{len(st.session_state.transcript.split())} كلمة")
            col2.metric("النماذج الناجحة", f"{len(successful)} / {len(results)}")
            col3.metric("متوسط الدرجة", f"{avg_overall:.1f} / 10")
            col4.metric("أعلى درجة", f"{best['Overall']} / 10")

            st.markdown("---")

            # --- Best model card ---
            cefr_color = get_cefr_color(best["CEFR"])
            st.markdown(f"""
            <div class="best-model">
                <h3 style="margin: 0;">🏆 أفضل تقييم: {best['Model']}</h3>
                <p style="margin: 0.5rem 0 0 0; font-size: 1.1rem;">
                    الدرجة العامة: <strong>{best['Overall']}/10</strong> &nbsp;|&nbsp;
                    مستوى CEFR:
                    <span style="background: {cefr_color}; color: white; padding: 0.2rem 0.6rem; border-radius: 12px;">
                        {best['CEFR']}
                    </span>
                </p>
            </div>
            """, unsafe_allow_html=True)

            st.markdown("---")

            # --- Charts ---
            chart_col1, chart_col2 = st.columns(2)
            with chart_col1:
                bar = render_bar_chart(results)
                if bar:
                    st.plotly_chart(bar, use_container_width=True)
            with chart_col2:
                st.plotly_chart(render_radar_chart(results), use_container_width=True)

            st.markdown("---")
            st.markdown("### 📋 تفاصيل التقييم لكل نموذج")

            # --- Per-model detailed cards ---
            for r in results:
                if not r["Success"]:
                    with st.expander(f"❌ {r['Model']} - فشل التقييم"):
                        st.error(r["Error"])
                    continue

                with st.expander(f"📊 {r['Model']} - الدرجة العامة: {r['Overall']}/10", expanded=False):
                    c1, c2, c3, c4, c5 = st.columns(5)
                    c1.metric("النطق",     r["Pronunciation"])
                    c2.metric("الطلاقة",   r["Fluency"])
                    c3.metric("النحو",     r["Grammar"])
                    c4.metric("المفردات",  r["Vocabulary"])
                    c5.metric("المستوى",   r["CEFR"])

                    sc1, sc2 = st.columns(2)
                    with sc1:
                        st.markdown("#### ✅ نقاط القوة")
                        for s in r["Strengths"]:
                            st.markdown(f'<div class="strength-item">{s}</div>', unsafe_allow_html=True)
                    with sc2:
                        st.markdown("#### ⚠️ نقاط الضعف")
                        for w in r["Weaknesses"]:
                            st.markdown(f'<div class="weakness-item">{w}</div>', unsafe_allow_html=True)

                    if r["Recommendation"]:
                        st.markdown("#### 💡 التوصية")
                        st.markdown(
                            f'<div class="recommendation-box">{r["Recommendation"]}</div>',
                            unsafe_allow_html=True,
                        )

                    if r["Feedback"]:
                        st.markdown("#### 📝 التقييم التفصيلي")
                        st.info(r["Feedback"])

            st.markdown("---")

            # --- Download CSV ---
            df = pd.DataFrame([{
                "Model":          r["Model"],
                "Success":        r["Success"],
                "Pronunciation":  r["Pronunciation"],
                "Fluency":        r["Fluency"],
                "Grammar":        r["Grammar"],
                "Vocabulary":     r["Vocabulary"],
                "Overall":        r["Overall"],
                "CEFR":           r["CEFR"],
                "Strengths":      " | ".join(r["Strengths"]) if isinstance(r["Strengths"], list) else r["Strengths"],
                "Weaknesses":     " | ".join(r["Weaknesses"]) if isinstance(r["Weaknesses"], list) else r["Weaknesses"],
                "Recommendation": r["Recommendation"],
                "Feedback":       r["Feedback"],
                "Error":          r["Error"],
            } for r in results])

            csv = df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
            st.download_button(
                "📥 تحميل النتائج بصيغة CSV",
                data=csv,
                file_name="arabic_assessment_results.csv",
                mime="text/csv",
                use_container_width=True,
            )
