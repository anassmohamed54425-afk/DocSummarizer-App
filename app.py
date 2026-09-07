import streamlit as st
import PyPDF2
import io
import re
from docx import Document
import datetime

# ========================================
# إعدادات الصفحة
# ========================================
st.set_page_config(
    page_title="تحليل المعامل الذكي",
    page_icon="🧪",
    layout="wide"
)

# ========================================
# CSS للشكل الاحترافي
# ========================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap');
    * { font-family: 'Cairo', sans-serif; }
    
    .main-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 40px;
        border-radius: 20px;
        text-align: center;
        color: white;
        margin-bottom: 30px;
    }
    .main-header h1 { font-size: 48px; font-weight: 700; margin: 0; }
    .main-header p { font-size: 18px; opacity: 0.9; margin: 10px 0 0; }
    
    .test-card {
        background: var(--secondary-background-color);
        padding: 20px;
        border-radius: 12px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.06);
        margin: 12px 0;
        border-right: 5px solid #667eea;
        color: var(--text-color);
    }
    .test-card .test-name {
        font-size: 18px;
        font-weight: 700;
        color: var(--text-color);
    }
    .test-card .test-result {
        font-size: 16px;
        color: var(--text-color);
        opacity: 0.8;
    }
    .test-normal { border-right-color: #27ae60 !important; }
    .test-high { border-right-color: #e74c3c !important; }
    .test-low { border-right-color: #f39c12 !important; }
    
    .metric-box {
        background: var(--secondary-background-color);
        padding: 20px;
        border-radius: 12px;
        text-align: center;
        border: 1px solid var(--border-color);
    }
    .metric-box .value { font-size: 28px; font-weight: 700; color: var(--text-color); }
    .metric-box .label { font-size: 14px; color: var(--text-color); opacity: 0.7; margin-top: 5px; }
    
    .stButton > button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        padding: 12px 30px;
        border-radius: 10px;
        font-weight: 600;
        font-size: 16px;
        width: 100%;
    }
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 25px rgba(102, 126, 234, 0.4);
    }
</style>
""", unsafe_allow_html=True)

# ========================================
# الشريط الجانبي
# ========================================
with st.sidebar:
    st.markdown("""
    <div style="padding: 20px 10px;">
        <h3 style="color: #667eea; border-bottom: 2px solid #667eea; padding-bottom: 10px;">🧪 تحليل المعامل</h3>
        <p style="color: var(--text-color); opacity: 0.7; font-size: 14px;">تطبيق ذكي لتحليل وتلخيص تقارير المعامل الطبية</p>
        <hr style="border-color: var(--border-color);">
        <h4 style="color: var(--text-color);">⚡ الميزات</h4>
        <ul style="color: var(--text-color); font-size: 14px; list-style: none; padding: 0;">
            <li>📄 قراءة TXT, PDF, DOCX</li>
            <li>🧪 استخراج التحاليل الطبية</li>
            <li>📊 مقارنة بالمعدلات الطبيعية</li>
            <li>🎯 تقييم النتائج (طبيعي/مرتفع/منخفض)</li>
            <li>📥 تصدير تقرير منسق</li>
        </ul>
        <hr style="border-color: var(--border-color);">
        <h4 style="color: var(--text-color);">📌 الإصدار</h4>
        <p style="color: var(--text-color); opacity: 0.7; font-size: 12px;">v1.0 - Lab Analyzer</p>
    </div>
    """, unsafe_allow_html=True)

# ========================================
# رأس الصفحة
# ========================================
st.markdown("""
<div class="main-header">
    <h1>🧪 تحليل المعامل الذكي</h1>
    <p>ارفع تقرير معمل، واستخرج التحاليل مع التقييم</p>
</div>
""", unsafe_allow_html=True)

# ========================================
# دوال التنظيف والتحليل
# ========================================

def clean_text(text):
    text = re.sub(r'[═─▄▀█░▒▓▔▕▖▗▘▙▚▛▜▝▞▟■□▢▣▤▥▦▧▨▩▪▫▬▭▮▯]', '', text)
    text = re.sub(r'\n\s*\n', '\n\n', text)
    text = re.sub(r'═+', '', text)
    text = re.sub(r'─+', '', text)
    return text.strip()

def read_file(uploaded_file):
    content = uploaded_file.read()
    filename = uploaded_file.name
    if filename.endswith('.pdf'):
        pdf_reader = PyPDF2.PdfReader(io.BytesIO(content))
        text = ""
        for page in pdf_reader.pages:
            text += page.extract_text()
        return text
    elif filename.endswith('.docx'):
        doc = Document(io.BytesIO(content))
        text = "\n".join([para.text for para in doc.paragraphs])
        return text
    else:
        return content.decode("utf-8")

# قاعدة بيانات التحاليل الطبية (اسم التحليل، الوحدة، النطاق الطبيعي)
medical_tests = {
    "سكر صائم": {"unit": "mg/dL", "normal": (70, 100)},
    "سكر فاطر": {"unit": "mg/dL", "normal": (70, 140)},
    "هيموجلوبين": {"unit": "g/dL", "normal": (12, 16)},
    "صفائح دموية": {"unit": "10^3/µL", "normal": (150, 400)},
    "كريات بيضاء": {"unit": "10^3/µL", "normal": (4, 11)},
    "كريات حمراء": {"unit": "10^6/µL", "normal": (4.5, 5.5)},
    "صوديوم": {"unit": "mEq/L", "normal": (135, 145)},
    "بوتاسيوم": {"unit": "mEq/L", "normal": (3.5, 5.0)},
    "كالسيوم": {"unit": "mg/dL", "normal": (8.5, 10.5)},
    "يوريا": {"unit": "mg/dL", "normal": (7, 20)},
    "كرياتينين": {"unit": "mg/dL", "normal": (0.6, 1.2)},
    "كولسترول": {"unit": "mg/dL", "normal": (125, 200)},
    "دهون ثلاثية": {"unit": "mg/dL", "normal": (50, 150)},
    "فيتامين د": {"unit": "ng/mL", "normal": (30, 100)},
    "حديد": {"unit": "µg/dL", "normal": (60, 170)},
    "فيريتين": {"unit": "ng/mL", "normal": (20, 500)},
}

def extract_lab_tests(text):
    """استخراج التحاليل الطبية من النص"""
    found_tests = []
    text_lower = text.lower()
    
    for test_name, test_info in medical_tests.items():
        # البحث عن اسم التحليل في النص
        if test_name.lower() in text_lower:
            # محاولة استخراج الرقم بعد اسم التحليل
            pattern = rf"{test_name}[:\s]*([0-9]+\.?[0-9]*)"
            match = re.search(pattern, text, re.IGNORECASE)
            
            if match:
                value = float(match.group(1))
                normal_min, normal_max = test_info["normal"]
                unit = test_info["unit"]
                
                # تقييم النتيجة
                if value < normal_min:
                    status = "منخفض ⬇️"
                    status_class = "test-low"
                elif value > normal_max:
                    status = "مرتفع ⬆️"
                    status_class = "test-high"
                else:
                    status = "طبيعي ✅"
                    status_class = "test-normal"
                
                found_tests.append({
                    "name": test_name,
                    "value": value,
                    "unit": unit,
                    "normal_range": f"{normal_min} - {normal_max}",
                    "status": status,
                    "status_class": status_class
                })
    
    return found_tests

# ========================================
# واجهة المستخدم
# ========================================
uploaded_file = st.file_uploader("📂 اختر ملف تقرير معمل", type=["txt", "pdf", "docx"])

if uploaded_file is not None:
    try:
        text = read_file(uploaded_file)
    except Exception as e:
        st.error(f"❌ مشكلة في قراءة الملف: {str(e)}")
        st.stop()

    clean_text_content = clean_text(text)

    with st.expander("📄 النص الأصلي"):
        st.text(clean_text_content[:1000] + ("..." if len(clean_text_content) > 1000 else ""))

    # ========================================
    # استخراج التحاليل
    # ========================================
    st.markdown("---")
    st.subheader("🧪 التحاليل المستخرجة")

    tests = extract_lab_tests(clean_text_content)

    if not tests:
        st.warning("⚠️ لم يتم العثور على تحاليل طبية في هذا الملف. تأكد من أن الملف يحتوي على أسماء تحاليل معروفة.")
    else:
        # عرض كل تحليل في بطاقة
        for test in tests:
            st.markdown(f"""
            <div class="test-card {test['status_class']}">
                <div class="test-name">🧪 {test['name']}</div>
                <div class="test-result">
                    <strong>النتيجة:</strong> {test['value']} {test['unit']}
                    <span style="margin: 0 10px;">|</span>
                    <strong>المعدل الطبيعي:</strong> {test['normal_range']} {test['unit']}
                    <span style="margin: 0 10px;">|</span>
                    <strong>التقييم:</strong> {test['status']}
                </div>
            </div>
            """, unsafe_allow_html=True)

        # ========================================
        # إحصائيات
        # ========================================
        st.markdown("---")
        st.subheader("📊 ملخص التحاليل")

        total = len(tests)
        normal = sum(1 for t in tests if "طبيعي" in t['status'])
        high = sum(1 for t in tests if "مرتفع" in t['status'])
        low = sum(1 for t in tests if "منخفض" in t['status'])

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.markdown(f"""
            <div class="metric-box">
                <div class="value">{total}</div>
                <div class="label">إجمالي التحاليل</div>
            </div>
            """, unsafe_allow_html=True)
        with col2:
            st.markdown(f"""
            <div class="metric-box">
                <div class="value" style="color: #27ae60;">{normal}</div>
                <div class="label">طبيعي ✅</div>
            </div>
            """, unsafe_allow_html=True)
        with col3:
            st.markdown(f"""
            <div class="metric-box">
                <div class="value" style="color: #e74c3c;">{high}</div>
                <div class="label">مرتفع ⬆️</div>
            </div>
            """, unsafe_allow_html=True)
        with col4:
            st.markdown(f"""
            <div class="metric-box">
                <div class="value" style="color: #f39c12;">{low}</div>
                <div class="label">منخفض ⬇️</div>
            </div>
            """, unsafe_allow_html=True)

    # ========================================
    # تحميل التقرير
    # ========================================
    st.markdown("---")
    st.subheader("📥 تحميل التقرير")

    # بناء التقرير النصي
    report_lines = []
    report_lines.append("=" * 65)
    report_lines.append("              🧪 تقرير تحليل المعامل")
    report_lines.append("=" * 65)
    report_lines.append("")
    report_lines.append(f"  📅 التاريخ          :  {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    report_lines.append("  " + "-" * 60)
    report_lines.append("")

    if tests:
        for test in tests:
            report_lines.append(f"  🧪 {test['name']}")
            report_lines.append(f"     النتيجة: {test['value']} {test['unit']}")
            report_lines.append(f"     المعدل الطبيعي: {test['normal_range']} {test['unit']}")
            report_lines.append(f"     التقييم: {test['status']}")
            report_lines.append("")
    else:
        report_lines.append("  ⚠️ لم يتم العثور على تحاليل طبية في هذا الملف.")
        report_lines.append("")

    report_lines.append("  " + "-" * 60)
    report_lines.append("")
    report_lines.append("  ✅ تم إنشاء التقرير بواسطة تطبيق تحليل المعامل الذكي")
    report_lines.append("  📌 v1.0 - Lab Analyzer")
    report_lines.append("")
    report_lines.append("=" * 65)

    report_text = "\n".join(report_lines)

    st.download_button(
        label="📥 تحميل تقرير التحاليل (TXT)",
        data=report_text,
        file_name=f"تقرير_تحاليل_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
        mime="text/plain"
    )

else:
    st.info("⏳ انتظر رفع ملف تقرير معمل لتحليله")
    st.markdown("""
    ### 🚀 طريقة الاستخدام:
    1. اضغط على زر **"اختر ملف تقرير معمل"**
    2. اختر ملف `.txt` أو `.pdf` أو `.docx`
    3. انتظر لحظات وستظهر التحاليل المستخرجة
    4. يمكنك تحميل التقرير
    """)