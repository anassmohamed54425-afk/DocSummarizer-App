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
    page_title="ملخص المستندات الذكي",
    page_icon="📄",
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
    
    .result-card {
        background: var(--secondary-background-color);
        padding: 25px;
        border-radius: 15px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.08);
        margin: 20px 0;
        border-right: 6px solid #667eea;
        font-size: 18px;
        line-height: 1.8;
        color: var(--text-color);
    }
    
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
    
    .test-card {
        background: var(--secondary-background-color);
        padding: 20px;
        border-radius: 12px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.06);
        margin: 12px 0;
        border-right: 5px solid #667eea;
        color: var(--text-color);
    }
    .test-card .test-name { font-size: 18px; font-weight: 700; }
    .test-card .test-result { font-size: 16px; opacity: 0.8; }
    .test-normal { border-right-color: #27ae60 !important; }
    .test-high { border-right-color: #e74c3c !important; }
    .test-low { border-right-color: #f39c12 !important; }
    
    .email-card {
        background: var(--secondary-background-color);
        padding: 20px;
        border-radius: 12px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.06);
        margin: 12px 0;
        border-right: 5px solid #667eea;
        color: var(--text-color);
    }
    .email-card .email-subject { font-size: 18px; font-weight: 700; }
    .email-card .email-summary { font-size: 15px; opacity: 0.9; margin: 5px 0; }
    .email-card .email-classification { font-size: 14px; opacity: 0.8; }
    .email-complaint { border-right-color: #e74c3c !important; }
    .email-inquiry { border-right-color: #3498db !important; }
    .email-order { border-right-color: #2ecc71 !important; }
    .email-reply { border-right-color: #f39c12 !important; }
    .email-other { border-right-color: #95a5a6 !important; }
</style>
""", unsafe_allow_html=True)

# ========================================
# الشريط الجانبي
# ========================================
with st.sidebar:
    st.markdown("""
    <div style="padding: 20px 10px;">
        <h3 style="color: #667eea; border-bottom: 2px solid #667eea; padding-bottom: 10px;">📊 تحليل المستندات</h3>
        <p style="color: var(--text-color); opacity: 0.7; font-size: 14px;">تطبيق ذكي لتحليل وتلخيص المستندات النصية</p>
        <hr style="border-color: var(--border-color);">
        <h4 style="color: var(--text-color);">⚡ الميزات</h4>
        <ul style="color: var(--text-color); font-size: 14px; list-style: none; padding: 0;">
            <li>📄 قراءة TXT, PDF, DOCX</li>
            <li>📝 تلخيص ذكي</li>
            <li>🏷️ تصنيف تلقائي</li>
            <li>📊 إحصائيات متقدمة</li>
            <li>🧪 تحليل المعامل</li>
            <li>📧 تلخيص الإيميلات</li>
            <li>📥 تصدير تقرير منسق</li>
        </ul>
        <hr style="border-color: var(--border-color);">
        <h4 style="color: var(--text-color);">📌 الإصدار</h4>
        <p style="color: var(--text-color); opacity: 0.7; font-size: 12px;">v2.2 - AI Summarizer + Lab Analyzer + Email Summarizer</p>
    </div>
    """, unsafe_allow_html=True)

# ========================================
# رأس الصفحة
# ========================================
st.markdown("""
<div class="main-header">
    <h1>📄 ملخص المستندات الذكي</h1>
    <p>رفع ملف، تلخيص، تصنيف، تحليل معامل، وتلخيص إيميلات</p>
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

def summarize_text(text, num_sentences=5):
    sentences = re.split(r'[.!؟]+', text)
    sentences = [s.strip() for s in sentences if len(s.strip()) > 10]
    
    if len(sentences) <= num_sentences:
        return text
    
    word_freq = {}
    stopwords = ['و', 'في', 'من', 'الى', 'على', 'عن', 'مع', 'هذا', 'ذلك', 'كان', 'قد', 'كل', 'لم', 'له', 'ما', 'لا', 'غير', 'بين', 'إن', 'أن', 'ثم', 'حيث', 'حتى', 'عند', 'نحو', 'مثل', 'بعد', 'قبل', 'أثناء', 'دون', 'بسبب', 'رغم', 'معظم', 'بعض', 'أي', 'أو', 'فإن', 'إذا', 'لقد', 'هذه', 'التي', 'الذي']
    
    for sentence in sentences:
        words = re.findall(r'\w+', sentence)
        for word in words:
            word = word.lower()
            if word not in stopwords and len(word) > 2:
                word_freq[word] = word_freq.get(word, 0) + 1
    
    if not word_freq:
        return ' '.join(sentences[:num_sentences])
    
    sorted_words = sorted(word_freq.items(), key=lambda x: x[1], reverse=True)
    top_words = set([word for word, freq in sorted_words[:10]])
    
    sentence_scores = {}
    for sentence in sentences:
        words = re.findall(r'\w+', sentence)
        score = sum(1 for word in words if word.lower() in top_words)
        sentence_scores[sentence] = score
    
    sorted_sentences = sorted(sentence_scores.items(), key=lambda x: x[1], reverse=True)
    summary = ' '.join([s for s, score in sorted_sentences[:num_sentences]])
    
    return summary if summary else ' '.join(sentences[:num_sentences])

def classify_text(text):
    categories = {
        "مالي": ["مال", "اقتصاد", "بنك", "استثمار", "سوق", "أسهم", "دولار", "ربح", "خسارة", "ضريبة"],
        "طبي": ["طبي", "صحي", "مرض", "علاج", "دواء", "جراحة", "تشخيص", "مستشفى", "طبيب", "صحة"],
        "تقني": ["تقني", "برمجة", "حاسوب", "ذكاء اصطناعي", "بيانات", "خوارزمية", "تطبيق", "موقع", "برنامج", "تكنولوجيا"],
        "قانوني": ["قانون", "محكمة", "عقد", "دعوى", "محامي", "حكم", "تشريع", "حقوق", "إجراء", "قضائي"],
        "تعليمي": ["تعليم", "مدرسة", "جامعة", "طالب", "معلم", "منهج", "دراسة", "بحث", "علمي", "أكاديمي"],
        "تسويقي": ["تسويق", "إعلان", "علامة تجارية", "عملاء", "مبيعات", "عرض", "ترويج", "منتج", "خدمة", "سوق"],
        "سياسي": ["سياسي", "حكومة", "برلمان", "انتخاب", "وزير", "رئيس", "قرار", "أمة", "دستور", "حزب"],
        "اجتماعي": ["اجتماعي", "مجتمع", "أسرة", "ثقافة", "سكان", "تنمية", "فقر", "بطالة", "تعاون", "تكافل"],
        "رياضي": ["رياضي", "كرة", "ملعب", "لاعب", "مدرب", "بطولة", "مباراة", "نادي", "جمباز", "سباق"],
        "ديني": ["ديني", "إسلامي", "مسجد", "صلاة", "قرآن", "حديث", "فتوى", "إيمان", "عقيدة", "عبادة"],
        "فني": ["فني", "فن", "موسيقى", "رسم", "مسرح", "سينما", "تمثيل", "غناء", "تشكيل", "أدب"]
    }
    
    text_lower = text.lower()
    category_scores = {}
    for category, keywords in categories.items():
        score = sum(1 for keyword in keywords if keyword in text_lower)
        category_scores[category] = score
    
    if max(category_scores.values()) == 0:
        return "عام", 0.5
    
    best_category = max(category_scores, key=category_scores.get)
    best_score = category_scores[best_category]
    max_possible = len(categories[best_category])
    confidence = best_score / max_possible if max_possible > 0 else 0
    
    return best_category, min(confidence, 0.95)

# ========================================
# دوال تحليل المعامل
# ========================================

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
    found_tests = []
    text_lower = text.lower()
    
    for test_name, test_info in medical_tests.items():
        if test_name.lower() in text_lower:
            pattern = rf"{test_name}[:\s]*([0-9]+\.?[0-9]*)"
            match = re.search(pattern, text, re.IGNORECASE)
            
            if match:
                value = float(match.group(1))
                normal_min, normal_max = test_info["normal"]
                unit = test_info["unit"]
                
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
# دوال تلخيص الإيميلات
# ========================================

def extract_emails(text):
    """استخراج الإيميلات من النص"""
    email_pattern = r'(?:From|Subject|To|Date|Message)[:\s]+.*?(?=(?:From|Subject|To|Date|Message)|\Z)'
    emails_raw = re.findall(email_pattern, text, re.DOTALL | re.IGNORECASE)
    
    if not emails_raw:
        emails_raw = re.split(r'\n\s*\n', text)
    
    emails = []
    for email_text in emails_raw:
        if len(email_text.strip()) > 20:
            subject_match = re.search(r'Subject[:\s]+(.*?)(?:\n|$)', email_text, re.IGNORECASE)
            subject = subject_match.group(1).strip() if subject_match else "بدون موضوع"
            
            body = re.sub(r'(From|Subject|To|Date|Message)[:\s]+.*?\n', '', email_text, flags=re.IGNORECASE)
            body = body.strip()
            
            if body:
                emails.append({
                    "subject": subject,
                    "body": body,
                    "full_text": email_text.strip()
                })
    
    return emails

def classify_email(text):
    text_lower = text.lower()
    
    complaint_keywords = ['شكوى', 'مشكلة', 'خطأ', 'تأخر', 'سيء', 'غير راض', 'فشل', 'عطل']
    inquiry_keywords = ['استفسار', 'سؤال', 'استعلام', 'عندي سؤال', 'أريد معرفة', 'كيف']
    order_keywords = ['طلب', 'شراء', 'أريد', 'احتياج', 'تسجيل', 'اشتراك', 'حجز']
    reply_keywords = ['رد', 'شكراً', 'تم الاستلام', 'حسناً', 'ممتاز', 'تمام']
    
    for word in complaint_keywords:
        if word in text_lower:
            return "شكوى", "email-complaint"
    for word in order_keywords:
        if word in text_lower:
            return "طلب شراء", "email-order"
    for word in inquiry_keywords:
        if word in text_lower:
            return "استفسار", "email-inquiry"
    for word in reply_keywords:
        if word in text_lower:
            return "رد", "email-reply"
    
    return "أخرى", "email-other"

def summarize_email(text, num_sentences=3):
    return summarize_text(text, num_sentences)

# ========================================
# دالة عرض التقرير الجميل
# ========================================
def display_report(summary, label, score, word_count, char_count, sentence_count, clean_text_content):
    report_html = f"""
    <div style="
        background: var(--secondary-background-color);
        padding: 30px;
        border-radius: 15px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.08);
        border-right: 5px solid #667eea;
        direction: rtl;
        text-align: right;
        font-family: 'Cairo', sans-serif;
        max-width: 900px;
        margin: 0 auto;
        color: var(--text-color);
    ">
        <h2 style="text-align: center; color: var(--text-color); border-bottom: 3px solid #667eea; padding-bottom: 15px; margin-bottom: 20px;">
            📄 تقرير تلخيص المستند
        </h2>
        
        <p style="text-align: center; color: var(--text-color); opacity: 0.7; font-size: 14px; margin-bottom: 20px;">
            التاريخ: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
        </p>
        
        <div style="
            background: var(--secondary-background-color);
            padding: 15px 20px;
            border-radius: 10px;
            margin: 15px 0;
            border: 1px solid var(--border-color);
        ">
            <p style="margin: 5px 0;"><strong>🏷️ التصنيف:</strong> {label}</p>
            <p style="margin: 5px 0;"><strong>📊 نسبة الثقة:</strong> {score:.2%}</p>
            <p style="margin: 5px 0;"><strong>📝 عدد الكلمات:</strong> {word_count}</p>
            <p style="margin: 5px 0;"><strong>🔤 عدد الأحرف:</strong> {char_count}</p>
            <p style="margin: 5px 0;"><strong>📖 عدد الجمل:</strong> {sentence_count}</p>
        </div>
        
        <div style="
            background: var(--secondary-background-color);
            padding: 15px 20px;
            border-radius: 10px;
            margin: 15px 0;
            border-right: 4px solid #667eea;
            border: 1px solid var(--border-color);
        ">
            <p style="font-weight: bold; margin: 0 0 5px 0;">📝 الملخص:</p>
            <p style="margin: 0; line-height: 1.8; color: var(--text-color);">{summary}</p>
        </div>
        
        <div style="
            background: var(--secondary-background-color);
            padding: 15px 20px;
            border-radius: 10px;
            margin: 15px 0;
            border-right: 4px solid #667eea;
            border: 1px solid var(--border-color);
        ">
            <p style="font-weight: bold; margin: 0 0 5px 0;">📄 النص الأصلي (مختصر):</p>
            <p style="margin: 0; line-height: 1.8; color: var(--text-color);">{clean_text_content[:500]}{'...' if len(clean_text_content) > 500 else ''}</p>
        </div>
        
        <div style="text-align: center; margin-top: 20px; padding-top: 15px; border-top: 1px solid var(--border-color); color: var(--text-color); opacity: 0.7; font-size: 12px;">
            ✅ تم إنشاء التقرير بواسطة تطبيق ملخص المستندات الذكي<br>
            📌 v2.2 - AI Summarizer
        </div>
    </div>
    """
    return report_html

# ========================================
# واجهة المستخدم (تبويبات)
# ========================================

tab1, tab2, tab3 = st.tabs(["📝 تلخيص وتصنيف", "🧪 تحليل المعامل", "📧 تلخيص الإيميلات"])

# ========================================
# التبويب الأول: تلخيص وتصنيف
# ========================================
with tab1:
    uploaded_file = st.file_uploader("📂 اختر ملف", type=["txt", "pdf", "docx"], key="summarizer")

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
        # التلخيص
        # ========================================
        st.markdown("---")
        st.subheader("📝 الملخص")

        if len(clean_text_content.split()) < 50:
            st.warning("⚠️ النص قصير جداً (أقل من 50 كلمة)")
            summary = clean_text_content
        else:
            with st.spinner("⏳ جاري تلخيص النص..."):
                try:
                    summary = summarize_text(clean_text_content, num_sentences=5)
                    st.success("✅ تم التلخيص بنجاح!")
                except Exception as e:
                    st.error(f"❌ مش قادر ألخص النص: {str(e)}")
                    summary = clean_text_content

        st.markdown(f"""
        <div class="result-card">
            {summary}
        </div>
        """, unsafe_allow_html=True)

        # ========================================
        # التصنيف
        # ========================================
        st.markdown("---")
        st.subheader("🏷️ التصنيف")

        with st.spinner("⏳ جاري تصنيف النص..."):
            try:
                label, score = classify_text(clean_text_content)
                st.success("✅ تم التصنيف بنجاح!")
            except Exception as e:
                st.error(f"❌ مش قادر أصنف النص: {str(e)}")
                label = "غير معروف"
                score = 0

        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f"""
            <div class="metric-box">
                <div class="value">{label}</div>
                <div class="label">التصنيف</div>
            </div>
            """, unsafe_allow_html=True)
        with col2:
            st.markdown(f"""
            <div class="metric-box">
                <div class="value">{score:.2%}</div>
                <div class="label">نسبة الثقة</div>
            </div>
            """, unsafe_allow_html=True)

        # ========================================
        # إحصائيات
        # ========================================
        st.markdown("---")
        st.subheader("📊 إحصائيات")

        word_count = len(clean_text_content.split())
        char_count = len(clean_text_content)
        sentence_count = len(re.findall(r'[.!؟]+', clean_text_content))

        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown(f"""
            <div class="metric-box">
                <div class="value">{word_count}</div>
                <div class="label">عدد الكلمات</div>
            </div>
            """, unsafe_allow_html=True)
        with col2:
            st.markdown(f"""
            <div class="metric-box">
                <div class="value">{char_count}</div>
                <div class="label">عدد الأحرف</div>
            </div>
            """, unsafe_allow_html=True)
        with col3:
            st.markdown(f"""
            <div class="metric-box">
                <div class="value">{sentence_count}</div>
                <div class="label">عدد الجمل</div>
            </div>
            """, unsafe_allow_html=True)

        # ========================================
        # عرض التقرير الجميل
        # ========================================
        st.markdown("---")
        st.subheader("📄 التقرير النهائي")

        report_html = display_report(
            summary, label, score, word_count, char_count, sentence_count, clean_text_content
        )
        st.markdown(report_html, unsafe_allow_html=True)

        # ========================================
        # تحميل التقرير
        # ========================================
        st.markdown("---")
        st.subheader("📥 تحميل التقرير")

        report_text = f"""
        ═══════════════════════════════════════════════════════════════════
                              📄 تقرير تلخيص المستند
        ═══════════════════════════════════════════════════════════════════

        📅 التاريخ          :  {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
        ──────────────────────────────────────────────────────────────────
        🏷️ التصنيف          :  {label}
        📊 نسبة الثقة       :  {score:.2%}
        📝 عدد الكلمات      :  {word_count}
        🔤 عدد الأحرف       :  {char_count}
        📖 عدد الجمل        :  {sentence_count}
        ──────────────────────────────────────────────────────────────────

        📝 الملخص:
        ──────────────────────────────────────────────────────────────────
    """
        for s in summary.replace('؟', '.').split('. '):
            if s.strip():
                report_text += f"        • {s.strip()}.\n"

        report_text += f"""
        ──────────────────────────────────────────────────────────────────

        📄 النص الأصلي (مختصر):
        ──────────────────────────────────────────────────────────────────
        {clean_text_content[:500]}{'...' if len(clean_text_content) > 500 else ''}
        ──────────────────────────────────────────────────────────────────

        ✅ تم إنشاء التقرير بواسطة تطبيق ملخص المستندات الذكي
        📌 v2.2 - AI Summarizer + Lab Analyzer + Email Summarizer
        ═══════════════════════════════════════════════════════════════════
        """

        st.download_button(
            label="📥 تحميل التقرير (TXT)",
            data=report_text,
            file_name=f"تقرير_{uploaded_file.name}.txt",
            mime="text/plain"
        )

    else:
        st.info("⏳ انتظر رفع ملف لتحليله")
        st.markdown("""
        ### 🚀 طريقة الاستخدام:
        1. اضغط على زر **"اختر ملف"**
        2. اختر ملف `.txt` أو `.pdf` أو `.docx`
        3. انتظر لحظات وستظهر النتيجة
        4. يمكنك تحميل التقرير
        """)

# ========================================
# التبويب الثاني: تحليل المعامل
# ========================================
with tab2:
    st.markdown("""
    <div style="background: var(--secondary-background-color); padding: 15px; border-radius: 10px; border-right: 4px solid #667eea; margin-bottom: 20px;">
        <h4 style="margin: 0; color: var(--text-color);">🧪 تحليل المعامل الطبية</h4>
        <p style="margin: 5px 0 0; color: var(--text-color); opacity: 0.7; font-size: 14px;">ارفع تقرير معمل، وسيتم استخراج التحاليل وتقييمها</p>
    </div>
    """, unsafe_allow_html=True)

    uploaded_file_lab = st.file_uploader("📂 اختر ملف تقرير معمل", type=["txt", "pdf", "docx"], key="lab")

    if uploaded_file_lab is not None:
        try:
            text = read_file(uploaded_file_lab)
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
            st.warning("⚠️ لم يتم العثور على تحاليل طبية في هذا الملف.")
        else:
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
            # إحصائيات التحاليل
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
        # تحميل تقرير التحاليل
        # ========================================
        st.markdown("---")
        st.subheader("📥 تحميل تقرير التحاليل")

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
        report_lines.append("  📌 v2.2 - Lab Analyzer")
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

# ========================================
# التبويب الثالث: تلخيص الإيميلات
# ========================================
with tab3:
    st.markdown("""
    <div style="background: var(--secondary-background-color); padding: 15px; border-radius: 10px; border-right: 4px solid #667eea; margin-bottom: 20px;">
        <h4 style="margin: 0; color: var(--text-color);">📧 تلخيص الإيميلات</h4>
        <p style="margin: 5px 0 0; color: var(--text-color); opacity: 0.7; font-size: 14px;">ارفع ملف إيميلات، وسيتم استخراج وتلخيص وتصنيف كل إيميل</p>
    </div>
    """, unsafe_allow_html=True)

    uploaded_file_email = st.file_uploader("📂 اختر ملف إيميلات", type=["txt", "pdf", "docx"], key="email")

    if uploaded_file_email is not None:
        try:
            text = read_file(uploaded_file_email)
        except Exception as e:
            st.error(f"❌ مشكلة في قراءة الملف: {str(e)}")
            st.stop()

        clean_text_content = clean_text(text)

        with st.expander("📄 النص الأصلي"):
            st.text(clean_text_content[:1000] + ("..." if len(clean_text_content) > 1000 else ""))

        # ========================================
        # استخراج الإيميلات
        # ========================================
        st.markdown("---")
        st.subheader("📧 الإيميلات المستخرجة")

        emails = extract_emails(clean_text_content)

        if not emails:
            st.warning("⚠️ لم يتم العثور على إيميلات في هذا الملف.")
        else:
            email_summaries = []
            for i, email in enumerate(emails):
                summary = summarize_email(email['body'], num_sentences=3)
                classification, class_name = classify_email(email['body'])
                
                email_summaries.append({
                    "number": i + 1,
                    "subject": email['subject'],
                    "summary": summary,
                    "classification": classification,
                    "class_name": class_name
                })
                
                st.markdown(f"""
                <div class="email-card {class_name}">
                    <div class="email-subject">📧 {email['subject']}</div>
                    <div class="email-summary"><strong>📝 الملخص:</strong> {summary}</div>
                    <div class="email-classification"><strong>🏷️ التصنيف:</strong> {classification}</div>
                </div>
                """, unsafe_allow_html=True)

            # ========================================
            # إحصائيات الإيميلات
            # ========================================
            st.markdown("---")
            st.subheader("📊 ملخص الإيميلات")

            total = len(emails)
            classifications = {}
            for es in email_summaries:
                classifications[es['classification']] = classifications.get(es['classification'], 0) + 1

            col1, col2 = st.columns(2)
            with col1:
                st.markdown(f"""
                <div class="metric-box">
                    <div class="value">{total}</div>
                    <div class="label">إجمالي الإيميلات</div>
                </div>
                """, unsafe_allow_html=True)
            with col2:
                st.markdown(f"""
                <div class="metric-box">
                    <div class="value" style="font-size: 20px;">{', '.join([f"{k}: {v}" for k, v in classifications.items()])}</div>
                    <div class="label">التصنيفات</div>
                </div>
                """, unsafe_allow_html=True)

        # ========================================
        # تحميل تقرير الإيميلات
        # ========================================
        st.markdown("---")
        st.subheader("📥 تحميل تقرير الإيميلات")

        report_lines = []
        report_lines.append("=" * 65)
        report_lines.append("              📧 تقرير تلخيص الإيميلات")
        report_lines.append("=" * 65)
        report_lines.append("")
        report_lines.append(f"  📅 التاريخ          :  {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        report_lines.append("  " + "-" * 60)
        report_lines.append("")

        if emails:
            for es in email_summaries:
                report_lines.append(f"  📧 الإيميل {es['number']}: {es['subject']}")
                report_lines.append(f"     الملخص: {es['summary']}")
                report_lines.append(f"     التصنيف: {es['classification']}")
                report_lines.append("")
        else:
            report_lines.append("  ⚠️ لم يتم العثور على إيميلات في هذا الملف.")
            report_lines.append("")

        report_lines.append("  " + "-" * 60)
        report_lines.append("")
        report_lines.append("  ✅ تم إنشاء التقرير بواسطة تطبيق تلخيص الإيميلات الذكي")
        report_lines.append("  📌 v2.2 - Email Summarizer")
        report_lines.append("")
        report_lines.append("=" * 65)

        report_text = "\n".join(report_lines)

        st.download_button(
            label="📥 تحميل تقرير الإيميلات (TXT)",
            data=report_text,
            file_name=f"تقرير_إيميلات_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
            mime="text/plain"
        )

    else:
        st.info("⏳ انتظر رفع ملف إيميلات لتحليله")