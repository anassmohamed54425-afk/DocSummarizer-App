import streamlit as st
import PyPDF2
import io
import re
from docx import Document
import datetime
import arabic_reshaper
from bidi.algorithm import get_display
from fpdf import FPDF

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
    
    * {
        font-family: 'Cairo', sans-serif;
    }
    
    .main-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 40px;
        border-radius: 20px;
        text-align: center;
        color: white;
        margin-bottom: 30px;
        box-shadow: 0 10px 40px rgba(102, 126, 234, 0.4);
    }
    
    .main-header h1 {
        font-size: 48px;
        font-weight: 700;
        margin: 0;
    }
    
    .main-header p {
        font-size: 18px;
        opacity: 0.9;
        margin: 10px 0 0;
    }
    
    .result-card {
        background: white;
        padding: 25px;
        border-radius: 15px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.08);
        margin: 20px 0;
        border-right: 6px solid #667eea;
        transition: transform 0.2s;
    }
    
    .result-card:hover {
        transform: translateY(-3px);
        box-shadow: 0 8px 30px rgba(0,0,0,0.12);
    }
    
    .metric-box {
        background: #f8f9fa;
        padding: 20px;
        border-radius: 12px;
        text-align: center;
        transition: all 0.2s;
    }
    
    .metric-box:hover {
        background: #e9ecef;
        transform: scale(1.02);
    }
    
    .metric-box .value {
        font-size: 28px;
        font-weight: 700;
        color: #2d3436;
    }
    
    .metric-box .label {
        font-size: 14px;
        color: #636e72;
        margin-top: 5px;
    }
    
    .stButton > button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        padding: 12px 30px;
        border-radius: 10px;
        font-weight: 600;
        font-size: 16px;
        transition: all 0.3s;
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
        <h3 style="color: #2d3436; border-bottom: 2px solid #667eea; padding-bottom: 10px;">📊 تحليل المستندات</h3>
        <p style="color: #636e72; font-size: 14px;">تطبيق ذكي لتحليل وتلخيص المستندات النصية</p>
        <hr>
        <h4>⚡ الميزات</h4>
        <ul style="color: #2d3436; font-size: 14px; list-style: none; padding: 0;">
            <li>📄 قراءة TXT, PDF, DOCX</li>
            <li>📝 تلخيص ذكي</li>
            <li>🏷️ تصنيف تلقائي</li>
            <li>📊 إحصائيات متقدمة</li>
            <li>📥 تصدير PDF و TXT</li>
        </ul>
        <hr>
        <h4>📌 الإصدار</h4>
        <p style="color: #636e72; font-size: 12px;">v2.0 - AI Summarizer</p>
    </div>
    """, unsafe_allow_html=True)

# ========================================
# رأس الصفحة
# ========================================
st.markdown("""
<div class="main-header">
    <h1>📄 ملخص المستندات الذكي</h1>
    <p>رفع ملف، تلخيص، تصنيف، وتقرير PDF</p>
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
# دالة إنشاء PDF (باستخدام fpdf2 مع دعم العربية)
# ========================================
def create_pdf(text, summary, label, score, word_count, char_count, sentence_count):
    class PDF(FPDF):
        def header(self):
            self.set_font('Arial', 'B', 16)
            self.cell(0, 10, 'تقرير تحليل المستند', 0, 1, 'C')
            self.ln(5)
    
    pdf = PDF()
    pdf.add_page()
    
    # دالة لتنسيق النص العربي
    def format_arabic(txt):
        try:
            reshaped = arabic_reshaper.reshape(txt)
            return get_display(reshaped)
        except:
            return txt
    
    # محاولة إضافة خط عربي
    try:
        # استخدام DejaVu (مدمج في fpdf2)
        pdf.add_font('DejaVu', '', 'DejaVuSans.ttf', uni=True)
        font_name = 'DejaVu'
    except:
        try:
            # محاولة استخدام Arial Unicode
            pdf.add_font('ArialUnicode', '', 'ArialUnicodeMS.ttf', uni=True)
            font_name = 'ArialUnicode'
        except:
            try:
                # محاولة استخدام NotoSans
                pdf.add_font('NotoSans', '', 'NotoSans-Regular.ttf', uni=True)
                font_name = 'NotoSans'
            except:
                # لو كل حاجة فشلت، استخدم Helvetica (مش هيدعم العربية)
                font_name = 'Helvetica'
    
    pdf.set_font(font_name, size=12)
    
    # التاريخ
    pdf.cell(0, 10, format_arabic(f"التاريخ: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}"), 0, 1, 'C')
    pdf.ln(5)
    
    # النتيجة
    pdf.set_font(font_name, size=14)
    pdf.cell(0, 10, format_arabic(f"التصنيف: {label}"), 0, 1, 'C')
    pdf.cell(0, 10, format_arabic(f"نسبة الثقة: {score:.2%}"), 0, 1, 'C')
    pdf.cell(0, 10, format_arabic(f"عدد الكلمات: {word_count}"), 0, 1, 'C')
    pdf.cell(0, 10, format_arabic(f"عدد الأحرف: {char_count}"), 0, 1, 'C')
    pdf.cell(0, 10, format_arabic(f"عدد الجمل: {sentence_count}"), 0, 1, 'C')
    pdf.ln(5)
    
    # الملخص
    pdf.set_font(font_name, size=12)
    pdf.multi_cell(0, 10, format_arabic(f"الملخص:\n{summary}"))
    pdf.ln(5)
    
    # النص الأصلي (مختصر)
    pdf.set_font(font_name, size=10)
    pdf.multi_cell(0, 8, format_arabic(f"النص الأصلي (مختصر):\n{text[:500]}..."))
    
    # حفظ PDF
    try:
        pdf_output = pdf.output(dest='S')
        return io.BytesIO(pdf_output.encode('latin1'))
    except Exception as e:
        st.error(f"❌ مشكلة في حفظ PDF: {str(e)}")
        return None

# ========================================
# واجهة المستخدم
# ========================================
uploaded_file = st.file_uploader("📂 اختر ملف", type=["txt", "pdf", "docx"])

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
    # تحميل التقرير (PDF + TXT)
    # ========================================
    st.markdown("---")
    st.subheader("📥 تحميل التقرير")

    report_text = f"""
    ═══════════════════════════════════════════════════════════════
                          📄 تقرير تلخيص المستند
    ═══════════════════════════════════════════════════════════════

    التصنيف: {label} (نسبة الثقة: {score:.2%})
    عدد الكلمات: {word_count}
    عدد الأحرف: {char_count}
    عدد الجمل: {sentence_count}

    ═══════════════════════════════════════════════════════════════
    الملخص:
    ═══════════════════════════════════════════════════════════════

    {summary}

    ═══════════════════════════════════════════════════════════════
    النص الأصلي (مختصر):
    ═══════════════════════════════════════════════════════════════

    {clean_text_content[:500]}{'...' if len(clean_text_content) > 500 else ''}

    ═══════════════════════════════════════════════════════════════
    ✅ تم إنشاء التقرير بواسطة تطبيق ملخص المستندات الذكي
    ═══════════════════════════════════════════════════════════════
    """

    col1, col2 = st.columns(2)
    
    with col1:
        st.download_button(
            label="📥 تحميل التقرير (TXT)",
            data=report_text,
            file_name=f"تقرير_{uploaded_file.name}.txt",
            mime="text/plain"
        )
    
    with col2:
        with st.spinner("⏳ جاري إنشاء PDF..."):
            try:
                pdf_buffer = create_pdf(
                    clean_text_content, summary, label, score,
                    word_count, char_count, sentence_count
                )
                if pdf_buffer:
                    st.download_button(
                        label="📥 تحميل التقرير (PDF)",
                        data=pdf_buffer,
                        file_name=f"تقرير_{uploaded_file.name}.pdf",
                        mime="application/pdf"
                    )
                else:
                    st.warning("⚠️ لا يمكن إنشاء PDF، استخدم TXT")
            except Exception as e:
                st.error(f"❌ مشكلة في إنشاء PDF: {str(e)}")

else:
    st.info("⏳ انتظر رفع ملف لتحليله")
    st.markdown("""
    ### 🚀 طريقة الاستخدام:
    1. اضغط على زر **"اختر ملف"**
    2. اختر ملف `.txt` أو `.pdf` أو `.docx`
    3. انتظر لحظات وستظهر النتيجة
    4. يمكنك تحميل التقرير بصيغة TXT أو PDF
    """)