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
    
    /* ===== تحسين شكل التقرير ===== */
    .report-container {
        background: linear-gradient(145deg, #ffffff, #f8f9fa);
        padding: 30px;
        border-radius: 20px;
        box-shadow: 0 8px 40px rgba(0,0,0,0.08);
        border: 1px solid #e9ecef;
        margin: 20px 0;
        direction: rtl;
        text-align: right;
        font-family: 'Cairo', sans-serif;
    }
    
    .report-header {
        text-align: center;
        padding-bottom: 15px;
        border-bottom: 3px solid #667eea;
        margin-bottom: 20px;
    }
    
    .report-header h2 {
        color: #2d3436;
        font-size: 28px;
        font-weight: 700;
        margin: 0;
    }
    
    .report-header p {
        color: #636e72;
        font-size: 14px;
        margin: 5px 0 0;
    }
    
    .report-section {
        background: #f8f9fa;
        padding: 15px 20px;
        border-radius: 12px;
        margin: 15px 0;
        border-right: 4px solid #667eea;
    }
    
    .report-section-title {
        font-weight: 700;
        color: #2d3436;
        font-size: 16px;
        margin-bottom: 5px;
    }
    
    .report-section-content {
        color: #495057;
        font-size: 15px;
        line-height: 1.8;
    }
    
    .report-metric-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
        gap: 15px;
        margin: 15px 0;
    }
    
    .report-metric-item {
        background: white;
        padding: 12px;
        border-radius: 10px;
        text-align: center;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    }
    
    .report-metric-item .metric-value {
        font-size: 22px;
        font-weight: 700;
        color: #2d3436;
    }
    
    .report-metric-item .metric-label {
        font-size: 12px;
        color: #636e72;
    }
    
    .report-footer {
        text-align: center;
        padding-top: 15px;
        border-top: 1px solid #e9ecef;
        margin-top: 20px;
        color: #636e72;
        font-size: 12px;
    }
    
    .report-badge {
        display: inline-block;
        background: #667eea;
        color: white;
        padding: 4px 14px;
        border-radius: 20px;
        font-size: 14px;
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
            <li>📥 تصدير تقرير منسق</li>
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
    <p>رفع ملف، تلخيص، تصنيف، وتقرير منسق</p>
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
# دالة إنشاء تقرير HTML جميل
# ========================================
def create_report_html(text, summary, label, score, word_count, char_count, sentence_count):
    """إنشاء تقرير بتنسيق HTML جميل"""
    
    # تقسيم الملخص إلى نقاط
    summary_points = summary.replace('؟', '.').split('. ')
    summary_bullets = []
    for s in summary_points:
        if s.strip():
            summary_bullets.append(f"<li>{s.strip()}.</li>")
    
    # تقطيع النص الأصلي
    text_preview = text[:500]
    if len(text) > 500:
        text_preview += "..."
    
    # تنسيق النص الأصلي (تحويل الأسطر الجديدة إلى <br>)
    text_preview_html = text_preview.replace('\n', '<br>')
    
    html = f"""
    <div class="report-container">
        <div class="report-header">
            <h2>📄 تقرير تلخيص المستند</h2>
            <p>تم الإنشاء: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
        </div>
        
        <div class="report-metric-grid">
            <div class="report-metric-item">
                <div class="metric-value">🏷️ {label}</div>
                <div class="metric-label">التصنيف</div>
            </div>
            <div class="report-metric-item">
                <div class="metric-value">{score:.2%}</div>
                <div class="metric-label">نسبة الثقة</div>
            </div>
            <div class="report-metric-item">
                <div class="metric-value">{word_count}</div>
                <div class="metric-label">عدد الكلمات</div>
            </div>
            <div class="report-metric-item">
                <div class="metric-value">{char_count}</div>
                <div class="metric-label">عدد الأحرف</div>
            </div>
            <div class="report-metric-item">
                <div class="metric-value">{sentence_count}</div>
                <div class="metric-label">عدد الجمل</div>
            </div>
        </div>
        
        <div class="report-section">
            <div class="report-section-title">📝 الملخص</div>
            <div class="report-section-content">
                <ul style="list-style: none; padding: 0; margin: 0;">
                    {''.join(summary_bullets)}
                </ul>
            </div>
        </div>
        
        <div class="report-section">
            <div class="report-section-title">📄 النص الأصلي (مختصر)</div>
            <div class="report-section-content">
                {text_preview_html}
            </div>
        </div>
        
        <div class="report-footer">
            <span class="report-badge">AI Summarizer v2.0</span>
            <span style="margin: 0 10px;">|</span>
            تم إنشاء التقرير بواسطة تطبيق ملخص المستندات الذكي
        </div>
    </div>
    """
    return html

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
    # عرض التقرير المنمنم (بشكل صحيح)
    # ========================================
    st.markdown("---")
    st.subheader("📄 التقرير النهائي")

    # إنشاء التقرير HTML
    report_html = create_report_html(
        clean_text_content, summary, label, score,
        word_count, char_count, sentence_count
    )
    
    # ✅ عرض التقرير بشكل صحيح (st.markdown مع unsafe_allow_html=True)
    st.markdown(report_html, unsafe_allow_html=True)

    # ========================================
    # تحميل التقرير (TXT)
    # ========================================
    st.markdown("---")
    st.subheader("📥 تحميل التقرير")

    # إنشاء نسخة نصية للتحميل
    report_lines = []
    report_lines.append("╔══════════════════════════════════════════════════════════════════════╗")
    report_lines.append("║                     📄 تقرير تلخيص المستند                         ║")
    report_lines.append("╚══════════════════════════════════════════════════════════════════════╝")
    report_lines.append("")
    report_lines.append(f"  📅 التاريخ          :  {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    report_lines.append("  ──────────────────────────────────────────────────────────────────")
    report_lines.append(f"  🏷️ التصنيف          :  {label}")
    report_lines.append(f"  📊 نسبة الثقة       :  {score:.2%}")
    report_lines.append(f"  📝 عدد الكلمات      :  {word_count}")
    report_lines.append(f"  🔤 عدد الأحرف       :  {char_count}")
    report_lines.append(f"  📖 عدد الجمل        :  {sentence_count}")
    report_lines.append("  ──────────────────────────────────────────────────────────────────")
    report_lines.append("")
    report_lines.append("  📝 الملخص:")
    report_lines.append("  ──────────────────────────────────────────────────────────────────")
    for s in summary.replace('؟', '.').split('. '):
        if s.strip():
            report_lines.append(f"    • {s.strip()}.")
    report_lines.append("")
    report_lines.append("  ──────────────────────────────────────────────────────────────────")
    report_lines.append("")
    report_lines.append("  📄 النص الأصلي (مختصر):")
    report_lines.append("  ──────────────────────────────────────────────────────────────────")
    text_preview = clean_text_content[:500]
    if len(clean_text_content) > 500:
        text_preview += "..."
    report_lines.append(f"    {text_preview.replace(chr(10), chr(10) + '    ')}")
    report_lines.append("")
    report_lines.append("  ──────────────────────────────────────────────────────────────────")
    report_lines.append("")
    report_lines.append("  ✅ تم إنشاء التقرير بواسطة تطبيق ملخص المستندات الذكي")
    report_lines.append("  📌 v2.0 - AI Summarizer")
    report_lines.append("")
    report_lines.append("╔══════════════════════════════════════════════════════════════════════╗")
    report_lines.append("║                       نهاية التقرير                                ║")
    report_lines.append("╚══════════════════════════════════════════════════════════════════════╝")

    report_text = "\n".join(report_lines)

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