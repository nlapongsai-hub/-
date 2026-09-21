import io
import json
import streamlit as st
from docxtpl import DocxTemplate
from google import genai
from google.genai import types

SYSTEM_PASSCODE = "0863449483"

st.set_page_config(
    page_title="EduPlan Pro | ระบบจัดทำโครงการสอนอัจฉริยะ",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Prompt:wght@300;400;500;600;700&display=swap');
    * { font-family: 'Prompt', sans-serif !important; }
    .hero-banner {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 50%, #00c6ff 100%);
        border-radius: 20px;
        padding: 35px 30px;
        color: white;
        margin-bottom: 25px;
        box-shadow: 0 10px 25px rgba(30, 60, 114, 0.15);
    }
    .hero-title { font-size: 2.2rem; font-weight: 700; margin-bottom: 10px; }
    .hero-desc { font-size: 1.05rem; opacity: 0.92; font-weight: 300; line-height: 1.6; }
    .box-header { font-size: 1.25rem; font-weight: 600; color: #1a2a4b; margin-bottom: 15px; }
    .stButton>button[kind="primary"] {
        background: linear-gradient(90deg, #1e3c72 0%, #2a5298 100%);
        border: none; border-radius: 12px; font-weight: 600; font-size: 1.1rem; padding: 14px 28px;
        box-shadow: 0 8px 18px rgba(30, 60, 114, 0.25);
    }
    .copyright-card {
        background: #f8fafc; border-left: 4px solid #1e3c72; border-radius: 10px; padding: 14px 16px; margin-top: 20px;
    }
    .copyright-title { font-size: 0.85rem; font-weight: 700; color: #1e3c72; text-transform: uppercase; }
    .copyright-text { font-size: 0.8rem; color: #475569; margin-top: 4px; line-height: 1.5; }
</style>
""", unsafe_allow_html=True)

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

def login_gate():
    col_l, col_center, col_r = st.columns([1, 1.8, 1])
    with col_center:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.markdown("""
        <div style="text-align: center; margin-bottom: 25px;">
            <div style="font-size: 4rem; margin-bottom: 10px;">🏛️</div>
            <h2 style="font-weight: 700; color: #1e3c72;">EduPlan Pro (AI Engine)</h2>
            <p style="color: #64748b; font-size: 0.95rem;">ระบบจัดทำโครงการสอนและแผนการจัดการเรียนรู้มาตรฐาน สอศ.</p>
        </div>
        """, unsafe_allow_html=True)
        passcode = st.text_input("🔑 กรุณากรอกรหัสปลดล็อกสิทธิ์เข้าใช้งาน:", type="password", placeholder="กรอกรหัสผ่าน 10 หลัก")
        if st.button("🔓 ยืนยันสิทธิ์และเข้าสู่ระบบ", type="primary", use_container_width=True):
            if passcode == SYSTEM_PASSCODE:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("❌ รหัสผ่านไม่ถูกต้อง กรุณาติดต่อผู้รับผิดชอบระบบ")
        st.markdown("""
        <div class="copyright-card">
            <div class="copyright-title">ลิขสิทธิ์และการพัฒนา</div>
            <div class="copyright-text">
                นวัตกรรมระบบปัญญาประดิษฐ์สกัดโครงสร้างหลักสูตรอาชีวศึกษา<br>
                <b>นายณัฐวุฒิ หล้าปงสาย</b> ครูผู้ช่วย วิทยาลัยเทคนิคจันทบุรี<br>
                <span style="color: #94a3b8;">All Rights Reserved © 2026</span>
            </div>
        </div>
        """, unsafe_allow_html=True)

if not st.session_state.authenticated:
    login_gate()
    st.stop()

def extract_course_plan(api_key: str, file_bytes: bytes, mime_type: str, total_weeks: int):
    client = genai.Client(api_key=api_key)
    prompt = f"""
คุณคือระบบสกัดข้อมูลเอกสารทางการศึกษา (Data Extraction Engine) 
จงอ่านไฟล์ 'แผนการจัดการเรียนรู้' ที่แนบมานี้อย่างละเอียด แล้วทำการ "คัดลอก (Extract)" ข้อมูลจริงจากแต่ละสัปดาห์ (สัปดาห์ที่ 1 ถึง {total_weeks}) เพื่อนำไปลงตารางโครงการสอน

ข้อกำหนด:
1. สกัดข้อมูลครบถ้วน {total_weeks} สัปดาห์
2. คัดลอกข้อความจริงจากหัวข้อในแผนการสอน ห้ามแต่งข้อมูลขึ้นมาใหม่:
   - week: ตัวเลขสัปดาห์ (1 ถึง {total_weeks})
   - unit_name: ข้อความจาก "ชื่อหน่วยการเรียนรู้"
   - topic_name: ข้อความจาก "ชื่อเรื่อง/งาน"
   - teaching_points: รายการจุดประสงค์เชิงพฤติกรรมหรือสมรรถนะประจำหน่วย 3-4 ข้อ
   - activities: รายการกิจกรรมหลัก 4 ขั้นตอนจากหัวข้อกิจกรรมการเรียนรู้
   - media: รายการสื่อและแหล่งการเรียนรู้
   - assessment: รายการเครื่องมือและวิธีการวัดประเมินผล
3. ตอบกลับเป็น Pure JSON Array เท่านั้น ห้ามมี Markdown Backticks
    """
    response = client.models.generate_content(
        model="gemini-3.6-flash",
        contents=[
            types.Part.from_bytes(data=file_bytes, mime_type=mime_type),
            prompt
        ],
        config=types.GenerateContentConfig(response_mime_type="application/json")
    )
    return json.loads(response.text)

with st.sidebar:
    st.markdown("""
    <div style="display: flex; align-items: center; gap: 10px; margin-bottom: 10px;">
        <span style="font-size: 1.8rem;">⚙️</span>
        <h3 style="margin: 0; color: #1e3c72; font-weight: 700;">การตั้งค่าระบบ</h3>
    </div>
    """, unsafe_allow_html=True)
    st.markdown('<div style="background: #e0f2fe; color: #0369a1; padding: 6px 12px; border-radius: 20px; font-size: 0.85rem; font-weight: 600; width: fit-content; margin-bottom: 15px;">🟢 STATUS: AUTHORIZED</div>', unsafe_allow_html=True)
    api_key = st.text_input("🔑 Google Gemini API Key:", type="password", placeholder="AIzaSy...")
    st.markdown("[✨ รับ Gemini API Key สำหรับประมวลผล](https://aistudio.google.com/)")
    st.markdown("---")
    st.markdown("#### 👤 ข้อมูลผู้จัดทำ")
    teacher_name = st.text_input("ชื่อ-สกุล ครูผู้สอน:", value="นายณัฐวุฒิ หล้าปงสาย")
    dept_name = st.text_input("แผนกวิชา:", value="การจัดการโลจิสติกส์และซัพพลายเชน")
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🔒 ล็อกระบบและออกจากระบบ", use_container_width=True):
        st.session_state.authenticated = False
        st.rerun()
    st.markdown("""
    <div class="copyright-card">
        <div class="copyright-title">ลิขสิทธิ์โปรแกรม</div>
        <div class="copyright-text">
            <b>นายณัฐวุฒิ หล้าปงสาย</b><br>
            ครูผู้ช่วย วิทยาลัยเทคนิคจันทบุรี<br>
            สำนักงานคณะกรรมการการอาชีวศึกษา
        </div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("""
<div class="hero-banner">
    <div class="hero-title">📋 ระบบจัดทำโครงการสอนอัตโนมัติ</div>
    <div class="hero-desc">
        วิเคราะห์และสกัดข้อมูลจากแผนการจัดการเรียนรู้ ลงสู่เทมเพลตโครงการสอนของวิทยาลัยแบบอัตโนมัติ<br>
        ถูกต้องตามมาตรฐาน สอศ. 100% จัดโครงสร้างตารางและควบคุมหน้ากระดาษอย่างเป็นระเบียบ
    </div>
</div>
""", unsafe_allow_html=True)

col1, col2 = st.columns([1, 1], gap="large")

with col1:
    st.markdown('<div class="box-header">📁 1. แนบเอกสารต้นทางและฟอร์ม</div>', unsafe_allow_html=True)
    template_file = st.file_uploader("แบบฟอร์มโครงการสอนวิทยาลัย (template.docx):", type=["docx"])
    plan_file = st.file_uploader("แผนการจัดการเรียนรู้ฉบับเต็ม (PDF หรือ Word):", type=["pdf", "docx"])

with col2:
    st.markdown('<div class="box-header">🎯 2. ระดับชั้นและปีการศึกษา</div>', unsafe_allow_html=True)
    c_deg, c_yr = st.columns(2)
    with c_deg:
        degree = st.selectbox("ระดับคุณวุฒิการศึกษา:", ["ปวส.", "ปวช."])
    with c_yr:
        if degree == "ปวช.":
            year = st.selectbox("ระดับชั้นปี:", [1, 2, 3])
            default_weeks = 18
            deg_title = "ประกาศนียบัตรวิชาชีพ (ปวช.)"
        else:
            year = st.selectbox("ระดับชั้นปี:", [1, 2])
            default_weeks = 15
            deg_title = "ประกาศนียบัตรวิชาชีพชั้นสูง (ปวส.)"
            
    weeks_input = st.number_input("จำนวนสัปดาห์ตลอดภาคเรียน:", min_value=1, max_value=22, value=default_weeks)
    st.markdown(f'<div style="background-color: #f1f5f9; padding: 12px 16px; border-radius: 12px; border-left: 4px solid #3b82f6; margin-top: 10px;"><b>📌 ข้อกำหนด:</b> {deg_title} ปีที่ {year} | {weeks_input} สัปดาห์</div>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

if st.button("🚀 สกัดข้อมูลและสร้างเอกสารโครงการสอน (Generate Word)", type="primary", use_container_width=True):
    if not api_key:
        st.warning("⚠️ กรุณากรอก Gemini API Key ที่แถบด้านซ้ายก่อนเริ่มการประมวลผล")
    elif not template_file or not plan_file:
        st.warning("⚠️ กรุณาอัปโหลดทั้งแบบฟอร์ม (template.docx) และแผนการจัดการเรียนรู้ให้ครบถ้วน")
    else:
        with st.status("⚡ กำลังประมวลผลเอกสารทางการศึกษา...", expanded=True) as status:
            try:
                st.write("🔍 กำลังอ่านและจัดโครงสร้างไฟล์แผนการสอน...")
                mime = "application/pdf" if plan_file.name.endswith(".pdf") else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                plan_bytes = plan_file.read()
                
                st.write("🤖 กำลังสกัด Teaching Points, กิจกรรม, สื่อ และการวัดผล...")
                extracted = extract_course_plan(api_key=api_key, file_bytes=plan_bytes, mime_type=mime, total_weeks=weeks_input)
                
                st.write("📝 กำลังเรนเดอร์ข้อมูลลงในแม่แบบเอกสาร Word...")
                CHECKED, UNCHECKED = "☑", "☐"
                context = {
                    "degree_title": deg_title,
                    "is_pvc": (degree == "ปวช."),
                    "cb_y1": CHECKED if year == 1 else UNCHECKED,
                    "cb_y2": CHECKED if year == 2 else UNCHECKED,
                    "cb_y3": CHECKED if year == 3 else UNCHECKED,
                    "plans": extracted
                }
                
                doc = DocxTemplate(template_file)
                doc.render(context)
                
                output = io.BytesIO()
                doc.save(output)
                output.seek(0)
                
                status.update(label="✅ ดำเนินการสร้างโครงการสอนสำเร็จเรียบร้อย!", state="complete", expanded=False)
                st.balloons()
                st.success("🎉 ระบบได้จัดทำโครงการสอนฉบับสมบูรณ์เรียบร้อยแล้ว")
                st.download_button(
                    label=f"📥 ดาวน์โหลดเอกสารโครงการสอน ({degree} {year}).docx",
                    data=output,
                    file_name=f"โครงการสอน_{degree}_{year}_{dept_name}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                status.update(label="❌ เกิดข้อผิดพลาดในการประมวลผล", state="error")
                st.error(f"รายละเอียดข้อผิดพลาด: {str(e)}")
