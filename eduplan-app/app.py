import io
import json
import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from google import genai
from google.genai import types

# ----------------------------------------------------
# 1. การตั้งค่าระบบและธีมพรีเมียม
# ----------------------------------------------------
SYSTEM_PASSCODE = "0863449483"

st.set_page_config(
    page_title="EduPlan Pro | ระบบจัดทำโครงการสอนอัตโนมัติ",
    page_icon="🎓",
    layout="wide"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Prompt:wght@300;400;500;600;700&display=swap');
    * { font-family: 'Prompt', sans-serif !important; }
    .hero-banner {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 50%, #00c6ff 100%);
        border-radius: 18px; padding: 30px; color: white; margin-bottom: 25px;
    }
    .hero-title { font-size: 2.2rem; font-weight: 700; margin-bottom: 8px; }
    .hero-desc { font-size: 1.05rem; opacity: 0.95; font-weight: 300; line-height: 1.5; }
    .box-header { font-size: 1.2rem; font-weight: 600; color: #1e3c72; margin-bottom: 12px; }
    .copyright-card {
        background: #f8fafc; border-left: 4px solid #1e3c72; border-radius: 8px; padding: 12px 16px; margin-top: 20px;
    }
</style>
""", unsafe_allow_html=True)

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

def login_gate():
    c1, c2, c3 = st.columns([1, 1.8, 1])
    with c2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.markdown("""
        <div style="text-align: center; margin-bottom: 20px;">
            <div style="font-size: 3.5rem;">🏛️</div>
            <h2 style="font-weight: 700; color: #1e3c72;">EduPlan Pro (AI Engine)</h2>
            <p style="color: #64748b;">ระบบสกัดตารางวิเคราะห์งานสู่โครงการสอนมาตรฐาน สอศ.</p>
        </div>
        """, unsafe_allow_html=True)
        passcode = st.text_input("🔑 รหัสปลดล็อกสิทธิ์เข้าใช้งาน:", type="password")
        if st.button("🔓 ปลดล็อกและเข้าสู่ระบบ", type="primary", use_container_width=True):
            if passcode == SYSTEM_PASSCODE:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("❌ รหัสผ่านไม่ถูกต้อง")
        st.markdown("""
        <div class="copyright-card">
            <b>ลิขสิทธิ์และการพัฒนา</b><br>
            นายณัฐวุฒิ หล้าปงสาย ครูผู้ช่วย แผนกวิชาการจัดการโลจิสติกส์และซัพพลายเชน<br>
            วิทยาลัยเทคนิคจันทบุรี | All Rights Reserved © 2026
        </div>
        """, unsafe_allow_html=True)

if not st.session_state.authenticated:
    login_gate()
    st.stop()

# ----------------------------------------------------
# 2. ฟังก์ชัน AI ประมวลผลตารางวิเคราะห์งาน (Gemini 3.6 Flash)
# ----------------------------------------------------
def extract_from_analysis_doc(api_key: str, file_bytes: bytes, mime_type: str, total_weeks: int, course_name: str):
    client = genai.Client(api_key=api_key)
    
    prompt = f"""
คุณคือผู้เชี่ยวชาญการจัดทำหลักสูตรและโครงการสอนอาชีวศึกษา (ปวช./ปวส.)
จงอ่านเนื้อหาจากไฟล์ 'ตารางวิเคราะห์งาน / ตารางวิเคราะห์หน่วยการเรียนรู้' ของวิชา '{course_name}' ที่แนบมานี้
แล้วทำการสังเคราะห์และกระจายเนื้อหาจัดทำเป็น 'ตารางโครงการสอนต่อภาคเรียน' ให้ครบจำนวน {total_weeks} สัปดาห์ (สัปดาห์ที่ 1 ถึง {total_weeks})

เกณฑ์การจัดทำข้อมูลแต่ละสัปดาห์:
1. week: ลำดับสัปดาห์ (1 ถึง {total_weeks})
2. topic_full: ข้อความ 2 ส่วน (มีขึ้นบรรทัดใหม่)
   - บรรทัดแรก: หน่วยที่ ... และชื่อหน่วย (แปลงมาจาก งานหลัก / Duty)
   - บรรทัดสอง: เรื่อง ... (แปลงมาจาก งานย่อย / Task)
3. teaching_points: จุดประสงค์เชิงพฤติกรรม 3-4 ข้อ สังเคราะห์จาก 'สมรรถนะย่อย', 'ความรู้', และ 'ทักษะ'
4. activities: กิจกรรมการจัดการเรียนรู้เชิงรุก (Active Learning 4 ขั้นตอน) สอดคล้องกับทักษะปฏิบัติ
5. media: สื่อและแหล่งการเรียนรู้ (เช่น ใบงาน, สไลด์, โปรแกรมจำลอง, แพลตฟอร์มดิจิทัล)
6. assessment: เครื่องมือและวิธีการวัดประเมินผล (เช่น แบบทดสอบ, Rubric ประเมินทักษะ, สังเกตพฤติกรรม)

ตอบกลับเป็น Pure JSON Array โดยตรง ห้ามมีเครื่องหมาย markdown code block ครอบ
"""

    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=[
            types.Part.from_bytes(data=file_bytes, mime_type=mime_type),
            prompt
        ],
        config=types.GenerateContentConfig(response_mime_type="application/json")
    )
    return json.loads(response.text)

def set_cell_font(cell, font_name="TH SarabunPSK", font_size=Pt(14), bold=False):
    for p in cell.paragraphs:
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.0
        for run in p.runs:
            run.font.name = font_name
            run.font.size = font_size
            run.font.bold = bold
            rPr = run._r.get_or_add_rPr()
            rFonts = OxmlElement('w:rFonts')
            rFonts.set(qn('w:ascii'), font_name)
            rFonts.set(qn('w:hAnsi'), font_name)
            rFonts.set(qn('w:cs'), font_name)
            rPr.append(rFonts)

def replace_placeholders_in_doc(doc, replacements):
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, val in replacements.items():
                    if key in cell.text:
                        for p in cell.paragraphs:
                            if key in p.text:
                                p.text = p.text.replace(key, str(val))
    for p in doc.paragraphs:
        for key, val in replacements.items():
            if key in p.text:
                p.text = p.text.replace(key, str(val))

# ----------------------------------------------------
# 3. ส่วนควบคุม UI
# ----------------------------------------------------
with st.sidebar:
    st.markdown("### ⚙️ การตั้งค่าระบบ")
    st.success("🟢 STATUS: AUTHORIZED")
    api_key = st.text_input("🔑 Gemini API Key:", type="password", placeholder="AIzaSy...")
    st.markdown("[👉 รับ API Key จาก Google AI Studio](https://aistudio.google.com/)")
    st.markdown("---")
    st.markdown("#### 👤 ข้อมูลครูผู้สอน")
    teacher_name = st.text_input("ชื่อ-สกุล:", value="นายณัฐวุฒิ หล้าปงสาย")
    dept_name = st.text_input("แผนกวิชา:", value="การจัดการโลจิสติกส์และซัพพลายเชน")
    if st.button("🔒 ล็อกระบบกลับ"):
        st.session_state.authenticated = False
        st.rerun()

st.markdown("""
<div class="hero-banner">
    <div class="hero-title">📋 ระบบจัดทำโครงการสอนอัตโนมัติ</div>
    <div class="hero-desc">
        แปลงตารางวิเคราะห์งาน / วิเคราะห์หน่วย สู่ตารางโครงการสอนตามแบบฟอร์มวิทยาลัยโดยตรง<br>
        จัดรูปแบบอักษรและตารางมาตรฐาน สอศ. ไม่บีบ ไม่ล้นหน้า
    </div>
</div>
""", unsafe_allow_html=True)

col1, col2 = st.columns(2, gap="large")

with col1:
    st.markdown('<div class="box-header">📁 1. เอกสารนำเข้า</div>', unsafe_allow_html=True)
    template_file = st.file_uploader("แบบฟอร์มวิทยาลัย (template.docx):", type=["docx"])
    analysis_file = st.file_uploader("ไฟล์ตารางวิเคราะห์งาน (Word / PDF):", type=["docx", "pdf"])

with col2:
    st.markdown('<div class="box-header">🎯 2. ข้อมูลวิชาและระดับชั้น</div>', unsafe_allow_html=True)
    course_code = st.text_input("รหัสวิชา:", value="31401-2007")
    course_name = st.text_input("ชื่อวิชา:", value="การจัดการโลจิสติกส์และซัพพลายเชน")
    
    cd1, cd2 = st.columns(2)
    with cd1:
        degree = st.selectbox("ระดับคุณวุฒิ:", ["ปวส.", "ปวช."])
    with cd2:
        year = st.selectbox("ชั้นปี:", [1, 2, 3] if degree == "ปวช." else [1, 2])
        
    weeks_target = 18 if degree == "ปวช." else 15
    weeks_input = st.number_input("จำนวนสัปดาห์ต่อภาคเรียน:", min_value=1, max_value=22, value=weeks_target)
    hours_per_week = st.number_input("จำนวนชั่วโมงต่อสัปดาห์:", min_value=1, max_value=10, value=4)
    semester = st.text_input("ภาคเรียนที่:", value="1/2569")

st.markdown("<br>", unsafe_allow_html=True)

# ----------------------------------------------------
# 4. ประมวลผลและสร้างไฟล์ Word
# ----------------------------------------------------
if st.button("🚀 ประมวลผลและสร้างโครงการสอน (Generate)", type="primary", use_container_width=True):
    if not api_key:
        st.warning("กรุณาระบุ Gemini API Key ในแถบซ้ายมือก่อนเริ่ม")
    elif not template_file or not analysis_file:
        st.warning("กรุณาแนบทั้ง 'แบบฟอร์มวิทยาลัย (template.docx)' และ 'ไฟล์ตารางวิเคราะห์งาน'")
    else:
        with st.status("⚡ กำลังประมวลผลโครงการสอน...", expanded=True) as status:
            try:
                st.write("📖 กำลังอ่านโครงสร้างข้อมูลจากตารางวิเคราะห์งาน...")
                mime = "application/pdf" if analysis_file.name.endswith(".pdf") else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                file_bytes = analysis_file.read()
                
                st.write("🤖 ส่งต่อ Gemini 3.6 Flash สังเคราะห์แผนการสอนรายสัปดาห์...")
                plans = extract_from_analysis_doc(
                    api_key=api_key,
                    file_bytes=file_bytes,
                    mime_type=mime,
                    total_weeks=weeks_input,
                    course_name=course_name
                )
                
                st.write("📝 บรรจุข้อมูลและแทนที่หัวกระดาษลงในแบบฟอร์มวิทยาลัย...")
                doc = Document(template_file)
                
                deg_full = "ประกาศนียบัตรวิชาชีพชั้นสูง (ปวส.)" if degree == "ปวส." else "ประกาศนียบัตรวิชาชีพ (ปวช.)"
                
                replacements = {
                    "{{ degree_title }}": deg_full,
                    "{{ course_code }}": course_code,
                    "{{ course_name }}": course_name,
                    "{{ hours_per_week }}": str(hours_per_week),
                    "{{ total_weeks }}": str(weeks_input),
                    "{{ semester }}": semester,
                    "{{ cb_y1 }}": "☑" if year == 1 else "☐",
                    "{{ cb_y2 }}": "☑" if year == 2 else "☐",
                    "{{ cb_y3 }}": "☑" if year == 3 else "☐"
                }
                replace_placeholders_in_doc(doc, replacements)
                
                target_table = None
                for tbl in doc.tables:
                    for row in tbl.rows:
                        row_text = " ".join([c.text for c in row.cells])
                        if "ส.ป." in row_text or "Teaching Point" in row_text:
                            target_table = tbl
                            break
                    if target_table:
                        break
                        
                if not target_table:
                    target_table = doc.tables[0]
                
                for item in plans:
                    row_cells = target_table.add_row().cells
                    row_cells[0].text = str(item.get("week", ""))
                    row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                    row_cells[1].text = item.get("topic_full", "")
                    
                    tp_list = item.get("teaching_points", [])
                    row_cells[2].text = "\n".join(tp_list) if isinstance(tp_list, list) else str(tp_list)
                    
                    act_list = item.get("activities", [])
                    row_cells[3].text = "\n".join([f"- {a}" for a in act_list]) if isinstance(act_list, list) else str(act_list)
                    
                    med_list = item.get("media", [])
                    row_cells[4].text = "\n".join([f"- {m}" for m in med_list]) if isinstance(med_list, list) else str(med_list)
                    
                    eval_list = item.get("assessment", [])
                    row_cells[5].text = "\n".join([f"- {e}" for e in eval_list]) if isinstance(eval_list, list) else str(eval_list)
                    
                    for cell in row_cells:
                        set_cell_font(cell, font_name="TH SarabunPSK", font_size=Pt(14))
                
                out_stream = io.BytesIO()
                doc.save(out_stream)
                out_stream.seek(0)
                
                status.update(label="✅ ดำเนินการสร้างโครงการสอนสำเร็จ!", state="complete", expanded=False)
                st.balloons()
                st.success("🎉 ระบบสร้างเอกสารโครงการสอนเสร็จสมบูรณ์เรียบร้อยแล้ว")
                
                st.download_button(
                    label=f"📥 ดาวน์โหลดโครงการสอน_{course_code}.docx",
                    data=out_stream,
                    file_name=f"โครงการสอน_{course_code}_{degree}_{year}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
                
            except Exception as err:
                status.update(label="❌ เกิดข้อผิดพลาด", state="error")
                st.error(f"รายละเอียดข้อผิดพลาด: {str(err)}")
