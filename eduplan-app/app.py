def extract_course_plan(api_key: str, content_data, total_weeks: int, course_name: str):
    client = genai.Client(api_key=api_key)
    prompt = f"""
คุณคือผู้เชี่ยวชาญการจัดทำหลักสูตรและโครงการสอนระดับอาชีวศึกษา (มาตรฐาน สอศ.)
จงอ่านข้อมูลจากตารางวิเคราะห์งาน/หลักสูตรวิชา '{course_name}' แล้วสังเคราะห์จัดทำเป็น 'ตารางโครงการสอนรายสัปดาห์' ให้ครบถ้วนจำนวน {total_weeks} สัปดาห์ (สัปดาห์ที่ 1 ถึง {total_weeks})

*** ข้อกำหนดข้อมูลในแต่ละสัปดาห์ (เน้นความเชื่อมโยง 100%): ***
1. week: ตัวเลขสัปดาห์ (1 ถึง {total_weeks})
2. topic: ชื่อหน่วยและเรื่อง (บรรทัดแรก: หน่วยที่... ชื่อหน่วย / บรรทัดสอง: เรื่อง...)
3. tp: จุดประสงค์เชิงพฤติกรรม 2-3 ข้อ (สังเคราะห์จากสมรรถนะย่อย ความรู้ และทักษะ)
4. act: ขั้นตอน Active Learning 4 ขั้นตอน (1.ขั้นนำเข้าสู่บทเรียน 2.ขั้นให้ความรู้/ศึกษาค้นคว้า 3.ขั้นฝึกปฏิบัติการ 4.ขั้นสรุปและประเมินผล)
5. media: สื่อการเรียนรู้ที่ **สอดคล้องกับสิ่งที่นักเรียนใช้ในกิจกรรมปฏิบัติจริง** (เช่น แบบฟอร์มใบสั่งซื้อ PO, ซอฟต์แวร์ ERP จำลอง, แผนผัง Layout คลังสินค้า พร้อมสไลด์และใบงาน)
6. assess: การวัดผลที่ **สอดคล้องกับกิจกรรมจริง** (เช่น แบบประเมินทักษะ Rubric, แบบทดสอบย่อย, ตรวจผลงานใบงาน)

ตอบกลับเป็น Pure JSON Array เท่านั้น ห้ามใส่ markdown code block ครอบ:
[
  {{
    "week": 1,
    "topic": "หน่วยที่ 1 ความรู้เบื้องต้นเกี่ยวกับซัพพลายเชน\\nเรื่อง ความหมายและขอบเขต",
    "tp": "- อธิบายความหมายและขอบเขตได้\\n- เปรียบเทียบความแตกต่างได้",
    "act": "1. ขั้นนำ: เปิดคลิปวิดีโอ...\\n2. ขั้นสอน: บรรยายสไลด์...\\n3. ขั้นปฏิบัติ: ทำใบงานวิเคราะห์...\\n4. ขั้นสรุป: สรุปร่วมกัน",
    "media": "- สไลด์มัลติมีเดีย หน่วยที่ 1\\n- คลิปวิดีโอวงจรชีวิตผลิตภัณฑ์\\n- ใบงานที่ 1.1 เรื่อง โครงสร้างซัพพลายเชน",
    "assess": "- แบบประเมินใบงานที่ 1.1\\n- แบบทดสอบย่อยท้ายคาบ\\n- แบบสังเกตพฤติกรรมการทำงาน"
  }}
]
"""
    if isinstance(content_data, str):
        full_contents = f"ข้อมูลตารางวิเคราะห์งาน:\n{content_data[:12000]}\n\n{prompt}"
    else:
        full_contents = [content_data, prompt]

    # ลิสต์โมเดลที่ใช้สลับหนี Error 503
    candidate_models = ["gemini-2.5-flash", "gemini-3.6-flash", "gemini-2.5-pro"]
    last_error = None

    for model_name in candidate_models:
        for attempt in range(2):
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=full_contents,
                    config=types.GenerateContentConfig(response_mime_type="application/json")
                )
                return json.loads(response.text)
            except Exception as e:
                last_error = e
                err_msg = str(e)
                # หากเจอ 503 หรือ UNAVAILABLE ให้หน่วงเวลาสั้นๆ แล้วลองอีกรอบ หรือสลับรุ่นโมเดลทันที
                if "503" in err_msg or "UNAVAILABLE" in err_msg:
                    time.sleep(3)
                    continue
                else:
                    break

    raise last_error
