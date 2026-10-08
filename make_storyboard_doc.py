import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)

def create_storyboard_script():
    doc = Document()
    
    # Margins
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # Styles
    normal_style = doc.styles['Normal']
    normal_font = normal_style.font
    normal_font.name = 'TH Sarabun PSK'
    normal_font.size = Pt(16)
    normal_font.color.rgb = RGBColor(0x22, 0x22, 0x22)

    def add_custom_heading(text, level):
        h = doc.add_heading(text, level=level)
        run = h.runs[0]
        run.font.name = 'TH Sarabun PSK'
        if level == 1:
            run.font.size = Pt(22)
            run.bold = True
            run.font.color.rgb = RGBColor(0xD3, 0x54, 0x00) # Vibrant Orange
        elif level == 2:
            run.font.size = Pt(18)
            run.bold = True
            run.font.color.rgb = RGBColor(0x2C, 0x3E, 0x50)
        return h

    # ------------------ COVER ------------------
    p_cov = doc.add_paragraph()
    p_cov.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cov.paragraph_format.space_before = Pt(50)
    
    r_sub = p_cov.add_run("คู่มือฉบับเล่าเรื่อง & บทสตอรี่บอร์ดอนิเมชัน (Storytelling & Animation Script)\n")
    r_sub.font.size = Pt(18)
    r_sub.bold = True
    r_sub.font.color.rgb = RGBColor(0x7F, 0x8C, 0x8D)

    r_main = p_cov.add_run("🎬 ตะลุยโลก Innovation Lab: ศน. โคราช\n\"เมื่อ Discord กลายเป็นห้องทดลองลับ ยกระดับการศึกษาลูกหลานเมืองย่าโม\"")
    r_main.font.size = Pt(24)
    r_main.bold = True
    r_main.font.color.rgb = RGBColor(0xD3, 0x54, 0x00)

    p_meta = doc.add_paragraph()
    p_meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_meta.paragraph_format.space_before = Pt(30)
    r_m = p_meta.add_run("ความยาวโดยประมาณ: 2.30 - 3.00 นาที | กลุ่มเป้าหมาย: ศึกษานิเทศก์ และครูแกนนำทุกสังกัดในโคราช\nโทนเรื่อง: อบอุ่น สนุกสนาน ทลายความกลัวเทคโนโลยี สร้างแรงบันดาลใจ")
    r_m.font.size = Pt(16)
    r_m.font.italic = True

    doc.add_page_break()

    # ------------------ CHARACTERS ------------------
    add_custom_heading("🎭 ตัวละครหลักในแอนิเมชัน (Main Characters)", level=1)
    
    char_table = doc.add_table(rows=1, cols=3)
    char_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    ch_h = char_table.rows[0].cells
    ch_h[0].text = "ตัวละคร"
    ch_h[1].text = "รูปลักษณ์และบุคลิก"
    ch_h[2].text = "หน้าที่ในเรื่อง"
    for c in ch_h:
        set_cell_background(c, 'D35400')
        for cp in c.paragraphs:
            for cr in cp.runs:
                cr.font.name = 'TH Sarabun PSK'
                cr.font.size = Pt(16)
                cr.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                cr.bold = True

    chars_data = [
        ("ศน. น้องฟ้า\n(ตัวแทน ศน. รุ่นใหม่)", "หญิงสาวสดใส สวมเสื้อสูทคลุมทับเสื้อผ้าไหมโคราช ถือแท็บเล็ตและแฟ้มงานพะรุงพะรัง แววตามุ่งมั่นแต่แอบเหนื่อยล้า", "เป็นกระจกสะท้อน 'ความเจ็บปวด (Pain Point)' ของ ศน. ยุคปัจจุบัน ที่งานเอกสารล้นมือและขาดที่ปรึกษา"),
        ("ศน. พี่เก่ง\n(ตัวแทน ศน. รุ่นเก๋า)", "ชายวัยกลางคน อารมณ์ดี ใจดี ประสบการณ์สูง แต่เริ่มตามเครื่องมือเทคโนโลยีไม่ทัน รู้สึกว่าแอปใหม่ๆ เล่นยาก", "เป็นตัวแทนของความกลัวเทคโนโลยี (Tech Hesitation) แต่เมื่อลองใช้แล้วพบว่าง่ายกว่าที่คิด"),
        ("ปราชญ์โคราช AI\n(มาสคอตนำทาง)", "หุ่นยนต์นกฮูกตัวกลมสุดน่ารัก ตาโตเป็นหน้าจอดิจิทัล มีผ้าพันคอไหมโคราชสีส้ม บินได้ มีประกายดาวรอบตัว", "เป็นไกด์พาชมเซิร์ฟเวอร์ ให้กำลังใจ ตอบคำถามฉับไว และเชื่อมโยงทุกคนเข้าหากัน")
    ]
    for c1, c2, c3 in chars_data:
        r = char_table.add_row().cells
        r[0].text = c1
        r[1].text = c2
        r[2].text = c3
        for cell in r:
            for cp in cell.paragraphs:
                for cr in cp.runs:
                    cr.font.name = 'TH Sarabun PSK'
                    cr.font.size = Pt(15)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # ------------------ SCRIPT TABLE ------------------
    add_custom_heading("🎞️ บทภาพและบทพากย์ละเอียด (Detailed Storyboard Script)", level=1)

    sb_table = doc.add_table(rows=1, cols=4)
    sb_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sb_h = sb_table.rows[0].cells
    sb_h[0].text = "ฉาก / เวลา"
    sb_h[1].text = "ภาพในแอนิเมชัน (Visual Scene)"
    sb_h[2].text = "บทพากย์ / เสียงบทสนทนา (Voiceover)"
    sb_h[3].text = "AI Visual Prompt (สำหรับเจนภาพ)"
    for c in sb_h:
        set_cell_background(c, '2C3E50')
        for cp in c.paragraphs:
            for cr in cp.runs:
                cr.font.name = 'TH Sarabun PSK'
                cr.font.size = Pt(16)
                cr.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                cr.bold = True

    story_scenes = [
        (
            "ฉากที่ 1\n(0:00 - 0:30)\nเปิดเรื่อง & ปัญหา",
            "ภาพห้องทำงาน ศน. น้องฟ้านั่งถอนหายใจ มีกองเอกสาร ว.PA แผนนิเทศท่วมโต๊ะ หน้าจอคอมเปิดไฟล์ไลน์ที่รูปโหลดไม่ขึ้นขึ้นกากบาทสีเทา พี่เก่งเดินเข้ามาพร้อมกาแฟ บ่นว่าอยากได้ไอเดียนิเทศ Active Learning ไปช่วยครูพรุ่งนี้",
            "เสียงพากย์นุ่มนวล:\n'เคยรู้สึกไหมครับ... งานนิเทศก็ต้องลง พื้นที่ก็ต้องไป ว.PA ก็ต้องทำ... จะหาไอเดียเจ๋งๆ หรือถามเกณฑ์ระเบียบที ก็เหมือนอยู่ตัวคนเดียวในสังกัด!'\nศน.น้องฟ้า: 'พี่เก่งคะ สื่อในไลน์หมดอายุอีกแล้วค่ะ!'",
            "Cute 2D Pixar style animation, overwhelmed Thai female educator at messy office desk with paper stacks, digital tablet, cute expressions, warm studio lighting."
        ),
        (
            "ฉากที่ 2\n(0:30 - 1:00)\nการปรากฏตัว",
            "หน้าจอสว่างวาบ หุ่นยนต์นกฮูก 'ปราชญ์โคราช AI' บินหมุนตัวออกมาจากจอแท็บเล็ต พร้อมโปรเจกต์โฮโลแกรมรูปประตูมิติดิจิทัล มีป้ายไฟเขียนว่า '🧪 Innovation Lab: ศน. โคราช'",
            "ปราชญ์โคราช AI (เสียงสดใสร่าเริง):\n'ฮัลโหลพี่น้อง ศน. โคราชบ้านเอ๋ง! ไม่ต้องปวดหัวอีกต่อไปแล้วครับ... ยินดีต้อนรับสู่ Innovation Lab พื้นที่วิชาการสุดล้ำบน Discord ที่ไม่ได้มีไว้เล่นเกม แต่มีไว้เสกไอเดียการศึกษา!'",
            "A cute friendly owl robot mascot wearing Thai Korat silk scarf flying out of a glowing futuristic tablet, magical educational laboratory background, 3D Pixar render."
        ),
        (
            "ฉากที่ 3\n(1:00 - 1:45)\nพาทัวร์แล็บ",
            "นกฮูกพาบินข้ามห้องต่างๆ ใน Discord:\n1. แวะ #สภากาแฟ เห็น ศน. ต่างสังกัดจิบกาแฟส่งสติ๊กเกอร์\n2. แวะ #คลังสื่อและนวัตกรรม มีการ์ดสื่อแยกสี Tag ค้นหาง่ายเพียงคลิกเดียว\n3. แวะ #โชว์ผลงาน ภาพครูและนักเรียนยิ้มแย้มจากการนิเทศ",
            "เสียงพากย์:\n'ที่นี่เราจัดระเบียบใหม่หมดจด! อยากคุยเล่นคลายเครียดแวะ #สภากาแฟ... อยากได้สื่อ แผนการสอน ไม่ต้องกลัวไฟล์หมดอายุ เพราะอยู่ใน #คลังสื่อและนวัตกรรม ค้นหาด้วย Tag ง่ายในคลิกเดียว!'",
            "Isometric cutaway view of a colorful digital house with coffee zone, library shelves with glowing tag badges, gallery wall of school success, cheerful aesthetic."
        ),
        (
            "ฉากที่ 4\n(1:45 - 2:20)\nคู่หู AI แท็กทีม",
            "น้องฟ้าพิมพ์ลงในห้อง AI: '@Korat Edu-Bot ขอไอเดียสะเต็ม ป.4'\nปราชญ์โคราชตาเปล่งแสง เสกแผนกิจกรรมออกมา 3 สเต็ปทันที!\nพี่เก่งถาม: 'แล้วเกณฑ์ ว.PA ละ?'\nนกฮูกชี้ไปที่ลิงก์ NotebookLM: 'กดตรงนี้ ตอบตรงตามระเบียบเป๊ะ 100% ครับ!'",
            "ศน.น้องฟ้า: 'ว้าว! ปราชญ์โคราชคิดไอเดียไวใน 3 วินาที!'\nปราชญ์โคราช: 'ส่วนเรื่องระเบียบราชการที่ต้องเป๊ะ 100% เรามีฐานข้อมูล NotebookLM ปักหมุดไว้ ค้นหาอ้างอิงตรงเป๊ะ ไม่มีหลอนครับ!'",
            "Two AI helper concept: one energetic brainstorming robot holding glowing lightbulb, one scholarly robot reading golden law handbook, smiling educators watching."
        ),
        (
            "ฉากที่ 5\n(2:20 - 2:50)\nภารกิจแรก & รวมพลัง",
            "ภาพหน้าจอชวนทำเควสต์แรก: ศน. สพป. ได้ป้ายสีฟ้า, สพม. สีเขียว, สช. สีเหลือง, อปท. สีส้ม ทุกคนเอามือมาจับรวมกันตรงกลาง กลายเป็นต้นไม้แห่งการเรียนรู้ที่สว่างไสว",
            "เสียงพากย์ปิดท้ายสุดอบอุ่น:\n'ภารกิจแรกของคุณ... แค่คลิกลิงก์เข้ามา แล้วพิมพ์ทักทายใน #สภากาแฟ เพื่อรับยศสีประจำสังกัด... ต่างบทบาท ต่างหน้าที่ แต่มีหัวใจเดียวกัน เพื่อเด็กโคราชบ้านเราครับ!'\nขึ้นหน้าจอ QR Code & Link เข้าร่วม",
            "Diverse group of Thai education supervisors wearing colorful ID badges (blue, green, yellow, orange) holding hands in circle around glowing futuristic tree of knowledge."
        )
    ]

    for sc_id, sc_vis, sc_voice, sc_prompt in story_scenes:
        r = sb_table.add_row().cells
        r[0].text = sc_id
        r[1].text = sc_vis
        r[2].text = sc_voice
        r[3].text = sc_prompt
        for cell in r:
            for cp in cell.paragraphs:
                for cr in cp.runs:
                    cr.font.name = 'TH Sarabun PSK'
                    cr.font.size = Pt(14)

    doc.add_page_break()

    # ------------------ AI TOOLS GUIDE ------------------
    add_custom_heading("🛠️ คู่มือเครื่องมือ AI ที่แนะนำในการผลิตแอนิเมชันจริง", level=1)
    
    p_tool = doc.add_paragraph()
    p_tool.add_run("คุณแสนยากรสามารถนำสคริปต์นี้ไปแปลงเป็นคลิปจริงได้อย่างรวดเร็ว โดยใช้เครื่องมือ AI ยอดนิยมดังนี้:\n\n"
                   "1. สร้างภาพนิ่งตัวละครและฉาก (Image Generation):\n"
                   "   • Midjourney / Flux / Leonardo.ai หรือ Antigravity AI\n"
                   "   • ใช้ Prompt ภาษาอังกฤษจากตารางคอลัมน์ที่ 4 ได้ทันที\n\n"
                   "2. ชุบชีวิตภาพนิ่งให้เคลื่อนไหวเป็นวิดีโอ (Image-to-Video):\n"
                   "   • Kling AI / Runway Gen-3 / Luma Dream Machine / Hailuo AI\n"
                   "   • นำภาพที่เจนได้มาอัปโหลด แล้วใส่การเคลื่อนไหว เช่น 'character smiles and waves hand, gentle camera zoom'\n\n"
                   "3. เสียงพากย์ภาษาไทยเป็นธรรมชาติ (AI Voiceover):\n"
                   "   • ElevenLabs (โมเดล Multilingual v2 ให้เสียงไทยที่น่าฟังและแสดงอารมณ์ได้ดีเยี่ยม)\n"
                   "   • Botnoi Voice (เสียงพากย์สไตล์ไทยแท้ เลือกบุคลิกข้าราชการ/อาจารย์ได้)\n\n"
                   "4. ตัดต่อและใส่ซับไตเติลอัตโนมัติ (Editing & Auto Subtitles):\n"
                   "   • CapCut (PC หรือมือถือ) มีระบบ Auto Captions ภาษาไทยที่แม่นยำ พร้อมใส่เสียงดนตรีประกอบแนวสดใส")

    out_dir = r"I:\My Drive\@Discord\Innovation_Lab_Manual"
    out_file = os.path.join(out_dir, "Animation_Story_Script.docx")
    doc.save(out_file)
    print(f"Storyboard Script created successfully at: {out_file}")

if __name__ == '__main__':
    create_storyboard_script()
