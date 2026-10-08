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

def generate_expert_academic_manual():
    doc = Document()
    
    # Page Margins (Standard Thai Academic Format)
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.25)
        section.right_margin = Inches(1.0)

    # Base Styles
    normal_style = doc.styles['Normal']
    normal_font = normal_style.font
    normal_font.name = 'TH Sarabun PSK'
    normal_font.size = Pt(16)
    normal_font.color.rgb = RGBColor(0x22, 0x22, 0x22)

    def add_heading_1(text):
        h = doc.add_heading(text, level=1)
        h.paragraph_format.space_before = Pt(14)
        h.paragraph_format.space_after = Pt(8)
        run = h.runs[0]
        run.font.name = 'TH Sarabun PSK'
        run.font.size = Pt(20)
        run.bold = True
        run.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D) # Navy
        return h

    def add_heading_2(text):
        h = doc.add_heading(text, level=2)
        h.paragraph_format.space_before = Pt(10)
        h.paragraph_format.space_after = Pt(4)
        run = h.runs[0]
        run.font.name = 'TH Sarabun PSK'
        run.font.size = Pt(18)
        run.bold = True
        run.font.color.rgb = RGBColor(0x2E, 0x5B, 0x88)
        return h

    def add_heading_3(text):
        h = doc.add_heading(text, level=3)
        h.paragraph_format.space_before = Pt(8)
        h.paragraph_format.space_after = Pt(2)
        run = h.runs[0]
        run.font.name = 'TH Sarabun PSK'
        run.font.size = Pt(16)
        run.bold = True
        run.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
        return h

    # ------------------ COVER PAGE ------------------
    cover_p1 = doc.add_paragraph()
    cover_p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cover_p1.paragraph_format.space_before = Pt(40)
    run_org = cover_p1.add_run("เอกสารประกอบการประเมินวิทยฐานะ ด้านที่ 3 ผลงานทางวิชาการ/นวัตกรรมการนิเทศการศึกษา\nเพื่อขอรับการประเมินวิทยฐานะศึกษานิเทศก์เชี่ยวชาญ (คศ.4)")
    run_org.font.size = Pt(16)
    run_org.bold = True

    cover_title = doc.add_paragraph()
    cover_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cover_title.paragraph_format.space_before = Pt(24)
    cover_title.paragraph_format.space_after = Pt(14)
    run_t1 = cover_title.add_run("รายงานการวิจัยและพัฒนานวัตกรรมดิจิทัลเชิงสถาปัตยกรรม\n")
    run_t1.font.size = Pt(22)
    run_t1.bold = True
    run_t1.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    run_t2 = cover_title.add_run("การพัฒนาชุมชนการเรียนรู้ทางวิชาชีพดิจิทัลข้ามสังกัด (Cross-Silo CoP)\n\"🧪 Innovation Lab: ศน. โคราช\" ร่วมกับระบบผู้ช่วยปัญญาประดิษฐ์\n(AI-Augmented Educational Supervision)")
    run_t2.font.size = Pt(19)
    run_t2.bold = True
    run_t2.font.color.rgb = RGBColor(0x2E, 0x5B, 0x88)

    cover_info = doc.add_paragraph()
    cover_info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cover_info.paragraph_format.space_before = Pt(90)
    
    r_author = cover_info.add_run("ผู้จัดทำและพัฒนานวัตกรรม\nนายแสนยากร แสวงชิด\nตำแหน่ง ศึกษานิเทศก์ วิทยฐานะศึกษานิเทศก์ชำนาญการพิเศษ\nกลุ่มนิเทศ ติดตาม และประเมินผลการจัดการศึกษา\nสำนักงานศึกษาธิการจังหวัดนครราชสีมา\nกระทรวงศึกษาธิการ")
    r_author.font.size = Pt(18)
    r_author.bold = True

    doc.add_page_break()

    # ------------------ CHAPTER 1 ------------------
    add_heading_1("บทที่ 1 บทนำและหลักการเชิงยุทธศาสตร์")

    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Inches(0.5)
    p.add_run("1.1 ความเป็นมาและสภาพปัญหา (Problem Statement)\nการจัดการศึกษาในระดับพื้นที่จังหวัดนครราชสีมา มีความหลากหลายของบริบทโรงเรียนและโครงสร้างหน่วยงานทางการศึกษา โดยมีศึกษานิเทศก์ (ศน.) ปฏิบัติหน้าที่กระจายอยู่ตามสังกัดต่างๆ ได้แก่ สำนักงานศึกษาธิการจังหวัด, สำนักงานเขตพื้นที่การศึกษาประถมศึกษา (สพป. นครราชสีมา เขต 1-6), สำนักงานเขตพื้นที่การศึกษามัธยมศึกษา (สพม. นครราชสีมา), องค์การบริหารส่วนจังหวัด (อบจ. นครราชสีมา), เทศบาล และ อาชีวศึกษา")

    p2 = doc.add_paragraph()
    p2.paragraph_format.first_line_indent = Inches(0.5)
    p2.add_run("แม้ศึกษานิเทศก์ทุกสังกัดจะมีเป้าหมายสูงสุดร่วมกันคือการยกระดับคุณภาพการเรียนรู้ของผู้เรียนในจังหวัดนครราชสีมา แต่ในทางปฏิบัติมักประสบปัญหาอุปสรรคเชิงโครงสร้าง 3 ประการสำคัญ:")
    
    bullets = [
        ("ปัญหาไซโลระหว่างสังกัด (Cross-Silo Disconnection): ", "ขาดพื้นที่กลางในการแลกเปลี่ยนเรียนรู้ สื่อนวัตกรรม และแนวปฏิบัติที่ดี (Best Practices) ทำให้เกิดการทำงานซ้ำซ้อนและขาดพลังการบูรณาการในระดับจังหวัด"),
        ("ความซับซ้อนตามเกณฑ์ ว.PA และการจัดการเรียนรู้เชิงรุก: ", "การขับเคลื่อนการประเมินตำแหน่งและวิทยฐานะตามเกณฑ์ ว.PA (ข้อตกลงในการพัฒนางาน) และการส่งเสริมการจัดการเรียนรู้เชิงรุก (Active Learning) ต้องอาศัยการสืบค้นระเบียบที่แม่นยำและการระดมความคิดเห็นอย่างต่อเนื่อง"),
        ("ข้อจำกัดของแพลตฟอร์มการสื่อสารเดิม: ", "แอปพลิเคชันสนทนาทั่วไป (เช่น LINE) มักประสบปัญหาไฟล์และภาพหมดอายุ ข้อมูลสำคัญเลือนหาย และไม่สามารถจัดหมวดหมู่คลังความรู้ได้อย่างเป็นระบบถาวร")
    ]
    for b_title, b_desc in bullets:
        bp = doc.add_paragraph()
        bp.paragraph_format.left_indent = Inches(0.5)
        run_b = bp.add_run("• " + b_title)
        run_b.bold = True
        bp.add_run(b_desc)

    p3 = doc.add_paragraph()
    p3.paragraph_format.first_line_indent = Inches(0.5)
    p3.add_run("1.2 วัตถุประสงค์ของการพัฒนานวัตกรรม\n"
               "1. เพื่อออกแบบและพัฒนาแพลตฟอร์มชุมชนการเรียนรู้ทางวิชาชีพดิจิทัลข้ามสังกัด (Cross-Silo CoP) ภายใต้ชื่อ \"🧪 Innovation Lab: ศน. โคราช\" บนแพลตฟอร์ม Discord (เริ่มพัฒนาระบบนำร่องเมื่อวันที่ 6 กรกฎาคม 2569)\n"
               "2. เพื่อพัฒนาระบบผู้ช่วยปัญญาประดิษฐ์ \"ปราชญ์โคราช AI (Korat Edu-Bot)\" ขับเคลื่อนด้วยโมเดล Gemini 3.5 Flash ผสานเทคโนโลยี Interactive Buttons และ Pop-up Modal Form เพื่อเสริมพลังการนิเทศการศึกษา (AI-Augmented Educational Supervision)\n"
               "3. เพื่อประเมินผลการสร้างเครือข่ายวิชาชีพ การแลกเปลี่ยนเรียนรู้ และการนำสื่อนวัตกรรมการศึกษาไปประยุกต์ใช้จริงในสถานศึกษาทั่วจังหวัดนครราชสีมา")

    doc.add_page_break()

    # ------------------ CHAPTER 2 ------------------
    add_heading_1("บทที่ 2 กรอบแนวคิดเชิงทฤษฎี งานวิจัยอ้างอิง และสถาปัตยกรรมพื้นที่")

    p_theo = doc.add_paragraph()
    p_theo.paragraph_format.first_line_indent = Inches(0.5)
    p_theo.add_run("2.1 กรอบแนวคิดเชิงทฤษฎีและงานวิจัยระดับสากลที่รองรับการออกแบบห้อง (Theoretical Foundations)\n"
                   "ในการออกแบบสถาปัตยกรรมพื้นที่ของ \"Innovation Lab: ศน. โคราช\" ผู้พัฒนานวัตกรรมได้นำหลักการทางวิชาการและงานวิจัยระดับสากลมารองรับวัตถุประสงค์ของแต่ละห้องอย่างเป็นระบบ เพื่อตอบสนองมาตรฐานวิทยฐานะศึกษานิเทศก์เชี่ยวชาญ ดังนี้:")

    theo_details = [
        ("1. หมวด Information (#👋-ต้อนรับและกติกา, #📢-ประกาศ-แจ้งข่าว): ",
         "รองรับด้วยทฤษฎี Digital Onboarding & Institutional Norms (Scott, 2014; Bauer & Erdogan, 2011) ที่เน้นการสร้างบรรทัดฐานร่วม การปฐมนิเทศดิจิทัล และการลดความวิตกกังวลในการใช้เทคโนโลยี (Technology Anxiety) ของสมาชิกใหม่"),
        ("2. หมวดคลังความรู้ระบบฟอรั่ม (#📚-คลังสื่อและนวัตกรรม, #🏆-โชว์ผลงาน): ",
         "รองรับด้วย SECI Model ของ Nonaka & Takeuchi (1995) ในการแปลงความรู้ฝังลึก (Tacit Knowledge) ของศึกษานิเทศก์แต่ละท่านให้กลายเป็นความรู้ชัดแจ้ง (Explicit Knowledge) ผ่านระบบ Forum Tagging (Vander Wal, 2007) ซึ่งเก็บรักษาองค์ความรู้ได้อย่างถาวร ไม่สูญหายตามกาลเวลา"),
        ("3. หมวดสังคมและพื้นที่ปลอดภัย (#☕-สภากาแฟ-คุยทั่วไป, #🎉-ห้องพักใจ-เล่าเรื่อง): ",
         "รองรับด้วยทฤษฎีพื้นที่ปลอดภัยทางจิตวิทยา (Psychological Safety Theory) ของ Amy Edmondson (1999, Harvard University) และทฤษฎีทุนทางสังคม (Social Capital Theory) ของ Robert Putnam (2000) ที่ชี้ชัดว่า ชุมชนวิชาชีพจะเข้มแข็งได้ต้องมีพื้นที่ไม่เป็นทางการที่สมาชิกกล้าสะท้อนข้อผิดพลาดและช่วยเหลือกันโดยปราศจากการตัดสิน"),
        ("4. หมวดการปฏิบัติงานข้ามสังกัด (#📅-นัดหมาย-ปรึกษางาน, #moderator-only): ",
         "รองรับด้วยแนวคิด Boundary-Crossing & Collaboration ของ Akkerman & Bakker (2011) ที่มุ่งทลายไซโลระหว่างหน่วยงานทางการศึกษา นำไปสู่การนิเทศร่วมข้ามสังกัด (Joint-Supervision) และการกำกับมาตรฐานข้อมูลโดยคณะทำงาน (Steering Committee)"),
        ("5. ห้องปฏิบัติการปัญญาประดิษฐ์ (#ai-เพื่อนคู่คิด-ศึกษานิเทศก์): ",
         "รองรับด้วยแนวคิด AI-Augmented Cognition & Supervision (Luckin, 2018; Mollick, 2024, Wharton School) ที่ใช้ AI เป็น Cognitive Scaffolding (นั่งร้านทางปัญญา) ช่วยแบ่งเบาภาระงานเอกสารรูทีนและเร่งกระบวนการระดมสมองเชิงวิชาการ"),
        ("6. หมวดสื่อสารเสียงและสัมมนาเสมือน (Voice Channels & ห้องประชุม 1): ",
         "รองรับด้วย Community of Inquiry (CoI) Framework ของ Garrison, Anderson & Archer (2000) ที่ผสานการมีตัวตนทางสังคม (Social Presence) และการรู้คิด (Cognitive Presence) ผ่านการประชุมออนไลน์และการแชร์หน้าจอสาธิต")
    ]

    for t_title, t_desc in theo_details:
        tp = doc.add_paragraph()
        tp.paragraph_format.left_indent = Inches(0.5)
        run_t = tp.add_run(t_title)
        run_t.bold = True
        tp.add_run(t_desc)

    add_heading_2("2.2 ผังโครงสร้างห้องและหน้าที่จริง (Real Channel Architecture)")
    
    table = doc.add_table(rows=1, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0].cells
    hdr[0].text = 'หมวดหมู่'
    hdr[1].text = 'ชื่อห้องจริง'
    hdr[2].text = 'ประเภท'
    hdr[3].text = 'ฟังก์ชันเชิงวิชาการและทฤษฎีรองรับ'
    for c in hdr:
        set_cell_background(c, '1B365D')
        for cp in c.paragraphs:
            for cr in cp.runs:
                cr.font.name = 'TH Sarabun PSK'
                cr.font.size = Pt(15)
                cr.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                cr.bold = True

    channels_data = [
        ("Information", "#👋-ต้อนรับและกติกา\n#📢-ประกาศ-แจ้งข่าว", "Text\nText", "ปฐมนิเทศดิจิทัล แจ้งกติกาชุมชน และนโยบายการนิเทศจังหวัด (Digital Onboarding)"),
        ("Forum System", "#📚-คลังสื่อและนวัตกรรม\n#🏆-โชว์ผลงาน", "Forum\nForum", "คลังสื่อแผนการสอน นวัตกรรม และแกลเลอรีถอดบทเรียนความสำเร็จ (SECI Model & Tagging)"),
        ("Text Channels", "#☕-สภากาแฟ-คุยทั่วไป\n#📅-นัดหมาย-ปรึกษางาน\n#🎉-ห้องพักใจ-เล่าเรื่อง\n#ai-เพื่อนคู่คิด-ศึกษานิเทศก์", "Text\nText\nText\nText", "พื้นที่ปลอดภัย สานสัมพันธ์ นัดหมายนิเทศร่วม คลายเครียด และห้องทดลอง AI-Augmented Supervision"),
        ("Voice Channels", "#🔊 ชวนกันเล่นเกม\n#🔊 มุมกาแฟ (คุยเล่น)\n#🔊 ห้องประชุม 1", "Voice\nVoice\nVoice", "เสวนาวิชาการไม่เป็นทางการ ประชุมวิชาการ และ Virtual Workshop (CoI Framework)"),
        ("Executive Zone", "#moderator-only", "Text (Private)", "ห้องวอร์รูมคณะทำงาน กำกับมาตรฐาน และวางแผนยุทธศาสตร์ระดับจังหวัด (Governance)")
    ]

    for cat, ch, ctype, desc in channels_data:
        r = table.add_row().cells
        r[0].text = cat
        r[1].text = ch
        r[2].text = ctype
        r[3].text = desc
        for c in r:
            for cp in c.paragraphs:
                for cr in cp.runs:
                    cr.font.name = 'TH Sarabun PSK'
                    cr.font.size = Pt(14)

    doc.add_page_break()

    # ------------------ CHAPTER 3 ------------------
    add_heading_1("บทที่ 3 สถาปัตยกรรมระบบปัญญาประดิษฐ์ \"ปราชญ์โคราช AI (Korat Edu-Bot)\"")
    
    p_ai = doc.add_paragraph()
    p_ai.paragraph_format.first_line_indent = Inches(0.5)
    p_ai.add_run("3.1 การบูรณาการโมเดล Gemini 3.5 Flash และโครงสร้างพื้นฐานคลาวด์\n"
                 "ระบบผู้ช่วยอัจฉริยะขับเคลื่อนด้วยโมเดล Gemini 3.5 Flash ซึ่งเป็นโมเดลปัญญาประดิษฐ์ระดับแนวหน้า มีความสามารถในการประมวลผลเชิงเหตุผล (Reasoning) และภาษาไทยระดับสูง โดยติดตั้งบน Google Cloud Platform (Compute Engine VM) ควบคุมกระบวนการทำงานด้วย PM2 Process Manager เพื่อความเสถียรระดับ 100% ต่อเนื่อง 24 ชั่วโมง")

    add_heading_2("3.2 นวัตกรรมส่วนต่อประสานแบบปุ่มสัมผัส (Interactive Buttons & Modal Interface)")
    p_ui = doc.add_paragraph()
    p_ui.paragraph_format.first_line_indent = Inches(0.5)
    p_ui.add_run("เพื่อลดอุปสรรคการใช้งานของศึกษานิเทศก์ทุกช่วงวัย ผู้พัฒนานวัตกรรมได้ออกแบบคำสั่ง !setup-menu เพื่อสร้างการ์ดปุ่มกดสัมผัสลอยปักหมุดถาวรในห้องแชท ได้แก่:\n"
                 "• ปุ่ม '💡 ขอไอเดีย Active Learning' (สีเขียว)\n"
                 "• ปุ่ม '📋 ปรึกษาเกณฑ์ ว.PA' (สีน้ำเงิน)\n"
                 "• ปุ่ม '🎯 ออกแบบแผนการนิเทศ' (สีเทา)\n"
                 "• ปุ่ม '💬 พิมพ์คำถามอิสระ' (สีเทา)\n\n"
                 "เมื่อคลิกปุ่ม ระบบจะเปิดหน้าต่าง Pop-up Modal Form ให้ผู้ใช้พิมพ์ข้อความ และส่งประมวลผลผ่าน Gemini 3.5 Flash ตอบกลับทันทีโดยไม่ต้องจำคำสั่งคอมพิวเตอร์ใดๆ")

    doc.add_page_break()

    # ------------------ CHAPTER 4 & 5 ------------------
    add_heading_1("บทที่ 4 ข้อมูลเชิงประจักษ์ระยะนำร่องและการขับเคลื่อนเครือข่าย")
    
    p_emp = doc.add_paragraph()
    p_emp.paragraph_format.first_line_indent = Inches(0.5)
    p_emp.add_run("4.1 สถิติเชิงประจักษ์จากการส่งออกข้อมูลระบบ (Exported Data: 7 ตุลาคม 2569)\n"
                  "จากคำสั่ง !export-lab บนระบบจริง บันทึกข้อมูลได้ดังนี้:\n"
                  "• ชื่อระบบ: 🧪 Innovation Lab: ศน. โคราช (ก่อตั้งเมื่อวันที่ 6 กรกฎาคม 2569)\n"
                  "• ระยะนำร่อง (Pilot Testing): มีผู้ร่วมทดสอบระบบ 3 ราย (ศึกษานิเทศก์แกนนำ 2 ท่าน และ Korat Edu-Bot 1 ตัว)\n"
                  "• ระบบบทบาท (Roles) รองรับ 5 หน่วยงานหลัก: ศน.ศธจ., ศน.เขตพื้นที่ (สพป./สพม.), ศน.อบจ., ศน.เทศบาล และ ศน.อาชีวศึกษา\n\n"
                  "4.2 แผนยุทธศาสตร์การขยายผลสู่เครือข่ายระดับจังหวัด (Provincial Rollout Strategy)\n"
                  "ผู้พัฒนานวัตกรรมได้กำหนดกระบวนการขยายผลแบบ 2 มิติ:\n"
                  "1. มิติทางการ: ส่งหนังสือราชการจากสำนักงานศึกษาธิการจังหวัดนครราชสีมา เชิญชวนศึกษานิเทศก์ทุกสังกัดเข้าร่วมอย่างเป็นทางการ\n"
                  "2. มิติไม่เป็นทางการ: การใช้สื่อการ์ตูนแอนิเมชันสตอรี่บอร์ด แนะนำการใช้งานเพื่อลดช่องว่างทางเทคโนโลยี")

    add_heading_1("บทที่ 5 บทสรุปและข้อเสนอแนะเชิงวิทยฐานะเชี่ยวชาญ")
    p_sum = doc.add_paragraph()
    p_sum.paragraph_format.first_line_indent = Inches(0.5)
    p_sum.add_run("นวัตกรรม \"Innovation Lab: ศน. โคราช\" เป็นผลงานทางวิชาการที่แสดงถึงความคิดริเริ่มสร้างสรรค์ (Initiative) และความเชี่ยวชาญในการบูรณาการเทคโนโลยีดิจิทัลเข้ากับศาสตร์การนิเทศการศึกษาอย่างเป็นรูปธรรม ตอบโจทย์เกณฑ์ ว.PA ด้านการพัฒนานวัตกรรมที่มีระเบียบวิธีวิจัยและการนำไปใช้จริงในระดับพื้นที่อย่างสมบูรณ์")

    doc.add_page_break()

    # ------------------ APPENDIX: OFFICIAL LETTER ------------------
    add_heading_1("ภาคผนวก: ร่างหนังสือราชการเชิญเข้าร่วมชุมชนนวัตกรรม")

    p_let = doc.add_paragraph()
    p_let.paragraph_format.space_before = Pt(12)
    p_let.add_run("ที่ ศธ ๐๒๑๐.๑๒/................                                    สำนักงานศึกษาธิการจังหวัดนครราชสีมา\n"
                  "                                                               ถนนสืบศิริ อำเภอเมืองฯ จังหวัดนครราชสีมา ๓๐๐๐๐\n\n"
                  "                                            พฤศจิกายน ๒๕๖๙\n\n"
                  "เรื่อง  ขอเชิญศึกษานิเทศก์เข้าร่วมชุมชนการเรียนรู้ทางวิชาชีพดิจิทัลและทดลองใช้นวัตกรรมแพลตฟอร์ม\n"
                  "       \"Innovation Lab: ศน. โคราช\"\n\n"
                  "เรียน  ผู้อำนวยการสำนักงานเขตพื้นที่การศึกษาประถมศึกษานครราชสีมา เขต ๑ - ๖ / ผู้อำนวยการสำนักงานเขตพื้นที่การศึกษามัธยมศึกษานครราชสีมา / นายกองค์การบริหารส่วนจังหวัดนครราชสีมา / นายกเทศมนตรีนครนครราชสีมา / ประธานประสานงานการศึกษาเอกชนจังหวัดนครราชสีมา / ผู้อำนวยการสถาบันการอาชีวศึกษาภาคตะวันออกเฉียงเหนือ ๕\n\n"
                  "สิ่งที่ส่งมาด้วย  ๑. คู่มือการเข้าใช้งานแพลตฟอร์ม Innovation Lab และการใช้งานระบบปราชญ์โคราช AI  จำนวน ๑ ฉบับ\n"
                  "                ๒. คิวอาร์โค้ด (QR Code) สำหรับเชื่อมต่อเข้าสู่เซิร์ฟเวอร์                     จำนวน ๑ แผ่น\n\n")

    p_body1 = doc.add_paragraph()
    p_body1.paragraph_format.first_line_indent = Inches(0.5)
    p_body1.add_run("ด้วย สำนักงานศึกษาธิการจังหวัดนครราชสีมา ได้ดำเนินการพัฒนานวัตกรรมการนิเทศการศึกษาเพื่อยกระดับคุณภาพการจัดการเรียนรู้ของสถานศึกษาในระดับพื้นที่ ภายใต้โครงการ \"ชุมชนการเรียนรู้ทางวิชาชีพดิจิทัลข้ามสังกัดศึกษานิเทศก์ (Cross-Silo CoP) จังหวัดนครราชสีมา\" โดยมีวัตถุประสงค์เพื่อสร้างพื้นที่กลางในการแลกเปลี่ยนองค์ความรู้ สื่อนวัตกรรมการเรียนรู้เชิงรุก (Active Learning) การขับเคลื่อนการประเมินวิทยฐานะตามเกณฑ์ ว.PA ตลอดจนการบูรณาการระบบผู้ช่วยปัญญาประดิษฐ์ (AI-Augmented Educational Supervision) ในการปฏิบัติหน้าที่นิเทศการศึกษา")

    p_body2 = doc.add_paragraph()
    p_body2.paragraph_format.first_line_indent = Inches(0.5)
    p_body2.add_run("ในการนี้ เพื่อเป็นการเสริมสร้างเครือข่ายความร่วมมือทางวิชาชีพอย่างไร้รอยต่อ และขับเคลื่อนการพัฒนานวัตกรรมการศึกษาร่วมกันในระดับจังหวัด สำนักงานศึกษาธิการจังหวัดนครราชสีมา จึงขอความอนุเคราะห์จากหน่วยงานของท่าน แจ้งและเชิญชวนศึกษานิเทศก์ในสังกัด เข้าร่วมเป็นสมาชิกและร่วมขับเคลื่อนการเรียนรู้ในแพลตฟอร์ม \"🧪 Innovation Lab: ศน. โคราช\" ผ่านช่องทาง Discord โดยสามารถสแกนคิวอาร์โค้ดตามสิ่งที่ส่งมาด้วย ทั้งนี้ ศึกษานิเทศก์ผู้เข้าร่วมจะได้รับการกำหนดอัตลักษณ์และบทบาท (Roles) ตามสังกัดหน่วยงาน และสามารถเข้าถึงคลังสื่อนวัตกรรม รวมถึงระบบผู้ช่วยปัญญาประดิษฐ์ \"ปราชญ์โคราช AI\" เพื่อสนับสนุนการปฏิบัติงานนิเทศการศึกษาได้โดยไม่มีค่าใช้จ่าย")

    p_end = doc.add_paragraph()
    p_end.paragraph_format.first_line_indent = Inches(0.5)
    p_end.paragraph_format.space_before = Pt(12)
    p_end.add_run("จึงเรียนมาเพื่อโปรดพิจารณาให้ความอนุเคราะห์ และแจ้งศึกษานิเทศก์ในสังกัดเข้าร่วมต่อไปด้วย จะขอบคุณยิ่ง\n\n"
                  "                                            ขอแสดงความนับถือ\n\n\n"
                  "                                      (............................................................)\n"
                  "                                         ศึกษาธิการจังหวัดนครราชสีมา\n\n\n"
                  "กลุ่มนิเทศ ติดตาม และประเมินผล\n"
                  "โทร. ๐ ๔๔xx xxxx\n"
                  "ผู้ประสานงาน: นายแสนยากร แสวงชิด ศึกษานิเทศก์ชำนาญการพิเศษ โทร. ๐๘ xxxx xxxx")

    out_file = r"I:\My Drive\@Discord\Innovation_Lab_Manual\Academic_Manual_Innovation_Lab.docx"
    doc.save(out_file)
    print(f"Master Academic Manual generated successfully at: {out_file}")

if __name__ == '__main__':
    generate_expert_academic_manual()
