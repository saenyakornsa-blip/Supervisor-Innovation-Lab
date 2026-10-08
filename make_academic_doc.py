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

def create_academic_manual():
    doc = Document()
    
    # Page Setup - Margins
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.25)
        section.right_margin = Inches(1.0)

    # Styles
    normal_style = doc.styles['Normal']
    normal_font = normal_style.font
    normal_font.name = 'TH Sarabun PSK'
    normal_font.size = Pt(16)
    normal_font.color.rgb = RGBColor(0x22, 0x22, 0x22)

    # Helper function for headings
    def add_custom_heading(text, level, space_before=12, space_after=6):
        h = doc.add_heading(text, level=level)
        h.paragraph_format.space_before = Pt(space_before)
        h.paragraph_format.space_after = Pt(space_after)
        run = h.runs[0]
        run.font.name = 'TH Sarabun PSK'
        if level == 1:
            run.font.size = Pt(20)
            run.bold = True
            run.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D) # Navy Blue
        elif level == 2:
            run.font.size = Pt(18)
            run.bold = True
            run.font.color.rgb = RGBColor(0x2E, 0x5B, 0x88)
        elif level == 3:
            run.font.size = Pt(16)
            run.bold = True
            run.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
        return h

    # ------------------ COVER PAGE ------------------
    cover_p1 = doc.add_paragraph()
    cover_p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cover_p1.paragraph_format.space_before = Pt(72)
    run_org = cover_p1.add_run("เอกสารประกอบการประเมินวิทยฐานะ ด้านที่ 3 ผลงานทางวิชาการ/นวัตกรรม")
    run_org.font.size = Pt(16)
    run_org.bold = True

    cover_title = doc.add_paragraph()
    cover_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cover_title.paragraph_format.space_before = Pt(36)
    cover_title.paragraph_format.space_after = Pt(18)
    run_t1 = cover_title.add_run("คู่มือการใช้งานและรายงานเชิงสถาปัตยกรรมนวัตกรรมดิจิทัล\n")
    run_t1.font.size = Pt(24)
    run_t1.bold = True
    run_t1.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    run_t2 = cover_title.add_run("ชุมชนการเรียนรู้ทางวิชาชีพดิจิทัลข้ามสังกัดศึกษานิเทศก์\n\"Innovation Lab: ศน. โคราช\" ร่วมกับระบบพี่เลี้ยงปัญญาประดิษฐ์ (AI-Augmented Supervision)")
    run_t2.font.size = Pt(20)
    run_t2.bold = True
    run_t2.font.color.rgb = RGBColor(0x33, 0x5C, 0x8D)

    cover_info = doc.add_paragraph()
    cover_info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cover_info.paragraph_format.space_before = Pt(120)
    
    r_author = cover_info.add_run("ผู้จัดทำและพัฒนานวัตกรรม\nนายแสนยากร แสวงชิด\nตำแหน่ง ศึกษานิเทศก์ วิทยฐานะศึกษานิเทศก์ชำนาญการพิเศษ\nสำนักงานศึกษาธิการจังหวัดนครราชสีมา\nกระทรวงศึกษาธิการ")
    r_author.font.size = Pt(18)
    r_author.bold = True

    doc.add_page_break()

    # ------------------ CHAPTER 1 ------------------
    add_custom_heading("บทที่ 1 บทนำและหลักการและเหตุผลเชิงยุทธศาสตร์", level=1)

    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Inches(0.5)
    p.add_run("1.1 ความเป็นมาและสภาพปัญหา (Problem Statement)\nการจัดการศึกษาในระดับพื้นที่จังหวัดนครราชสีมา มีความหลากหลายของบริบทโรงเรียนและโครงสร้างการบริหารจัดการ โดยมีศึกษานิเทศก์ (ศน.) สังกัดหน่วยงานทางการศึกษาต่างๆ ปฏิบัติหน้าที่นิเทศ กำกับ ติดตาม และพัฒนาคุณภาพการศึกษา ประกอบด้วย สำนักงานเขตพื้นที่การศึกษาประถมศึกษา (สพป.) นครราชสีมา เขต 1-6, สำนักงานเขตพื้นที่การศึกษามัธยมศึกษา (สพม.) นครราชสีมา, สำนักงานคณะกรรมการการส่งเสริมการศึกษาเอกชน (สช.), และองค์กรปกครองส่วนท้องถิ่น (อปท.)")

    p2 = doc.add_paragraph()
    p2.paragraph_format.first_line_indent = Inches(0.5)
    p2.add_run("อย่างไรก็ตาม ปัญหาสำคัญในเชิงโครงสร้างที่ส่งผลต่อประสิทธิภาพการนิเทศการศึกษา ได้แก่:")
    
    bullets = [
        ("ปัญหาการทำงานแบบแยกส่วน (Silo Mentality): ", "ศึกษานิเทศก์แต่ละสังกัดขาดพื้นที่กลางในการแลกเปลี่ยนองค์ความรู้ สื่อนวัตกรรม และประสบการณ์แก้ปัญหา ทั้งที่ปลายทางคือการพัฒนาผู้เรียนในผืนแผ่นดินโคราชเหมือนกัน"),
        ("ภาระงานเอกสารและความท้าทายตามเกณฑ์ ว.PA: ", "การขับเคลื่อนการประเมินตำแหน่งและวิทยฐานะตามเกณฑ์ ว.PA และการจัดการเรียนรู้เชิงรุก (Active Learning) ต้องอาศัยการเข้าถึงข้อมูลระเบียบและตัวอย่างที่รวดเร็ว ถูกต้อง"),
        ("ข้อจำกัดของช่องทางการสื่อสารเดิม: ", "การใช้แอปพลิเคชันส่งข้อความทั่วไป (เช่น LINE) มักประสบปัญหาไฟล์และรูปภาพหมดอายุ ข้อความสำคัญตกหล่น และไม่สามารถจัดหมวดหมู่คลังความรู้ได้อย่างเป็นระบบ ส่วนระบบ LMS แบบเดิม (เช่น Google Classroom) มีโครงสร้างแบบ Top-Down ไม่เอื้อต่อการมีปฏิสัมพันธ์แบบ Many-to-Many")
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
               "1. เพื่อออกแบบและพัฒนาแพลตฟอร์มชุมชนการเรียนรู้ทางวิชาชีพดิจิทัลข้ามสังกัด (Cross-Silo CoP) สำหรับศึกษานิเทศก์จังหวัดนครราชสีมา บนแพลตฟอร์ม Discord\n"
               "2. เพื่อพัฒนาระบบผู้ช่วยปัญญาประดิษฐ์แบบผสมผสาน (Hybrid AI Assistant: Korat Edu-Bot ร่วมกับ NotebookLM) เพื่อยกระดับสมรรถนะการนิเทศการศึกษา (AI-Augmented Supervision)\n"
               "3. เพื่อประเมินผลการยอมรับ การมีส่วนร่วม และการนำนวัตกรรมทางการศึกษาไปประยุกต์ใช้ในการปฏิบัติงานจริงของศึกษานิเทศก์")

    doc.add_page_break()

    # ------------------ CHAPTER 2 ------------------
    add_custom_heading("บทที่ 2 กรอบแนวคิดเชิงทฤษฎีและสถาปัตยกรรมนวัตกรรม", level=1)

    p_theo = doc.add_paragraph()
    p_theo.paragraph_format.first_line_indent = Inches(0.5)
    p_theo.add_run("2.1 กรอบแนวคิดเชิงทฤษฎี (Theoretical Framework)\nการพัฒนานวัตกรรม \"Innovation Lab: ศน. โคราช\" ตั้งอยู่บนฐานทฤษฎีและแนวคิดสากล 3 ประการหลัก:")

    theos = [
        ("1. ชุมชนการเรียนรู้ทางวิชาชีพ (Community of Practice - CoP) ของ Etienne Wenger: ", 
         "มุ่งเน้นการสร้าง 3 องค์ประกอบหลัก ได้แก่ ขอบเขตความสนใจร่วมกัน (Domain), ความสัมพันธ์และการแลกเปลี่ยนในชุมชน (Community), และคลังเครื่องมือ/แนวปฏิบัติร่วม (Practice)"),
        ("2. ทฤษฎีการเชื่อมโยงการเรียนรู้ดิจิทัล (Connectivism) ของ George Siemens: ", 
         "การเรียนรู้ในยุคดิจิทัลไม่ได้เกิดขึ้นเฉพาะในตัวบุคคล แต่อยู่ที่ความสามารถในการเชื่อมโยงโหนด (Nodes) ข้อมูล เครือข่าย และเครื่องมือดิจิทัลเข้าด้วยกัน"),
        ("3. การนิเทศการศึกษาเสริมพลังด้วยปัญญาประดิษฐ์ (AI-Augmented Educational Supervision): ", 
         "การนำเทคโนโลยี AI มาทำหน้าที่เป็น Co-Pilot หรือพี่เลี้ยงคู่คิด ช่วยลดภาระงานรูทีนและเร่งกระบวนการสืบค้นข้อมูลเชิงลึก")
    ]
    for t_title, t_desc in theos:
        tp = doc.add_paragraph()
        tp.paragraph_format.left_indent = Inches(0.5)
        run_t = tp.add_run(t_title)
        run_t.bold = True
        tp.add_run(t_desc)

    add_custom_heading("2.2 สถาปัตยกรรมโครงสร้างชุมชน (Channel Architecture)", level=2)
    p_arch = doc.add_paragraph()
    p_arch.paragraph_format.first_line_indent = Inches(0.5)
    p_arch.add_run("เซิร์ฟเวอร์ถูกออกแบบเชิงสถาปัตยกรรมข้อมูล (Information Architecture) โดยแบ่งพื้นที่ตามฟังก์ชันการทำงาน เพื่อลด Cognitive Overload:")

    # Table of Channels
    table = doc.add_table(rows=1, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'หมวดหมู่ (Category)'
    hdr_cells[1].text = 'ชื่อห้อง (Channels)'
    hdr_cells[2].text = 'วัตถุประสงค์เชิงการนิเทศและการจัดการความรู้'
    for cell in hdr_cells:
        set_cell_background(cell, '1B365D')
        for cp in cell.paragraphs:
            for cr in cp.runs:
                cr.font.name = 'TH Sarabun PSK'
                cr.font.size = Pt(16)
                cr.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                cr.bold = True

    channels_data = [
        ("Information\n(จุดประชาสัมพันธ์และคลังความรู้)", "#👋-ต้อนรับและกติกา\n#📢-ประกาศ-แจ้งข่าว\n#📚-คลังสื่อและนวัตกรรม\n#🏆-โชว์ผลงาน", "สร้างปฐมนิเทศดิจิทัล กระจายข่าวสารนโยบายสำคัญ และสร้างคลังสื่อนวัตกรรมระบบฟอรั่ม (Forum Tagging) พร้อมแสดงแกลเลอรีผลงานถอดบทเรียน"),
        ("Text Channels\n(ลานเสวนาและปรึกษา)", "#☕-สภากาแฟ-คุยทั่วไป\n#📅-นัดหมาย-ปรึกษางาน\n#🎉-ห้องพักใจ-เล่าเรื่อง\n#🤖-ai-เพื่อนคู่คิด-ศึกษานิเทศก์", "สร้างพื้นที่ปลอดภัย (Safe Space) สานสัมพันธ์ข้ามสังกัด ปรึกษาข้อราชการลดความเครียด และเป็นห้องปฏิบัติการสั่งการปัญญาประดิษฐ์"),
        ("Voice Channels\n(ห้องสนทนาเสียง/ประชุม)", "🔊 ชวนกันเล่นเกม\n🔊 มุมกาแฟ (คุยเล่น)\n🔊 ห้องประชุม 1", "รองรับการประชุมออนไลน์ การแชร์หน้าจออบรมเชิงปฏิบัติการ และการสื่อสารอย่างไม่เป็นทางการ")
    ]

    for cat, ch, desc in channels_data:
        row_cells = table.add_row().cells
        row_cells[0].text = cat
        row_cells[1].text = ch
        row_cells[2].text = desc
        for cell in row_cells:
            for cp in cell.paragraphs:
                for cr in cp.runs:
                    cr.font.name = 'TH Sarabun PSK'
                    cr.font.size = Pt(15)

    doc.add_page_break()

    # ------------------ CHAPTER 3 ------------------
    add_custom_heading("บทที่ 3 สถาปัตยกรรมระบบผู้ช่วย AI แบบแท็กทีม (Hybrid AI Architecture)", level=1)
    
    p_ai = doc.add_paragraph()
    p_ai.paragraph_format.first_line_indent = Inches(0.5)
    p_ai.add_run("นวัตกรรมนี้ได้บูรณาการระบบผู้ช่วย AI แบบ 2 ประสาน (Tag-Team Dual Architecture) เพื่อรองรับลักษณะงานทางวิชาการที่แตกต่างกันอย่างสิ้นเชิง:")

    add_custom_heading("3.1 ผู้ช่วยส่วนที่ 1: \"ปราชญ์โคราช AI\" (Fast Generative Co-Pilot)", level=2)
    p_ai1 = doc.add_paragraph()
    p_ai1.paragraph_format.left_indent = Inches(0.5)
    p_ai1.add_run("• โครงสร้างทางเทคนิค: ขับเคลื่อนด้วย Large Language Model (Google Gemini 1.5 Flash API) เชื่อมต่อผ่าน Node.js (Discord.js Library) โฮสต์บนระบบคลาวด์\n"
                  "• ภารกิจหลัก: ระดมสมอง (Brainstorming), การร่างเค้าโครงการนิเทศ, การออกแบบกิจกรรม Active Learning, การเขียนคำถามสะท้อนคิด (Reflective Questioning)\n"
                  "• กลไกคำสั่ง (Persona Engineering): ถูกโปรแกรมให้มีบทบาทเป็นกัลยาณมิตรทางวิชาการที่สุภาพ เสริมสร้างพลังใจ และเชี่ยวชาญศาสตร์การสอนร่วมสมัย")

    add_custom_heading("3.2 ผู้ช่วยส่วนที่ 2: \"ฐานข้อมูล NotebookLM\" (Grounded Regulatory Retrieval)", level=2)
    p_ai1 = doc.add_paragraph()
    p_ai1.paragraph_format.left_indent = Inches(0.5)
    p_ai1.add_run("• โครงสร้างทางเทคนิค: Retrieval-Augmented Generation (RAG) ปักหมุดในห้อง #📚-คลังสื่อและนวัตกรรม\n"
                  "• ภารกิจหลัก: ตอบคำถามและสืบค้นข้อบังคับที่มีความเที่ยงตรง 100% เช่น เกณฑ์ ว.PA, มาตรฐาน สมศ., กฎกระทรวง และคู่มือนิเทศประจำจังหวัด\n"
                  "• จุดเด่น: ไม่เกิดอาการหลอนข้อมูล (Hallucination) เนื่องจากตอบจากฐานเอกสารทางการที่อัปโหลดไว้พร้อมการอ้างอิงเลขหน้าอย่างแม่นยำ")

    add_custom_heading("3.3 แผนผังเปรียบเทียบการเลือกใช้งานระบบ AI", level=2)
    ai_table = doc.add_table(rows=1, cols=3)
    ai_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    ai_hdrs = ai_table.rows[0].cells
    ai_hdrs[0].text = 'ประเด็นการประเมิน'
    ai_hdrs[1].text = 'ปราชญ์โคราช AI (Bot)'
    ai_hdrs[2].text = 'ฐานข้อมูล NotebookLM'
    for c in ai_hdrs:
        set_cell_background(c, '2E5B88')
        for cp in c.paragraphs:
            for cr in cp.runs:
                cr.font.name = 'TH Sarabun PSK'
                cr.font.size = Pt(16)
                cr.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                cr.bold = True

    ai_data = [
        ("ลักษณะงานที่เหมาะสม", "สร้างสรรค์, ร่างโครงการ, ไอเดียสอน", "ระเบียบกฎหมาย, มาตรฐาน, ว.PA"),
        ("รูปแบบการตอบสนอง", "รวดเร็ว Real-time โต้ตอบในห้องแชท", "สังเคราะห์พร้อม Citation อ้างอิงเอกสาร"),
        ("ความแม่นยำของระเบียบ", "ปานกลาง (อาจมีความคลาดเคลื่อน)", "สูงมาก (ตอบเฉพาะจากคลังเอกสาร)")
    ]
    for row in ai_data:
        cells = ai_table.add_row().cells
        cells[0].text = row[0]
        cells[1].text = row[1]
        cells[2].text = row[2]
        for c in cells:
            for cp in c.paragraphs:
                for cr in cp.runs:
                    cr.font.name = 'TH Sarabun PSK'
                    cr.font.size = Pt(15)

    doc.add_page_break()

    # ------------------ CHAPTER 4 & 5 ------------------
    add_custom_heading("บทที่ 4 วัฒนธรรมชุมชน กฎบัตร และแนวทางการขับเคลื่อนเชิงประจักษ์", level=1)
    
    p_cul = doc.add_paragraph()
    p_cul.paragraph_format.first_line_indent = Inches(0.5)
    p_cul.add_run("เพื่อให้ชุมชนเกิดความยั่งยืน ไม่กลายเป็นกลุ่มร้าง จึงได้กำหนด \"กฎบัตร 3 ข้อ (Three Pillars of Culture)\" ในการขับเคลื่อน:")

    pillars = [
        ("1. วัฒนธรรมพื้นที่ปลอดภัย (Psychological Safety & Safe Space): ", "เปิดกว้างให้สมาชิกกล้าแลกเปลี่ยนนวัตกรรมที่อยู่ระหว่างการพัฒนา โดยปราศจากการตัดสินเชิงลบ เพื่อกระตุ้นวัฒนธรรมการช่วยเหลือและเติมเต็ม"),
        ("2. การแบ่งปันอย่างมีบริบท (Contextual Knowledge Sharing): ", "การเผยแพร่นวัตกรรมต้องระบุบริบทสั้นๆ 3 องค์ประกอบ (ปัญหาที่พบ -> วิธีการแก้ -> ผลสัมฤทธิ์) เพื่อให้เพื่อนศึกษานิเทศก์นำไปต่อยอดได้จริง"),
        ("3. เอกภาพบนความหลากหลาย (Cross-Silo Synergy): ", "การรวมพลังของ ศน. สพป., สพม., สช., และ อปท. โดยใช้สีประจำสังกัดสร้างอัตลักษณ์เชิงบวก (Identity Affirmation) แต่มีเป้าหมายเดียวกันคือเด็กและเยาวชนโคราช")
    ]
    for p_title, p_desc in pillars:
        pp = doc.add_paragraph()
        pp.paragraph_format.left_indent = Inches(0.5)
        run_p = pp.add_run(p_title)
        run_p.bold = True
        pp.add_run(p_desc)

    add_custom_heading("บทที่ 5 การประเมินผลและการสะท้อนคิดเพื่อขอรับการประเมินวิทยฐานะ", level=1)
    p_eval = doc.add_paragraph()
    p_eval.paragraph_format.first_line_indent = Inches(0.5)
    p_eval.add_run("ในการนำเสนอนวัตกรรมเพื่อขอรับการประเมินวิทยฐานะตามเกณฑ์ ว.PA (ด้านที่ 3) ได้กำหนดกรอบตัวชี้วัดเชิงประจักษ์ (Evidence-based Indicators) ไว้ดังนี้:")

    evidences = [
        ("มิติด้านปริมาณ (Quantitative Indicators): ", "จำนวนสมาชิกศึกษานิเทศก์ที่เข้าร่วมข้ามสังกัด, จำนวนคำถามและปฏิสัมพันธ์กับระบบ AI, จำนวนสื่อนวัตกรรมที่ได้รับการอัปโหลดในคลังสื่อ"),
        ("มิติด้านคุณภาพ (Qualitative Indicators): ", "ผลการประเมินความพึงพอใจของศึกษานิเทศก์, ตัวอย่างการนำสื่อหรือแผนนิเทศจากชุมชนไปขยายผลในโรงเรียน, การสะท้อนคิด (AAR) จากการนิเทศร่วม"),
        ("การเผยแพร่และขยายผล (Dissemination & Impact): ", "การเป็นต้นแบบเครือข่ายวิชาชีพดิจิทัลระดับจังหวัด ที่สามารถถอดบทเรียนขยายผลสู่จังหวัดอื่นในภูมิภาคได้")
    ]
    for e_title, e_desc in evidences:
        ep = doc.add_paragraph()
        ep.paragraph_format.left_indent = Inches(0.5)
        run_e = ep.add_run(e_title)
        run_e.bold = True
        ep.add_run(e_desc)

    # Output path
    out_dir = r"I:\My Drive\@Discord\Innovation_Lab_Manual"
    os.makedirs(out_dir, exist_ok=True)
    out_file = os.path.join(out_dir, "Academic_Manual_Innovation_Lab.docx")
    doc.save(out_file)
    print(f"Academic Manual created successfully at: {out_file}")

if __name__ == '__main__':
    create_academic_manual()
