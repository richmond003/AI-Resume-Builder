from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

class _Docx_Builder:
    def __init__(self):
        pass

    def add_run(self, paragraph, text, bold=False, italic=False, size=10, underline=False):
        run = paragraph.add_run(text)
        run.bold = bold
        run.italic = italic
        run.underline = underline
        run.font.size = Pt(size)
        return run
    
    def add_hyperlink(self, paragraph: Paragraph, text: str, url: str):
        """ adds clickable hyperlink to a paragraph """
        part = paragraph.part
        r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
        hyperlink = OxmlElement("w:hyperlink")
        hyperlink.set(qn("r:id"), r_id)
        run_elem = OxmlElement("w:r")
        rpr = OxmlElement("w:rPr")
        style = OxmlElement("w:rStyle")
        style.set(qn("w:val"), "Hyperlink")
        rpr.append(style)
        run_elem.append(rpr)
        t = OxmlElement("w:t")
        t.text = text
        run_elem.append(t)
        hyperlink.append(run_elem)
        paragraph._p.append(hyperlink)

    def section_heading(self, doc: Document, label):
        """ all-caps bold heading with a bottom border line """
        p = doc.add_paragraph()
        run = p.add_run(label.upper())
        run.bold = True
        run.font.size = Pt(11)

        # bottom border under the headaing
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "6")
        bottom.set(qn("w:space"), "1")
        bottom.set(qn("w:color"), "000000")
        pBdr.append(bottom)
        pPr.append(pBdr)

        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        return p

    def sub_heading(self,doc, title, date):
        """ bold job/school title with right-aligned date on the same line """
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)

        self.add_run(self, p, title, bold=True, size=10)

        # Tab to push date to the right margin
        self.add_run(p, "\t", size=10)
        self.add_run(p, date, italic=True, size=10)

        # Right-aligned tab stop at ~6.5 inches
        pPr = p._p.get_or_add_pPr()
        tabs = OxmlElement("w:tabs")
        tab = OxmlElement("w:tab")
        tab.set(qn("w:val"), "right")
        tab.set(qn("w:pos"), "9360")   # DXA: 9360 = 6.5 inches
        tabs.append(tab)
        pPr.append(tabs)
        return p
    
    def org_line(self, doc, text):
        """ Italic and underline orgization and location line """
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2)
        self.add_run(p, text, italic=True, underline=True, size=10)
        return p
    
    def bullet_point(self,doc, text):
        """ standard bullet point using the built-in list Bullet style """
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2)
        self.add_run(p, text, size=10)
        return p

    def skill_line(self, doc, label, value):
        """ Bold label follewed by normal value """
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        self.add_run(p, label, bold=True, size=10)
        self.add_run(p, " " + value, size=10)
        return p

class Resume_Builder(_Docx_Builder):
    def __init__(self):
        self.doc = Document()
        self.section = self.doc.sections[0]
        self.section.page_width = Inches(8.5)
        self.section.page_height= Inches(11)
        self.section.top_margin = Inches(0.5)
        self.section.bottom_margin = Inches(0.5)
        self.section.left_margin = Inches(0.75)
        self.section.right_margin = Inches(0.75)

        # Default document font
        self.doc.styles["Normal"].font.name = "Calibri"
        self.doc.styles["Normal"].font.size = Pt(10)

        self.create_doc()

    def _header(self, header_info):
        # Resume Name
        name_p = self.doc.add_paragraph()
        name_p.alignment =  WD_ALIGN_PARAGRAPH.CENTER
        self.add_run(name_p, text= header_info["name"] or "FUll NAME", bold=True, size=16)
        name_p.paragraph_format.space_after = Pt(2)

        # Contact info under name
        contact_p = self.doc.add_paragraph()
        contact_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self.add_run(contact_p, "email@example.com | (000)-000-0000 | City, State | ", size=10)
        self.add_hyperlink(contact_p, "linkedin.com/in/yourhandle", "https://linkedin.com/in/yourhandle")
        self.add_run(contact_p, " | ", size=10)
        self.add_hyperlink(contact_p, "github.com/yourhandle", "https://github.com/yourhandle")
        contact_p.paragraph_format.space_after = Pt(6)

    def _summary(self):
        self.section_heading(self.doc, "Summary")
        p = self.doc.add_paragraph()
        self.add_run(p, text="One-paragraph professional summary goes here.", size=10)
        p.paragraph_format.space_after = Pt(4)

    def _education(self):
        self.section_heading(self.doc, text="Education")
        self.sub_heading(self.doc, "University Name - Degree, Minor", "Expected: Month")
        self.bullet_point(self.doc, text="GPA: X.XX | Award Name (Year)")
        self.bullet_point(self.doc, "Relevant coursework: Course A, Course B, Course C")

    def _experience(self, experience):
        self.section_heading(self.doc, "Work Experience")
        for job in experience:
            self.sub_heading(self.doc, job["title"], job["year"])
            self.org_line(self.doc, "Organization Nanme | City, State")
            for point in job["experience"]:
                self.bullet_point(self.doc,  "Achievement / responsibility with metric.")


    def _projects(self, projects):
        pass


    def _technical_skills(self, skills):
        for skill in skills:
            self.skill_line(self.doc, f"{skill.key} : {", ". join(skill)}")
       

    def create_doc(self):
        self