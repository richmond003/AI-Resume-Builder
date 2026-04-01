from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
import json
from pathlib import Path
from docx2pdf import convert
from schema.schema import (
    Resume,
    Header,
    Education,
    Experience,
    Project,
    TechnicalSkills
)

class _Docx_Builder():
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

        self.add_run(p, title, bold=True, size=10)

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
        #p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        self.add_run(p, label, bold=True, size=10)
        self.add_run(p, " " + value, size=10)
        return p



class Resume_Builder(_Docx_Builder):
    def __init__(self, json_data: str):
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

        self.resume_data = json_data

        self.create_doc(data=self.resume_data)

    def _header(self, header: Header):
        # Resume Name
        name_p = self.doc.add_paragraph()
        name_p.alignment =  WD_ALIGN_PARAGRAPH.CENTER
        self.add_run(name_p, text= header["full_name"], bold=True, size=16)
        name_p.paragraph_format.space_after = Pt(2)

        # Contact info under name
        contact_p = self.doc.add_paragraph()
        contact_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self.add_run(contact_p, f"{header["email"]} | {header["number"]} | {header["location"]} | ", size=10)
        self.add_hyperlink(contact_p, f"{header["linkdin"]}", f"https://{header["linkdin"]}")
        self.add_run(contact_p, " | ", size=10)
        self.add_hyperlink(contact_p, f"{header["github"]}", f"https://{header["github"]}")
        contact_p.paragraph_format.space_after = Pt(6)

    def _summary(self, summary: str):
        self.section_heading(self.doc, "Summary")
        p = self.doc.add_paragraph()
        self.add_run(p, text=summary, size=10)
        p.paragraph_format.space_after = Pt(4)

    def _education(self, education: Education): 
        self.section_heading(self.doc, label="Education")
        self.sub_heading(self.doc, f"{education["school"]} - Degree, Minor", f"{education["expected"]}")
        self.bullet_point(self.doc, text=f"GPA: {education["GPA"]} | {education["award"]}")
        self.bullet_point(self.doc, f"Relevant coursework: {", ".join(education["Relevant Courses"])}")

    def _experience(self, experience: Experience):
        self.section_heading(self.doc, "Work Experience")
        for job in experience:
            self.sub_heading(self.doc, job["job_title"], job["timeline"])
            self.org_line(self.doc, f"{job["organization"]} | {job["location"]}")
            for task in job["responsiblities"]:
                self.bullet_point(self.doc,  task)


    def _projects(self, projects: Project):
        self.section_heading(self.doc, "Personal Projects")
        for project in projects:
            p = self.doc.add_paragraph()
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(2)
            self.add_run(p, f"{project["name"]} — {project["description"]} (", bold=True, size=10)
            self.add_hyperlink(p, "github", f"https://{project["link"]}")
            self.add_run(p, ")", bold=True, size=10)
            for point in project["bullet_points"]:
                self.bullet_point(self.doc, point)


    def _technical_skills(self, technical_skils: TechnicalSkills):
        self.section_heading(self.doc, "Technical Skills")
        for key, val in technical_skils.items():
            self.skill_line(self.doc, label=f"{key}:", value=f"{", ".join(val)}")
    
    def _export_resume(self, filename:str = "resume", output_dir: str ="resume_ark/", keep_docx: bool = False):
        docx_path = Path(output_dir) / f"{filename}.docx"
        pdf_path = Path(output_dir) / f"{filename}.pdf"

        # save docx
        self.doc.save(docx_path)

        # convert to docx to pdf
        convert(docx_path, pdf_path)
        
        if not keep_docx:
            docx_path.unlink()

        print(f"Saved: {pdf_path}")

    def create_doc(self , data: Resume):
        # resume header section
        self._header(data["header"])

        # resume summary section
        self._summary(data["summary"])

        # resume education section
        self._education(data["education"])

        # resume work expirence section
        self._experience(data["experience"])

        # resume personal project section
        self._projects(data["projects"])

        # resume technical skills section
        self._technical_skills(data["technical_skills"])
        
        # save doc
        self._export_resume()
        

if __name__ == "__main__":
  
    json_data_path = "./schema/resume_schema.json"
    data = {}
    if Path(json_data_path).exists():
        with open(json_data_path, "r") as file:
            data = json.load(file)
        Resume_Builder(data)
            
    else:
        print("couldn't find json file")
    


    
    