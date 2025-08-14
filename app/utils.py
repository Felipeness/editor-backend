from docx.text.paragraph import Paragraph
from docx.opc.constants import RELATIONSHIP_TYPE as RT

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def add_hyperlink(paragraph: Paragraph, url: str, text: str) -> None:
    part = paragraph.part
    # Use a constante oficial:
    r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    run = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")
    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0000EE")
    r_pr.append(u)
    r_pr.append(color)

    t = OxmlElement("w:t")
    t.text = text

    run.append(r_pr)
    run.append(t)
    hyperlink.append(run)


    paragraph._p.append(hyperlink)
