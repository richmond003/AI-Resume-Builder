import pdfplumber
from rich import print

def extract_text_pdf(pdf_path: str) -> str:
    content = ''

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                content +=  text + "\n"
    return content


def extract_links(pdf_path: str) -> list[dict]:
    links = []

    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            # annotations contain hyperlinks
            if page.annots:
                for annot in page.annots:
                    uri = annot.get("uri")          # the actual URL
                    if uri:
                        links.append({
                            "page": i + 1,
                            "url": uri,
                        })

    return links





if __name__  == "__main__":
    # data = extract_text_pdf("./resume_ark/resume_original.pdf")
    # print(data)
    links = extract_links("resume_ark/resume_original.pdf")

    for link in links:
        print(f"Page {link['page']}: {link['url']}")
