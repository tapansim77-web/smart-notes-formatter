
import streamlit as st
import markdown
from bs4 import BeautifulSoup, NavigableString, Tag
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from io import BytesIO

st.set_page_config(
    page_title="Smart Notes Formatter",
    page_icon="📘",
    layout="wide"
)

st.markdown("""
<style>
.main-title {
    font-size: 34px;
    font-weight: 800;
    color: #174A7E;
}
.subtitle {
    color: #64748b;
    font-size: 15px;
}
div.stButton > button {
    border-radius: 8px;
    font-weight: 600;
}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📘 Smart Notes Formatter</div>',
            unsafe_allow_html=True)
st.markdown(
    '<div class="subtitle">AI Notes to Professional Word Document</div>',
    unsafe_allow_html=True
)
st.divider()

def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), fill)
    tc_pr.append(shading)

def set_font(run, font_name, size, color=None, bold=None):
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
    run._element.rPr.rFonts.set(qn("w:cs"), font_name)
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = RGBColor(*color)
    if bold is not None:
        run.bold = bold

def add_page_number(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run("Page ")
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    run._r.addnext(fld)

def add_inline(paragraph, element, font_name, font_size, color):
    if isinstance(element, NavigableString):
        run = paragraph.add_run(str(element))
        set_font(run, font_name, font_size, color)
        return

    if not isinstance(element, Tag):
        return

    if element.name == "br":
        paragraph.add_run().add_break()
        return

    if element.name == "img":
        return

    if element.name == "a":
        for child in element.children:
            add_inline(paragraph, child, font_name, font_size, color)
        return

    for child in element.children:
        if isinstance(child, NavigableString):
            run = paragraph.add_run(str(child))
            set_font(run, font_name, font_size, color)
            if element.name in ("strong", "b"):
                run.bold = True
            if element.name in ("em", "i"):
                run.italic = True
            if element.name == "code":
                run.font.name = "Consolas"
        else:
            add_inline(paragraph, child, font_name, font_size, color)
            if paragraph.runs:
                run = paragraph.runs[-1]
                if element.name in ("strong", "b"):
                    run.bold = True
                if element.name in ("em", "i"):
                    run.italic = True
                if element.name == "del":
                    run.font.strike = True

def add_images_from_tag(tag, paragraph, uploaded_images):
    for img in tag.find_all("img"):
        src = img.get("src", "")
        filename = src.split("/")[-1]
        if filename in uploaded_images:
            try:
                paragraph.add_run().add_picture(
                    BytesIO(uploaded_images[filename]),
                    width=Inches(4.5)
                )
            except Exception:
                pass

def add_table(doc, table_tag, font_name, font_size):
    rows = table_tag.find_all("tr")
    if not rows:
        return

    max_cols = max(
        len(row.find_all(["th", "td"], recursive=False))
        for row in rows
    )

    table = doc.add_table(rows=0, cols=max_cols)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for row_index, html_row in enumerate(rows):
        cells = html_row.find_all(["th", "td"], recursive=False)
        new_cells = table.add_row().cells

        for col_index, html_cell in enumerate(cells):
            cell = new_cells[col_index]
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            if row_index == 0:
                set_cell_shading(cell, "DCEAF7")
            p = cell.paragraphs[0]
            add_inline(p, html_cell, font_name, font_size, (35, 45, 60))
            if row_index == 0:
                for run in p.runs:
                    run.bold = True

    doc.add_paragraph()

def add_blocks(container, doc, font_name, font_size,
               heading_colors, uploaded_images, list_level=0):

    for element in container.children:
        if isinstance(element, NavigableString):
            text = str(element).strip()
            if text:
                p = doc.add_paragraph()
                add_inline(p, element, font_name, font_size,
                           (35, 45, 60))
            continue

        if not isinstance(element, Tag):
            continue

        name = element.name

        if name in ("h1", "h2", "h3"):
            level = int(name[1])
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(6)
            color = heading_colors[level - 1]
            add_inline(p, element, font_name,
                       {1: 18, 2: 15, 3: 13}[level], color)
            for run in p.runs:
                run.bold = True

        elif name == "p":
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(6)
            p.paragraph_format.line_spacing = 1.15
            add_inline(p, element, font_name, font_size,
                       (35, 45, 60))
            add_images_from_tag(element, p, uploaded_images)

        elif name in ("ul", "ol"):
            style = "List Bullet" if name == "ul" else "List Number"
            for li in element.find_all("li", recursive=False):
                p = doc.add_paragraph(style=style)
                p.paragraph_format.left_indent = Inches(
                    0.25 * (list_level + 1)
                )
                add_inline(p, li, font_name, font_size,
                           (35, 45, 60))
                add_images_from_tag(li, p, uploaded_images)

        elif name == "table":
            add_table(doc, element, font_name, font_size)

        elif name == "blockquote":
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.3)
            add_inline(p, element, font_name, font_size,
                       (80, 90, 110))

        elif name == "pre":
            p = doc.add_paragraph()
            run = p.add_run(element.get_text())
            set_font(run, "Consolas", 9, (50, 50, 50))

        elif name == "hr":
            doc.add_paragraph()

        elif name == "div":
            add_blocks(element, doc, font_name, font_size,
                       heading_colors, uploaded_images, list_level)

        elif name == "img":
            add_images_from_tag(element, doc.add_paragraph(),
                                uploaded_images)

        else:
            if name not in ("script", "style"):
                add_blocks(element, doc, font_name, font_size,
                           heading_colors, uploaded_images, list_level)

def create_document(content, title, author, institution,
                    font_name, font_size, heading_colors,
                    uploaded_images, cover_page):

    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.75)
    section.bottom_margin = Inches(0.75)
    section.left_margin = Inches(0.85)
    section.right_margin = Inches(0.75)

    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = font_name
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
    normal._element.rPr.rFonts.set(qn("w:cs"), font_name)
    normal.font.size = Pt(font_size)

    for style_name in ["List Bullet", "List Number"]:
        style = styles[style_name]
        style.font.name = font_name
        style._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)

    header = section.header.paragraphs[0]
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = header.add_run(institution)
    set_font(r, font_name, 9, (100, 100, 100))

    footer = section.footer.paragraphs[0]
    add_page_number(footer)

    if cover_page:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(110)
        r = p.add_run(title)
        set_font(r, font_name, 25, (25, 70, 120), True)

        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run("अध्ययन सामग्री")
        set_font(r, font_name, 17, (30, 120, 95), True)

        doc.add_paragraph()
        for value in [author, institution]:
            if value.strip():
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run(value)
                set_font(r, font_name, 12, (70, 70, 70))

        doc.add_page_break()

    html = markdown.markdown(
        content,
        extensions=["tables", "fenced_code", "sane_lists"]
    )
    soup = BeautifulSoup(html, "html.parser")
    add_blocks(soup, doc, font_name, font_size,
               heading_colors, uploaded_images)

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# SIDEBAR SETTINGS
st.sidebar.header("⚙️ Document Settings")

title = st.sidebar.text_input("Document Title", "अध्ययन सामग्री")
author = st.sidebar.text_input("Prepared By", "")
institution = st.sidebar.text_input("Institution Name", "")

font_name = st.sidebar.selectbox(
    "Hindi Font",
    ["Noto Sans Devanagari", "Mangal", "Aparajita", "Arial"]
)

font_size = st.sidebar.select_slider(
    "Body Font Size",
    options=[10, 11, 12, 13, 14],
    value=12
)

cover_page = st.sidebar.checkbox("Create Cover Page", True)

st.sidebar.subheader("Heading Colors")
heading_colors = [
    (23, 74, 126),
    (20, 120, 90),
    (190, 95, 25)
]

st.sidebar.caption("Heading 1: Blue | Heading 2: Green | Heading 3: Orange")

st.subheader("📝 Paste Your Notes")

content = st.text_area(
    "Markdown / ChatGPT Notes",
    height=400,
    placeholder="""# मुख्य शीर्षक

## पहला अध्याय

यहाँ अपना नोट्स पेस्ट करें।

### महत्वपूर्ण बिंदु
- पहला बिंदु
- दूसरा बिंदु

| विषय | विवरण |
|---|---|
| उदाहरण | जानकारी |
"""
)

st.subheader("🖼️ Upload Images")
image_files = st.file_uploader(
    "Upload images (PNG/JPG)",
    type=["png", "jpg", "jpeg"],
    accept_multiple_files=True
)

uploaded_images = {
    file.name: file.getvalue()
    for file in image_files
} if image_files else {}

st.info(
    "चित्र जोड़ने के लिए Markdown में ![विवरण](filename.png) लिखें "
    "और उसी नाम की image upload करें।"
)

if st.button("📄 Generate Word Document", type="primary",
             use_container_width=True):

    if not content.strip():
        st.error("कृपया पहले अपने नोट्स पेस्ट करें।")
    else:
        try:
            with st.spinner("आपकी Word फाइल तैयार हो रही है..."):
                word_file = create_document(
                    content, title, author, institution,
                    font_name, font_size, heading_colors,
                    uploaded_images, cover_page
                )

            filename = "Smart_Notes_Formatter.docx"

            st.success("आपकी Word फाइल तैयार है!")

            st.download_button(
                "⬇️ Download Word Document",
                data=word_file.getvalue(),
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

        except Exception as e:
            st.error(f"फाइल बनाने में समस्या आई: {e}")

st.divider()
st.caption("Smart Notes Formatter | Version 1.0")
