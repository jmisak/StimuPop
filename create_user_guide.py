"""
StimuPop User Guide Generator

Creates a professional, graphically appealing DOCX user guide
for the StimuPop Excel to PowerPoint converter application.
"""

from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def set_cell_shading(cell, color_hex):
    """Set background color for a table cell."""
    shading = OxmlElement('w:shd')
    shading.set(qn('w:fill'), color_hex)
    cell._tc.get_or_add_tcPr().append(shading)


def add_horizontal_line(doc):
    """Add a horizontal line/divider."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)

    # Create bottom border
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '4472C4')
    pBdr.append(bottom)
    pPr.append(pBdr)


def create_styled_heading(doc, text, level=1):
    """Create a styled heading with custom formatting."""
    heading = doc.add_heading(text, level=level)

    # Style based on level
    if level == 1:
        heading.runs[0].font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)
        heading.runs[0].font.size = Pt(24)
    elif level == 2:
        heading.runs[0].font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)
        heading.runs[0].font.size = Pt(18)
    elif level == 3:
        heading.runs[0].font.color.rgb = RGBColor(0x44, 0x72, 0xC4)
        heading.runs[0].font.size = Pt(14)

    return heading


def add_info_box(doc, title, content, box_color="E7F3FF", border_color="2E74B5"):
    """Add a styled information box."""
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.cell(0, 0)

    # Set cell background
    set_cell_shading(cell, box_color)

    # Set border
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for border_name in ['top', 'left', 'bottom', 'right']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '12')
        border.set(qn('w:color'), border_color)
        tcBorders.append(border)
    tcPr.append(tcBorders)

    # Add title
    title_para = cell.paragraphs[0]
    title_run = title_para.add_run(title)
    title_run.bold = True
    title_run.font.size = Pt(12)
    title_run.font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)

    # Add content
    content_para = cell.add_paragraph()
    content_run = content_para.add_run(content)
    content_run.font.size = Pt(11)

    doc.add_paragraph()  # Spacing


def add_warning_box(doc, content):
    """Add a warning/caution box."""
    add_info_box(doc, "⚠️ Important", content, "FFF3CD", "856404")


def add_tip_box(doc, content):
    """Add a tip/hint box."""
    add_info_box(doc, "💡 Tip", content, "D4EDDA", "155724")


def add_step_table(doc, steps):
    """Add a numbered steps table with visual styling."""
    table = doc.add_table(rows=len(steps), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Set column widths
    table.columns[0].width = Inches(0.6)
    table.columns[1].width = Inches(5.4)

    for idx, (step_num, step_text) in enumerate(steps):
        # Number cell
        num_cell = table.cell(idx, 0)
        set_cell_shading(num_cell, "2E74B5")
        num_para = num_cell.paragraphs[0]
        num_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        num_run = num_para.add_run(str(step_num))
        num_run.bold = True
        num_run.font.size = Pt(14)
        num_run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

        # Text cell
        text_cell = table.cell(idx, 1)
        if idx % 2 == 0:
            set_cell_shading(text_cell, "F8F9FA")
        text_para = text_cell.paragraphs[0]
        text_para.paragraph_format.left_indent = Pt(6)
        text_run = text_para.add_run(step_text)
        text_run.font.size = Pt(11)

    doc.add_paragraph()  # Spacing


def create_user_guide():
    """Generate the complete user guide document."""
    doc = Document()

    # Set default font
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)

    # ========== TITLE PAGE ==========
    doc.add_paragraph()
    doc.add_paragraph()

    # Main title
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title.add_run("🎯 StimuPop")
    title_run.bold = True
    title_run.font.size = Pt(48)
    title_run.font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)

    # Subtitle
    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_run = subtitle.add_run("Excel to PowerPoint Converter")
    sub_run.font.size = Pt(24)
    sub_run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    doc.add_paragraph()

    # User Guide label
    guide_label = doc.add_paragraph()
    guide_label.alignment = WD_ALIGN_PARAGRAPH.CENTER
    guide_run = guide_label.add_run("USER GUIDE")
    guide_run.bold = True
    guide_run.font.size = Pt(18)
    guide_run.font.color.rgb = RGBColor(0x44, 0x72, 0xC4)

    doc.add_paragraph()
    doc.add_paragraph()

    # Version info box
    version_table = doc.add_table(rows=1, cols=1)
    version_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    version_cell = version_table.cell(0, 0)
    set_cell_shading(version_cell, "E7F3FF")
    version_para = version_cell.paragraphs[0]
    version_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    version_run = version_para.add_run("Version 2.2.0")
    version_run.font.size = Pt(14)
    version_run.font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)

    # Page break
    doc.add_page_break()

    # ========== TABLE OF CONTENTS ==========
    create_styled_heading(doc, "Table of Contents", 1)
    add_horizontal_line(doc)

    toc_items = [
        ("1.", "Introduction", "What is StimuPop?"),
        ("2.", "Getting Started", "Installation and first launch"),
        ("3.", "Preparing Your Excel File", "Data structure requirements"),
        ("4.", "Using StimuPop", "Step-by-step guide"),
        ("5.", "Configuration Options", "Customizing your output"),
        ("6.", "Troubleshooting", "Common issues and solutions"),
        ("7.", "Quick Reference", "Keyboard shortcuts and tips"),
    ]

    toc_table = doc.add_table(rows=len(toc_items), cols=3)
    for idx, (num, title, desc) in enumerate(toc_items):
        toc_table.cell(idx, 0).paragraphs[0].add_run(num).bold = True
        toc_table.cell(idx, 1).paragraphs[0].add_run(title).bold = True
        toc_table.cell(idx, 2).paragraphs[0].add_run(desc).font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    doc.add_page_break()

    # ========== SECTION 1: INTRODUCTION ==========
    create_styled_heading(doc, "1. Introduction", 1)
    add_horizontal_line(doc)

    create_styled_heading(doc, "What is StimuPop?", 2)

    intro_para = doc.add_paragraph()
    intro_para.add_run("StimuPop").bold = True
    intro_para.add_run(" is a powerful yet easy-to-use application that converts Excel spreadsheet data into professional PowerPoint presentations. It's designed to save you hours of manual work by automatically:")

    features = [
        "Extracting embedded images from Excel cells",
        "Creating individual slides for each data row",
        "Formatting text with custom fonts, sizes, and colors",
        "Supporting both portrait and landscape orientations",
        "Handling errors gracefully (missing images won't stop the process)"
    ]

    for feature in features:
        p = doc.add_paragraph(style='List Bullet')
        p.add_run("✓ ").font.color.rgb = RGBColor(0x28, 0xA7, 0x45)
        p.add_run(feature)

    doc.add_paragraph()
    add_tip_box(doc, "StimuPop is perfect for creating product catalogs, photo albums, training materials, and any presentation where each slide follows the same format.")

    # ========== SECTION 2: GETTING STARTED ==========
    doc.add_page_break()
    create_styled_heading(doc, "2. Getting Started", 1)
    add_horizontal_line(doc)

    create_styled_heading(doc, "Installation (Portable Version)", 2)

    doc.add_paragraph("The portable version requires no installation. Simply follow these steps:")

    install_steps = [
        ("1", "Extract the ZIP file to any folder on your computer"),
        ("2", "Open the extracted folder"),
        ("3", "Double-click 'StimuPop.bat' to launch the application"),
        ("4", "Wait for your web browser to open automatically"),
    ]
    add_step_table(doc, install_steps)

    add_info_box(doc, "📌 First Launch",
                 "The first time you run StimuPop, it may take 1-2 minutes to initialize. "
                 "Subsequent launches will be much faster.")

    create_styled_heading(doc, "System Requirements", 2)

    req_table = doc.add_table(rows=4, cols=2)
    req_table.style = 'Table Grid'
    requirements = [
        ("Operating System", "Windows 10 or later (64-bit)"),
        ("Disk Space", "~300 MB for portable version"),
        ("Memory", "4 GB RAM minimum"),
        ("Browser", "Chrome, Firefox, or Edge (latest version)"),
    ]
    for idx, (req, val) in enumerate(requirements):
        req_table.cell(idx, 0).paragraphs[0].add_run(req).bold = True
        set_cell_shading(req_table.cell(idx, 0), "F0F0F0")
        req_table.cell(idx, 1).paragraphs[0].add_run(val)

    doc.add_paragraph()

    # ========== SECTION 3: PREPARING YOUR EXCEL FILE ==========
    doc.add_page_break()
    create_styled_heading(doc, "3. Preparing Your Excel File", 1)
    add_horizontal_line(doc)

    doc.add_paragraph("For best results, structure your Excel file as follows:")

    create_styled_heading(doc, "Recommended Structure", 2)

    # Example table
    example_table = doc.add_table(rows=4, cols=4)
    example_table.style = 'Table Grid'

    headers = ["Column A\n(Skip)", "Column B\n(Image)", "Column C\n(Title)", "Column D\n(Description)"]
    for idx, header in enumerate(headers):
        cell = example_table.cell(0, idx)
        set_cell_shading(cell, "2E74B5")
        run = cell.paragraphs[0].add_run(header)
        run.bold = True
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        run.font.size = Pt(10)

    data_rows = [
        ["ID-001", "[Image 1]", "Product One", "Description of product one"],
        ["ID-002", "[Image 2]", "Product Two", "Description of product two"],
        ["ID-003", "[Image 3]", "Product Three", "Description of product three"],
    ]

    for row_idx, row_data in enumerate(data_rows, start=1):
        for col_idx, cell_data in enumerate(row_data):
            cell = example_table.cell(row_idx, col_idx)
            if row_idx % 2 == 0:
                set_cell_shading(cell, "F8F9FA")
            cell.paragraphs[0].add_run(cell_data).font.size = Pt(10)

    doc.add_paragraph()

    create_styled_heading(doc, "Image Options", 2)

    doc.add_paragraph("StimuPop supports three ways to include images:")

    img_options = [
        ("Embedded Images", "Paste images directly into Excel cells. This is the most reliable method."),
        ("File Paths", "Enter the full path to an image file (e.g., C:\\Images\\photo.jpg)"),
        ("URLs", "Enter a web URL to an image (e.g., https://example.com/image.png)"),
    ]

    for title, desc in img_options:
        p = doc.add_paragraph()
        p.add_run(f"• {title}: ").bold = True
        p.add_run(desc)

    doc.add_paragraph()
    add_warning_box(doc, "Images embedded directly in Excel cells provide the most consistent results. "
                        "File paths must be accessible from the computer running StimuPop.")

    # ========== SECTION 4: USING STIMUPOP ==========
    doc.add_page_break()
    create_styled_heading(doc, "4. Using StimuPop", 1)
    add_horizontal_line(doc)

    create_styled_heading(doc, "Step-by-Step Guide", 2)

    usage_steps = [
        ("1", "Launch StimuPop by double-clicking 'StimuPop.bat'"),
        ("2", "Click 'Browse files' under 'Upload Excel File' and select your .xlsx file"),
        ("3", "Optionally upload a PowerPoint template for custom styling"),
        ("4", "Set the Image Column (e.g., 'B' or the column header name)"),
        ("5", "Set the Text Columns (e.g., 'C,D,E,F' for multiple columns)"),
        ("6", "Adjust font size using the slider (default: 14pt)"),
        ("7", "Click the blue 'Generate Presentation' button"),
        ("8", "Wait for processing (progress bar shows status)"),
        ("9", "Click 'Download Presentation' to save your .pptx file"),
    ]
    add_step_table(doc, usage_steps)

    create_styled_heading(doc, "Understanding the Interface", 2)

    interface_desc = doc.add_paragraph()
    interface_desc.add_run("The StimuPop interface is divided into several sections:\n\n")

    sections = [
        ("Upload Files", "Where you select your Excel file and optional PowerPoint template"),
        ("Configuration", "Basic settings like image column, text columns, and font size"),
        ("Advanced Settings", "Layout options and per-column formatting (click to expand)"),
        ("Data Preview", "Shows a preview of your Excel data before generation"),
        ("Generate Button", "Starts the presentation generation process"),
    ]

    for section, desc in sections:
        p = doc.add_paragraph()
        p.add_run(f"📍 {section}: ").bold = True
        p.add_run(desc)

    # ========== SECTION 5: CONFIGURATION OPTIONS ==========
    doc.add_page_break()
    create_styled_heading(doc, "5. Configuration Options", 1)
    add_horizontal_line(doc)

    create_styled_heading(doc, "Basic Settings", 2)

    basic_settings = [
        ("Image Column", "Letter (A-Z) or name of the column containing images", "B"),
        ("Text Columns", "Comma-separated list of columns for text content", "C,D,E,F"),
        ("Font Size", "Default font size for all text (10-32pt)", "14"),
    ]

    settings_table = doc.add_table(rows=len(basic_settings)+1, cols=3)
    settings_table.style = 'Table Grid'

    # Header
    for idx, header in enumerate(["Setting", "Description", "Default"]):
        cell = settings_table.cell(0, idx)
        set_cell_shading(cell, "2E74B5")
        run = cell.paragraphs[0].add_run(header)
        run.bold = True
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    for row_idx, (setting, desc, default) in enumerate(basic_settings, start=1):
        settings_table.cell(row_idx, 0).paragraphs[0].add_run(setting).bold = True
        settings_table.cell(row_idx, 1).paragraphs[0].add_run(desc)
        settings_table.cell(row_idx, 2).paragraphs[0].add_run(default)

    doc.add_paragraph()

    create_styled_heading(doc, "Advanced Settings", 2)

    doc.add_paragraph("Click 'Advanced Settings' to access these options:")

    advanced_settings = [
        ("Image Width", "Width of images on slides (3.0-7.0 inches)", "5.5 inches"),
        ("Image Top Position", "Distance from top of slide to image", "0.5 inches"),
        ("Text Top Position", "Distance from top of slide to text area", "5.0 inches"),
        ("Slide Orientation", "Portrait (tall) or Landscape (wide)", "Portrait"),
    ]

    for setting, desc, default in advanced_settings:
        p = doc.add_paragraph()
        p.add_run(f"• {setting}: ").bold = True
        p.add_run(f"{desc} (Default: {default})")

    doc.add_paragraph()

    create_styled_heading(doc, "Per-Column Formatting", 2)

    doc.add_paragraph("In Advanced Settings, you can customize each text column individually:")

    format_options = ["Font size (8-48pt)", "Font family (Calibri, Arial, Times New Roman, etc.)",
                      "Text color (color picker)", "Bold and Italic styles"]

    for opt in format_options:
        p = doc.add_paragraph(style='List Bullet')
        p.add_run(opt)

    doc.add_paragraph()
    add_tip_box(doc, "Use larger, bold fonts for titles (Column C) and smaller regular fonts for descriptions (Column D) to create visual hierarchy.")

    # ========== SECTION 6: TROUBLESHOOTING ==========
    doc.add_page_break()
    create_styled_heading(doc, "6. Troubleshooting", 1)
    add_horizontal_line(doc)

    issues = [
        ("Application won't start",
         "• Ensure you extracted all files from the ZIP\n"
         "• Run as Administrator (right-click → Run as administrator)\n"
         "• Check that antivirus isn't blocking the application"),

        ("Browser doesn't open automatically",
         "• Manually open your browser and go to: http://localhost:8501\n"
         "• Try a different browser (Chrome recommended)"),

        ("'Column not found' error",
         "• Check that your column letter/name matches exactly\n"
         "• Column letters are case-insensitive (B = b)\n"
         "• Column names ARE case-sensitive"),

        ("Images not appearing on slides",
         "• Ensure images are embedded in cells, not floating\n"
         "• Check file paths are correct and accessible\n"
         "• Verify image files are JPG, PNG, or GIF format"),

        ("Presentation generation is slow",
         "• Large images take longer to process\n"
         "• Consider resizing images in Excel before upload\n"
         "• Reduce the number of slides if testing"),

        ("Text formatting not applied",
         "• Ensure you configured formatting in Advanced Settings\n"
         "• Check that column letters in formatting match your data"),
    ]

    for issue, solution in issues:
        p = doc.add_paragraph()
        p.add_run(f"❓ {issue}").bold = True
        p.add_run(f"\n{solution}")
        doc.add_paragraph()

    # ========== SECTION 7: QUICK REFERENCE ==========
    doc.add_page_break()
    create_styled_heading(doc, "7. Quick Reference", 1)
    add_horizontal_line(doc)

    create_styled_heading(doc, "Column Reference Cheat Sheet", 2)

    ref_table = doc.add_table(rows=6, cols=2)
    ref_table.style = 'Table Grid'

    ref_data = [
        ("Reference Type", "Example"),
        ("Single letter", "B"),
        ("Multiple letters", "C,D,E,F"),
        ("Column name", "Title"),
        ("Mixed", "B,Title,Description"),
        ("With spaces", "\"Product Name\""),
    ]

    for idx, (col1, col2) in enumerate(ref_data):
        if idx == 0:
            set_cell_shading(ref_table.cell(idx, 0), "2E74B5")
            set_cell_shading(ref_table.cell(idx, 1), "2E74B5")
            ref_table.cell(idx, 0).paragraphs[0].add_run(col1).font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            ref_table.cell(idx, 1).paragraphs[0].add_run(col2).font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            ref_table.cell(idx, 0).paragraphs[0].runs[0].bold = True
            ref_table.cell(idx, 1).paragraphs[0].runs[0].bold = True
        else:
            ref_table.cell(idx, 0).paragraphs[0].add_run(col1)
            code_run = ref_table.cell(idx, 1).paragraphs[0].add_run(col2)
            code_run.font.name = 'Consolas'

    doc.add_paragraph()

    create_styled_heading(doc, "Keyboard Shortcuts", 2)

    shortcuts = [
        ("Ctrl + C", "Stop the application (in command window)"),
        ("F5", "Refresh the browser page"),
        ("Ctrl + S", "Save downloaded file (in browser)"),
    ]

    for shortcut, action in shortcuts:
        p = doc.add_paragraph()
        shortcut_run = p.add_run(f"  {shortcut}  ")
        shortcut_run.bold = True
        shortcut_run.font.name = 'Consolas'
        p.add_run(f"  →  {action}")

    doc.add_paragraph()
    doc.add_paragraph()

    # Footer box
    add_info_box(doc, "📧 Need Help?",
                 "If you encounter issues not covered in this guide, please contact your system administrator "
                 "or the development team with:\n"
                 "• A description of the problem\n"
                 "• The exact error message (if any)\n"
                 "• Your Excel file (if possible)")

    # ========== FINAL PAGE ==========
    doc.add_page_break()

    doc.add_paragraph()
    doc.add_paragraph()
    doc.add_paragraph()

    thanks = doc.add_paragraph()
    thanks.alignment = WD_ALIGN_PARAGRAPH.CENTER
    thanks_run = thanks.add_run("Thank you for using StimuPop!")
    thanks_run.font.size = Pt(20)
    thanks_run.font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)
    thanks_run.bold = True

    doc.add_paragraph()

    version_final = doc.add_paragraph()
    version_final.alignment = WD_ALIGN_PARAGRAPH.CENTER
    version_final.add_run("Version 2.2.0").font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    # Save document
    doc.save('StimuPop_User_Guide.docx')
    print("[OK] User Guide created successfully: StimuPop_User_Guide.docx")


if __name__ == "__main__":
    create_user_guide()
