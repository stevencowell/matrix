import datetime
import zipfile
from xml.sax.saxutils import escape

XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
NAMESPACE_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NAMESPACE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def column_letter(idx: int) -> str:
    result = []
    while idx:
        idx, remainder = divmod(idx - 1, 26)
        result.append(chr(65 + remainder))
    return ''.join(reversed(result))


def build_content_types() -> str:
    parts = [
        XML_HEADER,
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '  <Default Extension="xml" ContentType="application/xml"/>',
        '  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
        '  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
        '  <Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
        '  <Override PartName="/xl/worksheets/sheet4.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
        '  <Override PartName="/xl/worksheets/sheet5.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
        '  <Override PartName="/xl/worksheets/sheet6.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
        '  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
        '  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
        '  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
        '</Types>'
    ]
    return '\n'.join(parts)


def build_root_rels() -> str:
    parts = [
        XML_HEADER,
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>',
        '  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>',
        '  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>',
        '</Relationships>'
    ]
    return '\n'.join(parts)


def build_app_xml(sheet_names) -> str:
    vector_entries = ''.join(f'<vt:lpstr>{escape(name)}</vt:lpstr>' for name in sheet_names)
    parts = [
        XML_HEADER,
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">',
        '  <Application>Microsoft Excel</Application>',
        '  <DocSecurity>0</DocSecurity>',
        '  <ScaleCrop>false</ScaleCrop>',
        '  <HeadingPairs>',
        '    <vt:vector size="2" baseType="variant">',
        '      <vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant>',
        f'      <vt:variant><vt:i4>{len(sheet_names)}</vt:i4></vt:variant>',
        '    </vt:vector>',
        '  </HeadingPairs>',
        '  <TitlesOfParts>',
        f'    <vt:vector size="{len(sheet_names)}" baseType="lpstr">{vector_entries}</vt:vector>',
        '  </TitlesOfParts>',
        '  <Company></Company>',
        '  <LinksUpToDate>false</LinksUpToDate>',
        '  <SharedDoc>false</SharedDoc>',
        '  <HyperlinksChanged>false</HyperlinksChanged>',
        '  <AppVersion>16.0300</AppVersion>',
        '</Properties>'
    ]
    return '\n'.join(parts)


def build_core_xml() -> str:
    timestamp = (
        datetime.datetime.now(datetime.timezone.utc)
        .replace(microsecond=0)
        .isoformat()
        .replace('+00:00', 'Z')
    )
    parts = [
        XML_HEADER,
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">',
        '  <dc:creator>Timetable Matrix Template Builder</dc:creator>',
        '  <cp:lastModifiedBy>Timetable Matrix Template Builder</cp:lastModifiedBy>',
        f'  <dcterms:created xsi:type="dcterms:W3CDTF">{timestamp}</dcterms:created>',
        f'  <dcterms:modified xsi:type="dcterms:W3CDTF">{timestamp}</dcterms:modified>',
        '</cp:coreProperties>'
    ]
    return '\n'.join(parts)


def build_styles_xml() -> str:
    parts = [
        XML_HEADER,
        f'<styleSheet xmlns="{NAMESPACE_MAIN}">',
        '  <fonts count="4">',
        '    <font><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/><family val="2"/></font>',
        '    <font><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/><b/></font>',
        '    <font><sz val="14"/><color rgb="FF1F4E78"/><name val="Calibri"/><family val="2"/><b/></font>',
        '    <font><sz val="12"/><color rgb="FF1F4E78"/><name val="Calibri"/><family val="2"/><b/></font>',
        '  </fonts>',
        '  <fills count="3">',
        '    <fill><patternFill patternType="none"/></fill>',
        '    <fill><patternFill patternType="gray125"/></fill>',
        '    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor rgb="FF1F4E78"/></patternFill></fill>',
        '  </fills>',
        '  <borders count="2">',
        '    <border><left/><right/><top/><bottom/><diagonal/></border>',
        '    <border>',
        '      <left style="thin"><color rgb="FF4F81BD"/></left>',
        '      <right style="thin"><color rgb="FF4F81BD"/></right>',
        '      <top style="thin"><color rgb="FF4F81BD"/></top>',
        '      <bottom style="thin"><color rgb="FF4F81BD"/></bottom>',
        '      <diagonal/>',
        '    </border>',
        '  </borders>',
        '  <cellStyleXfs count="1">',
        '    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>',
        '  </cellStyleXfs>',
        '  <cellXfs count="6">',
        '    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
        '    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>',
        '    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>',
        '    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>',
        '    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>',
        '    <xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>',
        '  </cellXfs>',
        '  <cellStyles count="1">',
        '    <cellStyle name="Normal" xfId="0" builtinId="0"/>',
        '  </cellStyles>',
        '</styleSheet>'
    ]
    return '\n'.join(parts)


def build_workbook_xml(sheet_names) -> str:
    sheet_entries = [f'    <sheet name="{escape(name)}" sheetId="{idx}" r:id="rId{idx}"/>' for idx, name in enumerate(sheet_names, start=1)]
    parts = [
        XML_HEADER,
        f'<workbook xmlns="{NAMESPACE_MAIN}" xmlns:r="{NAMESPACE_REL}">',
        '  <fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="22228"/>',
        '  <workbookPr defaultThemeVersion="164011"/>',
        '  <bookViews>',
        '    <workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="16380" activeTab="0"/>',
        '  </bookViews>',
        '  <sheets>',
        *sheet_entries,
        '  </sheets>',
        '</workbook>'
    ]
    return '\n'.join(parts)


def build_workbook_rels(sheet_count: int) -> str:
    rel_entries = [
        f'  <Relationship Id="rId{idx}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{idx}.xml"/>'
        for idx in range(1, sheet_count + 1)
    ]
    rel_entries.append(
        f'  <Relationship Id="rId{sheet_count + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    )
    parts = [
        XML_HEADER,
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        *rel_entries,
        '</Relationships>'
    ]
    return '\n'.join(parts)


def build_instructions_sheet() -> str:
    lines = [
        ("Head Teacher Allocation Template", 4),
        ("How to use this spreadsheet:", 5),
        ("1. Start in the Staff sheet to list all available teachers, their faculty, employment fraction and teaching limits.", 3),
        ("2. Use the Subjects sheet to capture each class that needs staffing, including student numbers and any special requirements.", 3),
        ("3. Record your timetable structure in the Lines sheet if your school uses lines or blocks.", 3),
        ("4. Build the teaching allocation in the Allocations sheet using the drop-down menus for subjects and teachers.", 3),
        ("5. Check the Coverage Summary to confirm every class has a teacher and loads stay within limits.", 3),
        ("Blue headings indicate areas to complete. Cells with drop-down menus pull their options from other sheets.", 3),
    ]
    dimension = f"A1:A{len(lines)}"
    rows_xml = [
        f'  <row r="{idx}" spans="1:1"><c r="A{idx}" s="{style}" t="inlineStr"><is><t xml:space="preserve">{escape(text)}</t></is></c></row>'
        for idx, (text, style) in enumerate(lines, start=1)
    ]
    parts = [
        XML_HEADER,
        f'<worksheet xmlns="{NAMESPACE_MAIN}" xmlns:r="{NAMESPACE_REL}">',
        f'  <dimension ref="{dimension}"/>',
        '  <sheetViews>',
        '    <sheetView workbookViewId="0" showGridLines="0">',
        '      <selection activeCell="A1" sqref="A1"/>',
        '    </sheetView>',
        '  </sheetViews>',
        '  <sheetFormatPr defaultRowHeight="15"/>',
        '  <cols><col min="1" max="1" width="120" customWidth="1"/></cols>',
        '  <sheetData>',
        *rows_xml,
        '  </sheetData>',
        '  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>',
        '</worksheet>'
    ]
    return '\n'.join(parts)


def build_data_sheet(headers, column_widths, left_align_columns, total_data_rows, data_validations=None) -> str:
    column_count = len(headers)
    total_rows = total_data_rows + 1
    last_column = column_letter(column_count)
    dimension = f"A1:{last_column}{total_rows}"
    header_cells = [
        f'<c r="{column_letter(col_idx)}1" s="1" t="inlineStr"><is><t xml:space="preserve">{escape(header)}</t></is></c>'
        for col_idx, header in enumerate(headers, start=1)
    ]
    data_rows = []
    left_align = set(left_align_columns)
    for row_idx in range(2, total_rows + 1):
        cells = []
        for col_idx in range(1, column_count + 1):
            style = 3 if col_idx in left_align else 2
            cell_ref = f"{column_letter(col_idx)}{row_idx}"
            cells.append(f'<c r="{cell_ref}" s="{style}"/>')
        data_rows.append(f'  <row r="{row_idx}" spans="1:{column_count}">{''.join(cells)}</row>')
    cols_xml = ''.join(
        f'<col min="{idx}" max="{idx}" width="{width}" customWidth="1"/>'
        for idx, width in enumerate(column_widths, start=1)
    )
    parts = [
        XML_HEADER,
        f'<worksheet xmlns="{NAMESPACE_MAIN}" xmlns:r="{NAMESPACE_REL}">',
        f'  <dimension ref="{dimension}"/>',
        '  <sheetViews>',
        '    <sheetView workbookViewId="0">',
        '      <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>',
        '      <selection pane="bottomLeft" activeCell="A2" sqref="A2"/>',
        '    </sheetView>',
        '  </sheetViews>',
        '  <sheetFormatPr defaultRowHeight="15"/>',
        f'  <cols>{cols_xml}</cols>',
        '  <sheetData>',
        f'  <row r="1" spans="1:{column_count}" ht="18" customHeight="1">{''.join(header_cells)}</row>',
        *data_rows,
        '  </sheetData>',
    ]
    if data_validations:
        parts.append(f'  <dataValidations count="{len(data_validations)}">')
        for validation in data_validations:
            allow_blank = '1' if validation.get('allow_blank', True) else '0'
            parts.append(
                f'    <dataValidation type="{validation["type"]}" allowBlank="{allow_blank}" showInputMessage="1" showErrorMessage="1" sqref="{validation["sqref"]}">'  # noqa: E501
            )
            parts.append(f'      <formula1>{escape(validation["formula"])}</formula1>')
            parts.append('    </dataValidation>')
        parts.append('  </dataValidations>')
    parts.append('  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>')
    parts.append('</worksheet>')
    return '\n'.join(parts)


def build_summary_sheet() -> str:
    column_widths = [18, 22, 20]
    cols_xml = ''.join(
        f'<col min="{idx}" max="{idx}" width="{width}" customWidth="1"/>'
        for idx, width in enumerate(column_widths, start=1)
    )
    rows = []
    rows.append('  <row r="1" spans="1:3"><c r="A1" s="4" t="inlineStr"><is><t xml:space="preserve">Class Coverage</t></is></c></row>')
    header_row = ''.join(
        f'<c r="{column_letter(col)}2" s="1" t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'
        for col, text in enumerate(["Subject Code", "Class Identifier", "Periods Assigned"], start=1)
    )
    rows.append(f'  <row r="2" spans="1:3">{header_row}</row>')
    blank_row = ''.join(f'<c r="{column_letter(col)}3" s="2"/>' for col in range(1, 4))
    rows.append(f'  <row r="3" spans="1:3">{blank_row}</row>')
    rows.append('  <row r="4" spans="1:3"><c r="A4" s="4" t="inlineStr"><is><t xml:space="preserve">Staff Load</t></is></c></row>')
    staff_header = ''.join(
        f'<c r="{column_letter(col)}5" s="1" t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'
        for col, text in enumerate(["Staff Name", "Allocated Periods", "Remaining Capacity"], start=1)
    )
    rows.append(f'  <row r="5" spans="1:3">{staff_header}</row>')
    for row_idx in range(6, 21):
        cells = []
        for col_idx in range(1, 4):
            style = 3 if col_idx == 1 else 2
            cell_ref = f"{column_letter(col_idx)}{row_idx}"
            cells.append(f'<c r="{cell_ref}" s="{style}"/>')
        rows.append(f'  <row r="{row_idx}" spans="1:3">{''.join(cells)}</row>')
    parts = [
        XML_HEADER,
        f'<worksheet xmlns="{NAMESPACE_MAIN}" xmlns:r="{NAMESPACE_REL}">',
        '  <dimension ref="A1:C20"/>',
        '  <sheetViews>',
        '    <sheetView workbookViewId="0" showGridLines="0">',
        '      <selection activeCell="A1" sqref="A1"/>',
        '    </sheetView>',
        '  </sheetViews>',
        '  <sheetFormatPr defaultRowHeight="15"/>',
        f'  <cols>{cols_xml}</cols>',
        '  <sheetData>',
        *rows,
        '  </sheetData>',
        '  <mergeCells count="2">',
        '    <mergeCell ref="A1:C1"/>',
        '    <mergeCell ref="A4:C4"/>',
        '  </mergeCells>',
        '  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>',
        '</worksheet>'
    ]
    return '\n'.join(parts)


def create_workbook(filename: str) -> None:
    sheet_names = [
        "Instructions",
        "Staff",
        "Subjects",
        "Lines",
        "Allocations",
        "Coverage Summary",
    ]
    staff_headers = [
        "Staff Code",
        "Staff Name",
        "Faculty/Department",
        "Role",
        "Employment Fraction (FTE)",
        "Max Teaching Periods",
        "Yard/Extra Duties",
        "Comments",
    ]
    subjects_headers = [
        "Subject Code",
        "Subject Name",
        "Year Level",
        "Line/Block",
        "Class Identifier",
        "Planned Class Size",
        "Room Type/Resources",
        "Special Notes",
    ]
    lines_headers = [
        "Year Level",
        "Line/Block",
        "Description",
        "Default Room",
        "Notes",
    ]
    allocations_headers = [
        "Year Level",
        "Line",
        "Subject Code",
        "Subject Name",
        "Class Identifier",
        "Periods/Cycle",
        "Teacher 1",
        "Teacher 1 Load",
        "Teacher 2",
        "Teacher 2 Load",
        "Room",
        "Shared Notes",
    ]

    staff_sheet = build_data_sheet(
        headers=staff_headers,
        column_widths=[15, 26, 20, 18, 24, 20, 20, 35],
        left_align_columns={2, 3, 8},
        total_data_rows=50,
    )
    subjects_sheet = build_data_sheet(
        headers=subjects_headers,
        column_widths=[18, 28, 14, 12, 18, 20, 24, 32],
        left_align_columns={2, 7, 8},
        total_data_rows=100,
    )
    lines_sheet = build_data_sheet(
        headers=lines_headers,
        column_widths=[14, 14, 36, 18, 30],
        left_align_columns={3, 5},
        total_data_rows=60,
    )
    allocations_sheet = build_data_sheet(
        headers=allocations_headers,
        column_widths=[12, 12, 16, 26, 18, 16, 20, 16, 20, 16, 14, 36],
        left_align_columns={4, 12},
        total_data_rows=100,
        data_validations=[
            {
                "type": "list",
                "sqref": "C2:C101",
                "formula": "Subjects!$A$2:$A$101",
            },
            {
                "type": "list",
                "sqref": "D2:D101",
                "formula": "Subjects!$B$2:$B$101",
            },
            {
                "type": "list",
                "sqref": "G2:G101 I2:I101",
                "formula": "Staff!$A$2:$A$51",
            },
        ],
    )
    summary_sheet = build_summary_sheet()

    with zipfile.ZipFile(filename, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", build_content_types())
        zf.writestr("_rels/.rels", build_root_rels())
        zf.writestr("docProps/app.xml", build_app_xml(sheet_names))
        zf.writestr("docProps/core.xml", build_core_xml())
        zf.writestr("xl/workbook.xml", build_workbook_xml(sheet_names))
        zf.writestr("xl/_rels/workbook.xml.rels", build_workbook_rels(len(sheet_names)))
        zf.writestr("xl/styles.xml", build_styles_xml())
        zf.writestr("xl/worksheets/sheet1.xml", build_instructions_sheet())
        zf.writestr("xl/worksheets/sheet2.xml", staff_sheet)
        zf.writestr("xl/worksheets/sheet3.xml", subjects_sheet)
        zf.writestr("xl/worksheets/sheet4.xml", lines_sheet)
        zf.writestr("xl/worksheets/sheet5.xml", allocations_sheet)
        zf.writestr("xl/worksheets/sheet6.xml", summary_sheet)


if __name__ == "__main__":
    create_workbook("HeadTeacher_Allocation_Template.xlsx")
