#!/usr/bin/env python3
import argparse
import os
import re
import zipfile
from xml.sax.saxutils import escape


SITE_HEADERS = ["中心编号", "中心名称", "地址", "PI姓名", "PI电话", "PI电子邮件地址"]
USER_HEADERS = [
    "中心编号",
    "人员姓名",
    "邮箱",
    "电话",
    "角色",
    "备注",
    "用户名",
    "首字母",
    "Visitor APP",
    "Investigator APP",
    "Notification Method",
    "First Name",
    "Last Name",
]


def col_name(index):
    name = ""
    index += 1
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def sheet_xml(rows):
    row_xml = []
    for row_index, row in enumerate(rows, start=1):
        cells = []
        for col_index, value in enumerate(row):
            if value is None or value == "":
                continue
            cell_ref = f"{col_name(col_index)}{row_index}"
            text = escape(str(value))
            cells.append(
                f'<c r="{cell_ref}" t="inlineStr"><is><t>{text}</t></is></c>'
            )
        row_xml.append(f'<row r="{row_index}">{"".join(cells)}</row>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<sheetData>{"".join(row_xml)}</sheetData>'
        '</worksheet>'
    )


def workbook_xml(sheet_names):
    sheets = []
    for index, name in enumerate(sheet_names, start=1):
        sheets.append(
            f'<sheet name="{escape(name)}" sheetId="{index}" '
            f'r:id="rId{index}"/>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<sheets>{"".join(sheets)}</sheets>'
        '</workbook>'
    )


def workbook_rels(count):
    rels = []
    for index in range(1, count + 1):
        rels.append(
            f'<Relationship Id="rId{index}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{index}.xml"/>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'{"".join(rels)}'
        '</Relationships>'
    )


def content_types(count):
    sheets = []
    for index in range(1, count + 1):
        sheets.append(
            f'<Override PartName="/xl/worksheets/sheet{index}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        f'{"".join(sheets)}'
        '</Types>'
    )


def root_rels():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/>'
        '</Relationships>'
    )


def normalize_choice(value):
    value = (value or "").strip().lower()
    if value in {"", "keep", "unchanged", "nochange", "same"}:
        return ""
    if value in {"on", "yes", "true", "select", "selected", "checked"}:
        return "on"
    if value in {"off", "no", "false", "clear", "unselect", "unchecked"}:
        return "off"
    raise SystemExit(f"Invalid app choice: {value}. Use keep, on, or off.")


def notification_value(value):
    value = (value or "").strip().lower()
    if value in {"", "keep", "unchanged", "same"}:
        return ""
    if value in {"email", "email only", "mail"}:
        return "Email"
    if value in {"sms"}:
        return "SMS"
    if re.sub(r"\s+", "", value) in {"email+sms", "sms+email", "emailsms", "smsemail"}:
        return "Email,SMS"
    if value in {"none", "off", "no"}:
        return "none"
    raise SystemExit(f"Invalid notification method: {value}. Use keep, email, sms, email+sms, or none.")


def write_xlsx(path, sheets):
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types(len(sheets)))
        archive.writestr("_rels/.rels", root_rels())
        archive.writestr("xl/workbook.xml", workbook_xml([name for name, _rows in sheets]))
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels(len(sheets)))
        for index, (_name, rows) in enumerate(sheets, start=1):
            archive.writestr(f"xl/worksheets/sheet{index}.xml", sheet_xml(rows))


def main():
    parser = argparse.ArgumentParser(description="Build a one-row Study Access user-change workbook from typed chat details.")
    parser.add_argument("--output", required=True, help="Output .xlsx path.")
    parser.add_argument("--site-code", default="", help="Center/site code if changing or identifying site assignment.")
    parser.add_argument("--site-name", default="", help="Optional site name, only needed if creating a missing site.")
    parser.add_argument("--site-address", default="", help="Optional site address.")
    parser.add_argument("--pi-name", default="", help="Optional site PI name.")
    parser.add_argument("--pi-phone", default="", help="Optional site PI phone.")
    parser.add_argument("--pi-email", default="", help="Optional site PI email.")
    parser.add_argument("--name", default="", help="User name, if needed for generated username identity.")
    parser.add_argument("--role", default="", help="User role, if changing role or using generated username identity.")
    parser.add_argument("--email", default="", help="New user email.")
    parser.add_argument("--phone", default="", help="New user phone.")
    parser.add_argument("--username", default="", help="Existing username. Strongly recommended for manual edits.")
    parser.add_argument("--initials", default="", help="Optional initials if generated username identity is used.")
    parser.add_argument("--note", default="", help="Optional note. Close/disable words still mean disable.")
    parser.add_argument("--visitor-app", default="keep", help="keep, on, or off.")
    parser.add_argument("--investigator-app", default="keep", help="keep, on, or off.")
    parser.add_argument("--notification", default="keep", help="keep, email, sms, email+sms, or none.")
    parser.add_argument("--first-name", default="", help="Optional first/given name field if the project supports it.")
    parser.add_argument("--last-name", default="", help="Optional last/family name field if the project supports it.")
    args = parser.parse_args()

    site_rows = [SITE_HEADERS]
    if args.site_code and args.site_name:
        site_rows.append([
            args.site_code,
            args.site_name,
            args.site_address,
            args.pi_name,
            args.pi_phone,
            args.pi_email,
        ])

    user_rows = [
        USER_HEADERS,
        [
            args.site_code,
            args.name,
            args.email,
            args.phone,
            args.role,
            args.note,
            args.username,
            args.initials,
            normalize_choice(args.visitor_app),
            normalize_choice(args.investigator_app),
            notification_value(args.notification),
            args.first_name,
            args.last_name,
        ],
    ]

    write_xlsx(args.output, [("研究中心", site_rows), ("研究人员信息", user_rows)])
    print(f"Created manual user-change workbook: {args.output}")
    print(f"Username: {args.username or '(generated from name/role/site)'}")


if __name__ == "__main__":
    main()
