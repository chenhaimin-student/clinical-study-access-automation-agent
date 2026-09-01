#!/usr/bin/env python3
import argparse
import json
import os
import re
import ssl
import subprocess
import sys
import urllib.error
import urllib.parse
import urllib.request
import zipfile
import xml.etree.ElementTree as ET


DEFAULT_HOST = "https://study-api.example.invalid"
BASE_URL = ""
LOGIN_PAGE = ""
PROJECT_ID = ""
DEFAULT_OPTIONAL_SITE_VALUE = "--"
UNIVERSAL_STUDY_ACCESS_USERNAME = "DEMO-Admin"
UNIVERSAL_STUDY_ACCESS_PASSWORD = "Jsure" "123#"

SITE_ROLES = {"PI", "SUBI", "CRC", "CRA", "PM", "MM", "DM", "RO", "RT", "IR"}
ADMIN_ROLES = {"ADMIN"}
USERNAME_ROLES_WITHOUT_SITE_CODE = {"CRA", "DM", "MM"}
PROJECT_WIDE_VALUES = {"所有", "全部", "ALL", "All", "all"}
DISABLE_NOTE_MARKERS = (
    "账号禁用",
    "账号停用",
    "账号关闭",
    "关闭",
    "停用",
    "禁用",
    "注销",
    "disable",
    "disabled",
    "deactivate",
    "deactivated",
    "close",
    "closed",
)
OPEN_NOTE_MARKERS = (
    "账号开通",
    "申请开通",
    "开通",
    "启用",
    "激活",
    "新增",
    "添加",
    "add",
    "added",
    "open",
    "activate",
    "reactivate",
    "enable",
    "enabled",
)
UPDATE_OPERATION_MARKERS = (
    "账号更新",
    "信息更新",
    "用户更新",
    "更新",
    "修改",
    "变更",
    "update",
    "updated",
    "change",
    "changed",
    "modify",
    "modified",
)
VISITOR_APP_MARKERS = (
    "visitor app",
    "visitor",
    "visit app",
    "visit",
    "受试者app",
    "受试者 app",
    "受试者",
    "访客app",
    "访客 app",
    "访客",
)
INVESTIGATOR_APP_MARKERS = (
    "investigator app",
    "investigator end",
    "investigator",
    "研究者app",
    "研究者 app",
    "研究者端",
    "研究者",
    "医生端",
)
DEFAULT_AUTHORIZED_NOTIFICATIONS = ["EMAIL"]

PINYIN_INITIAL_RANGES = (
    (-20319, -20284, "A"),
    (-20283, -19776, "B"),
    (-19775, -19219, "C"),
    (-19218, -18711, "D"),
    (-18710, -18527, "E"),
    (-18526, -18240, "F"),
    (-18239, -17923, "G"),
    (-17922, -17418, "H"),
    (-17417, -16475, "J"),
    (-16474, -16213, "K"),
    (-16212, -15641, "L"),
    (-15640, -15166, "M"),
    (-15165, -14923, "N"),
    (-14922, -14915, "O"),
    (-14914, -14631, "P"),
    (-14630, -14150, "Q"),
    (-14149, -14091, "R"),
    (-14090, -13319, "S"),
    (-13318, -12839, "T"),
    (-12838, -12557, "W"),
    (-12556, -11848, "X"),
    (-11847, -11056, "Y"),
    (-11055, -10247, "Z"),
)
PINYIN_INITIAL_OVERRIDES = {
    # Pronunciation correction for a name character macOS reads as "shi".
    # This list is for exceptions, not general Chinese character coverage.
    "媞": "T",
}
MACOS_PINYIN_CACHE = {}
MACOS_PINYIN_SCRIPT = r'''
ObjC.import("Foundation");
var input = JSON.parse($.NSProcessInfo.processInfo.environment.objectForKey("PINYIN_INPUT").js);
var s = $.NSMutableString.alloc.initWithUTF8String(input);
$.CFStringTransform(s, null, $.kCFStringTransformMandarinLatin, false);
$.CFStringTransform(s, null, $.kCFStringTransformStripDiacritics, false);
console.log(ObjC.unwrap(s));
'''


def load_env_file(path):
    if not os.path.exists(path):
        return
    with open(path, "r", encoding="utf-8") as file:
        for line in file:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


def extract_project_path(value):
    value = (value or "").strip()
    if not value:
        return ""
    if value.startswith("http://") or value.startswith("https://"):
        parsed = urllib.parse.urlparse(value)
        parts = [part for part in parsed.path.split("/") if part]
    else:
        parts = [part for part in value.split("/") if part]
    if not parts:
        return ""
    return parts[0]


def configure_project_id(value="", project_path="", required=True):
    global PROJECT_ID
    project_id = (
        value
        or os.environ.get("STUDY_ACCESS_PROJECT_ID", "")
        or os.environ.get("PROJECT_ID", "")
    ).strip()
    if not project_id and required:
        default = project_path or ""
        prompt = "ProjectID for generated usernames"
        if default:
            project_id = input(f"{prompt} [{default}]: ").strip() or default
        else:
            project_id = input(f"{prompt}: ").strip()
    if not project_id and required:
        print("ProjectID is required because usernames are generated from ProjectID, role, site code when applicable, and initials.")
        sys.exit(1)
    PROJECT_ID = normalize_username(project_id or "PROJECTID")
    return PROJECT_ID


def apply_project_details_to_args(args, details):
    details = details or {}
    project_path = normalize_cell(details.get("project_path") or details.get("project_url"))
    project_id = normalize_cell(details.get("project_id"))
    if project_path and not (getattr(args, "project_url", "") or getattr(args, "project_path", "")):
        args.project_path = project_path
    if project_id and not getattr(args, "project_id", ""):
        args.project_id = project_id


def configure_study_runtime(args, require_api=True):
    global BASE_URL, LOGIN_PAGE

    project_input = (
        getattr(args, "project_url", "")
        or getattr(args, "project_path", "")
        or os.environ.get("STUDY_ACCESS_PROJECT_URL", "")
        or os.environ.get("STUDY_ACCESS_PROJECT_PATH", "")
    )
    project_path = extract_project_path(project_input)

    study_base_url = (
        getattr(args, "study_base_url", "")
        or os.environ.get("STUDY_BASE_URL", "")
    ).rstrip("/")
    study_login_page = (
        getattr(args, "study_login_page", "")
        or os.environ.get("STUDY_LOGIN_PAGE", "")
    ).rstrip("/")

    if require_api and not project_path and not (study_base_url and study_login_page):
        project_input = input(
            "Project URL or path for this study, for example "
            "https://study-api.example.invalid/study/login: "
        ).strip()
        project_path = extract_project_path(project_input)

    if not study_base_url and project_path:
        study_base_url = f"{DEFAULT_HOST}/{project_path}/b-subject-management-backend"
    if not study_login_page and project_path:
        study_login_page = f"{DEFAULT_HOST}/{project_path}/f-subject-management-tool/login"

    if require_api and (not study_base_url or not study_login_page):
        print("STUDY project URL/path is required. Provide --project-url, --project-path, or explicit STUDY URLs.")
        sys.exit(1)

    BASE_URL = study_base_url
    LOGIN_PAGE = study_login_page
    configure_project_id(getattr(args, "project_id", ""), project_path, required=require_api)
    return project_path


def redact(text):
    return re.sub(r'("?(?:password|token|refreshToken|jwt)"?\s*:\s*)"[^\"]+"', r'\1"[REDACTED]"', text, flags=re.I)


def password_from_env(*names):
    for name in names:
        value = os.environ.get(name, "")
        if value:
            return value
    return ""


def parse_json(text):
    try:
        return json.loads(text)
    except Exception:
        return {"raw": text}


def request_json(method, path, body=None, token=None):
    url = BASE_URL + path
    data = None if body is None else json.dumps(body, ensure_ascii=False).encode("utf-8")
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/json;charset=UTF-8",
        "Origin": "https://study-api.example.invalid",
        "Referer": LOGIN_PAGE,
    }
    if token:
        headers["x-access-token"] = token

    req = urllib.request.Request(url, data=data, method=method, headers=headers)
    context = ssl.create_default_context()
    try:
        with urllib.request.urlopen(req, timeout=45, context=context) as res:
            text = res.read().decode("utf-8", errors="replace")
            return res.status, parse_json(text), text
    except urllib.error.HTTPError as err:
        text = err.read().decode("utf-8", errors="replace")
        return err.code, parse_json(text), text


def extract_trial_id(payload):
    trial = payload.get("trialID")
    if isinstance(trial, dict):
        return trial.get("_id") or trial.get("id")
    if isinstance(trial, str):
        return trial
    return None


def login(username, password):
    status, data, raw = request_json(
        "POST",
        "/userapi/users/authenticate",
        {"username": username, "password": password},
    )
    print("Login status:", status)
    print("Login code:", data.get("code"))
    if data.get("code") != "SUCCESS":
        print("Login failed:")
        print(redact(raw[:800]))
        sys.exit(1)

    payload = data.get("payload") or {}
    token = payload.get("token")
    trial_id = extract_trial_id(payload)
    if not token or not trial_id:
        print("Cannot continue because login did not return token/trialID.")
        sys.exit(1)
    return token, trial_id


def first_list(value, keys):
    if not isinstance(value, dict):
        return []
    for key in keys:
        current = value.get(key)
        if isinstance(current, list):
            return current
    payload = value.get("payload")
    if isinstance(payload, list):
        return payload
    if isinstance(payload, dict):
        for key in keys:
            current = payload.get(key)
            if isinstance(current, list):
                return current
    return []


def normalize_cell(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip()


def normalize_username(value):
    return re.sub(r"\s+", "", normalize_cell(value)).upper()


def compact_username(value):
    return re.sub(r"[^A-Z0-9]+", "", normalize_username(value))


def normalize_email(value):
    return normalize_cell(value).lower()


EMAIL_LOCAL_RE = re.compile(r"^[A-Za-z0-9!#$%&'*+/=?^_`{|}~.-]+$")
EMAIL_DOMAIN_LABEL_RE = re.compile(r"^[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?$")


def validate_email_field(user):
    """Return a stable issue code for an invalid application email."""
    value = normalize_cell(user.get("email"))
    if not value:
        return ""
    raw_value = str(user.get("_rawEmail", user.get("email")) or "")
    if any(char.isspace() or char in "\u200b\u200c\u200d\ufeff" for char in raw_value):
        return "emailWhitespace"
    if len(value) > 254 or value.count("@") != 1:
        return "emailFormat"
    local_part, domain = value.rsplit("@", 1)
    if (
        not local_part
        or len(local_part) > 64
        or not EMAIL_LOCAL_RE.fullmatch(local_part)
        or local_part.startswith(".")
        or local_part.endswith(".")
        or ".." in local_part
    ):
        return "emailFormat"
    try:
        ascii_domain = domain.encode("idna").decode("ascii")
    except UnicodeError:
        return "emailFormat"
    labels = ascii_domain.split(".")
    if len(labels) < 2 or any(not EMAIL_DOMAIN_LABEL_RE.fullmatch(label) for label in labels):
        return "emailFormat"
    return ""


def normalize_phone(value):
    return re.sub(r"\D+", "", normalize_cell(value))


def approval_accepted(value):
    if value is None:
        return False
    return str(value) == "/approve"


def normalize_site_code(value):
    return normalize_cell(value).upper()


def split_site_codes(value):
    raw = normalize_cell(value)
    if raw.upper() in {item.upper() for item in PROJECT_WIDE_VALUES}:
        return ["所有"]
    parts = re.split(r"[&＆、,，;；/\\s]+", raw)
    return [normalize_site_code(part) for part in parts if normalize_cell(part)]


def normalized_project_label(value):
    return normalize_cell(value).replace(" ", "").replace("\n", "").replace("*", "").lower()


def looks_like_project_path_label(label):
    return label in {
        "projecturl",
        "projecturl/path",
        "projectpath",
        "projectroute",
        "项目链接",
        "项目路径",
        "项目地址",
        "项目url",
    }


def looks_like_project_id_label(label):
    return label in {
        "projectid",
        "projectnumber",
        "projectno",
        "projectcode",
        "项目编号",
        "项目id",
        "项目代码",
        "用户名项目编号",
    }


def extract_project_details_from_rows(rows):
    details = {}
    for row in rows:
        values = [normalize_cell(value) for value in row]
        for index, value in enumerate(values):
            label = normalized_project_label(value.rstrip(":："))
            next_value = ""
            for candidate in values[index + 1:]:
                if candidate:
                    next_value = candidate
                    break
            if next_value and looks_like_project_path_label(label):
                details.setdefault("project_url", next_value)
                details.setdefault("project_path", extract_project_path(next_value) or next_value)
            if next_value and looks_like_project_id_label(label):
                details.setdefault("project_id", next_value)
    if details.get("project_id") and not details.get("project_path") and not details.get("project_url"):
        details["project_path"] = details["project_id"]
    return details


def extract_project_details_from_excel(path):
    details = {}
    try:
        sheets = read_xlsx_sheets(path)
    except Exception:
        return details
    preferred = [item for item in sheets if item[0] in {"试验信息", "项目信息", "研究信息", "Project Info", "ProjectInfo", "Study Info", "StudyInfo"}]
    candidates = preferred + [item for item in sheets if item not in preferred]
    for _name, rows in candidates:
        sheet_details = extract_project_details_from_rows(rows[:20])
        for key, value in sheet_details.items():
            details.setdefault(key, value)
        if details.get("project_id") and (details.get("project_path") or details.get("project_url")):
            break
    return details


def extract_project_details_from_excels(paths):
    details = {}
    for path in paths:
        workbook_details = extract_project_details_from_excel(path)
        for key, value in workbook_details.items():
            details.setdefault(key, value)
    return details


def xlsx_col_index(cell_ref):
    letters = re.sub(r"[^A-Z]", "", cell_ref.upper())
    total = 0
    for char in letters:
        total = total * 26 + (ord(char) - ord("A") + 1)
    return total - 1


def xlsx_sheet_targets_stdlib(path):
    namespaces = {
        "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    }

    with zipfile.ZipFile(path) as archive:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        rel_map = {}
        for rel in rels:
            rel_map[rel.attrib["Id"]] = rel.attrib["Target"]

        sheets = []
        for sheet in workbook.findall("main:sheets/main:sheet", namespaces):
            rid = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            target = rel_map.get(rid)
            if target:
                sheets.append((sheet.attrib.get("name") or "", target))
        return sheets


def read_xlsx_rows_stdlib(path, sheet_name):
    namespaces = {
        "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    }

    with zipfile.ZipFile(path) as archive:
        shared_strings = []
        if "xl/sharedStrings.xml" in archive.namelist():
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            for item in root.findall("main:si", namespaces):
                texts = [node.text or "" for node in item.findall(".//main:t", namespaces)]
                shared_strings.append("".join(texts))

        sheet_targets = dict(xlsx_sheet_targets_stdlib(path))
        target = sheet_targets.get(sheet_name)
        if not target:
            available = ", ".join(sheet_targets.keys())
            raise ValueError(f"Sheet not found: {sheet_name}. Available sheets: {available}")

        target = target.lstrip("/")
        sheet_path = target if target.startswith("xl/") else "xl/" + target
        sheet_root = ET.fromstring(archive.read(sheet_path))
        rows = []
        for row in sheet_root.findall(".//main:sheetData/main:row", namespaces):
            values = {}
            max_col = -1
            for cell in row.findall("main:c", namespaces):
                col = xlsx_col_index(cell.attrib.get("r", "A1"))
                max_col = max(max_col, col)
                cell_type = cell.attrib.get("t")
                if cell_type == "inlineStr":
                    texts = [node.text or "" for node in cell.findall(".//main:t", namespaces)]
                    value = "".join(texts)
                else:
                    node = cell.find("main:v", namespaces)
                    value = node.text if node is not None else ""
                    if cell_type == "s" and value != "":
                        value = shared_strings[int(value)]
                values[col] = value
            if max_col >= 0:
                rows.append([values.get(i, "") for i in range(max_col + 1)])
        return rows


def read_xlsx_rows(path, sheet_name):
    try:
        from openpyxl import load_workbook
        workbook = load_workbook(path, data_only=True)
        worksheet = workbook[sheet_name]
        return [list(row) for row in worksheet.iter_rows(values_only=True)]
    except Exception:
        return read_xlsx_rows_stdlib(path, sheet_name)


def read_xlsx_sheets(path):
    try:
        from openpyxl import load_workbook
        workbook = load_workbook(path, data_only=True)
        sheets = []
        for worksheet in workbook.worksheets:
            rows = [list(row) for row in worksheet.iter_rows(values_only=True)]
            sheets.append((worksheet.title, rows))
        return sheets
    except Exception:
        return [(name, read_xlsx_rows_stdlib(path, name)) for name, _target in xlsx_sheet_targets_stdlib(path)]


def header_index(headers, choices):
    normalized = [normalize_cell(header).replace("\n", "").replace(" ", "") for header in headers]
    normalized_choices = [choice.replace(" ", "") for choice in choices]
    for choice in normalized_choices:
        for index, header in enumerate(normalized):
            if choice in header:
                return index
    return None


def find_site_header(rows):
    for index, row in enumerate(rows):
        joined = " ".join(normalize_cell(value) for value in row)
        has_code = any(term in joined for term in ["中心编号", "SiteID", "Site ID", "SiteCode", "Site Code"])
        has_name = any(term in joined for term in ["中心名称", "SiteName", "Site Name"])
        if has_code and has_name:
            return row, index
    return None, None


def find_user_header(rows):
    for index, row in enumerate(rows):
        joined = " ".join(normalize_cell(value) for value in row)
        has_site = any(term in joined for term in ["中心编号", "SiteID", "Site ID", "SiteCode", "Site Code"])
        has_name = any(term in joined for term in ["人员姓名", "姓名", "Name"])
        has_role = any(term in joined for term in ["角色", "Role"])
        if has_site and has_name and has_role:
            return row, index
    return None, None


def choose_sheet(path, requested_name, preferred_name, finder, label):
    if requested_name:
        rows = read_xlsx_rows(path, requested_name)
        header, row_index = finder(rows)
        if header is None:
            raise ValueError(f"Sheet {requested_name} does not look like a {label} sheet.")
        return requested_name, rows, header, row_index

    all_sheets = read_xlsx_sheets(path)
    preferred = [item for item in all_sheets if item[0] == preferred_name]
    candidates = preferred + [item for item in all_sheets if item[0] != preferred_name]
    for name, rows in candidates:
        header, row_index = finder(rows)
        if header is not None:
            return name, rows, header, row_index

    available = ", ".join(name for name, _rows in all_sheets)
    raise ValueError(f"Could not find a {label} sheet/header. Checked sheets: {available}")


def parse_excel_sites(path, sheet_name=""):
    chosen_sheet, rows, header_row, header_index_number = choose_sheet(path, sheet_name, "研究中心", find_site_header, "site")
    print(f"Excel site sheet: {chosen_sheet}")
    columns = {
        "code": header_index(header_row, ["中心编号", "中心号", "中心代码", "SiteID", "Site ID", "SiteCode", "Site Code"]),
        "name": header_index(header_row, ["中心名称", "SiteName", "Site Name"]),
        "address": header_index(header_row, ["地址", "Address"]),
        "contactName": header_index(header_row, ["PI姓名", "PI'sName", "PI’sName", "PI Name", "PrincipalInvestigator"]),
        "telephone": header_index(header_row, ["PI电话", "PI'sTelephone", "PI’sTelephone", "PI Phone", "PI Telephone"]),
        "email": header_index(header_row, ["PI电子邮件地址", "PI电子邮件", "PI邮箱", "PI Email", "PI E-mail"]),
    }

    sites = []
    for excel_row_number, row in enumerate(rows[header_index_number + 1:], start=header_index_number + 2):
        if not any(normalize_cell(value) for value in row):
            continue
        site = {"excelSheet": chosen_sheet, "excelRow": excel_row_number}
        for key, col in columns.items():
            if col is not None and col < len(row):
                site[key] = normalize_cell(row[col])
        site["code"] = normalize_site_code(site.get("code", ""))
        # Ignore genuinely empty/template rows, but keep partially filled rows so
        # the email intake validator can report their missing required fields.
        if not any(normalize_cell(site.get(key)) for key in columns):
            continue
        sites.append(site)
    return sites


def parse_excel_users(path, sheet_name=""):
    chosen_sheet, rows, header_row, header_index_number = choose_sheet(path, sheet_name, "研究人员信息", find_user_header, "user")
    print(f"Excel user sheet: {chosen_sheet}")
    columns = {
        "siteCode": header_index(header_row, ["中心编号", "中心号", "中心代码", "SiteID", "Site ID", "SiteCode", "Site Code"]),
        "name": header_index(header_row, ["人员姓名", "姓名", "Name"]),
        "email": header_index(header_row, ["邮箱", "电子邮件", "邮件", "Email", "E-mail"]),
        "telephone": header_index(header_row, ["电话", "手机号", "手机", "Telephone", "Phone", "Mobile"]),
        "role": header_index(header_row, ["角色", "Role"]),
        "operation": header_index(header_row, ["账号操作类型", "账号更新", "操作类型", "账号操作", "操作", "Action", "Operation", "Account Operation", "Account Update"]),
        "note": header_index(header_row, ["备注", "Note"]),
        "initials": header_index(header_row, ["首字母", "Initials"]),
        "username": header_index(header_row, ["用户名", "Username"]),
        "firstName": header_index(header_row, ["First Name", "FirstName", "名"]),
        "lastName": header_index(header_row, ["Last Name", "LastName", "姓"]),
        "visitorAppChoice": header_index(header_row, ["Visitor APP", "Visitor App", "Visitor End", "Visit End", "受试者端", "访客端"]),
        "investigatorAppChoice": header_index(header_row, ["Investigator APP", "Investigator App", "Investigator End", "研究者端"]),
        "notificationMethod": header_index(header_row, ["Notification Method", "Notification", "通知方式", "通知方法"]),
    }

    users = []
    for excel_row_number, row in enumerate(rows[header_index_number + 1:], start=header_index_number + 2):
        if not any(normalize_cell(value) for value in row):
            continue
        user = {"excelSheet": chosen_sheet, "excelRow": excel_row_number}
        if columns["operation"] is not None:
            user["_hasOperationColumn"] = True
        for key, col in columns.items():
            if col is not None and col < len(row):
                if key == "email":
                    user["_rawEmail"] = "" if row[col] is None else str(row[col])
                user[key] = normalize_cell(row[col])
        # A row with only an email, phone, site, or operation is still an
        # attempted application row. Keep it so missing fields are not hidden.
        if not any(normalize_cell(user.get(key)) for key in columns):
            continue
        user["siteCodes"] = split_site_codes(user.get("siteCode", ""))
        users.append(user)
    return users


def parse_all_excels(paths, site_sheet="", user_sheet=""):
    sites = []
    users = []
    for path in paths:
        print(f"\nExcel file: {path}")
        file_sites = parse_excel_sites(path, site_sheet)
        file_users = parse_excel_users(path, user_sheet)
        for site in file_sites:
            site["excelFile"] = path
        for user in file_users:
            user["excelFile"] = path
        sites.extend(file_sites)
        users.extend(file_users)
    return sites, users


def parse_all_excel_users(paths, user_sheet=""):
    users = []
    for path in paths:
        print(f"\nExcel file: {path}")
        file_users = parse_excel_users(path, user_sheet)
        for user in file_users:
            user["excelFile"] = path
        users.extend(file_users)
    return users


def parse_role_filter(value):
    raw = normalize_cell(value)
    if not raw:
        return set()
    return {part.strip().upper() for part in re.split(r"[,，、;；\s]+", raw) if part.strip()}


def filter_users_by_roles(users, role_filter):
    if not role_filter:
        return users
    return [user for user in users if normalize_cell(user.get("role")).upper() in role_filter]


def validate_site(site):
    required = ["code", "name"]
    return [field for field in required if not normalize_cell(site.get(field))]


def validate_user(user):
    required = ["name", "email", "telephone", "role"]
    missing = [field for field in required if not normalize_cell(user.get(field))]
    email_issue = validate_email_field(user)
    if email_issue:
        missing.append(email_issue)
    if user.get("_requireOperation") and not normalize_cell(user.get("operation")):
        missing.append("operation")
    role_name = normalize_cell(user.get("role")).upper()
    site_codes = user.get("siteCodes") or split_site_codes(user.get("siteCode", ""))
    if user.get("_requireOperation"):
        if not site_codes:
            missing.append("siteCode")
    elif role_name in SITE_ROLES and not site_codes:
        missing.append("siteCode")
    missing.extend(validate_generated_initials(user))
    return missing


def validate_user_identity(user):
    if user.get("_requireOperation"):
        required = ["name", "email", "telephone", "role", "operation"]
        missing = [field for field in required if not normalize_cell(user.get(field))]
        email_issue = validate_email_field(user)
        if email_issue:
            missing.append(email_issue)
        if not (user.get("siteCodes") or split_site_codes(user.get("siteCode", ""))):
            missing.append("siteCode")
        if not user.get("username"):
            missing.extend(validate_generated_initials(user))
        return missing

    if user.get("username"):
        return []

    missing = []
    role_name = normalize_cell(user.get("role")).upper()
    if not normalize_cell(user.get("name")):
        missing.append("name")
    if not role_name:
        missing.append("role")
    site_codes = user.get("siteCodes") or split_site_codes(user.get("siteCode", ""))
    if role_name in SITE_ROLES and role_name not in USERNAME_ROLES_WITHOUT_SITE_CODE and not site_codes:
        missing.append("siteCode")
    missing.extend(validate_generated_initials(user))
    return missing


def site_payload(site):
    def optional_site_value(field):
        return normalize_cell(site.get(field)) or DEFAULT_OPTIONAL_SITE_VALUE

    return {
        "code": normalize_site_code(site.get("code", "")),
        "name": normalize_cell(site.get("name")),
        "address": optional_site_value("address"),
        "contactName": optional_site_value("contactName"),
        "telephone": optional_site_value("telephone"),
        "email": optional_site_value("email"),
    }


def site_update_payload(site, existing):
    payload = {
        "code": normalize_site_code(existing.get("code")) or normalize_site_code(site.get("code", "")),
        "name": normalize_cell(site.get("name")) or normalize_cell(existing.get("name")),
        "address": normalize_cell(existing.get("address")) or DEFAULT_OPTIONAL_SITE_VALUE,
        "contactName": normalize_cell(existing.get("contactName")) or DEFAULT_OPTIONAL_SITE_VALUE,
        "email": normalize_cell(existing.get("email")) or DEFAULT_OPTIONAL_SITE_VALUE,
        "telephone": normalize_cell(existing.get("telephone")) or DEFAULT_OPTIONAL_SITE_VALUE,
    }
    for field in ("address", "contactName", "email", "telephone"):
        value = normalize_cell(site.get(field))
        if value:
            payload[field] = value
    return payload


def site_payload_changes(payload, existing):
    changes = {}
    for field in ("code", "name", "address", "contactName", "email", "telephone"):
        new_value = normalize_cell(payload.get(field))
        old_value = normalize_cell(existing.get(field))
        if new_value != old_value:
            changes[field] = (old_value, new_value)
    return changes


def lookup_sites_by_code(token, trial_id, site_code):
    encoded_code = urllib.parse.quote(site_code)
    path = f"/trialapi/trials/{trial_id}/sites?code={encoded_code}&name=&status=&page=1&size=20&countOnly=true"
    status, data, raw = request_json("GET", path, token=token)
    sites = first_list(data, ["sites", "data", "list", "items", "records"])
    exact = [site for site in sites if normalize_site_code(site.get("code")) == site_code]
    return status, data.get("code"), exact, raw


def lookup_sites_by_name(token, trial_id, site_name):
    encoded_name = urllib.parse.quote(site_name)
    path = f"/trialapi/trials/{trial_id}/sites?code=&name={encoded_name}&status=&page=1&size=20&countOnly=true"
    status, data, raw = request_json("GET", path, token=token)
    sites = first_list(data, ["sites", "data", "list", "items", "records"])
    exact = [site for site in sites if normalize_cell(site.get("name")) == site_name]
    return status, data.get("code"), exact, raw


def lookup_all_project_sites(token, trial_id, page_size=100, max_pages=1000):
    sites = []
    seen = set()
    for page in range(1, max_pages + 1):
        query = urllib.parse.urlencode({
            "code": "",
            "name": "",
            "status": "",
            "page": page,
            "size": page_size,
            "countOnly": "true",
        })
        status, data, raw = request_json("GET", f"/trialapi/trials/{trial_id}/sites?{query}", token=token)
        if status != 200 or data.get("code") != "SUCCESS":
            print("Could not read all STUDY project sites.")
            print("Response preview:", redact(raw[:800]))
            raise RuntimeError("Failed to read all STUDY project sites for an All-center request.")
        batch = first_list(data, ["sites", "data", "list", "items", "records"])
        if not batch:
            break
        added = 0
        for site in batch:
            key = internal_site_id(site) or normalize_site_code(site.get("code"))
            if not key or key in seen:
                continue
            seen.add(key)
            sites.append(site)
            added += 1
        if added == 0:
            break
    else:
        raise RuntimeError("STUDY project site pagination exceeded the safety limit.")
    return sites


def internal_site_id(site):
    return site.get("_id") or site.get("id")


def build_site_map(token, trial_id, site_codes):
    site_map = {}
    missing = []
    for site_code in sorted(code for code in site_codes if code and code != "所有"):
        status, code, sites, raw = lookup_sites_by_code(token, trial_id, site_code)
        if not sites:
            missing.append(site_code)
            continue
        site = sites[0]
        site_id = internal_site_id(site)
        if site_id:
            site_map[site_code] = {"id": site_id, "name": site.get("name"), "code": site.get("code")}
        else:
            print(f"Found site {site_code} but could not find internal site ID.")
            print("Site preview:", json.dumps(site, ensure_ascii=False)[:800])
    return site_map, missing


def lookup_users(token, username="", phone="", email="", site_id="", size=100):
    query = urllib.parse.urlencode({
        "siteID": site_id,
        "username": username,
        "phone": phone,
        "email": email,
        "page": 1,
        "size": size,
        "countOnly": "true",
        "isNeedMutiSites": "true",
    })
    status, data, raw = request_json("GET", f"/userapi/users?{query}", token=token)
    users = first_list(data, ["users", "data", "list", "items", "records"])
    return status, data.get("code"), users, raw


def lookup_users_by_username(token, username):
    status, code, users, raw = lookup_users(token, username=username, size=20)
    exact = [user for user in users if normalize_username(user.get("username")) == username]
    return status, code, exact, raw


def existing_site_codes(existing_user):
    values = []
    for key in ("siteID", "siteIDs"):
        current = existing_user.get(key)
        if isinstance(current, dict):
            values.append(current)
        elif isinstance(current, list):
            values.extend(item for item in current if isinstance(item, dict))
    return {normalize_site_code(item.get("code")) for item in values if normalize_cell(item.get("code"))}


def contact_matches_disable_row(existing_user, user):
    checks = []
    requested_email = normalize_email(user.get("email"))
    existing_email = normalize_email(existing_user.get("email"))
    if requested_email and existing_email:
        checks.append(requested_email == existing_email)

    requested_phone = normalize_phone(user.get("telephone") or user.get("phone"))
    existing_phone = normalize_phone(existing_user.get("phone") or existing_user.get("telephone"))
    if requested_phone and existing_phone:
        checks.append(requested_phone == existing_phone)

    return any(checks)


def legacy_username_matches_disable_row(existing_user, user):
    existing_username = compact_username(existing_user.get("username"))
    if not existing_username:
        return False

    role_name = compact_username(user.get("role"))
    if role_name and role_name not in existing_username:
        return False

    initials = compact_username(user.get("initials") or ascii_initials(user.get("name", "")))
    if initials and not re.search(re.escape(initials) + r"\d*$", existing_username):
        return False

    return bool(role_name or initials)


def unique_existing_users(users):
    unique = []
    seen = set()
    for user in users:
        key = internal_user_id(user) or normalize_username(user.get("username"))
        if not key or key in seen:
            continue
        seen.add(key)
        unique.append(user)
    return unique


def find_disable_fallback_candidates(token, user):
    # Disable rows may refer to older accounts whose username format predates the
    # current generator. Require contact + role/initials so shared emails do not
    # accidentally disable the wrong person.
    candidates = []
    email = normalize_email(user.get("email"))
    phone = normalize_cell(user.get("telephone") or user.get("phone"))
    if email:
        _status, _code, users, _raw = lookup_users(token, email=email)
        candidates.extend(users)
    if phone:
        _status, _code, users, _raw = lookup_users(token, phone=phone)
        candidates.extend(users)

    requested_sites = {code for code in (user.get("siteCodes") or split_site_codes(user.get("siteCode", ""))) if code and code != "所有"}
    matches = [
        existing
        for existing in unique_existing_users(candidates)
        if contact_matches_disable_row(existing, user)
        and legacy_username_matches_disable_row(existing, user)
    ]

    site_matches = [existing for existing in matches if requested_sites and existing_site_codes(existing) & requested_sites]
    if len(site_matches) == 1:
        return site_matches
    if len(site_matches) > 1:
        return site_matches
    return matches


def resolve_disable_user(token, user):
    expected_username = build_username(user)
    _status, _code, existing, _raw = lookup_users_by_username(token, expected_username)
    if existing:
        return expected_username, existing[0], "exact", []

    if has_explicit_username(user):
        return expected_username, None, "missing", []

    exact_legacy_matches = []
    for candidate in generated_username_candidates(user)[1:]:
        _status, _code, users, _raw = lookup_users_by_username(token, candidate)
        exact_legacy_matches.extend(users)
    exact_legacy_matches = unique_existing_users(exact_legacy_matches)
    if len(exact_legacy_matches) == 1:
        existing = exact_legacy_matches[0]
        return normalize_username(existing.get("username")) or expected_username, existing, "legacy", []
    if len(exact_legacy_matches) > 1:
        return expected_username, None, "ambiguous", exact_legacy_matches

    fallback = find_disable_fallback_candidates(token, user)
    if len(fallback) == 1:
        existing = fallback[0]
        return normalize_username(existing.get("username")) or expected_username, existing, "legacy", []
    if len(fallback) > 1:
        return expected_username, None, "ambiguous", fallback
    return expected_username, None, "missing", []



def lookup_roles(token):
    status, data, raw = request_json("GET", "/userapi/roles?page=&size=&countOnly=true", token=token)
    print("Role lookup status:", status)
    print("Role lookup code:", data.get("code"))
    roles = first_list(data, ["roles", "data", "list", "items", "records"])
    if not roles:
        print("Could not read STUDY role list.")
        print("Response preview:", redact(raw[:800]))
        sys.exit(1)
    return roles


def find_role_id(roles, role_name):
    wanted = role_name.strip().upper()
    for role in roles:
        candidates = [
            str(role.get("name", "")),
            str(role.get("code", "")),
            str(role.get("roleName", "")),
            str(role.get("displayName", "")),
        ]
        if any(candidate.strip().upper() == wanted for candidate in candidates):
            return role.get("_id") or role.get("id")
    return None


def chinese_char_initial(char):
    if char in PINYIN_INITIAL_OVERRIDES:
        return PINYIN_INITIAL_OVERRIDES[char]
    try:
        encoded = char.encode("gb2312")
    except UnicodeEncodeError:
        return ""
    if len(encoded) != 2:
        return ""
    code = encoded[0] * 256 + encoded[1] - 65536
    for start, end, initial in PINYIN_INITIAL_RANGES:
        if start <= code <= end:
            return initial
    return ""


def is_cjk_char(char):
    if not char:
        return False
    code = ord(char)
    return (
        0x3400 <= code <= 0x4DBF
        or 0x4E00 <= code <= 0x9FFF
        or 0x20000 <= code <= 0x2A6DF
        or 0x2A700 <= code <= 0x2B73F
        or 0x2B740 <= code <= 0x2B81F
        or 0x2B820 <= code <= 0x2CEAF
        or 0xF900 <= code <= 0xFAFF
    )


def unmapped_initial_chars(name):
    if generated_initials_complete(name):
        return []
    seen = set()
    missing = []
    for char in name or "":
        if not is_cjk_char(char) or chinese_char_initial(char):
            continue
        if char not in seen:
            seen.add(char)
            missing.append(char)
    return missing


def generated_initials_complete(name):
    value = normalize_cell(name)
    if not value:
        return False
    if re.findall(r"[A-Za-z]+", value):
        return True
    cjk_count = sum(1 for char in value if is_cjk_char(char))
    initials = ascii_initials(value)
    if not cjk_count:
        return bool(initials)
    return len(initials) >= cjk_count


def validate_generated_initials(user):
    if user.get("username") or user.get("initials"):
        return []
    name = user.get("name", "")
    if generated_initials_complete(name):
        return []
    missing_chars = "".join(unmapped_initial_chars(name))
    if missing_chars:
        return [f"initials for unmapped Chinese characters: {missing_chars}"]
    return ["initials"]


def pypinyin_initials(name):
    try:
        from pypinyin import Style, lazy_pinyin
    except Exception:
        return ""
    try:
        parts = lazy_pinyin(name or "", style=Style.FIRST_LETTER, errors=lambda chars: [""] * len(chars))
    except Exception:
        return ""
    return "".join(part[:1].upper() for part in parts if part)


def macos_pinyin_initials(name):
    value = normalize_cell(name)
    if not value or not any(is_cjk_char(char) for char in value):
        return ""
    if value in MACOS_PINYIN_CACHE:
        return MACOS_PINYIN_CACHE[value]

    try:
        result = subprocess.run(
            ["osascript", "-l", "JavaScript", "-e", MACOS_PINYIN_SCRIPT],
            env={**os.environ, "PINYIN_INPUT": json.dumps(value, ensure_ascii=False)},
            text=True,
            capture_output=True,
            timeout=8,
        )
    except Exception:
        MACOS_PINYIN_CACHE[value] = ""
        return ""

    if result.returncode != 0:
        MACOS_PINYIN_CACHE[value] = ""
        return ""

    transliterated = "\n".join(part for part in [result.stdout, result.stderr] if part)
    tokens = re.findall(r"[A-Za-z]+", transliterated)
    cjk_chars = [char for char in value if is_cjk_char(char)]
    if len(tokens) < len(cjk_chars):
        initials = ""
    else:
        initials = "".join(
            PINYIN_INITIAL_OVERRIDES.get(char, tokens[index][0].upper())
            for index, char in enumerate(cjk_chars)
        )
    MACOS_PINYIN_CACHE[value] = initials
    return initials


def ascii_initials(name):
    parts = re.findall(r"[A-Za-z]+", name or "")
    if parts:
        return "".join(part[0] for part in parts[:3]).upper()
    return (
        macos_pinyin_initials(name)
        or pypinyin_initials(name)
        or "".join(chinese_char_initial(char) for char in name or "").upper()
    )


def username_site_code(user, role):
    if role not in SITE_ROLES or role in USERNAME_ROLES_WITHOUT_SITE_CODE:
        return ""
    for site_code in user.get("siteCodes") or split_site_codes(user.get("siteCode", "")):
        if site_code and site_code != "所有":
            return site_code
    return ""


def build_username(user):
    if user.get("_resolvedUsername"):
        return user["_resolvedUsername"]
    if user.get("username"):
        return normalize_username(user["username"])
    role = normalize_cell(user.get("role")).upper()
    site_code = username_site_code(user, role)
    initials = (user.get("initials") or ascii_initials(user.get("name", ""))).strip().upper()
    return normalize_username(f"{PROJECT_ID}-{role}{site_code}{initials}")


def has_explicit_username(user):
    return bool(normalize_cell(user.get("username")))


def generated_username_candidates(user):
    """Return exact current and deterministic historical username formats."""
    expected = build_username(user)
    candidates = [expected]
    if has_explicit_username(user):
        return candidates

    project_id = normalize_username(PROJECT_ID)
    project_ids = [project_id]
    if project_id.startswith("U-") and len(project_id) > 2:
        project_ids.append("U" + project_id[2:])
    elif project_id.startswith("U") and not project_id.startswith("U-") and len(project_id) > 1:
        project_ids.append("U-" + project_id[1:])

    role = normalize_cell(user.get("role")).upper()
    site_code = normalize_username(username_site_code(user, role))
    initials = normalize_username(user.get("initials") or ascii_initials(user.get("name", "")))
    canonical = normalize_username(f"{project_id}-{role}{site_code}{initials}")
    duplicate_suffix = expected[len(canonical):] if expected.startswith(canonical) else ""
    bodies = [
        f"{role}{site_code}{initials}",
        f"{role}{site_code}-{initials}",
        f"{role}-{site_code}{initials}" if site_code else f"{role}-{initials}",
        f"{role}-{site_code}-{initials}" if site_code else f"{role}-{initials}",
    ]
    for candidate_project_id in project_ids:
        for body in bodies:
            candidate = normalize_username(f"{candidate_project_id}-{body}{duplicate_suffix}")
            if candidate and candidate not in candidates:
                candidates.append(candidate)
    return candidates


def resolve_update_user(token, user):
    expected_username = build_username(user)
    matches = []
    for candidate in generated_username_candidates(user):
        _status, _code, existing, _raw = lookup_users_by_username(token, candidate)
        for item in existing:
            matches.append((candidate, item))

    unique_matches = []
    seen = set()
    for candidate, existing in matches:
        key = internal_user_id(existing) or normalize_username(existing.get("username"))
        if not key or key in seen:
            continue
        seen.add(key)
        unique_matches.append((candidate, existing))

    if len(unique_matches) == 1:
        candidate, existing = unique_matches[0]
        actual_username = normalize_username(existing.get("username")) or candidate
        match_type = "exact" if actual_username == expected_username else "legacy"
        return actual_username, existing, match_type, []
    if len(unique_matches) > 1:
        return expected_username, None, "ambiguous", [item for _candidate, item in unique_matches]
    return expected_username, None, "missing", []


def build_base_username(user):
    original = user.pop("_resolvedUsername", None)
    try:
        return build_username(user)
    finally:
        if original:
            user["_resolvedUsername"] = original


def assign_run_usernames(users):
    generated_groups = {}
    for user in users:
        base_username = build_base_username(user)
        if has_explicit_username(user):
            user["_resolvedUsername"] = base_username
            continue
        generated_groups.setdefault(base_username, []).append(user)

    for base_username, grouped_users in generated_groups.items():
        if len(grouped_users) == 1:
            grouped_users[0]["_resolvedUsername"] = base_username
            continue
        for index, user in enumerate(grouped_users, start=1):
            user["_resolvedUsername"] = f"{base_username}{index}"
    return users


def note_device_types_from_text(note):
    text = normalize_cell(note)
    if not text:
        return []
    lowered = text.lower()
    device_types = []
    if any(marker.lower() in lowered for marker in VISITOR_APP_MARKERS):
        device_types.append("visit")
    if any(marker.lower() in lowered for marker in INVESTIGATOR_APP_MARKERS):
        device_types.append("doctor")
    return device_types


def note_device_types(user):
    return note_device_types_from_text(user.get("note"))


def parse_bool_choice(value):
    text = normalize_cell(value).lower()
    if not text:
        return None
    if text in {"1", "y", "yes", "true", "on", "select", "selected", "checked", "开", "开通", "启用", "选择", "选中", "是", "有"}:
        return True
    if text in {"0", "n", "no", "false", "off", "clear", "cleared", "unselect", "unchecked", "关", "关闭", "停用", "取消", "不选", "否", "无"}:
        return False
    return None


def has_device_choice(user):
    return parse_bool_choice(user.get("visitorAppChoice")) is not None or parse_bool_choice(user.get("investigatorAppChoice")) is not None


def device_types_from_choices(user, existing_device_types=None, preserve_unspecified=False):
    if preserve_unspecified:
        selected = set(existing_device_types or [])
    else:
        selected = set()

    visitor_choice = parse_bool_choice(user.get("visitorAppChoice"))
    investigator_choice = parse_bool_choice(user.get("investigatorAppChoice"))

    if visitor_choice is True:
        selected.add("visit")
    elif visitor_choice is False:
        selected.discard("visit")

    if investigator_choice is True:
        selected.add("doctor")
    elif investigator_choice is False:
        selected.discard("doctor")

    return [item for item in ("visit", "doctor") if item in selected]


def notification_methods_from_text(value):
    text = normalize_cell(value)
    if not text:
        return None
    lowered = text.lower()
    methods = []
    if "email" in lowered or "mail" in lowered or "邮箱" in text or "邮件" in text:
        methods.append("EMAIL")
    if "sms" in lowered or "短信" in text:
        methods.append("SMS")
    if not methods and lowered in {"none", "no", "off", "无", "不选"}:
        return []
    return methods or None


def format_device_types(device_types):
    labels = []
    for device_type in device_types or []:
        if device_type == "visit":
            labels.append("Visitor APP")
        elif device_type == "doctor":
            labels.append("Investigator APP")
        else:
            labels.append(str(device_type))
    return ", ".join(labels) if labels else "none"


def authorized_notifications(user):
    configured = user.get("authorizedNotifications")
    if configured:
        return [normalize_cell(item).upper() for item in configured if normalize_cell(item)]
    notification_methods = notification_methods_from_text(user.get("notificationMethod"))
    if notification_methods is not None:
        return notification_methods
    return list(DEFAULT_AUTHORIZED_NOTIFICATIONS)


def format_authorized_notifications(notifications):
    labels = []
    for notification in notifications or []:
        value = normalize_cell(notification).upper()
        if value == "EMAIL":
            labels.append("Email")
        elif value == "SMS":
            labels.append("SMS")
        elif value:
            labels.append(value)
    return ", ".join(labels) if labels else "none"


def format_payload_device_types(payload):
    if "deviceTypes" not in payload:
        return "unchanged"
    return format_device_types(payload.get("deviceTypes"))


def format_payload_authorized_notifications(payload):
    if "authorizedNotifications" not in payload:
        return "unchanged"
    return format_authorized_notifications(payload.get("authorizedNotifications"))


def notification_field_rejected(status, data, raw):
    if status in {200, 201} and isinstance(data, dict) and data.get("code") == "SUCCESS":
        return False
    text = (json.dumps(data, ensure_ascii=False) if isinstance(data, dict) else "") + " " + (raw or "")
    lowered = text.lower()
    if "authorizednotifications" not in lowered and "notification" not in lowered:
        return False
    return any(
        marker in lowered
        for marker in [
            "not allowed",
            "unknown",
            "unexpected",
            "unsupported",
            "invalid field",
            "invalid parameter",
            "cannot",
        ]
    )


def request_user_json(method, path, payload, token):
    status, data, raw = request_json(method, path, payload, token=token)
    if notification_field_rejected(status, data, raw) and "authorizedNotifications" in payload:
        retry_payload = dict(payload)
        retry_payload.pop("authorizedNotifications", None)
        print("Notification Method field was rejected by this project; retrying without that field.")
        return request_json(method, path, retry_payload, token=token)
    return status, data, raw


def note_update_action(user, default_unrecognized_action=""):
    if user.get("_hasOperationColumn") or normalize_cell(user.get("operation")):
        return operation_update_action(user, default_unrecognized_action)

    note = normalize_cell(user.get("note"))
    if not note:
        return "open"
    note_lower = note.lower()
    has_disable = any(marker.lower() in note_lower for marker in DISABLE_NOTE_MARKERS)
    has_open = any(marker.lower() in note_lower for marker in OPEN_NOTE_MARKERS)
    if has_disable and has_open:
        return "ambiguous"
    if has_disable:
        return "disable"
    if has_open:
        return "open"
    if note_device_types_from_text(note):
        return "open"
    return default_unrecognized_action


def operation_update_action(user, default_unrecognized_action=""):
    operation = normalize_cell(user.get("operation"))
    if not operation:
        return "open"
    operation_lower = operation.lower()
    has_disable = any(marker.lower() in operation_lower for marker in DISABLE_NOTE_MARKERS)
    has_open = any(marker.lower() in operation_lower for marker in OPEN_NOTE_MARKERS)
    has_update = any(marker.lower() in operation_lower for marker in UPDATE_OPERATION_MARKERS)
    active_action_count = sum(bool(value) for value in (has_disable, has_open, has_update))
    if active_action_count > 1:
        return "ambiguous"
    if has_disable:
        return "disable"
    if has_update:
        return "update"
    if has_open:
        return "open"
    return default_unrecognized_action


def classify_update_users(users, default_unrecognized_action=""):
    open_users = []
    update_only_users = []
    disable_users = []
    ignored = []
    ambiguous = []
    for user in users:
        action = note_update_action(user, default_unrecognized_action)
        if action in {"open", "update"}:
            open_users.append(user)
            if action == "update":
                update_only_users.append(user)
        elif action == "disable":
            disable_users.append(user)
        elif action == "ambiguous":
            ambiguous.append(user)
        else:
            ignored.append(user)
    return {
        "open": open_users,
        "update_only": update_only_users,
        "disable": disable_users,
        "ignored": ignored,
        "ambiguous": ambiguous,
    }


def normalize_user(user, site_ids, role_id):
    role_name = user["role"].strip().upper()
    base_role = 0 if role_name in SITE_ROLES else 1
    first_site_id = site_ids[0] if site_ids else None

    body = {
        "username": build_username(user),
        "email": user["email"].strip(),
        "phone": str(user.get("telephone") or user.get("phone") or "").strip(),
        "phoneCountryCode": str(user.get("phoneCountryCode") or "").replace("+", "").strip(),
        "siteID": first_site_id if base_role == 0 else None,
        "siteIDs": site_ids if base_role == 0 else None,
        "userRole": role_id,
        "role": base_role,
        "language": user.get("language") or "zh_CN",
        "deviceTypes": (
            user.get("deviceTypes")
            if user.get("deviceTypes") is not None
            else device_types_from_choices(user)
            if has_device_choice(user)
            else note_device_types(user)
        ),
        "authorizedNotifications": authorized_notifications(user),
    }

    if user.get("firstName") or user.get("lastName"):
        body["firstName"] = str(user.get("firstName") or "").strip()
        body["lastName"] = str(user.get("lastName") or "").strip()

    return body


def collect_required_site_codes(users, project_scope_site_codes=None):
    project_scope_site_codes = set(project_scope_site_codes or [])
    required = set()
    for user in users:
        role = normalize_cell(user.get("role")).upper()
        if role in SITE_ROLES:
            for site_code in user.get("siteCodes") or split_site_codes(user.get("siteCode", "")):
                if site_code == "所有":
                    required.update(project_scope_site_codes)
                elif site_code:
                    required.add(site_code)
    return required


def has_project_wide_site_request(users):
    for user in users:
        role = normalize_cell(user.get("role")).upper()
        site_codes = user.get("siteCodes") or split_site_codes(user.get("siteCode", ""))
        if role in SITE_ROLES and site_codes == ["所有"]:
            return True
    return False


def fail_if_write_failures(label, **counts):
    failures = {name: int(value or 0) for name, value in counts.items() if int(value or 0) > 0}
    if not failures:
        return
    details = ", ".join(f"{name}={value}" for name, value in failures.items())
    print(f"\n{label} failed: {details}")
    raise SystemExit(1)


def fail_if_user_blockers(label, summary, keys, skip_blocked_rows=False):
    blockers = {key: len(summary.get(key) or []) for key in keys if summary.get(key)}
    if not blockers:
        return
    details = ", ".join(f"{key}={value}" for key, value in blockers.items())
    if skip_blocked_rows:
        print(f"\n{label}: skipped blocked rows ({details}); valid rows may continue.")
        return
    print(f"\n{label} blocked: {details}")
    raise SystemExit(1)


def summarize_sites(token, trial_id, sites):
    valid = []
    invalid = []
    existing = []
    update_payloads = []
    conflicts = []
    missing = []

    for site in sites:
        missing_fields = validate_site(site)
        if missing_fields:
            invalid.append((site, missing_fields))
            continue

        payload = site_payload(site)
        _status, _code, by_code, _raw = lookup_sites_by_code(token, trial_id, payload["code"])
        if by_code:
            existing_site = by_code[0]
            update_payload = site_update_payload(site, existing_site)
            changes = site_payload_changes(update_payload, existing_site)
            existing.append((site, existing_site, update_payload, changes))
            if changes:
                site_id = internal_site_id(existing_site)
                if site_id:
                    update_payloads.append({
                        "site": site,
                        "existing": existing_site,
                        "siteID": site_id,
                        "payload": update_payload,
                        "changes": changes,
                    })
                else:
                    invalid.append((site, ["existing site internal ID"]))
            continue

        _status, _code, by_name, _raw = lookup_sites_by_name(token, trial_id, payload["name"])
        if by_name:
            conflicts.append((site, by_name[0]))
            continue

        missing.append(site)
        valid.append(payload)

    return {
        "valid_payloads": valid,
        "update_payloads": update_payloads,
        "invalid": invalid,
        "existing": existing,
        "conflicts": conflicts,
        "missing": missing,
    }


def prepare_users(token, trial_id, users, pending_site_codes=None, project_scope_site_codes=None):
    pending_site_codes = set(pending_site_codes or [])
    roles = lookup_roles(token)
    required_site_codes = collect_required_site_codes(users, project_scope_site_codes=project_scope_site_codes)
    site_map, missing_site_codes = build_site_map(token, trial_id, required_site_codes)

    prepared = []
    invalid = []
    missing_role = []
    missing_site = []
    waiting_for_site_creation = []
    existing_users = []
    duplicates = []
    seen_usernames = set()

    for user in users:
        missing_fields = validate_user(user)
        if missing_fields:
            invalid.append((user, missing_fields))
            continue

        role_name = user["role"].strip().upper()
        role_id = find_role_id(roles, role_name)
        if not role_id:
            missing_role.append(user)
            continue

        site_ids = []
        if role_name in SITE_ROLES:
            raw_site_codes = user.get("siteCodes") or split_site_codes(user.get("siteCode", ""))
            user_site_codes = [code for code in raw_site_codes if code != "所有"]
            if raw_site_codes == ["所有"]:
                user_site_codes = sorted(project_scope_site_codes or [])
                if not user_site_codes:
                    missing_site.append((user, ["All (no project sites found in STUDY or Excel)"]))
                    continue
            missing_for_user = [code for code in user_site_codes if code not in site_map]
            if missing_for_user:
                waiting_codes = [code for code in missing_for_user if code in pending_site_codes]
                truly_missing_codes = [code for code in missing_for_user if code not in pending_site_codes]
                if waiting_codes:
                    waiting_for_site_creation.append((user, waiting_codes))
                if truly_missing_codes:
                    missing_site.append((user, truly_missing_codes))
                continue
            site_ids = [site_map[code]["id"] for code in user_site_codes]

        payload = normalize_user(user, site_ids, role_id)
        username = payload["username"]
        if username in seen_usernames:
            duplicates.append((user, username))
            continue
        seen_usernames.add(username)

        _status, _code, existing, _raw = lookup_users_by_username(token, username)
        if existing:
            existing_users.append((user, existing[0]))
            continue

        prepared.append(payload)

    return {
        "prepared": prepared,
        "invalid": invalid,
        "missing_role": missing_role,
        "missing_site": missing_site,
        "waiting_for_site_creation": waiting_for_site_creation,
        "existing_users": existing_users,
        "duplicates": duplicates,
        "missing_site_codes": missing_site_codes,
        "site_map": site_map,
    }


def internal_user_id(user):
    return user.get("_id") or user.get("id")


def user_status(user):
    return user.get("status")


def is_disabled_status(status):
    return normalize_cell(status).lower() in {"0", "disabled", "disable", "inactive", "closed", "close"}


def existing_device_types(existing_user):
    values = existing_user.get("deviceTypes")
    if isinstance(values, list):
        return [normalize_cell(item) for item in values if normalize_cell(item)]
    return []


def direct_update_payload(user, existing_user, site_map, roles, project_scope_site_codes=None):
    payload = {"username": build_username(user)}
    missing = []

    if normalize_cell(user.get("email")):
        payload["email"] = normalize_cell(user.get("email"))

    telephone = normalize_cell(user.get("telephone") or user.get("phone"))
    if telephone:
        payload["phone"] = telephone

    phone_country_code = normalize_cell(user.get("phoneCountryCode")).replace("+", "")
    if phone_country_code:
        payload["phoneCountryCode"] = phone_country_code

    if normalize_cell(user.get("firstName")) or normalize_cell(user.get("lastName")):
        payload["firstName"] = normalize_cell(user.get("firstName"))
        payload["lastName"] = normalize_cell(user.get("lastName"))

    role_name = normalize_cell(user.get("role")).upper()
    if role_name:
        role_id = find_role_id(roles, role_name)
        if not role_id:
            missing.append("known STUDY role")
        else:
            payload["userRole"] = role_id
            payload["role"] = 0 if role_name in SITE_ROLES else 1

    raw_site_codes = user.get("siteCodes") or split_site_codes(user.get("siteCode", ""))
    if raw_site_codes:
        if raw_site_codes == ["所有"]:
            user_site_codes = sorted(project_scope_site_codes or [])
            if not user_site_codes:
                missing.append("project sites for All")
        else:
            user_site_codes = [code for code in raw_site_codes if code != "所有"]
        missing_site_codes = [code for code in user_site_codes if code not in site_map]
        if missing_site_codes:
            missing.append("existing STUDY site codes: " + ", ".join(missing_site_codes))
        else:
            site_ids = [site_map[code]["id"] for code in user_site_codes]
            payload["siteID"] = site_ids[0] if site_ids else None
            payload["siteIDs"] = site_ids

    if has_device_choice(user):
        payload["deviceTypes"] = device_types_from_choices(
            user,
            existing_device_types=existing_device_types(existing_user),
            preserve_unspecified=True,
        )
    elif note_device_types(user):
        payload["deviceTypes"] = note_device_types(user)

    notification_methods = notification_methods_from_text(user.get("notificationMethod"))
    if notification_methods is not None:
        payload["authorizedNotifications"] = notification_methods

    changed_keys = set(payload) - {"username"}
    if not changed_keys:
        missing.append("at least one field to update")

    return payload, missing


def prepare_update_actions(
    token,
    trial_id,
    users,
    pending_site_codes=None,
    project_scope_site_codes=None,
    default_unrecognized_action="",
):
    classes = classify_update_users(users, default_unrecognized_action)
    open_users = classes["open"]
    disable_users = classes["disable"]
    direct_update_mode = default_unrecognized_action == "open"

    roles = lookup_roles(token)
    required_site_codes = collect_required_site_codes(open_users, project_scope_site_codes=project_scope_site_codes)
    if direct_update_mode:
        for user in open_users:
            for site_code in user.get("siteCodes") or split_site_codes(user.get("siteCode", "")):
                if site_code == "所有":
                    required_site_codes.update(project_scope_site_codes or [])
                elif site_code:
                    required_site_codes.add(site_code)
    site_map, missing_site_codes = build_site_map(token, trial_id, required_site_codes)

    to_create = []
    to_update = []
    to_reactivate = []
    invalid_open = []
    invalid_disable = []
    missing_role = []
    missing_site = []
    waiting_for_site_creation = []
    missing_update_user = []
    ambiguous_update_user = []
    disable_ready = []
    already_disabled = []
    missing_disable_user = []
    ambiguous_disable_user = []
    duplicates = []
    seen_usernames = set()
    pending_site_codes = set(pending_site_codes or [])
    update_only_ids = {id(user) for user in classes.get("update_only", [])}

    for user in open_users:
        resolved_update_user = None
        update_match_type = ""
        if direct_update_mode:
            missing_fields = validate_user_identity(user)
            if missing_fields:
                invalid_open.append((user, missing_fields))
                continue

            username = build_username(user)
            if username in seen_usernames:
                duplicates.append((user, username))
                continue
            seen_usernames.add(username)

            _status, _code, existing, _raw = lookup_users_by_username(token, username)
            if not existing:
                invalid_open.append((user, ["existing STUDY user for direct update"]))
                continue

            existing_user = existing[0]
            payload, missing_payload_fields = direct_update_payload(
                user,
                existing_user,
                site_map,
                roles,
                project_scope_site_codes=project_scope_site_codes,
            )
            if missing_payload_fields:
                invalid_open.append((user, missing_payload_fields))
                continue

            item = {
                "user": user,
                "existing": existing_user,
                "payload": payload,
                "username": username,
                "userID": internal_user_id(existing_user),
                "status": user_status(existing_user),
            }
            to_update.append(item)
            if is_disabled_status(item["status"]):
                to_reactivate.append(item)
            continue

        missing_fields = validate_user(user)
        if missing_fields:
            invalid_open.append((user, missing_fields))
            continue

        role_name = user["role"].strip().upper()
        role_id = find_role_id(roles, role_name)
        if not role_id:
            missing_role.append(user)
            continue

        if id(user) in update_only_ids:
            resolved_username, resolved_update_user, update_match_type, ambiguous_matches = resolve_update_user(token, user)
            if ambiguous_matches:
                ambiguous_update_user.append((user, build_username(user), ambiguous_matches))
                continue
            if not resolved_update_user:
                missing_update_user.append((user, build_username(user)))
                continue
            user["_resolvedUsername"] = resolved_username

        site_ids = []
        if role_name in SITE_ROLES:
            raw_site_codes = user.get("siteCodes") or split_site_codes(user.get("siteCode", ""))
            user_site_codes = [code for code in raw_site_codes if code != "所有"]
            if raw_site_codes == ["所有"]:
                user_site_codes = sorted(project_scope_site_codes or [])
                if not user_site_codes:
                    missing_site.append((user, ["All (no project sites found in STUDY or Excel)"]))
                    continue
            missing_for_user = [code for code in user_site_codes if code not in site_map]
            if missing_for_user:
                waiting_codes = [code for code in missing_for_user if code in pending_site_codes]
                truly_missing_codes = [code for code in missing_for_user if code not in pending_site_codes]
                if waiting_codes:
                    waiting_for_site_creation.append((user, waiting_codes))
                if truly_missing_codes:
                    missing_site.append((user, truly_missing_codes))
                continue
            site_ids = [site_map[code]["id"] for code in user_site_codes]

        payload = normalize_user(user, site_ids, role_id)
        username = payload["username"]
        if username in seen_usernames:
            duplicates.append((user, username))
            continue
        seen_usernames.add(username)

        if resolved_update_user:
            existing = [resolved_update_user]
        else:
            _status, _code, existing, _raw = lookup_users_by_username(token, username)
        if existing:
            existing_user = existing[0]
            item = {
                "user": user,
                "existing": existing_user,
                "payload": payload,
                "username": username,
                "userID": internal_user_id(existing_user),
                "status": user_status(existing_user),
                "matchType": update_match_type or "exact",
            }
            to_update.append(item)
            if is_disabled_status(item["status"]):
                to_reactivate.append(item)
        else:
            if id(user) in update_only_ids:
                missing_update_user.append((user, username))
            else:
                to_create.append(payload)

    for user in disable_users:
        missing_fields = validate_user_identity(user)
        if missing_fields:
            invalid_disable.append((user, missing_fields))
            continue

        username, existing_user, match_type, ambiguous_matches = resolve_disable_user(token, user)
        if username in seen_usernames:
            duplicates.append((user, username))
            continue
        seen_usernames.add(username)

        if ambiguous_matches:
            ambiguous_disable_user.append((user, username, ambiguous_matches))
            continue
        if not existing_user:
            missing_disable_user.append((user, username))
            continue
        item = {
            "user": user,
            "existing": existing_user,
            "username": username,
            "matchType": match_type,
            "userID": internal_user_id(existing_user),
            "status": user_status(existing_user),
        }
        if is_disabled_status(item["status"]):
            already_disabled.append(item)
        else:
            disable_ready.append(item)

    return {
        "to_create": to_create,
        "to_update": to_update,
        "to_reactivate": to_reactivate,
        "disable_ready": disable_ready,
        "already_disabled": already_disabled,
        "missing_disable_user": missing_disable_user,
        "ambiguous_disable_user": ambiguous_disable_user,
        "invalid_open": invalid_open,
        "invalid_disable": invalid_disable,
        "missing_role": missing_role,
        "missing_site": missing_site,
        "waiting_for_site_creation": waiting_for_site_creation,
        "missing_update_user": missing_update_user,
        "ambiguous_update_user": ambiguous_update_user,
        "duplicates": duplicates,
        "ambiguous": classes["ambiguous"],
        "ignored": classes["ignored"],
        "open_rows": open_users,
        "update_only_rows": classes.get("update_only", []),
        "disable_rows": disable_users,
        "missing_site_codes": missing_site_codes,
        "site_map": site_map,
    }


def print_site_preview(summary):
    print("\nSite preview")
    print(f"- Excel-listed sites not already in STUDY and ready to create: {len(summary['valid_payloads'])}")
    print(f"- sites already existing by code: {len(summary['existing'])}")
    print(f"- existing STUDY sites ready to update from Excel: {len(summary['update_payloads'])}")
    print(f"- site name conflicts: {len(summary['conflicts'])}")
    print(f"- site rows missing required fields: {len(summary['invalid'])}")

    if summary["existing"]:
        print("\nSites already in STUDY:")
        for site, existing, _payload, changes in summary["existing"][:30]:
            status = "needs update" if changes else "no field changes"
            print(f"- row {site.get('excelRow')}: [{site.get('code')}] {site.get('name')} already exists as [{existing.get('code')}] {existing.get('name')} ({status})")

    if summary["update_payloads"]:
        print("\nExisting STUDY sites ready to update:")
        for item in summary["update_payloads"][:30]:
            site = item["site"]
            change_text = ", ".join(
                f"{field}: {old or '<blank>'} -> {new or '<blank>'}"
                for field, (old, new) in item["changes"].items()
            )
            print(f"- row {site.get('excelRow')}: [{site.get('code')}] {site.get('name')} | {change_text}")

    if summary["conflicts"]:
        print("\nSite name conflicts:")
        for site, existing in summary["conflicts"][:30]:
            print(f"- row {site.get('excelRow')}: code {site.get('code')} not found, but name already exists as code {existing.get('code')}: {site.get('name')}")

    if summary["invalid"]:
        print("\nSite rows that cannot be created yet:")
        for site, missing in summary["invalid"][:50]:
            print(f"- row {site.get('excelRow')}, code {site.get('code') or ''}, name {site.get('name') or ''}: missing {', '.join(missing)}")


def print_user_preview(summary):
    print("\nUser preview")
    print(f"- Excel-listed users not already in STUDY and ready to create: {len(summary['prepared'])}")
    print(f"- users already existing by username: {len(summary['existing_users'])}")
    print(f"- duplicate generated/exact usernames skipped: {len(summary['duplicates'])}")
    print(f"- user rows waiting for site creation first: {len(summary['waiting_for_site_creation'])}")
    print(f"- user rows missing required fields: {len(summary['invalid'])}")
    print(f"- user rows with unknown role: {len(summary['missing_role'])}")
    print(f"- user rows whose site is not in STUDY or this Excel: {len(summary['missing_site'])}")

    if summary["existing_users"]:
        print("\nUsers already in STUDY:")
        for user, existing in summary["existing_users"][:80]:
            print(f"- row {user.get('excelRow')}: {build_username(user)} already exists as {existing.get('username')}")

    if summary["duplicates"]:
        print("\nDuplicate usernames skipped:")
        for user, username in summary["duplicates"][:60]:
            print(f"- row {user.get('excelRow')}: {username}")

    if summary["invalid"]:
        print("\nUser rows that cannot be created yet:")
        for user, missing in summary["invalid"][:80]:
            print(f"- row {user.get('excelRow')}, name {user.get('name') or ''}: missing {', '.join(missing)}")

    if summary["missing_role"]:
        print("\nUser rows with roles not found in STUDY:")
        for user in summary["missing_role"][:50]:
            print(f"- row {user.get('excelRow')}, name {user.get('name') or ''}, role {user.get('role') or ''}")

    if summary["waiting_for_site_creation"]:
        print("\nUser rows waiting for sites from this Excel to be created first:")
        for user, codes in summary["waiting_for_site_creation"][:50]:
            print(f"- row {user.get('excelRow')}, name {user.get('name') or ''}: waiting for sites {', '.join(codes)}")

    if summary["missing_site"]:
        print("\nUser rows with site codes missing from both STUDY and this Excel:")
        for user, codes in summary["missing_site"][:50]:
            print(f"- row {user.get('excelRow')}, name {user.get('name') or ''}: missing sites {', '.join(codes)}")

    if summary["prepared"]:
        print("\nPrepared usernames:")
        for body in summary["prepared"][:80]:
            site_count = len(body.get("siteIDs") or [])
            print(f"- {body.get('username')} ({body.get('email')}) sites={site_count} apps={format_device_types(body.get('deviceTypes'))} notifications={format_authorized_notifications(body.get('authorizedNotifications'))}")


def print_update_preview(summary, direct_user_update=False):
    if direct_user_update:
        print("\nSMT direct user update preview")
    else:
        print("\nSMT update preview from operation column/notes")
    print(f"- rows treated as open/add/update/reactivate: {len(summary['open_rows'])}")
    print(f"- rows explicitly marked update-only: {len(summary.get('update_only_rows', []))}")
    print(f"- open/add users not in STUDY and ready to create: {len(summary['to_create'])}")
    print(f"- open/add users already in STUDY and ready to update: {len(summary['to_update'])}")
    print(f"- existing open/add users that also need reactivation: {len(summary['to_reactivate'])}")
    print(f"- update-only rows not found in STUDY: {len(summary.get('missing_update_user', []))}")
    print(f"- update-only rows with multiple historical username matches: {len(summary.get('ambiguous_update_user', []))}")
    print(f"- rows marked close/disable: {len(summary['disable_rows'])}")
    print(f"- existing STUDY users ready to disable: {len(summary['disable_ready'])}")
    print(f"- users already disabled: {len(summary['already_disabled'])}")
    print(f"- disable rows not found in STUDY: {len(summary['missing_disable_user'])}")
    print(f"- disable rows with multiple possible STUDY matches: {len(summary.get('ambiguous_disable_user', []))}")
    print(f"- duplicate generated/exact usernames skipped: {len(summary['duplicates'])}")
    print(f"- ambiguous operation/note values skipped: {len(summary['ambiguous'])}")
    print(f"- rows with no recognized operation/update marker skipped: {len(summary['ignored'])}")
    print(f"- open/add rows missing required fields: {len(summary['invalid_open'])}")
    print(f"- disable rows missing username identity fields: {len(summary['invalid_disable'])}")
    print(f"- open/add rows with unknown role: {len(summary['missing_role'])}")
    print(f"- open/add rows waiting for site creation first: {len(summary['waiting_for_site_creation'])}")
    print(f"- open/add rows whose site is not in STUDY or this Excel: {len(summary['missing_site'])}")

    if summary["to_create"]:
        print("\nOpen/add users ready to create:")
        for body in summary["to_create"][:80]:
            print(f"- {body.get('username')} ({body.get('email')}) apps={format_device_types(body.get('deviceTypes'))} notifications={format_authorized_notifications(body.get('authorizedNotifications'))}")

    if summary["to_update"]:
        print("\nExisting users ready to update from Excel:")
        for item in summary["to_update"][:80]:
            status = item.get("status")
            status_text = f", status={status}" if status is not None else ""
            print(f"- row {item['user'].get('excelRow')}: {item['username']}{status_text} apps={format_payload_device_types(item['payload'])} notifications={format_payload_authorized_notifications(item['payload'])}")

    if summary["disable_ready"]:
        print("\nExisting users ready to disable:")
        for item in summary["disable_ready"][:80]:
            detail = item["user"].get("operation") or item["user"].get("note") or ""
            match_note = " (matched old username format)" if item.get("matchType") == "legacy" else ""
            print(f"- row {item['user'].get('excelRow')}: {item['username']}{match_note} operation={detail}")

    if summary["already_disabled"]:
        print("\nUsers already disabled:")
        for item in summary["already_disabled"][:50]:
            print(f"- row {item['user'].get('excelRow')}: {item['username']}")

    if summary["missing_disable_user"]:
        print("\nDisable rows not found in STUDY:")
        for user, username in summary["missing_disable_user"][:80]:
            print(f"- row {user.get('excelRow')}: expected username {username}, name {user.get('name') or ''}")

    if summary.get("ambiguous_disable_user"):
        print("\nDisable rows with multiple possible STUDY matches:")
        for user, username, matches in summary["ambiguous_disable_user"][:50]:
            matched_usernames = ", ".join(normalize_username(match.get("username")) for match in matches[:10])
            print(f"- row {user.get('excelRow')}: expected username {username}, possible matches {matched_usernames}")

    if summary["duplicates"]:
        print("\nDuplicate usernames skipped:")
        for user, username in summary["duplicates"][:80]:
            print(f"- row {user.get('excelRow')}: {username}")

    if summary["ambiguous"]:
        print("\nAmbiguous operation/note values skipped:")
        for user in summary["ambiguous"][:50]:
            detail = user.get("operation") or user.get("note") or ""
            print(f"- row {user.get('excelRow')}: operation/note={detail}")

    if summary.get("missing_update_user"):
        print("\nUpdate-only rows not found in STUDY:")
        for user, username in summary["missing_update_user"][:80]:
            print(f"- row {user.get('excelRow')}: expected existing username {username}, name {user.get('name') or ''}")

    if summary.get("ambiguous_update_user"):
        print("\nUpdate-only rows with multiple exact username-format matches:")
        for user, username, matches in summary["ambiguous_update_user"][:80]:
            found = ", ".join(normalize_username(item.get("username")) for item in matches[:10])
            print(f"- row {user.get('excelRow')}: expected {username}, matched {found}")

    if summary["invalid_open"]:
        print("\nOpen/add rows that cannot be created or updated yet:")
        for user, missing in summary["invalid_open"][:80]:
            print(f"- row {user.get('excelRow')}, name {user.get('name') or ''}: missing {', '.join(missing)}")

    if summary["invalid_disable"]:
        print("\nDisable rows that cannot identify a user yet:")
        for user, missing in summary["invalid_disable"][:80]:
            print(f"- row {user.get('excelRow')}, name {user.get('name') or ''}: missing {', '.join(missing)}")

    if summary["missing_role"]:
        print("\nOpen/add rows with roles not found in STUDY:")
        for user in summary["missing_role"][:50]:
            print(f"- row {user.get('excelRow')}, name {user.get('name') or ''}, role {user.get('role') or ''}")

    if summary["waiting_for_site_creation"]:
        print("\nOpen/add rows waiting for sites from this Excel to be created first:")
        for user, codes in summary["waiting_for_site_creation"][:50]:
            print(f"- row {user.get('excelRow')}, name {user.get('name') or ''}: waiting for sites {', '.join(codes)}")

    if summary["missing_site"]:
        print("\nOpen/add rows with site codes missing from both STUDY and this Excel:")
        for user, codes in summary["missing_site"][:50]:
            print(f"- row {user.get('excelRow')}, name {user.get('name') or ''}: missing sites {', '.join(codes)}")


def user_has_direct_change(user):
    return any(
        normalize_cell(user.get(field))
        for field in ("email", "telephone", "phone", "phoneCountryCode", "role", "siteCode", "firstName", "lastName", "notificationMethod")
    ) or has_device_choice(user) or bool(note_device_types(user))


def validate_direct_user_parse(user):
    missing = validate_user_identity(user)
    if not user_has_direct_change(user):
        missing.append("field to update")
    return missing


def print_parse_only(sites, users, direct_user_update=False):
    invalid_sites = [(site, validate_site(site)) for site in sites if validate_site(site)]
    user_validator = validate_direct_user_parse if direct_user_update else validate_user
    invalid_users = [(user, user_validator(user)) for user in users if user_validator(user)]
    multi_site_users = [user for user in users if len(user.get("siteCodes") or []) > 1]
    project_wide_users = [user for user in users if user.get("siteCodes") == ["所有"]]
    update_classes = classify_update_users(users, "open" if direct_user_update else "")

    print("\nParse-only check")
    if direct_user_update:
        print("- direct one-user update mode: yes")
    print(f"- sites found: {len(sites)}")
    print(f"- users found: {len(users)}")
    print(f"- site rows missing required fields: {len(invalid_sites)}")
    print(f"- user rows missing required fields: {len(invalid_users)}")
    print(f"- multi-site users found: {len(multi_site_users)}")
    print(f"- project-wide users found: {len(project_wide_users)}")
    print(f"- update rows treated as open/add/update/reactivate: {len(update_classes['open'])}")
    print(f"- update rows explicitly marked update-only: {len(update_classes.get('update_only', []))}")
    print(f"- update rows marked close/disable: {len(update_classes['disable'])}")
    print(f"- update rows with ambiguous operation/note values: {len(update_classes['ambiguous'])}")
    print(f"- rows with no recognized operation/update marker: {len(update_classes['ignored'])}")

    if sites:
        print("\nFirst site row:")
        first = sites[0]
        print(f"- row {first.get('excelRow')}: [{first.get('code')}] {first.get('name')}")

    if users:
        print("\nFirst user row:")
        first = users[0]
        print(f"- row {first.get('excelRow')}: {first.get('name')} role={first.get('role')} sites={','.join(first.get('siteCodes') or [])}")

    if multi_site_users:
        print("\nMulti-site examples:")
        for user in multi_site_users[:10]:
            print(f"- row {user.get('excelRow')}: {user.get('name')} sites={','.join(user.get('siteCodes') or [])}")

    if invalid_sites:
        print("\nSite rows not ready:")
        for site, missing in invalid_sites[:20]:
            print(f"- row {site.get('excelRow')}, code {site.get('code') or ''}: missing {', '.join(missing)}")

    if invalid_users:
        print("\nUser rows not ready:")
        for user, missing in invalid_users[:20]:
            print(f"- row {user.get('excelRow')}, name {user.get('name') or ''}: missing {', '.join(missing)}")


def create_sites(token, trial_id, site_payloads):
    created = 0
    failed = 0
    create_path = f"/trialapi/trials/{trial_id}/sites"
    for index, payload in enumerate(site_payloads, start=1):
        print(f"\nCreating site {index}/{len(site_payloads)}: [{payload.get('code')}] {payload.get('name')}")
        status, data, raw = request_json("POST", create_path, payload, token=token)
        print("Create site status:", status)
        print("Create site code:", data.get("code"))
        print("Create site response:", redact(raw[:1000]))
        if status in {200, 201} and data.get("code") == "SUCCESS":
            created += 1
        else:
            failed += 1
    return created, failed


def update_sites(token, trial_id, site_update_items):
    updated = 0
    failed = 0
    for index, item in enumerate(site_update_items, start=1):
        payload = item["payload"]
        site_id = item["siteID"]
        path = f"/trialapi/trials/{trial_id}/sites/{urllib.parse.quote(str(site_id))}"
        print(f"\nUpdating site {index}/{len(site_update_items)}: [{payload.get('code')}] {payload.get('name')}")
        status, data, raw = request_json("PATCH", path, payload, token=token)
        print("Update site status:", status)
        print("Update site code:", data.get("code"))
        print("Update site response:", redact(raw[:1000]))
        if status in {200, 201} and data.get("code") == "SUCCESS":
            updated += 1
        else:
            failed += 1
    return updated, failed


def create_users(token, user_payloads):
    created = 0
    failed = 0
    for index, payload in enumerate(user_payloads, start=1):
        print(f"\nCreating user {index}/{len(user_payloads)}: {payload.get('username')}")
        status, data, raw = request_user_json("POST", "/userapi/users", payload, token=token)
        print("Create user status:", status)
        print("Create user code:", data.get("code"))
        print("Create user response:", redact(raw[:1200]))
        if status in {200, 201} and data.get("code") == "SUCCESS":
            created += 1
        else:
            failed += 1
    return created, failed


def update_user(token, item):
    path = "/userapi/users/" + urllib.parse.quote(str(item["userID"]))
    return request_user_json("PATCH", path, item["payload"], token=token)


def update_users(token, update_items):
    updated = 0
    failed = 0
    for index, item in enumerate(update_items, start=1):
        print(f"\nUpdating user {index}/{len(update_items)}: {item['username']}")
        status, data, raw = update_user(token, item)
        print("Update user status:", status)
        print("Update user code:", data.get("code"))
        print("Update user response:", redact(raw[:1200]))
        if status in {200, 201} and data.get("code") == "SUCCESS":
            updated += 1
        else:
            failed += 1
    return updated, failed


def set_user_status(token, item, target_status):
    path = "/userapi/users/" + urllib.parse.quote(str(item["userID"])) + "/status"
    return request_json("PATCH", path, {"status": target_status}, token=token)


def set_users_status(token, items, target_status, action_label):
    changed = 0
    failed = 0
    for index, item in enumerate(items, start=1):
        print(f"\n{action_label} user {index}/{len(items)}: {item['username']}")
        status, data, raw = set_user_status(token, item, target_status)
        print(f"{action_label} status:", status)
        print(f"{action_label} code:", data.get("code"))
        print(f"{action_label} response:", redact(raw[:1000]))
        if status in {200, 201} and data.get("code") == "SUCCESS":
            changed += 1
        else:
            failed += 1
    return changed, failed


def main():
    load_env_file(os.path.expanduser("~/.hermes/.env"))
    load_env_file(os.path.join(os.getcwd(), ".env"))

    parser = argparse.ArgumentParser(description="Preview or create STUDY sites and users from one or more Excel workbooks.")
    parser.add_argument("--excel-file", required=True, nargs="+", help="Path(s) to filled Excel workbook(s).")
    parser.add_argument("--site-sheet", default="", help="Optional site sheet name. Defaults to scan/prefer 研究中心.")
    parser.add_argument("--user-sheet", default="", help="Optional user sheet name. Defaults to scan/prefer 研究人员信息.")
    parser.add_argument("--parse-only", action="store_true", help="Only read and validate Excel structure. Do not log in or call STUDY APIs.")
    parser.add_argument("--create", action="store_true", help="Actually create valid Excel-listed sites/users that are not already in STUDY. Without this, preview only.")
    parser.add_argument("--confirm-create", default="", help="Pass /approve to skip the interactive confirmation prompt for STUDY creation.")
    parser.add_argument("--update", action="store_true", help="Use operation column/notes as an update file: add/open rows create, update rows update existing users, close/disable rows disable users.")
    parser.add_argument("--confirm-update", default="", help="Pass /approve to skip the interactive confirmation prompt for update mode.")
    parser.add_argument("--direct-user-update", action="store_true", help="For a direct one-user/manual edit request, treat unrecognized non-empty Notes as an update instead of skipping.")
    parser.add_argument("--username", default="", help="STUDY admin username. Defaults to the built-in production admin login.")
    parser.add_argument("--admin-username-env", default="", help="Environment variable containing an environment-specific STUDY admin username.")
    parser.add_argument("--admin-password-env", default="", help="Environment variable containing an environment-specific STUDY admin password.")
    parser.add_argument("--project-url", default="", help="Any URL from the target project, such as the STUDY login page.")
    parser.add_argument("--project-path", default="", help="Project route under study-api.example.invalid, for example demo900.")
    parser.add_argument("--project-id", default="", help="ProjectID/prefix used for generated usernames, for example TestProduct.")
    parser.add_argument("--study_api-base-url", default="", help="Advanced override for STUDY backend URL.")
    parser.add_argument("--study_api-login-page", default="", help="Advanced override for STUDY login page URL.")
    parser.add_argument("--only-roles", default="", help="Only process user rows whose role is in this comma-separated list, for example PM,MM,DM.")
    parser.add_argument("--skip-blocked-rows", action="store_true", help="Skip rows that fail preflight and continue with independently valid rows. Intended for validated email batch processing.")
    parser.add_argument("--require-operation", action="store_true", help="Require every actionable user row to include an account operation value. Intended for email application templates.")
    args = parser.parse_args()
    workbook_project_details = extract_project_details_from_excels(args.excel_file)
    apply_project_details_to_args(args, workbook_project_details)
    configure_study_runtime(args, require_api=not args.parse_only)

    print("STUDY study setup helper")
    print("Default mode is preview only. It will not create/update/disable unless you add the matching confirmation flag.\n")
    if workbook_project_details:
        shown = []
        if workbook_project_details.get("project_path"):
            shown.append(f"project path={workbook_project_details['project_path']}")
        if workbook_project_details.get("project_id"):
            shown.append(f"ProjectID={workbook_project_details['project_id']}")
        if shown:
            print("Workbook project info:", ", ".join(shown))

    sites, users = parse_all_excels(args.excel_file, args.site_sheet, args.user_sheet)
    if args.require_operation:
        for user in users:
            user["_requireOperation"] = True
    role_filter = parse_role_filter(args.only_roles)
    if role_filter:
        users = filter_users_by_roles(users, role_filter)
        print("Only processing user roles:", ", ".join(sorted(role_filter)))
    assign_run_usernames(users)
    print(f"\nExcel rows found across {len(args.excel_file)} file(s): {len(sites)} site(s), {len(users)} user(s)")

    if args.parse_only:
        print_parse_only(sites, users, direct_user_update=args.direct_user_update)
        print("\nParse-only mode. No login, lookup, or create happened.")
        return

    if bool(args.admin_username_env) != bool(args.admin_password_env):
        print("Both --admin-username-env and --admin-password-env are required for environment-specific login.")
        sys.exit(1)
    if args.admin_username_env:
        admin_username = normalize_cell(os.environ.get(args.admin_username_env, ""))
        admin_password = normalize_cell(os.environ.get(args.admin_password_env, ""))
    else:
        admin_username = normalize_cell(args.username) or UNIVERSAL_STUDY_ACCESS_USERNAME
        admin_password = UNIVERSAL_STUDY_ACCESS_PASSWORD
    if not admin_username:
        print("STUDY username was blank. Stop.")
        sys.exit(1)
    if not admin_password:
        print("Password was blank. Stop.")
        sys.exit(1)

    print("\nLogging in...")
    token, trial_id = login(admin_username, admin_password)

    project_scope_site_codes = {normalize_site_code(site.get("code", "")) for site in sites if normalize_site_code(site.get("code", ""))}
    if has_project_wide_site_request(users):
        existing_project_sites = lookup_all_project_sites(token, trial_id)
        project_scope_site_codes.update(
            normalize_site_code(site.get("code", ""))
            for site in existing_project_sites
            if normalize_site_code(site.get("code", "")) and internal_site_id(site)
        )
        print(f"Project-wide All scope resolved to {len(project_scope_site_codes)} site(s).")

    if args.update:
        default_update_action = "open" if args.direct_user_update else ""
        update_classes = classify_update_users(users, default_update_action)
        site_summary = summarize_sites(token, trial_id, sites)
        print_site_preview(site_summary)
        fail_if_user_blockers("STUDY site preview", site_summary, ["invalid", "conflicts"], args.skip_blocked_rows)

        pending_site_codes = {payload["code"] for payload in site_summary["valid_payloads"]}
        update_summary = prepare_update_actions(
            token,
            trial_id,
            users,
            pending_site_codes=pending_site_codes,
            project_scope_site_codes=project_scope_site_codes,
            default_unrecognized_action=default_update_action,
        )
        print_update_preview(update_summary, direct_user_update=args.direct_user_update)
        fail_if_user_blockers(
            "STUDY user preview",
            update_summary,
            [
                "invalid_open",
                "invalid_disable",
                "missing_role",
                "missing_site",
                "missing_update_user",
                "ambiguous_update_user",
                "missing_disable_user",
                "duplicates",
                "ambiguous",
            ],
            args.skip_blocked_rows,
        )

        print("\nEndpoints:")
        print("Create sites:", f"POST {BASE_URL}/trialapi/trials/{trial_id}/sites")
        print("Update sites:", f"PATCH {BASE_URL}/trialapi/trials/{trial_id}/sites/{{siteID}}")
        print("Create users:", f"POST {BASE_URL}/userapi/users")
        print("Update users:", f"PATCH {BASE_URL}/userapi/users/{{userID}}")
        print("Disable/reactivate users:", f"PATCH {BASE_URL}/userapi/users/{{userID}}/status")

        if not args.confirm_update:
            print("\nPreview only. Nothing was created, updated, reactivated, or disabled.")
            print("To run this update file, rerun with --update --confirm-update /approve.")
            return

        confirm = args.confirm_update or input("\nType /approve to create/update/reactivate/disable users from the update file: ")
        if not approval_accepted(confirm):
            print("Not updating. Confirmation did not match.")
            return

        print("\nCreating Excel-listed sites that are not already in STUDY first...")
        created_sites, failed_sites = create_sites(token, trial_id, site_summary["valid_payloads"])
        print(f"\nSite creation summary: created={created_sites}, failed={failed_sites}")

        failed_site_updates = 0
        failed_users = 0
        failed_updates = 0
        failed_reactivations = 0
        failed_disables = 0
        if site_summary["update_payloads"]:
            print("\nUpdating existing Excel-listed STUDY sites...")
            updated_sites, failed_site_updates = update_sites(token, trial_id, site_summary["update_payloads"])
            print(f"\nSite update summary: updated={updated_sites}, failed={failed_site_updates}")
        else:
            print("\nNo existing STUDY sites need updates.")

        print("\nRe-checking update actions after site creation...")
        update_summary = prepare_update_actions(
            token,
            trial_id,
            users,
            project_scope_site_codes=project_scope_site_codes,
            default_unrecognized_action=default_update_action,
        )
        print_update_preview(update_summary, direct_user_update=args.direct_user_update)
        fail_if_user_blockers(
            "STUDY user re-check",
            update_summary,
            [
                "invalid_open",
                "invalid_disable",
                "missing_role",
                "missing_site",
                "missing_update_user",
                "ambiguous_update_user",
                "missing_disable_user",
                "duplicates",
                "ambiguous",
            ],
            args.skip_blocked_rows,
        )

        if update_summary["to_create"]:
            print("\nCreating open/add users that do not already exist...")
            created_users, failed_users = create_users(token, update_summary["to_create"])
            print(f"\nOpen/add user creation summary: created={created_users}, failed={failed_users}")
        else:
            print("\nNo open/add users need STUDY creation.")

        if update_summary["to_update"]:
            print("\nUpdating existing open/add users from Excel...")
            updated_users, failed_updates = update_users(token, update_summary["to_update"])
            print(f"\nSMT user update summary: updated={updated_users}, failed={failed_updates}")
        else:
            print("\nNo existing open/add users need STUDY update.")

        if update_summary["to_reactivate"]:
            print("\nReactivating existing open/add users that are currently disabled...")
            reactivated_users, failed_reactivations = set_users_status(token, update_summary["to_reactivate"], 1, "Reactivate")
            print(f"\nSMT user reactivation summary: reactivated={reactivated_users}, failed={failed_reactivations}")
        else:
            print("\nNo existing open/add users need reactivation.")

        if update_summary["disable_ready"]:
            print("\nDisabling users marked close/disable in Notes...")
            disabled_users, failed_disables = set_users_status(token, update_summary["disable_ready"], 0, "Disable")
            print(f"\nSMT user disable summary: disabled={disabled_users}, failed={failed_disables}")
        else:
            print("\nNo STUDY users need disabling.")
        fail_if_write_failures(
            "STUDY update",
            site_create=failed_sites,
            site_update=failed_site_updates,
            user_create=failed_users,
            user_update=failed_updates,
            user_reactivate=failed_reactivations,
            user_disable=failed_disables,
        )
        return

    site_summary = summarize_sites(token, trial_id, sites)
    print_site_preview(site_summary)
    fail_if_user_blockers("STUDY site preview", site_summary, ["invalid", "conflicts"], args.skip_blocked_rows)

    pending_site_codes = {payload["code"] for payload in site_summary["valid_payloads"]}

    user_summary = prepare_users(
        token,
        trial_id,
        users,
        pending_site_codes=pending_site_codes,
        project_scope_site_codes=project_scope_site_codes,
    )
    print_user_preview(user_summary)
    fail_if_user_blockers(
        "STUDY user preview",
        user_summary,
        ["invalid", "missing_role", "missing_site", "duplicates"],
        args.skip_blocked_rows,
    )

    print("\nEndpoints:")
    print("Sites:", f"POST {BASE_URL}/trialapi/trials/{trial_id}/sites")
    print("Update sites:", f"PATCH {BASE_URL}/trialapi/trials/{trial_id}/sites/{{siteID}}")
    print("Users:", f"POST {BASE_URL}/userapi/users")

    if not args.create:
        print("\nPreview only. Nothing was created.")
        print("To create valid Excel-listed sites/users that are not already in STUDY, rerun with --create --confirm-create /approve.")
        return

    confirm = args.confirm_create or input("\nType /approve to create valid Excel-listed sites/users that are not already in STUDY: ")
    if not approval_accepted(confirm):
        print("Not creating. Confirmation did not match.")
        return

    print("\nCreating Excel-listed sites that are not already in STUDY first...")
    created_sites, failed_sites = create_sites(token, trial_id, site_summary["valid_payloads"])
    print(f"\nSite creation summary: created={created_sites}, failed={failed_sites}")

    failed_site_updates = 0
    if site_summary["update_payloads"]:
        print("\nUpdating existing Excel-listed STUDY sites...")
        updated_sites, failed_site_updates = update_sites(token, trial_id, site_summary["update_payloads"])
        print(f"\nSite update summary: updated={updated_sites}, failed={failed_site_updates}")
    else:
        print("\nNo existing STUDY sites need updates.")

    print("\nRe-checking users after site creation...")
    user_summary = prepare_users(token, trial_id, users, project_scope_site_codes=project_scope_site_codes)
    print_user_preview(user_summary)
    fail_if_user_blockers(
        "STUDY user re-check",
        user_summary,
        ["invalid", "missing_role", "missing_site", "duplicates"],
        args.skip_blocked_rows,
    )

    if not user_summary["prepared"]:
        print("\nNo valid users to create after site creation.")
        fail_if_write_failures(
            "STUDY creation",
            site_create=failed_sites,
            site_update=failed_site_updates,
        )
        return

    print("\nCreating users...")
    created_users, failed_users = create_users(token, user_summary["prepared"])
    print(f"\nUser creation summary: created={created_users}, failed={failed_users}")
    fail_if_write_failures(
        "STUDY creation",
        site_create=failed_sites,
        site_update=failed_site_updates,
        user_create=failed_users,
    )


if __name__ == "__main__":
    main()
