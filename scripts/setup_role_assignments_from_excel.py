#!/usr/bin/env python3
import argparse
import json
import os
import re
import ssl
import sys
import urllib.error
import urllib.parse
import urllib.request

from setup_study_from_excel import (
    DEFAULT_HOST,
    SITE_ROLES,
    approval_accepted,
    assign_run_usernames,
    apply_project_details_to_args,
    build_username,
    classify_update_users,
    configure_project_id,
    extract_project_path,
    extract_project_details_from_excels,
    first_list,
    generated_username_candidates,
    load_env_file,
    normalize_cell,
    normalize_username,
    parse_all_excel_users,
    redact,
    UNIVERSAL_STUDY_ACCESS_PASSWORD,
    UNIVERSAL_STUDY_ACCESS_USERNAME,
    validate_user,
    validate_user_identity,
)


ROLE_BASE_URL = ""
ROLE_LOGIN_PAGE = ""


def configure_role_runtime(args, require_api=True):
    global ROLE_BASE_URL, ROLE_LOGIN_PAGE

    project_input = (
        getattr(args, "project_url", "")
        or getattr(args, "project_path", "")
        or os.environ.get("STUDY_ACCESS_PROJECT_URL", "")
        or os.environ.get("STUDY_ACCESS_PROJECT_PATH", "")
    )
    project_path = extract_project_path(project_input)

    role_base_url = (
        getattr(args, "role_base_url", "")
        or os.environ.get("ROLE_BASE_URL", "")
    ).rstrip("/")
    role_login_page = (
        getattr(args, "role_login_page", "")
        or os.environ.get("ROLE_LOGIN_PAGE", "")
    ).rstrip("/")

    if require_api and not project_path and not (role_base_url and role_login_page):
        project_input = input(
            "Project URL or path for this study, for example "
            "https://role-api.example.invalid/role/login: "
        ).strip()
        project_path = extract_project_path(project_input)

    if not role_base_url and project_path:
        role_base_url = f"{DEFAULT_HOST}/{project_path}/b-report-management-backend"
    if not role_login_page and project_path:
        role_login_page = f"{DEFAULT_HOST}/{project_path}/f-report-management-tool/login"

    if require_api and (not role_base_url or not role_login_page):
        print("ROLE project URL/path is required. Provide --project-url, --project-path, or explicit ROLE URLs.")
        sys.exit(1)

    ROLE_BASE_URL = role_base_url
    ROLE_LOGIN_PAGE = role_login_page
    configure_project_id(getattr(args, "project_id", ""), project_path, required=True)
    return project_path


def parse_json(text):
    try:
        return json.loads(text)
    except Exception:
        return {"raw": text}


def request_json(method, path, body=None, token=None):
    url = ROLE_BASE_URL + path
    data = None if body is None else json.dumps(body, ensure_ascii=False).encode("utf-8")
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/json;charset=UTF-8",
        "Origin": "https://role-api.example.invalid",
        "Referer": ROLE_LOGIN_PAGE,
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
    print("ROLE login status:", status)
    print("ROLE login code:", data.get("code"))
    if data.get("code") != "SUCCESS":
        print("ROLE login failed:")
        print(redact(raw[:800]))
        sys.exit(1)

    payload = data.get("payload") or {}
    token = payload.get("token")
    trial_id = extract_trial_id(payload)
    if not token or not trial_id:
        print("Cannot continue because ROLE login did not return token/trialID.")
        sys.exit(1)
    return token, trial_id


def parse_role_map(items):
    mapping = {}
    for item in items or []:
        if "=" not in item:
            raise ValueError(f"Invalid --role-map value: {item}. Use EXCEL_ROLE=ROLE_ROLE.")
        left, right = item.split("=", 1)
        mapping[left.strip().upper()] = right.strip()
    return mapping


def parse_role_filter(value):
    raw = normalize_cell(value)
    if not raw:
        return set()
    return {part.strip().upper() for part in re.split(r"[,，、;；\s]+", raw) if part.strip()}


def filter_users_by_roles(users, role_filter):
    if not role_filter:
        return users
    return [user for user in users if normalize_cell(user.get("role")).upper() in role_filter]


def role_base_from_excel(role_name):
    return 0 if role_name.strip().upper() in SITE_ROLES else 1


def lookup_role_assignments(token):
    status, data, raw = request_json("GET", "/userapi/roles?page=1&size=500&countOnly=true", token=token)
    print("ROLE role lookup status:", status)
    print("ROLE role lookup code:", data.get("code"))
    roles = first_list(data, ["roles", "data", "list", "items", "records"])
    if not roles:
        print("Could not read ROLE role list.")
        print("Response preview:", redact(raw[:1000]))
        sys.exit(1)
    return roles


def role_name_candidates(role):
    return [
        str(role.get("name", "")),
        str(role.get("code", "")),
        str(role.get("roleName", "")),
        str(role.get("displayName", "")),
    ]


def matching_role_assignments(roles, excel_role_name, role_map):
    excel_role = excel_role_name.strip().upper()
    wanted = role_map.get(excel_role, excel_role).strip().upper()

    matches = []
    for role in roles:
        names = [name.strip().upper() for name in role_name_candidates(role) if name.strip()]
        if wanted in names:
            matches.append(role)
    return matches


def find_role_role(roles, excel_role_name, role_map):
    excel_role = excel_role_name.strip().upper()
    wanted_base_role = role_base_from_excel(excel_role)
    matches = matching_role_assignments(roles, excel_role_name, role_map)

    active_matches = [role for role in matches if role.get("status", 1) == 1]
    base_matches = [role for role in active_matches if normalize_base_role(role.get("baseRole")) in (None, wanted_base_role)]
    chosen = (base_matches or [None])[0]
    if not chosen:
        return None

    return {
        "id": chosen.get("_id") or chosen.get("id"),
        "name": chosen.get("name") or chosen.get("roleName") or chosen.get("displayName") or excel_role,
        "baseRole": chosen.get("baseRole"),
    }


def normalize_base_role(value):
    if value is None or value == "":
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return value


def query_role_users(token, username, assign_type, page=1, size=20):
    query = urllib.parse.urlencode({
        "siteID": "",
        "username": username,
        "phone": "",
        "assignType": assign_type,
        "userType": "",
        "page": page,
        "size": size,
        "countOnly": "true",
    })
    status, data, raw = request_json("GET", f"/userapi/users?{query}", token=token)
    users = first_list(data, ["users", "data", "list", "items", "records"])
    return status, data.get("code"), users, raw


def lookup_role_user_id(token, username):
    path = "/userapi/users/null?userName=" + urllib.parse.quote(username)
    status, data, raw = request_json("GET", path, token=token)
    payload = data.get("payload") if isinstance(data, dict) else None
    if status == 200 and data.get("code") == "SUCCESS" and isinstance(payload, dict) and payload.get("_id"):
        returned_username = normalize_username(payload.get("username") or payload.get("userName") or "")
        if returned_username and returned_username != username:
            return None
        return {
            "_id": payload.get("_id"),
            "username": returned_username or username,
            "email": payload.get("email"),
            "source": "username lookup",
        }
    return None


def find_role_user(token, username):
    for assign_type in ("isAssign", "notAssign"):
        status, code, users, raw = query_role_users(token, username, assign_type)
        if status != 200 or code != "SUCCESS":
            continue
        exact = [user for user in users if normalize_username(user.get("username")) == username]
        if exact:
            user = exact[0]
            user["source"] = assign_type
            return user
    return lookup_role_user_id(token, username)


def resolve_role_user(token, user, allow_legacy=False):
    """Resolve one exact ROLE user, optionally checking deterministic legacy formats."""
    expected_username = build_username(user)
    candidates = generated_username_candidates(user) if allow_legacy else [expected_username]
    matches = []
    for candidate in candidates:
        role_user = find_role_user(token, candidate)
        if role_user and role_user.get("_id"):
            matches.append((candidate, role_user))

    unique_matches = []
    seen = set()
    for candidate, role_user in matches:
        key = role_user.get("_id") or normalize_username(role_user.get("username"))
        if not key or key in seen:
            continue
        seen.add(key)
        unique_matches.append((candidate, role_user))

    if len(unique_matches) == 1:
        candidate, role_user = unique_matches[0]
        actual_username = normalize_username(role_user.get("username")) or candidate
        match_type = "exact" if actual_username == expected_username else "legacy"
        return actual_username, role_user, match_type, []
    if len(unique_matches) > 1:
        return expected_username, None, "ambiguous", [item for _candidate, item in unique_matches]
    return expected_username, None, "missing", []


def current_role_id(user):
    role = user.get("userRole")
    if isinstance(role, dict):
        return role.get("_id") or role.get("id")
    if isinstance(role, str):
        return role
    return None


def current_role_name(user):
    role = user.get("userRole")
    if isinstance(role, dict):
        return role.get("name") or role.get("roleName") or role.get("displayName") or ""
    return ""


def unique_users_by_username(users):
    unique = []
    seen = set()
    duplicates = []
    for user in users:
        username = build_username(user)
        if username in seen:
            duplicates.append((user, username))
            continue
        seen.add(username)
        unique.append(user)
    return unique, duplicates


def prepare_assignments(token, users, roles, role_map, direct_update_mode=False, legacy_user_ids=None):
    prepared = []
    invalid = []
    missing_role = []
    missing_user = []
    no_change = []
    duplicates = []
    role_base_mismatch = []
    user_role_mismatch = []
    ambiguous_user = []
    seen_usernames = set()
    legacy_user_ids = set(legacy_user_ids or [])

    for user in users:
        if direct_update_mode:
            missing_fields = validate_user_identity(user)
            if not normalize_cell(user.get("role")):
                missing_fields.append("role")
        else:
            missing_fields = validate_user(user)
        if missing_fields:
            invalid.append((user, missing_fields))
            continue

        excel_role = normalize_cell(user.get("role")).upper()
        role_role = find_role_role(roles, excel_role, role_map)
        if not role_role or not role_role.get("id"):
            expected_base_role = role_base_from_excel(excel_role)
            wrong_base_matches = [
                role for role in matching_role_assignments(roles, excel_role, role_map)
                if normalize_base_role(role.get("baseRole")) not in (None, expected_base_role)
            ]
            if wrong_base_matches:
                role_base_mismatch.append((user, expected_base_role, wrong_base_matches))
                continue
            missing_role.append(user)
            continue

        username, role_user, match_type, ambiguous_matches = resolve_role_user(
            token,
            user,
            allow_legacy=id(user) in legacy_user_ids,
        )
        if ambiguous_matches:
            ambiguous_user.append((user, build_username(user), ambiguous_matches))
            continue
        if username in seen_usernames:
            duplicates.append((user, username))
            continue
        seen_usernames.add(username)
        if not role_user or not role_user.get("_id"):
            missing_user.append((user, username))
            continue

        current_id = current_role_id(role_user)
        role_user_base_role = normalize_base_role(role_user.get("role"))
        role_role_base_role = normalize_base_role(role_role.get("baseRole"))
        item = {
            "excelRow": user.get("excelRow"),
            "name": user.get("name"),
            "username": username,
            "userID": role_user.get("_id"),
            "excelRole": excel_role,
            "rmtUserBaseRole": role_user_base_role,
            "rmtRoleID": role_role["id"],
            "rmtRoleName": role_role["name"],
            "rmtRoleBaseRole": role_role_base_role,
            "currentRoleID": current_id,
            "currentRoleName": current_role_name(role_user),
            "source": role_user.get("source", ""),
            "matchType": match_type,
        }
        if role_user_base_role is not None and role_role_base_role is not None and role_user_base_role != role_role_base_role:
            user_role_mismatch.append(item)
            continue
        if current_id and current_id == role_role["id"]:
            no_change.append(item)
        else:
            prepared.append(item)

    return {
        "prepared": prepared,
        "invalid": invalid,
        "missing_role": missing_role,
        "missing_user": missing_user,
        "duplicates": duplicates,
        "no_change": no_change,
        "role_base_mismatch": role_base_mismatch,
        "user_role_mismatch": user_role_mismatch,
        "ambiguous_user": ambiguous_user,
    }


def prepare_update_actions(token, users, roles, role_map, default_unrecognized_action=""):
    classes = classify_update_users(users, default_unrecognized_action)
    direct_update_mode = default_unrecognized_action == "open"
    update_only_ids = {id(user) for user in classes.get("update_only", [])}
    assignment_summary = prepare_assignments(
        token,
        classes["open"],
        roles,
        role_map,
        direct_update_mode,
        legacy_user_ids=update_only_ids,
    )

    disable_ready = []
    disable_not_found = []
    invalid_disable = []
    duplicate_disable = []
    ambiguous_disable_user = []
    seen_disable_usernames = set()

    for user in classes["disable"]:
        missing_fields = validate_user_identity(user)
        if missing_fields:
            invalid_disable.append((user, missing_fields))
            continue

        username, role_user, _match_type, ambiguous_matches = resolve_role_user(token, user, allow_legacy=True)
        if ambiguous_matches:
            ambiguous_disable_user.append((user, build_username(user), ambiguous_matches))
            continue
        if username in seen_disable_usernames:
            duplicate_disable.append((user, username))
            continue
        seen_disable_usernames.add(username)

        if not role_user or not role_user.get("_id"):
            disable_not_found.append((user, username))
            continue
        disable_ready.append({
            "excelRow": user.get("excelRow"),
            "name": user.get("name"),
            "username": username,
            "userID": role_user.get("_id"),
            "source": role_user.get("source", ""),
        })

    return {
        "assignments": assignment_summary,
        "disable_ready": disable_ready,
        "disable_not_found": disable_not_found,
        "invalid_disable": invalid_disable,
        "duplicate_disable": duplicate_disable,
        "ambiguous_disable_user": ambiguous_disable_user,
        "open_rows": classes["open"],
        "update_only_rows": classes.get("update_only", []),
        "disable_rows": classes["disable"],
        "ambiguous": classes["ambiguous"],
        "ignored": classes["ignored"],
    }


def print_role_list(roles):
    print("\nAvailable ROLE roles:")
    for role in roles:
        print(f"- {role.get('name')} | id={role.get('_id') or role.get('id')} | baseRole={role.get('baseRole')} | status={role.get('status')}")


def print_preview(summary, roles):
    print("\nRMT assignment preview")
    print(f"- users ready for ROLE role assignment: {len(summary['prepared'])}")
    print(f"- users already assigned to requested role: {len(summary['no_change'])}")
    print(f"- duplicate generated usernames skipped: {len(summary['duplicates'])}")
    print(f"- user rows missing required fields: {len(summary['invalid'])}")
    print(f"- user rows with unknown ROLE role: {len(summary['missing_role'])}")
    print(f"- user rows where ROLE role exists but has wrong type: {len(summary['role_base_mismatch'])}")
    print(f"- user rows where ROLE user type does not match target role type: {len(summary['user_role_mismatch'])}")
    print(f"- generated usernames not found in ROLE: {len(summary['missing_user'])}")
    print(f"- update rows with multiple historical username matches: {len(summary['ambiguous_user'])}")

    if summary["prepared"]:
        print("\nAssignments ready:")
        for item in summary["prepared"][:100]:
            current = item.get("currentRoleName") or item.get("currentRoleID") or "not assigned/unknown"
            print(f"- row {item['excelRow']}: {item['username']} {current} -> {item['rmtRoleName']} ({item['source']})")

    if summary["no_change"]:
        print("\nAlready correct:")
        for item in summary["no_change"][:40]:
            print(f"- row {item['excelRow']}: {item['username']} already has {item['rmtRoleName']}")

    if summary["duplicates"]:
        print("\nDuplicate generated usernames skipped:")
        for user, username in summary["duplicates"][:40]:
            print(f"- row {user.get('excelRow')}: {username}")

    if summary["invalid"]:
        print("\nRows not ready:")
        for user, missing in summary["invalid"][:60]:
            print(f"- row {user.get('excelRow')}, name {user.get('name') or ''}: missing {', '.join(missing)}")

    if summary["missing_role"]:
        print("\nRoles not found in ROLE:")
        for user in summary["missing_role"][:60]:
            print(f"- row {user.get('excelRow')}: Excel role {user.get('role')}")
        print_role_list(roles)

    if summary["role_base_mismatch"]:
        print("\nRMT roles found by name, but with the wrong type:")
        for user, expected_base_role, matches in summary["role_base_mismatch"][:60]:
            found = ", ".join(
                f"{role.get('name')} baseRole={role.get('baseRole')}"
                for role in matches[:8]
            )
            print(f"- row {user.get('excelRow')}: Excel role {user.get('role')} needs baseRole={expected_base_role}, found {found}")
        print("\nFor these rows, create/map to an ROLE role with the required baseRole, or remove those rows from ROLE assignment.")
        print_role_list(roles)

    if summary["user_role_mismatch"]:
        print("\nRMT user type does not match target role type:")
        for item in summary["user_role_mismatch"][:60]:
            print(f"- row {item['excelRow']}: {item['username']} user role={item['rmtUserBaseRole']} but target role {item['rmtRoleName']} baseRole={item['rmtRoleBaseRole']}")
        print("\nThese rows are blocked because the existing ROLE/STUDY user account was created with a different account type than the target ROLE role.")
        print("Do not rerun assignment for these rows until the user accounts are fixed or recreated with the correct account type.")

    if summary["missing_user"]:
        print("\nUsers not found in ROLE by generated username:")
        for user, username in summary["missing_user"][:80]:
            print(f"- row {user.get('excelRow')}: expected username {username}, name {user.get('name') or ''}")

    if summary["ambiguous_user"]:
        print("\nUpdate rows with multiple exact historical username matches:")
        for user, username, matches in summary["ambiguous_user"][:80]:
            found = ", ".join(normalize_username(item.get("username")) for item in matches[:8])
            print(f"- row {user.get('excelRow')}: expected username {username}, found {found}")


def print_update_preview(summary, roles, direct_user_update=False):
    assignments = summary["assignments"]
    if direct_user_update:
        print("\nRMT direct user role update preview")
    else:
        print("\nRMT update preview from operation column/notes")
    print(f"- rows treated as open/add/update/reactivate: {len(summary['open_rows'])}")
    print(f"- rows explicitly marked update-only: {len(summary.get('update_only_rows', []))}")
    print(f"- open/add rows ready for ROLE role assignment: {len(assignments['prepared'])}")
    print(f"- open/add rows already assigned to requested role: {len(assignments['no_change'])}")
    print(f"- rows marked close/disable: {len(summary['disable_rows'])}")
    print(f"- existing ROLE users ready to disable/remove from ROLE login: {len(summary['disable_ready'])}")
    print(f"- disable rows not found in ROLE: {len(summary['disable_not_found'])}")
    print(f"- disable rows missing username identity fields: {len(summary['invalid_disable'])}")
    print(f"- duplicate disable usernames skipped: {len(summary['duplicate_disable'])}")
    print(f"- disable rows with multiple historical username matches: {len(summary['ambiguous_disable_user'])}")
    print(f"- ambiguous operation/note values skipped: {len(summary['ambiguous'])}")
    print(f"- rows with no recognized operation/update marker skipped: {len(summary['ignored'])}")
    print(f"- open/add duplicate generated usernames skipped: {len(assignments['duplicates'])}")
    print(f"- open/add rows missing required fields: {len(assignments['invalid'])}")
    print(f"- open/add rows with unknown ROLE role: {len(assignments['missing_role'])}")
    print(f"- open/add rows where ROLE role exists but has wrong type: {len(assignments['role_base_mismatch'])}")
    print(f"- open/add rows where ROLE user type does not match target role type: {len(assignments['user_role_mismatch'])}")
    print(f"- open/add generated usernames not found in ROLE: {len(assignments['missing_user'])}")
    print(f"- update rows with multiple historical username matches: {len(assignments['ambiguous_user'])}")

    if assignments["prepared"]:
        print("\nRMT assignments ready:")
        for item in assignments["prepared"][:100]:
            current = item.get("currentRoleName") or item.get("currentRoleID") or "not assigned/unknown"
            print(f"- row {item['excelRow']}: {item['username']} {current} -> {item['rmtRoleName']} ({item['source']})")

    if assignments["no_change"]:
        print("\nRMT roles already correct:")
        for item in assignments["no_change"][:40]:
            print(f"- row {item['excelRow']}: {item['username']} already has {item['rmtRoleName']}")

    if summary["disable_ready"]:
        print("\nRMT users ready to disable/remove:")
        for item in summary["disable_ready"][:80]:
            print(f"- row {item['excelRow']}: {item['username']} ({item['source']})")

    if summary["disable_not_found"]:
        print("\nDisable rows not found in ROLE:")
        for user, username in summary["disable_not_found"][:80]:
            print(f"- row {user.get('excelRow')}: expected username {username}, name {user.get('name') or ''}")

    if summary["invalid_disable"]:
        print("\nDisable rows that cannot identify a user yet:")
        for user, missing in summary["invalid_disable"][:80]:
            print(f"- row {user.get('excelRow')}, name {user.get('name') or ''}: missing {', '.join(missing)}")

    if summary["duplicate_disable"]:
        print("\nDuplicate disable usernames skipped:")
        for user, username in summary["duplicate_disable"][:80]:
            print(f"- row {user.get('excelRow')}: {username}")

    if summary["ambiguous_disable_user"]:
        print("\nDisable rows with multiple exact historical username matches:")
        for user, username, matches in summary["ambiguous_disable_user"][:80]:
            found = ", ".join(normalize_username(item.get("username")) for item in matches[:8])
            print(f"- row {user.get('excelRow')}: expected username {username}, found {found}")

    if summary["ambiguous"]:
        print("\nAmbiguous operation/note values skipped:")
        for user in summary["ambiguous"][:50]:
            detail = user.get("operation") or user.get("note") or ""
            print(f"- row {user.get('excelRow')}: operation/note={detail}")

    if assignments["invalid"]:
        print("\nOpen/add rows not ready for ROLE assignment:")
        for user, missing in assignments["invalid"][:60]:
            print(f"- row {user.get('excelRow')}, name {user.get('name') or ''}: missing {', '.join(missing)}")

    if assignments["missing_role"]:
        print("\nOpen/add roles not found in ROLE:")
        for user in assignments["missing_role"][:60]:
            print(f"- row {user.get('excelRow')}: Excel role {user.get('role')}")
        print_role_list(roles)

    if assignments["role_base_mismatch"]:
        print("\nOpen/add ROLE roles found by name, but with the wrong type:")
        for user, expected_base_role, matches in assignments["role_base_mismatch"][:60]:
            found = ", ".join(
                f"{role.get('name')} baseRole={role.get('baseRole')}"
                for role in matches[:8]
            )
            print(f"- row {user.get('excelRow')}: Excel role {user.get('role')} needs baseRole={expected_base_role}, found {found}")
        print_role_list(roles)

    if assignments["user_role_mismatch"]:
        print("\nOpen/add ROLE user type does not match target role type:")
        for item in assignments["user_role_mismatch"][:60]:
            print(f"- row {item['excelRow']}: {item['username']} user role={item['rmtUserBaseRole']} but target role {item['rmtRoleName']} baseRole={item['rmtRoleBaseRole']}")

    if assignments["missing_user"]:
        print("\nOpen/add users not found in ROLE by generated username:")
        for user, username in assignments["missing_user"][:80]:
            print(f"- row {user.get('excelRow')}: expected username {username}, name {user.get('name') or ''}")

    if assignments["ambiguous_user"]:
        print("\nUpdate rows with multiple exact historical username matches:")
        for user, username, matches in assignments["ambiguous_user"][:80]:
            found = ", ".join(normalize_username(item.get("username")) for item in matches[:8])
            print(f"- row {user.get('excelRow')}: expected username {username}, found {found}")


def print_parse_only(users):
    invalid = [(user, validate_user(user)) for user in users if validate_user(user)]
    ready = [user for user in users if not validate_user(user)]
    update_classes = classify_update_users(users)
    print("\nRMT parse-only check")
    print(f"- user rows found: {len(users)}")
    print(f"- user rows ready for username generation: {len(ready)}")
    print(f"- user rows missing required fields: {len(invalid)}")
    print(f"- update rows treated as open/add/update/reactivate: {len(update_classes['open'])}")
    print(f"- update rows explicitly marked update-only: {len(update_classes.get('update_only', []))}")
    print(f"- update rows marked close/disable: {len(update_classes['disable'])}")
    print(f"- update rows with ambiguous operation/note values: {len(update_classes['ambiguous'])}")
    print(f"- rows with no recognized operation/update marker: {len(update_classes['ignored'])}")

    if ready:
        print("\nGenerated username examples:")
        for user in ready[:30]:
            print(f"- row {user.get('excelRow')}: {build_username(user)} role={user.get('role')}")

    if invalid:
        print("\nRows not ready:")
        for user, missing in invalid[:40]:
            print(f"- row {user.get('excelRow')}, name {user.get('name') or ''}: missing {', '.join(missing)}")


def assign_role_role(token, item):
    path = "/userapi/localUsers/" + urllib.parse.quote(item["userID"])
    body = {"userID": item["userID"], "userRole": item["rmtRoleID"]}
    status, data, raw = request_json("PATCH", path, body, token=token)
    return status, data, raw


def assign_roles(token, assignments):
    assigned = 0
    failed = 0
    for index, item in enumerate(assignments, start=1):
        print(f"\nAssigning ROLE role {index}/{len(assignments)}: {item['username']} -> {item['rmtRoleName']}")
        status, data, raw = assign_role_role(token, item)
        print("Assign status:", status)
        print("Assign code:", data.get("code"))
        print("Assign response:", redact(raw[:1000]))
        if status in {200, 201} and data.get("code") == "SUCCESS":
            assigned += 1
        else:
            failed += 1
    return assigned, failed


def delete_role_user(token, item):
    path = "/userapi/users/" + urllib.parse.quote(item["userID"])
    status, data, raw = request_json("DELETE", path, token=token)
    return status, data, raw


def delete_role_users(token, items):
    disabled = 0
    failed = 0
    for index, item in enumerate(items, start=1):
        print(f"\nDisabling/removing ROLE user {index}/{len(items)}: {item['username']}")
        status, data, raw = delete_role_user(token, item)
        print("ROLE disable status:", status)
        print("ROLE disable code:", data.get("code"))
        print("ROLE disable response:", redact(raw[:1000]))
        if status in {200, 201} and data.get("code") == "SUCCESS":
            disabled += 1
        else:
            failed += 1
    return disabled, failed


def assignment_blocker_count(summary, skip_blocked_rows=False):
    if skip_blocked_rows:
        return 0
    return sum(
        len(summary[key])
        for key in (
            "invalid",
            "missing_role",
            "missing_user",
            "duplicates",
            "role_base_mismatch",
            "user_role_mismatch",
            "ambiguous_user",
        )
    )


def update_blocker_count(summary, skip_blocked_rows=False):
    if skip_blocked_rows:
        return 0
    return (
        assignment_blocker_count(summary["assignments"])
        + len(summary["invalid_disable"])
        + len(summary["ambiguous_disable_user"])
        + len(summary["ambiguous"])
    )


def main():
    load_env_file(os.path.expanduser("~/.hermes/.env"))
    load_env_file(os.path.join(os.getcwd(), ".env"))

    parser = argparse.ArgumentParser(description="Preview or assign ROLE user-tab roles from one or more Excel workbooks.")
    parser.add_argument("--excel-file", required=True, nargs="+", help="Path(s) to filled Excel workbook(s).")
    parser.add_argument("--user-sheet", default="", help="Optional user sheet name. Defaults to scan/prefer 研究人员信息.")
    parser.add_argument("--parse-only", action="store_true", help="Only read Excel and show generated usernames. Do not log in or call ROLE APIs.")
    parser.add_argument("--assign", action="store_true", help="Actually assign ROLE roles. Without this, preview only.")
    parser.add_argument("--confirm-assign", default="", help="Pass /approve to skip the interactive confirmation prompt for ROLE assignment.")
    parser.add_argument("--update", action="store_true", help="Use operation column/notes as an update file: add/update rows assign roles; close/disable rows disable/remove ROLE users.")
    parser.add_argument("--confirm-update", default="", help="Pass /approve to skip the interactive confirmation prompt for update mode.")
    parser.add_argument("--direct-user-update", action="store_true", help="For a direct one-user/manual edit request, treat unrecognized non-empty Notes as an update instead of skipping.")
    parser.add_argument("--username", default="", help="ROLE admin username. Defaults to the built-in production admin login.")
    parser.add_argument("--admin-username-env", default="", help="Environment variable containing an environment-specific ROLE admin username.")
    parser.add_argument("--admin-password-env", default="", help="Environment variable containing an environment-specific ROLE admin password.")
    parser.add_argument("--project-url", default="", help="Any URL from the target project, such as the ROLE login page.")
    parser.add_argument("--project-path", default="", help="Project route under study-api.example.invalid, for example demo900.")
    parser.add_argument("--project-id", default="", help="ProjectID/prefix used for generated usernames, for example TestProduct.")
    parser.add_argument("--role_api-base-url", default="", help="Advanced override for ROLE backend URL.")
    parser.add_argument("--role_api-login-page", default="", help="Advanced override for ROLE login page URL.")
    parser.add_argument("--role-map", action="append", default=[], help="Map Excel role to ROLE role, for example --role-map CRC=RT.")
    parser.add_argument("--only-roles", default="", help="Only process user rows whose role is in this comma-separated list, for example PM,MM,DM.")
    parser.add_argument("--skip-blocked-rows", action="store_true", help="Skip rows that fail preflight and continue with independently valid rows. Intended for validated email batch processing.")
    parser.add_argument("--require-operation", action="store_true", help="Require every actionable user row to include an account operation value. Intended for email application templates.")
    args = parser.parse_args()
    workbook_project_details = extract_project_details_from_excels(args.excel_file)
    apply_project_details_to_args(args, workbook_project_details)
    configure_role_runtime(args, require_api=not args.parse_only)

    print("ROLE role assignment helper")
    print("Default mode is preview only. It will not assign or disable unless you add the matching confirmation flag.\n")
    if workbook_project_details:
        shown = []
        if workbook_project_details.get("project_path"):
            shown.append(f"project path={workbook_project_details['project_path']}")
        if workbook_project_details.get("project_id"):
            shown.append(f"ProjectID={workbook_project_details['project_id']}")
        if shown:
            print("Workbook project info:", ", ".join(shown))

    users = parse_all_excel_users(args.excel_file, args.user_sheet)
    if args.require_operation:
        for user in users:
            user["_requireOperation"] = True
    role_filter = parse_role_filter(args.only_roles)
    if role_filter:
        users = filter_users_by_roles(users, role_filter)
        print("Only processing user roles:", ", ".join(sorted(role_filter)))
    assign_run_usernames(users)
    print(f"\nExcel user rows found across {len(args.excel_file)} file(s): {len(users)}")

    if args.parse_only:
        print_parse_only(users)
        print("\nParse-only mode. No login, lookup, or ROLE assignment happened.")
        return

    role_map = parse_role_map(args.role_map)
    if role_map:
        print("Role map:", ", ".join(f"{key}->{value}" for key, value in role_map.items()))

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
        print("ROLE username was blank. Stop.")
        sys.exit(1)
    if not admin_password:
        print("Password was blank. Stop.")
        sys.exit(1)

    print("\nLogging in to ROLE...")
    token, trial_id = login(admin_username, admin_password)
    print("ROLE trial ID:", trial_id)

    roles = lookup_role_assignments(token)

    if args.update:
        default_update_action = "open" if args.direct_user_update else ""
        summary = prepare_update_actions(token, users, roles, role_map, default_update_action)
        print_update_preview(summary, roles, direct_user_update=args.direct_user_update)

        print("\nEndpoints:")
        print("ROLE role assignment:", f"PATCH {ROLE_BASE_URL}/userapi/localUsers/{{userID}}")
        print("ROLE user disable/remove:", f"DELETE {ROLE_BASE_URL}/userapi/users/{{userID}}")

        if not args.confirm_update:
            print("\nPreview only. Nothing was assigned or disabled in ROLE.")
            print("To run this ROLE update file, rerun with --update --confirm-update /approve.")
            return

        blocked = update_blocker_count(summary, args.skip_blocked_rows)
        if blocked:
            print(
                "\nNo ROLE changes were made because one or more rows are blocked. "
                "Fix the blocked rows above, then rerun."
            )
            raise SystemExit(1)

        confirm = args.confirm_update or input("\nType /approve to assign open/add rows and disable close rows in ROLE: ")
        if not approval_accepted(confirm):
            print("Not updating ROLE. Confirmation did not match.")
            return

        failed_disabled = 0
        failed = 0
        if summary["disable_ready"]:
            disabled, failed_disabled = delete_role_users(token, summary["disable_ready"])
            print(f"\nRMT disable/remove summary: disabled={disabled}, failed={failed_disabled}")
        else:
            print("\nNo ROLE users need disabling/removal.")

        assignments = summary["assignments"]
        if assignments["prepared"]:
            assigned, failed = assign_roles(token, assignments["prepared"])
            print(f"\nRMT role assignment summary: assigned={assigned}, failed={failed}")
        else:
            print("\nNo ROLE role assignments to make for open/add rows.")
        if failed_disabled or failed:
            print(
                "\nRMT update failed: "
                f"disable_failures={failed_disabled}, assignment_failures={failed}"
            )
            raise SystemExit(1)
        return

    summary = prepare_assignments(token, users, roles, role_map)
    print_preview(summary, roles)

    print("\nEndpoint:")
    print("ROLE role assignment:", f"PATCH {ROLE_BASE_URL}/userapi/localUsers/{{userID}}")

    if not args.assign:
        print("\nPreview only. Nothing was assigned.")
        print("To assign valid ROLE roles, rerun with --assign --confirm-assign /approve.")
        return


    blocked = assignment_blocker_count(summary, args.skip_blocked_rows)
    if blocked:
        print("\nNo ROLE role assignments were made because one or more rows are blocked. Fix the blocked rows above, then rerun.")
        raise SystemExit(1)

    confirm = args.confirm_assign or input("\nType /approve to assign valid ROLE roles: ")
    if not approval_accepted(confirm):
        print("Not assigning. Confirmation did not match.")
        return

    if not summary["prepared"]:
        print("\nNo ROLE role assignments to make.")
        return

    assigned, failed = assign_roles(token, summary["prepared"])
    print(f"\nRMT role assignment summary: assigned={assigned}, failed={failed}")
    if failed:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
