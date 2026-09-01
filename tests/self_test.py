#!/usr/bin/env python3
import os
import sys
import tempfile
from email.message import EmailMessage

TEST_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(TEST_DIR)
SCRIPT_DIR = os.path.join(PROJECT_ROOT, 'scripts')
sys.path.insert(0, SCRIPT_DIR)

import email_intake
import manual_user_change_to_excel as manual
import setup_role_assignments_from_excel as role_api
import setup_study_from_excel as study_api


FAILURES = []


def check(name, condition, details=""):
    if condition:
        print(f"PASS: {name}")
    else:
        message = f"FAIL: {name}"
        if details:
            message += f" - {details}"
        print(message)
        FAILURES.append(message)


def fake_success_list(key, items):
    return 200, {"code": "SUCCESS", "payload": {key: items}}, ""


def test_approval_is_exact():
    check("approval accepts exact /approve", study_api.approval_accepted("/approve") is True)
    for value in ["/APPROVE", "approve", "CREATE", "UPDATE", "ASSIGN", "yes", " /approve "]:
        check(f"approval rejects {value!r}", study_api.approval_accepted(value) is False)


def test_site_code_exact_matching():
    check("site code keeps 1 literal", study_api.normalize_site_code(1) == "1")
    check("site code keeps 01 literal", study_api.normalize_site_code("01") == "01")

    def run_case(requested, returned, expected_count):
        original = study_api.request_json

        def fake_request_json(method, path, body=None, token=None):
            return fake_success_list("sites", [{"_id": "site-id", "code": returned, "name": "返回中心"}])

        study_api.request_json = fake_request_json
        try:
            matches = study_api.lookup_sites_by_code("token", "trial", requested)[2]
        finally:
            study_api.request_json = original
        check(f"site exact match Excel {requested!r} vs STUDY {returned!r}", len(matches) == expected_count)

    run_case("1", "1", 1)
    run_case("1", "01", 0)
    run_case("01", "1", 0)
    run_case("01", "01", 1)
    run_case("01", "CN011", 0)


def test_site_scope_resolution():
    check("single site id remains one code", study_api.split_site_codes("CN012") == ["CN012"])
    check("multiple site ids split on ampersand", study_api.split_site_codes("CN012&CN028") == ["CN012", "CN028"])
    check("multiple site ids split on full-width ampersand", study_api.split_site_codes("CN012＆CN028") == ["CN012", "CN028"])
    check("All normalizes to project-wide scope", study_api.split_site_codes("All") == ["所有"])

    original_request = study_api.request_json
    pages = []

    def fake_all_sites(method, path, body=None, token=None):
        pages.append(path)
        if "page=1" in path:
            return fake_success_list(
                "sites",
                [
                    {"_id": "site-12", "code": "CN012", "name": "中心12"},
                    {"_id": "site-28", "code": "CN028", "name": "中心28"},
                ],
            )
        return fake_success_list("sites", [])

    study_api.request_json = fake_all_sites
    try:
        all_sites = study_api.lookup_all_project_sites("token", "trial")
    finally:
        study_api.request_json = original_request
    check("All site lookup reads all server sites", [site["code"] for site in all_sites] == ["CN012", "CN028"], str(all_sites))
    check("All site lookup paginates until empty", len(pages) == 2, str(pages))

    original_roles = study_api.lookup_roles
    original_site_map = study_api.build_site_map
    original_users = study_api.lookup_users_by_username

    def fake_roles(token):
        return [{"_id": "role-pm", "name": "PM"}]

    def fake_site_map(token, trial_id, site_codes):
        return (
            {code: {"id": f"id-{code}", "name": code, "code": code} for code in site_codes},
            [],
        )

    def fake_users(token, username):
        return 200, "SUCCESS", [], ""

    def prepared_payload(site_value, project_scope=None):
        user = {
            "username": f"TEST-{site_value}",
            "name": "测试用户",
            "email": "user1@example.test",
            "telephone": "13800138000",
            "role": "PM",
            "siteCode": site_value,
            "siteCodes": study_api.split_site_codes(site_value),
        }
        summary = study_api.prepare_users(
            "token",
            "trial",
            [user],
            project_scope_site_codes=set(project_scope or []),
        )
        return summary, summary["prepared"][0] if summary["prepared"] else None

    study_api.lookup_roles = fake_roles
    study_api.build_site_map = fake_site_map
    study_api.lookup_users_by_username = fake_users
    try:
        single_summary, single = prepared_payload("CN012")
        multi_summary, multi = prepared_payload("CN012&CN028")
        all_summary, all_payload = prepared_payload("All", {"CN012", "CN028"})
        empty_all_summary, empty_all = prepared_payload("All", set())
    finally:
        study_api.lookup_roles = original_roles
        study_api.build_site_map = original_site_map
        study_api.lookup_users_by_username = original_users

    check("single site payload has one site id", single and single["siteIDs"] == ["id-CN012"], str(single_summary))
    check(
        "multiple site payload has both site ids",
        multi and multi["siteIDs"] == ["id-CN012", "id-CN028"],
        str(multi_summary),
    )
    check(
        "All payload expands to every project site id",
        all_payload and all_payload["siteIDs"] == ["id-CN012", "id-CN028"],
        str(all_summary),
    )
    check(
        "All with no project sites is blocked before API write",
        empty_all is None and len(empty_all_summary["missing_site"]) == 1,
        str(empty_all_summary),
    )

    try:
        study_api.fail_if_write_failures("test", user_create=1)
        write_failure_raised = False
    except SystemExit as exc:
        write_failure_raised = exc.code == 1
    check("write failures propagate a non-zero exit", write_failure_raised)


def test_site_name_and_study_username_exact_matching():
    original = study_api.request_json

    def fake_site_name(method, path, body=None, token=None):
        return fake_success_list("sites", [{"_id": "site-id", "code": "1", "name": "测试中心ABC"}])

    study_api.request_json = fake_site_name
    try:
        fuzzy_name_matches = study_api.lookup_sites_by_name("token", "trial", "测试中心")[2]
    finally:
        study_api.request_json = original
    check("site name rejects fuzzy result", fuzzy_name_matches == [])

    def fake_study_user(method, path, body=None, token=None):
        return fake_success_list("users", [{"_id": "user-id", "username": "DEMO400-USER-EXAMPLE-OLD"}])

    study_api.request_json = fake_study_user
    try:
        fuzzy_user_matches = study_api.lookup_users_by_username("token", "DEMO400-USER-EXAMPLE")[2]
    finally:
        study_api.request_json = original
    check("STUDY username rejects fuzzy result", fuzzy_user_matches == [])


def test_role_username_exact_matching():
    original = role_api.request_json

    def fake_role_fuzzy(method, path, body=None, token=None):
        if path.startswith("/userapi/users?"):
            return fake_success_list("users", [{"_id": "bad-list", "username": "DEMO400-USER-EXAMPLE-OLD"}])
        if path.startswith("/userapi/users/null"):
            return 200, {"code": "SUCCESS", "payload": {"_id": "bad-fallback", "username": "DEMO400-USER-EXAMPLE-OLD"}}, ""
        return 200, {"code": "SUCCESS", "payload": {}}, ""

    role_api.request_json = fake_role_fuzzy
    try:
        fuzzy = role_api.find_role_user("token", "DEMO400-USER-EXAMPLE")
    finally:
        role_api.request_json = original
    check("ROLE username rejects fuzzy list and fallback", fuzzy is None)

    def fake_role_exact(method, path, body=None, token=None):
        if path.startswith("/userapi/users?"):
            return fake_success_list("users", [])
        if path.startswith("/userapi/users/null"):
            return 200, {"code": "SUCCESS", "payload": {"_id": "good", "username": "DEMO400-USER-EXAMPLE"}}, ""
        return 200, {"code": "SUCCESS", "payload": {}}, ""

    role_api.request_json = fake_role_exact
    try:
        exact = role_api.find_role_user("token", "DEMO400-USER-EXAMPLE")
    finally:
        role_api.request_json = original
    check("ROLE username accepts exact fallback", bool(exact and exact.get("_id") == "good"))


def test_username_generation_rules():
    study_api.configure_project_id("demo012", required=False)
    users = [
        {"name": "Example User", "role": "CRC", "siteCode": "023", "siteCodes": ["023"]},
        {"name": "Example User", "role": "CRC", "siteCode": "023", "siteCodes": ["023"]},
        {"name": "Example User", "role": "CRA", "siteCode": "023", "siteCodes": ["023"]},
        {"username": "demo099-user-explicit", "name": "Example User", "role": "CRC", "siteCode": "067", "siteCodes": ["067"]},
    ]
    study_api.assign_run_usernames(users)
    built = [study_api.build_username(user) for user in users]
    check("duplicate generated usernames get suffix 1", built[0] == "DEMO012-CRC023EU1", built[0])
    check("duplicate generated usernames get suffix 2", built[1] == "DEMO012-CRC023EU2", built[1])
    check("CRA username omits site code", built[2] == "DEMO012-CRAEU", built[2])
    check("explicit username uppercases only", built[3] == "DEMO099-USER-EXPLICIT", built[3])


def test_update_note_and_direct_payload_rules():
    check("blank note is open", study_api.note_update_action({"note": ""}) == "open")
    check("unrecognized note is skipped normally", study_api.note_update_action({"note": "change email"}) == "")
    check("unrecognized note is open in direct mode", study_api.note_update_action({"note": "change email"}, "open") == "open")
    check("disable note disables", study_api.note_update_action({"note": "please disable"}) == "disable")
    check("ambiguous note blocks", study_api.note_update_action({"note": "open then disable"}) == "ambiguous")
    check("operation open ignores date note", study_api.note_update_action({"operation": "账号开通", "note": "2026.7.22", "_hasOperationColumn": True}) == "open")
    check("operation update is update-only", study_api.note_update_action({"operation": "账号更新", "note": "2026.7.22", "_hasOperationColumn": True}) == "update")
    check("operation disable disables", study_api.note_update_action({"operation": "账号禁用", "note": "2026.7.22", "_hasOperationColumn": True}) == "disable")
    check("visitor wording selects visitor app", study_api.note_device_types({"note": "Visitor APP"}) == ["visit"])
    check("investigator wording selects investigator app", study_api.note_device_types({"note": "Investigator End"}) == ["doctor"])

    study_api.configure_project_id("DEMO400", required=False)
    payload, missing = study_api.direct_update_payload(
        {
            "username": "demo400-user-example",
            "email": "user2@example.test",
            "visitorAppChoice": "off",
            "investigatorAppChoice": "on",
            "notificationMethod": "Email",
        },
        {"deviceTypes": ["visit"]},
        {},
        [],
        set(),
    )
    check("direct payload has no missing fields", missing == [], str(missing))
    check("direct payload uppercases username", payload.get("username") == "DEMO400-USER-EXAMPLE", str(payload))
    check("direct payload preserves/changes app selections", payload.get("deviceTypes") == ["doctor"], str(payload))
    check("direct payload sends email only when requested", payload.get("authorizedNotifications") == ["EMAIL"], str(payload))
    check("direct payload does not send blank phone", "phone" not in payload, str(payload))


def test_disable_finds_legacy_username_format():
    study_api.configure_project_id("demo100test", required=False)
    original = study_api.request_json

    def fake_request_json(method, path, body=None, token=None):
        if path.startswith("/userapi/users?"):
            parsed = study_api.urllib.parse.urlparse(path)
            query = study_api.urllib.parse.parse_qs(parsed.query)
            if query.get("username", [""])[0] == "DEMO100TEST-CRC01EX":
                return fake_success_list("users", [])
            if query.get("email", [""])[0] == "user3@example.test":
                return fake_success_list("users", [
                    {
                        "_id": "old-user",
                        "username": "T-DEMO100-CRC-EX",
                        "email": "user3@example.test",
                        "phone": "10000000003",
                        "siteID": {"code": "03", "name": "示例第三中心"},
                        "status": 1,
                    },
                    {
                        "_id": "other-user",
                        "username": "DEMO100TEST-CRC01OTHER",
                        "email": "user3@example.test",
                        "phone": "10000000003",
                        "siteID": {"code": "01", "name": "示例第一中心"},
                        "status": 1,
                    },
                ])
        return fake_success_list("users", [])

    study_api.request_json = fake_request_json
    try:
        username, existing, match_type, ambiguous = study_api.resolve_disable_user(
            "token",
            {
                "excelRow": 5,
                "siteCode": "01",
                "siteCodes": ["01"],
                "name": "Example X",
                "email": "user3@example.test",
                "telephone": "10000000003",
                "role": "CRC",
                "operation": "账号禁用",
            },
        )
    finally:
        study_api.request_json = original

    check("disable fallback finds legacy username", bool(existing and existing.get("_id") == "old-user"), str(existing))
    check("disable fallback reports legacy match", match_type == "legacy", match_type)
    check("disable fallback uses existing username", username == "T-DEMO100-CRC-EX", username)
    check("disable fallback is not ambiguous", ambiguous == [], str(ambiguous))


def test_update_username_legacy_candidates_are_deterministic():
    original_project_id = study_api.PROJECT_ID
    try:
        study_api.PROJECT_ID = "DEMO300"
        demo300_user = {"name": "EX", "initials": "EX", "role": "CRA"}
        demo300_candidates = study_api.generated_username_candidates(demo300_user)
        check("DEMO300 update checks historical role separator", "DEMO300-CRA-EX" in demo300_candidates, str(demo300_candidates))

        study_api.PROJECT_ID = "U-DEMO401"
        all_sites = study_api.split_site_codes("All")
        ubr_user = {"name": "EX", "initials": "EX", "role": "SUBI", "siteCode": "All", "siteCodes": all_sites}
        ubr_candidates = study_api.generated_username_candidates(ubr_user)
        check("U-DEMO401 update checks historical project separator", "UDEMO401-SUBIEX" in ubr_candidates, str(ubr_candidates))
        check("current U-DEMO401 username remains first", ubr_candidates[0] == "U-DEMO401-SUBIEX", str(ubr_candidates))
    finally:
        study_api.PROJECT_ID = original_project_id


def test_study_update_missing_user_blocks_before_pending_site():
    original_project_id = study_api.PROJECT_ID
    originals = (study_api.lookup_roles, study_api.build_site_map, study_api.resolve_update_user)
    user = {
        "excelRow": 4,
        "name": "EX",
        "initials": "EX",
        "email": "user4@example.test",
        "telephone": "13800000000",
        "role": "SUBI",
        "siteCode": "007",
        "siteCodes": ["007"],
        "operation": "update",
    }
    try:
        study_api.PROJECT_ID = "U-DEMO401"
        study_api.lookup_roles = lambda token: [{"_id": "role-id", "name": "SUBI"}]
        study_api.build_site_map = lambda token, trial_id, codes: ({}, set(codes))
        study_api.resolve_update_user = lambda token, row: (study_api.build_username(row), None, "missing", [])
        summary = study_api.prepare_update_actions("token", "trial", [user], pending_site_codes={"007"})
    finally:
        study_api.PROJECT_ID = original_project_id
        study_api.lookup_roles, study_api.build_site_map, study_api.resolve_update_user = originals
    check("missing update user blocks in first preflight", len(summary["missing_update_user"]) == 1, str(summary))
    check("missing update user is not deferred behind site creation", len(summary["waiting_for_site_creation"]) == 0, str(summary))


def test_role_legacy_lookup_only_applies_to_update_rows():
    original_project_id = study_api.PROJECT_ID
    original_find = role_api.find_role_user
    calls = []
    user = {
        "excelRow": 4,
        "name": "EX",
        "initials": "EX",
        "email": "user4@example.test",
        "telephone": "13800000000",
        "role": "SUBI",
        "siteCode": "All",
        "siteCodes": study_api.split_site_codes("All"),
    }
    roles = [{"_id": "role-id", "name": "SUBI", "baseRole": 0, "status": 1}]

    def fake_find(token, username):
        calls.append(username)
        if username == "UDEMO401-SUBIEX":
            return {"_id": "user-id", "username": username, "role": 0, "source": "notAssign"}
        return None

    try:
        study_api.PROJECT_ID = "U-DEMO401"
        role_api.find_role_user = fake_find
        update_summary = role_api.prepare_assignments("token", [user], roles, {}, legacy_user_ids={id(user)})
        update_calls = list(calls)
        calls.clear()
        create_summary = role_api.prepare_assignments("token", [user], roles, {})
        create_calls = list(calls)
    finally:
        study_api.PROJECT_ID = original_project_id
        role_api.find_role_user = original_find

    update_items = update_summary["prepared"] + update_summary["no_change"]
    check("ROLE update resolves historical username", bool(update_items) and update_items[0]["username"] == "UDEMO401-SUBIEX", str(update_summary))
    check("ROLE update records legacy match", bool(update_items) and update_items[0]["matchType"] == "legacy", str(update_summary))
    check("ROLE update checked historical candidate", "UDEMO401-SUBIEX" in update_calls, str(update_calls))
    check("ROLE new account does not use historical candidate", "UDEMO401-SUBIEX" not in create_calls and len(create_summary["missing_user"]) == 1, str(create_calls))


def test_role_ambiguous_legacy_matches_block_all_writes():
    original_project_id = study_api.PROJECT_ID
    original_find = role_api.find_role_user
    user = {
        "excelRow": 4,
        "name": "EX",
        "initials": "EX",
        "email": "user5@example.test",
        "telephone": "13800000000",
        "role": "CRA",
        "siteCode": "001",
        "siteCodes": ["001"],
    }
    roles = [{"_id": "role-id", "name": "CRA", "baseRole": 0, "status": 1}]

    def fake_find(token, username):
        if username in {"DEMO300-CRAEX", "DEMO300-CRA-EX"}:
            return {"_id": username, "username": username, "role": 0, "source": "notAssign"}
        return None

    try:
        study_api.PROJECT_ID = "DEMO300"
        role_api.find_role_user = fake_find
        summary = role_api.prepare_assignments("token", [user], roles, {}, legacy_user_ids={id(user)})
    finally:
        study_api.PROJECT_ID = original_project_id
        role_api.find_role_user = original_find
    check("multiple exact historical ROLE users are ambiguous", len(summary["ambiguous_user"]) == 1, str(summary))
    check("ambiguous ROLE user produces no prepared write", not summary["prepared"], str(summary))
    check("ambiguous ROLE user counts as pre-write blocker", role_api.assignment_blocker_count(summary) == 1, str(summary))


def test_role_main_blocks_before_patch_or_delete():
    names = (
        "extract_project_details_from_excels",
        "apply_project_details_to_args",
        "configure_role_runtime",
        "parse_all_excel_users",
        "login",
        "lookup_role_assignments",
        "prepare_update_actions",
        "print_update_preview",
        "assign_roles",
        "delete_role_users",
    )
    originals = {name: getattr(role_api, name) for name in names}
    original_argv = list(sys.argv)
    writes = []
    assignments = {
        "prepared": [{"username": "SHOULD-NOT-WRITE"}],
        "invalid": [],
        "missing_role": [],
        "missing_user": [],
        "duplicates": [],
        "no_change": [],
        "role_base_mismatch": [],
        "user_role_mismatch": [],
        "ambiguous_user": [(None, "AMBIGUOUS", [])],
    }
    summary = {
        "assignments": assignments,
        "disable_ready": [{"username": "SHOULD-NOT-DELETE"}],
        "disable_not_found": [],
        "invalid_disable": [],
        "duplicate_disable": [],
        "ambiguous_disable_user": [],
        "open_rows": [],
        "update_only_rows": [],
        "disable_rows": [],
        "ambiguous": [],
        "ignored": [],
    }
    try:
        role_api.extract_project_details_from_excels = lambda paths: {}
        role_api.apply_project_details_to_args = lambda args, details: None
        role_api.configure_role_runtime = lambda args, require_api=True: "demo401"
        role_api.parse_all_excel_users = lambda paths, sheet: []
        role_api.login = lambda username, password: ("token", "trial")
        role_api.lookup_role_assignments = lambda token: []
        role_api.prepare_update_actions = lambda *args, **kwargs: summary
        role_api.print_update_preview = lambda *args, **kwargs: None
        role_api.assign_roles = lambda *args, **kwargs: writes.append("PATCH") or (1, 0)
        role_api.delete_role_users = lambda *args, **kwargs: writes.append("DELETE") or (1, 0)
        sys.argv = [
            "setup_role_assignments_from_excel.py",
            "--excel-file", "dummy.xlsx",
            "--project-url", "demo401",
            "--project-id", "U-DEMO401",
            "--username", "U-DEMO-Admin",
            "--update",
            "--confirm-update", "/approve",
        ]
        try:
            role_api.main()
        except SystemExit as exc:
            blocked = exc.code == 1
        else:
            blocked = False
    finally:
        sys.argv = original_argv
        for name, value in originals.items():
            setattr(role_api, name, value)
    check("ROLE apply exits when any row is blocked", blocked)
    check("ROLE apply performs zero writes when blocked", writes == [], str(writes))


def test_manual_workbook_roundtrip():
    study_api.configure_project_id("DEMO400", required=False)
    with tempfile.TemporaryDirectory() as temp_dir:
        path = os.path.join(temp_dir, "manual-user-change.xlsx")
        manual.write_xlsx(
            path,
            [
                ("研究中心", [manual.SITE_HEADERS]),
                (
                    "研究人员信息",
                    [
                        manual.USER_HEADERS,
                        ["", "", "user2@example.test", "", "", "", "demo400-user-example", "", "off", "on", "Email", "", ""],
                    ],
                ),
            ],
        )
        sites, users = study_api.parse_all_excels([path])
        study_api.assign_run_usernames(users)
        check("manual workbook has no site rows", len(sites) == 0)
        check("manual workbook has one user row", len(users) == 1)
        check("manual workbook username parses exactly", study_api.build_username(users[0]) == "DEMO400-USER-EXAMPLE")
        check("manual workbook visitor app parses off", study_api.parse_bool_choice(users[0].get("visitorAppChoice")) is False)
        check("manual workbook investigator app parses on", study_api.parse_bool_choice(users[0].get("investigatorAppChoice")) is True)
        check("manual workbook notification parses email", study_api.notification_methods_from_text(users[0].get("notificationMethod")) == ["EMAIL"])


def test_orange_template_layout_roundtrip():
    study_api.configure_project_id("DEMO001", required=False)
    with tempfile.TemporaryDirectory() as temp_dir:
        path = os.path.join(temp_dir, "orange-template.xlsx")
        manual.write_xlsx(
            path,
            [
                (
                    "试验信息",
                    [
                        ["试验信息", "申办方", "示例申办方"],
                        ["", "方案编号", "DEMO-PROTOCOL-001"],
                        ["", "项目编号", "DEMO001"],
                        ["", "项目名称", "示例自动化测试"],
                    ],
                ),
                (
                    "中心信息",
                    [
                        ["研究中心编号*", "研究中心名称*", "地址*", "PI姓名*", "PI电话*", "PI电子邮件*"],
                        ["SITE-001", "示例研究中心", "示例市示例区", "示例市", "10000000001", "user6@example.test"],
                    ],
                ),
                (
                    "用户申请",
                    [
                        ["研究中心编号*", "姓名*", "邮箱*", "电话*", "角色*", "账号操作类型*", "备注（请填写申请日期YYYY-MM-DD）"],
                        ["SITE-001", "示例用户甲", "user7@example.test", "10000000002", "PI", "账号开通", "2026.7.22"],
                        ["SITE-001", "示例用户乙", "user8@example.test", "10000000003", "SUBI", "账号更新", "2026.7.22"],
                        ["SITE-001", "示例用户丙", "user9@example.test", "10000000004", "SUBI", "账号禁用", "2026.7.22"],
                    ],
                ),
            ],
        )
        sites, users = study_api.parse_all_excels([path])
        study_api.assign_run_usernames(users)
        classes = study_api.classify_update_users(users)
        check("orange template site sheet parses", len(sites) == 1)
        check("orange template PI email header parses", sites[0].get("email") == "user6@example.test", str(sites[0]))
        check("orange template user sheet parses", len(users) == 3)
        check("orange template operation column parses", users[0].get("operation") == "账号开通", str(users[0]))
        check("orange template open/update rows parse", len(classes["open"]) == 2)
        check("orange template update-only row parses", len(classes.get("update_only", [])) == 1)
        check("orange template disable row parses", len(classes["disable"]) == 1)
        check("orange template date notes are not ignored", len(classes["ignored"]) == 0)
        project_details = study_api.extract_project_details_from_excels([path])
        check("orange template project number parses as ProjectID", project_details.get("project_id") == "DEMO001", str(project_details))
        check("orange template project number falls back to project path", project_details.get("project_path") == "DEMO001", str(project_details))


def test_site_optional_fields_default_to_dash():
    minimal_site = {
        "code": "CN002",
        "name": "示例测试中心",
        "address": "",
        "contactName": "",
        "telephone": "",
        "email": "",
    }
    payload = study_api.site_payload(minimal_site)
    check("minimal site with code and name is valid", study_api.validate_site(minimal_site) == [], str(study_api.validate_site(minimal_site)))
    check("missing optional site address becomes dash", payload.get("address") == "--", str(payload))
    check("missing optional site PI name becomes dash", payload.get("contactName") == "--", str(payload))
    check("missing optional site phone becomes dash", payload.get("telephone") == "--", str(payload))
    check("missing optional site email becomes dash", payload.get("email") == "--", str(payload))
    check("site missing code is invalid", study_api.validate_site({"name": "示例测试中心"}) == ["code"])
    check("site missing name is invalid", study_api.validate_site({"code": "CN002"}) == ["name"])


def test_existing_site_updates_preserve_blank_optional_fields():
    excel_site = {
        "code": "01",
        "name": "新第一中心",
        "address": "",
        "contactName": "新PI",
        "telephone": "",
        "email": "",
    }
    existing_site = {
        "_id": "site-01",
        "code": "01",
        "name": "第一中心",
        "address": "旧地址",
        "contactName": "旧PI",
        "telephone": "10000000001",
        "email": "user10@example.test",
    }
    payload = study_api.site_update_payload(excel_site, existing_site)
    changes = study_api.site_payload_changes(payload, existing_site)
    check("site update changes filled name", payload.get("name") == "新第一中心", str(payload))
    check("site update changes filled PI", payload.get("contactName") == "新PI", str(payload))
    check("site update preserves blank address", payload.get("address") == "旧地址", str(payload))
    check("site update preserves blank phone", payload.get("telephone") == "10000000001", str(payload))
    check("site update preserves blank email", payload.get("email") == "user10@example.test", str(payload))
    check("site update reports changed fields only", set(changes) == {"name", "contactName"}, str(changes))


def test_application_email_format_validation():
    check("normal email format is accepted", study_api.validate_email_field({"email": "user11@example.test"}) == "")
    for value in (
        "user @example.com",
        "user@ example.com",
        " user12@example.test",
        "user12@example.test ",
        "user\tuser13@example.test",
        "user\nuser14@example.test",
    ):
        issue = study_api.validate_email_field({"email": value.strip(), "_rawEmail": value})
        check(f"email whitespace is rejected: {value!r}", issue == "emailWhitespace", issue)
    for value in (
        "userexample.com",
        "user@@example.com",
        ".user12@example.test",
        "user@-example.test",
        "user@example",
        "user@example..test",
    ):
        issue = study_api.validate_email_field({"email": value, "_rawEmail": value})
        check(f"invalid email syntax is rejected: {value!r}", issue == "emailFormat", issue)


def test_row_level_skip_prepares_only_complete_users():
    study_api.configure_project_id("TEST001", required=False)
    users = []
    for index in range(10):
        users.append({
            "excelRow": index + 2,
            "siteCode": "CN001",
            "siteCodes": ["CN001"],
            "name": f"User {index + 1}",
            "initials": f"U{index + 1}",
            "email": f"user{index + 1}@example.com",
            "telephone": f"1380000000{index}",
            "role": "ADMIN",
            "operation": "账号开通",
            "_requireOperation": True,
        })
    users[6]["email"] = "bad @ example.com"
    users[6]["_rawEmail"] = "bad @ example.com"
    users[7]["telephone"] = ""
    users[8]["role"] = ""
    users[9]["operation"] = ""
    originals = {
        "lookup_roles": study_api.lookup_roles,
        "build_site_map": study_api.build_site_map,
        "lookup_users_by_username": study_api.lookup_users_by_username,
    }
    try:
        study_api.lookup_roles = lambda _token: [{"_id": "role-admin", "name": "ADMIN"}]
        study_api.build_site_map = lambda _token, _trial, _codes: ({}, [])
        study_api.lookup_users_by_username = lambda _token, _username: (200, "SUCCESS", [], "{}")
        summary = study_api.prepare_users("token", "trial", users)
    finally:
        for name, value in originals.items():
            setattr(study_api, name, value)

    check("six complete users are prepared", len(summary["prepared"]) == 6, str(summary))
    check("four incomplete users are isolated", len(summary["invalid"]) == 4, str(summary))
    missing_fields = [fields for _user, fields in summary["invalid"]]
    check("email containing spaces is rejected", ["emailWhitespace"] in missing_fields, str(missing_fields))
    check("telephone is required", ["telephone"] in missing_fields, str(missing_fields))
    check("role is required", ["role"] in missing_fields, str(missing_fields))
    check("account operation is required", ["operation"] in missing_fields, str(missing_fields))
    try:
        study_api.fail_if_user_blockers("test", summary, ["invalid"], skip_blocked_rows=True)
        continued = True
    except SystemExit:
        continued = False
    check("email row-level mode continues past isolated blockers", continued)


def test_email_intake_job_roundtrip():
    with tempfile.TemporaryDirectory() as temp_dir:
        workbook_path = os.path.join(temp_dir, "orange-template.xlsx")
        manual.write_xlsx(
            workbook_path,
            [
                (
                    "试验信息",
                    [
                        ["试验信息", "申办方", "示例申办方"],
                        ["", "方案编号", "DEMO-PROTOCOL-001"],
                        ["", "项目编号", "DEMO001"],
                        ["", "项目名称", "示例自动化测试"],
                    ],
                ),
                (
                    "中心信息",
                    [
                        ["研究中心编号*", "研究中心名称*", "地址*", "PI姓名*", "PI电话*", "PI电子邮件*"],
                        ["SITE-001", "示例研究中心", "示例市示例区", "示例市", "10000000001", "user6@example.test"],
                    ],
                ),
                (
                    "用户申请",
                    [
                        ["研究中心编号*", "姓名*", "邮箱*", "电话*", "角色*", "账号操作类型*", "备注（请填写申请日期YYYY-MM-DD）"],
                        ["SITE-001", "示例用户甲", "user7@example.test", "10000000002", "PI", "账号更新", "2026.7.22"],
                    ],
                ),
            ],
        )

        message = EmailMessage()
        message["From"] = "support@example.test"
        message["Subject"] = "Study Access account update"
        message["Message-ID"] = "<user18@example.test>"
        message.set_content("Project URL/path: demo900\nProjectID: DEMO001\nPlease process the attached workbook.")
        with open(workbook_path, "rb") as file:
            message.add_attachment(
                file.read(),
                maintype="application",
                subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                filename="用户申请.xlsx",
            )

        job = email_intake.process_message_bytes(
            message.as_bytes(),
            os.path.join(temp_dir, "jobs"),
            mode="auto",
            scope="both",
            live_preview=False,
            timeout=120,
        )
        preview_path = os.path.join(job["job_dir"], "preview.txt")
        preview = open(preview_path, "r", encoding="utf-8").read()
        preview_email_path = os.path.join(job["job_dir"], "preview_email.txt")
        preview_email = open(preview_email_path, "r", encoding="utf-8").read()
        chat_message_path = os.path.join(job["job_dir"], "chat_message.md")
        chat_message = open(chat_message_path, "r", encoding="utf-8").read()
        dingtalk_message_path = os.path.join(job["job_dir"], "dingtalk_message.md")
        dingtalk_message = open(dingtalk_message_path, "r", encoding="utf-8").read()
        check("email intake creates preview job", job.get("status") == "previewed", str(job))
        check("email intake saves attachment", len(job.get("excel_files", [])) == 1)
        check("email intake extracts project id", job.get("project_id") == "DEMO001", str(job))
        check("email intake infers update mode from operation column", job.get("mode") == "update", str(job))
        check("email intake preview contains STUDY parse", "STUDY study setup helper" in preview)
        check("email intake preview contains ROLE parse", "ROLE role assignment helper" in preview)
        check("email intake writes preview email", os.path.exists(preview_email_path))
        check("preview email asks for job-specific approval", f"/approve {job['job_id']}" in preview_email)
        check("email intake writes Hermes chat message", os.path.exists(chat_message_path))
        check("Hermes chat message asks for job-specific approval", f"/approve {job['job_id']}" in chat_message)
        check("Hermes chat message includes compact preview", "STUDY study setup helper" in chat_message)
        check("email intake writes DingTalk message", os.path.exists(dingtalk_message_path))
        check("DingTalk message asks for job-specific approval", f"/approve {job['job_id']}" in dingtalk_message)
        check("DingTalk message includes compact preview", "STUDY study setup helper" in dingtalk_message)


def test_email_intake_auto_apply():
    with tempfile.TemporaryDirectory() as temp_dir:
        workbook_path = os.path.join(temp_dir, "orange-template.xlsx")
        manual.write_xlsx(
            workbook_path,
            [
                (
                    "试验信息",
                    [
                        ["试验信息", "申办方", "示例申办方"],
                        ["", "方案编号", "DEMO-PROTOCOL-001"],
                        ["", "项目编号", "DEMO001"],
                        ["", "项目名称", "示例自动化测试"],
                    ],
                ),
                (
                    "中心信息",
                    [
                        ["研究中心编号*", "研究中心名称*", "地址*", "PI姓名*", "PI电话*", "PI电子邮件*"],
                        ["SITE-001", "示例研究中心", "", "", "", ""],
                    ],
                ),
                (
                    "用户申请",
                    [
                        ["研究中心编号*", "姓名*", "邮箱*", "电话*", "角色*", "账号操作类型*", "备注（请填写申请日期YYYY-MM-DD）"],
                        ["SITE-001", "示例用户甲", "user7@example.test", "10000000002", "PI", "账号更新", "2026.7.22"],
                    ],
                ),
            ],
        )

        message = EmailMessage()
        message["From"] = "support@example.test"
        message["Subject"] = "Study Access auto apply"
        message["Message-ID"] = "<user19@example.test>"
        message.set_content("Please process the attached workbook. Project info is in the workbook.")
        with open(workbook_path, "rb") as file:
            message.add_attachment(
                file.read(),
                maintype="application",
                subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                filename="用户申请.xlsx",
            )

        env_keys = [
            "STUDY_ACCESS_PROJECT_URL",
            "STUDY_ACCESS_PROJECT_PATH",
            "STUDY_ACCESS_PROJECT_ID",
            "PROJECT_ID",
        ]
        saved_env = {key: os.environ.get(key) for key in env_keys}
        original_run = email_intake.run_commands_with_results
        calls = []

        def fake_run(commands, timeout):
            calls.append(commands)
            results = []
            for command in commands:
                results.append(
                    {
                        "command": command,
                        "stdout": "OK " + os.path.basename(command[2]),
                        "stderr": "",
                        "returncode": 0,
                    }
                )
            return email_intake.format_command_results(results), results

        try:
            for key in env_keys:
                os.environ.pop(key, None)
            email_intake.run_commands_with_results = fake_run
            job = email_intake.process_message_bytes(
                message.as_bytes(),
                os.path.join(temp_dir, "jobs"),
                mode="auto",
                scope="both",
                live_preview=False,
                timeout=120,
                auto_apply=True,
            )
        finally:
            email_intake.run_commands_with_results = original_run
            for key, value in saved_env.items():
                if value is None:
                    os.environ.pop(key, None)
                else:
                    os.environ[key] = value

        chat_message = open(job["chat_message_path"], "r", encoding="utf-8").read()
        preview_chat_message = open(job["preview_chat_message_path"], "r", encoding="utf-8").read()
        flattened_preview = [part for command in calls[0] for part in command] if calls else []
        flattened_apply = [part for command in calls[1] for part in command] if len(calls) > 1 else []
        check("email intake auto apply marks applied", job.get("status") == "applied", str(job))
        check("email intake auto apply runs preview then apply", len(calls) == 2, str(calls))
        check("auto apply uses project from workbook", "DEMO001" in flattened_preview, str(calls))
        check("auto apply uses universal username", "DEMO-Admin" in flattened_preview and "DEMO-Admin" in flattened_apply, str(calls))
        check("auto apply stores workbook project details", job.get("workbook_project_details", {}).get("project_id") == "DEMO001", str(job))
        check("auto apply forces live preview", "--parse-only" not in flattened_preview, str(calls))
        check("auto apply infers update mode", "--update" in flattened_apply, str(calls))
        check("auto apply sends internal confirm only", "--confirm-update" in flattened_apply and "/approve" in flattened_apply, str(calls))
        check("auto apply writes apply result", os.path.exists(job.get("apply_path", "")), str(job))
        check("auto apply chat message is final result", "apply result" in chat_message and "To approve this exact email job" not in chat_message)
        check("auto apply keeps separate approval preview artifact", "To approve this exact email job" in preview_chat_message)


def test_email_intake_management_gate():
    with tempfile.TemporaryDirectory() as temp_dir:
        message = EmailMessage()
        message["From"] = "Client Person <user20@example.test>"
        message["Subject"] = "账号申请"
        message["Message-ID"] = "<user21@example.test>"
        message.set_content("请完全不留存任何个人信息，并处理附件中的账号申请。")
        message.add_attachment(
            b"This attachment must not be saved or parsed.",
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="request.xlsx",
        )

        sent_messages = []
        original_send = email_intake.send_email_message

        def fake_send(to_addresses, subject, body, attachments=None, cc_addresses=None, **_kwargs):
            sent_messages.append({
                "to": list(to_addresses),
                "subject": subject,
                "body": body,
                "attachments": list(attachments or []),
                "cc": list(cc_addresses or []),
            })

        try:
            email_intake.send_email_message = fake_send
            job = email_intake.process_message_bytes(
                message.as_bytes(),
                os.path.join(temp_dir, "jobs"),
                auto_apply=True,
            )
        finally:
            email_intake.send_email_message = original_send

        check("management gate blocks restricted request", job.get("status") == "management_required", str(job))
        check(
            "management gate sends exact customer reply",
            len(sent_messages) == 1
            and sent_messages[0]["body"] == email_intake.MANAGEMENT_REQUIRED_REPLY,
            str({"sent": sent_messages, "error": job.get("management_reply_error")}),
        )
        if sent_messages:
            check("management gate replies only to client", sent_messages[0]["to"] == ["user20@example.test"] and not sent_messages[0]["cc"], str(sent_messages))
            check("management gate sends no attachment", not sent_messages[0]["attachments"], str(sent_messages))
        else:
            check("management gate replies only to client", False, "no reply was captured")
            check("management gate sends no attachment", False, "no reply was captured")
        check("management gate creates no job files", not os.path.exists(os.path.join(temp_dir, "jobs")), str(os.listdir(temp_dir)))


def test_email_intake_auto_apply_blocks_missing_project_config():
    with tempfile.TemporaryDirectory() as temp_dir:
        workbook_path = os.path.join(temp_dir, "orange-template.xlsx")
        manual.write_xlsx(
            workbook_path,
            [
                ("研究中心", [manual.SITE_HEADERS]),
                (
                    "用户申请",
                    [
                        ["研究中心编号*", "姓名*", "邮箱*", "电话*", "角色*", "账号操作类型*"],
                        ["SITE-001", "示例用户甲", "user7@example.test", "10000000002", "PI", "账号更新"],
                    ],
                )
            ],
        )

        message = EmailMessage()
        message["From"] = "support@example.test"
        message["Subject"] = "Study Access missing config"
        message["Message-ID"] = "<user22@example.test>"
        message.set_content("Please process the attached workbook.")
        with open(workbook_path, "rb") as file:
            message.add_attachment(
                file.read(),
                maintype="application",
                subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                filename="用户申请.xlsx",
            )

        env_keys = [
            "STUDY_ACCESS_PROJECT_URL",
            "STUDY_ACCESS_PROJECT_PATH",
            "STUDY_ACCESS_PROJECT_ID",
            "PROJECT_ID",
        ]
        saved_env = {key: os.environ.get(key) for key in env_keys}
        original_run = email_intake.run_commands_with_results
        calls = []

        def fake_run(commands, timeout):
            calls.append(commands)
            results = []
            for command in commands:
                results.append({"command": command, "stdout": "parse OK", "stderr": "", "returncode": 0})
            return email_intake.format_command_results(results), results

        try:
            for key in env_keys:
                os.environ.pop(key, None)
            email_intake.run_commands_with_results = fake_run
            job = email_intake.process_message_bytes(
                message.as_bytes(),
                os.path.join(temp_dir, "jobs"),
                mode="auto",
                scope="both",
                auto_apply=True,
            )
        finally:
            email_intake.run_commands_with_results = original_run
            for key, value in saved_env.items():
                if value is None:
                    os.environ.pop(key, None)
                else:
                    os.environ[key] = value

        chat_message = open(job["chat_message_path"], "r", encoding="utf-8").read()
        check("auto apply blocks when project config missing", job.get("status") == "auto_apply_blocked", str(job))
        check("auto apply missing config does not run apply", len(calls) == 1, str(calls))
        check("auto apply blocked message is clear", "was not auto-applied" in chat_message and "Missing:" in chat_message)


def test_email_intake_project_conflict_blocks_auto_apply():
    with tempfile.TemporaryDirectory() as temp_dir:
        workbook_path = os.path.join(temp_dir, "orange-template.xlsx")
        manual.write_xlsx(
            workbook_path,
            [
                ("研究中心", [manual.SITE_HEADERS]),
                (
                    "试验信息",
                    [
                        ["试验信息", "申办方", "示例申办方"],
                        ["", "方案编号", "DEMO-PROTOCOL-001"],
                        ["", "项目编号", "DEMO001"],
                    ],
                ),
                (
                    "用户申请",
                    [
                        ["研究中心编号*", "姓名*", "邮箱*", "电话*", "角色*", "账号操作类型*"],
                        ["SITE-001", "示例用户甲", "user7@example.test", "10000000002", "PI", "账号更新"],
                    ],
                ),
            ],
        )

        message = EmailMessage()
        message["From"] = "support@example.test"
        message["Subject"] = "Study Access conflicting project"
        message["Message-ID"] = "<user23@example.test>"
        message.set_content("Project URL/path: wrongproject\nPlease process the attached workbook.")
        with open(workbook_path, "rb") as file:
            message.add_attachment(
                file.read(),
                maintype="application",
                subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                filename="用户申请.xlsx",
            )

        env_keys = []
        saved_env = {key: os.environ.get(key) for key in env_keys}
        original_run = email_intake.run_commands_with_results
        calls = []

        def fake_run(commands, timeout):
            calls.append(commands)
            results = []
            for command in commands:
                results.append({"command": command, "stdout": "parse OK", "stderr": "", "returncode": 0})
            return email_intake.format_command_results(results), results

        try:
            email_intake.run_commands_with_results = fake_run
            job = email_intake.process_message_bytes(
                message.as_bytes(),
                os.path.join(temp_dir, "jobs"),
                mode="auto",
                scope="both",
                auto_apply=True,
            )
        finally:
            email_intake.run_commands_with_results = original_run
            for key, value in saved_env.items():
                if value is None:
                    os.environ.pop(key, None)
                else:
                    os.environ[key] = value

        chat_message = open(job["chat_message_path"], "r", encoding="utf-8").read()
        flattened_preview = [part for command in calls[0] for part in command] if calls else []
        check("project conflict blocks auto apply", job.get("status") == "auto_apply_blocked", str(job))
        check("project conflict is recorded", bool(job.get("project_conflicts")), str(job))
        check("project conflict uses parse-only preview", "--parse-only" in flattened_preview and len(calls) == 1, str(calls))
        check("project conflict message is clear", "Project info conflict" in chat_message, chat_message)


def test_email_client_recipient_detection():
    message = EmailMessage()
    message["From"] = "support@example.test"
    message["Subject"] = "FW: Study Access request"
    body = (
        "---------- Forwarded message ---------\n"
        "From: Client Person <user24@example.test>\n"
        "To: support@example.test\n"
        "Project URL/path: demo900\n"
        "ProjectID: DEMO001\n"
    )
    message.set_content(body)

    env_keys = ["STUDY_ACCESS_SUPPORT_FORWARDERS", "STUDY_ACCESS_EMAIL_PREVIEW_TO", "STUDY_ACCESS_ALLOW_SUPPORT_REPLY", "AGENT_EMAIL"]
    saved_env = {key: os.environ.get(key) for key in env_keys}
    try:
        os.environ["STUDY_ACCESS_SUPPORT_FORWARDERS"] = "support@example.test"
        os.environ.pop("STUDY_ACCESS_EMAIL_PREVIEW_TO", None)
        os.environ.pop("STUDY_ACCESS_ALLOW_SUPPORT_REPLY", None)
        os.environ["AGENT_EMAIL"] = "access.agent@example.test"
        check("forwarded client email is detected", email_intake.client_email_from_message(message, body) == "user24@example.test")
        check("configured recipients use client not support", email_intake.configured_preview_recipients(message, body) == ["user24@example.test"])

        direct_message = EmailMessage()
        direct_message["From"] = "Client Person <user24@example.test>"
        direct_message["To"] = "access.agent@example.test"
        direct_message["Cc"] = "Project PM <user26@example.test>, PM Agent <access.agent@example.test>, user24@example.test"
        direct_message.set_content("ProjectID: DEMO001")
        check(
            "direct project PM cc is retained and agent/client are filtered",
            email_intake.result_cc_recipients(
                direct_message,
                direct_message.get_content(),
                ["user24@example.test"],
            )
            == ["user26@example.test"],
        )

        forwarded_body = body.replace(
            "To: support@example.test\n",
            "To: support@example.test\nCc: Project PM <user27@example.test>\n",
        )
        check(
            "forwarded project PM cc is detected",
            email_intake.result_cc_recipients(message, forwarded_body, ["user24@example.test"])
            == ["user27@example.test"],
        )

        os.environ["STUDY_ACCESS_EMAIL_PREVIEW_TO"] = "support@example.test"
        check(
            "fixed support recipient is blocked by default",
            email_intake.configured_preview_recipients(message, body)
            == ["user24@example.test"],
        )
    finally:
        for key, value in saved_env.items():
            if value is None:
                os.environ.pop(key, None)
            else:
                os.environ[key] = value


def test_email_intake_auto_apply_emails_client_and_audits():
    with tempfile.TemporaryDirectory() as temp_dir:
        workbook_path = os.path.join(temp_dir, "orange-template.xlsx")
        manual.write_xlsx(
            workbook_path,
            [
                ("研究中心", [manual.SITE_HEADERS]),
                (
                    "用户申请",
                    [
                        ["研究中心编号*", "姓名*", "邮箱*", "电话*", "角色*", "账号操作类型*"],
                        ["SITE-001", "示例用户甲", "user7@example.test", "10000000002", "PI", "账号更新"],
                    ],
                )
            ],
        )

        message = EmailMessage()
        message["From"] = "support@example.test"
        message["Subject"] = "FW: Study Access account update"
        message["Message-ID"] = "<user28@example.test>"
        message["Cc"] = "Project PM <user26@example.test>, support@example.test"
        message.set_content(
            "---------- Forwarded message ---------\n"
            "From: Client Person <user24@example.test>\n"
            "To: support@example.test\n"
            "Project URL/path: demo900\n"
            "ProjectID: DEMO001\n"
        )
        with open(workbook_path, "rb") as file:
            message.add_attachment(
                file.read(),
                maintype="application",
                subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                filename="用户申请.xlsx",
            )

        env_keys = [
            "STUDY_ACCESS_PROJECT_URL",
            "STUDY_ACCESS_PROJECT_PATH",
            "STUDY_ACCESS_PROJECT_ID",
            "PROJECT_ID",
            "STUDY_ACCESS_SUPPORT_FORWARDERS",
            "STUDY_ACCESS_EMAIL_PREVIEW_TO",
            "STUDY_ACCESS_ALLOW_SUPPORT_REPLY",
        ]
        saved_env = {key: os.environ.get(key) for key in env_keys}
        original_run = email_intake.run_commands_with_results
        original_send = email_intake.send_email_message
        sent_messages = []

        def fake_run(commands, timeout):
            results = []
            for command in commands:
                results.append({"command": command, "stdout": "SUCCESS", "stderr": "", "returncode": 0})
            return email_intake.format_command_results(results), results

        def fake_send(to_addresses, subject, body, attachments=None, cc_addresses=None, **_kwargs):
            sent_messages.append(
                {
                    "to": list(to_addresses),
                    "cc": list(cc_addresses or []),
                    "subject": subject,
                    "body": body,
                    "attachments": [str(path) for path in attachments or []],
                }
            )

        try:
            for key in env_keys:
                os.environ.pop(key, None)
            os.environ["STUDY_ACCESS_SUPPORT_FORWARDERS"] = "support@example.test"
            email_intake.run_commands_with_results = fake_run
            email_intake.send_email_message = fake_send
            job = email_intake.process_message_bytes(
                message.as_bytes(),
                os.path.join(temp_dir, "jobs"),
                mode="auto",
                scope="both",
                auto_apply=True,
                send_email_preview=True,
            )
        finally:
            email_intake.run_commands_with_results = original_run
            email_intake.send_email_message = original_send
            for key, value in saved_env.items():
                if value is None:
                    os.environ.pop(key, None)
                else:
                    os.environ[key] = value

        audit_log = open(job["audit_log_path"], "r", encoding="utf-8").read()
        check("auto apply stores forwarder email", job.get("forwarder_email") == "support@example.test", str(job))
        check("auto apply stores client email", job.get("client_email") == "user24@example.test", str(job))
        check("auto apply stores client name", job.get("client_name") == "Client Person", str(job))
        check("auto apply stores project PM cc", job.get("result_cc_recipients") == ["user26@example.test"], str(job))
        check("auto apply result email goes to client", sent_messages and sent_messages[0]["to"] == ["user24@example.test"], str(sent_messages))
        check("auto apply result email cc goes to project PM", sent_messages[0]["cc"] == ["user26@example.test"], str(sent_messages))
        check("auto apply result email does not go to support", "support@example.test" not in sent_messages[0]["to"], str(sent_messages))
        check("auto apply result email greets client", "Dear Client Person," in sent_messages[0]["body"], sent_messages[0]["body"])
        check("auto apply result email uses completion template", "已根据附件请求处理好相应的账号申请/修改操作，请及时联系用户登录邮箱查看。" in sent_messages[0]["body"], sent_messages[0]["body"])
        check("auto apply result email keeps technical detail out of body", "$ " not in sent_messages[0]["body"], sent_messages[0]["body"])
        check("auto apply result email hides job id", "Job ID:" not in sent_messages[0]["body"], sent_messages[0]["body"])
        check("auto apply result email hides status", "Status:" not in sent_messages[0]["body"], sent_messages[0]["body"])
        check("auto apply result email does not mention attachments", "处理明细" not in sent_messages[0]["body"], sent_messages[0]["body"])
        check("full success email omits processed file list", "已处理完整记录的文件" not in sent_messages[0]["body"], sent_messages[0]["body"])
        check("full success email omits zero-skip counts", "信息完整的用户记录" not in sent_messages[0]["body"] and "已跳过的信息不完整用户记录" not in sent_messages[0]["body"], sent_messages[0]["body"])
        check("audit log records email receipt", '"event": "email_received"' in audit_log)
        check("audit log records preview", '"event": "preview_finished"' in audit_log)
        check("audit log records apply", '"event": "apply_finished"' in audit_log)
        check("audit log records client email", '"event": "client_result_email_sent"' in audit_log)


def test_email_intake_isolates_incomplete_workbooks():
    with tempfile.TemporaryDirectory() as temp_dir:
        ready_path = os.path.join(temp_dir, "ready.xlsx")
        incomplete_path = os.path.join(temp_dir, "incomplete.xlsx")
        manual.write_xlsx(
            ready_path,
            [
                ("Sites", [manual.SITE_HEADERS]),
                (
                    "Users",
                    [
                        ["Site ID", "Name", "Email", "Phone", "Role", "Account Operation"],
                        ["CN001", "Ready User", "user29@example.test", "13800000001", "PI", "Account Opening"],
                    ],
                ),
            ],
        )
        manual.write_xlsx(
            incomplete_path,
            [
                ("Sites", [manual.SITE_HEADERS]),
                (
                    "Users",
                    [
                        ["Site ID", "Name", "Email", "Phone", "Role", "Account Operation"],
                        ["CN001", "Complete Row", "user30@example.test", "13800000002", "PI", "Account Opening"],
                        ["CN002", "Missing Phone", "user31@example.test", "", "PI", "Account Opening"],
                    ],
                ),
            ],
        )

        message = EmailMessage()
        message["From"] = "Client Person <user24@example.test>"
        message["Cc"] = "Project PM <user26@example.test>"
        message["Subject"] = "Batch account request"
        message["Message-ID"] = "<user32@example.test>"
        message.set_content("Project URL/path: demo900\nProjectID: TEST001\n")
        for path in (ready_path, incomplete_path):
            with open(path, "rb") as file:
                message.add_attachment(
                    file.read(),
                    maintype="application",
                    subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    filename=os.path.basename(path),
                )

        original_run = email_intake.run_commands_with_results
        original_send = email_intake.send_email_message
        calls = []
        sent_messages = []

        def fake_run(commands, timeout):
            calls.append(commands)
            results = [
                {"command": command, "stdout": "SUCCESS", "stderr": "", "returncode": 0}
                for command in commands
            ]
            return email_intake.format_command_results(results), results

        def fake_send(to_addresses, subject, body, attachments=None, cc_addresses=None, **_kwargs):
            sent_messages.append({
                "to": list(to_addresses),
                "cc": list(cc_addresses or []),
                "subject": subject,
                "body": body,
                "attachments": list(attachments or []),
            })

        try:
            email_intake.run_commands_with_results = fake_run
            email_intake.send_email_message = fake_send
            job = email_intake.process_message_bytes(
                message.as_bytes(),
                os.path.join(temp_dir, "jobs"),
                mode="setup",
                scope="both",
                auto_apply=True,
                send_email_preview=True,
            )
        finally:
            email_intake.run_commands_with_results = original_run
            email_intake.send_email_message = original_send

        flattened = [part for call in calls for command in call for part in command]
        report = job.get("excel_validation") or {}
        incomplete = report.get("incomplete_files") or []
        check("complete workbook is still auto-applied", job.get("status") == "applied" and len(calls) == 2, str(job))
        check("one workbook contains a skipped incomplete row", len(incomplete) == 1 and incomplete[0].get("filename") == "incomplete.xlsx", str(report))
        check("mixed workbook remains eligible for its complete row", any(str(part).endswith("incomplete.xlsx") for part in flattened), str(calls))
        check("complete workbook is passed to STUDY/ROLE", any(str(part).endswith("ready.xlsx") for part in flattened), str(calls))
        check("email batch commands enable row-level skipping", "--skip-blocked-rows" in flattened, str(calls))
        missing_issues = incomplete[0].get("issues") if incomplete else []
        check(
            "missing row reports sheet row and telephone",
            bool(missing_issues)
            and missing_issues[0].get("row") == 3
            and "telephone" in missing_issues[0].get("fields", []),
            str(missing_issues),
        )
        check("complete row in the mixed workbook is allowed to continue", any(str(part).endswith("incomplete.xlsx") for part in flattened), str(calls))
        check("client receives one consolidated result email", len(sent_messages) == 1 and sent_messages[0]["to"] == ["user24@example.test"], str(sent_messages))
        check("project PM is copied on validation reminder", sent_messages[0]["cc"] == ["user26@example.test"], str(sent_messages))
        check("client email uses concise partial-result template", "已成功处理附件内信息完整的 2 条用户账号申请/更新记录。" in sent_messages[0]["body"] and "以下 1 条用户信息缺失，暂未处理，请补充：" in sent_messages[0]["body"], sent_messages[0]["body"])
        check("client email omits file names", "ready.xlsx" not in sent_messages[0]["body"] and "incomplete.xlsx" not in sent_messages[0]["body"], sent_messages[0]["body"])
        check("client email identifies the missing phone row", "第3行" in sent_messages[0]["body"] and "联系电话" in sent_messages[0]["body"], sent_messages[0]["body"])
        check("client reminder contains no technical command output", "$ " not in sent_messages[0]["body"] and "Job ID:" not in sent_messages[0]["body"], sent_messages[0]["body"])


def test_email_validation_six_complete_four_incomplete():
    with tempfile.TemporaryDirectory() as temp_dir:
        path = os.path.join(temp_dir, "ten-users.xlsx")
        headers = ["Site ID", "Name", "Email", "Phone", "Role", "Account Operation"]
        rows = [headers]
        for index in range(10):
            rows.append([
                f"CN{index + 1:03d}",
                f"User {index + 1}",
                f"user{index + 1}@example.com",
                f"1380000000{index}",
                "PI",
                "Account Opening",
            ])
        rows[7][0] = ""
        rows[8][1] = ""
        rows[9][2] = "bad @ example.com"
        rows[10][5] = ""
        manual.write_xlsx(
            path,
            [
                ("Sites", [manual.SITE_HEADERS]),
                ("Users", rows),
            ],
        )

        result = email_intake.validate_excel_attachment(path, "auto")
        fields = [issue.get("fields") for issue in result.get("issues") or []]
        check("ten-row workbook keeps six complete users", result.get("complete_user_count") == 6, str(result))
        check("ten-row workbook reports four incomplete users", len(result.get("issues") or []) == 4, str(result))
        check("mixed ten-row workbook remains processable", result.get("processable") is True, str(result))
        check("missing center code is reported", ["siteCode"] in fields, str(fields))
        check("missing name is reported", ["name"] in fields, str(fields))
        check("email containing spaces is reported", ["emailWhitespace"] in fields, str(fields))
        check("missing account operation is reported", ["operation"] in fields, str(fields))
        report = {
            "files": [result],
            "incomplete_files": [result],
        }
        email_body = email_intake.build_apply_email_text(
            {"status": "applied", "client_name": "Client", "excel_validation": report},
            "",
        )
        check("result email reports six processed users", "已成功处理附件内信息完整的 6 条用户账号申请/更新记录。" in email_body, email_body)
        check("result email reports four incomplete users", "以下 4 条用户信息缺失，暂未处理，请补充：" in email_body, email_body)
        check("result email omits legacy file and count blocks", "已处理完整记录的文件" not in email_body and "信息完整的用户记录" not in email_body and "已跳过的信息不完整用户记录" not in email_body, email_body)
        check("result email explains that email cannot contain spaces", "邮箱中不能包含空格" in email_body, email_body)


def test_email_approval_parser():
    check("approval email parser accepts exact job command", email_intake.extract_approval_job_id("/approve abc123DEF_45") == "abc123DEF_45")
    check("approval email parser rejects no job id", email_intake.extract_approval_job_id("/approve") == "")
    check("approval email parser rejects uppercase approve", email_intake.extract_approval_job_id("/APPROVE abc123") == "")
    check("approval email parser rejects quoted approval line", email_intake.extract_approval_job_id("> /approve abc123") == "")
    check(
        "management gate detects restricted request",
        email_intake.management_gate_match("账号申请", "请完全不留存个人信息")
        == "privacy_instruction",
    )
    check("management gate ignores generic privacy footer", not email_intake.management_gate_match("账号申请", "Please see our privacy policy."))


def test_uat_prefix_routing_is_isolated():
    unchanged = email_intake.resolve_project_details({}, {"project_id": "DEMO300", "project_path": "DEMO300"})
    check("stable project keeps its existing path", unchanged["project_url"] == "DEMO300", str(unchanged))

    for project_id, expected_path in (("DEMO100", "demo100"), ("DEMO200", "demo200"), ("DEMO201", "demo201")):
        stable = email_intake.resolve_project_details({}, {"project_id": project_id, "project_path": project_id})
        stable_job = dict(stable, mode="update", scope="both", excel_files=["request.xlsx"])
        stable_joined = " ".join(" ".join(command) for command in email_intake.preview_commands(stable_job, True, "both"))
        check(
            f"{project_id} keeps production route and login",
            stable["project_url"].lower() == expected_path
            and "--username DEMO-Admin" in stable_joined
            and "--admin-username-env" not in stable_joined
            and "--admin-password-env" not in stable_joined,
            str(stable),
        )

    ubr = email_intake.resolve_project_details({}, {"project_id": "UDEMO401", "project_path": "UDEMO401"})
    u_dash = email_intake.resolve_project_details({}, {"project_id": "U-DEMO401", "project_path": "U-DEMO401"})
    u_demo400 = email_intake.resolve_project_details({}, {"project_id": "U-DEMO400", "project_path": "U-DEMO400"})
    u_demo400_suffix = email_intake.resolve_project_details({}, {"project_id": "U-DEMO400-01", "project_path": "U-DEMO400-01"})
    u_suffix = email_intake.resolve_project_details({}, {"project_id": "U-DEMO401-301-01", "project_path": "U-DEMO401-301-01"})
    prod_suffix = email_intake.resolve_project_details({}, {"project_id": "DEMO100-201-01", "project_path": "DEMO100-201-01"})
    new_uat = email_intake.resolve_project_details({}, {"project_id": "U-ABC123-201-01", "project_path": "U-ABC123-201-01"})
    new_br_style_uat = email_intake.resolve_project_details({}, {"project_id": "U-BR005-301-01", "project_path": "U-BR005-301-01"})
    invalid_legacy_uat = email_intake.resolve_project_details({}, {"project_id": "UABC123", "project_path": "UABC123"})
    new_prod = email_intake.resolve_project_details({}, {"project_id": "ABC123", "project_path": "ABC123"})
    new_prod_suffix = email_intake.resolve_project_details({}, {"project_id": "ABC123-201-01", "project_path": "ABC123-201-01"})
    malformed_uat = email_intake.resolve_project_details({}, {"project_id": "U-ABC_123", "project_path": "U-ABC_123"})
    empty_suffix_uat = email_intake.resolve_project_details({}, {"project_id": "U-ABC123--01", "project_path": "U-ABC123--01"})
    check("UDEMO401 without U- is treated as PROD", ubr["environment"] == "prod" and ubr["project_url"] == "udemo401" and ubr["project_id"] == "UDEMO401", str(ubr))
    check("U-DEMO401 preserves its hyphen and resolves to demo401", u_dash["project_url"] == "demo401" and u_dash["trial_id"] == "U-DEMO401" and u_dash["project_id"] == "U-DEMO401", str(u_dash))
    check("U-DEMO400 resolves to its confirmed lowercase path", u_demo400["project_url"] == "demo400" and u_demo400["trial_id"] == "U-DEMO400" and u_demo400["project_id"] == "U-DEMO400", str(u_demo400))
    check("U-DEMO400 suffix stays in the complete ID but not the URL", u_demo400_suffix["project_url"] == "demo400" and u_demo400_suffix["project_id"] == "U-DEMO400-01" and u_demo400_suffix["project_suffix"] == "01", str(u_demo400_suffix))
    check("UAT suffixes stay in the complete ID but not the URL", u_suffix["project_url"] == "demo401" and u_suffix["project_id"] == "U-DEMO401-301-01" and u_suffix["project_suffix"] == "301-01", str(u_suffix))
    check("PROD suffixes stay in the complete ID but not the URL", prod_suffix["project_url"] == "demo100" and prod_suffix["project_id"] == "DEMO100-201-01" and prod_suffix["project_suffix"] == "201-01", str(prod_suffix))
    check("new standard UAT project auto-resolves lowercase URL", new_uat["project_url"] == "abc123" and new_uat["project_id"] == "U-ABC123-201-01" and new_uat["project_suffix"] == "201-01", str(new_uat))
    check("new BR-style UAT project uses only its first code segment for URL", new_br_style_uat["project_url"] == "br005" and new_br_style_uat["project_id"] == "U-BR005-301-01" and new_br_style_uat["project_suffix"] == "301-01", str(new_br_style_uat))
    check("UABC123 without U- is treated as PROD", invalid_legacy_uat["environment"] == "prod" and invalid_legacy_uat["project_url"] == "uabc123", str(invalid_legacy_uat))
    check("new exact PROD project auto-resolves lowercase URL", new_prod["project_url"] == "abc123" and new_prod["project_id"] == "ABC123", str(new_prod))
    check("new suffixed PROD project auto-resolves lowercase URL", new_prod_suffix["project_url"] == "abc123" and new_prod_suffix["project_suffix"] == "201-01", str(new_prod_suffix))
    check("U- project with underscore still follows UAT prefix rule", malformed_uat["environment"] == "uat" and malformed_uat["project_url"] == "abc_123" and not malformed_uat["project_conflicts"], str(malformed_uat))
    check("U- project with repeated suffix separator still follows UAT prefix rule", empty_suffix_uat["environment"] == "uat" and empty_suffix_uat["project_url"] == "abc123" and empty_suffix_uat["project_suffix"] == "-01" and not empty_suffix_uat["project_conflicts"], str(empty_suffix_uat))

    saved_overrides = os.environ.get("STUDY_ACCESS_PROJECT_ROUTES")
    os.environ["STUDY_ACCESS_PROJECT_ROUTES"] = "ZX900=custom-zx"
    try:
        overridden = email_intake.resolve_project_details({}, {"project_id": "U-ZX900-02", "project_path": "U-ZX900-02"})
    finally:
        if saved_overrides is None:
            os.environ.pop("STUDY_ACCESS_PROJECT_ROUTES", None)
        else:
            os.environ["STUDY_ACCESS_PROJECT_ROUTES"] = saved_overrides
    check("special route override wins over lowercase inference", overridden["project_url"] == "custom-zx" and overridden["project_suffix"] == "02", str(overridden))

    check("configured UAT route has complete routing config", not email_intake.missing_live_config("both", u_dash), str(email_intake.missing_live_config("both", u_dash)))

    job = dict(u_suffix, mode="update", scope="both", excel_files=["request.xlsx"])
    commands = email_intake.preview_commands(job, True, "both")
    joined = " ".join(" ".join(command) for command in commands)
    check("UAT commands use resolved path and preserve the complete ID", "--project-url demo401" in joined and "--project-id U-DEMO401-301-01" in joined, joined)
    check("UAT commands use unified UAT username", "--username U-DEMO-Admin" in joined, joined)
    check("UAT commands do not use project-specific credential references", "--admin-username-env" not in joined and "--admin-password-env" not in joined, joined)
    check("UAT commands do not fall back to production username", "--username DEMO-Admin" not in joined, joined)

    new_uat_job = dict(new_uat, mode="update", scope="both", excel_files=["request.xlsx"])
    new_uat_joined = " ".join(" ".join(command) for command in email_intake.preview_commands(new_uat_job, True, "both"))
    check("new inferred UAT command uses lowercase URL and full ID", "--project-url abc123" in new_uat_joined and "--project-id U-ABC123-201-01" in new_uat_joined, new_uat_joined)
    check("new inferred UAT command uses only UAT admin", "--username U-DEMO-Admin" in new_uat_joined and "--username DEMO-Admin" not in new_uat_joined, new_uat_joined)

    new_prod_job = dict(new_prod, mode="update", scope="both", excel_files=["request.xlsx"])
    new_prod_joined = " ".join(" ".join(command) for command in email_intake.preview_commands(new_prod_job, True, "both"))
    check("new inferred PROD command uses lowercase URL and full ID", "--project-url abc123" in new_prod_joined and "--project-id ABC123" in new_prod_joined, new_prod_joined)
    check("new inferred PROD command uses only production admin", "--username DEMO-Admin" in new_prod_joined and "--username U-DEMO-Admin" not in new_prod_joined, new_prod_joined)

    legacy_job = dict(unchanged, mode="update", scope="both", excel_files=["request.xlsx"])
    legacy_commands = email_intake.preview_commands(legacy_job, True, "both")
    legacy_joined = " ".join(" ".join(command) for command in legacy_commands)
    check("stable project retains universal username command", "--username DEMO-Admin" in legacy_joined, legacy_joined)


def main():
    tests = [
        test_approval_is_exact,
        test_site_code_exact_matching,
        test_site_scope_resolution,
        test_site_name_and_study_username_exact_matching,
        test_role_username_exact_matching,
        test_username_generation_rules,
        test_update_note_and_direct_payload_rules,
        test_disable_finds_legacy_username_format,
        test_update_username_legacy_candidates_are_deterministic,
        test_study_update_missing_user_blocks_before_pending_site,
        test_role_legacy_lookup_only_applies_to_update_rows,
        test_role_ambiguous_legacy_matches_block_all_writes,
        test_role_main_blocks_before_patch_or_delete,
        test_manual_workbook_roundtrip,
        test_orange_template_layout_roundtrip,
        test_site_optional_fields_default_to_dash,
        test_existing_site_updates_preserve_blank_optional_fields,
        test_application_email_format_validation,
        test_row_level_skip_prepares_only_complete_users,
        test_email_intake_job_roundtrip,
        test_email_intake_auto_apply,
        test_email_intake_management_gate,
        test_email_intake_auto_apply_blocks_missing_project_config,
        test_email_intake_project_conflict_blocks_auto_apply,
        test_email_client_recipient_detection,
        test_email_intake_auto_apply_emails_client_and_audits,
        test_email_intake_isolates_incomplete_workbooks,
        test_email_validation_six_complete_four_incomplete,
        test_email_approval_parser,
        test_uat_prefix_routing_is_isolated,
    ]
    for test in tests:
        print(f"\n== {test.__name__} ==")
        test()
    if FAILURES:
        print(f"\nSelf-test failed: {len(FAILURES)} failure(s)")
        for failure in FAILURES:
            print(f"- {failure}")
        sys.exit(1)
    print("\nSelf-test passed.")


if __name__ == "__main__":
    main()
