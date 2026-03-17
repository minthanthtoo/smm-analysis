import time
from pathlib import Path
from typing import Any

from flask import Flask, abort, jsonify, render_template, request

from smm_services import *
from smm_styled import build_main_sheet_html_cached, clear_styled_cache, parse_clamped_int
from smm_runtime import resolve_runtime_host_port

app = Flask(__name__)


def clear_workbook_caches() -> None:
    clear_data_caches()
    clear_styled_cache()


def canonical_manageable_role(raw_role: Any) -> str | None:
    role = normalize_token(raw_role)
    if role in {"rsm", "regionalmanager", "regional_manager"}:
        return ROLE_RSM
    if role == "asm":
        return ROLE_ASM
    if role in {"user", "staff"}:
        return ROLE_USER
    return None


@app.route("/")
def index() -> str:
    return render_template("index.html", static_version=int(time.time()))


@app.route("/api/access/context")
def api_access_context():
    principal = principal_from_request()
    return jsonify(access_context_payload(principal))


@app.route("/api/access/assign-rsm", methods=["POST"])
def api_access_assign_rsm():
    principal = principal_from_request()
    if not principal_can_manage_rsm(principal):
        return jsonify({"error": "Only Owner can assign RSM users."}), 403

    payload = request_json_payload()
    username = normalize_username(payload.get("username") or request.form.get("username"))
    if not username:
        return jsonify({"error": "Missing required field: username"}), 400

    display_name = payload.get("display_name") or request.form.get("display_name")
    regions = normalize_region_list(payload.get("regions") or request.form.get("regions"))

    access = load_access_control()
    ensure_access_user(access, username, role=ROLE_RSM, display_name=display_name)
    access["users"][username]["role"] = ROLE_RSM
    access.setdefault("rsm_regions", {})[username] = regions
    save_access_control(access)

    refreshed = resolve_principal(principal["username"], principal["selected_region"])
    return jsonify(access_context_payload(refreshed))


@app.route("/api/access/assign-user-to-rsm", methods=["POST"])
def api_access_assign_user_to_rsm():
    principal = principal_from_request()
    if not principal_can_manage_rsm(principal):
        return jsonify({"error": "Only Owner can assign users to RSM."}), 403

    payload = request_json_payload()
    username = normalize_username(payload.get("username") or request.form.get("username"))
    rsm_username = normalize_username(payload.get("rsm_username") or request.form.get("rsm_username"))
    if not username or not rsm_username:
        return jsonify({"error": "Missing required fields: username and rsm_username"}), 400

    access = load_access_control()
    ensure_access_user(access, username, role=ROLE_USER)
    ensure_access_user(access, rsm_username, role=ROLE_RSM)
    if normalize_user_role(access["users"][rsm_username].get("role")) != ROLE_RSM:
        return jsonify({"error": f"Target user is not an RSM: {rsm_username}"}), 400

    access.setdefault("user_to_rsm", {})[username] = rsm_username
    save_access_control(access)

    refreshed = resolve_principal(principal["username"], principal["selected_region"])
    return jsonify(access_context_payload(refreshed))


@app.route("/api/access/set-user-role", methods=["POST"])
def api_access_set_user_role():
    principal = principal_from_request()
    if not principal_can_manage_rsm(principal):
        return jsonify({"error": "Only Owner can add users or switch roles."}), 403

    payload = request_json_payload()
    username = normalize_username(payload.get("username") or request.form.get("username"))
    if not username:
        return jsonify({"error": "Missing required field: username"}), 400
    if username == "owner":
        return jsonify({"error": "Owner role cannot be modified."}), 400

    target_role = canonical_manageable_role(payload.get("role") or request.form.get("role"))
    if not target_role:
        return jsonify({"error": "Invalid role. Allowed roles: user, rsm, asm."}), 400

    display_name = payload.get("display_name") or request.form.get("display_name")
    requested_rsm = normalize_username(payload.get("rsm_username") or request.form.get("rsm_username"))

    access = load_access_control()
    user_record = ensure_access_user(access, username, role=target_role, display_name=display_name)
    user_record["role"] = target_role

    rsm_regions = access.setdefault("rsm_regions", {})
    user_to_rsm = access.setdefault("user_to_rsm", {})
    asm_townships = access.setdefault("asm_townships", {})

    if target_role != ROLE_RSM:
        rsm_regions.pop(username, None)
        for candidate, manager in list(user_to_rsm.items()):
            if candidate != username and normalize_username(manager) == username:
                user_to_rsm.pop(candidate, None)

    if target_role != ROLE_ASM:
        asm_townships.pop(username, None)

    if target_role == ROLE_RSM:
        regions = normalize_region_list(payload.get("regions") or request.form.get("regions"))
        user_to_rsm.pop(username, None)
        if regions:
            rsm_regions[username] = regions
        else:
            rsm_regions.setdefault(username, [])
    else:
        manager_rsm = requested_rsm or normalize_username(user_to_rsm.get(username))
        if manager_rsm:
            if manager_rsm == username:
                return jsonify({"error": "User cannot be mapped to self as RSM."}), 400
            ensure_access_user(access, manager_rsm, role=ROLE_RSM)
            if normalize_viewer_role(access["users"][manager_rsm].get("role")) != ROLE_RSM:
                return jsonify({"error": f"Target user is not an RSM: {manager_rsm}"}), 400
            user_to_rsm[username] = manager_rsm
        else:
            user_to_rsm.pop(username, None)

        if target_role == ROLE_ASM:
            manager_rsm = normalize_username(user_to_rsm.get(username))
            if not manager_rsm:
                return jsonify({"error": "ASM role requires rsm_username (or existing RSM mapping)."}), 400
            asm_townships.setdefault(username, {})

    save_access_control(access)

    refreshed = resolve_principal(principal["username"], principal["selected_region"])
    return jsonify(access_context_payload(refreshed))


@app.route("/api/access/assign-asm", methods=["POST"])
def api_access_assign_asm():
    principal = principal_from_request()
    if not principal_can_manage_asm(principal):
        return jsonify({"error": "Only Owner or RSM can assign ASM users."}), 403

    payload = request_json_payload()
    asm_username = normalize_username(
        payload.get("asm_username") or payload.get("username") or request.form.get("asm_username")
    )
    if not asm_username:
        return jsonify({"error": "Missing required field: asm_username"}), 400

    if principal["role"] == ROLE_OWNER:
        target_rsm = normalize_username(
            payload.get("rsm_username") or request.form.get("rsm_username")
        )
        if not target_rsm:
            return jsonify({"error": "Owner must provide rsm_username."}), 400
    else:
        target_rsm = principal["username"]

    region = normalize_region_token(payload.get("region") or request.form.get("region"))
    if not region:
        return jsonify({"error": "Missing required field: region"}), 400

    access = load_access_control()
    ensure_access_user(access, target_rsm, role=ROLE_RSM)
    if normalize_user_role(access["users"][target_rsm].get("role")) != ROLE_RSM:
        return jsonify({"error": f"Target manager is not an RSM: {target_rsm}"}), 400

    target_rsm_regions = normalize_region_list(access.get("rsm_regions", {}).get(target_rsm, []))
    if principal["role"] == ROLE_RSM and region not in set(principal.get("rsm_regions", [])):
        return jsonify({"error": f"RSM cannot assign ASM outside managed regions: {region}"}), 403
    if region not in target_rsm_regions:
        target_rsm_regions.append(region)
        access.setdefault("rsm_regions", {})[target_rsm] = sorted(set(target_rsm_regions))

    display_name = payload.get("display_name") or request.form.get("display_name")
    ensure_access_user(access, asm_username, role=ROLE_ASM, display_name=display_name)
    access["users"][asm_username]["role"] = ROLE_ASM
    access.setdefault("user_to_rsm", {})[asm_username] = target_rsm

    towns_raw = payload.get("townships")
    if towns_raw is None:
        towns_raw = request.form.get("townships")
    townships = normalize_township_list(towns_raw)
    access.setdefault("asm_townships", {}).setdefault(asm_username, {})[region] = townships
    save_access_control(access)

    refreshed = resolve_principal(principal["username"], principal["selected_region"])
    return jsonify(access_context_payload(refreshed))


@app.route("/api/access/set-asm-townships", methods=["POST"])
def api_access_set_asm_townships():
    principal = principal_from_request()
    if not principal_can_manage_asm(principal):
        return jsonify({"error": "Only Owner or RSM can set ASM township permissions."}), 403

    payload = request_json_payload()
    asm_username = normalize_username(
        payload.get("asm_username") or payload.get("username") or request.form.get("asm_username")
    )
    region = normalize_region_token(payload.get("region") or request.form.get("region"))
    if not asm_username or not region:
        return jsonify({"error": "Missing required fields: asm_username and region"}), 400

    access = load_access_control()
    manager_rsm = normalize_username(access.get("user_to_rsm", {}).get(asm_username))
    if not manager_rsm:
        return jsonify({"error": f"ASM is not mapped to any RSM: {asm_username}"}), 400

    if principal["role"] == ROLE_RSM and manager_rsm != principal["username"]:
        return jsonify({"error": "RSM can only manage own ASM users."}), 403
    if principal["role"] == ROLE_RSM and region not in set(principal.get("rsm_regions", [])):
        return jsonify({"error": f"RSM cannot assign townships outside managed regions: {region}"}), 403

    towns_raw = payload.get("townships")
    if towns_raw is None:
        towns_raw = request.form.get("townships")
    townships = normalize_township_list(towns_raw)
    access.setdefault("asm_townships", {}).setdefault(asm_username, {})[region] = townships
    save_access_control(access)

    refreshed = resolve_principal(principal["username"], principal["selected_region"])
    return jsonify(access_context_payload(refreshed))


@app.route("/api/files")
def api_files():
    principal = principal_from_request()
    region_scope = principal["selected_region"]
    allowed_regions = set(principal["allowed_regions"])
    role = principal["role"]

    visible_entries: list[dict[str, Any]] = []
    for entry in principal["entries"]:
        region = str(entry["region"])
        if role != ROLE_OWNER and region not in allowed_regions:
            continue
        if region_scope != ALL_REGION_TOKEN and region != region_scope:
            continue
        visible_entries.append(entry)

    files_payload = []
    rsm_regions = set(principal.get("rsm_regions", []))
    for entry in visible_entries:
        file_region = str(entry["region"])
        uploaded = is_uploaded_workbook_entry(entry)
        can_modify = False
        if uploaded and role == ROLE_OWNER:
            can_modify = True
        elif uploaded and role == ROLE_RSM and file_region in rsm_regions:
            can_modify = True
        files_payload.append(
            {
                "name": str(entry["name"]),
                "region": file_region,
                "uploaded": uploaded,
                "can_update": can_modify,
                "can_delete": can_modify,
            }
        )

    return jsonify(
        {
            "files": files_payload,
            "selected_region": region_scope,
            "current_user": principal["username"],
            "viewer_role": principal["role"],
        }
    )


@app.route("/api/files/<path:filename>", methods=["PATCH"])
def api_file_update_region(filename: str):
    principal = principal_from_request()
    entries = discover_workbook_entries()
    entry = workbook_entry_by_name(entries, filename)
    if not entry:
        return jsonify({"error": f"Unknown workbook: {filename}"}), 404
    if not is_uploaded_workbook_entry(entry):
        return jsonify({"error": "Only uploaded files can be updated."}), 403

    payload = request_json_payload()
    target_region = normalize_region_token(payload.get("region") or request.form.get("region"))
    if not target_region:
        return jsonify({"error": "Missing required field: region"}), 400

    source_region = str(entry["region"])
    if not principal_can_manage_region(principal, source_region):
        return jsonify({"error": "No permission to update this file."}), 403
    if not principal_can_manage_region(principal, target_region):
        return jsonify({"error": "No permission to move file to this region."}), 403

    registry = load_workbook_registry()
    registry[str(entry["name"])] = {"region": target_region}
    save_workbook_registry(registry)
    clear_workbook_caches()

    return jsonify(
        {
            "file": str(entry["name"]),
            "region": target_region,
            "updated": True,
        }
    )


@app.route("/api/files/<path:filename>", methods=["DELETE"])
def api_file_delete(filename: str):
    principal = principal_from_request()
    entries = discover_workbook_entries()
    entry = workbook_entry_by_name(entries, filename)
    if not entry:
        return jsonify({"error": f"Unknown workbook: {filename}"}), 404
    if not is_uploaded_workbook_entry(entry):
        return jsonify({"error": "Only uploaded files can be deleted."}), 403

    file_region = str(entry["region"])
    if not principal_can_manage_region(principal, file_region):
        return jsonify({"error": "No permission to delete this file."}), 403

    path = Path(entry["path"])
    try:
        path.unlink()
    except OSError as exc:
        return jsonify({"error": f"Failed to delete file: {exc}"}), 500

    registry = load_workbook_registry()
    registry.pop(str(entry["name"]), None)
    save_workbook_registry(registry)
    clear_workbook_caches()
    return jsonify({"deleted": str(entry["name"])})


@app.route("/api/workbook/rename-sheet", methods=["POST"])
def api_workbook_rename_sheet():
    principal = principal_from_request()
    if principal["role"] not in {ROLE_OWNER, ROLE_RSM}:
        return jsonify({"error": "Only Owner or RSM can rename workbook sheets."}), 403

    payload = request_json_payload()
    workbook_kind = str(payload.get("workbook") or request.form.get("workbook") or "").strip().lower()
    if workbook_kind not in {"main", "reference"}:
        return jsonify({"error": "Missing or invalid field: workbook (main/reference)."}), 400

    old_sheet_name = payload.get("old_sheet_name") or request.form.get("old_sheet_name")
    new_sheet_name = payload.get("new_sheet_name") or request.form.get("new_sheet_name")

    try:
        selection = workbook_selection_for_principal(
            principal,
            payload.get("main") or request.form.get("main") or request.args.get("main"),
            payload.get("reference") or request.form.get("reference") or request.args.get("reference"),
        )
    except FileNotFoundError as exc:
        return jsonify({"error": str(exc)}), 404
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400

    target_path = selection["main_path"] if workbook_kind == "main" else selection["reference_path"]
    entry = workbook_entry_by_path(principal["entries"], target_path)
    if not entry:
        return jsonify({"error": f"Unknown workbook entry: {target_path.name}"}), 404

    workbook_region = str(entry.get("region") or "")
    if not principal_can_manage_region(principal, workbook_region):
        return jsonify({"error": "No permission to edit sheets in this workbook."}), 403

    try:
        renamed_from, renamed_to = rename_sheet_in_workbook(
            target_path,
            str(old_sheet_name or ""),
            str(new_sheet_name or ""),
        )
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except OSError as exc:
        return jsonify({"error": f"Failed to save workbook: {exc}"}), 500

    clear_workbook_caches()
    return jsonify(
        {
            "updated": True,
            "workbook": target_path.name,
            "workbook_kind": workbook_kind,
            "region": workbook_region,
            "old_sheet_name": renamed_from,
            "new_sheet_name": renamed_to,
        }
    )


@app.route("/api/workbooks")
def api_workbooks():
    principal = principal_from_request()
    try:
        selection = workbook_selection_for_principal(
            principal,
            None,
            None,
        )
    except FileNotFoundError as exc:
        return jsonify(
            {
                "error": str(exc),
                "workbooks": [],
                "regions": (
                    principal["allowed_regions"]
                    if not principal["allow_all_regions"]
                    else principal["regions_universe"]
                ),
                "viewer_role": principal["role"],
                "selected_region": principal["selected_region"],
                "current_user": principal["username"],
            }
        )

    return jsonify(
        {
            "workbooks": selection["available"],
            "regions": principal["allowed_regions"] if not principal["allow_all_regions"] else principal["regions_universe"],
            "viewer_role": principal["role"],
            "selected_region": selection["selected_region"],
            "current_user": principal["username"],
            "default_main": selection["default_main"],
            "default_reference": selection["default_reference"],
        }
    )


@app.route("/api/upload-workbooks", methods=["POST"])
def api_upload_workbooks():
    principal = principal_from_request()
    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "No Excel files were uploaded."}), 400

    if principal["role"] not in {ROLE_OWNER, ROLE_RSM}:
        return jsonify({"error": "Only Owner or RSM can upload files."}), 403

    upload_region = request.form.get("upload_region")

    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    saved_files: list[str] = []
    uploaded_regions: list[dict[str, str]] = []
    skipped_files: list[str] = []
    denied_files: list[dict[str, str]] = []
    registry = load_workbook_registry()

    for file_storage in files:
        original_name = (file_storage.filename or "").strip()
        if not original_name:
            continue

        suffix = Path(original_name).suffix.lower()
        if suffix not in SUPPORTED_EXTENSIONS:
            skipped_files.append(original_name)
            continue

        region = upload_region_for_file(upload_region, original_name)
        if not principal_can_upload_to_region(principal, region):
            denied_files.append(
                {
                    "name": original_name,
                    "reason": f"No upload permission for region {region}",
                }
            )
            continue

        output_name = safe_uploaded_filename(original_name)
        file_storage.save(UPLOAD_DIR / output_name)
        saved_files.append(output_name)
        uploaded_regions.append({"file": output_name, "region": region})
        registry[output_name] = {"region": region}

    if not saved_files:
        status_code = 403 if denied_files else 400
        return (
            jsonify(
                {
                    "error": "No files were uploaded.",
                    "skipped_files": skipped_files,
                    "denied_files": denied_files,
                }
            ),
            status_code,
        )

    save_workbook_registry(registry)
    clear_workbook_caches()
    refreshed_principal = resolve_principal(
        requested_user=principal["username"],
        requested_region=request.form.get("region"),
    )
    selection = workbook_selection_for_principal(
        refreshed_principal,
        None,
        None,
    )
    return jsonify(
        {
            "uploaded_files": saved_files,
            "uploaded_regions": uploaded_regions,
            "skipped_files": skipped_files,
            "denied_files": denied_files,
            "workbooks": selection["available"],
            "regions": (
                refreshed_principal["allowed_regions"]
                if not refreshed_principal["allow_all_regions"]
                else refreshed_principal["regions_universe"]
            ),
            "viewer_role": refreshed_principal["role"],
            "selected_region": selection["selected_region"],
            "current_user": refreshed_principal["username"],
            "default_main": selection["default_main"],
            "default_reference": selection["default_reference"],
        }
    )


@app.route("/api/sheets")
def api_sheets():
    principal = principal_from_request()
    try:
        selection = workbook_selection_for_principal(
            principal,
            request.args.get("main"),
            request.args.get("reference"),
        )
    except FileNotFoundError as exc:
        return jsonify({"error": str(exc)}), 404
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400

    data = load_viewer_data(selection["main_path"], selection["reference_path"])
    data = filter_data_by_allowed_townships(data, principal["allowed_townships"])
    return jsonify(
        {
            "main_workbook": data["main_workbook"],
            "reference_workbook": data["reference_workbook"],
            "version": data["version"],
            "available_workbooks": selection["available"],
            "main_sheet_names": data["main_sheet_names"],
            "reference_sheet_names": data["reference_sheet_names"],
            "regions": (
                principal["allowed_regions"]
                if not principal["allow_all_regions"]
                else principal["regions_universe"]
            ),
            "viewer_role": principal["role"],
            "selected_region": selection["selected_region"],
            "current_user": principal["username"],
            "sheets": data["sheet_index"],
            "main_sheet_tabs": data["main_sheet_tabs"],
            "reference_sheet_tabs": data["reference_sheet_tabs"],
        }
    )


@app.route("/api/sheet/<canonical>")
def api_sheet(canonical: str):
    principal = principal_from_request()
    try:
        selection = workbook_selection_for_principal(
            principal,
            request.args.get("main"),
            request.args.get("reference"),
        )
    except FileNotFoundError as exc:
        return jsonify({"error": str(exc)}), 404
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400

    data = filter_data_by_allowed_townships(
        load_viewer_data(selection["main_path"], selection["reference_path"]),
        principal["allowed_townships"],
    )
    if canonical not in data["main"] and canonical not in data["reference"]:
        abort(404, description=f"Unknown sheet: {canonical}")

    return jsonify(
        {
            "canonical": canonical,
            "version": data["version"],
            "main": data["main"].get(canonical),
            "reference": data["reference"].get(canonical),
        }
    )


@app.route("/api/main-styled/<canonical>")
def api_main_styled(canonical: str):
    principal = principal_from_request()
    try:
        selection = workbook_selection_for_principal(
            principal,
            request.args.get("main"),
            request.args.get("reference"),
        )
    except FileNotFoundError as exc:
        return jsonify({"error": str(exc)}), 404
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400

    data = filter_data_by_allowed_townships(
        load_viewer_data(selection["main_path"], selection["reference_path"]),
        principal["allowed_townships"],
    )
    if canonical not in data["main"]:
        abort(404, description=f"Unknown sheet: {canonical}")

    month_keys_csv = request.args.get("month_keys", "")
    mode = request.args.get("mode", "past_months")
    if mode not in {"past_months", "same_month_years"}:
        mode = "past_months"
    n_value = parse_clamped_int(request.args.get("n"), default=6, min_value=1, max_value=60)
    month_value = parse_clamped_int(request.args.get("month"), default=0, min_value=0, max_value=12)
    main_stat = selection["main_path"].stat()
    sheet_name = data["main"][canonical]["sheet_name"]
    html_payload = build_main_sheet_html_cached(
        str(selection["main_path"]),
        main_stat.st_mtime_ns,
        sheet_name,
        month_keys_csv,
        mode,
        n_value,
        month_value,
    )
    return jsonify(
        {
            "version": data["version"],
            "canonical": canonical,
            "filterable": True,
            "sheet_name": html_payload["sheet_name"],
            "row_count": html_payload["row_count"],
            "col_count": html_payload["col_count"],
            "frozen_count": html_payload["frozen_count"],
            "frozen_rows": html_payload.get("frozen_rows", 0),
            "frozen_columns": html_payload.get("frozen_columns", html_payload["frozen_count"]),
            "selected_month_labels": html_payload["selected_month_labels"],
            "available_months": html_payload["available_months"],
            "hidden_rows_skipped": html_payload.get("hidden_rows_skipped", 0),
            "hidden_columns_skipped": html_payload.get("hidden_columns_skipped", 0),
            "html": html_payload["html"],
        }
    )


@app.route("/api/main-styled-sheet")
def api_main_styled_sheet():
    principal = principal_from_request()
    try:
        selection = workbook_selection_for_principal(
            principal,
            request.args.get("main"),
            request.args.get("reference"),
        )
    except FileNotFoundError as exc:
        return jsonify({"error": str(exc)}), 404
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400

    data = filter_data_by_allowed_townships(
        load_viewer_data(selection["main_path"], selection["reference_path"]),
        principal["allowed_townships"],
    )
    sheet_name = request.args.get("sheet")
    if not sheet_name:
        return jsonify({"error": "Missing required query parameter: sheet"}), 400

    tab_match = next(
        (tab for tab in data["main_sheet_tabs"] if tab["sheet_name"] == sheet_name),
        None,
    )
    if not tab_match:
        abort(404, description=f"Unknown main sheet: {sheet_name}")

    month_keys_csv = request.args.get("month_keys", "")
    mode = request.args.get("mode", "past_months")
    if mode not in {"past_months", "same_month_years"}:
        mode = "past_months"
    n_value = parse_clamped_int(request.args.get("n"), default=6, min_value=1, max_value=60)
    month_value = parse_clamped_int(request.args.get("month"), default=0, min_value=0, max_value=12)
    main_stat = selection["main_path"].stat()
    html_payload = build_main_sheet_html_cached(
        str(selection["main_path"]),
        main_stat.st_mtime_ns,
        sheet_name,
        month_keys_csv,
        mode,
        n_value,
        month_value,
    )
    return jsonify(
        {
            "version": data["version"],
            "canonical": tab_match["canonical"],
            "filterable": tab_match["filterable"],
            "sheet_name": html_payload["sheet_name"],
            "row_count": html_payload["row_count"],
            "col_count": html_payload["col_count"],
            "frozen_count": html_payload["frozen_count"],
            "frozen_rows": html_payload.get("frozen_rows", 0),
            "frozen_columns": html_payload.get("frozen_columns", html_payload["frozen_count"]),
            "selected_month_labels": html_payload["selected_month_labels"],
            "available_months": html_payload["available_months"],
            "hidden_rows_skipped": html_payload.get("hidden_rows_skipped", 0),
            "hidden_columns_skipped": html_payload.get("hidden_columns_skipped", 0),
            "html": html_payload["html"],
        }
    )


if __name__ == "__main__":
    host, port = resolve_runtime_host_port()
    print(f"Starting viewer at http://{host}:{port}")
    app.run(host=host, port=port, debug=False)
