#!/usr/bin/env python3
"""
GAS 웹앱 스냅샷(gas_snapshot.json) -> Supabase PostgreSQL 일괄 이관 스크립트
"""
import json
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request

SUPABASE_URL = os.environ.get("SUPABASE_URL", "https://uyqvguzglviqegxmxqyn.supabase.co").rstrip("/")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY", "sb_publishable_f8JwxMgpgs6AVSEumSIV2g_KlK6o9GA")
DEFAULT_SNAPSHOT = "/Users/kimver2/.gemini/antigravity-cli/brain/e4f7a26c-f9d6-41a5-998c-01719b56801d/scratch/gas_snapshot.json"


def to_int(val, default=0):
    if val is None or val == "":
        return default
    if isinstance(val, (int, float)):
        return int(val)
    cleaned = re.sub(r"[^0-9\-]", "", str(val))
    if not cleaned or cleaned == "-":
        return default
    try:
        return int(cleaned)
    except ValueError:
        return default


def to_float_or_none(val):
    if val is None or val == "":
        return None
    try:
        return float(val)
    except (ValueError, TypeError):
        return None


def supabase_request(method, path, payload=None, prefer="return=minimal"):
    url = f"{SUPABASE_URL}/rest/v1/{path}"
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8") if payload is not None else None
    headers = {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": "application/json",
        "Prefer": prefer,
    }
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            body = resp.read().decode("utf-8")
            return resp.status, body
    except urllib.error.HTTPError as e:
        err_body = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTP {e.code} on {method} {path}: {err_body}")


def batch_upsert(table, rows, on_conflict=None, batch_size=200):
    if not rows:
        return 0
    total = 0
    for i in range(0, len(rows), batch_size):
        chunk = rows[i : i + batch_size]
        path = table
        prefer = "return=minimal"
        if on_conflict:
            path = f"{table}?on_conflict={urllib.parse.quote(on_conflict)}"
            prefer = "resolution=merge-duplicates,return=minimal"
        supabase_request("POST", path, chunk, prefer=prefer)
        total += len(chunk)
    return total


def main():
    snap_path = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_SNAPSHOT
    with open(snap_path, "r", encoding="utf-8") as f:
        snap = json.load(f)

    dashboard = snap["dashboard"]
    settings = snap["settings"]
    finance = snap["finance"]

    # 1. members
    members_rows = []
    seen_members = set()
    for idx, m in enumerate(settings.get("members", [])):
        name = str(m.get("name", "")).strip()
        if not name or name in seen_members:
            continue
        seen_members.add(name)
        members_rows.append({
            "name": name,
            "role": str(m.get("role", "")).strip(),
            "phone": str(m.get("phone", "")).strip(),
            "email": str(m.get("email", "")).strip(),
            "receive_briefing": bool(m.get("receive", True)),
            "note": str(m.get("note", "")).strip(),
            "sort_order": idx,
        })
    print(f"[1/9] members: {batch_upsert('members', members_rows, 'name')}건 이관")

    # 2. task_categories
    task_cat_rows = []
    seen_tasks = set()
    for idx, t in enumerate(settings.get("tasks", [])):
        t_name = str(t.get("taskName", "")).strip()
        if not t_name or t_name in seen_tasks:
            continue
        seen_tasks.add(t_name)
        task_cat_rows.append({
            "task_name": t_name,
            "category": str(t.get("category", "")).strip(),
            "note": str(t.get("note", "")).strip(),
            "sort_order": idx,
        })
    print(f"[2/9] task_categories: {batch_upsert('task_categories', task_cat_rows, 'task_name')}건 이관")

    # 3. partners
    partner_rows = []
    seen_partners = set()
    for idx, p in enumerate(settings.get("partners", [])):
        name = str(p.get("name", "")).strip()
        biz = str(p.get("biz", "")).strip()
        if not name:
            continue
        key = (biz, name)
        if key in seen_partners:
            continue
        seen_partners.add(key)
        partner_rows.append({
            "name": name,
            "biz": biz,
            "category": str(p.get("category", "")).strip(),
            "priority": str(p.get("priority", "")).strip(),
            "manager": str(p.get("manager", "")).strip(),
            "phone": str(p.get("phone", "")).strip(),
            "status": str(p.get("status", "활성")).strip() or "활성",
            "feature": str(p.get("feature", "")).strip(),
            "order_method": str(p.get("orderMethod", "")).strip(),
            "note": str(p.get("note", "")).strip(),
            "dropdown_val": str(p.get("dropdownVal", "")).strip(),
            "bank_account": str(p.get("bankAccount", "")).strip(),
            "sort_order": idx,
        })
    print(f"[3/9] partners: {batch_upsert('partners', partner_rows, 'biz,name')}건 이관")

    # 4. sites & 5. site_tasks & 6. estimates
    site_info_map = dashboard.get("siteInfoMap", {}) or {}
    norm_info_map = {}
    for k, v in site_info_map.items():
        norm_info_map[k] = v
        norm_info_map[re.sub(r"\s+", "", k)] = v

    sites_rows = []
    seen_sites = set()
    site_task_rows = []
    estimate_rows = []

    for idx, s in enumerate(dashboard.get("sites", [])):
        s_name = str(s.get("siteName", "")).strip()
        raw_scope = str(s.get("scope", "site")).strip()
        if raw_scope == "newSite":
            ed = s.get("estimateData") or {}
            est_name = str(ed.get("siteName") or s_name).strip()
            if est_name:
                estimate_rows.append({
                    "site_customer": est_name,
                    "stage": str(ed.get("status") or "신규 문의").strip(),
                    "inquiry_date": str(ed.get("inquiryDate") or "").strip(),
                    "priority": str(ed.get("priority") or "보통").strip(),
                    "work_scope": str(ed.get("scope") or "").strip(),
                    "consult_memo": str(ed.get("memo") or "").strip(),
                    "source": str(ed.get("source") or "").strip(),
                    "first_consult": str(ed.get("consultDate") or "").strip(),
                    "expected_start_date": str(ed.get("expectedDate") or "").strip(),
                    "expected_amount": to_int(ed.get("expectedAmount"), 0),
                    "quote_sent_date": str(ed.get("estimateSendDate") or "").strip(),
                    "client_name": str(ed.get("clientName") or "").strip(),
                    "phone": str(ed.get("clientPhone") or "").strip(),
                    "manager": str(ed.get("manager") or "").strip(),
                    "quote_file": str(ed.get("estimateFile") or "").strip(),
                    "contract_amount": to_int(ed.get("contractAmount"), 0),
                    "next_action": str(ed.get("nextAction") or "").strip(),
                    "hold_reason": str(ed.get("failReason") or "").strip(),
                    "visit_date": str(ed.get("visitDate") or "").strip(),
                    "site_address": str(ed.get("address") or "").strip(),
                })
            continue

        if not s_name:
            continue
        scope = "completed" if s.get("isArchived") else (raw_scope if raw_scope in ("site", "single", "as") else "site")

        if s_name not in seen_sites:
            seen_sites.add(s_name)
            info = norm_info_map.get(s_name) or norm_info_map.get(re.sub(r"\s+", "", s_name)) or {}
            sites_rows.append({
                "site_name": s_name,
                "scope": scope,
                "address": str(info.get("address", "")).strip(),
                "door_password": str(info.get("doorPassword", "")).strip(),
                "office_phone": str(info.get("officePhone", "")).strip(),
                "consent_done": bool(info.get("consentDone", False)),
                "elevator_cover_done": bool(info.get("elevatorCoverDone", False)),
                "office_notice_done": bool(info.get("officeNoticeDone", False)),
                "memo": str(info.get("memo", "")).strip(),
                "drive_url": str(info.get("driveUrl", "")).strip(),
                "color_index": to_int(s.get("colorCode"), idx % 12),
            })

        for t in s.get("tasks", []):
            mgr_raw = str(t.get("manager", "")).strip()
            mgr_parts = [x.strip() for x in mgr_raw.split(",") if x.strip()]
            m1 = mgr_parts[0] if len(mgr_parts) >= 1 else ""
            m2 = mgr_parts[1] if len(mgr_parts) >= 2 else ""
            t_scope = "completed" if t.get("isArchived") else (str(t.get("scope") or scope).strip())
            if t_scope not in ("site", "single", "as", "completed"):
                t_scope = scope
            site_task_rows.append({
                "site_name": s_name,
                "scope": t_scope,
                "task_date": str(t.get("dateDisplay") or t.get("startDateStr") or "").strip(),
                "task_name": str(t.get("taskName", "")).strip(),
                "partner": str(t.get("partner", "")).strip(),
                "is_shared": bool(t.get("isShared", False)),
                "note": str(t.get("memo", "")).strip(),
                "photos": [],
                "manager_1": m1,
                "manager_2": m2,
                "status": str(t.get("status") or "예정").strip(),
                "calendar_id": str(t.get("id") or t.get("taskId") or "").strip(),
                "legacy_updated_at": to_int(t.get("dateMs"), 0) or None,
            })

    # siteInfoMap에만 존재하는 현장도 보충
    for raw_k, info in site_info_map.items():
        s_name = str(info.get("siteName") or raw_k).strip()
        if not s_name or s_name in seen_sites:
            continue
        seen_sites.add(s_name)
        sites_rows.append({
            "site_name": s_name,
            "scope": "site",
            "address": str(info.get("address", "")).strip(),
            "door_password": str(info.get("doorPassword", "")).strip(),
            "office_phone": str(info.get("officePhone", "")).strip(),
            "consent_done": bool(info.get("consentDone", False)),
            "elevator_cover_done": bool(info.get("elevatorCoverDone", False)),
            "office_notice_done": bool(info.get("officeNoticeDone", False)),
            "memo": str(info.get("memo", "")).strip(),
            "drive_url": str(info.get("driveUrl", "")).strip(),
            "color_index": 0,
        })

    print(f"[4/9] sites: {batch_upsert('sites', sites_rows, 'site_name')}건 이관")
    # site_tasks, estimates는 중복 적재 방지를 위해 기존 데이터 비우고 재적재
    supabase_request("DELETE", "site_tasks?id=not.is.null")
    print(f"[5/9] site_tasks: {batch_upsert('site_tasks', site_task_rows)}건 이관")
    supabase_request("DELETE", "estimates?id=not.is.null")
    print(f"[6/9] estimates: {batch_upsert('estimates', estimate_rows)}건 이관")

    # 7. finance_collections
    col_rows = []
    seen_col = set()
    for c in finance.get("collection", []):
        lid = str(c.get("id", "")).strip()
        if not lid or lid in seen_col:
            continue
        seen_col.add(lid)
        col_rows.append({
            "legacy_id": lid,
            "site_name": str(c.get("siteName", "")).strip(),
            "period": str(c.get("period", "")).strip(),
            "manager": str(c.get("manager", "")).strip(),
            "category": str(c.get("category", "")).strip(),
            "status": str(c.get("status", "진행중")).strip() or "진행중",
            "client_name": str(c.get("clientName", "")).strip(),
            "contract_amount": to_int(c.get("contractAmount"), 0),
            "change_amount": to_int(c.get("changeAmount"), 0),
            "change_desc": str(c.get("changeDesc", "")).strip(),
            "progress_rate": to_float_or_none(c.get("progressRate")),
            "total_amount": to_int(c.get("totalAmount"), 0),
            "down_payment": to_int(c.get("downPayment"), 0),
            "start_payment": to_int(c.get("startPayment"), 0),
            "mid_payment_1": to_int(c.get("midPayment1"), 0),
            "mid_payment_2": to_int(c.get("midPayment2"), 0),
            "mid_payment_3": to_int(c.get("midPayment3"), 0),
            "final_payment": to_int(c.get("finalPayment"), 0),
            "uncollected_amount": to_int(c.get("uncollectedAmount"), 0),
            "tax_invoice": str(c.get("taxInvoice", "미발행")).strip() or "미발행",
            "note": str(c.get("note", "")).strip(),
        })
    print(f"[7/9] finance_collections: {batch_upsert('finance_collections', col_rows, 'legacy_id')}건 이관")

    # 8. finance_expenses
    exp_rows = []
    seen_exp = set()
    for e in finance.get("expense", []):
        lid = str(e.get("id", "")).strip()
        if not lid or lid in seen_exp:
            continue
        seen_exp.add(lid)
        exp_rows.append({
            "legacy_id": lid,
            "settle_month": str(e.get("settleMonth", "")).strip(),
            "item": str(e.get("item", "")).strip(),
            "category": str(e.get("category", "")).strip(),
            "vendor": str(e.get("vendor", "")).strip(),
            "account_number": str(e.get("accountNumber") or e.get("category2") or "").strip(),
            "site_name": str(e.get("siteName", "")).strip(),
            "claim_amount": to_int(e.get("claimAmount"), 0),
            "paid_amount": to_int(e.get("paidAmount"), 0),
            "balance": to_int(e.get("balance"), 0),
            "pay_date": str(e.get("payDate", "")).strip(),
            "tax_proof": "",
            "note": str(e.get("note", "")).strip(),
        })
    print(f"[8/9] finance_expenses: {batch_upsert('finance_expenses', exp_rows, 'legacy_id')}건 이관")

    # 9. finance_fixed_master & app_settings
    fixed_rows = []
    seen_fixed = set()
    for f_item in finance.get("fixedMaster", []):
        lid = str(f_item.get("id", "")).strip()
        if not lid or lid in seen_fixed:
            continue
        seen_fixed.add(lid)
        fixed_rows.append({
            "legacy_id": lid,
            "is_active": str(f_item.get("active", "Y")).strip().upper() == "Y",
            "item": str(f_item.get("item", "고정비")).strip() or "고정비",
            "category": str(f_item.get("category", "")).strip(),
            "vendor": str(f_item.get("vendor", "")).strip(),
            "account_number": str(f_item.get("accountNumber") or f_item.get("category2") or "").strip(),
            "base_amount": to_int(f_item.get("baseAmount"), 0),
            "pay_day": str(f_item.get("payDay", "")).strip(),
            "note": str(f_item.get("note", "")).strip(),
        })
    print(f"[9/9] finance_fixed_master: {batch_upsert('finance_fixed_master', fixed_rows, 'legacy_id')}건 이관")

    app_settings_rows = [
        {
            "key": "auth_config",
            "value": {"default_pin": "0112", "admin_user": "admin", "admin_pin": "007077"},
        }
    ]
    batch_upsert("app_settings", app_settings_rows, "key")
    print("[보너스] app_settings: 기본 인증 설정(auth_config) 반영 완료")


if __name__ == "__main__":
    main()
