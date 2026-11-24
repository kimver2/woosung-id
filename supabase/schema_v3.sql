-- ============================================================================
-- 우성아이디 통합 현장·재무 대시보드 v3 — Supabase PostgreSQL 보안 스키마 (Idempotent)
-- ============================================================================
-- 실행 방법:
--   Supabase Dashboard -> SQL Editor -> New query -> 전체 붙여넣기 -> Run
--   (여러 번 반복 실행해도 기존 데이터 손상 없이 안전하게 보안 정책 및 RPC가 재적용됩니다.)
-- ============================================================================

-- 1. 대시보드 전체/도메인별 스냅샷 고속 동기화 & 호환용 테이블
create table if not exists public.ws_state_store (
  key text primary key,
  payload jsonb not null default '{}'::jsonb,
  updated_at timestamptz not null default now(),
  updated_by text default ''
);

-- 2. 현장 정규화 테이블 (개별 현장 조회·확장용)
create table if not exists public.ws_sites (
  site_name text primary key,
  status text default '',
  meta jsonb default '{}'::jsonb,
  drive_folder_url text default '',
  updated_at timestamptz not null default now()
);

-- 3. 공정/일정 정규화 테이블 (개별 공정 조회·필터링·확장용)
create table if not exists public.ws_tasks (
  id text primary key,
  row_index integer,
  site_name text default '',
  category text default 'site',
  work_date text default '',
  end_date text default '',
  assignee text default '',
  content text default '',
  status text default '',
  memo text default '',
  raw_json jsonb default '{}'::jsonb,
  updated_at timestamptz not null default now()
);

-- 4. 재무·원장 정규화 테이블 (수금/비용마감/고정비 마스터)
create table if not exists public.ws_finance_store (
  record_key text primary key,
  category text default '',
  payload jsonb not null default '{}'::jsonb,
  updated_at timestamptz not null default now()
);

-- 5. 보안 PIN 해시 전용 테이블 (REST 외부 직접 노출 완전 차단)
create table if not exists public.ws_secrets (
  id integer primary key default 1 check (id = 1),
  shared_pin_hash text not null,
  admin_pin_hash text not null,
  updated_at timestamptz not null default now()
);

-- 조회 성능 최적화 인덱스
create index if not exists idx_ws_sites_status on public.ws_sites (status);
create index if not exists idx_ws_tasks_site_name on public.ws_tasks (site_name);
create index if not exists idx_ws_tasks_work_date on public.ws_tasks (work_date);
create index if not exists idx_ws_tasks_category on public.ws_tasks (category);
create index if not exists idx_ws_tasks_status on public.ws_tasks (status);
create index if not exists idx_ws_finance_category on public.ws_finance_store (category);

-- updated_at 자동 갱신 트리거 함수
create or replace function public.ws_v3_set_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

drop trigger if exists trg_ws_state_store_updated_at on public.ws_state_store;
create trigger trg_ws_state_store_updated_at
  before update on public.ws_state_store
  for each row execute function public.ws_v3_set_updated_at();

drop trigger if exists trg_ws_sites_updated_at on public.ws_sites;
create trigger trg_ws_sites_updated_at
  before update on public.ws_sites
  for each row execute function public.ws_v3_set_updated_at();

drop trigger if exists trg_ws_tasks_updated_at on public.ws_tasks;
create trigger trg_ws_tasks_updated_at
  before update on public.ws_tasks
  for each row execute function public.ws_v3_set_updated_at();

drop trigger if exists trg_ws_finance_store_updated_at on public.ws_finance_store;
create trigger trg_ws_finance_store_updated_at
  before update on public.ws_finance_store
  for each row execute function public.ws_v3_set_updated_at();

-- ============================================================================
-- SHA-256 PIN 해시 헬퍼 및 보안 시드 마이그레이션 (평문 auth_config -> ws_secrets)
-- ============================================================================
create or replace function public.ws_sha256_pin(p_pin text)
returns text
language sql
immutable
as $$
  select encode(sha256(convert_to('ws_v3_pin_v1::' || coalesce(btrim(p_pin), ''), 'UTF8')), 'hex');
$$;

do $$
declare
  v_auth jsonb;
  v_shared text;
  v_admin text;
begin
  select payload into v_auth
  from public.ws_state_store
  where key = 'auth_config';

  v_shared := nullif(btrim(coalesce(v_auth->>'sharedPin', '')), '');
  v_admin := nullif(btrim(coalesce(v_auth->>'adminPin', '')), '');

  if not exists (select 1 from public.ws_secrets where id = 1) then
    insert into public.ws_secrets (id, shared_pin_hash, admin_pin_hash)
    values (
      1,
      public.ws_sha256_pin(coalesce(v_shared, '0112')),
      public.ws_sha256_pin(coalesce(v_admin, '007077'))
    );
  else
    if v_shared is not null then
      update public.ws_secrets set shared_pin_hash = public.ws_sha256_pin(v_shared), updated_at = now() where id = 1;
    end if;
    if v_admin is not null then
      update public.ws_secrets set admin_pin_hash = public.ws_sha256_pin(v_admin), updated_at = now() where id = 1;
    end if;
  end if;

  -- 평문 PIN이 담겨 있던 auth_config 행 즉시 제거
  delete from public.ws_state_store where key = 'auth_config';

  -- Realtime 경량 알림용 sync_signal 초기 행 보장 (개인정보 0%)
  insert into public.ws_state_store (key, payload, updated_by)
  values ('sync_signal', '{"init":true}'::jsonb, 'system')
  on conflict (key) do nothing;
end $$;

-- ============================================================================
-- SECURITY DEFINER 보안 RPC 함수군 (외부 공유 격리 + 내부 PIN 해시 검증)
-- ============================================================================

-- 1) 고객/협력업체 공유 링크 단건 전용 조회 (타 현장·재무·내부 데이터 접근 완전 차단)
create or replace function public.ws_get_public_share(p_share_key text)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  v_row record;
begin
  if p_share_key is null or (p_share_key not like 'client_share::%' and p_share_key not like 'partner_share::%') then
    return null;
  end if;
  if length(p_share_key) < 20 or length(p_share_key) > 200 then
    return null;
  end if;

  select payload, updated_at into v_row
  from public.ws_state_store
  where key = p_share_key;

  if not found then
    return null;
  end if;

  return jsonb_build_object(
    'payload', v_row.payload,
    'updated_at', v_row.updated_at
  );
end;
$$;

-- 2) 내부 대시보드 PIN 해시 검증 (평문 PIN 네트워크 미전송)
create or replace function public.ws_rpc_verify_pin(p_user text, p_pin_hash text)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  v_sec record;
  v_clean_user text := regexp_replace(coalesce(p_user, ''), '\s+', '', 'g');
  v_lower_user text := lower(regexp_replace(coalesce(p_user, ''), '\s+', '', 'g'));
  v_hash text := lower(btrim(coalesce(p_pin_hash, '')));
  v_members text[] := array['차정률', '조항만', '송지은', '전예린', '한송이'];
  v_set_payload jsonb;
  v_dash_payload jsonb;
  v_matched text := null;
  v_m text;
begin
  select shared_pin_hash, admin_pin_hash into v_sec from public.ws_secrets where id = 1;
  if not found or v_hash = '' then
    return jsonb_build_object('success', false, 'error', '인증 설정을 확인할 수 없습니다.');
  end if;

  if v_lower_user = 'admin' then
    if v_hash = v_sec.admin_pin_hash then
      return jsonb_build_object(
        'success', true,
        'user', 'admin',
        'isAdmin', true,
        'usedAdminPin', true,
        'adminPinConfigured', true
      );
    end if;
    return jsonb_build_object('success', false, 'error', '비밀번호(PIN)가 일치하지 않습니다.');
  end if;

  select payload into v_set_payload from public.ws_state_store where key = 'settings_main';
  if v_set_payload is not null and jsonb_typeof(v_set_payload->'members') = 'array' then
    select array_agg(distinct coalesce(elem->>'name', trim(both '"' from elem::text)))
    into v_members
    from jsonb_array_elements(v_set_payload->'members') as elem
    where coalesce(elem->>'name', trim(both '"' from elem::text)) <> '';
  else
    select payload into v_dash_payload from public.ws_state_store where key = 'dashboard_main';
    if v_dash_payload is not null and jsonb_typeof(v_dash_payload->'members') = 'array' then
      select array_agg(distinct trim(both '"' from elem::text))
      into v_members
      from jsonb_array_elements(v_dash_payload->'members') as elem
      where trim(both '"' from elem::text) <> '';
    end if;
  end if;

  if v_members is not null then
    foreach v_m in array v_members loop
      if lower(regexp_replace(coalesce(v_m, ''), '\s+', '', 'g')) = v_lower_user then
        v_matched := v_m;
        exit;
      end if;
    end loop;
  end if;

  if v_matched is null then
    return jsonb_build_object('success', false, 'error', '등록되지 않은 사용자입니다. (구성원 설정 확인)');
  end if;

  if v_lower_user = '조항만' then
    if v_hash = v_sec.admin_pin_hash or v_hash = v_sec.shared_pin_hash then
      return jsonb_build_object(
        'success', true,
        'user', v_matched,
        'isAdmin', false,
        'usedAdminPin', (v_hash = v_sec.admin_pin_hash),
        'adminPinConfigured', true
      );
    end if;
    return jsonb_build_object('success', false, 'error', '비밀번호(PIN)가 일치하지 않습니다.');
  end if;

  if v_hash = v_sec.shared_pin_hash then
    return jsonb_build_object(
      'success', true,
      'user', v_matched,
      'isAdmin', false,
      'usedAdminPin', false,
      'adminPinConfigured', true
    );
  end if;

  return jsonb_build_object('success', false, 'error', '비밀번호(PIN)가 일치하지 않습니다.');
end;
$$;

-- 3) 인증된 내부 세션 전용 상태 조회 RPC
create or replace function public.ws_rpc_select_states(p_pin_hash text, p_keys text[] default null)
returns table(key text, payload jsonb, updated_at timestamptz, updated_by text)
language plpgsql
security definer
set search_path = public
as $$
declare
  v_sec record;
  v_hash text := lower(btrim(coalesce(p_pin_hash, '')));
begin
  select shared_pin_hash, admin_pin_hash into v_sec from public.ws_secrets where id = 1;
  if not found or (v_hash <> v_sec.shared_pin_hash and v_hash <> v_sec.admin_pin_hash) then
    raise exception 'UNAUTHORIZED_PIN_HASH';
  end if;

  if p_keys is not null and array_length(p_keys, 1) > 0 then
    return query
      select s.key, s.payload, s.updated_at, s.updated_by
      from public.ws_state_store s
      where s.key = any(p_keys)
        and s.key <> 'auth_config';
  else
    return query
      select s.key, s.payload, s.updated_at, s.updated_by
      from public.ws_state_store s
      where s.key <> 'auth_config';
  end if;
end;
$$;

-- 4) 인증된 내부 세션 전용 상태 저장 + 경량 sync_signal 브로드캐스트 RPC
create or replace function public.ws_rpc_upsert_states(p_pin_hash text, p_rows jsonb)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  v_sec record;
  v_hash text := lower(btrim(coalesce(p_pin_hash, '')));
  v_elem jsonb;
  v_key text;
  v_last_mut text := '';
  v_keys text[] := array[]::text[];
  v_count integer := 0;
begin
  select shared_pin_hash, admin_pin_hash into v_sec from public.ws_secrets where id = 1;
  if not found or (v_hash <> v_sec.shared_pin_hash and v_hash <> v_sec.admin_pin_hash) then
    raise exception 'UNAUTHORIZED_PIN_HASH';
  end if;

  if p_rows is null or jsonb_typeof(p_rows) <> 'array' then
    return jsonb_build_object('ok', true, 'count', 0);
  end if;

  for v_elem in select * from jsonb_array_elements(p_rows)
  loop
    v_key := coalesce(v_elem->>'key', '');
    if v_key <> '' and v_key <> 'auth_config' then
      insert into public.ws_state_store (key, payload, updated_at, updated_by)
      values (
        v_key,
        coalesce(v_elem->'payload', '{}'::jsonb),
        coalesce((v_elem->>'updated_at')::timestamptz, now()),
        coalesce(v_elem->>'updated_by', '')
      )
      on conflict (key) do update set
        payload = excluded.payload,
        updated_at = excluded.updated_at,
        updated_by = excluded.updated_by;

      v_count := v_count + 1;
      if v_key <> 'sync_signal' then
        v_keys := array_append(v_keys, v_key);
        if coalesce(v_elem->'payload'->>'_clientMutId', '') <> '' then
          v_last_mut := v_elem->'payload'->>'_clientMutId';
        end if;
      end if;
    end if;
  end loop;

  if array_length(v_keys, 1) > 0 then
    insert into public.ws_state_store (key, payload, updated_at, updated_by)
    values (
      'sync_signal',
      jsonb_build_object(
        'changedKeys', to_jsonb(v_keys),
        '_clientMutId', v_last_mut,
        '_updatedMs', (extract(epoch from clock_timestamp()) * 1000)::bigint
      ),
      'rpc_signal'
    )
    on conflict (key) do update set
      payload = excluded.payload,
      updated_at = excluded.updated_at,
      updated_by = excluded.updated_by;
  end if;

  return jsonb_build_object('ok', true, 'count', v_count);
end;
$$;

-- 5) 정규화 보조 테이블(ws_sites, ws_tasks, ws_finance_store) 인증 동기화 RPC
create or replace function public.ws_rpc_sync_table(
  p_pin_hash text,
  p_table text,
  p_op text,
  p_rows jsonb default '[]'::jsonb,
  p_ids text[] default null
)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  v_sec record;
  v_hash text := lower(btrim(coalesce(p_pin_hash, '')));
  v_res jsonb := '[]'::jsonb;
begin
  select shared_pin_hash, admin_pin_hash into v_sec from public.ws_secrets where id = 1;
  if not found or (v_hash <> v_sec.shared_pin_hash and v_hash <> v_sec.admin_pin_hash) then
    raise exception 'UNAUTHORIZED_PIN_HASH';
  end if;

  if p_table = 'ws_sites' then
    if p_op = 'select' then
      select coalesce(jsonb_agg(row_to_json(t)), '[]'::jsonb) into v_res from public.ws_sites t;
      return v_res;
    elsif p_op = 'upsert' then
      insert into public.ws_sites (site_name, status, meta, drive_folder_url, updated_at)
      select
        coalesce(x->>'site_name', ''),
        coalesce(x->>'status', ''),
        coalesce(x->'meta', '{}'::jsonb),
        coalesce(x->>'drive_folder_url', ''),
        coalesce((x->>'updated_at')::timestamptz, now())
      from jsonb_array_elements(p_rows) x
      where coalesce(x->>'site_name', '') <> ''
      on conflict (site_name) do update set
        status = excluded.status,
        meta = excluded.meta,
        drive_folder_url = excluded.drive_folder_url,
        updated_at = excluded.updated_at;
      return jsonb_build_object('ok', true);
    elsif p_op = 'delete' and p_ids is not null then
      delete from public.ws_sites where site_name = any(p_ids);
      return jsonb_build_object('ok', true);
    end if;
  elsif p_table = 'ws_tasks' then
    if p_op = 'select' then
      select coalesce(jsonb_agg(row_to_json(t)), '[]'::jsonb) into v_res from public.ws_tasks t;
      return v_res;
    elsif p_op = 'upsert' then
      insert into public.ws_tasks (id, row_index, site_name, category, work_date, end_date, assignee, content, status, memo, raw_json, updated_at)
      select
        coalesce(x->>'id', ''),
        (x->>'row_index')::integer,
        coalesce(x->>'site_name', ''),
        coalesce(x->>'category', 'site'),
        coalesce(x->>'work_date', ''),
        coalesce(x->>'end_date', ''),
        coalesce(x->>'assignee', ''),
        coalesce(x->>'content', ''),
        coalesce(x->>'status', ''),
        coalesce(x->>'memo', ''),
        coalesce(x->'raw_json', '{}'::jsonb),
        coalesce((x->>'updated_at')::timestamptz, now())
      from jsonb_array_elements(p_rows) x
      where coalesce(x->>'id', '') <> ''
      on conflict (id) do update set
        row_index = excluded.row_index,
        site_name = excluded.site_name,
        category = excluded.category,
        work_date = excluded.work_date,
        end_date = excluded.end_date,
        assignee = excluded.assignee,
        content = excluded.content,
        status = excluded.status,
        memo = excluded.memo,
        raw_json = excluded.raw_json,
        updated_at = excluded.updated_at;
      return jsonb_build_object('ok', true);
    elsif p_op = 'delete' and p_ids is not null then
      delete from public.ws_tasks where id = any(p_ids);
      return jsonb_build_object('ok', true);
    end if;
  elsif p_table = 'ws_finance_store' then
    if p_op = 'select' then
      select coalesce(jsonb_agg(row_to_json(t)), '[]'::jsonb) into v_res from public.ws_finance_store t;
      return v_res;
    elsif p_op = 'upsert' then
      insert into public.ws_finance_store (record_key, category, payload, updated_at)
      select
        coalesce(x->>'record_key', ''),
        coalesce(x->>'category', ''),
        coalesce(x->'payload', '{}'::jsonb),
        coalesce((x->>'updated_at')::timestamptz, now())
      from jsonb_array_elements(p_rows) x
      where coalesce(x->>'record_key', '') <> ''
      on conflict (record_key) do update set
        category = excluded.category,
        payload = excluded.payload,
        updated_at = excluded.updated_at;
      return jsonb_build_object('ok', true);
    elsif p_op = 'delete' and p_ids is not null then
      delete from public.ws_finance_store where record_key = any(p_ids);
      return jsonb_build_object('ok', true);
    end if;
  end if;

  return jsonb_build_object('ok', false, 'error', 'UNSUPPORTED_OPERATION');
end;
$$;

-- 6) 공용/관리자 PIN 변경 RPC
create or replace function public.ws_rpc_update_pin(
  p_current_pin_hash text,
  p_field text,
  p_new_pin_hash text
)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  v_sec record;
  v_cur text := lower(btrim(coalesce(p_current_pin_hash, '')));
  v_new text := lower(btrim(coalesce(p_new_pin_hash, '')));
begin
  select shared_pin_hash, admin_pin_hash into v_sec from public.ws_secrets where id = 1;
  if not found or v_new = '' then
    return jsonb_build_object('success', false, 'error', 'INVALID_PIN_HASH');
  end if;

  if p_field = 'adminPin' then
    if v_cur <> v_sec.admin_pin_hash and v_cur <> v_sec.shared_pin_hash then
      return jsonb_build_object('success', false, 'error', 'UNAUTHORIZED');
    end if;
    update public.ws_secrets set admin_pin_hash = v_new, updated_at = now() where id = 1;
  else
    if v_cur <> v_sec.shared_pin_hash and v_cur <> v_sec.admin_pin_hash then
      return jsonb_build_object('success', false, 'error', 'UNAUTHORIZED');
    end if;
    update public.ws_secrets set shared_pin_hash = v_new, updated_at = now() where id = 1;
  end if;

  return jsonb_build_object('success', true, 'adminPinConfigured', true);
end;
$$;

-- ============================================================================
-- 권한 및 RLS(Row Level Security) 최소 권한 잠금
-- ============================================================================
grant usage on schema public to anon, authenticated;

-- 기존 전체 개방 권한 회수 후 sync_signal 조회만 허용
revoke all on table public.ws_secrets from anon, authenticated;
revoke all on table public.ws_sites from anon, authenticated;
revoke all on table public.ws_tasks from anon, authenticated;
revoke all on table public.ws_finance_store from anon, authenticated;
revoke all on table public.ws_state_store from anon, authenticated;
grant select on table public.ws_state_store to anon, authenticated;

grant execute on function public.ws_get_public_share(text) to anon, authenticated;
grant execute on function public.ws_rpc_verify_pin(text, text) to anon, authenticated;
grant execute on function public.ws_rpc_select_states(text, text[]) to anon, authenticated;
grant execute on function public.ws_rpc_upsert_states(text, jsonb) to anon, authenticated;
grant execute on function public.ws_rpc_sync_table(text, text, text, jsonb, text[]) to anon, authenticated;
grant execute on function public.ws_rpc_update_pin(text, text, text) to anon, authenticated;

alter table public.ws_state_store enable row level security;
alter table public.ws_sites enable row level security;
alter table public.ws_tasks enable row level security;
alter table public.ws_finance_store enable row level security;
alter table public.ws_secrets enable row level security;

drop policy if exists "ws_state_store_anon_auth_all" on public.ws_state_store;
drop policy if exists "ws_sites_anon_auth_all" on public.ws_sites;
drop policy if exists "ws_tasks_anon_auth_all" on public.ws_tasks;
drop policy if exists "ws_finance_store_anon_auth_all" on public.ws_finance_store;
drop policy if exists "ws_state_store_anon_sync_signal_select" on public.ws_state_store;

-- 익명(anon) 직접 테이블 접근은 오직 경량 알림 키(sync_signal) SELECT만 허용 (개인정보·재무 0%)
create policy "ws_state_store_anon_sync_signal_select"
  on public.ws_state_store
  for select
  to anon, authenticated
  using (key = 'sync_signal');

-- WAL 용량 폭증 방지를 위해 REPLICA IDENTITY DEFAULT 적용
alter table public.ws_state_store replica identity default;
alter table public.ws_sites replica identity default;
alter table public.ws_tasks replica identity default;
alter table public.ws_finance_store replica identity default;

-- Supabase Realtime Publication 등록 (ws_state_store의 sync_signal 변경 브로드캐스트)
do $$
declare
  already_added boolean;
begin
  if exists (select 1 from pg_publication where pubname = 'supabase_realtime') then
    select exists (
      select 1
      from pg_publication_tables
      where pubname = 'supabase_realtime'
        and schemaname = 'public'
        and tablename = 'ws_state_store'
    ) into already_added;

    if not already_added then
      begin
        alter publication supabase_realtime add table public.ws_state_store;
      exception
        when duplicate_object then null;
        when undefined_object then null;
      end;
    end if;
  end if;
end $$;
