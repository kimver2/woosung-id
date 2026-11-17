-- ============================================================================
-- 우성아이디 통합 현장·재무 대시보드 v3 — Supabase PostgreSQL 스키마 (Idempotent)
-- ============================================================================
-- 실행 방법:
--   Supabase Dashboard -> SQL Editor -> New query -> 전체 붙여넣기 -> Run
--   (여러 번 반복 실행해도 기존 데이터 손상 없이 안전하게 재적용됩니다.)
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

-- API 역할(anon / authenticated) 권한 부여
grant usage on schema public to anon, authenticated;
grant select, insert, update, delete on table public.ws_state_store to anon, authenticated;
grant select, insert, update, delete on table public.ws_sites to anon, authenticated;
grant select, insert, update, delete on table public.ws_tasks to anon, authenticated;
grant select, insert, update, delete on table public.ws_finance_store to anon, authenticated;

-- RLS(Row Level Security) 활성화 및 허용 정책 (Idempotent)
alter table public.ws_state_store enable row level security;
alter table public.ws_sites enable row level security;
alter table public.ws_tasks enable row level security;
alter table public.ws_finance_store enable row level security;

drop policy if exists "ws_state_store_anon_auth_all" on public.ws_state_store;
create policy "ws_state_store_anon_auth_all"
  on public.ws_state_store
  for all
  to anon, authenticated
  using (true)
  with check (true);

drop policy if exists "ws_sites_anon_auth_all" on public.ws_sites;
create policy "ws_sites_anon_auth_all"
  on public.ws_sites
  for all
  to anon, authenticated
  using (true)
  with check (true);

drop policy if exists "ws_tasks_anon_auth_all" on public.ws_tasks;
create policy "ws_tasks_anon_auth_all"
  on public.ws_tasks
  for all
  to anon, authenticated
  using (true)
  with check (true);

drop policy if exists "ws_finance_store_anon_auth_all" on public.ws_finance_store;
create policy "ws_finance_store_anon_auth_all"
  on public.ws_finance_store
  for all
  to anon, authenticated
  using (true)
  with check (true);

-- Realtime 변경사항 전체 필드 전송 설정
alter table public.ws_state_store replica identity full;
alter table public.ws_sites replica identity full;
alter table public.ws_tasks replica identity full;
alter table public.ws_finance_store replica identity full;

-- Supabase Realtime Publication 등록 (이미 등록된 경우 안전하게 무시)
do $$
declare
  tbl text;
  already_added boolean;
begin
  if exists (select 1 from pg_publication where pubname = 'supabase_realtime') then
    foreach tbl in array array['ws_state_store', 'ws_sites', 'ws_tasks', 'ws_finance_store']
    loop
      select exists (
        select 1
        from pg_publication_tables
        where pubname = 'supabase_realtime'
          and schemaname = 'public'
          and tablename = tbl
      ) into already_added;

      if not already_added then
        begin
          execute format('alter publication supabase_realtime add table public.%I', tbl);
        exception
          when duplicate_object then null;
          when undefined_object then null;
        end;
      end if;
    end loop;
  end if;
end $$;
