-- ============================================================================
-- 우성아이디 통합 현장·재무 대시보드 v3 — Supabase PostgreSQL 초기 스키마
-- ============================================================================
-- 기존 Google Sheets(02_거래처, 03_구성원, 04_공종, 현장정보, 현장데이터/단건/AS/완료현장,
-- 신규견적, 05_ToDo) 및 우성아이디_재무DB(수금현황, 비용마감, 비용마감_고정) 1:1 매핑
--
-- 실행 방법:
--   Supabase Dashboard -> SQL Editor -> New query -> 전체 복사/붙여넣기 후 Run
-- ============================================================================

-- 0. 확장 모듈 및 공통 updated_at 자동 갱신 트리거 함수
CREATE EXTENSION IF NOT EXISTS pgcrypto;

CREATE OR REPLACE FUNCTION public.set_updated_at()
RETURNS TRIGGER
LANGUAGE plpgsql
AS $$
BEGIN
  NEW.updated_at = now();
  RETURN NEW;
END;
$$;

-- ============================================================================
-- 1. 마스터 기준정보 테이블
-- ============================================================================

-- 1-1. 사내 구성원 (03_구성원)
CREATE TABLE IF NOT EXISTS public.members (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  name text NOT NULL UNIQUE,
  role text NOT NULL DEFAULT '',
  phone text NOT NULL DEFAULT '',
  email text NOT NULL DEFAULT '',
  receive_briefing boolean NOT NULL DEFAULT true,
  note text NOT NULL DEFAULT '',
  sort_order integer NOT NULL DEFAULT 0,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now()
);

-- 1-2. 표준 공종 마스터 (04_공종)
CREATE TABLE IF NOT EXISTS public.task_categories (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  task_name text NOT NULL UNIQUE,
  category text NOT NULL DEFAULT '',
  note text NOT NULL DEFAULT '',
  sort_order integer NOT NULL DEFAULT 0,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now()
);

-- 1-3. 협력업체 마스터 (02_거래처)
CREATE TABLE IF NOT EXISTS public.partners (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  name text NOT NULL,
  biz text NOT NULL DEFAULT '',
  category text NOT NULL DEFAULT '',
  priority text NOT NULL DEFAULT '',
  manager text NOT NULL DEFAULT '',
  phone text NOT NULL DEFAULT '',
  status text NOT NULL DEFAULT '활성',
  feature text NOT NULL DEFAULT '',
  order_method text NOT NULL DEFAULT '',
  note text NOT NULL DEFAULT '',
  dropdown_val text NOT NULL DEFAULT '',
  bank_account text NOT NULL DEFAULT '',
  sort_order integer NOT NULL DEFAULT 0,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now(),
  CONSTRAINT partners_biz_name_unique UNIQUE (biz, name)
);

-- ============================================================================
-- 2. 현장 정보 & 공정 일정 테이블
-- ============================================================================

-- 2-1. 현장 기본정보 (현장정보 시트)
CREATE TABLE IF NOT EXISTS public.sites (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  site_name text NOT NULL UNIQUE,
  scope text NOT NULL DEFAULT 'site' CHECK (scope IN ('site', 'single', 'as', 'completed')),
  address text NOT NULL DEFAULT '',
  door_password text NOT NULL DEFAULT '', -- 앞자리 0 보존 (예: '0304')
  office_phone text NOT NULL DEFAULT '',
  consent_done boolean NOT NULL DEFAULT false,
  elevator_cover_done boolean NOT NULL DEFAULT false,
  office_notice_done boolean NOT NULL DEFAULT false,
  memo text NOT NULL DEFAULT '',
  drive_url text NOT NULL DEFAULT '',
  color_index integer NOT NULL DEFAULT 0,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now()
);

-- 2-2. 현장 공정 일정 (현장데이터, 단건, AS, 완료현장 통합)
CREATE TABLE IF NOT EXISTS public.site_tasks (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  site_name text NOT NULL,
  scope text NOT NULL DEFAULT 'site' CHECK (scope IN ('site', 'single', 'as', 'completed')),
  task_date text NOT NULL DEFAULT '', -- 'YYYY.MM.DD' 또는 'YYYY-MM-DD HH:mm'
  task_name text NOT NULL DEFAULT '',
  partner text NOT NULL DEFAULT '',
  is_shared boolean NOT NULL DEFAULT false,
  note text NOT NULL DEFAULT '',
  photos jsonb NOT NULL DEFAULT '[]'::jsonb, -- [{url, name, driveId, ...}]
  manager_1 text NOT NULL DEFAULT '',
  manager_2 text NOT NULL DEFAULT '',
  status text NOT NULL DEFAULT '예정', -- '예정' | '진행 중' | '완료' | 'AS'
  calendar_id text NOT NULL DEFAULT '',
  legacy_updated_at bigint,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_site_tasks_site_name ON public.site_tasks (site_name);
CREATE INDEX IF NOT EXISTS idx_site_tasks_scope ON public.site_tasks (scope);
CREATE INDEX IF NOT EXISTS idx_site_tasks_task_date ON public.site_tasks (task_date);

-- ============================================================================
-- 3. 신규 견적 & To-Do 테이블
-- ============================================================================

-- 3-1. 신규 견적 파이프라인 (신규견적 시트 20개 컬럼)
CREATE TABLE IF NOT EXISTS public.estimates (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  site_customer text NOT NULL,
  stage text NOT NULL DEFAULT '신규 문의',
  inquiry_date text NOT NULL DEFAULT '',
  priority text NOT NULL DEFAULT '보통',
  work_scope text NOT NULL DEFAULT '',
  consult_memo text NOT NULL DEFAULT '',
  source text NOT NULL DEFAULT '',
  first_consult text NOT NULL DEFAULT '',
  expected_start_date text NOT NULL DEFAULT '',
  expected_amount bigint NOT NULL DEFAULT 0,
  quote_sent_date text NOT NULL DEFAULT '',
  client_name text NOT NULL DEFAULT '',
  phone text NOT NULL DEFAULT '',
  manager text NOT NULL DEFAULT '',
  quote_file text NOT NULL DEFAULT '',
  contract_amount bigint NOT NULL DEFAULT 0,
  next_action text NOT NULL DEFAULT '',
  hold_reason text NOT NULL DEFAULT '',
  visit_date text NOT NULL DEFAULT '',
  site_address text NOT NULL DEFAULT '',
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_estimates_stage ON public.estimates (stage);

-- 3-2. To-Do 할 일 (05_ToDo 시트)
CREATE TABLE IF NOT EXISTS public.todos (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  legacy_id text UNIQUE,
  due_date text NOT NULL DEFAULT '',
  status text NOT NULL DEFAULT '미완료',
  manager text NOT NULL DEFAULT '',
  title text NOT NULL,
  tag text NOT NULL DEFAULT '',
  detail text NOT NULL DEFAULT '',
  site_name text NOT NULL DEFAULT '',
  is_important boolean NOT NULL DEFAULT false,
  completed_at timestamptz,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_todos_status ON public.todos (status);
CREATE INDEX IF NOT EXISTS idx_todos_manager ON public.todos (manager);

-- ============================================================================
-- 4. 재무·원장 테이블 (우성아이디_재무DB)
-- ============================================================================

-- 4-1. 현장 수금현황 (수금현황 시트 21개 컬럼)
CREATE TABLE IF NOT EXISTS public.finance_collections (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  legacy_id text UNIQUE,
  site_name text NOT NULL,
  period text NOT NULL DEFAULT '',
  manager text NOT NULL DEFAULT '',
  category text NOT NULL DEFAULT '',
  status text NOT NULL DEFAULT '진행중',
  client_name text NOT NULL DEFAULT '',
  contract_amount bigint NOT NULL DEFAULT 0,
  change_amount bigint NOT NULL DEFAULT 0,
  change_desc text NOT NULL DEFAULT '',
  progress_rate numeric,
  total_amount bigint NOT NULL DEFAULT 0,
  down_payment bigint NOT NULL DEFAULT 0,
  start_payment bigint NOT NULL DEFAULT 0,
  mid_payment_1 bigint NOT NULL DEFAULT 0,
  mid_payment_2 bigint NOT NULL DEFAULT 0,
  mid_payment_3 bigint NOT NULL DEFAULT 0,
  final_payment bigint NOT NULL DEFAULT 0,
  uncollected_amount bigint NOT NULL DEFAULT 0,
  tax_invoice text NOT NULL DEFAULT '미발행',
  note text NOT NULL DEFAULT '',
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_finance_collections_site_name ON public.finance_collections (site_name);

-- 4-2. 월별 비용마감 (비용마감 시트)
CREATE TABLE IF NOT EXISTS public.finance_expenses (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  legacy_id text UNIQUE,
  settle_month text NOT NULL DEFAULT '', -- 'YY.MM' (예: '26.11')
  item text NOT NULL DEFAULT '',
  category text NOT NULL DEFAULT '',
  vendor text NOT NULL DEFAULT '',
  account_number text NOT NULL DEFAULT '', -- 기존 구분2 / 계좌번호
  site_name text NOT NULL DEFAULT '', -- 고정비는 '공통'
  claim_amount bigint NOT NULL DEFAULT 0,
  paid_amount bigint NOT NULL DEFAULT 0,
  balance bigint NOT NULL DEFAULT 0,
  pay_date text NOT NULL DEFAULT '',
  tax_proof text NOT NULL DEFAULT '', -- 세금계산서 / 현금영수증 / 카드 / 무증빙
  note text NOT NULL DEFAULT '',
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_finance_expenses_month ON public.finance_expenses (settle_month);
CREATE INDEX IF NOT EXISTS idx_finance_expenses_site_name ON public.finance_expenses (site_name);

-- 4-3. 고정비 마스터 (비용마감_고정 시트)
CREATE TABLE IF NOT EXISTS public.finance_fixed_master (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  legacy_id text UNIQUE,
  is_active boolean NOT NULL DEFAULT true,
  item text NOT NULL DEFAULT '고정비',
  category text NOT NULL DEFAULT '',
  vendor text NOT NULL DEFAULT '',
  account_number text NOT NULL DEFAULT '',
  base_amount bigint NOT NULL DEFAULT 0,
  pay_day text NOT NULL DEFAULT '',
  note text NOT NULL DEFAULT '',
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now()
);

-- ============================================================================
-- 5. 앱 전역 설정 (공용 PIN, 알림톡 설정 등)
-- ============================================================================
CREATE TABLE IF NOT EXISTS public.app_settings (
  key text PRIMARY KEY,
  value jsonb NOT NULL DEFAULT '{}'::jsonb,
  updated_at timestamptz NOT NULL DEFAULT now()
);

-- ============================================================================
-- 6. updated_at 자동 갱신 트리거 등록
-- ============================================================================
DO $$
DECLARE
  t text;
BEGIN
  FOREACH t IN ARRAY ARRAY[
    'members',
    'task_categories',
    'partners',
    'sites',
    'site_tasks',
    'estimates',
    'todos',
    'finance_collections',
    'finance_expenses',
    'finance_fixed_master',
    'app_settings'
  ]
  LOOP
    EXECUTE format('DROP TRIGGER IF EXISTS trg_%I_updated_at ON public.%I;', t, t);
    EXECUTE format(
      'CREATE TRIGGER trg_%I_updated_at BEFORE UPDATE ON public.%I FOR EACH ROW EXECUTE FUNCTION public.set_updated_at();',
      t, t
    );
  END LOOP;
END $$;

-- ============================================================================
-- 7. Row Level Security (RLS) 정책
-- ============================================================================
-- 주의: 현재 v3 초기 개발·이관 단계에서는 Publishable key(anon) 기반 웹앱 연동을 위해
-- anon/authenticated 모두 접근 가능하도록 설정했습니다.
-- 추후 Supabase Auth 로그인 전환 시 `USING (auth.role() = 'authenticated')` 등으로
-- 정책을 강화할 수 있습니다.
DO $$
DECLARE
  t text;
BEGIN
  FOREACH t IN ARRAY ARRAY[
    'members',
    'task_categories',
    'partners',
    'sites',
    'site_tasks',
    'estimates',
    'todos',
    'finance_collections',
    'finance_expenses',
    'finance_fixed_master',
    'app_settings'
  ]
  LOOP
    EXECUTE format('ALTER TABLE public.%I ENABLE ROW LEVEL SECURITY;', t);
    EXECUTE format('DROP POLICY IF EXISTS "allow_all_access" ON public.%I;', t);
    EXECUTE format(
      'CREATE POLICY "allow_all_access" ON public.%I FOR ALL TO anon, authenticated USING (true) WITH CHECK (true);',
      t
    );
  END LOOP;
END $$;

-- ============================================================================
-- 8. Realtime (실시간 동기화) 활성화
-- ============================================================================
DO $$
DECLARE
  t text;
BEGIN
  FOREACH t IN ARRAY ARRAY[
    'sites',
    'site_tasks',
    'estimates',
    'todos',
    'finance_collections',
    'finance_expenses'
  ]
  LOOP
    IF NOT EXISTS (
      SELECT 1
      FROM pg_publication_tables
      WHERE pubname = 'supabase_realtime'
        AND schemaname = 'public'
        AND tablename = t
    ) THEN
      EXECUTE format('ALTER PUBLICATION supabase_realtime ADD TABLE public.%I;', t);
    END IF;
  END LOOP;
END $$;
