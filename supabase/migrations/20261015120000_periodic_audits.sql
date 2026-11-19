-- Periodic (quarterly / monthly / half-yearly) internal audits.
-- Kept fully separate from the annual Audit Status page, which stays on task_logs.

-- Checklist templates, usually uploaded from Excel.
-- items: [{ "id": "...", "section": "...", "item": "...", "guidance": "..." }]
create table public.periodic_audit_templates (
    id uuid primary key default gen_random_uuid(),
    name text not null unique,
    items jsonb not null default '[]'::jsonb,
    created_at timestamptz not null default now(),
    updated_at timestamptz not null default now()
);

-- One client running one checklist for one financial year.
create table public.periodic_audit_engagements (
    id uuid primary key default gen_random_uuid(),
    client_id uuid not null references public.clients(client_id) on delete cascade,
    template_id uuid not null references public.periodic_audit_templates(id) on delete restrict,
    fy text not null,                                   -- e.g. FY_2026_27
    frequency text not null check (frequency in ('monthly', 'quarterly', 'half_yearly')),
    assigned_to uuid references public.users(id) on delete set null,
    is_archived boolean not null default false,
    created_at timestamptz not null default now(),
    unique (client_id, template_id, fy)
);

-- One row per period (Q1, M04, H1 ...) holding the due date and overall stage.
create table public.periodic_audit_periods (
    id uuid primary key default gen_random_uuid(),
    engagement_id uuid not null references public.periodic_audit_engagements(id) on delete cascade,
    period text not null,
    due_date date,
    status text not null default 'Not Started',
    remarks text,
    updated_by uuid references public.users(id) on delete set null,
    updated_at timestamptz not null default now(),
    unique (engagement_id, period)
);

-- Status of each checklist item within a period.
create table public.periodic_audit_item_status (
    id uuid primary key default gen_random_uuid(),
    engagement_id uuid not null references public.periodic_audit_engagements(id) on delete cascade,
    period text not null,
    item_id text not null,
    status text not null default 'Pending',             -- Pending | Done | Query | N/A
    remark text,
    updated_by uuid references public.users(id) on delete set null,
    updated_at timestamptz not null default now(),
    unique (engagement_id, period, item_id)
);
create index periodic_audit_item_status_eng_idx on public.periodic_audit_item_status (engagement_id, period);

do $$
declare
    t text;
begin
    foreach t in array array['periodic_audit_templates', 'periodic_audit_engagements', 'periodic_audit_periods', 'periodic_audit_item_status'] loop
        execute format('alter table public.%I enable row level security', t);
        execute format(
            'create policy "Office OS staff full access" on public.%I
                 for all to authenticated
                 using (((select auth.jwt()) -> ''app_metadata'' ->> ''app_user_id'') is not null)
                 with check (((select auth.jwt()) -> ''app_metadata'' ->> ''app_user_id'') is not null)',
            t);
    end loop;
end $$;
