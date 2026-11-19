-- Lock the database down to logged-in Office OS staff.
--
-- The browser now uses the public (publishable) key. RLS was already enabled on every
-- table with no policies, so that key alone can read nothing. Staff get access through a
-- Supabase Auth session issued by the `office-login` edge function after a name + PIN
-- check; that function stamps `app_metadata.app_user_id` on the auth user.
--
-- app_metadata can only be written with a secret key, so self sign-ups and anonymous
-- sign-ins (which also get the `authenticated` role) never pass this check.

do $$
declare
    t text;
begin
    for t in select tablename from pg_tables where schemaname = 'public' loop
        execute format(
            'create policy "Office OS staff full access" on public.%I
                 for all to authenticated
                 using (((select auth.jwt()) -> ''app_metadata'' ->> ''app_user_id'') is not null)
                 with check (((select auth.jwt()) -> ''app_metadata'' ->> ''app_user_id'') is not null)',
            t);
    end loop;
end $$;

-- Superseded: it let any authenticated session (including self sign-ups) touch invoices.
drop policy if exists "Allow Authenticated Access" on public.invoices;

-- Failed login attempts, used by office-login to lock out PIN guessing.
-- RLS on with no policies: only the edge function (secret key) can read or write it.
create table public.login_attempts (
    id bigint generated always as identity primary key,
    name_key text not null,
    attempted_at timestamptz not null default now()
);
create index login_attempts_name_time_idx on public.login_attempts (name_key, attempted_at);
alter table public.login_attempts enable row level security;
