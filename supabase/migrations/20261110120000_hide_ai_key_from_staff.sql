-- The OpenRouter key (app_settings.office_os_gemini_key) was readable by every logged-in
-- staff browser. Only Code.js needs it, and it uses the secret key, which bypasses RLS.
-- Staff keep full access to every other settings row.
drop policy if exists "Office OS staff full access" on public.app_settings;

create policy "Office OS staff full access" on public.app_settings
    for all to authenticated
    using (
        ((select auth.jwt()) -> 'app_metadata' ->> 'app_user_id') is not null
        and setting_key <> 'office_os_gemini_key'
    )
    with check (
        ((select auth.jwt()) -> 'app_metadata' ->> 'app_user_id') is not null
        and setting_key <> 'office_os_gemini_key'
    );
