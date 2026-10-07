-- Migration: 20261008000000_cms_dashboard_support_and_fixes.sql
-- Comprehensive CMS Control Plane, Support Tickets, Delete Unused Keys, and Accurate Days Extension

BEGIN;

-- 1. Ensure owner accounts are registered
INSERT INTO public.license_owner_emails (email) VALUES
  ('zenithopensourceprojects@gmail.com'),
  ('roshhellwett@gmail.com')
ON CONFLICT (email) DO NOTHING;

INSERT INTO public.license_admins (user_id, role)
SELECT id, 'owner' FROM auth.users WHERE lower(email) IN ('zenithopensourceprojects@gmail.com', 'roshhellwett@gmail.com')
ON CONFLICT (user_id) DO UPDATE SET role = 'owner';

INSERT INTO public.license_accounts (user_id, email, trial_started_at, trial_expires_at)
SELECT id, lower(email), now(), now() + interval '365 days'
FROM auth.users
WHERE id NOT IN (SELECT user_id FROM public.license_accounts)
ON CONFLICT (user_id) DO NOTHING;

-- 2. Support Tickets Table
CREATE TABLE IF NOT EXISTS public.support_tickets (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  user_id uuid REFERENCES auth.users(id) ON DELETE SET NULL,
  name text NOT NULL,
  email text NOT NULL,
  category text NOT NULL DEFAULT 'general',
  subject text NOT NULL,
  message text NOT NULL,
  status text NOT NULL DEFAULT 'pending',
  admin_notes text,
  replied_at timestamptz,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS support_tickets_status_idx ON public.support_tickets(status);
CREATE INDEX IF NOT EXISTS support_tickets_created_idx ON public.support_tickets(created_at DESC);
CREATE INDEX IF NOT EXISTS support_tickets_email_idx ON public.support_tickets(email);

ALTER TABLE public.support_tickets ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "Anyone can submit support tickets" ON public.support_tickets;
CREATE POLICY "Anyone can submit support tickets"
  ON public.support_tickets FOR INSERT
  WITH CHECK (true);

DROP POLICY IF EXISTS "Support tickets select policy" ON public.support_tickets;
CREATE POLICY "Support tickets select policy"
  ON public.support_tickets FOR SELECT
  USING (
    (auth.uid() IS NOT NULL AND user_id = auth.uid()) OR
    EXISTS (SELECT 1 FROM public.license_admins WHERE user_id = auth.uid()) OR
    EXISTS (SELECT 1 FROM public.license_owner_emails WHERE email = (SELECT lower(email) FROM auth.users WHERE id = auth.uid()))
  );

DROP POLICY IF EXISTS "Support tickets update policy" ON public.support_tickets;
CREATE POLICY "Support tickets update policy"
  ON public.support_tickets FOR UPDATE
  USING (
    EXISTS (SELECT 1 FROM public.license_admins WHERE user_id = auth.uid()) OR
    EXISTS (SELECT 1 FROM public.license_owner_emails WHERE email = (SELECT lower(email) FROM auth.users WHERE id = auth.uid()))
  );

DROP POLICY IF EXISTS "Support tickets delete policy" ON public.support_tickets;
CREATE POLICY "Support tickets delete policy"
  ON public.support_tickets FOR DELETE
  USING (
    EXISTS (SELECT 1 FROM public.license_admins WHERE user_id = auth.uid()) OR
    EXISTS (SELECT 1 FROM public.license_owner_emails WHERE email = (SELECT lower(email) FROM auth.users WHERE id = auth.uid()))
  );

GRANT SELECT, INSERT ON public.support_tickets TO anon, authenticated;
GRANT ALL ON public.support_tickets TO service_role;
GRANT UPDATE, DELETE ON public.support_tickets TO authenticated;

-- 3. Replace public.license_command with full CMS functionality
CREATE OR REPLACE FUNCTION public.license_command(p_actor_id uuid, p_action text, p_payload jsonb DEFAULT '{}'::jsonb)
 RETURNS jsonb
 LANGUAGE plpgsql
 SECURITY DEFINER
 SET search_path TO 'public', 'pg_temp'
AS $function$
declare
  v_user auth.users%rowtype;
  v_target auth.users%rowtype;
  v_account public.license_accounts%rowtype;
  v_target_account public.license_accounts%rowtype;
  v_device public.license_devices%rowtype;
  v_new_device public.license_devices%rowtype;
  v_key public.license_keys%rowtype;
  v_role text;
  v_bucket text;
  v_attempts integer;
  v_limit integer;
  v_window interval;
  v_state text;
  v_expiry timestamptz;
  v_days integer;
  v_offset integer;
  v_query text;
  v_entry jsonb;
  v_before jsonb;
  v_metadata jsonb := '{}';
  v_result jsonb;
  v_target_id uuid;
  v_note text;
  v_client_ip text;
  v_auto_key jsonb;
  v_existing_accounts_on_ip integer;
  v_deleted_count integer;
  v_ticket_id uuid;
begin
  select * into v_user from auth.users where id = p_actor_id;
  if not found then return jsonb_build_object('error', 'Sign in again.', 'code', 'AUTH_REQUIRED', 'httpStatus', 401); end if;
  if v_user.email_confirmed_at is null then
    return jsonb_build_object('state', 'email_unconfirmed', 'canUse', false, 'isAdmin', false,
      'email', lower(v_user.email), 'deviceHint', null, 'trialStartedAt', null, 'trialExpiresAt', null,
      'license', null, 'daysRemaining', 0, 'banReason', null, 'serverNow', now(), 'verifyUntil', now());
  end if;

  -- Admin check for admin_% actions
  if p_action like 'admin\_%' escape '\' then
    select role into v_role from public.license_admins where user_id = p_actor_id;
    if v_role is null and not exists (select 1 from public.license_owner_emails where email = lower(v_user.email)) then
      return jsonb_build_object('error', 'Administrator access is required for this action.', 'code', 'ADMIN_REQUIRED', 'httpStatus', 403);
    end if;
  end if;

  perform pg_advisory_xact_lock(743602610);

  -- Strict IP Tracking & Rate Limiting
  v_client_ip := coalesce(trim(p_payload->>'clientIp'), '');
  if v_client_ip <> '' and not exists (select 1 from public.license_admins where user_id = p_actor_id)
     and not exists (select 1 from public.license_owner_emails where email = lower(v_user.email)) then
    if exists (select 1 from public.license_banned_ips where ip = v_client_ip) then
      return jsonb_build_object('error', 'Network access is blocked due to abusive behavior. Contact support.', 'code', 'IP_BANNED', 'httpStatus', 403);
    end if;

    select count(distinct actor_user_id) into v_existing_accounts_on_ip
      from public.license_ip_tracking
      where ip = v_client_ip and created_at > now() - interval '24 hours' and actor_user_id <> p_actor_id;

    if v_existing_accounts_on_ip >= 2 then
      insert into public.license_banned_ips (ip, reason)
        values (v_client_ip, 'Automated fraud protection: more than 2 accounts created from this IP in 24 hours')
        on conflict (ip) do nothing;
      return jsonb_build_object('error', 'Multiple account creation from this network is blocked. Contact support.', 'code', 'IP_RATE_LIMITED', 'httpStatus', 429);
    end if;

    insert into public.license_ip_tracking (ip, actor_user_id, action)
      values (v_client_ip, p_actor_id, p_action);
  end if;

  -- User Rate Limits
  v_bucket := case when p_action = 'activate' then 'activation' when p_action like 'admin%' then 'admin' else 'verification' end;
  v_limit := case v_bucket when 'activation' then 5 when 'admin' then 100 else 120 end;
  v_window := case v_bucket when 'activation' then interval '10 minutes' else interval '1 minute' end;
  insert into public.license_rate_limits values (p_actor_id, v_bucket, now(), 1)
    on conflict (actor_user_id, bucket) do update set
      attempts = case when license_rate_limits.window_started_at + v_window <= now() then 1 else license_rate_limits.attempts + 1 end,
      window_started_at = case when license_rate_limits.window_started_at + v_window <= now() then now() else license_rate_limits.window_started_at end
    returning attempts into v_attempts;
  if v_attempts > v_limit then
    return jsonb_build_object('error', 'Too many attempts. Please retry later.', 'code', 'RATE_LIMITED', 'httpStatus', 429);
  end if;

  -- 1. ADMIN LIST (Full CMS Aggregation)
  if p_action = 'admin_list' then
    v_offset := greatest(0, least(100000, coalesce((p_payload->>'offset')::integer, 0)));
    v_query := lower(coalesce(p_payload->>'query', ''));
    return jsonb_build_object(
      'keys', coalesce((select jsonb_agg(t) from (select id, key_hint, raw_key,
        case when status = 'active' and expires_at <= now() then 'expired' else status end as status,
        duration_days, bound_user_id, bound_email, bound_device_id, activated_at, expires_at, revoked_at, created_at,
        case when status = 'unused' then duration_days
             when expires_at is null then duration_days
             else greatest(0, ceil(extract(epoch from (expires_at - now())) / 86400)) end as days_remaining
        from public.license_keys
        where v_query = '' or position(v_query in lower(coalesce(bound_email, '') || key_hint || coalesce(raw_key, '') || status)) > 0
        order by created_at desc, id limit 150 offset v_offset) t), '[]'::jsonb),
      'accounts', coalesce((select jsonb_agg(t) from (select
        u.id as user_id,
        lower(u.email) as email,
        coalesce(a.trial_started_at, u.created_at) as trial_started_at,
        coalesce(a.trial_expires_at, u.created_at + interval '30 days') as trial_expires_at,
        a.banned_at,
        a.ban_reason,
        a.created_at,
        a.active_device_id,
        greatest(0, ceil(extract(epoch from (coalesce(a.trial_expires_at, u.created_at + interval '30 days') - now())) / 86400)) as days_remaining,
        case
          when a.banned_at is not null then 'banned'
          when exists(select 1 from public.license_keys k where k.bound_user_id = u.id and k.status = 'active' and k.expires_at > now()) then 'licensed'
          when coalesce(a.trial_expires_at, u.created_at + interval '30 days') > now() then 'trial'
          else 'expired'
        end as status
        from auth.users u
        left join public.license_accounts a on u.id = a.user_id
        where v_query = '' or position(v_query in lower(coalesce(u.email, ''))) > 0
        order by u.created_at desc limit 150 offset v_offset) t), '[]'::jsonb),
      'tickets', coalesce((select jsonb_agg(t) from (select * from public.support_tickets
        where v_query = '' or position(v_query in lower(name || ' ' || email || ' ' || subject || ' ' || message || ' ' || status)) > 0
        order by created_at desc limit 100 offset v_offset) t), '[]'::jsonb),
      'devices', coalesce((select jsonb_agg(t) from (select id, user_id, install_id_hint, status, ban_reason, label, first_seen_at, last_seen_at from public.license_devices order by first_seen_at desc, id limit 100 offset v_offset) t), '[]'::jsonb),
      'events', coalesce((select jsonb_agg(t) from (select * from public.license_events order by created_at desc, id limit 100 offset v_offset) t), '[]'::jsonb),
      'stats', jsonb_build_object(
        'totalUsers', (select count(*) from auth.users),
        'activeLicenses', (select count(*) from public.license_keys where status = 'active' and expires_at > now()),
        'unusedKeys', (select count(*) from public.license_keys where status = 'unused'),
        'pendingTickets', (select count(*) from public.support_tickets where status = 'pending'),
        'totalTickets', (select count(*) from public.support_tickets)
      ),
      'offset', v_offset, 'pageSize', 100);
  end if;

  -- 2. ADMIN GENERATE KEYS
  if p_action = 'admin_generate' then
    v_days := (p_payload->>'durationDays')::integer;
    if v_days is null or v_days not between 1 and 3650 or jsonb_typeof(p_payload->'keys') <> 'array'
      or jsonb_array_length(p_payload->'keys') not between 1 and 100 then
      return jsonb_build_object('error', 'Specify 1–100 keys and 1–3650 days.', 'httpStatus', 400);
    end if;
    for v_entry in select value from jsonb_array_elements(p_payload->'keys') loop
      insert into public.license_keys(key_hash, key_hint, raw_key, duration_days, created_by)
        values (v_entry->>'hash', v_entry->>'hint', v_entry->>'rawKey', v_days, p_actor_id);
    end loop;
    insert into public.license_events(actor_user_id, event_type, metadata) values
      (p_actor_id, p_action, jsonb_build_object('count', jsonb_array_length(p_payload->'keys'), 'durationDays', v_days));
    return jsonb_build_object('success', true);
  end if;

  -- 3. ADMIN DELETE UNUSED KEYS
  if p_action = 'admin_delete_unused_keys' then
    delete from public.license_keys where status = 'unused';
    get diagnostics v_deleted_count = row_count;
    insert into public.license_events(actor_user_id, event_type, metadata) values
      (p_actor_id, p_action, jsonb_build_object('deletedCount', v_deleted_count));
    return jsonb_build_object('success', true, 'deletedCount', v_deleted_count);
  end if;

  -- 4. ADMIN DELETE SINGLE KEY
  if p_action = 'admin_delete_key' then
    select * into v_key from public.license_keys where id = (p_payload->>'licenseId')::uuid;
    if not found then return jsonb_build_object('error', 'License key not found.', 'httpStatus', 404); end if;
    delete from public.license_keys where id = v_key.id;
    insert into public.license_events(actor_user_id, target_user_id, license_id, event_type, metadata) values
      (p_actor_id, v_key.bound_user_id, v_key.id, p_action, jsonb_build_object('deletedKeyHint', v_key.key_hint, 'status', v_key.status));
    return jsonb_build_object('success', true);
  end if;

  -- 5. ADMIN ADJUST TRIAL / EXTEND USER
  if p_action in ('admin_ban_user', 'admin_unban_user', 'admin_adjust_trial') then
    v_target_id := (p_payload->>'userId')::uuid;
    select * into v_target from auth.users where id = v_target_id;
    if not found then return jsonb_build_object('error', 'User account not found.', 'httpStatus', 404); end if;
    
    -- Ensure row in license_accounts exists
    select * into v_account from public.license_accounts where user_id = v_target_id;
    if not found then
      insert into public.license_accounts (user_id, email, trial_started_at, trial_expires_at)
      values (v_target_id, lower(v_target.email), now(), now() + interval '30 days')
      returning * into v_account;
    end if;
    v_before := to_jsonb(v_account);

    if p_action = 'admin_ban_user' then
      v_note := trim(p_payload->>'reason');
      if coalesce(length(v_note), 0) not between 3 and 1000 then return jsonb_build_object('error', 'A ban reason (3–1000 characters) is required.', 'httpStatus', 400); end if;
      update public.license_accounts set banned_at = now(), ban_reason = v_note where user_id = v_target_id;
    elsif p_action = 'admin_unban_user' then
      update public.license_accounts set banned_at = null, ban_reason = null where user_id = v_target_id;
    else
      v_days := (p_payload->>'deltaDays')::integer;
      if v_days is null or v_days = 0 or abs(v_days) > 3650 then return jsonb_build_object('error', 'Specify a nonzero adjustment within ±3650 days.', 'httpStatus', 400); end if;
      if v_days > 0 then
        update public.license_accounts
           set trial_expires_at = greatest(coalesce(trial_expires_at, now()), now()) + v_days * interval '1 day'
         where user_id = v_target_id;
        update public.license_keys
           set expires_at = greatest(coalesce(expires_at, now()), now()) + v_days * interval '1 day',
               duration_days = duration_days + v_days,
               status = 'active'
         where bound_user_id = v_target_id
           and status in ('active', 'expired');
      else
        update public.license_accounts
           set trial_expires_at = trial_expires_at + v_days * interval '1 day'
         where user_id = v_target_id;
        update public.license_keys
           set expires_at = expires_at + v_days * interval '1 day',
               duration_days = greatest(1, duration_days + v_days)
         where bound_user_id = v_target_id
           and status = 'active';
      end if;
    end if;
    insert into public.license_events(actor_user_id, target_user_id, event_type, metadata) values
      (p_actor_id, v_target_id, p_action, jsonb_build_object('before', v_before, 'deltaDays', v_days, 'after', (select to_jsonb(a) from public.license_accounts a where a.user_id = v_target_id)));
    return jsonb_build_object('success', true);
  end if;

  -- 6. ADMIN SUPPORT TICKETS CONTROLS
  if p_action = 'admin_update_ticket' then
    v_ticket_id := (p_payload->>'ticketId')::uuid;
    if not exists (select 1 from public.support_tickets where id = v_ticket_id) then
      return jsonb_build_object('error', 'Support ticket not found.', 'httpStatus', 404);
    end if;
    update public.support_tickets
       set status = coalesce(trim(p_payload->>'status'), status),
           admin_notes = coalesce(trim(p_payload->>'adminNotes'), admin_notes),
           replied_at = case when p_payload->>'replied' = 'true' then now() else replied_at end,
           updated_at = now()
     where id = v_ticket_id;
    return jsonb_build_object('success', true);
  end if;

  if p_action = 'admin_delete_ticket' then
    v_ticket_id := (p_payload->>'ticketId')::uuid;
    delete from public.support_tickets where id = v_ticket_id;
    return jsonb_build_object('success', true);
  end if;

  -- 7. ADMIN DEVICE CONTROLS
  if p_action in ('admin_ban_device', 'admin_unban_device') then
    select * into v_device from public.license_devices where id = (p_payload->>'deviceId')::uuid;
    if not found then return jsonb_build_object('error', 'Device not found.', 'httpStatus', 404); end if;
    v_note := trim(p_payload->>'reason');
    if p_action = 'admin_ban_device' and coalesce(length(v_note), 0) not between 3 and 1000 then return jsonb_build_object('error', 'A device ban reason is required.', 'httpStatus', 400); end if;
    update public.license_devices set status = case when p_action = 'admin_ban_device' then 'banned' else 'active' end,
      ban_reason = case when p_action = 'admin_ban_device' then v_note else null end where id = v_device.id;
    insert into public.license_events(actor_user_id, target_user_id, device_id, event_type, metadata) values
      (p_actor_id, v_device.user_id, v_device.id, p_action, jsonb_build_object('reason', v_note));
    return jsonb_build_object('success', true);
  end if;

  -- 8. ADMIN KEY CONTROLS: ADJUST DAYS, REVOKE, DURATION, EXPIRY, TRANSFER
  if p_action like 'admin%' then
    select * into v_key from public.license_keys where id = (p_payload->>'licenseId')::uuid;
    if not found then return jsonb_build_object('error', 'License not found.', 'httpStatus', 404); end if;
    v_target_id := v_key.bound_user_id;
    v_before := to_jsonb(v_key) - 'key_hash';
    if p_action = 'admin_revoke' then
      update public.license_keys set status = 'revoked', revoked_at = now() where id = v_key.id;
    elsif p_action = 'admin_set_duration' then
      v_days := (p_payload->>'durationDays')::integer;
      if v_key.status <> 'unused' or v_days is null or v_days not between 1 and 3650 then return jsonb_build_object('error', 'Only unused keys can have a duration set (1–3650 days).', 'httpStatus', 400); end if;
      update public.license_keys set duration_days = v_days where id = v_key.id;
    elsif p_action in ('admin_adjust_days', 'admin_set_expiry') then
      if v_key.status = 'revoked' then return jsonb_build_object('error', 'Revoked keys cannot be extended.', 'httpStatus', 400); end if;
      if p_action = 'admin_set_expiry' then
        v_expiry := (p_payload->>'expiresAt')::timestamptz;
        if v_key.status not in ('active', 'expired') or v_expiry is null then return jsonb_build_object('error', 'An activated key and valid expiry are required.', 'httpStatus', 400); end if;
        update public.license_keys set expires_at = v_expiry, status = case when v_expiry > now() then 'active' else 'expired' end where id = v_key.id;
        if v_target_id is not null then
          update public.license_accounts set trial_expires_at = greatest(coalesce(trial_expires_at, now()), v_expiry) where user_id = v_target_id;
        end if;
      else
        v_days := (p_payload->>'deltaDays')::integer;
        if v_days is null or v_days = 0 or abs(v_days) > 3650 then return jsonb_build_object('error', 'Specify a nonzero adjustment within ±3650 days.', 'httpStatus', 400); end if;
        if v_key.status = 'unused' then
          if v_key.duration_days + v_days not between 1 and 3650 then return jsonb_build_object('error', 'Duration must remain within 1–3650 days.', 'httpStatus', 400); end if;
          update public.license_keys set duration_days = duration_days + v_days where id = v_key.id;
        else
          update public.license_keys
             set expires_at = case when v_days > 0 then greatest(coalesce(expires_at, now()), now()) else expires_at end + v_days * interval '1 day',
                 duration_days = greatest(1, duration_days + v_days),
                 status = case when (case when v_days > 0 then greatest(coalesce(expires_at, now()), now()) else expires_at end + v_days * interval '1 day') > now() then 'active' else 'expired' end
           where id = v_key.id;
          if v_target_id is not null then
            if v_days > 0 then
              update public.license_accounts set trial_expires_at = greatest(coalesce(trial_expires_at, now()), now()) + v_days * interval '1 day' where user_id = v_target_id;
            else
              update public.license_accounts set trial_expires_at = trial_expires_at + v_days * interval '1 day' where user_id = v_target_id;
            end if;
          end if;
        end if;
      end if;
    elsif p_action = 'admin_reset_device' then
      if v_target_id is null then return jsonb_build_object('error', 'This key has no bound account.', 'httpStatus', 400); end if;
      update public.license_keys set bound_device_id = null where bound_user_id = v_target_id and status = 'active';
      update public.license_devices set status = 'retired' where id = (select active_device_id from public.license_accounts where user_id = v_target_id);
      update public.license_accounts set active_device_id = null where user_id = v_target_id;
      insert into public.license_events(actor_user_id, target_user_id, license_id, event_type, metadata) values
        (p_actor_id, v_target_id, v_key.id, p_action, jsonb_build_object('before', v_before));
      return jsonb_build_object('success', true);
    elsif p_action = 'admin_transfer' then
      v_note := trim(p_payload->>'verificationNote');
      if coalesce(length(v_note), 0) not between 12 and 1000 then return jsonb_build_object('error', 'Document ownership verification (12–1000 characters).', 'httpStatus', 400); end if;
      if v_key.status <> 'active' then return jsonb_build_object('error', 'Only activated, non-revoked keys can be transferred.', 'httpStatus', 400); end if;
      select * into v_target from auth.users where lower(email) = lower(trim(p_payload->>'targetEmail')) and email_confirmed_at is not null;
      if not found then return jsonb_build_object('error', 'The target must have a verified account.', 'httpStatus', 404); end if;
      select * into v_account from public.license_accounts where user_id = v_target_id;
      select * into v_target_account from public.license_accounts where user_id = v_target.id;
      if v_target_account.banned_at is not null then return jsonb_build_object('error', 'The target account is suspended.', 'httpStatus', 403); end if;
      if coalesce((p_payload->>'preserveDevice')::boolean, false) then
        if v_key.bound_device_id is null then return jsonb_build_object('error', 'This license has no device to preserve.', 'httpStatus', 400); end if;
        select * into v_new_device from public.license_devices where id = v_key.bound_device_id;
        if not found or v_new_device.status <> 'active' then return jsonb_build_object('error', 'A banned or retired device cannot be transferred.', 'httpStatus', 403); end if;
        if exists(select 1 from public.license_keys where bound_device_id = v_new_device.id and bound_user_id = v_target_id and status = 'active' and id <> v_key.id) then
          return jsonb_build_object('error', 'Move the other licenses off this device first.', 'httpStatus', 409);
        end if;
      else
        if length(coalesce(p_payload->>'installHash', '')) <> 64 then return jsonb_build_object('error', 'Specify the destination installation ID.', 'httpStatus', 400); end if;
        select * into v_new_device from public.license_devices where install_id_hash = p_payload->>'installHash';
        if found and (v_new_device.user_id <> v_target.id or v_new_device.status = 'banned') then return jsonb_build_object('error', 'Destination is banned or belongs to another account.', 'httpStatus', 409); end if;
        if not found then
          insert into public.license_devices(user_id, install_id_hash, install_id_hint) values (v_target.id, p_payload->>'installHash', p_payload->>'installHint') returning * into v_new_device;
        end if;
      end if;
      if v_target_account.active_device_id is not null and v_target_account.active_device_id <> v_new_device.id then
        return jsonb_build_object('error', 'The target already has a different bound device. Reset it first.', 'httpStatus', 409);
      end if;
      if not exists(select 1 from public.license_accounts where user_id = v_target.id) then
        insert into public.license_accounts(user_id, email, active_device_id, trial_started_at, trial_expires_at)
          values (v_target.id, lower(v_target.email), v_new_device.id, coalesce(v_account.trial_started_at, now()), coalesce(v_account.trial_expires_at, now()));
      end if;
      update public.license_devices set user_id = v_target.id, status = 'active' where id = v_new_device.id;
      update public.license_accounts set active_device_id = v_new_device.id where user_id = v_target.id;
      update public.license_keys set bound_user_id = v_target.id, bound_email = lower(v_target.email), bound_device_id = v_new_device.id where id = v_key.id;
      if v_target_id <> v_target.id then update public.license_accounts set trial_expires_at = least(trial_expires_at, now()) where user_id = v_target_id; end if;
      v_metadata := jsonb_build_object('verificationNote', v_note, 'previousUserId', v_target_id, 'targetEmail', lower(v_target.email), 'deviceId', v_new_device.id);
      v_target_id := v_target.id;
    else
      return jsonb_build_object('error', 'Unknown administrator action.', 'httpStatus', 400);
    end if;
    insert into public.license_events(actor_user_id, target_user_id, license_id, device_id, event_type, metadata) values
      (p_actor_id, v_target_id, v_key.id, null, p_action, v_metadata || jsonb_build_object('before', v_before, 'after', (select to_jsonb(k) - 'key_hash' from public.license_keys k where k.id = v_key.id)));
    return jsonb_build_object('success', true);
  end if;

  -- 9. CLIENT ACTIONS: status, heartbeat, activate
  if p_action not in ('status', 'heartbeat', 'activate') then return jsonb_build_object('error', 'Unknown licensing action.', 'httpStatus', 400); end if;
  if length(coalesce(p_payload->>'installHash', '')) <> 64 then return jsonb_build_object('error', 'Installation ID is required.', 'httpStatus', 400); end if;

  select * into v_account from public.license_accounts where user_id = p_actor_id;
  select * into v_device from public.license_devices where install_id_hash = p_payload->>'installHash';

  -- Device conflict check
  if found and v_device.user_id <> p_actor_id
     and not exists (select 1 from public.license_admins where user_id = p_actor_id)
     and not exists (select 1 from public.license_owner_emails where email = lower(v_user.email)) then
    insert into public.license_events(actor_user_id, target_user_id, device_id, event_type, metadata)
      values (p_actor_id, v_device.user_id, v_device.id, 'hardware_conflict', jsonb_build_object('reason', 'Device already bound to another account'));
    return jsonb_build_object('error', 'This device is bound to another user account. Transfer the license or use a separate workstation.',
      'code', 'DEVICE_MISMATCH', 'httpStatus', 409);
  end if;

  if not found then
    insert into public.license_devices(user_id, install_id_hash, install_id_hint)
      values (p_actor_id, p_payload->>'installHash', p_payload->>'installHint') returning * into v_device;
  elsif v_device.user_id = p_actor_id and v_device.status = 'active' then
    update public.license_devices set last_seen_at = now() where id = v_device.id;
  end if;

  if v_account.user_id is null then
    insert into public.license_accounts(user_id, email, active_device_id, trial_started_at, trial_expires_at)
      values (p_actor_id, lower(v_user.email), v_device.id, now(), now() + interval '30 days') returning * into v_account;
    insert into public.license_events(actor_user_id, target_user_id, device_id, event_type) values (p_actor_id, p_actor_id, v_device.id, 'trial_started');

    v_auto_key := p_payload->'autoKey';
    if v_auto_key is not null and v_auto_key->>'keyHash' is not null then
      insert into public.license_keys (
        key_hash, key_hint, raw_key, duration_days,
        status, bound_user_id, bound_email, bound_device_id,
        activated_at, expires_at, created_by
      ) values (
        v_auto_key->>'keyHash', v_auto_key->>'keyHint', v_auto_key->>'rawKey', 30,
        'active', p_actor_id, lower(v_user.email), v_device.id,
        now(), now() + interval '30 days', p_actor_id
      ) returning * into v_key;
      insert into public.license_events(actor_user_id, target_user_id, device_id, event_type, metadata)
        values (p_actor_id, p_actor_id, v_device.id, 'auto_key_created', jsonb_build_object('key_hint', v_auto_key->>'keyHint'));
    end if;
  elsif v_account.active_device_id is null then
    update public.license_accounts set active_device_id = v_device.id where user_id = p_actor_id;
  end if;

  if v_account.banned_at is not null then v_state := 'banned';
  elsif v_device.status = 'banned' then v_state := 'device_banned';
  elsif v_account.active_device_id is not null and v_account.active_device_id <> v_device.id then v_state := 'device_mismatch';
  end if;

  if v_state is null and p_action = 'activate' then
    select * into v_key from public.license_keys where key_hash = p_payload->>'keyHash' for update;
    if not found then return jsonb_build_object('error', 'Activation key was not recognized.', 'code', 'KEY_NOT_FOUND', 'httpStatus', 404); end if;
    if v_key.status = 'revoked' then return jsonb_build_object('error', 'This activation key was revoked by administrator.', 'code', 'KEY_REVOKED', 'httpStatus', 409); end if;
    if v_key.status = 'unused' then
      update public.license_keys set status = 'active', bound_user_id = p_actor_id, bound_email = lower(v_user.email),
        bound_device_id = v_device.id, activated_at = now(), expires_at = now() + (v_key.duration_days * interval '1 day') where id = v_key.id returning * into v_key;
      update public.license_accounts set trial_expires_at = greatest(coalesce(trial_expires_at, now()), v_key.expires_at) where user_id = p_actor_id;
      insert into public.license_events(actor_user_id, target_user_id, license_id, device_id, event_type) values (p_actor_id, p_actor_id, v_key.id, v_device.id, 'license_activated');
    elsif v_key.bound_user_id <> p_actor_id then
      return jsonb_build_object('error', 'This key is activated on another account.', 'code', 'KEY_BOUND', 'httpStatus', 409);
    elsif v_key.status = 'active' and v_key.bound_device_id is null then
      update public.license_keys set bound_device_id = v_device.id where id = v_key.id and bound_device_id is null;
      insert into public.license_events(actor_user_id, target_user_id, license_id, device_id, event_type) values (p_actor_id, p_actor_id, v_key.id, v_device.id, 'license_rebound');
    end if;
  end if;

  select * into v_key from public.license_keys where bound_user_id = p_actor_id and bound_device_id = v_device.id
    order by (status = 'active' and expires_at > now()) desc, expires_at desc nulls last limit 1;
  if v_state is null then
    v_state := case when v_key.status = 'active' and v_key.expires_at > now() and v_key.bound_email <> lower(v_user.email) then 'email_mismatch'
      when v_key.status = 'active' and v_key.expires_at > now() then 'licensed'
      when v_key.status = 'revoked' then 'revoked'
      when v_account.trial_expires_at > now() then 'trial' else 'expired' end;
  end if;
  v_expiry := case when v_state = 'licensed' then v_key.expires_at when v_state = 'trial' then v_account.trial_expires_at else null end;
  select role into v_role from public.license_admins where user_id = p_actor_id;
  if v_role is null and exists (select 1 from public.license_owner_emails where email = lower(v_user.email)) then
    v_role := 'owner';
  end if;

  -- Ensure owners and admins always have full access
  if v_role is not null then
    v_state := coalesce(v_state, 'licensed');
    if v_expiry is null or v_expiry <= now() then
      v_expiry := now() + interval '365 days';
    end if;
  end if;

  return jsonb_build_object('state', v_state,
    'canUse', (v_role is not null or v_state in ('trial', 'licensed')),
    'isAdmin', v_role is not null,
    'email', lower(v_user.email), 'deviceHint', v_device.install_id_hint,
    'trialStartedAt', v_account.trial_started_at, 'trialExpiresAt', v_account.trial_expires_at,
    'daysRemaining', case when v_role is not null then 365 when v_expiry is null then 0 else greatest(0, ceil(extract(epoch from (v_expiry - now())) / 86400)) end,
    'banReason', coalesce(v_account.ban_reason, v_device.ban_reason),
    'license', case when v_key.id is null then null else jsonb_build_object('id', v_key.id, 'keyHint', v_key.key_hint, 'rawKey', v_key.raw_key,
      'status', case when v_key.status = 'active' and v_key.expires_at <= now() then 'expired' else v_key.status end,
      'activatedAt', v_key.activated_at, 'expiresAt', v_key.expires_at,
      'daysRemaining', greatest(0, ceil(extract(epoch from (v_key.expires_at - now())) / 86400))) end,
    'serverNow', now(), 'verifyUntil', least(coalesce(v_expiry, now()), now() + interval '90 seconds'));
end;
$function$;

GRANT EXECUTE ON FUNCTION public.license_command(uuid, text, jsonb) TO authenticated, service_role, anon;

COMMIT;
