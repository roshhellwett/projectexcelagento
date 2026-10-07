-- Migration: 20261008010000_fix_support_rls_and_owner_privileges.sql
-- Fixes:
-- 1. Corrects owner email to zenithprojects@icloud.com
-- 2. Eliminates "permission denied for table license_admins" on support ticket submission via SECURITY DEFINER is_admin function and RPC
-- 3. Grants permanent lifetime root privileges to Super Admin / Owner accounts, removing trial restrictions, expirations, ban vulnerability, and day counters

BEGIN;

-- 1. Ensure Owner emails table exists and has correct emails
CREATE TABLE IF NOT EXISTS public.license_owner_emails (
  email text PRIMARY KEY
);

DELETE FROM public.license_owner_emails WHERE email = 'zenithopensourceprojects@gmail.com';
INSERT INTO public.license_owner_emails (email) VALUES
  ('roshhellwett@gmail.com'),
  ('zenithprojects@icloud.com')
ON CONFLICT (email) DO NOTHING;

-- Update license_admins with owner roles
INSERT INTO public.license_admins (user_id, role)
SELECT id, 'owner' FROM auth.users WHERE lower(email) IN ('zenithprojects@icloud.com', 'roshhellwett@gmail.com')
ON CONFLICT (user_id) DO UPDATE SET role = 'owner';

-- 2. Create SECURITY DEFINER function to check admin status without leaking license_admins table permissions
CREATE OR REPLACE FUNCTION public.is_admin(p_user_id uuid)
RETURNS boolean
LANGUAGE plpgsql
SECURITY DEFINER
STABLE
SET search_path = public, auth, pg_temp
AS $$
BEGIN
  IF p_user_id IS NULL THEN
    RETURN false;
  END IF;

  RETURN EXISTS (
    SELECT 1 FROM public.license_admins WHERE user_id = p_user_id
  ) OR EXISTS (
    SELECT 1 FROM public.license_owner_emails loe
    JOIN auth.users u ON lower(u.email) = loe.email
    WHERE u.id = p_user_id
  );
END;
$$;

GRANT EXECUTE ON FUNCTION public.is_admin(uuid) TO anon, authenticated, service_role;

-- 3. Fix support_tickets RLS policies so normal and anonymous users can submit tickets with zero permission issues
DROP POLICY IF EXISTS "Anyone can submit support tickets" ON public.support_tickets;
CREATE POLICY "Anyone can submit support tickets"
  ON public.support_tickets FOR INSERT
  WITH CHECK (true);

DROP POLICY IF EXISTS "Support tickets select policy" ON public.support_tickets;
CREATE POLICY "Support tickets select policy"
  ON public.support_tickets FOR SELECT
  USING (
    (auth.uid() IS NOT NULL AND user_id = auth.uid()) OR
    public.is_admin(auth.uid())
  );

DROP POLICY IF EXISTS "Support tickets update policy" ON public.support_tickets;
CREATE POLICY "Support tickets update policy"
  ON public.support_tickets FOR UPDATE
  USING (public.is_admin(auth.uid()));

DROP POLICY IF EXISTS "Support tickets delete policy" ON public.support_tickets;
CREATE POLICY "Support tickets delete policy"
  ON public.support_tickets FOR DELETE
  USING (public.is_admin(auth.uid()));

GRANT SELECT, INSERT ON public.support_tickets TO anon, authenticated;
GRANT ALL ON public.support_tickets TO service_role;
GRANT UPDATE, DELETE ON public.support_tickets TO authenticated;

-- 4. Dedicated RPC to submit support tickets with zero permission friction
CREATE OR REPLACE FUNCTION public.submit_support_ticket(
  p_name text,
  p_email text,
  p_category text,
  p_subject text,
  p_message text,
  p_user_id uuid DEFAULT NULL
)
RETURNS jsonb
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = public, pg_temp
AS $$
DECLARE
  v_ticket_id uuid;
BEGIN
  INSERT INTO public.support_tickets (name, email, category, subject, message, user_id, status)
  VALUES (
    trim(p_name),
    lower(trim(p_email)),
    coalesce(nullif(trim(p_category), ''), 'general'),
    trim(p_subject),
    trim(p_message),
    p_user_id,
    'pending'
  )
  RETURNING id INTO v_ticket_id;

  RETURN jsonb_build_object('success', true, 'id', v_ticket_id);
END;
$$;

GRANT EXECUTE ON FUNCTION public.submit_support_ticket(text, text, text, text, text, uuid) TO anon, authenticated, service_role;

-- 5. Update license_command with Admin vs Normal User differentiation and root immunity
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
  v_target_device public.license_devices%rowtype;
  v_key public.license_keys%rowtype;
  v_auto_key jsonb;
  v_state text;
  v_expiry timestamptz;
  v_role text;
  v_is_owner boolean := false;
  v_is_admin boolean := false;
  v_client_ip text;
  v_attempts integer;
  v_bucket text;
  v_limit integer;
  v_window interval;
  v_days integer;
  v_count integer;
  v_entry jsonb;
  v_query text;
  v_offset integer;
  v_deleted_count integer;
  v_delta_days integer;
  v_new_expiry timestamptz;
  v_now_expiry timestamptz;
  v_ip_record public.license_ip_tracking%rowtype;
begin
  select * into v_user from auth.users where id = p_actor_id;
  if not found then
    return jsonb_build_object('error', 'Sign in again to verify access.', 'code', 'AUTH_REQUIRED', 'httpStatus', 401);
  end if;

  select * into v_account from public.license_accounts where user_id = p_actor_id;
  v_client_ip := coalesce(p_payload->>'clientIp', '');

  -- Determine if actor is Admin or Owner
  select role into v_role from public.license_admins where user_id = p_actor_id;
  if v_role is null and exists (select 1 from public.license_owner_emails where email = lower(v_user.email)) then
    v_role := 'owner';
  end if;
  v_is_owner := (v_role = 'owner');
  v_is_admin := (v_role is not null);

  -- IP Rate Limiting & Tracking
  if v_client_ip <> '' then
    if v_client_ip <> '' and not v_is_admin then
      select count(*) into v_count from public.license_ip_tracking
        where ip = v_client_ip and created_at > now() - interval '24 hours';
      if v_count >= 10 and p_action = 'status' and v_account.user_id is null then
        return jsonb_build_object('error', 'Network access rate limit exceeded.', 'code', 'IP_RATE_LIMITED', 'httpStatus', 429);
      end if;
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

  -- 1. ADMIN LIST (Full CMS Aggregation with distinct Admin/Owner representation)
  if p_action = 'admin_list' then
    if not v_is_admin then
      return jsonb_build_object('error', 'Access restricted to administrators.', 'code', 'FORBIDDEN', 'httpStatus', 403);
    end if;

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
        case
          when exists (select 1 from public.license_admins la where la.user_id = u.id and la.role = 'owner')
            or exists (select 1 from public.license_owner_emails loe where loe.email = lower(u.email)) then 'owner'
          when exists (select 1 from public.license_admins la where la.user_id = u.id) then 'admin'
          else null
        end as role,
        (exists (select 1 from public.license_admins la where la.user_id = u.id)
          or exists (select 1 from public.license_owner_emails loe where loe.email = lower(u.email))) as is_admin,
        coalesce(a.trial_started_at, u.created_at) as trial_started_at,
        case
          when exists (select 1 from public.license_admins la where la.user_id = u.id)
            or exists (select 1 from public.license_owner_emails loe where loe.email = lower(u.email)) then null
          else coalesce(a.trial_expires_at, u.created_at + interval '30 days')
        end as trial_expires_at,
        a.banned_at,
        a.ban_reason,
        a.created_at,
        a.active_device_id,
        case
          when exists (select 1 from public.license_admins la where la.user_id = u.id)
            or exists (select 1 from public.license_owner_emails loe where loe.email = lower(u.email)) then null
          else greatest(0, ceil(extract(epoch from (coalesce(a.trial_expires_at, u.created_at + interval '30 days') - now())) / 86400))
        end as days_remaining,
        case
          when exists (select 1 from public.license_admins la where la.user_id = u.id and la.role = 'owner')
            or exists (select 1 from public.license_owner_emails loe where loe.email = lower(u.email)) then 'owner'
          when exists (select 1 from public.license_admins la where la.user_id = u.id) then 'admin'
          when a.banned_at is not null then 'banned'
          when exists(select 1 from public.license_keys k where k.bound_user_id = u.id and k.status = 'active' and k.expires_at > now()) then 'licensed'
          when coalesce(a.trial_expires_at, u.created_at + interval '30 days') > now() then 'trial'
          else 'expired'
        end as status
        from auth.users u
        left join public.license_accounts a on u.id = a.user_id
        where v_query = '' or position(v_query in lower(coalesce(u.email, ''))) > 0
        order by (case when exists (select 1 from public.license_owner_emails loe where loe.email = lower(u.email)) then 1 else 2 end), u.created_at desc limit 150 offset v_offset) t), '[]'::jsonb),
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
    if not v_is_admin then
      return jsonb_build_object('error', 'Access restricted to administrators.', 'code', 'FORBIDDEN', 'httpStatus', 403);
    end if;

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
    if not v_is_admin then
      return jsonb_build_object('error', 'Access restricted to administrators.', 'code', 'FORBIDDEN', 'httpStatus', 403);
    end if;

    delete from public.license_keys where status = 'unused';
    get diagnostics v_deleted_count = row_count;
    insert into public.license_events(actor_user_id, event_type, metadata) values
      (p_actor_id, p_action, jsonb_build_object('deletedCount', v_deleted_count));
    return jsonb_build_object('success', true, 'deletedCount', v_deleted_count);
  end if;

  -- 4. ADMIN DELETE SINGLE KEY
  if p_action = 'admin_delete_key' then
    if not v_is_admin then
      return jsonb_build_object('error', 'Access restricted to administrators.', 'code', 'FORBIDDEN', 'httpStatus', 403);
    end if;

    select * into v_key from public.license_keys where id = (p_payload->>'licenseId')::uuid;
    if not found then return jsonb_build_object('error', 'License key not found.', 'httpStatus', 404); end if;
    delete from public.license_keys where id = v_key.id;
    insert into public.license_events(actor_user_id, target_user_id, license_id, event_type, metadata) values
      (p_actor_id, v_key.bound_user_id, v_key.id, p_action, jsonb_build_object('deletedKeyHint', v_key.key_hint, 'status', v_key.status));
    return jsonb_build_object('success', true);
  end if;

  -- 5. ADMIN ADJUST TRIAL / DAYS (Protected against modifying admins)
  if p_action = 'admin_adjust_trial' then
    if not v_is_admin then
      return jsonb_build_object('error', 'Access restricted to administrators.', 'code', 'FORBIDDEN', 'httpStatus', 403);
    end if;

    select * into v_target from auth.users where id = (p_payload->>'userId')::uuid;
    if not found then return jsonb_build_object('error', 'User not found.', 'httpStatus', 404); end if;

    -- Root Protection
    if exists (select 1 from public.license_admins where user_id = v_target.id) or
       exists (select 1 from public.license_owner_emails where email = lower(v_target.email)) then
      return jsonb_build_object('error', 'Administrator and Owner accounts have permanent lifetime access and do not require manual extensions.', 'httpStatus', 400);
    end if;

    v_delta_days := coalesce((p_payload->>'deltaDays')::integer, 30);
    select * into v_target_account from public.license_accounts where user_id = v_target.id;
    v_now_expiry := coalesce(v_target_account.trial_expires_at, now());
    v_new_expiry := greatest(now(), v_now_expiry) + (v_delta_days * interval '1 day');

    if v_target_account.user_id is null then
      insert into public.license_accounts(user_id, email, trial_started_at, trial_expires_at)
        values (v_target.id, lower(v_target.email), now(), v_new_expiry);
    else
      update public.license_accounts set trial_expires_at = v_new_expiry where user_id = v_target.id;
    end if;

    insert into public.license_events(actor_user_id, target_user_id, event_type, metadata) values
      (p_actor_id, v_target.id, p_action, jsonb_build_object('deltaDays', v_delta_days, 'newExpiry', v_new_expiry));
    return jsonb_build_object('success', true, 'newExpiry', v_new_expiry);
  end if;

  -- 6. ADMIN BAN USER (Root Protected)
  if p_action = 'admin_ban_user' then
    if not v_is_admin then
      return jsonb_build_object('error', 'Access restricted to administrators.', 'code', 'FORBIDDEN', 'httpStatus', 403);
    end if;

    select * into v_target from auth.users where id = (p_payload->>'userId')::uuid;
    if not found then return jsonb_build_object('error', 'User not found.', 'httpStatus', 404); end if;

    -- Root Protection
    if exists (select 1 from public.license_admins where user_id = v_target.id) or
       exists (select 1 from public.license_owner_emails where email = lower(v_target.email)) then
      return jsonb_build_object('error', 'Administrator and Owner accounts are root protected and cannot be banned or restricted.', 'httpStatus', 400);
    end if;

    update public.license_accounts set banned_at = now(), ban_reason = coalesce(p_payload->>'reason', 'Banned by admin')
      where user_id = v_target.id;
    insert into public.license_events(actor_user_id, target_user_id, event_type, metadata) values
      (p_actor_id, v_target.id, p_action, jsonb_build_object('reason', p_payload->>'reason'));
    return jsonb_build_object('success', true);
  end if;

  -- 7. ADMIN UNBAN USER
  if p_action = 'admin_unban_user' then
    if not v_is_admin then
      return jsonb_build_object('error', 'Access restricted to administrators.', 'code', 'FORBIDDEN', 'httpStatus', 403);
    end if;

    update public.license_accounts set banned_at = null, ban_reason = null where user_id = (p_payload->>'userId')::uuid;
    insert into public.license_events(actor_user_id, target_user_id, event_type) values
      (p_actor_id, (p_payload->>'userId')::uuid, p_action);
    return jsonb_build_object('success', true);
  end if;

  -- 8. ADMIN SUPPORT TICKET ACTIONS
  if p_action = 'admin_update_ticket' then
    if not v_is_admin then
      return jsonb_build_object('error', 'Access restricted to administrators.', 'code', 'FORBIDDEN', 'httpStatus', 403);
    end if;

    update public.support_tickets
    set status = coalesce(p_payload->>'status', status),
        admin_notes = coalesce(p_payload->>'adminNotes', admin_notes),
        replied_at = case when (p_payload->>'markReplied')::boolean = true then now() else replied_at end,
        updated_at = now()
    where id = (p_payload->>'ticketId')::uuid;
    return jsonb_build_object('success', true);
  end if;

  if p_action = 'admin_delete_ticket' then
    if not v_is_admin then
      return jsonb_build_object('error', 'Access restricted to administrators.', 'code', 'FORBIDDEN', 'httpStatus', 403);
    end if;

    delete from public.support_tickets where id = (p_payload->>'ticketId')::uuid;
    return jsonb_build_object('success', true);
  end if;

  -- 9. USER LICENSE VERIFICATION & HEARTBEAT (verify_license)
  select * into v_device from public.license_devices where install_id_hash = p_payload->>'installHash';
  if v_device.id is null then
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
    end if;
  elsif v_account.active_device_id is null then
    update public.license_accounts set active_device_id = v_device.id where user_id = p_actor_id;
  end if;

  if not v_is_admin then
    if v_account.banned_at is not null then v_state := 'banned';
    elsif v_device.status = 'banned' then v_state := 'device_banned';
    elsif v_account.active_device_id is not null and v_account.active_device_id <> v_device.id then v_state := 'device_mismatch';
    end if;
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
    v_state := case
      when v_key.status = 'active' and v_key.expires_at > now() and v_key.bound_email <> lower(v_user.email) then 'email_mismatch'
      when v_key.status = 'active' and v_key.expires_at > now() then 'licensed'
      when v_key.status = 'revoked' then 'revoked'
      when v_account.trial_expires_at > now() then 'trial' else 'expired' end;
  end if;

  -- Ensure owners and admins have permanent lifetime immunity and root access
  if v_is_admin then
    v_state := 'licensed';
    v_expiry := null;
  else
    v_expiry := case when v_state = 'licensed' then v_key.expires_at when v_state = 'trial' then v_account.trial_expires_at else null end;
  end if;

  return jsonb_build_object('state', v_state,
    'canUse', (v_is_admin or v_state in ('trial', 'licensed')),
    'isAdmin', v_is_admin,
    'isOwner', v_is_owner,
    'role', v_role,
    'isLifetime', v_is_admin,
    'email', lower(v_user.email), 'deviceHint', v_device.install_id_hint,
    'trialStartedAt', v_account.trial_started_at,
    'trialExpiresAt', case when v_is_admin then null else v_account.trial_expires_at end,
    'daysRemaining', case when v_is_admin then null when v_expiry is null then 0 else greatest(0, ceil(extract(epoch from (v_expiry - now())) / 86400)) end,
    'banReason', case when v_is_admin then null else coalesce(v_account.ban_reason, v_device.ban_reason) end,
    'license', case when v_key.id is null then null else jsonb_build_object('id', v_key.id, 'keyHint', v_key.key_hint, 'rawKey', v_key.raw_key,
      'status', case when v_key.status = 'active' and v_key.expires_at <= now() then 'expired' else v_key.status end,
      'activatedAt', v_key.activated_at, 'expiresAt', v_key.expires_at,
      'daysRemaining', greatest(0, ceil(extract(epoch from (v_key.expires_at - now())) / 86400))) end,
    'serverNow', now(), 'verifyUntil', now() + interval '90 seconds');
end;
$function$;

GRANT EXECUTE ON FUNCTION public.license_command(uuid, text, jsonb) TO authenticated, service_role, anon;

COMMIT;
