-- Migration: Anti-abuse hardware device protection, IP rate limiting, and auto-provisioning 30-day keys
BEGIN;

ALTER TABLE IF EXISTS public.license_keys ADD COLUMN IF NOT EXISTS raw_key text;
ALTER TABLE IF EXISTS public.license_keys ALTER COLUMN created_by DROP NOT NULL;

CREATE TABLE IF NOT EXISTS public.license_ip_tracking (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  ip text NOT NULL,
  actor_user_id uuid NOT NULL,
  action text NOT NULL,
  created_at timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_license_ip_tracking ON public.license_ip_tracking(ip, created_at);

CREATE TABLE IF NOT EXISTS public.license_banned_ips (
  ip text PRIMARY KEY,
  banned_at timestamptz NOT NULL DEFAULT now(),
  reason text NOT NULL
);

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
    if v_role is null or (v_role = 'support' and p_action not in ('admin_list', 'admin_reset_device', 'admin_transfer')) then
      return jsonb_build_object('error', 'Administrator access is required for this action.', 'code', 'ADMIN_REQUIRED', 'httpStatus', 403);
    end if;
  end if;

  perform pg_advisory_xact_lock(743602610);

  -- Strict IP Tracking & Rate Limiting
  v_client_ip := coalesce(trim(p_payload->>'clientIp'), '');
  if v_client_ip <> '' and not exists (select 1 from public.license_admins where user_id = p_actor_id) then
    if exists (select 1 from public.license_banned_ips where ip = v_client_ip) then
      return jsonb_build_object('error', 'Network access is blocked due to abusive behavior. Contact zenithprojects@icloud.com.', 'code', 'IP_BANNED', 'httpStatus', 403);
    end if;

    select count(distinct actor_user_id) into v_existing_accounts_on_ip
      from public.license_ip_tracking
      where ip = v_client_ip and created_at > now() - interval '24 hours' and actor_user_id <> p_actor_id;

    if v_existing_accounts_on_ip >= 2 then
      insert into public.license_banned_ips (ip, reason)
        values (v_client_ip, 'Automated fraud protection: more than 2 accounts created from this IP in 24 hours')
        on conflict (ip) do nothing;
      return jsonb_build_object('error', 'Multiple account creation from this network is blocked. Contact zenithprojects@icloud.com.', 'code', 'IP_RATE_LIMITED', 'httpStatus', 429);
    end if;

    insert into public.license_ip_tracking (ip, actor_user_id, action)
      values (v_client_ip, p_actor_id, p_action);
  end if;

  -- User Rate Limits
  v_bucket := case when p_action = 'activate' then 'activation' when p_action like 'admin%' then 'admin' else 'verification' end;
  v_limit := case v_bucket when 'activation' then 5 when 'admin' then 60 else 120 end;
  v_window := case v_bucket when 'activation' then interval '10 minutes' else interval '1 minute' end;
  insert into public.license_rate_limits values (p_actor_id, v_bucket, now(), 1)
    on conflict (actor_user_id, bucket) do update set
      attempts = case when license_rate_limits.window_started_at + v_window <= now() then 1 else license_rate_limits.attempts + 1 end,
      window_started_at = case when license_rate_limits.window_started_at + v_window <= now() then now() else license_rate_limits.window_started_at end
    returning attempts into v_attempts;
  if v_attempts > v_limit then
    return jsonb_build_object('error', 'Too many attempts. Please retry later.', 'code', 'RATE_LIMITED', 'httpStatus', 429);
  end if;

  -- Admin Actions
  if p_action = 'admin_list' then
    v_offset := greatest(0, least(100000, coalesce((p_payload->>'offset')::integer, 0)));
    v_query := lower(coalesce(p_payload->>'query', ''));
    return jsonb_build_object(
      'keys', coalesce((select jsonb_agg(t) from (select id, key_hint, raw_key,
        case when status = 'active' and expires_at <= now() then 'expired' else status end as status,
        duration_days, bound_user_id, bound_email, bound_device_id, activated_at, expires_at, revoked_at, created_at
        from public.license_keys where v_query = '' or position(v_query in lower(coalesce(bound_email, '') || key_hint || coalesce(raw_key, '') || status)) > 0
        order by created_at desc, id limit 100 offset v_offset) t), '[]'::jsonb),
      'accounts', coalesce((select jsonb_agg(t) from (select a.*,
        case when a.banned_at is not null then 'banned' when exists(select 1 from public.license_keys k where k.bound_user_id = a.user_id and k.status = 'active' and k.expires_at > now()) then 'licensed'
          when a.trial_expires_at > now() then 'trial' else 'expired' end as status
        from public.license_accounts a where v_query = '' or position(v_query in lower(email)) > 0
        order by created_at desc, user_id limit 100 offset v_offset) t), '[]'::jsonb),
      'devices', coalesce((select jsonb_agg(t) from (select id, user_id, install_id_hint, status, ban_reason, label, first_seen_at, last_seen_at from public.license_devices order by first_seen_at desc, id limit 100 offset v_offset) t), '[]'::jsonb),
      'events', coalesce((select jsonb_agg(t) from (select * from public.license_events order by created_at desc, id limit 100 offset v_offset) t), '[]'::jsonb),
      'offset', v_offset, 'pageSize', 100);
  end if;

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

  if p_action in ('admin_ban_user', 'admin_unban_user', 'admin_adjust_trial') then
    v_target_id := (p_payload->>'userId')::uuid;
    select * into v_account from public.license_accounts where user_id = v_target_id;
    if not found then return jsonb_build_object('error', 'Account not found.', 'httpStatus', 404); end if;
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
      update public.license_accounts set trial_expires_at = trial_expires_at + v_days * interval '1 day' where user_id = v_target_id;
      update public.license_keys set expires_at = expires_at + v_days * interval '1 day' where bound_user_id = v_target_id and status = 'active';
    end if;
    insert into public.license_events(actor_user_id, target_user_id, event_type, metadata) values
      (p_actor_id, v_target_id, p_action, jsonb_build_object('before', v_before, 'after', (select to_jsonb(a) from public.license_accounts a where a.user_id = v_target_id)));
    return jsonb_build_object('success', true);
  end if;

  if p_action in ('admin_ban_device', 'admin_unban_device') then
    select * into v_device from public.license_devices where id = (p_payload->>'deviceId')::uuid;
    if not found then return jsonb_build_object('error', 'Device not found.', 'httpStatus', 404); end if;
    v_note := trim(p_payload->>'reason');
    if p_action = 'admin_ban_device' and coalesce(length(v_note), 0) not between 3 and 1000 then return jsonb_build_object('error', 'A device ban reason is required.', 'httpStatus', 400); end if;
    update public.license_devices set status = case when p_action = 'admin_ban_device' then 'banned' else 'active' end,
      ban_reason = case when p_action = 'admin_ban_device' then v_note else null end where id = v_device.id;
    insert into public.license_events(actor_user_id, device_id, event_type, metadata) values
      (p_actor_id, v_device.id, p_action, jsonb_build_object('status', case when p_action = 'admin_ban_device' then 'banned' else 'active' end, 'reason', v_note));
    return jsonb_build_object('success', true);
  end if;

  if p_action = 'admin_reset_device' then
    select * into v_key from public.license_keys where id = (p_payload->>'licenseId')::uuid;
    if not found then return jsonb_build_object('error', 'License not found.', 'httpStatus', 404); end if;
    update public.license_keys set bound_device_id = null where id = v_key.id;
    if v_key.bound_user_id is not null then
      update public.license_accounts set active_device_id = null where user_id = v_key.bound_user_id;
    end if;
    insert into public.license_events(actor_user_id, target_user_id, license_id, device_id, event_type) values
      (p_actor_id, v_key.bound_user_id, v_key.id, v_key.bound_device_id, p_action);
    return jsonb_build_object('success', true);
  end if;

  if p_action in ('admin_adjust_days', 'admin_set_duration', 'admin_set_expiry', 'admin_revoke') then
    select * into v_key from public.license_keys where id = (p_payload->>'licenseId')::uuid;
    if not found then return jsonb_build_object('error', 'License not found.', 'httpStatus', 404); end if;
    v_before := to_jsonb(v_key);
    if p_action = 'admin_revoke' then
      update public.license_keys set status = 'revoked', revoked_at = now() where id = v_key.id;
    elsif p_action = 'admin_adjust_days' then
      v_days := (p_payload->>'deltaDays')::integer;
      if v_days is null or v_days = 0 or abs(v_days) > 3650 then return jsonb_build_object('error', 'Specify a nonzero adjustment within ±3650 days.', 'httpStatus', 400); end if;
      update public.license_keys set expires_at = case when status = 'active' then expires_at + v_days * interval '1 day' else expires_at end,
        duration_days = duration_days + v_days where id = v_key.id;
    elsif p_action = 'admin_set_duration' then
      v_days := (p_payload->>'durationDays')::integer;
      if v_days is null or v_days not between 1 and 3650 then return jsonb_build_object('error', 'Specify a duration between 1 and 3650 days.', 'httpStatus', 400); end if;
      update public.license_keys set duration_days = v_days,
        expires_at = case when activated_at is not null then activated_at + v_days * interval '1 day' else null end where id = v_key.id;
    else
      v_expiry := (p_payload->>'expiresAt')::timestamptz;
      if v_expiry is null or v_expiry <= now() then return jsonb_build_object('error', 'Expiry must be in the future.', 'httpStatus', 400); end if;
      update public.license_keys set expires_at = v_expiry, duration_days = greatest(1, ceil(extract(epoch from (v_expiry - coalesce(activated_at, now()))) / 86400)) where id = v_key.id;
    end if;
    insert into public.license_events(actor_user_id, target_user_id, license_id, event_type, metadata) values
      (p_actor_id, v_key.bound_user_id, v_key.id, p_action, jsonb_build_object('before', v_before, 'after', (select to_jsonb(k) from public.license_keys k where k.id = v_key.id)));
    return jsonb_build_object('success', true);
  end if;

  if p_action = 'admin_transfer' then
    select * into v_key from public.license_keys where id = (p_payload->>'licenseId')::uuid;
    if not found then return jsonb_build_object('error', 'License not found.', 'httpStatus', 404); end if;
    v_note := trim(p_payload->>'verificationNote');
    if coalesce(length(v_note), 0) not between 12 and 1000 then return jsonb_build_object('error', 'Verification note must be 12–1000 characters.', 'httpStatus', 400); end if;
    select * into v_target from auth.users where lower(email) = lower(trim(p_payload->>'targetEmail'));
    if not found then return jsonb_build_object('error', 'Target user does not exist.', 'httpStatus', 404); end if;
    if v_target.id = v_key.bound_user_id then return jsonb_build_object('error', 'License is already bound to this user.', 'httpStatus', 409); end if;
    select * into v_target_account from public.license_accounts where user_id = v_target.id;
    if v_target_account.banned_at is not null then return jsonb_build_object('error', 'Target account is banned.', 'httpStatus', 403); end if;
    if coalesce((p_payload->>'preserveDevice')::boolean, false) then
      v_new_device := null;
    else
      if length(coalesce(p_payload->>'installId', '')) < 16 then return jsonb_build_object('error', 'Destination installation ID is required.', 'httpStatus', 400); end if;
      insert into public.license_devices(user_id, install_id_hash, install_id_hint, label)
        values (v_target.id, 'manual_transfer:' || md5(p_payload->>'installId'), substring(trim(p_payload->>'installId') from 1 for 8) || '…' || substring(trim(p_payload->>'installId') from length(trim(p_payload->>'installId')) - 5 for 6), 'Assigned by transfer')
        returning * into v_new_device;
    end if;
    v_before := to_jsonb(v_key);
    update public.license_keys set bound_user_id = v_target.id, bound_email = lower(v_target.email),
      bound_device_id = coalesce(v_new_device.id, case when coalesce((p_payload->>'preserveDevice')::boolean, false) then bound_device_id else null end)
      where id = v_key.id;
    if v_target_account.user_id is null then
      insert into public.license_accounts(user_id, email, active_device_id, trial_started_at, trial_expires_at)
        values (v_target.id, lower(v_target.email), v_new_device.id, now(), now() + interval '30 days');
    elsif v_new_device.id is not null then
      update public.license_accounts set active_device_id = v_new_device.id where user_id = v_target.id;
    end if;
    insert into public.license_events(actor_user_id, target_user_id, license_id, device_id, event_type, metadata) values
      (p_actor_id, v_target.id, v_key.id, v_new_device.id, p_action, jsonb_build_object('before', v_before, 'after', (select to_jsonb(k) from public.license_keys k where k.id = v_key.id), 'note', v_note));
    return jsonb_build_object('success', true);
  end if;

  -- Client actions: status, heartbeat, activate
  if p_action not in ('status', 'heartbeat', 'activate') then
    return jsonb_build_object('error', 'Unknown action.', 'httpStatus', 400);
  end if;
  if length(coalesce(p_payload->>'installHash', '')) <> 64 then return jsonb_build_object('error', 'Installation ID is required.', 'httpStatus', 400); end if;

  select * into v_account from public.license_accounts where user_id = p_actor_id;
  select * into v_device from public.license_devices where install_id_hash = p_payload->>'installHash';

  -- ANTI-ABUSE: If hardware device is already registered to another account, INSTANT BLOCK
  if found and v_device.user_id <> p_actor_id and not exists (select 1 from public.license_admins where user_id = p_actor_id) then
    update public.license_devices set status = 'banned', ban_reason = 'Anti-abuse violation: Multiple accounts attempted on this hardware device.' where id = v_device.id;
    if v_account.user_id is not null then
      update public.license_accounts set banned_at = now(), ban_reason = 'Anti-abuse violation: Attempted to register multiple accounts on one hardware device.' where user_id = p_actor_id;
    end if;
    insert into public.license_events(actor_user_id, target_user_id, device_id, event_type, metadata)
      values (p_actor_id, p_actor_id, v_device.id, 'abuse_multi_account_blocked', jsonb_build_object('existing_user_id', v_device.user_id, 'installHint', p_payload->>'installHint'));
    return jsonb_build_object('error', 'Access blocked: This hardware device is already registered to another account. Multiple trial accounts per device are strictly prohibited. Contact zenithprojects@icloud.com.', 'code', 'DEVICE_BANNED', 'httpStatus', 403);
  end if;

  if v_device.id is null then
    if v_account.user_id is not null and v_account.active_device_id is not null
       and not exists (select 1 from public.license_admins where user_id = p_actor_id) then
      return jsonb_build_object('error', 'Your account is bound to another installation. Contact support for a device transfer.', 'code', 'DEVICE_MISMATCH', 'httpStatus', 409);
    end if;
    insert into public.license_devices(user_id, install_id_hash, install_id_hint)
      values (p_actor_id, p_payload->>'installHash', p_payload->>'installHint') returning * into v_device;
    if exists (select 1 from public.license_admins where user_id = p_actor_id) then
      update public.license_accounts set active_device_id = v_device.id where user_id = p_actor_id;
      v_account.active_device_id := v_device.id;
    end if;
  end if;

  -- First-time account creation
  if v_account.user_id is null then
    insert into public.license_accounts(user_id, email, active_device_id, trial_started_at, trial_expires_at)
      values (p_actor_id, lower(v_user.email), v_device.id, now(), now() + interval '30 days')
      returning * into v_account;
    insert into public.license_events(actor_user_id, target_user_id, device_id, event_type)
      values (p_actor_id, p_actor_id, v_device.id, 'trial_started');
  elsif v_account.active_device_id is null then
    update public.license_accounts set active_device_id = v_device.id where user_id = p_actor_id;
    v_account.active_device_id := v_device.id;
  end if;

  -- AUTO-CREATE AND BIND ACTIVATION KEY FOR FIRST-TIME USERS
  v_auto_key := p_payload->'autoKey';
  if v_auto_key is not null and not exists (select 1 from public.license_keys where bound_user_id = p_actor_id and status = 'active') then
    insert into public.license_keys(
      key_hash,
      key_hint,
      raw_key,
      status,
      duration_days,
      bound_user_id,
      bound_email,
      bound_device_id,
      activated_at,
      expires_at,
      created_by
    ) values (
      v_auto_key->>'keyHash',
      v_auto_key->>'keyHint',
      v_auto_key->>'rawKey',
      'active',
      30,
      p_actor_id,
      lower(v_user.email),
      v_device.id,
      now(),
      now() + interval '30 days',
      p_actor_id
    );
    insert into public.license_events(actor_user_id, target_user_id, device_id, event_type, metadata)
      values (p_actor_id, p_actor_id, v_device.id, 'auto_key_created', jsonb_build_object('key_hint', v_auto_key->>'keyHint'));
  end if;

  v_state := case when v_account.banned_at is not null then 'banned' when v_device.status = 'banned' then 'device_banned'
    when (v_device.status = 'retired' or v_account.active_device_id <> v_device.id)
         and not exists (select 1 from public.license_admins where user_id = p_actor_id) then 'device_mismatch' else null end;
  if p_action = 'activate' and v_state is not null then return jsonb_build_object('error', 'Account or device access is suspended or mismatched.', 'code', upper(v_state), 'httpStatus', 403); end if;
  update public.license_devices set last_seen_at = now() where id = v_device.id;
  update public.license_accounts set email = lower(v_user.email) where user_id = p_actor_id;

  if p_action = 'activate' then
    select * into v_key from public.license_keys where key_hash = p_payload->>'keyHash';
    if not found then return jsonb_build_object('error', 'Invalid activation key.', 'code', 'INVALID_KEY', 'httpStatus', 403); end if;
    if v_key.status = 'revoked' then return jsonb_build_object('error', 'This key has been revoked.', 'code', 'REVOKED', 'httpStatus', 403); end if;
    if v_key.status = 'active' and v_key.expires_at <= now() then return jsonb_build_object('error', 'This key has expired.', 'code', 'EXPIRED', 'httpStatus', 410); end if;
    if v_key.bound_user_id is not null and (v_key.bound_user_id <> p_actor_id or v_key.bound_email <> lower(v_user.email)) then return jsonb_build_object('error', 'This key is bound to another email.', 'code', 'EMAIL_MISMATCH', 'httpStatus', 403); end if;
    if v_key.bound_device_id is not null and v_key.bound_device_id <> v_device.id then return jsonb_build_object('error', 'This key is bound to another installation.', 'code', 'DEVICE_MISMATCH', 'httpStatus', 403); end if;
    if v_key.status = 'unused' then
      update public.license_keys set status = 'active', bound_user_id = p_actor_id, bound_email = lower(v_user.email),
        bound_device_id = v_device.id, activated_at = now(), expires_at = now() + duration_days * interval '1 day' where id = v_key.id;
      insert into public.license_events(actor_user_id, target_user_id, license_id, device_id, event_type) values (p_actor_id, p_actor_id, v_key.id, v_device.id, 'license_activated');
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

  return jsonb_build_object('state', v_state, 'canUse', v_state in ('trial', 'licensed'), 'isAdmin', v_role is not null,
    'email', lower(v_user.email), 'deviceHint', v_device.install_id_hint,
    'trialStartedAt', v_account.trial_started_at, 'trialExpiresAt', v_account.trial_expires_at,
    'daysRemaining', case when v_expiry is null then 0 else greatest(0, ceil(extract(epoch from (v_expiry - now())) / 86400)) end,
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
