-- Migration: 20261009000000_simplify_licensing_remove_devices.sql
-- Simplifies activation & trial management:
-- 1. All email sign-ups automatically receive a 30-day trial from registration.
-- 2. Once the 30-day trial expires, user pays for an activation key (30 days, 60 days, etc.) to continue.
-- 3. Completely removes complex device identity, machine IDs, HWID, and device lockouts.
-- 4. Keys are bound strictly to user accounts without device constraints.

BEGIN;

-- 1. Update handle_new_user trigger function to guarantee 30-day trial creation on sign-up
CREATE OR REPLACE FUNCTION public.handle_new_user()
RETURNS trigger
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = public, pg_temp
AS $$
BEGIN
  -- Populate public.profiles
  INSERT INTO public.profiles (id, email, full_name, avatar_url, tier, subscription_status)
  VALUES (
    new.id,
    new.email,
    coalesce(new.raw_user_meta_data->>'full_name', split_part(new.email, '@', 1)),
    new.raw_user_meta_data->>'avatar_url',
    'free',
    'active'
  )
  ON CONFLICT (id) DO UPDATE SET
    email = excluded.email,
    updated_at = now();

  -- Populate public.license_accounts with 30-day trial
  INSERT INTO public.license_accounts (user_id, email, trial_started_at, trial_expires_at)
  VALUES (
    new.id,
    lower(new.email),
    now(),
    now() + interval '30 days'
  )
  ON CONFLICT (user_id) DO UPDATE SET
    email = lower(excluded.email);

  RETURN new;
END;
$$;

-- 2. Make device ID columns nullable and optional on license_accounts and license_keys
ALTER TABLE public.license_accounts ALTER COLUMN active_device_id DROP NOT NULL;
ALTER TABLE public.license_keys ALTER COLUMN bound_device_id DROP NOT NULL;
ALTER TABLE public.license_keys DROP CONSTRAINT IF EXISTS license_binding_complete;
ALTER TABLE public.license_keys ADD CONSTRAINT license_binding_complete
  CHECK (((status <> 'active'::text) OR ((bound_user_id IS NOT NULL) AND (bound_email IS NOT NULL) AND (activated_at IS NOT NULL) AND (expires_at IS NOT NULL))));

-- 3. Replace public.license_command with simple, clean account-based logic
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
  v_key public.license_keys%rowtype;
  v_state text;
  v_expiry timestamptz;
  v_base_time timestamptz;
  v_new_expiry timestamptz;
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
  v_now_expiry timestamptz;
begin
  select * into v_user from auth.users where id = p_actor_id;
  if not found then
    return jsonb_build_object('error', 'Sign in again to verify access.', 'code', 'AUTH_REQUIRED', 'httpStatus', 401);
  end if;

  -- Ensure account record exists with 30-day trial from user registration
  select * into v_account from public.license_accounts where user_id = p_actor_id;
  if v_account.user_id is null then
    insert into public.license_accounts(user_id, email, trial_started_at, trial_expires_at)
      values (
        p_actor_id,
        lower(v_user.email),
        coalesce(v_user.created_at, now()),
        coalesce(v_user.created_at, now()) + interval '30 days'
      )
      on conflict (user_id) do update set email = lower(excluded.email)
      returning * into v_account;
  end if;

  v_client_ip := coalesce(p_payload->>'clientIp', '');

  -- Determine if actor is Admin or Owner
  select role into v_role from public.license_admins where user_id = p_actor_id;
  if v_role is null and exists (select 1 from public.license_owner_emails where email = lower(v_user.email)) then
    v_role := 'owner';
  end if;
  v_is_owner := (v_role = 'owner');
  v_is_admin := (v_role is not null);

  -- IP tracking if provided
  if v_client_ip <> '' then
    insert into public.license_ip_tracking (ip, actor_user_id, action)
      values (v_client_ip, p_actor_id, p_action);
  end if;

  -- Rate limits
  v_bucket := case when p_action = 'activate' then 'activation' when p_action like 'admin%' then 'admin' else 'verification' end;
  v_limit := case v_bucket when 'activation' then 20 when 'admin' then 200 else 300 end;
  v_window := case v_bucket when 'activation' then interval '10 minutes' else interval '1 minute' end;
  insert into public.license_rate_limits values (p_actor_id, v_bucket, now(), 1)
    on conflict (actor_user_id, bucket) do update set
      attempts = case when license_rate_limits.window_started_at + v_window <= now() then 1 else license_rate_limits.attempts + 1 end,
      window_started_at = case when license_rate_limits.window_started_at + v_window <= now() then now() else license_rate_limits.window_started_at end
    returning attempts into v_attempts;
  if v_attempts > v_limit then
    return jsonb_build_object('error', 'Too many attempts. Please retry later.', 'code', 'RATE_LIMITED', 'httpStatus', 429);
  end if;

  -- 1. ADMIN LIST (Overview, users, keys, support)
  if p_action = 'admin_list' then
    if not v_is_admin then
      return jsonb_build_object('error', 'Access restricted to administrators.', 'code', 'FORBIDDEN', 'httpStatus', 403);
    end if;

    v_offset := greatest(0, least(100000, coalesce((p_payload->>'offset')::integer, 0)));
    v_query := lower(coalesce(p_payload->>'query', ''));
    return jsonb_build_object(
      'keys', coalesce((select jsonb_agg(t) from (select id, key_hint, raw_key,
        case when status = 'active' and expires_at <= now() then 'expired' else status end as status,
        duration_days, bound_user_id, bound_email, activated_at, expires_at, revoked_at, created_at,
        case when status = 'unused' then duration_days
             when expires_at is null then duration_days
             else greatest(0, ceil(extract(epoch from (expires_at - now())) / 86400)) end as days_remaining
        from public.license_keys
        where v_query = '' or position(v_query in lower(coalesce(bound_email, '') || key_hint || coalesce(raw_key, '') || status)) > 0
        order by created_at desc, id limit 200 offset v_offset) t), '[]'::jsonb),
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
        order by (case when exists (select 1 from public.license_owner_emails loe where loe.email = lower(u.email)) then 1 else 2 end), u.created_at desc limit 200 offset v_offset) t), '[]'::jsonb),
      'tickets', coalesce((select jsonb_agg(t) from (select * from public.support_tickets
        where v_query = '' or position(v_query in lower(name || ' ' || email || ' ' || subject || ' ' || message || ' ' || status)) > 0
        order by created_at desc limit 100 offset v_offset) t), '[]'::jsonb),
      'devices', '[]'::jsonb,
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

  -- 2. ADMIN GENERATE KEYS (30 days, 60 days, etc.)
  if p_action = 'admin_generate' then
    if not v_is_admin then
      return jsonb_build_object('error', 'Access restricted to administrators.', 'code', 'FORBIDDEN', 'httpStatus', 403);
    end if;

    v_days := (p_payload->>'durationDays')::integer;
    if v_days is null or v_days not between 1 and 3650 or jsonb_typeof(p_payload->'keys') <> 'array'
      or jsonb_array_length(p_payload->'keys') not between 1 and 100 then
      return jsonb_build_object('error', 'Specify 1–100 keys and duration between 1 and 3650 days.', 'httpStatus', 400);
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

  -- 5. ADMIN ADJUST TRIAL / EXTEND DAYS
  if p_action = 'admin_adjust_trial' then
    if not v_is_admin then
      return jsonb_build_object('error', 'Access restricted to administrators.', 'code', 'FORBIDDEN', 'httpStatus', 403);
    end if;

    select * into v_target from auth.users where id = (p_payload->>'userId')::uuid;
    if not found then return jsonb_build_object('error', 'User not found.', 'httpStatus', 404); end if;

    if exists (select 1 from public.license_admins where user_id = v_target.id) or
       exists (select 1 from public.license_owner_emails where email = lower(v_target.email)) then
      return jsonb_build_object('error', 'Administrator and Owner accounts have permanent lifetime access.', 'httpStatus', 400);
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

  -- 6. ADMIN BAN / UNBAN USER
  if p_action = 'admin_ban_user' then
    if not v_is_admin then
      return jsonb_build_object('error', 'Access restricted to administrators.', 'code', 'FORBIDDEN', 'httpStatus', 403);
    end if;

    select * into v_target from auth.users where id = (p_payload->>'userId')::uuid;
    if not found then return jsonb_build_object('error', 'User not found.', 'httpStatus', 404); end if;

    if exists (select 1 from public.license_admins where user_id = v_target.id) or
       exists (select 1 from public.license_owner_emails where email = lower(v_target.email)) then
      return jsonb_build_object('error', 'Administrator and Owner accounts cannot be banned.', 'httpStatus', 400);
    end if;

    update public.license_accounts set banned_at = now(), ban_reason = coalesce(p_payload->>'reason', 'Banned by admin')
      where user_id = v_target.id;
    insert into public.license_events(actor_user_id, target_user_id, event_type, metadata) values
      (p_actor_id, v_target.id, p_action, jsonb_build_object('reason', p_payload->>'reason'));
    return jsonb_build_object('success', true);
  end if;

  if p_action = 'admin_unban_user' then
    if not v_is_admin then
      return jsonb_build_object('error', 'Access restricted to administrators.', 'code', 'FORBIDDEN', 'httpStatus', 403);
    end if;

    update public.license_accounts set banned_at = null, ban_reason = null where user_id = (p_payload->>'userId')::uuid;
    insert into public.license_events(actor_user_id, target_user_id, event_type) values
      (p_actor_id, (p_payload->>'userId')::uuid, p_action);
    return jsonb_build_object('success', true);
  end if;

  -- 7. ADMIN SUPPORT TICKETS
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

  -- 8. USER KEY ACTIVATION (Simple paid key activation for 30d, 60d, etc.)
  if p_action = 'activate' then
    select * into v_key from public.license_keys
      where key_hash = p_payload->>'keyHash' or (p_payload->>'rawKey' is not null and raw_key = p_payload->>'rawKey')
      for update;

    if not found then
      return jsonb_build_object('error', 'Activation key was not recognized.', 'code', 'KEY_NOT_FOUND', 'httpStatus', 404);
    end if;

    if v_key.status = 'revoked' then
      return jsonb_build_object('error', 'This activation key was revoked.', 'code', 'KEY_REVOKED', 'httpStatus', 409);
    end if;

    if v_key.status = 'active' and v_key.bound_user_id is not null and v_key.bound_user_id <> p_actor_id then
      return jsonb_build_object('error', 'This activation key has already been used by another account.', 'code', 'KEY_BOUND', 'httpStatus', 409);
    end if;

    if v_key.status = 'unused' then
      -- Add duration_days to current active access or start from now if expired
      v_base_time := greatest(now(), coalesce(v_account.trial_expires_at, now()));
      v_new_expiry := v_base_time + (v_key.duration_days * interval '1 day');

      update public.license_keys
        set status = 'active',
            bound_user_id = p_actor_id,
            bound_email = lower(v_user.email),
            activated_at = now(),
            expires_at = v_new_expiry
        where id = v_key.id
        returning * into v_key;

      update public.license_accounts
        set trial_expires_at = v_new_expiry
        where user_id = p_actor_id
        returning * into v_account;

      insert into public.license_events(actor_user_id, target_user_id, license_id, event_type, metadata)
        values (p_actor_id, p_actor_id, v_key.id, 'license_activated',
          jsonb_build_object('durationDays', v_key.duration_days, 'newExpiry', v_new_expiry));
    end if;
  end if;

  -- Check ban status
  if not v_is_admin and v_account.banned_at is not null then
    v_state := 'banned';
  end if;

  -- Determine user state and expiration
  if v_state is null then
    select * into v_key from public.license_keys
      where bound_user_id = p_actor_id and status = 'active' and expires_at > now()
      order by expires_at desc limit 1;

    if v_key.id is not null then
      v_state := 'licensed';
      v_expiry := greatest(v_key.expires_at, coalesce(v_account.trial_expires_at, v_key.expires_at));
    elsif v_account.trial_expires_at > now() then
      v_state := 'trial';
      v_expiry := v_account.trial_expires_at;
    else
      v_state := 'expired';
      v_expiry := v_account.trial_expires_at;
    end if;
  end if;

  -- Owners and Admins have unrestricted lifetime access
  if v_is_admin then
    v_state := 'licensed';
    v_expiry := null;
  end if;

  return jsonb_build_object(
    'state', v_state,
    'canUse', (v_is_admin or v_state in ('trial', 'licensed')),
    'isAdmin', v_is_admin,
    'isOwner', v_is_owner,
    'role', v_role,
    'isLifetime', v_is_admin,
    'email', lower(v_user.email),
    'trialStartedAt', v_account.trial_started_at,
    'trialExpiresAt', case when v_is_admin then null else v_account.trial_expires_at end,
    'daysRemaining', case when v_is_admin then null when v_expiry is null then 0 else greatest(0, ceil(extract(epoch from (v_expiry - now())) / 86400)) end,
    'banReason', case when v_is_admin then null else v_account.ban_reason end,
    'license', case when v_key.id is null then null else jsonb_build_object(
      'id', v_key.id,
      'keyHint', v_key.key_hint,
      'rawKey', v_key.raw_key,
      'status', case when v_key.status = 'active' and v_key.expires_at <= now() then 'expired' else v_key.status end,
      'activatedAt', v_key.activated_at,
      'expiresAt', v_key.expires_at,
      'daysRemaining', greatest(0, ceil(extract(epoch from (v_key.expires_at - now())) / 86400))
    ) end,
    'serverNow', now()
  );
end;
$function$;

-- 4. Replace public.license_command_guarded to remove all device checks
CREATE OR REPLACE FUNCTION public.license_command_guarded(
  p_actor_id uuid,
  p_action text,
  p_payload jsonb DEFAULT '{}'::jsonb
)
RETURNS jsonb
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = public, pg_temp
AS $function$
DECLARE
  v_role text;
  v_email text;
BEGIN
  -- Re-assert owner privileges if registered owner email
  SELECT lower(email) INTO v_email FROM auth.users WHERE id = p_actor_id;
  IF EXISTS (SELECT 1 FROM public.license_owner_emails WHERE email = v_email) THEN
    INSERT INTO public.license_admins(user_id, role)
    VALUES (p_actor_id, 'owner')
    ON CONFLICT (user_id) DO UPDATE SET role = excluded.role;
    UPDATE public.license_accounts
       SET banned_at = null, ban_reason = null
     WHERE user_id = p_actor_id;
  END IF;

  RETURN public.license_command(p_actor_id, p_action, p_payload);
END;
$function$;

-- 5. Set proper execution permissions
REVOKE ALL ON FUNCTION public.license_command(uuid, text, jsonb) FROM PUBLIC, anon, authenticated;
GRANT EXECUTE ON FUNCTION public.license_command(uuid, text, jsonb) TO service_role;
REVOKE ALL ON FUNCTION public.license_command_guarded(uuid, text, jsonb) FROM PUBLIC, anon, authenticated;
GRANT EXECUTE ON FUNCTION public.license_command_guarded(uuid, text, jsonb) TO service_role;

COMMIT;
