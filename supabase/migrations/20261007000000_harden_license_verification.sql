-- Upgrade an already-provisioned database; do not rerun the initial migration.
begin;

-- An active entitlement may await support-approved rebinding. Ownership, email
-- and original expiry remain mandatory; a null device never grants access.
alter table public.license_keys drop constraint license_binding_complete;
alter table public.license_keys add constraint license_binding_complete check (
  status <> 'active' or (bound_user_id is not null and bound_email is not null
    and activated_at is not null and expires_at is not null)
);
create index if not exists license_accounts_created_idx on public.license_accounts(created_at desc, user_id);
create index if not exists license_keys_created_idx on public.license_keys(created_at desc, id);
create index if not exists license_devices_created_idx on public.license_devices(first_seen_at desc, id);

-- Keep the proven admin mutations while replacing the high-volume verification
-- path. The existing admin transaction takes the EXCLUSIVE advisory lock; user
-- verification takes the corresponding SHARED lock, so unrelated users run in
-- parallel and an admin revoke/ban/transfer cannot race a verification commit.
alter function public.license_command(uuid, text, jsonb) rename to license_command_v1;
revoke all on function public.license_command_v1(uuid, text, jsonb) from public, anon, authenticated, service_role;

create function public.license_command(p_actor_id uuid, p_action text, p_payload jsonb default '{}')
returns jsonb language plpgsql security definer set search_path = public, pg_temp as $$
declare
  v_user auth.users%rowtype;
  v_account public.license_accounts%rowtype;
  v_device public.license_devices%rowtype;
  v_key public.license_keys%rowtype;
  v_role text;
  v_result jsonb;
  v_state text;
  v_expiry timestamptz;
  v_attempts integer;
  v_bucket text;
  v_window interval;
  v_limit integer;
begin
  select * into v_user from auth.users where id = p_actor_id;
  if not found or v_user.email is null then
    return jsonb_build_object('error', 'Sign in again.', 'code', 'AUTH_REQUIRED', 'httpStatus', 401);
  end if;
  if p_action like 'admin\_%' escape '\' then
    select role into v_role from public.license_admins where user_id = p_actor_id;
    if v_user.email_confirmed_at is null or v_role is null
      or exists(select 1 from public.license_accounts where user_id = p_actor_id and banned_at is not null) then
      return jsonb_build_object('error', 'A verified administrator account is required.', 'code', 'ADMIN_REQUIRED', 'httpStatus', 403);
    end if;
    if p_action = 'admin_access' then
      return jsonb_build_object('isAdmin', true, 'adminRole', v_role);
    end if;
    -- The legacy transaction returns domain errors as JSON. A subtransaction
    -- ensures every rejected admin command rolls back intermediate writes (for
    -- example a destination device created before a transfer conflict).
    begin
      v_result := public.license_command_v1(p_actor_id, p_action, p_payload);
      if v_result ? 'error' then raise exception using errcode = 'PZ001', message = 'Rejected admin command'; end if;
    exception when sqlstate 'PZ001' then
      return v_result;
    end;
    return v_result;
  end if;
  if p_action not in ('status', 'heartbeat', 'activate') then
    return jsonb_build_object('error', 'Unknown licensing action.', 'httpStatus', 400);
  end if;
  if length(coalesce(p_payload->>'installHash', '')) <> 64 then
    return jsonb_build_object('error', 'Installation ID is required.', 'httpStatus', 400);
  end if;
  if v_user.email_confirmed_at is null then
    return jsonb_build_object('state', 'email_unconfirmed', 'canUse', false, 'isAdmin', false,
      'email', lower(v_user.email), 'deviceHint', null, 'trialStartedAt', null, 'trialExpiresAt', null,
      'license', null, 'daysRemaining', 0, 'banReason', null, 'serverNow', now(), 'verifyUntil', now());
  end if;
  perform pg_advisory_xact_lock_shared(743602610);
  -- Actor-scoped serialization covers initial enrollment and attempt counters;
  -- a key row lock below covers concurrent redemption by DIFFERENT accounts.
  perform pg_advisory_xact_lock(hashtextextended('excelagento:user:' || p_actor_id::text, 0));
  select role into v_role from public.license_admins where user_id = p_actor_id;
  v_bucket := case when p_action = 'activate' then 'activation' else 'verification' end;
  v_limit := case when p_action = 'activate' then 10 else 120 end;
  v_window := case when p_action = 'activate' then interval '10 minutes' else interval '1 minute' end;
  insert into public.license_rate_limits values (p_actor_id, v_bucket, now(), 1)
    on conflict (actor_user_id, bucket) do update set
      attempts = case when license_rate_limits.window_started_at + v_window <= now() then 1 else least(license_rate_limits.attempts + 1, v_limit + 1) end,
      window_started_at = case when license_rate_limits.window_started_at + v_window <= now() then now() else license_rate_limits.window_started_at end
    returning attempts into v_attempts;
  if v_attempts > v_limit then
    return jsonb_build_object('error', 'Too many attempts. Please retry later.', 'code', 'RATE_LIMITED', 'httpStatus', 429);
  end if;
  select * into v_account from public.license_accounts where user_id = p_actor_id;
  select * into v_device from public.license_devices where install_id_hash = p_payload->>'installHash';
  if v_device.id is null then
    if v_account.active_device_id is not null and v_role is null then
      return jsonb_build_object('error', 'Contact support to transfer this account to your new installation.', 'code', 'DEVICE_MISMATCH', 'httpStatus', 409);
    end if;
    insert into public.license_devices(user_id, install_id_hash, install_id_hint)
      values (p_actor_id, p_payload->>'installHash', p_payload->>'installHint')
      on conflict (install_id_hash) do nothing;
    select * into v_device from public.license_devices where install_id_hash = p_payload->>'installHash';
  end if;
  if v_device.user_id <> p_actor_id then
    return jsonb_build_object('error', 'This installation belongs to another account. Contact support for recovery.', 'code', 'DEVICE_BOUND', 'httpStatus', 409);
  end if;
  if v_account.user_id is null then
    insert into public.license_accounts(user_id, email, active_device_id)
      values (p_actor_id, lower(v_user.email), v_device.id) returning * into v_account;
    insert into public.license_events(actor_user_id, target_user_id, device_id, event_type)
      values (p_actor_id, p_actor_id, v_device.id, 'trial_started');
  elsif v_account.active_device_id is null and v_device.status = 'active' then
    update public.license_accounts set active_device_id = v_device.id where user_id = p_actor_id;
    v_account.active_device_id := v_device.id;
  end if;
  v_state := case when v_account.banned_at is not null then 'banned'
    when v_device.status = 'banned' then 'device_banned'
    when v_device.status = 'retired' then 'device_mismatch'
    when v_account.active_device_id is distinct from v_device.id and v_role is null then 'device_mismatch' else null end;
  if p_action = 'activate' and v_state is not null then
    return jsonb_build_object('error', 'Account or device access is suspended or mismatched.', 'code', upper(v_state), 'httpStatus', 403);
  end if;
  -- Avoid unnecessary row versions at every heartbeat/multi-tab refresh.
  update public.license_devices set last_seen_at = now() where id = v_device.id and last_seen_at < now() - interval '5 minutes';
  update public.license_accounts set email = lower(v_user.email) where user_id = p_actor_id and email is distinct from lower(v_user.email);
  if p_action = 'activate' then
    select * into v_key from public.license_keys where key_hash = p_payload->>'keyHash' for update;
    if not found then return jsonb_build_object('error', 'Invalid activation key.', 'code', 'INVALID_KEY', 'httpStatus', 403); end if;
    if v_key.status = 'revoked' then return jsonb_build_object('error', 'This key has been revoked.', 'code', 'REVOKED', 'httpStatus', 403); end if;
    if v_key.status = 'active' and v_key.expires_at <= now() then return jsonb_build_object('error', 'This key has expired.', 'code', 'EXPIRED', 'httpStatus', 410); end if;
    if v_key.bound_user_id is not null and (v_key.bound_user_id <> p_actor_id or v_key.bound_email <> lower(v_user.email)) then
      return jsonb_build_object('error', 'This key is bound to another email.', 'code', 'EMAIL_MISMATCH', 'httpStatus', 403);
    end if;
    if v_key.bound_device_id is not null and v_key.bound_device_id <> v_device.id then
      return jsonb_build_object('error', 'This key is bound to another installation.', 'code', 'DEVICE_MISMATCH', 'httpStatus', 403);
    end if;
    if v_key.status = 'unused' then
      update public.license_keys set status = 'active', bound_user_id = p_actor_id, bound_email = lower(v_user.email),
        bound_device_id = v_device.id, activated_at = now(), expires_at = now() + duration_days * interval '1 day' where id = v_key.id;
      insert into public.license_events(actor_user_id, target_user_id, license_id, device_id, event_type)
        values (p_actor_id, p_actor_id, v_key.id, v_device.id, 'license_activated');
    elsif v_key.bound_device_id is null then
      update public.license_keys set bound_device_id = v_device.id where id = v_key.id;
      insert into public.license_events(actor_user_id, target_user_id, license_id, device_id, event_type)
        values (p_actor_id, p_actor_id, v_key.id, v_device.id, 'license_rebound');
    end if;
  end if;
  select * into v_key from public.license_keys where bound_user_id = p_actor_id and bound_device_id = v_device.id
    order by (status = 'active' and expires_at > now()) desc, expires_at desc nulls last, id limit 1;
  if v_state is null then
    v_state := case when v_key.status = 'active' and v_key.expires_at > now() and v_key.bound_email <> lower(v_user.email) then 'email_mismatch'
      when v_key.status = 'active' and v_key.expires_at > now() then 'licensed'
      when exists(select 1 from public.license_keys where bound_user_id = p_actor_id and status = 'active' and bound_device_id is null and expires_at > now()) then 'activation_required'
      when v_key.status = 'revoked' then 'revoked'
      when v_account.trial_expires_at > now() then 'trial' else 'expired' end;
  end if;
  v_expiry := case when v_state = 'licensed' then v_key.expires_at when v_state = 'trial' then v_account.trial_expires_at else null end;
  return jsonb_build_object('state', v_state, 'canUse', v_state in ('trial', 'licensed'),
    'isAdmin', v_role is not null and v_account.banned_at is null, 'adminRole', v_role,
    'email', lower(v_user.email), 'deviceHint', v_device.install_id_hint,
    'trialStartedAt', v_account.trial_started_at, 'trialExpiresAt', v_account.trial_expires_at,
    'daysRemaining', case when v_expiry is null then 0 else greatest(0, ceil(extract(epoch from (v_expiry - now())) / 86400)) end,
    'banReason', coalesce(v_account.ban_reason, v_device.ban_reason),
    'license', case when v_key.id is null then null else jsonb_build_object('id', v_key.id, 'keyHint', v_key.key_hint,
      'status', case when v_key.status = 'active' and v_key.expires_at <= now() then 'expired' else v_key.status end,
      'activatedAt', v_key.activated_at, 'expiresAt', v_key.expires_at,
      'daysRemaining', greatest(0, ceil(extract(epoch from (v_key.expires_at - now())) / 86400))) end,
    'serverNow', now(), 'verifyUntil', least(coalesce(v_expiry, now()), now() + interval '90 seconds'));
end;
$$;
revoke all on function public.license_command(uuid, text, jsonb) from public, anon, authenticated;
grant execute on function public.license_command(uuid, text, jsonb) to service_role;
commit;
