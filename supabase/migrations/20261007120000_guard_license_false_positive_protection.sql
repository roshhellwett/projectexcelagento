-- Keep anti-abuse controls strict without permanently banning a legitimate user or shared IP
-- during normal status/heartbeat traffic. The older function remains the transaction engine; this
-- guarded entry point performs a serialized, non-destructive preflight before calling it.
begin;

create table if not exists public.license_owner_emails (
  email text primary key,
  created_at timestamptz not null default now()
);

insert into public.license_owner_emails(email)
values ('zenithopensourceprojects@gmail.com')
on conflict (email) do nothing;

alter table public.license_owner_emails enable row level security;
revoke all on public.license_owner_emails from public, anon, authenticated;
grant select, insert on public.license_owner_emails to service_role;

revoke all on function public.license_command(uuid, text, jsonb) from public, anon, authenticated;
grant execute on function public.license_command(uuid, text, jsonb) to service_role;

create or replace function public.license_command_guarded(
  p_actor_id uuid,
  p_action text,
  p_payload jsonb default '{}'::jsonb
)
returns jsonb
language plpgsql
security definer
set search_path = public, pg_temp
as $function$
declare
  v_role text;
  v_email text;
  v_device_user uuid;
  v_client_ip text;
  v_other_accounts integer;
begin
  perform pg_advisory_xact_lock(743602610);

  select lower(email) into v_email from auth.users where id = p_actor_id;
  select role into v_role from public.license_admins where user_id = p_actor_id;
  if exists (select 1 from public.license_owner_emails where email = v_email) then
    insert into public.license_admins(user_id, role)
    values (p_actor_id, 'owner')
    on conflict (user_id) do update set role = excluded.role;
    -- The configured owner must remain recoverable after an accidental anti-abuse false positive.
    update public.license_accounts
       set banned_at = null, ban_reason = null
     where user_id = p_actor_id;
    update public.license_devices
       set status = 'active', ban_reason = null
     where user_id = p_actor_id and status = 'banned';
    v_role := 'owner';
  end if;
  if p_action like 'admin\_%' escape '\' or v_role is not null then
    return public.license_command(p_actor_id, p_action, p_payload);
  end if;

  v_client_ip := coalesce(trim(p_payload->>'clientIp'), '');
  if v_client_ip <> '' then
    if exists (select 1 from public.license_banned_ips where ip = v_client_ip) then
      return jsonb_build_object(
        'error', 'Too many account attempts from this network. Contact support for review.',
        'code', 'IP_RATE_LIMITED',
        'httpStatus', 429
      );
    end if;

    select count(distinct actor_user_id)
      into v_other_accounts
      from public.license_ip_tracking
     where ip = v_client_ip
       and created_at > now() - interval '24 hours'
       and actor_user_id <> p_actor_id;

    -- Do not add another permanent IP ban from a normal verification request. A third distinct
    -- account in one day is rate-limited and can be reviewed by the administrator instead.
    if v_other_accounts >= 2 then
      return jsonb_build_object(
        'error', 'Multiple account creation from this network is temporarily limited. Contact support if this is a shared network.',
        'code', 'IP_RATE_LIMITED',
        'httpStatus', 429
      );
    end if;
  end if;

  select user_id into v_device_user
    from public.license_devices
   where install_id_hash = p_payload->>'installHash';

  -- A device already linked to another account is a recoverable ownership conflict. It must not
  -- ban the shared device or the account merely because a user signed out and another user signed in.
  if v_device_user is not null and v_device_user <> p_actor_id then
    return jsonb_build_object(
      'error', 'This installation is already linked to another account. Contact support for recovery.',
      'code', 'DEVICE_BOUND',
      'httpStatus', 409
    );
  end if;

  return public.license_command(p_actor_id, p_action, p_payload);
end;
$function$;

revoke all on function public.license_command_guarded(uuid, text, jsonb) from public, anon, authenticated;
grant execute on function public.license_command_guarded(uuid, text, jsonb) to service_role;

commit;
