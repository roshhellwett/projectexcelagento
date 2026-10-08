-- Harden the public boundaries added by the licensing/support migrations.
-- Client input must never choose another account's ticket owner, and the privileged
-- license command must only be callable by the service-role Edge Function.

BEGIN;

-- The Edge Function authenticates the bearer token and calls this control plane through its
-- service-role client. Re-granting the underlying SECURITY DEFINER function to browser roles would
-- allow callers to submit an arbitrary p_actor_id.
REVOKE ALL ON FUNCTION public.license_command(uuid, text, jsonb) FROM PUBLIC, anon, authenticated;
GRANT EXECUTE ON FUNCTION public.license_command(uuid, text, jsonb) TO service_role;
REVOKE ALL ON FUNCTION public.license_command_guarded(uuid, text, jsonb) FROM PUBLIC, anon, authenticated;
GRANT EXECUTE ON FUNCTION public.license_command_guarded(uuid, text, jsonb) TO service_role;

-- Direct inserts remain a compatibility fallback for older clients, but an authenticated caller
-- can only attach a ticket to their own account. Anonymous tickets must have a null owner.
DROP POLICY IF EXISTS "Anyone can submit support tickets" ON public.support_tickets;
CREATE POLICY "Anyone can submit support tickets"
  ON public.support_tickets FOR INSERT
  WITH CHECK (user_id IS NULL OR user_id = auth.uid());

-- The RPC derives ownership from the current JWT. p_user_id is retained in the signature for
-- client compatibility but is intentionally ignored; browser input cannot impersonate another
-- account or make a ticket visible in somebody else's support history.
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
  v_user_id uuid := auth.uid();
  v_name text := trim(coalesce(p_name, ''));
  v_email text := lower(trim(coalesce(p_email, '')));
  v_category text := trim(coalesce(p_category, 'general'));
  v_subject text := trim(coalesce(p_subject, ''));
  v_message text := trim(coalesce(p_message, ''));
BEGIN
  IF length(v_name) NOT BETWEEN 1 AND 200
     OR v_email !~* '^[^[:space:]@]+@[^[:space:]@]+\.[^[:space:]@]+$'
     OR length(v_email) > 320
     OR length(v_category) NOT BETWEEN 1 AND 50
     OR length(v_subject) NOT BETWEEN 1 AND 200
     OR length(v_message) NOT BETWEEN 10 AND 10000 THEN
    RAISE EXCEPTION 'Please provide a valid name, email, subject, and a message between 10 and 10000 characters.'
      USING ERRCODE = '22023';
  END IF;

  INSERT INTO public.support_tickets (name, email, category, subject, message, user_id, status)
  VALUES (v_name, v_email, v_category, v_subject, v_message, v_user_id, 'pending')
  RETURNING id INTO v_ticket_id;

  RETURN jsonb_build_object('success', true, 'id', v_ticket_id);
END;
$$;

GRANT EXECUTE ON FUNCTION public.submit_support_ticket(text, text, text, text, text, uuid)
  TO anon, authenticated, service_role;

COMMIT;
