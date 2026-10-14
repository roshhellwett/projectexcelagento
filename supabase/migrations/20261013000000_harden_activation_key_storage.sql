-- Never persist plaintext activation keys.
--
-- The Edge Function shows a generated batch once in its authenticated response. The database only
-- needs the HMAC key hash and a short hint for administration. Clear any values written by older
-- migrations and keep the legacy column as a nullable compatibility field so applying this change
-- does not require rewriting the long licensing transaction function in the same migration.
BEGIN;

UPDATE public.license_keys
SET raw_key = NULL
WHERE raw_key IS NOT NULL;

CREATE OR REPLACE FUNCTION public.clear_license_plaintext_key()
RETURNS trigger
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = public, pg_temp
AS $$
BEGIN
  NEW.raw_key := NULL;
  RETURN NEW;
END;
$$;

DROP TRIGGER IF EXISTS license_keys_clear_plaintext_key ON public.license_keys;
CREATE TRIGGER license_keys_clear_plaintext_key
  BEFORE INSERT OR UPDATE OF raw_key ON public.license_keys
  FOR EACH ROW
  EXECUTE FUNCTION public.clear_license_plaintext_key();

COMMENT ON COLUMN public.license_keys.raw_key IS
  'Deprecated compatibility column. Plaintext activation keys are cleared by trigger; use key_hash and key_hint.';

COMMIT;
