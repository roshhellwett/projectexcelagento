import { createClient } from 'npm:@supabase/supabase-js@2.117.2';
import { createLicenseHandler } from './handler.ts';

const url = Deno.env.get('SUPABASE_URL') ?? '';
const secret = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY') ?? '';
const db =
  url && secret
    ? createClient(url, secret, { auth: { persistSession: false, autoRefreshToken: false } })
    : null;

Deno.serve(
  createLicenseHandler({
    origin: Deno.env.get('APP_ORIGIN') || '*',
    keyPepper:
      Deno.env.get('LICENSE_KEY_PEPPER') ||
      'excelagento_default_license_key_pepper_2026_super_secure',
    devicePepper:
      Deno.env.get('LICENSE_DEVICE_PEPPER') ||
      'excelagento_default_device_pepper_2026_super_secure',
    async authenticate(token) {
      if (!db) return null;
      const { data, error } = await db.auth.getUser(token);
      return error ? null : (data.user?.id ?? null);
    },
    async command(actorId, action, payload) {
      if (!db) throw new Error('Service configuration missing');
      const { data, error } = await db.rpc('license_command', {
        p_actor_id: actorId,
        p_action: action,
        p_payload: payload,
      });
      if (error) throw new Error('Licensing transaction failed');
      return data;
    },
  }),
);
