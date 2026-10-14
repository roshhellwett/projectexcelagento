import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { basename, dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { inspectLicenseEntry } from './release-check-rules.mjs';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const errors = [];

// This check covers tracked private environment files, selected Vercel configuration, the web
// environment example, and direct license-function environment fallbacks. It is intentionally a
// limited static release-hygiene check rather than a complete secret scanner or deployment test.

function readJson(relativePath) {
  try {
    return JSON.parse(readFileSync(join(root, relativePath), 'utf8'));
  } catch {
    errors.push(`${relativePath} could not be parsed.`);
    return null;
  }
}

function readText(relativePath) {
  try {
    return readFileSync(join(root, relativePath), 'utf8');
  } catch {
    errors.push(`${relativePath} could not be read.`);
    return '';
  }
}

function trackedFiles() {
  try {
    return execFileSync('git', ['ls-files', '-z'], { cwd: root, encoding: 'utf8' })
      .split('\0')
      .filter(Boolean);
  } catch {
    errors.push('git metadata could not be read; release checks need a repository checkout.');
    return [];
  }
}

const tracked = trackedFiles();
const privateEnvironmentFiles = tracked.filter((file) => {
  const name = basename(file);
  return name === '.env' || (name.startsWith('.env.') && !name.endsWith('.env.example'));
});
if (privateEnvironmentFiles.length > 0) {
  errors.push(`private environment files are tracked: ${privateEnvironmentFiles.join(', ')}`);
}

const deploymentFiles = ['vercel.json', 'apps/web/vercel.json'];
const credentialPattern = /\b(?:sk-or-v1|sk-|gsk_|AIzaSy)[A-Za-z0-9_-]{8,}\b/u;
for (const relativePath of deploymentFiles) {
  const config = readJson(relativePath);
  if (!config) continue;

  if (config.framework !== 'vite') errors.push(`${relativePath} must declare the Vite framework.`);
  if (config.outputDirectory !== 'dist')
    errors.push(`${relativePath} must publish the dist directory.`);
  if (relativePath === 'vercel.json' && config.buildCommand !== 'pnpm build') {
    errors.push('vercel.json must build through the root pnpm build command.');
  }
  if (relativePath === 'apps/web/vercel.json' && config.buildCommand !== 'vite build') {
    errors.push('apps/web/vercel.json must build through vite build.');
  }

  const serialized = JSON.stringify(config);
  if (credentialPattern.test(serialized)) {
    errors.push(`${relativePath} contains a provider credential-shaped value.`);
  }
  if (serialized.includes('http://localhost:') || serialized.includes('http://127.0.0.1:')) {
    errors.push(`${relativePath} must not allow localhost connections in its production CSP.`);
  }
  for (const [name, value] of Object.entries(config.env ?? {})) {
    if (/key|token|secret|password/i.test(name) && typeof value === 'string' && value.trim()) {
      errors.push(`${relativePath} contains a non-empty secret-like environment value (${name}).`);
    }
  }
}

const envExamplePath = join(root, 'apps/web/.env.example');
try {
  const example = readFileSync(envExamplePath, 'utf8');
  for (const name of ['VITE_SUPABASE_URL', 'VITE_SUPABASE_ANON_KEY', 'VITE_AUTH_REDIRECT_ORIGIN']) {
    if (!new RegExp(`^${name}=`, 'm').test(example)) {
      errors.push(`apps/web/.env.example is missing ${name}.`);
    }
  }
} catch {
  errors.push('apps/web/.env.example is missing or unreadable.');
}

const licenseEntry = readText('supabase/functions/license/index.ts');
if (licenseEntry) {
  errors.push(...inspectLicenseEntry(licenseEntry));
}

if (errors.length > 0) {
  globalThis.console.error('Release hygiene checks failed (limited static scope):');
  for (const error of errors) globalThis.console.error(`- ${error}`);
  globalThis.process.exitCode = 1;
} else {
  globalThis.console.info(
    'Release hygiene checks passed: selected deployment and license-boundary checks passed (limited static scope).',
  );
}
