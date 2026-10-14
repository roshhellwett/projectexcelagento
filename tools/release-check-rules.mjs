function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function nonEmptyFallbackLiterals(source, environmentName) {
  const pattern = new RegExp(
    `Deno\\s*\\.\\s*env\\s*\\.\\s*get\\s*\\(\\s*['"]${escapeRegExp(environmentName)}['"]\\s*\\)\\s*(?:\\|\\||\\?\\?)\\s*(['"])([^'"]*)\\1`,
    'gu',
  );

  return Array.from(source.matchAll(pattern), (match) => match[2]).filter(
    (fallback) => fallback.length > 0,
  );
}

/**
 * Check only the direct, quoted environment fallbacks in the license Edge Function entry point.
 * This is a static hygiene check, not a general secret scanner or a deployment validation.
 */
export function inspectLicenseEntry(source) {
  const errors = [];
  const pepperFallbacks = nonEmptyFallbackLiterals(source, 'LICENSE_KEY_PEPPER');
  if (pepperFallbacks.length > 0) {
    errors.push('license Edge Function must not fall back to a checked-in key pepper.');
  }

  const originFallbacks = nonEmptyFallbackLiterals(source, 'APP_ORIGIN');
  if (originFallbacks.includes('*')) {
    errors.push('license Edge Function must not default to a wildcard application origin.');
  } else if (originFallbacks.length > 0) {
    errors.push('license Edge Function must not use a non-empty APP_ORIGIN fallback literal.');
  }

  return errors;
}
