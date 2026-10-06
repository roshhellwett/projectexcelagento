# Security policy

## Scope

ExcelAgento processes workbooks in the browser and can send selected context to a model provider
when a user configures BYOK. The public deployment does not provide a shared API key. Agent memory
is 100% local in-browser with zero external database dependencies.

## Report a vulnerability

Please do not publish exploit details in a public issue. Contact the project maintainers through
the security contact listed by the repository owner, or use the private security-reporting channel
provided by the hosting organization. Include:

- affected version or commit;
- a minimal reproduction using synthetic data only;
- impact and prerequisites;
- any suggested mitigation.

Remove secrets and personal/customer data before sending a report. We will acknowledge reports when
we can, investigate with a reproducible case, and coordinate disclosure for confirmed issues.

## Safe disclosure boundaries

Do not test against another person's workbook, API key, provider account, or the
production deployment without explicit authorization. Use local fixtures and synthetic keys.
