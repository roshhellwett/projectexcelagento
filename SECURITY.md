# Security policy

## Scope

ExcelAgento processes workbooks in the browser and can send selected context to a model provider
when a user configures BYOK. The public deployment does not provide a shared API key. Agent memory
is 100% local in-browser with zero external database dependencies.

## Report a vulnerability

Please do not publish exploit details in a public issue or pull request. Open the repository's
[Security tab](https://github.com/roshhellwett/projectexcelagento/security) and use **Report a
vulnerability** when GitHub private vulnerability reporting is enabled. If that private form is
not available, contact the repository owner through [GitHub](https://github.com/roshhellwett) and
request a private channel before sharing details.

Include:

- affected version or commit;
- a minimal reproduction using synthetic data only;
- impact and prerequisites;
- any suggested mitigation.

Remove secrets and personal/customer data before sending a report. Use enough detail to support
investigation; disclosure coordination can be discussed after an issue is confirmed.

## Safe disclosure boundaries

Do not test against another person's workbook, API key, provider account, or the
production deployment without explicit authorization. Use local fixtures and synthetic keys.
