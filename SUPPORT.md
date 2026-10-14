# Support

ExcelAgento is an open-source project supported through the repository's public issue tracker.
Please search existing issues before opening a new one and choose the closest issue template.

## Setup and usage questions

For installation, local development, and product-boundary questions:

1. Read the [README](README.md), especially **Quick start**, **Quality gates**, and **Local recovery
   and formula safety**.
2. Check the package-level notes in [`packages/engine/README.md`](packages/engine/README.md),
   [`packages/agent/README.md`](packages/agent/README.md), and
   [`packages/evals/README.md`](packages/evals/README.md) when the question is about a specific
   package.
3. If the answer is still unclear, open a [support issue](https://github.com/roshhellwett/projectexcelagento/issues/new)
   with the setup steps, command output, and the smallest reproducible example.

## Bugs and feature ideas

- Use the repository's **Bug report** template for a reproducible problem.
- Use the **Feature request** template for a proposed user outcome or workflow.
- Include the commit or version, operating system and browser where relevant, and the exact
  commands you ran.

## Protect workbook and provider data

Use synthetic workbook data in issues. Never include API keys, customer workbooks, provider
payloads, private logs, access tokens, or other secrets. Redact any output that could identify a
person or organization.

## Security reports

Do not use a public issue for a vulnerability. Follow [`SECURITY.md`](SECURITY.md) for the private
reporting path and disclosure boundaries.
