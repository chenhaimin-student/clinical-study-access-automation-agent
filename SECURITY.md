# Security Policy

## Supported code

Only the latest commit on `main` is considered for security fixes.

## Reporting

Use GitHub's private vulnerability reporting feature when available. If it is not available, open a minimal issue that contains no secret, personal data, customer information, private endpoint, or exploit payload, and ask for a private contact channel.

Do not submit real credentials, production email, workbooks, logs, databases or screenshots.

## Operational guidance

- start in preview mode;
- use isolated development accounts and fictional data;
- store credentials outside the repository;
- require explicit confirmation before mutations;
- rotate any credential that may have been exposed;
- verify each account and role result after an API operation;
- do not use this portfolio repository as a production deployment package.
