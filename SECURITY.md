# Security rules for the Graph development demo

This public repository must never contain client secrets, passwords, bearer/access/refresh tokens, credentials, private account screenshots, downloaded Graph response JSON, or Microsoft 365 tenant-private data. The only supported Graph operation in the demo is an explicit, delegated, read-only `GET /v1.0/me?$select=id` with `User.Read`.

The local `.env.local` file is gitignored. The client/tenant IDs it holds are public-app configuration values, not secure secrets. Never place any actual secret behind Vite's `VITE_` prefix. Consent must be performed by the user or existing authorised tenant administrator, not automated by this demo.

This project is intended for localhost development and synthetic tests. Do not deploy it, expand permissions, merge to `main`, change an app registration, or use Microsoft Graph write calls without separate review and approval.
