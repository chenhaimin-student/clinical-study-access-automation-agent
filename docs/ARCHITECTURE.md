# Architecture

The project separates message intake, validation, approval, external-system adapters and result verification.

1. `email_intake.py` parses mail headers, bodies and attachments, applies scope and privacy policy, and creates a job record.
2. Workbook parsing converts site and user rows into normalized operation records.
3. Preview mode validates every row without changing external state.
4. Explicit approval enables the Study and Role API adapters.
5. Each adapter records per-item outcomes and performs a read-after-write verification.
6. The result layer writes an auditable summary and prepares recipient-aware notifications.

The original deployment used Hermes as the Agent gateway. The public repository keeps that dependency optional and isolates the author-written patch in `integrations/hermes`.

The repository does not include a production database. Job storage is file-based and should point to an ignored `.runtime` directory for local evaluation.
