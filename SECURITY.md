# Security and Client Operations

## Required setup

Copy `.env.example` into the deployment environment and set three separate secrets:

- `AILYN_LOGIN_PASSWORD` protects the workspace.
- `AILYN_ADMIN_PASSWORD` protects restore and upgrade actions.
- `AILYN_UPDATE_SIGNING_KEY` authorizes signed Python upgrades.

Keep the data directory outside the repository with `AILYN_DATA_DIR` when possible. The app stores the SQLite database, Excel exports, photos, receipt archives, and ZIP backups there. Backups should be copied to a second physical or cloud location because a local backup cannot protect against disk loss.

## Included protections

- Authentication is required by default and repeated failures are throttled.
- Restore uploads reject path traversal, missing application databases, oversized archives, and ZIP bombs.
- Image and update uploads have size limits.
- User-controlled report text is HTML-escaped.
- SQLite and backup files use owner-only permissions where supported.
- Signed upgrades require an HMAC signature and create a complete backup first.

## Recommended client features

1. Add an append-only audit log for login, create, edit, delete, restore, export, and budget actions.
2. Add role-based accounts so an accountant can export data without being able to restore or upgrade the system.
3. Add duplicate and anomaly review using the existing searchable-record helpers before approving monthly reports.
4. Add scheduled off-site encrypted backups with retention rules and a visible last-successful-backup status.
5. Add a monthly close workflow that locks an approved period and records who approved it.
6. Add attachment checksums and a document download manifest for receipts and work-proof photos.