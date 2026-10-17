# Project Rules — Pipecal Pro (`ryanyv/invoice_app`)

## Google Jules Delegation

Use **Google Jules** (`/opt/homebrew/bin/jules`) whenever suitable background or parallel coding tasks arise in this project.

- **Repository**: `ryanyv/invoice_app` (Git `origin` is `https://github.com/ryanyv/invoice_app.git`).
- **CLI Binary**: `/opt/homebrew/bin/jules`
- **When to Delegate to Jules**:
  - Self-contained background tasks such as writing or expanding unit/integration tests, isolated bug fixes, documentation updates, or bounded refactors where asynchronous remote execution helps.
  - Keep tightly coupled interactive changes or local desktop packaging (`electron-builder`) local.
- **Pre-Check**:
  - Before delegating, verify repository access with `/opt/homebrew/bin/jules remote list --repo` and confirm `ryanyv/invoice_app` is listed. If authentication fails (`401 UNAUTHENTICATED`), inform the user to run `jules login`.
- **Workflow**:
  1. Ensure the base branch/commits needed by Jules are pushed to `origin` (Jules operates on the remote GitHub repository, not uncommitted local files).
  2. Create a session:
     ```bash
     /opt/homebrew/bin/jules new --repo ryanyv/invoice_app "<clear task description, base branch, scope, constraints, and acceptance criteria>"
     ```
  3. Record the returned session ID and URL.
  4. Monitor status with `/opt/homebrew/bin/jules remote list --session` (avoid creating duplicate sessions for the same task).
  5. Inspect and pull completed work with `/opt/homebrew/bin/jules remote pull --session <id>` (and `--apply` after reviewing the diff and running `npm test` locally).
  6. Preserve existing local edits and verify all tests pass before committing, pushing, or publishing releases.
