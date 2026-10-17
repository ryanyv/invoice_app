# Project Rules — Pipecal Pro (`ryanyv/invoice_app`)

## Continuous Versioning, Desktop Builds & Release Policy

Every change or feature made to this project must follow this strict release cycle:

1. **Bump Version Number**:
   - Increment the semantic version number for every functional change or feature.
   - Update `package.json` (`"version"`).
   - Update `src/components/Navbar.tsx` (the Persian version badge, e.g., `نسخه ۲.۲`).
   - Update test assertions in `src/services/__tests__/challengerM6_2Adversarial.test.ts` (suite 2.1 checks `pkg.version`).

2. **Generate Desktop Builds**:
   - Build macOS DMG packages: `npm run electron:build:mac`
   - Build Windows installer & portable packages: `npm run electron:build:win`
   - Maintain dot-named release copies (`Pipecal.Pro-<version>-arm64.dmg`, `Pipecal.Pro.Setup.<version>.exe`, `Pipecal.Pro.<version>.exe`) and calculate their SHA-256 checksums.

3. **Preserve All Previous Versions**:
   - **Never delete or overwrite older releases**.
   - All past versions (e.g., `v2.0.0`, `v2.1.0`, `v2.2.0`, etc.) must remain accessible and downloadable both in the local `release/` directory and on GitHub Releases.

4. **Update README.md & Publish GitHub Release**:
   - Update `README.md` (both English and Persian sections):
     - New version badge and primary direct download links.
     - Summary of newly added features and enhancements.
     - New SHA-256 checksum table.
     - Retain archive links to all previous releases.
   - Publish a new GitHub Release via `gh release create v<version>` uploading the binaries with bilingual release notes (English & Persian) and SHA-256 checksums.

5. **Test & Git Workflow**:
   - Run `npm test` to verify all automated test suites pass.
   - Stage modified files explicitly by path (never blanket `git add .`).
   - Follow Conventional Commits and push to remote (`git push origin main`).

---

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
