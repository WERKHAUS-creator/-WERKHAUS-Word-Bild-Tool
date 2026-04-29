# Branching and Release Policy

## Branch roles
- `develop`: Active development, testing, fixes, and iterative work.
- `main`: Stable live baseline only, validated and release-ready.

## Rules
- No feature or experiment commits directly on `main`.
- Every production change must flow: `develop` -> reviewed PR -> `main`.
- Azure deployment is allowed only from `main`.

## Pre-merge checks (`develop` -> `main`)
1. `git status` must be clean.
2. No unintended files (`node_modules`, logs, archives, temp files).
3. `npm run lint` passes.
4. `npm run build` passes.
5. `npm run validate` passes.
6. Manual smoke-check for Word insertion modes.
7. Change summary and risk note documented in PR.

## Pre-deploy checks (`main` -> Azure)
1. Confirm current branch is `main`.
2. Confirm target commit is intended release.
3. Ensure no local-only URLs in production manifest.
4. Verify GitHub Secrets contain deployment token only (no plain tokens in repo).
5. Trigger workflow and validate deployment URL.

## Security rules
- Never commit secrets, tokens, certificates, or private keys.
- Keep deployment token in `AZURE_STATIC_WEB_APPS_API_TOKEN` only.
- Keep local development configuration separate from production configuration.
