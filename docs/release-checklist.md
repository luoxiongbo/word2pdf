# Open-Source Release Checklist

Use this before pushing to GitHub public repository.

## Repository Metadata

- [ ] `package.json`:
  - [ ] `author` updated
  - [ ] `repository.url` updated
  - [ ] `description` and `keywords` validated
- [ ] `LICENSE` copyright owner updated

## Documentation

- [ ] `README.md` reflects actual behavior
- [ ] `docs/operations.md` commands are executable
- [ ] `docs/architecture.md` matches current code
- [ ] `SECURITY.md` has valid private reporting path

## Assets

- [ ] Add Web UI screenshot to `docs/images/web-ui-screenshot.png`
- [ ] Remove sensitive sample files
- [ ] Confirm no private paths/usernames leaked in docs

## Code Quality

- [ ] `npm test` passes
- [ ] Manual conversion sanity check completed
- [ ] No accidental debug prints or temporary hacks

## Packaging Hygiene

- [ ] `.gitignore` covers generated files
- [ ] No large generated artifacts committed (`*.pdf`, temp files)
- [ ] Dependency versions reviewed

## Final Sanity

- [ ] Fresh clone bootstrap tested
- [ ] README Quick Start works end-to-end
- [ ] Issue/PR templates visible in GitHub UI
