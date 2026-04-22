# Contributing Guide (Word-to-PDF)

Thanks for your interest in improving this project.

## Scope

This repository currently contains two conversion paths:

1. Node CLI (`bin/`, `lib/`)
2. Web converter (`converter_from_downloads.py`)

Please keep PRs focused and avoid unrelated refactors.

## Development Setup

### Node environment

```bash
npm install
npm test
```

### Web environment

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python3 converter_from_downloads.py
```

## Coding Guidelines

- Keep changes minimal and targeted.
- Preserve backward compatibility for existing commands where possible.
- Prefer explicit error messages over silent fallbacks.
- Keep conversion behavior deterministic for repeatability.

## Pull Request Checklist

- [ ] Problem statement is clear
- [ ] Change is scoped and documented
- [ ] `npm test` passes (or rationale provided)
- [ ] README/docs updated if behavior changed
- [ ] No generated artifacts accidentally included (PDF/temp/cache)

## Commit Message Suggestions

Use conventional-style prefixes when possible:

- `feat:` new user-visible feature
- `fix:` bug fix
- `docs:` documentation-only change
- `refactor:` code structure without behavior change
- `test:` tests only

## Reporting Issues

Please use the GitHub issue templates:
- Bug report
- Feature request

For security-sensitive reports, see `SECURITY.md`.
