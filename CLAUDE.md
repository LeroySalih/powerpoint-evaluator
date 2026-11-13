# powerpoint-evaluator — rules

Checks pupil PowerPoint (.pptx) work against a JSON requirements spec. Source of truth for design: `docs/DESIGN.md`.

## Git workflow
- **Commit to `main` and push `main`.** No branches, no pull requests (owner's standing instruction). Commit small; pull/rebase before pushing; run `pytest` first.

## Public repo, so:
- Never commit mark schemes, model answers, pupil files or pupil names/IDs. Those live in `private/` (git-ignored). Reference images live in `specs/<spec>/assets/` (git-ignored).
- Never log pupil names or file contents. Reports carry only what the caller sent.

## Rules
- Inspect the .pptx directly (zip + lxml); do not depend on PowerPoint. Never execute anything from an uploaded file.
- Every check returns `pass`, `fail` or `review` (needs a human/vision judge). Never guess a pass.
- Check by effective value (slide > layout > master), not by where the pupil happened to put it.
- Tests that need private material skip when it is absent.
