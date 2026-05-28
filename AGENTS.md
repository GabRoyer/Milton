# Agent Guidelines

## Effort Workflow

An effort is a scoped body of work that may span one or more phases and pull requests.

- Put each effort on its own branch with a name relevant to the effort.
- Keep each phase on its own branch.
- Stack phase branches on the previous phase branch, not directly on `main`, unless the phase is intentionally independent.
- Use branch names that make the stack obvious, for example `effort/phase-1`, `effort/phase-2`, or `effort/dev-profiles-phase-2.1`.

## Design Docs

Design docs should include a phased implementation plan before implementation starts.

- Each phase should roughly map to one pull request.
- A phase can be split into dot phases when the work is too large or risky, for example `Phase 2.1`, `Phase 2.2`, and `Phase 2.3`.
- Plans should describe the intended branch name, PR scope, dependencies on earlier phases, and validation approach.
- If implementation reveals that a phase is larger than planned, split it before continuing rather than letting the PR grow unchecked.

## PR Size

Keep pull requests small enough to review carefully.

- Aim for under 500 changed lines per PR, excluding tests.
- Imported libraries, lockfile churn from dependency installation, and generated artifacts do not count toward the 500-line budget.
- Tests can exceed the budget when they are necessary to cover the change.
- Before opening a PR, inspect the diff and decide whether the implementation should be split into another phase or dot phase.
- Use judgment: a slightly larger cohesive PR is better than an artificial split that obscures behavior, but large mixed-scope PRs should be avoided.

## Stacked PRs

For multi-phase efforts, open PRs as a stack.

- Phase 1 targets the effort's base branch, usually `main`.
- Phase 2 targets the Phase 1 branch.
- Phase 3 targets the Phase 2 branch, and so on.
- Push each phase branch and open a PR for that phase before starting the next phase when practical.
- Keep later phase branches rebased or merged as needed when earlier phases change.

## PR Descriptions

Every phased PR description should start with navigation links.

Use this block at the top:

```md
Previous PR: <link or none>
Next PR: <link or pending>
```

- Update the previous PR when the next PR exists.
- Update the next PR when a following PR is opened.
- Use `none` only for the first PR in a stack.
- Use `pending` when the adjacent PR is planned but not opened yet.

## Implementation Discipline

- Keep implementation PRs aligned with the relevant design-doc phase.
- Do not combine unrelated phases just because they touch nearby files.
- Include focused validation in each phase.
- Keep generated local development files out of source control.
- If a phase needs follow-up work, record it in the design doc or PR description before moving on.

## Code Quality

- Avoid barrel import files where practical. Prefer importing from the module that owns the symbol so dependencies stay explicit and refactors remain local.
- Add concise JSDoc headers to TypeScript and JavaScript functions, interfaces, and interface/class fields. Headers should describe purpose or contract, not restate the symbol name; generated code and copied external declarations are exempt.
