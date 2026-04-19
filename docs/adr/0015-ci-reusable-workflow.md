# ADR-0015: Adopt a Shared Reusable GitHub Actions CI Workflow

- Status: Accepted
- Date: 2026-04-19
- Deciders: @AdaskoTheBeAsT
- Release: v3.0.0
- Related commit(s): v3.0.0 changes (`.github/workflows/ci.yml`)

## Context

Before v3.0.0 the repository had an analyzer stack (ADR-0006) and tests, but no automated CI. Pull requests could drift from green for weeks because nobody was automatically building against all 10 target frameworks, running tests, or publishing a SonarQube report. The author maintains several sibling packages (`AdaskoTheBeAsT.Interop.Threading`, `AdaskoTheBeAsT.*.Analyzer`, …) that all need the same build-test-Sonar-publish pipeline; copying and maintaining a full `ci.yml` per repository is duplicative and hard to keep in sync.

## Decision

The repository uses a single CI workflow at `.github/workflows/ci.yml` that delegates to the shared reusable workflow at `AdaskoTheBeAsT/github-actions/.github/workflows/dotnet-build-sonarqube-nuget.yml@v1`. Per-repository settings (`.NET versions`, runner OS, solution name, Sonar organization / project key, whether to honour `global.json`) are passed as inputs from repository variables, and secrets (Sonar token, NuGet API key) are passed through. The workflow runs on pull requests, pushes to `main`, and version tags (`v*`) so tags can automatically trigger NuGet publication.

Permissions are intentionally minimal: `checks: write`, `contents: read`, `issues: write`, `pull-requests: read`.

## Consequences

- Positive: every push and PR compiles the full 10-TFM matrix, runs tests, and publishes Sonar analysis results.
- Positive: changes to the pipeline (e.g. a new .NET version, a security fix in an action) are made once in the shared repo and every consumer inherits them.
- Positive: the local `ci.yml` stays short (a dozen lines) and its intent is obvious.
- Negative: the shared workflow introduces a runtime dependency on a private/external GitHub Actions repository; it must be versioned (`@v1`) and kept stable.
- Negative: opinionated pipeline choices (matrix strategy, reporting, caching) live outside this repo; contributors must look at the shared workflow to understand what runs.

## Alternatives Considered

- Inline `ci.yml` in every repo: rejected because it fragments the pipeline and makes hot-fixes expensive.
- Composite actions only: rejected because composite actions cannot supply entire top-level `jobs:` blocks; they would still require a wrapper per repo.
- No CI, rely on local builds: rejected as incompatible with ADR-0006 (warnings-as-errors only helps if it actually runs on every change).

## References

- `.github/workflows/ci.yml`
- [GitHub Actions reusable workflows](https://docs.github.com/en/actions/using-workflows/reusing-workflows)
