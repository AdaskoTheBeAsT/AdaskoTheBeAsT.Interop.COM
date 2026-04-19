# ADR-0001: Record Architecture Decisions

- Status: Accepted
- Date: 2026-04-19
- Deciders: @AdaskoTheBeAsT
- Release: v3.0.0

## Context

The project has accumulated a number of non-trivial, long-lived architectural choices (registration-free COM strategy, the Result pattern, the static `Executor` API, the multi-framework target matrix, the analyzer policy, and the v3.0.0 redesign around `IDisposable` and `IComExecutor`). Those decisions were, until now, only visible via the commit history, the README, and the code itself, which makes them hard to rediscover and easy to accidentally contradict.

## Decision

We adopt lightweight [Markdown Architectural Decision Records (MADR)](https://adr.github.io/madr/) stored under `docs/adr/`. Every architecturally significant decision is captured as a numbered ADR. Historical decisions are recreated retrospectively based on the commit chronology so newcomers can reason about the codebase without replaying the git history.

## Consequences

- Positive: the reasoning behind key choices is discoverable alongside the source code.
- Positive: reviewers can point at ADRs when refusing or accepting changes that would break a documented contract.
- Negative: every substantive change that touches the public surface or cross-cutting concerns now requires an ADR entry, which adds a small authoring cost.
- Follow-up: when a decision is superseded, update its status and cross-link the new ADR.

## Alternatives Considered

- Keeping the design narrative in the README: rejected because the README is a user-facing document, not a design journal, and it already struggles to stay focused on adoption.
- Maintaining a single `DECISIONS.md`: rejected because it collapses all history into one file and makes it hard to deprecate individual entries.

## References

- [Michael Nygard, "Documenting Architecture Decisions"](https://cognitect.com/blog/2011/11/15/documenting-architecture-decisions)
- [adr.github.io](https://adr.github.io/)
