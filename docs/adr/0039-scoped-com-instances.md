# 0039. Share COM instance lifetimes

**Status:** Accepted in 5.1.0

## Context

Several integrations need a transient or retained COM instance backed by an activation scope.

## Decision

Use `ScopedComInstance<T>` with caller-provided activation and instance factories. The activation scope remains responsible for the COM wrapper; each owner enforces thread ownership and reverse lease disposal order.

## Consequences

Callers retain control of classes, module verification, and error translation. Callers that nest scopes across different owners must respect the underlying activation scope order.
