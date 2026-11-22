# 0040. Make legacy Nano Server data access setup opt-in

**Status:** Accepted in 5.1.0

## Context

Private legacy data access components on x86 Nano Server may need metadata that is absent from a minimal image. Registration-free COM clients should not pay for or trigger those writes by default.

## Decision

Expose separate, explicit entry points for data access metadata alone or Jet 4.0 plus data access metadata. Gate them on x86 Nano Server with the compatibility forwarder present, leave existing values in place, and only write the 32-bit current-user registry. Ordinary activation does not use these entry points.

## Consequences

Setup is process-memoized and uses a trusted private provider directory. Run it at application startup before activating components that need it. It is not a general-purpose installer and does not remove the metadata on scope disposal.
