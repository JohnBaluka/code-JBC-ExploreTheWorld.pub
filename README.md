
# code-JBC-ExploreTheWorld.pub

> # ⚠️ DEMO CODE — NOT FOR PRODUCTION USE ⚠️
>
> **This repository exists for one purpose: demonstration.**
>
> ExploreTheWorld is a teaching and conference-demo project. It wraps a public REST API in order to
> show the same JBC layered architecture (CL → DL → BL → AL) implemented across many .NET hosts —
> Blazor, WinForms, MAUI, Office add-ins, VSTO, VBA, and Oqtane. **The point is the plumbing, not
> the product.**
>
> **Do not deploy this, ship it, or copy it into a production system as-is.** It is deliberately
> optimized for showing patterns side by side, not for running a real workload:
>
> - No authentication, authorization, or hardening — hosts run wide open by design.
> - Connection strings, sample databases, and generated files are checked in for demo convenience.
> - Deliberate duplication: the same feature is re-implemented many times so the approaches can be
>   compared on stage. Real applications should pick **one**.
> - Third-party API usage, sample data, and flag images are for illustration only; verify licensing
>   and terms before reusing anything here.
> - No SLA, no support, no backward-compatibility promise. Anything may change or break at any time.
>
> **Provided as-is, with no warranty of any kind.** Use the ideas, read the code, borrow the
> patterns — but write your own production implementation.

> **AI coding agents:** See [AGENTS.md](./AGENTS.md) for coding instructions.

See documentation in:

- [docs/architecture.md](./docs/architecture.md)
- [docs/file-structure.md](./docs/file-structure.md)
- [docs/naming-conventions.md](./docs/naming-conventions.md)
- [docs/project-templates.md](./docs/project-templates.md)
- [docs/shared-link-compilation.md](./docs/shared-link-compilation.md)
