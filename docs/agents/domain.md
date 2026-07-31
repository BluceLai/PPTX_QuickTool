# Domain Docs

How the engineering skills should consume this repo's domain documentation when exploring the codebase.

## Before Exploring, Read These

- `CONTEXT.md` at the repo root, or
- `CONTEXT-MAP.md` at the repo root if it exists. It points at one `CONTEXT.md` per context. Read each one relevant to the topic.
- `docs/adr/`. Read ADRs that touch the area you are about to work in. In multi-context repos, also check `src/<context>/docs/adr/` for context-scoped decisions.

If any of these files do not exist, proceed silently. Do not flag their absence or suggest creating them upfront. The domain-modeling workflow creates them lazily when terms or decisions actually get resolved.

## File Structure

Single-context repo, which is the layout for this project:

```text
/
|-- CONTEXT.md
|-- docs/
|   `-- adr/
|       |-- 0001-example-decision.md
|       `-- 0002-example-decision.md
`-- src/
```

Multi-context repo, only if `CONTEXT-MAP.md` exists at the root:

```text
/
|-- CONTEXT-MAP.md
|-- docs/
|   `-- adr/
`-- src/
    |-- context-a/
    |   |-- CONTEXT.md
    |   `-- docs/
    |       `-- adr/
    `-- context-b/
        |-- CONTEXT.md
        `-- docs/
            `-- adr/
```

## Use The Glossary's Vocabulary

When output names a domain concept in an issue title, refactor proposal, hypothesis, or test name, use the term as defined in `CONTEXT.md`. Do not drift to synonyms the glossary explicitly avoids.

If the concept is not in the glossary yet, either reconsider the language or note the gap for domain modeling.

## Flag ADR Conflicts

If output contradicts an existing ADR, surface it explicitly rather than silently overriding it.
