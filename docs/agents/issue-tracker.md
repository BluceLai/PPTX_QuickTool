# Issue Tracker: GitHub

Issues and PRDs for this repo live as GitHub issues in `BluceLai/PPTX_QuickTool`. Use the `gh` CLI for issue operations.

## Conventions

- Create an issue: `gh issue create --title "..." --body "..."`
- Read an issue: `gh issue view <number> --comments`
- List issues: `gh issue list --state open --json number,title,body,labels,comments`
- Comment on an issue: `gh issue comment <number> --body "..."`
- Apply or remove labels: `gh issue edit <number> --add-label "..."` / `--remove-label "..."`
- Close an issue: `gh issue close <number> --comment "..."`

The intended GitHub repository is:

```text
https://github.com/BluceLai/PPTX_QuickTool.git
```

When the local Git remote is configured, `gh` can infer the repo automatically from inside this clone.

## Pull Requests As A Triage Surface

PRs as a request surface: no.

Set this to `yes` if this repo later treats external pull requests as feature requests. When enabled, PRs should use the same triage states and labels as issues.

## When A Skill Says "Publish To The Issue Tracker"

Create a GitHub issue.

## When A Skill Says "Fetch The Relevant Ticket"

Run `gh issue view <number> --comments`.

## Wayfinding Operations

Used by the wayfinder workflow. The map is a single GitHub issue with child issues as tickets.

- Map: a single issue labelled `wayfinder:map`, holding notes, decisions so far, and open questions.
- Child ticket: an issue linked to the map as a GitHub sub-issue when available. If sub-issues are not available, add the child to a task list in the map body and put `Part of #<map>` at the top of the child body.
- Blocking: use GitHub native issue dependencies when available. If dependencies are not available, fall back to a `Blocked by: #<n>, #<n>` line at the top of the child body.
- Frontier query: list open child tickets, skip assigned tickets and tickets with open blockers, then take the first available ticket in map order.
- Claim: assign the ticket to the current developer.
- Resolve: comment with the answer or result, close the ticket, and append a context pointer to the map issue.
