# Planner CLI

A command-line client for Microsoft Planner. It talks to the Microsoft Graph API so people and agents can read and change Planner work from a shell, with compact output.

## Language

**Plan**:
A Microsoft Planner board owned by a Microsoft 365 group; it holds Buckets and Tasks.
_Avoid_: Board, project

**Bucket**:
A named column inside a Plan that groups Tasks.
_Avoid_: Column, list, lane

**Task**:
A unit of work in a Plan, placed in one Bucket, with a title, progress, priority, dates, and Assignees.
_Avoid_: Card, ticket, issue

**Details**:
The description and checklist of a Task, stored apart from the Task itself.
_Avoid_: Body, notes

**Assignee**:
A member of the Plan's group that a Task is given to.
_Avoid_: Owner, assigned user

**Priority**:
A Task's urgency, one of `urgent`, `important`, `medium`, `low`.
_Avoid_: Severity, importance

**Profile**:
A named set of sign-in and defaults (account, default Plan, default Bucket) selected with `-p`.
_Avoid_: Account, environment, context

**Default Plan**:
The Plan a command uses when none is given, set per Profile.
_Avoid_: Current plan, active plan

**Bulk create**:
Creating many Tasks in one `add` call from NDJSON or a JSON array.
_Avoid_: Import, batch
