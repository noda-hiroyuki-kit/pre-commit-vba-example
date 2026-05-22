---
name: pr-review-thread-workflow
description: "Handle PR review threads sequentially with mandatory rationale comments, commit/push checks, and thread resolution rules"
---

# PR Review Thread Workflow (Repository-specific)

Use this workflow when the user asks to process Copilot/GitHub PR review comments for this repository.

## Environment

- **OS**: Windows
- **Shell**: PowerShell
- **VCS**: git
- **Tools**: MCP GitHub tools, VS Code

## Default Execution Mode

Use these defaults unless the user explicitly overrides them:

- Process review comments one by one in numeric/user-specified order.
- User performs VBA edits and staging; agent performs verification, commit, push, and comment drafting.
- Once user says "fixed and staged", proceed to commit and push without asking again.
- After each push, immediately provide the review reply text that includes acceptance/decline and commit ID.

## Goals

- Process review threads one by one in the user-specified order.
- Keep user control over source edits when requested.
- Ensure every resolved thread has a traceable rationale comment.
- Avoid accidental code changes outside requested scope.

## VBA Code Editing Workflow

**Important**: User edits are made only in Excel VBE, not directly in repository files.

1. Open the macro-enabled workbook (`.xlsm`, `.xlsb`, `.xlam`, `.xls`, or similar)
2. Press Alt+F11 or right-click -> Edit -> Trigger VBE
3. Edit VBA code in the VBE module browser
4. Save the workbook (Ctrl+S)
5. Pre-commit hook automatically extracts and updates `<workbook>.xls?.VBA/` folder (for example, `example-app.xlsm.VBA/`)
6. `<workbook>.xls?.VBA/` folder contents are auto-generated; **do not edit directly**
7. Stage the updated `.xls?` file and corresponding `<workbook>.xls?.VBA/` changes
8. Agent verifies staged diff and commits with Conventional Commits message
9. Agent pushes branch and drafts review reply comment

## Standard Operating Procedure

1. Refresh active PR data and collect unresolved threads.
2. Group threads by topic and identify duplicates.
3. For each thread, propose one of: implement, decline, or defer.
4. Wait for user decision when policy-sensitive or ambiguous.
5. If implementation is selected:
   - **User edits VBA code in Excel VBE** (not directly in `.xls?.VBA` folder)
   - Save the macro-enabled workbook
   - Verify staged diff: `git diff --cached --stat`
   - Confirm exact changes: `git diff --cached -- <workbook>.xls?` (for example, `git diff --cached -- example-app.xlsm`)
   - Commit with Conventional Commits style
   - Push branch
6. Before resolving thread(s), post a PR comment that includes:
   - target thread IDs
   - implemented commit hash(es) or explicit decline rationale
7. Resolve only the threads covered by the posted rationale.
8. Re-fetch unresolved threads and repeat until zero.

## One-By-One Operational Flow

For each thread, follow this exact sequence:

1. Identify target thread ID and requested action (implement/decline).
2. Wait for user signal that edits are done and staged.
3. Verify staged files and staged diff summary.
4. Commit using Conventional Commits.
5. Push current branch to origin.
6. Return English review comment in Markdown code block.
7. Move to next thread.

## Commit Message Convention (Repository Rule)

- Format: `<type>(<scope>): <subject>`
- `<scope>` must be module name in lowercase kebab-case.
- Examples:
  - `fix(trailing-whitespace-module): add workbook cleanup error handling`
  - `refactor(trailing-whitespace-module): centralize unchanged-code guard in replacement`
  - `fix(branch-version-resolver): surface branch resolution failures`
- Keep one thread topic per commit.

## Review Reply Templates (English)

### Accepted template

```markdown
Accepted.
Addressed in commit <COMMIT_ID>.
<One concise sentence describing the implemented change and why it resolves the concern.>
```

### Declined template

```markdown
Declined.
No code changes were made.

<One concise sentence explaining the intentional design choice and why current behavior is acceptable for this repository.>
```

## Decision Rules for This Repository

- If a thread category was already decided in prior rounds, keep consistency unless user explicitly changes policy.
- Treat repeated duplicate comments as a batch: one rationale comment can cover multiple thread IDs.
- If a tool cannot reply inside the specific thread, post PR-level comment and explicitly list thread IDs.

## Expected Agent Responsibilities

- Never edit `<workbook>.xls?.VBA/` directly.
- Never ask the user to repeat standing instructions already encoded in this skill.
- If staged changes exist for the current thread, proceed with commit/push autonomously.
- If commit/push succeeds, always report commit hash and provide ready-to-post reply text.

## Commit and Verification Checklist

- Verify macro-enabled workbook was edited in Excel VBE, not directly in `.xls?.VBA` folder
- Confirm staged files: `git diff --cached --stat`
- Confirm exact changes: `git diff --cached -- <workbook>.xls?` (for example, `git diff --cached -- example-app.xlsm`)
- Use a clear Conventional Commit message
- Push immediately after commit
- Keep commits narrowly scoped to the thread topic

## Communication Rules

- Always present current unresolved set before action.
- Always state what was closed, what was changed, and why.
- Never resolve a thread silently.
- If no unresolved threads remain, report count = 0 explicitly.

## Progress Reporting Template

- New unresolved threads: <count>
- Closed this step: <thread ID>
- Commit created: <hash or none>
- Push status: <done or failed>
- Remaining unresolved: <count>

## Safety Rules

- Do not revert unrelated local changes.
- Do not use destructive git commands.
- If unexpected workspace changes appear, stop and ask user.
- **Do not edit files directly in `<workbook>.xls?.VBA/` folders** - these are auto-generated by pre-commit hook.

## Code Search Example (Windows PowerShell)

For querying VBA code across a project workbook:

```powershell
Get-ChildItem -Recurse "<workbook>.xls?.VBA\code" -File |
  Where-Object { $_.Extension -in '.bas','.cls','.frm' } |
  Select-String -Pattern '<regex_pattern>' |
  ForEach-Object { "{0}:{1}: {2}" -f $_.Path, $_.LineNumber, $_.Line.Trim() }
```

Example: `example-app.xlsm.VBA\code` folder
