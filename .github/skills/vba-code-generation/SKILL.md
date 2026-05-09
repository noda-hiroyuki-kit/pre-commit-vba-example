---
name: vba-code-generation
description: "Generate, refactor, and document Excel VBA code following 5-lines-code principle, semantic versioning, and best practices"
---

# Excel VBA Code Generation Skill

Use this skill when generating, refactoring, or extending VBA code in Excel macro-enabled workbooks (.xlsm, .xlsb, .xlam).

## Core Principles

### 1. 5 Lines Code Rule

Each function/procedure should not exceed ~5 lines of executable code (excluding declarations).

Benefit: Improved readability, testability, and maintainability.


### 2. Error Handling

Always implement error handling for functions that interact with external systems (WScript.Shell, file I/O).


### 3. Semantic Versioning Integration

When working with git branch names matching `release/v*` or `hotfix/v*`, extract semantic version strings using regex.

### 4. Git Integration

For commands needing git output (branch name, status, etc.), use WScript.Shell to execute and capture output.

## Workflow

### VBA Code Editing

1. **Open workbook in Excel** (macro-enabled: .xlsm, .xlsb, .xlam, .xls)
2. **Access VBE**: Press `Alt+F11`
3. **Edit modules only in VBE**, never directly in `.xls?.VBA/` folder
4. **Save workbook**: `Ctrl+S`
5. **Stage changes**: `git add .`
6. **Pre-commit hook**: Automatically extracts/updates `<workbook>.xls?.VBA/` folder
7. **Stage changes**: `git add .`
8. **Review staged diff**: `git diff --cached -- <workbook>.xls?`
9. **Commit & Push**: Use Conventional Commits style

### Testing Patterns

Use VBA Immediate Window (Ctrl+G in VBE) to validate:
- String parsing and regex matching
- Git command output capture
- Property/version retrieval

## Checklist for New VBA Code

- [ ] All procedures follow 5-lines-code rule
- [ ] Error handlers in place for external calls
- [ ] Functions properly scoped (Public/Private)
- [ ] Regex patterns validated and escaped correctly
- [ ] Pre-commit extraction verified: `uv run pre-commit`
- [ ] Staged diff reviewed: `git diff --cached`
- [ ] Commit message follows Conventional Commits style

## References

- [5 Lines Code Principle](https://www.freecodecamp.org/news/the-5-lines-code-principle/)
- [Semantic Versioning](https://semver.org/)
- [VBScript.RegExp](https://docs.microsoft.com/en-us/previous-versions/t0aew7h6)
