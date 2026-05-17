---
name: vba-code-generation
description: "Generate, refactor, and document Excel VBA code following 5-lines-code principle, semantic versioning, and best practices"
---

# Excel VBA Code Generation Skill

Use this skill when generating, refactoring, or extending VBA code in Excel macro-enabled workbooks (.xlsm, .xlsb, .xlam).

> **Note**: This skill requires macro-enabled workbooks. For non-macro-enabled files (.xlsx, .xls), convert to .xlsm first or ensure the VBE is accessible. If VBE is inaccessible, check security settings: File > Options > Trust Center > Trust Center Settings > Macro Settings. If VBE remains inaccessible despite adjusting security settings, provide a detailed error message and consult IT support or check for Group Policy restrictions.

## Core Principles

Decision rule for conflicts: if the 5-lines-code rule conflicts with safe error handling or required behavior, keep safe error handling and required behavior first, then refactor into helper procedures to restore the 5-lines-code rule.


### 1. 5 Lines Code Rule

Each function/procedure should target 5 lines of executable code, excluding declarations; when unavoidable for safe error handling or required behavior, exceed temporarily and then refactor into helper procedures.

**Definition of executable code**: Counts all code statements that perform actions (assignments, method calls, control flow, inline error-handling statements, and each statement in multi-statement one-liners separated by :). Excludes: comments, blank lines, variable/constant declarations, and function/procedure signatures.

Benefit: Improved readability, testability, and maintainability.


### 2. Error Handling

Always implement error handling for functions interacting with WScript.Shell, file I/O, databases, external APIs, or other external dependencies.


### 3. Semantic Versioning Integration

When working with git branch names matching `release/v*` or `hotfix/v*`, extract semantic version strings using regex.

### 4. Git Integration

For commands needing git output (branch name, status, etc.), use WScript.Shell to execute and capture output.

**Error Handling for Git**: Always check if git is installed and configured. If git is not available, provide a clear message indicating "Git is not installed or configured" and suggest installing git or checking the system PATH configuration.

### 5. Variable Declaration Standards

Follow these standards for variable declarations and type usage:

- **One variable per Dim statement**: Each `Dim` statement declares exactly one variable. Avoid comma-separated declarations.
  ```vba
  ' Good
  Dim vbMod As VBIDE.CodeModule
  Dim lineNum As Long

  ' Avoid
  Dim vbMod As Object, lineNum As Long
  ```

- **Avoid Object type**: Use specific type declarations instead of `Object`. Prefer concrete types like `VBIDE.CodeModule`, `Worksheet`, `Range`, etc.
  ```vba
  ' Good
  Dim vbMod As VBIDE.CodeModule

  ' Avoid
  Dim vbMod As Object
  ```

- **Declare variables just before use**: Place `Dim` statements immediately before the variable is first used, improving readability and reducing scope.
  ```vba
  ' Good
  Set vbMod = ThisWorkbook.VBProject.VBComponents("TestController").CodeModule
  Dim lineNum As Long
  For lineNum = 1 To vbMod.CountOfLines
      ' ...
  Next lineNum

  ' Avoid
  Dim vbMod As Object
  Dim lineNum As Long
  ' ... other code ...
  Set vbMod = ThisWorkbook.VBProject.VBComponents("TestController").CodeModule
  ```

## Workflow

### VBA Code Editing

1. **Open workbook in Excel** (macro-enabled: .xlsm, .xlsb, .xlam, .xls)
2. **Access VBE**: Press `Alt+F11`
3. **Edit modules only in VBE**, never directly in `.xls?.VBA/` folder
4. **Save workbook**: Use VBA Immediate Window `Excel.Application.DisplayAlerts = False: ThisWorkbook.Save: Excel.Application.DisplayAlerts = True`
5. **Stage changes**: `git add .`
6. **Run pre-commit**: `uv run pre-commit` (this may update generated files)
7. **Re-stage only if pre-commit changed files**: `git add .`
8. **Review staged diff**: `git diff --cached -- <workbook>.xls?`
9. **Commit and push**: Use Conventional Commits style

### Testing Patterns

Use VBA Immediate Window (Ctrl+G in VBE) to validate:
- String parsing and regex matching
- Git command output capture
- Property/version retrieval

## Checklist for New VBA Code

- [ ] All procedures follow 5-lines-code rule
- [ ] Error handlers in place for external calls
- [ ] Functions properly scoped (Public/Private)
- [ ] Each `Dim` declares one variable only
- [ ] Specific types used (no `Object` type)
- [ ] Variables declared just before use
- [ ] Regex patterns validated and escaped correctly
- [ ] Pre-commit extraction verified: `uv run pre-commit`
- [ ] Staged diff reviewed: `git diff --cached`
- [ ] Commit message follows Conventional Commits style

## References

- [5 Lines Code Principle](https://dev.to/kanani_nirav/the-five-lines-of-code-principle-why-less-is-more-in-programming-31j6)
- [Semantic Versioning](https://semver.org/spec/v2.0.0.html)
- [Use RegExp classes included in VBE](https://devblogs.microsoft.com/microsoft365dev/how-to-prepare-vba-projects-for-vbscript-deprecation/#use-regexp-classes-included-in-vbe)
