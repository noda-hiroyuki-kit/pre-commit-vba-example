# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.1.3] - 2026-09-25

### Added

- Add `zizmor` and `pinact` for GitHub Actions security checks and version pinning.  
  GitHub Actions のセキュリティチェックとバージョン固定のために `zizmor` と `pinact` を追加.

### Changed

- Update the pre-commit-vba hook to v0.4.3.  
    pre-commit-vba フックを v0.4.3 に更新.
- Replace `pre-commit` with `prek` and use its built-in hooks.  
  `pre-commit` を `prek` に置き換え、組み込みフックを使用.
- Update the automated hook update workflow to use `prek` and pinned GitHub Actions versions.  
  自動フック更新ワークフローを `prek` とバージョン固定済みの GitHub Actions に更新.

## [0.1.2] - 2026-09-13

### Changed

- Update the pre-commit-vba hook to v0.4.1.  
    pre-commit-vba フックを v0.4.1 に更新.
- Update development dependency lock versions.  
    開発依存関係のロックバージョンを更新.

## [0.1.1] - 2026-06-22

### Changed

- Refactor BranchVersionResolver by extracting `resolveVersionText` and helper subs for setting values and raising errors.  
    BranchVersionResolver で `resolveVersionText` と、値設定・エラー送出の補助サブルーチンを切り出すようにリファクタリング.
- Align RegisterProduct naming between form modules.  
    フォーム関連モジュール間の RegisterProduct 命名を統一.

### Fixed

- Add explicit type to label parameter in register form validation helper.  
    登録フォームのバリデーションヘルパーで label パラメータの型を明示.
- Add Rubberduck ignore annotation for unused test mock procedure warning.  
    テスト用モックの未使用警告に対して Rubberduck の ignore 注釈を追加.

## [0.1.0] - 2026-06-21

### Added

- Release as first version.  
    最初のバージョンとしてリリース.

[Unreleased]: https://github.com/noda-hiroyuki-kit/pre-commit-vba-example/compare/v0.1.3...HEAD
[0.1.3]: https://github.com/noda-hiroyuki-kit/pre-commit-vba-example/releases/tag/v0.1.3
[0.1.2]: https://github.com/noda-hiroyuki-kit/pre-commit-vba-example/releases/tag/v0.1.2
[0.1.1]: https://github.com/noda-hiroyuki-kit/pre-commit-vba-example/releases/tag/v0.1.1
[0.1.0]: https://github.com/noda-hiroyuki-kit/pre-commit-vba-example/releases/tag/v0.1.0
