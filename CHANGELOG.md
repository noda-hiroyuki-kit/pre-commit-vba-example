# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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

[Unreleased]: https://github.com/noda-hiroyuki-kit/pre-commit-vba-example/compare/v0.1.1...HEAD
[0.1.1]: https://github.com/noda-hiroyuki-kit/pre-commit-vba-example/releases/tag/v0.1.1
[0.1.0]: https://github.com/noda-hiroyuki-kit/pre-commit-vba-example/releases/tag/v0.1.0
