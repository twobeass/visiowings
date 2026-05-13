# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

From version 0.6.1 onwards, this changelog is generated automatically by
[release-please](https://github.com/googleapis/release-please) from
[Conventional Commit](https://www.conventionalcommits.org/) messages.

## [1.0.1](https://github.com/twobeass/visiowings/compare/1.0.0...1.0.1) (2026-05-13)


### Bug Fixes

* **release:** unbreak the 1.0.0 release pipeline ([b3f6543](https://github.com/twobeass/visiowings/commit/b3f65434756ed67816fc57b2ddf8156eb1f8574c))
* **release:** unbreak the 1.0.0 release pipeline (-&gt; 1.0.1) ([6d5ff99](https://github.com/twobeass/visiowings/commit/6d5ff991828ca4c4e8fa9b44709024b6bffea5db))

## [1.0.0](https://github.com/twobeass/visiowings/compare/0.6.1...1.0.0) (2026-05-13)


### Features

* add typed exceptions, retry helper, and structured logging ([5815429](https://github.com/twobeass/visiowings/commit/5815429c858e2f67cdce93f1e6a840d9ada6e375))
* **cli:** project config + init wizard + opt-out update check ([7bc6f05](https://github.com/twobeass/visiowings/commit/7bc6f058bcd65e97b74367d8ff5e8d351cfc0a47))
* **tests:** migrate UAT suite in-tree, drop external vbatest dependency ([f19aef6](https://github.com/twobeass/visiowings/commit/f19aef6522bcf525717449dad1f1d0e5241ab7ea))
* **uat-§C1:** add --non-interactive flag to `visiowings init` ([203f0cc](https://github.com/twobeass/visiowings/commit/203f0ccecd4fd05976c955a53fb2bb961c6ed2ce))
* **uat-§C4:** emit structured [DEBUG] logs at the CLI checkpoints ([a02b873](https://github.com/twobeass/visiowings/commit/a02b873ac1410449ef37e8c4de14ec696cb50194))


### Bug Fixes

* address CodeQL findings (cyclic import + 4 silent excepts) ([b5141fa](https://github.com/twobeass/visiowings/commit/b5141fa44c716fe85dc1dadba8cbdd6e2ea97fba))
* **cli:** break cli ↔ interactive import cycle by inverting the dep ([a3fcaa3](https://github.com/twobeass/visiowings/commit/a3fcaa305308abfc5232b1423b44f3fa66c8442d))
* harden COM error handling, threading, BOM decoding and CLI input ([61981eb](https://github.com/twobeass/visiowings/commit/61981eb410461e7c47bb6ba43da36aaa7bb9ecb2))
* **tests:** ignore UAT collection on non-Windows runners ([29f95d4](https://github.com/twobeass/visiowings/commit/29f95d4b996c1a4240a4dfce43579c3e46d0af7d))
* **tests:** make tests/uat runnable on a real Windows + Visio machine ([2b9bf6e](https://github.com/twobeass/visiowings/commit/2b9bf6eb21ed63194a6fca91204620096bbccb80))
* **tests:** make tests/uat runnable on a real Windows + Visio machine ([ddd4c2d](https://github.com/twobeass/visiowings/commit/ddd4c2dff0375dec5a67a17df9ff08156d7e0e70))
* **tests:** satisfy CodeQL flow analysis in _require_user_opened_visio ([b8a2e98](https://github.com/twobeass/visiowings/commit/b8a2e98bb145829ff3449f1378ec41de67ea5e15))
* **uat-§B1:** make entry-point pipx-installable on non-Windows hosts ([25bb4ba](https://github.com/twobeass/visiowings/commit/25bb4bae48681d86c13025d82a4886a573127016))
* **uat-§C1:** reconfigure stdout to utf-8 so init banner survives cp1252 ([54f8be7](https://github.com/twobeass/visiowings/commit/54f8be738e8eb3106080452e5cf26ef44268e25d))
* **uat-§C3:** no Python traceback on user-facing CLI errors ([8c9214e](https://github.com/twobeass/visiowings/commit/8c9214e2d551b9cf3ab4e8e9b58b68fe2a156117))
* **uat-iter3-#1:** dedupe Option Explicit after VBComponents.Import ([8c6d2a8](https://github.com/twobeass/visiowings/commit/8c6d2a88124d28d68682b1bb019c9b2a860ac6fc))
* **uat-iter3-#2:** honor --force in the import batch-conflict prompt ([e358c11](https://github.com/twobeass/visiowings/commit/e358c11ec561c084355bf4e55d031ac55dc091f9))
* **uat-iter3-#3:** opt-in --ephemeral flag clears Visio's dirty marker ([69f5653](https://github.com/twobeass/visiowings/commit/69f5653997083ea2d0e68fe4d202e954cc672e9d))
* **uat-iter3-#4:** refuse to lose a module to a silent codepage mismatch ([9eaadbf](https://github.com/twobeass/visiowings/commit/9eaadbf5cc99c55e6a5ad680077594a9586243d3))
* **uat-iter4-#5:** export exits non-zero when any document write fails ([412ad3d](https://github.com/twobeass/visiowings/commit/412ad3d7e07af8eeb5a4e2add9b4c41f3f7ad879))
* **uat-iter4-#6:** make `import --rd` actually re-import on a folder move ([8e2aee4](https://github.com/twobeass/visiowings/commit/8e2aee4707f03eb68c20ef065edb4ea65b3e6dbb))
* **uat-iter4-#7:** bidirectional polling never prompts on stdin ([2056c00](https://github.com/twobeass/visiowings/commit/2056c002bb5c5580f9d6ab0d47c306baf12da073))


### Documentation

* add comprehensive Human UAT checklist for release sign-off ([e78365c](https://github.com/twobeass/visiowings/commit/e78365c8969620064af89abb5cdf8a06556498c4))
* document Phase H tooling (just, nox, Sigstore, Scorecard, .vscode) ([826a068](https://github.com/twobeass/visiowings/commit/826a068dc6943ddbb632432f3ca329cc7be0807c))
* migrate to MkDocs Material with GitHub Pages auto-deploy ([8a0b695](https://github.com/twobeass/visiowings/commit/8a0b695b8c0176cba6b0c16f0b4d29537ab7f10b))


### Continuous Integration

* add release-please + PyPI Trusted Publishing + SBOM ([db5265b](https://github.com/twobeass/visiowings/commit/db5265b4b9ac67d25218c60d9822b176d04597c3))
* align release-please tag pattern with existing 0.x.y tags ([ebfb22c](https://github.com/twobeass/visiowings/commit/ebfb22c340e911731ec7a6e691fc054a9edde069))
* bump pre-commit ruff to v0.15.12 + drop dead `force = True` write ([e6a6475](https://github.com/twobeass/visiowings/commit/e6a6475782277c1be7db4be45253c4090bcf2307))
* make Dependency Review advisory when graph is unavailable ([7cf7915](https://github.com/twobeass/visiowings/commit/7cf7915d4aa23b58edafaf9a434eeb24e9a01222))
* **release-please:** switch to scoped PAT instead of GITHUB_TOKEN ([5fc891d](https://github.com/twobeass/visiowings/commit/5fc891ddfe149475b792184c76e8b1ed3db83163))
* **release-please:** use scoped PAT instead of GITHUB_TOKEN ([cfebdd1](https://github.com/twobeass/visiowings/commit/cfebdd16ac3c472ba54b2e8c3c0dc7d139517bc1))
* replace stub CI with full quality-gate pipeline ([ce29ab3](https://github.com/twobeass/visiowings/commit/ce29ab324a41ec2dd05d860f2a8ddeb1709a1519))
* turn the pipeline green — fix lint, types, build, security gates ([c1ef77e](https://github.com/twobeass/visiowings/commit/c1ef77e6537b79b3f5e2bf82d2c28b2660b035a2))


### Chores

* declare 1.0.0 stable ([e159501](https://github.com/twobeass/visiowings/commit/e159501d415d02536a735c47fbb73f4b8c96059b))

## [0.6.1] — 2026-05-08

This is the first release whose changelog is managed by release-please.
Earlier history is preserved in the Git log.

### Highlights

- VBA editor for Microsoft Visio with VS Code live-sync.
- Multi-document support (drawings + stencils + templates).
- Rubberduck `@Folder` annotation support.
- Auto-detection of locale codepage (cp1252, cp1251, cp1250, cp932, cp936,
  cp949, cp950, …).

[0.6.1]: https://github.com/twobeass/visiowings/releases/tag/0.6.1
