# visiowings

**VBA Editor for Microsoft Visio with VS Code integration.**

Edit Visio VBA modules in your favourite editor with live sync, Git
support, and modern tooling. Inspired by [xlwings](https://www.xlwings.org/).

## At a glance

- 🔁 **Live sync** — save in VS Code, see the change in Visio.
- 🔄 **Bidirectional** — changes in Visio flow back to VS Code (optional).
- 📚 **Multi-document** — drawings, stencils, and templates side by side.
- 🦆 **Rubberduck** — `@Folder` annotations preserved as directories.
- 🌍 **Locale-aware** — automatic codepage detection for cp1252, cp1251,
  cp932, cp936, cp949, cp950, …
- 📦 **Single command install** — `pipx install visiowings`.

## Get started

<div class="grid cards" markdown>

- :material-rocket-launch:{ .lg .middle } **Installation**

    ---

    Install with `pipx` and connect to a running Visio instance.

    [:octicons-arrow-right-24: Installation guide](getting-started/installation.md)

- :material-cog:{ .lg .middle } **Configuration**

    ---

    Persist defaults in `.visiowings.toml` and use `visiowings init`.

    [:octicons-arrow-right-24: Configuration](getting-started/configuration.md)

- :material-translate:{ .lg .middle } **Codepages**

    ---

    Locale matrix and how visiowings picks the right encoding.

    [:octicons-arrow-right-24: Codepages](getting-started/codepages.md)

- :material-source-branch:{ .lg .middle } **Releasing**

    ---

    The fully automated release pipeline.

    [:octicons-arrow-right-24: Releasing](contributing/releasing.md)

</div>

## Project status

`visiowings` is currently in 0.x; small breaking changes between minor
versions are possible while the API stabilises. The full source code lives
on [GitHub](https://github.com/twobeass/visiowings).

Issues and pull requests are very welcome — please read
[CONTRIBUTING](https://github.com/twobeass/visiowings/blob/main/CONTRIBUTING.md)
first.
