# Installation

## Requirements

- Windows 10/11 (visiowings talks to Visio via COM, which is Windows-only)
- Python 3.10 or newer
- Microsoft Visio with VBA support enabled

??? tip "Enable VBA access in Visio"

    Visio refuses to expose its VBA project to external tools by default.
    To enable it:

    1. Open Visio → **File** → **Options** → **Trust Center**.
    2. Click **Trust Center Settings…**
    3. Select **Macro Settings**.
    4. Tick :material-checkbox-marked-outline: **Trust access to the VBA
       project object model**.
    5. Restart Visio.

## Install with pipx (recommended)

[`pipx`](https://pipx.pypa.io/) installs visiowings into an isolated
virtual environment and adds the `visiowings` command to your PATH.

=== "Windows (PowerShell)"

    ```powershell
    py -m pip install --user pipx
    py -m pipx ensurepath
    pipx install visiowings
    ```

=== "Linux / WSL"

    ```bash
    python3 -m pip install --user pipx
    python3 -m pipx ensurepath
    pipx install visiowings
    ```

To upgrade later:

```bash
pipx upgrade visiowings
```

## Install with pip

```bash
pip install visiowings
```

## Standalone Windows EXE

Download `visiowings.exe` from the
[latest GitHub Release](https://github.com/twobeass/visiowings/releases/latest)
and place it on your PATH. The `.exe.sigstore` file next to it is a
Sigstore signature you can verify with
[`sigstore-python`](https://github.com/sigstore/sigstore-python).

## Install from source

```bash
git clone https://github.com/twobeass/visiowings.git
cd visiowings
pip install -e ".[dev]"
pre-commit install
```

This is the recommended setup for development; see
[CONTRIBUTING](https://github.com/twobeass/visiowings/blob/main/CONTRIBUTING.md).

## Verify the install

```bash
visiowings --version
visiowings --help
```
