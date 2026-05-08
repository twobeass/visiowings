# Security Policy

## Supported versions

Only the latest minor release of `visiowings` receives security fixes during
its `0.x` phase. Once `1.0` is published, the latest two minor versions will
be supported.

| Version | Supported          |
|---------|--------------------|
| 0.6.x   | :white_check_mark: |
| < 0.6   | :x:                |

## Reporting a vulnerability

**Please do not open a public issue for security problems.**

Use GitHub Security Advisories to report privately:

  https://github.com/twobeass/visiowings/security/advisories/new

You should receive an acknowledgement within 72 hours. We aim to publish a
fix within 30 days for high-severity issues.

## Scope

In scope:

- Code execution via crafted VBA module files passed to `visiowings import`.
- Path traversal in export/import directory handling.
- Encoding-related bugs that cause data corruption on round-trip.
- Insecure release / supply chain issues (signing, build pipeline).

Out of scope:

- Vulnerabilities in Microsoft Visio, Microsoft Office, or Windows itself.
- Vulnerabilities in `pywin32` or `watchdog` (please report upstream).
- Self-XSS in error messages.
- Issues that require a malicious user already to have write access to the
  user's Visio document.

## Coordinated disclosure

We follow a 90-day coordinated disclosure timeline. If you believe a longer
embargo is appropriate, mention it in your initial report.
