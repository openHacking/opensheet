# Security policy

OpenSheet 0.1.0 is a development preview. The active development branch receives fixes; no stable support matrix or security response SLA is promised.

## Reporting a vulnerability

Do not put exploit payloads, private workbooks or secrets in public issues. Use the repository's **Security → Report a vulnerability** option when private vulnerability reporting is enabled. If that option is unavailable, open an issue requesting a private contact channel without revealing vulnerability details. There is no dedicated security email address at this stage.

## Security boundaries

- Plugins execute in the host page's JavaScript context. Capabilities reduce accidental misuse; they are not a security sandbox. Install trusted plugins only.
- Models and adapters validate schemas and resource limits. The demo uses a Worker, file-size and ZIP prechecks, and parsing timeouts; these do not guarantee memory isolation.
- Formulas do not use `eval`, network functions or arbitrary code. HTML and LaTeX exports escape text. Delimited exports enable formula-injection protection by default.
- Treat JSON, external SheetJS objects and plugin output as untrusted. Server hosts need their own authentication, authorization, process isolation and resource limits.
- The demo contains no upload service, analytics or telemetry. Integrating hosts control their own privacy behavior.
