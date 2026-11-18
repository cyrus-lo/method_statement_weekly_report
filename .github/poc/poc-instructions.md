# Project Overview
A proof-of-concept should be implemented as a single file whenever feasible. The goal is to validate one idea quickly with the least amount of structure, setup, and abstraction necessary.

# General Guidelines
- **Single File First**: Default to one file that contains the full implementation. Only split into multiple files if the user explicitly asks for it, if third-party open source code must be included, or if a single file is no longer practical.
- **Minimal Implementation**: Implement only what is required to prove the concept. Avoid full application scaffolding, extra layers, or production-oriented architecture.
- **Communication with Users**: Clarify the target outcome, show the result once it works, and keep the implementation aligned with the user's stated scope.
- **Documentation**: Keep documentation lightweight. Only add or update documentation when the user asks for it or when the POC requirements change.

# Development Workflow
- **Implementation Shape**: Prefer one runnable file for the POC. Keep logic together so the user can inspect, run, and modify it easily.
- **POC Root Folder**: If the POC requires multiple files or bundled third-party open source components, create a dedicated POC root folder named `test_xxxx/` and store all relevant POC files inside it.
- **Node.js POCs**: If the POC uses Node.js, use a dedicated `test_xxxx/` root folder and run `npm install` inside that folder root.
- **Python Environment Setup**: For Python POCs, use only the root-level `./.venv`. Reuse `./.venv` if it already exists. Otherwise, create a Python 3.12 `venv` at `./.venv`, activate it, and use it for all install and run steps. Do not use Conda or nested project virtual environments for POC setup.
- **No Unnecessary Scaffolding**: Do not create separate frontend, backend, database, Docker, or deployment setup unless the user explicitly requests them.
- **Dependencies**: Avoid adding dependencies unless they are required to demonstrate the core idea.
- **Versioning**: Use clear, descriptive change history when version control is involved, but keep the implementation itself minimal.

# Tech Approach
- **Choose the Simplest Stack**: Use the smallest viable toolset for the requested POC.
- **Optimize for Speed of Validation**: Favor readability and directness over extensibility.
- **Keep It Self-Contained**: Inputs, processing, and output should stay in the same file whenever possible.

# Coding Standards
- **Consistency**: Match the style of the surrounding workspace when editing existing code.
- **File Naming**: Name single-file POCs using the pattern test_xxxx.{file type}, where xxxx is a short descriptive identifier.
- **Folder Naming**: Name multi-file POC root folders using the pattern test_xxxx/ and keep all relevant POC files under that folder.
- **Naming**: Use descriptive names that make the POC flow easy to follow.
- **Comments**: Prefer self-documenting code. Add short comments only where the intent would otherwise be unclear.

