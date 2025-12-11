# AI Agent Tools

A collection of utility modules designed to support the implementation of AI agent tools. This repository provides various helper libraries and tools that can be used as building blocks for AI-powered automation workflows.

## Overview

This repository serves as a monorepo containing multiple independent modules, each focused on specific functionality that AI agents commonly need. The modules are designed to be:

- **Modular**: Each module can be used independently
- **Well-documented**: Comprehensive documentation for each module
- **AI-friendly**: Designed with AI agent integration in mind

## Available Modules

| Module                                       | Description                                                                                | Documentation                     |
| -------------------------------------------- | ------------------------------------------------------------------------------------------ | --------------------------------- |
| [PowerPoint Template Filler](msoffice/pptx/) | Fill JSON data into PowerPoint slide placeholders with support for text, images, and lists | [README](msoffice/pptx/README.md) |

## Getting Started

Each module has its own setup instructions and dependencies. Navigate to the specific module directory and follow the README instructions.

### Example: PowerPoint Template Filler

```bash
cd msoffice/pptx
python -m venv .venv
.venv\Scripts\activate  # Windows
pip install -r requirements.txt
python cli.py -i samples/input.json
```

## Repository Structure

```text
ai-agent-tools/
├── msoffice/
│   └── pptx/           # PowerPoint template data filler module
├── package.json        # Code commit tooling configuration
├── LICENSE             # Apache License 2.0
└── README.md           # This file
```

## Roadmap

Future modules planned for this repository:

- Additional MS Office automation tools (Excel, Word)
- File system utilities for AI agents
- Data transformation helpers
- API integration tools

## Contributing

Contributions are welcome! Please ensure your commits follow the [Conventional Commits](https://www.conventionalcommits.org/) specification.

### Code Commit

This repository uses the following tools for conventional commits:

- **Prettier** - Code formatting
- **Commitizen** - Conventional commit messages
- **Commitlint** - Commit message linting
- **Husky** - Git hooks
- **lint-staged** - Run linters on staged files

### Setup Environment for Code Commit

```bash
npm install
```

## License

This project is licensed under the Apache License 2.0 - see the [LICENSE](LICENSE) file for details.
