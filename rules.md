# Project Rules — PDF_BULK_FILL

These rules apply to all contributors and AI agents working on `PDF_BULK_FILL`.

## 1. Repository & Project Board
- **Repository**: `git@github.com:vipulasrilanka/PDF_BULK_FILL.git`
- **Project Board**: `https://github.com/users/vipulasrilanka/projects/2`
- All work must be tracked as GitHub Issues in `vipulasrilanka/PDF_BULK_FILL`.
- Never modify files outside this repository (specifically, keep the legacy reference `../src/` folder untouched).

## 2. Ticket Workflow & Approval Gate
1. **Plan First**: Every feature, bugfix, or chore must have a GitHub Issue describing scope, requirements reference, and acceptance criteria.
2. **Approval Required Before Implementation**: No ticket may be implemented until explicitly approved by the project owner (`vipulasrilanka`).
3. **Update & Close**: Once implementation and tests pass, update the GitHub Issue with a summary of changes, test results, and code coverage percentage, then close the issue.

## 3. Testing & Code Coverage Gate
- Every ticket **must** include corresponding automated test cases (`pytest`) covering happy paths, edge cases, and error handling.
- **Minimum 80% code coverage** (`--cov-fail-under=80`) is strictly required.
- All tests must pass **before** any code is committed or pushed to GitHub.

## 4. Folder Organization
```text
PDF_BULK_FILL/
├── rules.md                  # Common project & workflow rules
├── README.md                 # Project overview & usage
├── docs/
│   └── REQUIREMENTS.md       # Brief functional & non-functional requirements
├── src/
│   └── pdf_bulk_fill/        # Application source code
├── tests/                    # Automated unit & integration tests
├── templates/
│   ├── pdf/                  # PDF form & letterhead templates (e.g., RR-Form.pdf)
│   └── excel/                # Mapping template workbooks (e.g., TEMPLATE.xlsx)
└── data/
    ├── input/                # Source Excel files & employee ID images (gitignored)
    └── output/               # Generated output PDFs (gitignored)
```

## 5. Git & Security Hygiene
- Never commit sensitive employee data, real NIC images, passwords, or generated output PDFs (`data/input/*` and `data/output/*` remain gitignored, except `.gitkeep`).
- Reference the relevant issue number (e.g., `Refs #1` or `Closes #1`) in commit messages.
- Keep commits focused to a single approved ticket.
