# Workspace Folder Structure

```text
method_statement/
├── django/
│   ├── .env
│   ├── .env.example
│   ├── .venv/
│   ├── .vscode/
│   │   └── settings.json
│   ├── output/
│   ├── Method_Statement_Interaction_Weekly_Report.py
│   ├── Method_Statement_Interaction_Weekly_Report_template.xlsx
│   ├── steps.md
│   ├── requirements.txt
│   ├── sessions.sql
│   └── submission_counts.sql
│
├── ecm/
│   ├── Step_1_Create_ECM_Record.py
│   ├── Step_2_Create_SMM_Summary.py
│   ├── requirements.txt
│   └── steps.md
├── .github/
├── .gitignore
├── folder_structure.md
└── hello.py
```

## Workflow ownership

- `django/`: completed PostgreSQL/Django reporting workflow and its local environment, credentials, SQL, template, documentation, and generated reports.
- `ecm/`: ECM and SMM weekly reporting workflow, including record creation, summary generation, and process documentation.
- Root files: shared repository guidance, Git ignore rules, and the folder structure document.
