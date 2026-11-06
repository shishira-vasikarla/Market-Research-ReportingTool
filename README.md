# Streamlit Q&A and Reporting — SQLite upgrade

Run in Codespaces:

```bash
pip install -r requirements.txt
streamlit run app.py
```

The app creates `project_history.db` automatically in its working directory.

- `projects`: uploaded CSV data and project metadata (CSV contents stored as JSON records)
- `answers`: one answer per company/question combination and pointer to the selected report version
- `answer_versions`: original, regenerated and manually edited answer snapshots
- `reports`: saved report snapshots and the exact answer-version IDs used to generate them

Create a project once, generate Q&A, compare versions, select and approve the version for the report, then generate/export. Later, resume the project and load an existing report without regenerating Q&A.

