# Upload v13 to GitHub by Web Page

This version is still a single-file Streamlit deploy package.

Upload these files directly into your GitHub repo root:

```text
app.py
requirements.txt
packages.txt
runtime.txt
README.md
UPLOAD_TO_GITHUB.md
```

Steps:

1. Open your GitHub repo.
2. Click **Add file**.
3. Click **Upload files**.
4. Drag these files into the upload area.
5. Click **Commit changes**.
6. Streamlit Cloud should redeploy automatically if connected to the repo.
7. If it does not, open Streamlit Cloud → Manage app → Reboot / Redeploy.

Do not upload the zip file itself.
