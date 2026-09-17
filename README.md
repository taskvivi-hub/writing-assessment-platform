# Writing Assessment Platform - Deploy Ready

This package is ready for GitHub + Streamlit Community Cloud deployment.

## Files to upload to GitHub
Upload all files and folders in this package, including:
- `app.py`
- `requirements.txt`
- `tasks.json`
- `.streamlit/config.toml`
- `.gitignore`
- `README.md`
- `SECRETS_EXAMPLE.toml`

Do NOT upload your real API key to GitHub.

## Step 1 - Create a GitHub repository
1. Go to GitHub.
2. Create a new repository, for example:
   `writing-assessment-platform`
3. Upload all files from this package.
4. Commit the files.

## Step 2 - Deploy in Streamlit Community Cloud
1. Open Streamlit Community Cloud.
2. Click **Create app** / **New app**.
3. Connect your GitHub account if needed.
4. Choose the repository you just created.
5. Main file path:
   `app.py`
6. Deploy.

After a short wait, Streamlit will create a public URL similar to:
`https://your-app-name.streamlit.app`

## Step 3 - Add Secrets
In the Streamlit app:
1. Open **App settings**.
2. Open **Secrets**.
3. Add:

```toml
OPENAI_API_KEY = "YOUR_OPENAI_API_KEY"
OPENAI_MODEL = "gpt-5.6-luna"
APP_BASE_URL = "https://YOUR-ACTUAL-APP-NAME.streamlit.app"
```

4. Save.
5. Restart / reboot the app if needed.

Important:
- Replace `YOUR_OPENAI_API_KEY` with your actual OpenAI API key.
- Replace the APP_BASE_URL value after Streamlit gives you the actual public URL.
- Never place the real API key in `app.py`, README, or GitHub.

## How the platform works

### Teacher
The teacher enters only:
- Task Title
- Genre / Writing Type

Then the platform generates a student task link.

### Student
The student:
1. Opens the shared task link.
2. Uploads a clear writing image.
3. Clicks **Submit for Assessment**.
4. Sees four scores and a total score out of 16.
5. Clicks **See Revision Suggestions**.
6. Sees:
   - Grammar corrections
   - Spelling corrections
   - Punctuation corrections
   - Brief Content & Organization Suggestions

## Rubric
- Content & Task Fulfillment
- Organization & Coherence
- Language Use
- Genre & Professional Appropriacy

Each dimension = 1-4 points.
Maximum = 16 points.

## Important limitation
The teacher enters only Task Title and Genre / Writing Type.
Therefore the system must not invent missing assignment requirements.
If later you want stricter Content & Task Fulfillment scoring, add a teacher-only Task Instructions field.

## Storage note
This prototype uses `tasks.json`.
On Streamlit Community Cloud, local file storage may not be permanent after restarts/redeployments.

For real long-term classroom use, the next version should use a persistent database such as:
- Supabase
- PostgreSQL
- Firebase / Firestore

That version can also add:
- teacher login
- student name / student ID
- submission records
- CSV export
- privacy / retention controls
