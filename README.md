# Student Management API Tester & Bulk Upload

A local web application providing Swagger UI for testing LMS API requests from imported Postman collections, as well as an interactive **Bulk Student Upload** system for creating students from Excel (`.xlsx`) and CSV files.

## Run

### On Windows
Simply double-click `run.bat` (or run `run.bat` in Command Prompt / PowerShell).
It automatically detects Python, launches the application, and opens [http://127.0.0.1:8000](http://127.0.0.1:8000) in your default browser.

### On Linux / macOS
From this directory:

```sh
python app.py
```

Open [http://127.0.0.1:8000](http://127.0.0.1:8000).

## Features

### 1. Bulk Student Upload (Excel & CSV)
Switch to the **Bulk Student Upload** tab in the web interface to upload students in batches:
- **Supported Formats**: Excel workbooks (`.xlsx`) and CSV files (`.csv`). No third-party packages required (built with pure standard library).
- **Template Downloads**: Download sample templates directly from the UI or via:
  - `GET /api/bulk-upload/template.xlsx`
  - `GET /api/bulk-upload/template.csv`
- **Flexible Column Mapping**: Automatically recognizes column variations (case-insensitive, spaces/underscores ignored):
  - `firstName` (aliases: `First Name`, `fname`, `first`)
  - `lastName` (aliases: `Last Name`, `lname`, `last`, `surname`)
  - `middleName` (aliases: `Middle Name`, `middle`)
  - `primaryEmail` (aliases: `Primary Email`, `Email`, `emailid`)
  - `secondaryEmail` (aliases: `Secondary Email`, `altemail`)
  - `contactNo` (aliases: `Contact No`, `Phone`, `Mobile`, `Phone No`)
  - `address` (aliases: `Address`, `City`, `Location`)
  - `batchUniqueId` (aliases: `Batch Unique ID`, `Batch ID`, `Batch Code`)
  - `studentUniqueId` (aliases: `Student Unique ID`, `Student ID`, `Roll No`, `Reg No`)
  - `secondaryLangauge` (aliases: `Secondary Language`, `Language`)
  - `classStartDate` (aliases: `Class Start Date`, `Start Date`)
- **Default / Fallback Inputs**: Apply default `batchUniqueId`, `classStartDate`, or `secondaryLangauge` to rows that do not have them.
- **Pre-upload Validation**: Displays an interactive preview table showing mapped columns, row validity, and missing fields.
- **Real-Time Live Progress**: Watch each student's creation status live (`✓ Success` or `✗ Failed: reason`) with a visual progress bar.
- **Report Export**: Download a full execution report as a `.csv` file upon completion.

### 2. Student Password Reset by Email
Switch to the **Reset Password** tab in the web interface to reset and change student LMS login passwords:
- **Dedicated REST API**: `POST /api/student/reset-password` accepts:
  ```json
  {
    "email": "student@riceindia.org",
    "password": "NewPassword@123",
    "environment": "production"
  }
  ```
  *(Also supports multiple emails via `"emails": ["student1@...", "student2@..."]`)*.
- **LMS Environment Selector**:
  - `production` -> `https://lms.ricesmart.in/change-password-by-email`
  - `uat` -> `https://lms-uat.ricesmart.in/change-password-by-email`
  - `riceedu` -> `https://lms.riceedu.in/change-password-by-email`
  - `riceedu_uat` -> `https://lms-uat.riceedu.in/change-password-by-email`
- **Credential-Safe Proxy**: Automatically applies authorized server credentials (`Authorization: cd4e20a844e7cdfaa9d04d4d0555f7`) without exposing them to the client.
- **Web UI Controls**: Live email input, password generator with one-click strong password creation, show/hide toggle, and instant visual status reporting.

### 3. Swagger API Tester & Credential-Safe Proxy
The app serves a local Swagger UI backed by `/openapi.json` and proxies requests through `/api/proxy/{id}`:
- **Rice Smart Production Server**: `https://lms.ricesmart.in/api/v1/rice/student` (Production)
- **Rice Smart Testing / UAT Server**: `https://lms-uat.ricesmart.in/api/v1/rice/student` (UAT)
- **Rice Edu Production Server**: `https://lms.riceedu.in/api/v1/rice/student` (Rice Edu)
- **Rice Edu Testing / UAT Server**: `https://lms-uat.riceedu.in/api/v1/rice/student` (Rice Edu)
- Covers student, batch, batch-group, and defaulter endpoints across Production, UAT, Production/development, and Rice Edu environments.
- Uses each request's Postman authentication header, falling back to `postman/secrets/api_key.txt` when needed. Credentials are never exposed to the browser.
- The Rice Smart student POST is prefilled from `postman/secrets/student_request_example.json`.

### 3. Postman Integration & Upload
The repository is configured for Postman with `.postman/resources.yaml` linked to workspace `409b9255-4de7-47a9-8042-112bd0ef4aab`:

- **Option A (Automated Cloud Upload)**:
  Upload the collection directly to your Postman Cloud workspace via:
  ```sh
  python upload_to_postman.py --key PMAK-YOUR_KEY
  ```
  *(Or run `python upload_to_postman.py` and enter your key when prompted)*.

- **Option B (Postman Local View / VS Code Extension)**:
  Open this folder in Postman Desktop or the Postman VS Code extension. The collections in `postman/collections/` and spec in `postman/specs/` are detected automatically in Local View; click **Push / Sync** to sync to the remote workspace.

- **Option C (1-Click File Import)**:
  In Postman, click **Import** and choose:
  - `postman/Student-Management-API.postman_collection.json` (Full collection with Bulk Upload & proxy endpoints)
  - `postman/specs/openapi.json` (Full OpenAPI 3.0 specification)

To use a different port, set `PORT` or `API_TESTER_PORT` before running the server (defaults to port `8000`, binding to `0.0.0.0`).

---

## Free Cloud Deployment

This project is configured out-of-the-box for free cloud hosting with zero maintenance:

### 1. Deploy to Vercel (Free Serverless)
[![Deploy with Vercel](https://vercel.com/button)](https://vercel.com/new/clone?repository-url=https://github.com/AryaDXD/rice_api)

1. Push this repository to your GitHub account (`https://github.com/AryaDXD/rice_api`).
2. Go to [Vercel](https://vercel.com) and sign in with GitHub.
3. Click **Add New...** -> **Project** and select `rice_api`.
4. Leave all settings at their defaults (Vercel automatically detects [`vercel.json`](file:///home/rouqh/ATC/rice_api/vercel.json) and [`api/index.py`](file:///home/rouqh/ATC/rice_api/api/index.py)).
5. Click **Deploy**. Your app will be live with a free SSL domain (e.g. `https://rice-api.vercel.app`).

### 2. Deploy to Render (Free Web Service)
[![Deploy to Render](https://render.com/images/deploy-to-render-button.svg)](https://render.com/deploy?repo=https://github.com/AryaDXD/rice_api)

1. Sign up for a free account at [Render](https://render.com).
2. Click **New +** -> **Web Service**.
3. Connect your `rice_api` GitHub repository.
4. Render automatically picks up [`render.yaml`](file:///home/rouqh/ATC/rice_api/render.yaml) and [`Procfile`](file:///home/rouqh/ATC/rice_api/Procfile) (`python app.py`).
5. Choose the **Free** instance type and click **Create Web Service**. Your app is live at `https://rice-api.onrender.com`.

### 3. Deploy to Railway / Koyeb / Docker
- **Railway**: Connect repo -> Railway detects [`Procfile`](file:///home/rouqh/ATC/rice_api/Procfile) and deploys automatically.
- **Docker**: Build and run anywhere using the included [`Dockerfile`](file:///home/rouqh/ATC/rice_api/Dockerfile):
  ```sh
  docker build -t rice-api .
  docker run -p 8000:8000 rice-api
  ```

