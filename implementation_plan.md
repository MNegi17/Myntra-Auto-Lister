# Implementation Plan - Headless Cloud Deployment Support

Deploying to a cloud hosting platform (like Render, Heroku, or a VPS) requires the application to run in a headless, server-side environment. This means we cannot use Windows-specific Tkinter file browsers (`filedialog`) or read local hard drive paths (like `C:\` or Google Drive `G:\`). 

This plan adapts the project to automatically switch to **Cloud Mode** when deployed, introducing web-based file uploads and direct browser downloads.

## Proposed Changes

### 1. Configuration & API Key Support
We will allow the app to detect if it is running in a cloud environment via an environment variable (`IS_CLOUD=true`). We will also support setting the Gemini API Key as an environment variable (`GEMINI_API_KEY`).

#### [MODIFY] [vision_classifier.py](file:///c:/Users/Manann/Desktop/Myntra%20Auto%20Lister%20V2/vision_classifier.py)
- Modify `get_api_key()` to first check `os.environ.get("GEMINI_API_KEY")` before falling back to reading `gemini_api_key.txt`.

### 2. Flask Backend Adaptation
We will add secure file upload endpoints so the browser can send Excel sheets to the server, save them in a temporary upload directory, and pass their server paths to the core generation logic.

#### [MODIFY] [web_app.py](file:///c:/Users/Manann/Desktop/Myntra%20Auto%20Lister%20V2/web_app.py)
- Add `/api/config` to check if `IS_CLOUD` is enabled.
- Add `/api/upload` to securely handle multipart file uploads of Excel sheets.
- Update `/api/run` to:
  - Auto-generate a `output_dir` (e.g., in `static/temp_workspace/<run_id>`) in Cloud Mode if no folder is provided.
  - Return the run ID and output file name so the browser knows where to download the result.

### 3. Frontend Adaptation
We will modify the web UI so that it dynamically switches behavior based on the `is_cloud` configuration returned by the backend.

#### [MODIFY] [index.html](file:///c:/Users/Manann/Desktop/Myntra%20Auto%20Lister%20V2/static/index.html)
- Add hidden standard HTML `<input type="file" accept=".xlsx, .xls">` fields for:
  - Item Directory File
  - Content Sheet File
  - Myntra SKU Template File
- Change card layout so that when in Cloud Mode, "Select file" buttons activate these file upload elements, showing a loading indicator during transmission.
- Hide local directory selections (like Output Folder and Product Images Folder) in Cloud Mode.

#### [MODIFY] [script.js](file:///c:/Users/Manann/Desktop/Myntra%20Auto%20Lister%20V2/static/script.js)
- Fetch configuration `/api/config` on startup.
- Handle uploading selected files asynchronously to `/api/upload` and updating the hidden fields with their temporary server paths.
- Hide/clear default Windows paths in Cloud Mode.
- Show clear indications when files are uploading and when the final generated sheet is ready to download.

### 4. Cloud Deployment Setup
We will add configurations required by cloud providers (like Render or Heroku) to build the environment and launch the Flask server using a production-grade WSGI server (`gunicorn`).

#### [NEW] [requirements.txt](file:///c:/Users/Manann/Desktop/Myntra%20Auto%20Lister%20V2/requirements.txt)
- Create a requirements file containing all necessary dependencies:
  - `Flask`
  - `Flask-Cors`
  - `pandas`
  - `openpyxl`
  - `google-generativeai`
  - `Pillow`
  - `gunicorn` (for production server startup)

#### [NEW] [Procfile](file:///c:/Users/Manann/Desktop/Myntra%20Auto%20Lister%20V2/Procfile)
- Create a Heroku/Render launch script:
  ```text
  web: gunicorn web_app:app
  ```

---

## Verification Plan

### Automated/Unit Verification
- Run Flask locally in `IS_CLOUD=true` mode and test the upload of:
  - Item Directory Excel
  - Myntra SKU Template
- Test triggering the generation and verify that it downloads the populated template directly in the browser window.
- Check environment variables fallback for the Gemini API key.

### Manual Verification
- Guide the user on how to set up the repository on GitHub and deploy it to a free hosting provider like **Render** or a VPS.
