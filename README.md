# MN Myntra Auto Lister PRO

**MN Myntra Auto Lister PRO** is a premium, AI-powered catalog automation web application built for fashion e-commerce vendors. The tool automates the tedious, manual process of catalog listing on Myntra by mapping size variants, prices, barcodes, and fabric attributes into Myntra's official SKU Excel templates in seconds. It also integrates **Google Gemini Vision AI** to analyze product photos and automatically classify visual design attributes like necklines, prints, patterns, and sleeve styles.

---

## 🚀 Key Features

* **Dual-Execution Engine**: Automatically switches between **Local Desktop Mode** (with native file/folder browser pickers) and **Cloud Mode** (with secure web file uploads and direct browser downloads).
* **AI Vision Auto-Fill**: Scan front/back images of garments via Google Gemini 2.5 Flash to auto-populate mandatory styling fields in Myntra templates.
* **Trained Learnings Gallery**: Review, correct, and save AI predictions. Trained models sync across all sessions.
* **Centralized Cloud Database**: Support for MongoDB Atlas to store and synchronize AI learnings across multiple users, regions, and devices in real-time.
* **Base64 Image Encoding**: Encodes gallery preview thumbnails into Base64 data strings inside the database to eliminate external image hosting dependencies.
* **Dynamic Sizing Mapping**: Automatically maps warehouse sizes to standard Myntra size charts based on category, gender, and brand criteria (e.g. *Purple United Kids* and *toothless*).
* **Live Execution Console**: Real-time server-sent logs stream directly into an interactive monospaced terminal in the UI.

---

## 🛠️ Technology Stack

* **Backend**: Python, Flask, Flask-CORS, PyMongo
* **Data Engineering**: Pandas, OpenPyXL
* **Artificial Intelligence**: Google Generative AI (Gemini 2.5 Flash), Pillow
* **Frontend**: HTML5, Vanilla CSS3 (Glassmorphic Design), Vanilla ES6 JavaScript
* **Production Web Server**: Gunicorn

---

## ⚙️ Configuration & Environment Variables

The application reads configurations dynamically from environment variables:

| Variable | Type | Description | Required in Cloud |
| :--- | :--- | :--- | :--- |
| `IS_CLOUD` | Boolean | Set to `true` to force headless browser mode (uploads/downloads). | Yes |
| `GEMINI_API_KEY` | String | Your API Key from Google AI Studio. | Optional (highly recommended) |
| `MONGO_URI` | String | MongoDB Atlas connection string for persistent cloud database. | Optional (local fallback) |
| `PORT` | Integer | Bind port for Flask/Gunicorn (defaults to `5000`). | No |

---

## 💻 Local Installation & Setup

1. **Clone the repository**:
   ```bash
   git clone https://github.com/MNegi17/Myntra-Auto-Lister.git
   cd Myntra-Auto-Lister
   ```

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Configure API Key (Local)**:
   Create a file named `gemini_api_key.txt` in the root folder and paste your Gemini API Key inside it.

4. **Launch the web application**:
   ```bash
   python web_app.py
   ```
   Open your browser and navigate to `http://127.0.0.1:5000`.

---

## ☁️ Cloud Deployment Instructions (e.g. Render)

1. **Create a Web Service** on Render and link your GitHub repository.
2. **Configure the build parameters**:
   * **Language**: `Python`
   * **Build Command**: `pip install -r requirements.txt`
   * **Start Command**: `gunicorn web_app:app`
   * **Instance Type**: `Free`
3. **Set Environment Variables** under Render's *Environment* tab:
   * Add `IS_CLOUD` = `true`
   * Add `GEMINI_API_KEY` = *your_google_ai_studio_api_key*
   * Add `MONGO_URI` = *your_mongodb_atlas_connection_string* (see instructions below)
4. **Deploy** the service. Your application will be live at `https://your-service-name.onrender.com`.

---

## 🗄️ MongoDB Atlas Database Setup (Free Tier)

To preserve AI learnings across cloud container restarts and share database entries among different users:
1. Sign up for a free account at [MongoDB Atlas](https://www.mongodb.com/cloud/atlas) and build an **M0 Free Shared Tier** cluster.
2. In the **Network Access** tab, add a rule to **Allow Access from Anywhere** (`0.0.0.0/0`).
3. In the **Database Access** tab, create a database user with read/write privileges.
4. Go to **Database** -> Click **Connect** -> Choose **Connect your application** (Python driver) and copy your connection string.
5. Replace `<password>` with your database user's password, and set this string as the `MONGO_URI` environment variable in your Render settings. The app will automatically sync classifications in real-time.

---

## 📝 License

Distributed under the MIT License. See `LICENSE` for more information.
