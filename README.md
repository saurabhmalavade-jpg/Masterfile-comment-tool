# 📋 Masterfile Comment Extractor

A Streamlit web app that extracts cell comments from any product Masterfile (.xlsx) and generates a clean, formatted Excel report — grouped by Unique ID (SKU / EAN / MPN) with comments split into **Confirmation** and **Needs Action** columns based on cell highlight colour.

---

## How It Works

| Cell Highlight in Masterfile | Report Column |
|---|---|
| 🟠 Orange (`FFC000`) | Confirmation Comment |
| 🟡 Yellow (`FFFF00`) | Needs Action Comment |

### Output Report has 2 sheets:
- **Summary** — one row per SKU showing total comments and attributes affected
- **Comment Report** — full detail with SKU | Attribute | Confirmation Comment | Needs Action Comment

---

## Run Locally

### 1. Clone the repo
```bash
git clone https://github.com/YOUR_USERNAME/masterfile-comment-tool.git
cd masterfile-comment-tool
```

### 2. Create a virtual environment (recommended)
```bash
python -m venv venv
source venv/bin/activate        # Mac/Linux
venv\Scripts\activate           # Windows
```

### 3. Install dependencies
```bash
pip install -r requirements.txt
```

### 4. Run the app
```bash
streamlit run app.py
```

The app opens at `http://localhost:8501`

---

## Deploy on Streamlit Cloud (Free Hosting)

### Step 1 — Push to GitHub
```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/masterfile-comment-tool.git
git push -u origin main
```

### Step 2 — Deploy on Streamlit Cloud
1. Go to **[share.streamlit.io](https://share.streamlit.io)**
2. Sign in with your GitHub account
3. Click **"New app"**
4. Select your repo → branch: `main` → file: `app.py`
5. Click **"Deploy"** — your app goes live in ~2 minutes

Your app URL will be:
```
https://YOUR_USERNAME-masterfile-comment-tool-app-XXXXX.streamlit.app
```

---

## Project Structure

```
masterfile-comment-tool/
├── app.py              # Main Streamlit application
├── requirements.txt    # Python dependencies
└── README.md           # This file
```

---

## Supported Marketplaces
Works with any Masterfile format — Amazon DE/UK, eBay DE/UK/AU, OTTO, CDiscount, MediaMarkt, Walmart, and more.

---

## Comment Colour Convention
- **Orange highlight** → Confirmation needed (e.g. price confirmation, SKU structure sign-off)
- **Yellow highlight** → Action needed (e.g. missing mandatory field, character limit exceeded)
