# R2D Reconciliation Tool

A user-friendly web interface for the R2D reconciliation tool. Non-technical users can drag-and-drop Excel files and get reconciliation results instantly.

## 🚀 Quick Start (Streamlit Community Cloud)

1. **Visit the deployed app**: (https://r2d-reconciliation-tool-v78hhyxf4pyqzpz66drnue.streamlit.app/)
2. **Upload your Excel file** with R2D and Chase sheets
3. **Configure sheet names** (defaults: "Repayments to Date" and "Chase")
4. **Click "Run Reconciliation"**
5. **Download your results**

## 📁 Files

- `streamlit_app.py` - The main Streamlit web application
- `r2d_recon.py` - Core reconciliation logic (do not modify)
- `requirements.txt` - Python dependencies
- `launch_app.py` - Local launcher script
- `Start_R2D_Tool.sh` - Shell script launcher (macOS/Linux)
- `Start_R2D_Tool.bat` - Batch script launcher (Windows)

## 🛠 Local Development

### Option 1: Using Python Virtual Environment

```bash
# Clone the repository
git clone https://github.com/loganalters123/r2d-reconciliation-tool.git
cd r2d-reconciliation-tool

# Create virtual environment
python3 -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run the app
streamlit run streamlit_app.py
```

### Option 2: Using the Launcher Scripts

**macOS/Linux:**
```bash
chmod +x Start_R2D_Tool.sh
./Start_R2D_Tool.sh
```

**Windows:**
```cmd
Start_R2D_Tool.bat
```

The app will open in your browser at `http://localhost:8501`

## 🌐 Deploy to Streamlit Community Cloud

### Prerequisites
- GitHub account
- This repository forked or copied to your GitHub

### Deployment Steps

1. **Go to [share.streamlit.io](https://share.streamlit.io)**

2. **Click "New app"**

3. **Fill in the details:**
   - Repository: `your-username/r2d-reconciliation-tool`
   - Branch: `main` (or `master`)
   - Main file path: `streamlit_app.py`
   - App URL: Choose a unique name like `your-name-r2d-tool`

4. **Click "Deploy!"**

5. **Wait for deployment** (usually 2-3 minutes)

6. **Share the URL** with your team: `https://your-app-name.streamlit.app`

### Environment Variables (if needed)
If your app requires any environment variables, add them in the Streamlit Cloud dashboard under "Settings" → "Secrets".

## 📊 How to Use

1. **Prepare your Excel file** with:
   - A sheet containing R2D data (default name: "Repayments to Date")
   - A sheet containing Chase data (default name: "Chase")

2. **Upload the file** using the drag-and-drop interface

3. **Configure settings:**
   - Sheet names (if different from defaults)
   - Optional date filter for debits

4. **Run reconciliation** and download results

## 🔧 Troubleshooting

### Common Issues

**"Could not import r2d_recon module"**
- Make sure `r2d_recon.py` is in the same directory as `streamlit_app.py`

**"Sheet not found" errors**
- Check that your Excel file has sheets with the exact names you specified
- Sheet names are case-sensitive

**"No data found" errors**
- Ensure your sheets contain data in the expected format
- Check that column headers match what the reconciliation algorithm expects

### Local Development Issues

**Streamlit not found**
```bash
pip install streamlit
```

**Other dependency issues**
```bash
pip install -r requirements.txt
```

**Permission issues on macOS/Linux**
```bash
chmod +x Start_R2D_Tool.sh
```

## 📞 Support

For technical support or feature requests, please:
1. Check the troubleshooting section above
2. Open an issue on GitHub
3. Contact the development team

## 🔒 Security Note

This tool processes financial data. When using the cloud version:
- Files are processed temporarily and not stored permanently
- Use the local version for highly sensitive data
- Ensure compliance with your organization's data policies

## Deploy to Streamlit Community Cloud (Free Public URL)

Follow these step-by-step instructions to deploy your app and get a public URL that anyone can access:

### Step 1: Create GitHub Repository

1. **Create a new GitHub repository:**
   - Go to [github.com](https://github.com) and sign in
   - Click "New repository" 
   - Name it `r2d-reconciliation-tool`
   - Make it **Public** (required for free Streamlit Community Cloud)
   - Click "Create repository"

2. **Upload your files to GitHub:**
   - Click "uploading an existing file"
   - Drag and drop these files:
     - `streamlit_app.py`
     - `r2d_recon.py` 
     - `requirements.txt`
     - `README.md`
   - Add commit message: "Initial commit - R2D reconciliation tool"
   - Click "Commit changes"

### Step 2: Deploy to Streamlit Community Cloud

1. **Go to Streamlit Community Cloud:**
   - Visit [share.streamlit.io](https://share.streamlit.io)
   - Click "Sign up with GitHub" (or sign in if you have an account)

2. **Connect your repository:**
   - Click "New app"
   - **Repository:** Select `your-username/r2d-reconciliation-tool`
   - **Branch:** `main` (default)
   - **Main file path:** `streamlit_app.py`
   - Click "Deploy!"

3. **Wait for deployment:**
   - Streamlit will install dependencies and start your app
   - This takes 2-3 minutes on first deployment
   - You'll see logs showing the deployment progress

### Step 3: Get Your Public URL

Once deployed, you'll get a URL like:
```
https://r2d-reconciliation-tool-[randomstring].streamlit.app
```

**Share this URL with your team!** Anyone can access it without installing anything.

### Step 4: Update Your App

To make changes:

1. Edit files locally
2. Upload updated files to GitHub (replace existing ones)
3. Streamlit will automatically redeploy within 1-2 minutes

### Troubleshooting

**Deployment fails?**
- Check that all files are uploaded to GitHub
- Ensure `requirements.txt` contains all dependencies
- Make sure repository is **Public**

**App shows errors?**
- Check the logs in Streamlit Community Cloud
- Verify your `r2d_recon.py` file works locally first
- Check that sheet names match your Excel files

**Need to update?**
- Just replace files in GitHub
- Streamlit auto-redeploys
- No need to manually restart

### Alternative: Local Network Sharing

If you prefer to run locally but share with your team:

```bash
streamlit run streamlit_app.py --server.address 0.0.0.0
```

Then share your local IP address (e.g., `http://192.168.1.100:8501`) with team members on the same network.

## Usage

1. **Upload Excel file** - Drag and drop your `.xlsx` file
2. **Configure settings** - Set sheet names (defaults: "Repayments to Date", "Chase") 
3. **Set date filter** - Optionally ignore debits before a certain date
4. **Run reconciliation** - Click the button and wait for processing
5. **Download results** - Get the reconciled Excel file with timestamp

## Support

- The app validates inputs and shows helpful error messages
- Check that your Excel file has the expected sheet names and column structure
- Output file includes timestamp: `Repayments_to_Date_recon-YYYY-MM-DD.xlsx`
