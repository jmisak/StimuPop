# Quick Start Guide

Get up and running in 3 minutes!

## Step 1: Install Python
**Windows:** Download from [python.org](https://www.python.org/downloads/)  
**Mac:** Already installed, or use `brew install python3`  
**Linux:** `sudo apt install python3 python3-pip`

## Step 2: Install Dependencies
Open terminal/command prompt in this folder and run:

```bash
pip install -r requirements.txt
```

**Having issues?** Try: `pip3 install -r requirements.txt`

## Step 3: Run the App

### Windows:
Double-click `run.bat` 

OR in command prompt:
```bash
streamlit run app.py
```

### Mac/Linux:
In terminal:
```bash
./run.sh
```

OR:
```bash
streamlit run app.py
```

## Step 4: Use the App

1. Your browser will open to `http://localhost:8501`
2. Upload `sample_data.xlsx` to test
3. Click "Generate Presentation"
4. Download your PowerPoint!

## Troubleshooting

### "Command not found: streamlit"
Run: `pip install streamlit` then try again

### "Port already in use"
Run: `streamlit run app.py --server.port 8502`

### Images not loading
- Check image URLs are publicly accessible
- Try the sample file first - it uses placeholder.com

## Next Steps

See [README.md](README.md) for complete documentation.

**Need help?** Check the "How to Use" section in the web app!
