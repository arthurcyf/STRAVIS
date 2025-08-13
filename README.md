# STRAVIS_Automation
Final iteration of automation code, specifically for STRAVIS

# Changes for arthur:
PyAutoIt - a thin python wrapper around the AutoIt COM interface
- Problem is that the device running the app needs to have AutoIt running in the background
- But can be solved by using a virtual environment and deploying it
python-uiautomation - A pure-Python wrapper around Microsoft’s UI Automation (UIA) API.

# PyAutoIt
Steps to take:
1. Prepare the machine (installing the dependencies)
2. Write the python automation and integrate into GUI
3. Package and deploy

# python-uiautomation
1. Setup
2. Writing the automation
3. Integrate with the GUI
4. Packaging and deployment

"""
1) Create and activate a venv (PowerShell as Admin recommended):
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1   # If blocked: Set-ExecutionPolicy -Scope CurrentUser RemoteSigned

2) Install deps:
   pip install -r requirements.txt

3) Run the UI (preferably as Administrator):
   streamlit run app.py

4) Bring STRAVIS to the foreground when prompted and let it run.

Troubleshooting:
- If controls aren't found, ensure STRAVIS is on the primary monitor and scaling is 100%.
- Run the terminal **as Administrator** for better UIA access.
- If the Save dialog path differs, adjust click_save_as_tree_item().
- Replace the hard 20s sleep with a smarter waiter if load times vary.
"""