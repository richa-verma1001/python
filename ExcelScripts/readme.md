
python -m venv venv    // Create
.\venv\Scripts\activate   // Activate 
deactivate   // deactivate
pip install pandas openpyxl
python FilterExcel.py



Github Key Check

dir $env:USERPROFILE\.ssh
Set-Service -Name ssh-agent -StartupType Automatic
Start-Service ssh-agent
Get-Content $env:USERPROFILE\.ssh\id_ed25519.pub | clip

Add key to Github
    Settings → SSH and GPG keys → New SSH key

Test connection - 
ssh -T git@github.com