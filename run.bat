cmd /c start /position (-1204, -232)
powershell python C:\Users\Sage\OneDrive\Desktop\gitRepos\ClashOfClans_clanManagement\main.py automation all
powershell git pull
powershell git add Clash.xlsx
powershell git commit -m 'updating the clash file from bat automatic run'
powershell git push