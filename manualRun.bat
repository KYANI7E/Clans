cmd /c start /position (-1204, -232)
powershell python C:\Users\Sage\OneDrive\Desktop\gitRepos\ClashOfClans_clanManagement\main.py -m -ng
powershell git pull
powershell git add Clash.xlsx
powershell git commit -m 'updating the clash file from bat manual run'
powershell git push
powershell .\Clash.xlsx