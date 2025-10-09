# get latest download url
$URL = "https://api.github.com/repos/microsoft/winget-cli/releases/latest"
$URL = (Invoke-WebRequest -Uri $URL).Content | ConvertFrom-Json |
        Select-Object -ExpandProperty "assets" |
        Where-Object "browser_download_url" -Match '.msixbundle' |
        Select-Object -ExpandProperty "browser_download_url"

# download
Invoke-WebRequest -Uri $URL -OutFile "Setup.msix" -UseBasicParsing

# install
Add-AppxPackage -Path "Setup.msix"

# delete file
Remove-Item "Setup.msix"

#start new powershell process to install omnissa so that winget is available in the current session
Start-Process "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Command `"iwr -useb https://raw.githubusercontent.com/ijsselgemeenten/osdcloud/refs/heads/main/winget-apps.ps1 | iex`""
