#upload device
# Mapping tussen leesbare namen en echte GroupTags
$GroupTagMap = @{
    "Persoonlijk Device"       = "GRIJPersonalDevice"
    "NUC Werkplek"             = "GRIJNUC"
    "Werkplein Cursus Laptop"  = "GRIJWERKPLEIN"
    "Werkcentrum Rijnmond"     = "WCR"
    "Shared Device"            = "GRIJSharedDevice"
    "Publieke Zuil (Krimpen)"  = "KRIPUBLICZUIL"
    "Wallboard (Krimpen)"      = "KRIKCCWALLBOARD"
    "Gezondheidscentrum"       = "GRIJGEZONDHEIDSCENTRUM"
    "Werkplein Kiosk"          = "WERKPLEINKIOSK"
}

# Form aanmaken
$form = New-Object System.Windows.Forms.Form
$form.Text = "Intune Device Profiel Selectie"
$form.Size = New-Object System.Drawing.Size(500,280)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::WhiteSmoke

# Titel
$title = New-Object System.Windows.Forms.Label
$title.Text = "Selecteer een Intune Device Profiel"
$title.Font = New-Object System.Drawing.Font("Segoe UI",16,[System.Drawing.FontStyle]::Bold)
$title.ForeColor = [System.Drawing.Color]::SteelBlue
$title.AutoSize = $true
$title.Location = New-Object System.Drawing.Point(20,20)
$form.Controls.Add($title)

# Subtitel
$subtitle = New-Object System.Windows.Forms.Label
$subtitle.Text = "Kies hieronder welk profiel je wilt toepassen voor dit device:"
$subtitle.Font = New-Object System.Drawing.Font("Segoe UI",10)
$subtitle.AutoSize = $true
$subtitle.Location = New-Object System.Drawing.Point(22,60)
$form.Controls.Add($subtitle)

# Dropdown
$dropdown = New-Object System.Windows.Forms.ComboBox
$dropdown.Location = New-Object System.Drawing.Point(25,100)
$dropdown.Size = New-Object System.Drawing.Size(430,35)
$dropdown.Font = New-Object System.Drawing.Font("Segoe UI",12)
$dropdown.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$dropdown.Items.AddRange($GroupTagMap.Keys)

# Standaard selectie -> Persoonlijk Device
$defaultKey = "Persoonlijk Device"
$dropdown.SelectedIndex = $dropdown.Items.IndexOf($defaultKey)
$form.Controls.Add($dropdown)

# OK knop
$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "Start uitrol"
$okButton.Font = New-Object System.Drawing.Font("Segoe UI",11,[System.Drawing.FontStyle]::Bold)
$okButton.BackColor = [System.Drawing.Color]::SteelBlue
$okButton.ForeColor = [System.Drawing.Color]::White
$okButton.FlatStyle = 'Flat'
$okButton.Size = New-Object System.Drawing.Size(120,40)
$okButton.Location = New-Object System.Drawing.Point(335,170)
$okButton.Add_Click({
    $form.Tag = $GroupTagMap[$dropdown.SelectedItem]
    $form.Close()
})
$form.Controls.Add($okButton)

# Cancel knop
$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = "Annuleren"
$cancelButton.Font = New-Object System.Drawing.Font("Segoe UI",11)
$cancelButton.BackColor = [System.Drawing.Color]::LightGray
$cancelButton.ForeColor = [System.Drawing.Color]::Black
$cancelButton.FlatStyle = 'Flat'
$cancelButton.Size = New-Object System.Drawing.Size(120,40)
$cancelButton.Location = New-Object System.Drawing.Point(200,170)
$cancelButton.Add_Click({
    $form.Tag = $null
    $form.Close()
})
$form.Controls.Add($cancelButton)

# Form tonen
$form.ShowDialog() | Out-Null
$GroupTag = $form.Tag

# Als er geen keuze is gemaakt (cancel)
if ([string]::IsNullOrWhiteSpace($GroupTag)) {
    Write-Host "Geen GroupTag gekozen, script gestopt." -ForegroundColor Red
    exit
}

Write-Host "Gekozen GroupTag: $GroupTag" -ForegroundColor Green

# Settings
$OSName = 'Windows 11 24H2 x64'
$OSEdition = 'Pro'
$OSActivation = Volume'
$OSLanguage = 'nl-nl'
$GroupTag = "$GroupTag"
$TimeServerUrl = "time.cloudflare.com"
$OutputFile = "X:\AutopilotHash.csv"
$TenantID = [Environment]::GetEnvironmentVariable('OSDCloudAPTenantID','Machine') # $env:OSDCloudAPTenantID doesn't work within WinPe
$AppID = [Environment]::GetEnvironmentVariable('OSDCloudAPAppID','Machine')
$AppSecret = [Environment]::GetEnvironmentVariable('OSDCloudAPAppSecret','Machine')

#Set Global OSDCloud Vars
$Global:MyOSDCloud = [ordered]@{
    BrandColor = "#0096FF"
    Restart = [bool]$False
    RecoveryPartition = [bool]$True
    OEMActivation = [bool]$True
    WindowsUpdate = [bool]$True
    WindowsUpdateDrivers = [bool]$True
    WindowsDefenderUpdate = [bool]$True
    SetTimeZone = [bool]$True
    ClearDiskConfirm = [bool]$False
    ShutdownSetupComplete = [bool]$false
    SyncMSUpCatDriverUSB = [bool]$True
    CheckSHA1 = [bool]$True
}

# Largely reworked from https://github.com/jbedrech/WinPE_Autopilot/tree/main
Write-Host "Autopilot Device Registration Version 1.0"

# Set the time
$DateTime = (Invoke-WebRequest -Uri $TimeServerUrl -UseBasicParsing).Headers.Date
Set-Date -Date $DateTime

# Download required files
$oa3tool = 'https://raw.githubusercontent.com/{Redacted}/{Redacted}/main/oa3tool.exe'
$pcpksp = 'https://raw.githubusercontent.com/{Redacted}/{Redacted}/main/PCPKsp.dll'
$inputxml = 'https://raw.githubusercontent.com/{Redacted}/{Redacted}/main/input.xml'
$oa3cfg = 'https://raw.githubusercontent.com/{Redacted}/{Redacted}/main/OA3.cfg'

Invoke-WebRequest $oa3tool -OutFile $PSScriptRoot\oa3tool.exe
Invoke-WebRequest $pcpksp -OutFile X:\Windows\System32\PCPKsp.dll
Invoke-WebRequest $inputxml -OutFile $PSScriptRoot\input.xml
Invoke-WebRequest $oa3cfg -OutFile $PSScriptRoot\OA3.cfg

# Create OA3 Hash
If((Test-Path X:\Windows\System32\wpeutil.exe) -and (Test-Path X:\Windows\System32\PCPKsp.dll))
{
	#Register PCPKsp
	rundll32 X:\Windows\System32\PCPKsp.dll,DllInstall
}

#Change Current Diretory so OA3Tool finds the files written in the Config File 
&cd $PSScriptRoot

#Get SN from WMI
$serial = (Get-WmiObject -Class Win32_BIOS).SerialNumber

#Run OA3Tool
&$PSScriptRoot\oa3tool.exe /Report /ConfigFile=$PSScriptRoot\OA3.cfg /NoKeyCheck

#Check if Hash was found
If (Test-Path $PSScriptRoot\OA3.xml) 
{
	#Read Hash from generated XML File
	[xml]$xmlhash = Get-Content -Path "$PSScriptRoot\OA3.xml"
	$hash=$xmlhash.Key.HardwareHash

	$computers = @()
	$product=""
	# Create a pipeline object
	$c = New-Object psobject -Property @{
 		"Device Serial Number" = $serial
		"Windows Product ID" = $product
		"Hardware Hash" = $hash
		"Group Tag" = $GroupTag
	}
	
 	$computers += $c
	$computers | Select "Device Serial Number", "Windows Product ID", "Hardware Hash", "Group Tag" | ConvertTo-CSV -NoTypeInformation | % {$_ -replace '"',''} | Out-File $OutputFile
}

# Upload the hash
Start-Sleep 30

#Get Modules needed for Installation
#PSGallery Support
Invoke-Expression(Invoke-RestMethod sandbox.osdcloud.com)
Install-Module WindowsAutoPilotIntune -SkipPublisherCheck -Force

#Connection
Connect-MSGraphApp -Tenant $TenantId -AppId $AppId -AppSecret $AppSecret

#Import Autopilot CSV to Tenant
Import-AutoPilotCSV -csvFile $OutputFile

Write-Host "Start-OSDCloud -OSName $OSName -OSEdition $OSEdition -OSActivation $OSActivation -OSLanguage $OSLanguage"
Start-OSDCloud -OSName $OSName -OSEdition $OSEdition -OSActivation $OSActivation -OSLanguage $OSLanguage
