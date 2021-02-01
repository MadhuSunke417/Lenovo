#notes
#This script will download latest BIOS from lenovo and extarct to same directory.
#install - extract and install WINUPTP64.EXE with -s switch
#Author: Madhu Sunke

$inTS = $false
$testmode = $false
try{
    $tsenv = New-Object -COMObject Microsoft.SMS.TSEnvironment
    Write-Host "BIOS running inside TS env"
    $inTS = $true
}catch{
    Write-Host "BIOS running outside TS env"
}

$Model = ((Get-WmiObject -Class Win32_ComputerSystem | Select-Object -ExpandProperty Model).SubString(0, 4)).Trim()
$BIOSVer = ((Get-WmiObject -Class win32_bios | Select-Object -ExpandProperty SMBIOSBIOSVersion).Split("(")[-1]).trimend(")").trim()
#$systemFamily = (Get-WmiObject -Class Win32_ComputerSystem | Select-Object -ExpandProperty SystemFamily).trim()
$DriverURL = "https://download.lenovo.com/cdrt/td/catalogv2.xml"

Write-Host "Loading Lenovo Catalog XML...." -ForegroundColor Yellow

if (($DriverURL.StartsWith("https://")) -OR ($DriverURL.StartsWith("http://"))) {
    try { $testOnlineConfig = Invoke-WebRequest -Uri $DriverURL -UseBasicParsing } catch { <# nothing to see here. Used to make webrequest silent #> }
    if ($testOnlineConfig.StatusDescription -eq "OK") {
        try {
            $webClient = New-Object System.Net.WebClient
            $webClient.Encoding = [System.Text.Encoding]::UTF8
            $Xml = [xml]$webClient.DownloadString($DriverURL)
            Write-host "Successfully loaded $DriverURL"
        }
        catch {
            $ErrorMessage = $_.Exception.Message
            Write-Host "Error, could not read $DriverURL" 
            Write-Host "Error message: $ErrorMessage"
            Exit 1
        }
    }
    else {
        Write-Host "The provided URL to the config does not reply or does not come back OK"
        Exit 1
    }
}

$DriverPath = "$env:ProgramData\LenovoBIOSUpdate"
$extractedPath = "$DriverPath\Extracted"

if(-not(Test-Path $DriverPath)){
    New-Item -Path $DriverPath -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
}

$ModelDriverPackInfo = $Xml.ModelList.Model | Where-Object -FilterScript {$_.Types.Type -match $Model} 
$BIOSInfo = $ModelDriverPackInfo.BIOS | Select-Object -Property *
$Downloadurl = $BIOSInfo.'#text'
$BIOSVerfromLen = $BIOSInfo.version
$TargetFileName = $Downloadurl.Split("/")[-1]
$AlreadyDownloadedFileName = (Get-ChildItem $DriverPath -Filter *.exe).Name
$TargetFilePathName = "$($DriverPath)\$($TargetFileName)"
Write-Host "Available BIOS version from Lenovo for $Model is $BIOSVerfromLen"

if(([version]$BIOSVer -eq $BIOSVerfromLen) -and (-not $testmode)){
    Write-Host "Machine already installed with latest version of BIOS : $BIOSVer"
    exit
}

if ((Test-Path $TargetFilePathName) -and (Test-Path "$extractedPath\$($TargetFileName.Split(".")[0])"))
{
    Write-Output "Aleady Contains Latest BIOS in Expanded Folder"
    if($inTS){
        $tsenv.value('LenovoBIOS') = "Success"
        }
        Write-Host "----------------------------" -ForegroundColor DarkGray
    }else{  
        if(Test-Path $extractedPath -ErrorAction SilentlyContinue)
          {
              Remove-Item $extractedPath -Force -Recurse -ErrorAction SilentlyContinue
            }
            if($AlreadyDownloadedFileName -ne $TargetFileName){
                Invoke-WebRequest -Uri $Downloadurl -OutFile $TargetFilePathName -UseBasicParsing
            }else{
                Write-Host "Skip Download and start extract Process"
            }
            $LenovoSilentSwitches = @(
            "/VERYSILENT"
            "/DIR=$extractedPath\$($TargetFileName.Split(".")[0])"
            "/Extract=YES"
            )
    Start-Process -FilePath $TargetFilePathName -ArgumentList $LenovoSilentSwitches -Verb RunAs
    # Wait for Lenovo BIOS Process To Finish
    While ((Get-Process) | Where-Object {$_.Name -eq $TargetFileName.Split(".")[0]}) {
    Write-Host "Waiting for extract process (Process: $TargetFileName) to complete..  Next check in 10 seconds"
    Start-Sleep -seconds 10
}
    }
#Double Check & Set TS Var
if($inTS){
    if ((Test-Path $TargetFilePathName) -and (Test-Path "$extractedPath\$($TargetFileName.Split(".")[0])"))
    {
        Write-Output "Confirmed Download and setting TSVar LenovoBIOS"
        $tsenv.value('LenovoBIOS') = "Success"
    }
    $finalTSVar =  $tsenv.value('LenovoBIOS')
    Write-Output "Final TS var value $finalTSVar"
    Write-Output "___________________________________________________"
    }