param (
  [parameter(Mandatory=$true)]
  [string]$ADGroup
)

#Requires -RunAsAdministrator
#This allows use of TLSv12
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

#these lines should be run on the local server (where the script is running).
#When you download the exe, put it in a path that other servers can access 
$downloadUrl = 'https://update.code.visualstudio.com/latest/win32-x64/stable'
# grab latest vscode version number from Content-Disposition header and remove all text but number
$latestVer = (((Invoke-WebRequest $downloadUrl).Headers.'Content-Disposition').Split('"')[1]).Replace("VSCodeSetup-x64-", "").Replace(".exe", "")
#This is where the update exe will be saved
$updateFilepath = 'C:\Windows\Temp\vscode-stable.exe'
#Gets the names of all the computers in the specified Active Directory Group
$computers = Get-ADGroupMember -Identity "$ADGroup" | Select-Object -ExpandProperty Name
#Creates an empty array that outdated computers will be added to if there are any
$outdatedComputers =@()

#Checks the version of vscode on each computer in the ADGroup and compares it to the latest version of vscode 
foreach($computer in $computers) {
  $version = Invoke-Command -ComputerName "$computer" -ScriptBlock {(code -v)[0]}
  #If the latest version if vscode is greater than the currently installed version, the computer will be added to the array outdatedComputers
  if ($version -And $latestVer -gt $version) {
    Write-Host "$computer needs an update..."
    $outdatedComputers += "$computer"
  }
}

#Checks if the array outdatedComputers has any computers in it
#If outdatedComputers is empty, all computers have the latest version of vscode
if($outdatedComputers) {
  $exeArgs = '/verysilent /tasks=addtopath,addcontextmenufiles,addcontextmenufolders'
  Write-Output "The following computers have an outdated version of VSCode: $outdatedComputers"
  #Downloads the vscode update exe and saves it to updateFilepath
  Invoke-WebRequest -Uri $downloadUrl -OutFile $updateFilepath
  #Runs the update exe on each computer that has an outdated version of vscode
  foreach($computer in $outdatedComputers) {
    $codeRunning = @(Get-Process -Name Code -ComputerName $computer -ErrorAction Ignore).Count -gt 0
    #If the computer being updated is the computer the script is being run on, the script will not need to copy the update exe
    if($computer -eq $env:ComputerName) {
      #$codeRunning = @(Get-Process -name Code -ErrorAction Ignore).Count -gt 0
      if ($codeRunning) {
        Write-Host "Stopping VSCode on $computer..."
        Stop-Process -Name Code
      }
      Write-Host "Updating $computer..."
      #Runs the exe on the local computer
      Start-Process -FilePath $updateFilepath -ArgumentList $exeArgs
    }
    else{
      #$codeRunning = Invoke-Command -ComputerName $computer -ScriptBlock {@(Get-Process -name Code -ErrorAction Ignore).Count} -gt 0
      if ($codeRunning) {
        Write-Host "Stopping VSCode on $computer..."
        Invoke-Command -ComputerName $computer -ScriptBlock {Get-Process -Name Code | Stop-Process}
      }
      Write-Host "Copying update file..."
      #Copys the update exe from the local machine to the admin share of the computer being updated
      Copy-Item -Path $updateFilepath -Destination "\\$computer\\c$\windows\temp\vscode-stable.exe"
      Write-Host "Updating $computer..."
      #Runs the exe on the remote computer
      Invoke-Command -ComputerName $computer -ScriptBlock {Start-Process -FilePath $Using:updateFilepath -ArgumentList $Using:exeArgs -Wait}
    }
    Write-Host "$computer successfully updated"
  }
  Write-Host "All computers have successfully been updated"
}
else {
  Write-Output "All computers have the latest version of VSCode"
}
