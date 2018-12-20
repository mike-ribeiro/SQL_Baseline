#--------------------------------------------------------------------------------------------------------------
#-- NAME: 
#-- AUTHOR: Michael Ribeiro, Microsoft (miribeir@microsoft.com)
#-- DATE: 
#--
#--
#-- COMMENTS: This script is a sample of how to automate the upload of performance counters into a SQL database. 
#--   Creates a backup folder and then archives the performance counters
#-- 
#-- PRE-REQS: 
#--   Database SQLBaseline needs to exist
#--   Machine that runs this script needs to have access to the SQL server hosting the SQLSaeline db
#--   
#--
#-- DISCLAIMER: This Sample Code is provided for the purpose of illustration only and is not intended to be used
#-- in a production environment.  THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY
#-- OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY 
#-- AND/OR FITNESS FOR A PARTICULAR PURPOSE.  We grant You a nonexclusive, royalty-free right to use and modify the
#-- Sample Code and to reproduce and distribute the object code form of the Sample Code, provided that 
#-- You agree: (i) to not use Our name, logo, or trademarks to market Your software product in which the Sample Code 
#-- is embedded; (ii) to include a valid copyright notice on Your software product in which the Sample Code is embedded; 
#-- and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims or lawsuits, 
#-- including attorneys’ fees, that arise or result from the use or distribution of the Sample Code.
#--
#------------------------------------------------------------------------------------------------------------------

#-- Folders and File locations --#

$sourceDir = "C:\PerfLogs\data\MS"
$destFolder = "C:\PerfLogs\data\MS\backups\" 
$sharedFolder = "Temp"
$destination = "$destFolder$(get-date -f yyyyMMddHHmmss).zip" 
$baselineCounters = "$SourceDir\Filtered.txt"

#uncomment below as this is specific for NBS environment
#[array]$PaymentsServers = ("Server1", "Server2", "Server3")



if((dir $sourceDir -Filter *".blg" | measure).Count -gt 0)
{

Write-Host "Checking for an ODBC connection..."
#Check if the DSN exists
if (!(Test-Path 'HKCU:\SOFTWARE\ODBC\ODBC.INI\PMCSQLBaseline'))
{
    Write-Host "ODBC connection not found. Creating a new DSN..."
    #Creating a DSN connection
    $HKCUPath1 = "HKCU:\SOFTWARE\ODBC\ODBC.INI\PMCSQLBaseline"
    $HKCUPath2 = "HKCU:\SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources\"

    md $HKCUPath1 -ErrorAction Stop
    Set-ItemProperty -path $HKCUPath1 -name Driver -Value (Get-ItemProperty "HKLM:\SOFTWARE\Wow6432Node\ODBC\ODBCINST.INI\SQL Server").Driver
    Set-ItemProperty -path $HKCUPath1 -name Description -Value PMCSQLBaseline
    Set-ItemProperty -path $HKCUPath1 -name Server -Value $env:COMPUTERNAME
    Set-ItemProperty -path $HKCUPath1 -name LastUser -Value $env:USERNAME
    Set-ItemProperty -path $HKCUPath1 -name Trusted_Connection -Value "Yes"
    Set-ItemProperty -path $HKCUPath1 -name Database -Value SQLBaseline

    Set-ItemProperty -path $HKCUPath2 -name "PMCSQLBaseline" -Value "SQL Server"

    Write-Host "ODBC connection created"
}

Write-Host "Using PMCSQLBaseline connection to upload data..."

Add-Type -AssemblyName "System.io.compression.filesystem"


if(!(Test-Path $destFolder))
{
    New-Item -ItemType Directory -Force -Path $destFolder
}

if(Test-Path $destination) { Remove-Item $destination }

#Get the zip archive to update with files. 
#This will create a Zip file if it doesn't already exist
$Archive = [System.IO.Compression.ZipFile]::Open( $destination, "Update" )



#We are now going to upload each collected BLG file into the SQL Baseline DB.
#For this we'll be using the relog method to filter out the necessary performance counters
#for the SQL Baseline
[array]$PaymentsServers = ("<local server>")
#Get-ChildItem -Path $sourceDir -Filter *".blg" | ForEach-Object { 
#   $fileName = $_.Name
#   Write-Host "Processing $fileName into the DB..."
Foreach($server in $PaymentsServers){     
    Write-Host "Processing following server:  $server"
    $processFolder = $server+$sharedFolder
    Write-Host "Processing following folder:  $processFolder"
  Get-ChildItem -Path $processFolder -Filter *".blg" | Where-Object {$_.CreationTime -gt (get-date).AddHours(-24)} | ForEach-Object { 
    $fileName = $_.FullName
    Write-Host "Processing $fileName into the DB..."


   #Call relog to upload the filtered BLG file into the DB   
    $result = Invoke-Command -ScriptBlock {    

   RELOG $_.FullName -cf $baselineCounters -f SQL -o SQL:PMCSQLBaseline!LatestRun   
    
   } -ErrorAction SilentlyContinue
   
   #After sucessfull upload to the DB we can know zip and archive the files
   remove $null = [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($Archive, $_.FullName, $fileName, [System.IO.Compression.CompressionLevel]::Optimal)

  remove$info = ConvertFrom-String($result | Out-String)
  removeif($info.P15 -eq 'Error:'){ Break}

  #Delete the file that has been processed into the DB
  removeRemove-item $_.FullName;
   
  } -ErrorAction Stop
}
$Archive.Dispose()
Get-Item $destination

}