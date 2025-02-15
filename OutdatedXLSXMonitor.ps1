﻿############################################################
#                                                          #
#  Script: OutdatedXLSXMonitor                             #
#  Last Updated: 2025-02-05                  Version 1.1   #
#  Written by: Fabio Siqueira                              #
#                                                          #
############################################################
# PowerShell Script to monitor outdated .csv, .txt, .xlsx files.

############################################################

# Settings
$DirectoryToScan = "C:\..." # Directory to Scan
$TimeThresholdHours = 1 # 1 Hour

$MailTo = "User@Mail.com" # User Sending the Mail Message
$MailFrom = "User@Mail.com" # User Receiving the Mail Message

$SMTPServer = "SMTP.Server.com" # SMTP Server in use
#$Port = 25
#$Port = 465
#$Port = 587


# Verify Files
Function Get-OutdatedFiles {
    param (
        [string]$Directory,
        [int]$TimeThreshold
    )

    $Now = Get-Date
    $TimeThresholdSpan = New-TimeSpan -Hours $TimeThreshold
    $OutdatedFiles = @()

    Get-ChildItem -Path $Directory -Recurse -Include "*.csv", "*.txt", "*.xlsx" | ForEach-Object {
        $LastModified = $_.LastWriteTime
        $TimeDifference = $Now - $LastModified

        if ($TimeDifference -gt $TimeThresholdSpan) {
            $OutdatedFiles += [PSCustomObject]@{
                FileName = $_.FullName
                LastModified = $LastModified
            }
        }
    }

    return $OutdatedFiles
}


# Message Settings
    $Message = New-Object System.Net.Mail.MailMessage
    $Message.From = $MailFrom
    $Message.To.Add($MailTo)
    $Message.Subject = $Subject
    $Message.Body = $Body
    $Message.IsBodyHtml = $false


# Running Script
Write-Host "Starting file monitoring..."
$OutdatedFiles = Get-OutdatedFiles -Directory $DirectoryToScan -TimeThreshold $TimeThresholdHours

if ($OutdatedFiles.Count -gt 0) {
    Write-Host "$($OutdatedFiles.Count) outdated file(s) found. Sending alert..."

    $Body = "The following files have not been updated for more than $TimeThresholdHours hours:`n`n"
    $OutdatedFiles | ForEach-Object {
        $Body += "- $($_.FileName) (Last modified: $($_.LastModified))`n"
    }

    Send-MailMessage -From $MailFrom -To $MailTo -Subject "Alert: files not updated" -Body $Body -UseSsl -SmtpServer $SMTPServer 
} else {
    Write-Host "All files are up to date."
}
