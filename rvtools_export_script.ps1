# =============================================================================================================
# Script:    RVTools_Report.ps1
# Date:      July, 2024
# Referance: https://www.vgemba.net/vmware/RVTools-Export/
# =============================================================================================================
<#
.SYNOPSIS
This script connects to multiple VMware vCenter environments, collects virtual machine inventory data using RVTools,
and exports the data to Excel files. The resulting files are then optionally merged and summarized using pandas,
filtered by selected columns, and saved to a structured output directory. Final reports can be sent via email
to designated recipients.
	
#>

[string] $RVToolsPath = "C:\Program Files (x86)\Robware\RVTools"

set-location $RVToolsPath

# =====================================
# RVTools TAV - A-site-vmwarename
# =====================================
[string] $VCServer = "x.x.x.x"                    
[string] $User = "username@vsphere.local"                                                    
[string] $EncryptedPassword = "_RVToolsV2encripytedpassword"
[string] $XlsxDir1 = "C:\RVTools"
[string] $XlsxFile1 = "A-site-vmwarename.xlsx"
[string] $XlsxFileoutput = "C:\RVTools\A-site-vmwarename.xlsx"


# Start cli of RVTools for A-site-vmwarename
Write-Host "Start export for vCenter $VCServer" -ForegroundColor DarkYellow
$Arguments = "-u $User -p $EncryptedPassword -s $VCServer -c ExportvInfo2xlsx -d $XlsxDir1 -f $XlsxFile1 -DBColumnNames -ExcludeCustomAnnotations"

Write-Host $Arguments
$Process = Start-Process -FilePath "C:\Program Files (x86)\Robware\RVTools\RVTools.exe" -ArgumentList $Arguments -NoNewWindow -Wait -PassThru

if($Process.ExitCode -eq -1)
{
    Write-Host "Error: Export failed! RVTools returned exitcode -1, probably a connection error! Script is stopped" -ForegroundColor Red
    exit 1
}
$OutputFile = $XlsxFileoutput


# =====================================
# RVTools TAV - B-site-vmwarename
# =====================================

[string] $VCServer = "x.x.x.x"                    
[string] $User = "username@vsphere.local"                                                    
[string] $EncryptedPassword = "_RVToolsV2encripytedpassword"
[string] $XlsxDir1 = "C:\RVTools"
[string] $XlsxFile1 = "B-site-vmwarename.xlsx"
[string] $XlsxFileoutput = "C:\RVTools\B-site-vmwarename.xlsx"

# Start cli of RVTools for B-site-vmwarename
Write-Host "Start export for vCenter $VCServer" -ForegroundColor DarkYellow
$Arguments = "-u $User -p $EncryptedPassword -s $VCServer -c ExportvInfo2xlsx -d $XlsxDir1 -f $XlsxFile1 -DBColumnNames -ExcludeCustomAnnotations"

Write-Host $Arguments

$Process = Start-Process -FilePath ".\RVTools.exe" -ArgumentList $Arguments -NoNewWindow -Wait -PassThru

if($Process.ExitCode -eq -1)
{
    Write-Host "Error: Export failed! RVTools returned exitcode -1, probably a connection error! Script is stopped" -ForegroundColor Red
    exit 1
}

$OutputFile = $XlsxFileoutput

# =====================================
# RVTools Send Mail
# =====================================

set path=%path%;"C:\Program Files (x86)\Robware\RVTools"

[string] $SMTPserver="relay-mail-server-ip"
[string] $SMTPport="25"
[string] $Mailto="emrahuludag@gmail.com"
[string] $Mailto="group-mail@domain.local"
[string] $Mailfrom="rvtools.report@domain.local"
[string] $Mailsubject="RVTools Inventory Report"
[string] $AttachmentDir="C:\rvtools\"
[string] $XlsxDir = "C:\RVTools\"
[string] $XlsxFile1 = "A-site-vmwarename.xlsx"
[string] $XlsxFile2 = "B-site-vmwarename.xlsx"

# =====================================
# Start RVTools Merge All files
# =====================================

$inputFiles = "$XlsxDir\$XlsxFile1;$XlsxDir\$XlsxFile2"

.\RVToolsMergeExcelFiles.exe -input "$inputFiles" -output "C:\rvtools\Rvtools_merged.xlsx" -overwrite -verbose

# ====================================
# Sending Mail
# ====================================
.\rvtoolssendmail.exe /smtpserver $SMTPserver /smtpport $SMTPport /mailto $Mailto /mailfrom $Mailfrom /mailsubject "RVTools Inventory Report Merged" /attachment C:\rvtools\Rvtools_merged.xlsx
.\rvtoolssendmail.exe /smtpserver $SMTPserver /smtpport $SMTPport /mailto $Mailto /mailfrom $Mailfrom /mailsubject "RVTools Inventory A-site-vmwarename.xlsx Report " /attachment $XlsxDir\$XlsxFile1
.\rvtoolssendmail.exe /smtpserver $SMTPserver /smtpport $SMTPport /mailto $Mailto /mailfrom $Mailfrom /mailsubject "RVTools Inventory B-site-vmwarename.xlsx Report " /attachment $XlsxDir\$XlsxFile2