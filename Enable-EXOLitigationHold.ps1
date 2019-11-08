#Requires -Version 5.1
<#PSScriptInfo

.VERSION 1.0.0

.GUID 6294d02e-207f-411b-a76e-1485011e98c5

.AUTHOR June Castillote

.COMPANYNAME lazyexchangeadmin.com

.COPYRIGHT Copyright (c) 2019 June Castillote

.TAGS

.LICENSEURI

.PROJECTURI https://github.com/junecastillote/Enable-EXOLitigationHold

.ICONURI

.EXTERNALMODULEDEPENDENCIES

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES

.PRIVATEDATA

#>

<#

.DESCRIPTION
 Script to enable litigation hold for all mailbox with ExchangeOnlineEnterprise Mailbox Plan

#>

[cmdletbinding()]
Param(
    [parameter()]
    [switch]$UseSSL,

    [parameter()]
    [switch]$SendEmail,

    [parameter()]
    [mailaddress]$From,

    [parameter()]
    [mailaddress[]]$To,

    [parameter()]
    [pscredential]$Credential,

    [parameter()]
    [string]$smtpServer,

    [parameter()]
    [int]$Port,

    [parameter()]
    [switch]$ListOnly,

    [parameter()]
    [string]$ReportDirectory = ($env:windir + '\temp'),

    [parameter()]
    [ValidateSet('CSV', 'HTML', 'ALL')]
    [string]$ReportType = 'ALL',

    [parameter()]
    [switch]$AttachCSVWhenPossible
)

if ($ListOnly) {
    Write-Output '[TEST MODE]'
}

if ($SendEmail) {
    if (!$From) {
        Write-Output "From address is missing."
    }

    if (!$To) {
        Write-Output "To address is missing."
    }

    if (!$smtpServer) {
        Write-Output "SMTP Server address is missing."
    }

    if (!$Port) {
        Write-Output "SMTP Server Port is missing."
    }
}

$ScriptInfo = (Test-ScriptFileInfo -Path "$($PSScriptRoot)\$($MyInvocation.MyCommand.Name)")
Write-Output "....................................."
Write-Output "Name      : $($ScriptInfo.Name)"
Write-Output "Version   : $($ScriptInfo.Version)"
Write-Output "....................................."
$tz = ([System.TimeZoneInfo]::Local).DisplayName.ToString().Split(" ")[0]
$today = Get-Date -Format "MMMM dd, yyyy hh:mm tt"

if (!(Test-Path $ReportDirectory)) {
    $null = New-Item -ItemType Directory -Path $ReportDirectory -Force
}

"Last Run: $today" | Out-File ($ReportDirectory + "\Remediate-Exchange-Online-Litigation-Hold.txt")

try {
    $OrgInfo = Get-OrganizationConfig -Erroraction stop
    $Organization = $OrgInfo.DisplayName
}
catch {
    Write-Output "Remote Exchange Online PowerShell Session is required."
    break
}

$css_string = @'
    #tbl
    {
        font-family:"Tahoma";
        width:auto;
        border-collapse:collapse;
    }
    #tbl td, #tbl th
    {
        font-size:13px;
        border-bottom: 1px solid #ccc;
        padding-top:10px;
        padding-bottom:10px;
        padding-left:10px;
        padding-right:10px;
    }
    #tbl td.head
    {
        font-size:14px;
        border: none;
        padding-top:10px;
        padding-bottom:10px;
        padding-left:10px;
        padding-right:10px;
    }
    #tbl th
    {
        font-size:14px;
        background-color:#fff;
        text-align:left;
        vertical-align:top;
        color: #7a7a52;
    }
    #tbl th.section
    {
        font-family:"Tahoma";
        font-weight: bold;
        font-size:26px;
        text-align:left;
        padding-top:10px;
        padding-bottom:10px;
        padding-left:10px;
        padding-right:10px;
        background-color:#fff;
        color:#000;
        vertical-align:center;
        border: none;
    }
    #tbl td
    {
        text-align:left;
        vertical-align:top;
        font-weight: lighter;
        color: #7a7a52;
        width: fit-content
    }
    #tbl td.wrap
    {
        word-break: break-all;
    }
    #legend
    {
        font-family:"Tahoma";
        width: auto;
        border-collapse:collapse;
        font-size:10px;
        text-align: center;
    }
    #legend td
    {
        font-size:10px;
        border: none;
        padding-top:5px;
        padding-bottom:5px;
        padding-left:5px;
        padding-right:5px;
    }
    #settings
    {
        font-family:"Tahoma";
        width: auto;
        border-collapse:collapse;
        font-size:10px;
        text-align: left;
    }
    #settings td
    {
        border: none;
        padding-top:0px;
        padding-bottom:0px;
        padding-left:0px;
        padding-right:0px;
        color: #7a7a52;
    }
    .red
    {
        background-color: red;
        color: #fff;
    }
    .green
    {
        background-color: green;
        color: #fff;
    }
    .gray
    {
        background-color: gray;
        color: #fff;
    }
'@

$subject = "Exchange Online Litigation Hold Remediation Report"

if ($ListOnly) {
    $subject = ('[TEST MODE] ' + $subject)
}

$fileSuffix = "{0:yyyy_MM_dd}" -f [datetime]$today
$outputCsvFile = ($ReportDirectory + "\$($Organization)-LitigationHold_Remediation_Report-$($fileSuffix).csv").Replace(" ", "_")
$outputHTMLFile = ($ReportDirectory + "\$($Organization)-LitigationHold_Remediation_Report-$($fileSuffix).html").Replace(" ", "_")

if (Test-Path $outputCsvFile) { $null = Remove-Item -Path $outputCsvFile -Force -Confirm:$false }
if (Test-Path $outputHTMLFile) { $null = Remove-Item -Path $outputHTMLFile -Force -Confirm:$false }

Write-Output 'Getting mailbox list with E3, E5 or Plan 2 License..'
## 1. Get all mailbox with plan
$mailboxList = Get-Mailbox -ResultSize Unlimited -Filter 'mailboxplan -ne $null -and litigationholdenabled -eq $false'
## 2. Get all mailbox with "ExchangeOnlineEnterprise*" plan
$mailboxList = $mailboxList | Where-Object { $_.MailboxPlan -like "ExchangeOnlineEnterprise*" }
Write-Output "Found $([int]$mailboxList.count) mailbox with disabled litigation hold"
## 3. Enable Litigation Hold
if ($mailboxList.count -gt 0) {
    if (!$ListOnly) {
        try {
            foreach ($mailbox in $mailboxList) {
                Set-Mailbox -Identity $mailbox.SamAccountName -LitigationHoldEnabled $true -WarningAction SilentlyContinue -ErrorAction STOP
                Write-Output "Enable Litigation Hold for Mailbox: $($Mailbox.Name) - SUCCESS"
            }
        }
        catch {
            Write-Output "Enable Litigation Hold for Mailbox: $($Mailbox.Name) - FAIL"
            Write-Output $_.Exception.Message
        }
    }
}
else {
    break
}

#Reset mailboxList and get updated values.
if (!$ListOnly) {
    Write-Output 'Getting updated values. Please wait...'
    $mailboxList = @($mailboxList | ForEach-Object { Get-Mailbox -Identity $_.Identity })
}

#if ($mailboxList.count -gt 0) {
Write-Output 'Writing report..'

## Create CSV report
if ($ReportType -eq 'CSV' -or $ReportType -eq 'ALL') {
    $mailboxList | Select-Object Name, UserPrincipalName, SamAccountName,
    @{Name = 'WhenMailboxCreated'; Expression = { '{0:dd/MMM/yyyy}' -f $_.WhenMailboxCreated } },
    LitigationHoldEnabled,
    LitigationHoldDate,
    LitigationHoldDuration,
    LitigationHoldOwner |
    Export-CSV -NoTypeInformation $outputCsvFile
}

## create the HTML report
## html title
$html = "<html><head><title>[$($Organization)] $($subject)</title><meta http-equiv=""Content-Type"" content=""text/html; charset=ISO-8859-1"" />"
$html += '<style type="text/css">'
$html += $css_string
$html += '</style></head><body>'

## heading
$html += '<table id="tbl">'
$html += '<tr><td class="head"> </td></tr>'
$html += '<tr><th class="section">' + $subject + '</th></tr>'
$html += '<tr><td class="head"><b>' + $Organization + '</b><br>' + $today + ' ' + $tz + '</td></tr>'
$html += '<tr><td class="head"> </td></tr>'
$html += '</table>'
$html += '<table id="tbl">'

if ($ReportType -ne 'CSV') {
    $html += '<tr><th>Name</th><th>UPN</th><th>Created</th><th>Enabled</th><th>When</th><th>Duration</th><th>Owner</th></tr>'

    foreach ($mailbox in $mailboxList) {
        $WhenMailboxCreated = '{0:dd-MMM-yyyy}' -f $mailbox.WhenMailboxCreated
        $LitigationHoldDate = '{0:dd-MMM-yyyy}' -f $mailbox.LitigationHoldDate
        ## data values
        $html += "<tr><td>$($mailbox.Name)</td><td>$($mailbox.UserPrincipalName)</td><td>$($WhenMailboxCreated)</td>"
        $html += "<td>$($mailbox.LitigationHoldEnabled)</td><td>$($LitigationHoldDate)</td>"
        $html += "<td>$($mailbox.LitigationHoldDuration)</td><td>$($mailbox.LitigationHoldOwner)</td></tr>"

        # if (!$ListOnly) {
        #     Set-Mailbox -Identity $mailbox.SamAccountName -LitigationHoldEnabled $true -WarningAction SilentlyContinue
        # }
    }
    $html += '</table>'
}
elseif ($ReportType -ne 'HTML') {
    $html += '<tr><th>Please see attached CSV file</th></tr>'
    $html += '</table>'
}

$html += '<table id="tbl">'
$html += '<tr><td class="head"> </td></tr>'
$html += '<tr><td class="head"> </td></tdr>'
$html += '<tr><td class="head">Source: ' + $env:COMPUTERNAME + '<br>'
$html += 'Script Directory: ' + (Resolve-Path $PSScriptRoot).Path + '<br>'
$html += 'Report Directory: ' + (Resolve-Path $ReportDirectory).Path + '<br>'
$html += '<a href="' + $ScriptInfo.ProjectURI.ToString() + '">' + $ScriptInfo.Name.ToString() + ' v' + $ScriptInfo.Version.ToString() + ' </a><br>'
$html += '<tr><td class="head"> </td></tr>'
$html += '</table>'
$html += '</html>'
$html | Out-File $outputHTMLFile -Encoding UTF8
Write-Output "Report saved in $ReportDirectory"

if ($sendEmail -eq $true) {
    Write-Output 'Sending email..'
    [string]$html = Get-Content $outputHTMLFile -Raw -Encoding UTF8

    $mailParams = @{
        SmtpServer                 = $smtpServer
        Port                       = $Port
        To                         = $To
        From                       = $From
        Subject                    = "[$($Organization)] $subject"
        DeliveryNotificationOption = 'OnFailure'
        BodyAsHTML                 = $true
        Body                       = (Get-Content $outputHTMLFile -Raw -Encoding UTF8)
    }

    if ($Credential) { $mailParams += @{Credential = $Credential } }
    if ($UseSSL) { $mailParams += @{UseSSL = $true } }

    ## if ReportType is CSV only, attach the CSV file.
    if ($ReportType -ne 'HTML' -or $AttachCSVWhenPossible) {
        if (Test-Path $outputCsvFile) {
            $mailParams += @{Attachments = $outputCsvFile }
        }
    }

    # ## if AttachCSVWhenPossible and ReportType is CSV or ALL
    # if ($AttachCSVWhenPossible -and ($ReportType -eq 'All' -or $ReportType -eq 'CSV')) {
    #     $mailParams += @{Attachments = $outputCsvFile }
    # }

    ## Send email
    try {
        Send-MailMessage @mailParams -ErrorAction STOP
    }
    catch {
        Write-Output $_.Exception.Message
    }
}

## if CSV only, delete HTML file
if ($ReportType -eq 'CSV') {
    Remove-Item -Path $outputHTMLFile -Confirm:$false -Force
}
#}
