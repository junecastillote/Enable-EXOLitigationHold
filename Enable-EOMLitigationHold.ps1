
Function Enable-EOMLitigationHold {


    [cmdletbinding()]
    Param(
        [parameter(Mandatory)]
        [pscredential]$Credential,    

        [parameter()]
        [switch]$SendEmail,

        [parameter()]
        [mailaddress]$From,

        [parameter()]
        [mailaddress[]]$To,

        [parameter()]
        [switch]$ListOnly,

        [parameter()]
        [string]$ReportDirectory = ($env:windir + '\temp'),

        [parameter()]
        [switch]$SkipConnect

    )

    if ($SendEmail) {
        if (!$From) {
            Write-Host "From address is missing."
        }

        if (!$To) {
            Write-Host "To address is missing."
        }
    }

    if (!$SkipConnect) {
        #discard all PSSession
        Get-PSSession | Remove-PSSession -Confirm:$false

        #create new Exchange Online Session
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $credential -Authentication Basic -AllowRedirection -WarningAction SilentlyContinue
        Import-PSSession $Session -DisableNameChecking | Out-Null
    }

    $Moduleinfo = (Test-ModuleManifest -Path $PSScriptRoot\Remediate-Exchange-Online-Litigation-Hold.psd1)
    $tz = ([System.TimeZoneInfo]::Local).DisplayName.ToString().Split(" ")[0]
    $today = Get-Date -Format "MMMM dd, yyyy hh:mm tt"

    $Organization = (Get-OrganizationConfig).DisplayName
    $css_string = Get-Content ($PSScriptRoot + '\style.css') -Raw

    $subject = "Exchange Online Mailbox Litigation Hold Remediation Report"
    $smtpServer = "smtp.office365.com"
    $smtpPort = "587"

    $fileSuffix = "{0:yyyy_MM_dd}" -f [datetime]$today
    $outputCsvFile = ($ReportDirectory + "\$($Organization)-LitigationHold_Remediation_Report-$($fileSuffix).csv").Replace(" ", "_")
    $outputHTMLFile = ($ReportDirectory + "\$($Organization)-LitigationHold_Remediation_Report-$($fileSuffix).html").Replace(" ", "_")
    $mailboxList = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox, SharedMailbox -Filter { LitigationHoldEnabled -eq $false }
    $mailboxList = $mailboxList | Where-Object { $_.MailboxPlan -like "ExchangeOnlineEnterprise*" }
    Write-Host "Found $($mailboxList.count) mailbox"

    if ($mailboxList.count -gt 0) {
        $mailboxList | Select-Object Name, UserPrincipalName, SamAccountName, @{Name = 'WhenMailboxCreated'; Expression = { '{0:dd/MMM/yyyy}' -f $_.WhenMailboxCreated } } | Export-CSV -NoTypeInformation $outputCsvFile
    
	
        #create the HTML report
        #html title
        $html = "<html><head><title>[$($Organization)] $($subject)</title><meta http-equiv=""Content-Type"" content=""text/html; charset=ISO-8859-1"" />"
        $html += '<style type="text/css">'
        $html += $css_string
        $html += '</style></head><body>'
		
        #heading
        $html += '<table id="tbl">'
        $html += '<tr><td class="head"> </td></tr>'
        $html += '<tr><th class="section">' + $subject + '</th></tr>'
        $html += '<tr><td class="head"><b>' + $Organization + '</b><br>' + $today + ' ' + $tz + '</td></tr>'
        $html += '<tr><td class="head"> </td></tr>'
        $html += '</table>'
        $html += '<table id="tbl">'
        $html += '<tr><th>Name</th><th>UPN</th><th>Mailbox Created Date</th></tr>'
		
        foreach ($mailbox in $mailboxList) {	
            $mailboxCreateDate = '{0:dd-MMM-yyyy}' -f $mailbox.WhenMailboxCreated
            #data values
            $html += "<tr><td>$($mailbox.Name)</td><td>$($mailbox.UserPrincipalName)</td><td>$($mailboxCreateDate)</td></tr>"
		
            if (!$ListOnly) {
                Set-Mailbox -Identity $mailbox.SamAccountName -LitigationHoldEnabled $true
            }
        }
        $html += '</table>'

        $html += '<table id="tbl">'
        $html += '<tr><td class="head"> </td></tr>'
        $html += '<tr><td class="head"> </td></tr>'
        $html += '<tr><td class="head">Source: ' + $env:COMPUTERNAME + '<br>'
        $html += 'Script Directory: ' + (Resolve-Path $PSScriptRoot).Path + '<br>'
        $html += 'Report Directory: ' + (Resolve-Path $ReportDirectory).Path + '<br>'
        $html += '<a href="' + $ModuleInfo.ProjectURI.ToString() + '">' + $ModuleInfo.Name.ToString() + ' v' + $ModuleInfo.Version.ToString() + ' </a><br>'
        $html += '<tr><td class="head"> </td></tr>'
        $html += '</table>'
        $html += '</html>'
        $html | Out-File $outputHTMLFile -Encoding UTF8
        Write-Host "Report saved in $($outputHTMLFile)"
	
        if ($sendEmail -eq $true) {
            [string]$html = Get-Content $outputHTMLFile -Raw -Encoding UTF8
            Send-MailMessage -SmtpServer $smtpServer -Port $smtpPort -To $Tos -From $From -Subject "[$($Organization)] $subject" -Body $html -BodyAsHTML -Credential $onlineCredential -UseSSL
        }
    }
}
