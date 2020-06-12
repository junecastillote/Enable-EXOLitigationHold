# This helper script assumes that you will use Office 365 SMTP Relay to send the email report.

## Get Office 365 Credential
## $Credential = Get-Credential

$params = @{
    Credential = $Credential
    SendEmail = $false
    ListOnly = $true
    ExclusionList = (Get-Content "$PSScriptRoot\ExclusionList.txt")
    ## The FROM address must be an existing Office 365 mailbox.
    ## And you credential must have SendAs permission to use it.
    #From = 'mailer365@poshlab.ml'
    #To = 'june@poshlab.ml'
    #smtpServer = 'smtp.office365.com'
    #Port = 587
    #UseSSL = $true
    ReportDirectory = "$PSScriptRoot\Reports"
    ## Valid values are: CSV, HTML, ALL
    #ReportType = 'ALL'
    #AttachCSVWhenPossible = $true
}

## This assumes that the script is in the same working directory.
## If not, then replace $PSScriptRoot with the actual path
## Or if the script is installed and/or the location is added in your $env:PATH,
## then omit the location and just use the script name - Remediate-LitigationHold.ps1 @params
."$PSScriptRoot\Remediate-LitigationHold.ps1" @params
