$params = @{
    Credential = $Credential ## Get-Credential
    SendEmail = $false
    ListOnly = $true
    From = 'mailer365@poshlab.ml'
    To = 'june@poshlab.ml'
    smtpServer = 'smtp.office365.com'
    Port = 587
    UseSSL = $true
    ReportDirectory = "$PSScriptRoot\Reports"
    ReportType = 'ALL' ## CSV, HTML, ALL
    AttachCSVWhenPossible = $true
}

.$PSScriptRoot\Enable-LitigationHold.ps1 @params