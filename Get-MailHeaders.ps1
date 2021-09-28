# this is a simpleish script to use the graph api to collect the headers from an email
# knowing the message-id of the email and mailbox where the email is sitting
#
# Usage: .\Get-MailHeaders.ps1 -MessageId "<interesting-message-id@domain.org>" -Mailbox "upn-of-mailbox@domain.org" -Tenant domain.onmicrosoft.com -Format Table
[CmdletBinding()]
param (
    [Parameter(Mandatory=$True)]
    [string]
    $MessageId,
    [Parameter(Mandatory=$True)]
    [string]
    $Mailbox,
    [Parameter(Mandatory=$False)]
    #[ValidateSet('tenant1', 'tenant2', 'tenant3')]
    [String[]]
    $Tenant = "contoso", # the bit before .onmicrosoft.com, e.g. contoso.onmicrosoft.com
    [Parameter(Mandatory=$False)]
    [ValidateSet('Table', 'GridView', 'List', 'Raw')]
    [String[]]
    $Format = "Table"
)

if(-not (Get-InstalledModule -Name Microsoft.Graph -ErrorAction SilentlyContinue)){
    Write-Host "Installing Microsoft.Graph modules..."
    Install-Module -Name Microsoft.Graph -Scope CurrentUser
}

if(Get-Command Graph){
    # this is a function I setup just to simplify the connection
    Graph -tenant $tenant
} else {
    Import-Module Microsoft.Graph
    $splat = @{
        # these should be changed to values relevant to your app/tenant
        $clientid = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"     # the clientid of an app with the permissions needed
        $certhumb = "0000000000000000000000000000000000000000" # thumbprint of a certificate credential used with the app
        $tenantid = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"     # the tenant of the app, can be contoso.onmicrosoft.com format
    }
    Try {
        Connect-MgGraph @splat
    } Catch { # laziest error catching ever!
        Connect-MgGraph -UseDeviceAuthentication
    }
}

# this gets the headers we're looking for
$headers = Get-MgUserMessage -UserId $Mailbox -Filter "InternetMessageId eq '$MessageId'" -Property InternetMessageHeaders | Select-Object -ExpandProperty InternetMessageHeaders

# thouse outputs the headers in a nice way
# call using the -Raw switch if you want the output as a variable
switch ($Format) {
    "Table" {$headers | Format-Table -Wrap}
    "GridView" {$headers | Out-GridView}
    "List" {$headers | Format-List}
    "Raw" {$headers}
}
