param([string] $varITGKey,
      [string] $varPasswordID)

if (-not ([System.Management.Automation.PSTypeName]'ServerCertificateValidationCallback').Type)
{
$certCallback = @"
    using System;
    using System.Net;
    using System.Net.Security;
    using System.Security.Cryptography.X509Certificates;
    public class ServerCertificateValidationCallback
    {
        public static void Ignore()
        {
            if(ServicePointManager.ServerCertificateValidationCallback ==null)
            {
                ServicePointManager.ServerCertificateValidationCallback += 
                    delegate
                    (
                        Object obj, 
                        X509Certificate certificate, 
                        X509Chain chain, 
                        SslPolicyErrors errors
                    )
                    {
                        return true;
                    };
            }
        }
    }
"@
    Add-Type $certCallback
 }
[ServerCertificateValidationCallback]::Ignore()

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

if (Get-Module -ListAvailable -Name MSOnline) {
    Import-Module MSOnline
} else {
    Install-Module MSOnline -Force
}

$key = "$varITGKey"
$ITGbaseURI = "https://api.itglue.com"
$assettypeID = 107594
 
$headers = @{
    "x-api-key" = $key
}
 
Function Get-StringHash([String] $String, $HashName = "MD5") { 
    $StringBuilder = New-Object System.Text.StringBuilder 
    [System.Security.Cryptography.HashAlgorithm]::Create($HashName).ComputeHash([System.Text.Encoding]::UTF8.GetBytes($String))| % { 
        [Void]$StringBuilder.Append($_.ToString("x2")) 
    } 
    $StringBuilder.ToString() 
}
     
function Get-ITGlueItem($Resource) {
    $array = @()
 
    $body = Invoke-RestMethod -Method get -Uri "$ITGbaseUri/$Resource" -Headers $headers -ContentType application/vnd.api+json
    $array += $body.data
    Write-Host "Retrieved $($array.Count) items"
 
    if ($body.links.next) {
        do {
            $body = Invoke-RestMethod -Method get -Uri $body.links.next -Headers $headers -ContentType application/vnd.api+json
            $array += $body.data
            Write-Host "Retrieved $($array.Count) items"
        } while ($body.links.next)
    }
    return $array
}
 
$passwords = Get-ITGlueItem -Resource passwords/$varPasswordID 

$ITGluepasswords = @()

foreach($password in $passwords){
    $details = Get-ITGlueItem -Resource passwords/$($password.id) 
    if(($details.attributes.'password-category-name' -eq 'Office 365') -and ($password.id -eq "$varPasswordID")){

    $customers = [ordered]@{
        Customer        = $details.attributes.'organization-name'
        Category = $details.attributes.'password-category-name'
        Username = $details.attributes.username
        Password = $details.attributes.password
        PasswordID = $password.id      
    }
    $object = New-Object psobject -Property $customers
    $ITGluepasswords += $object

    }
}    

$o365user = $ITGluepasswords.username 
$o365pass = $ITGluepasswords.password 
$pass= convertto-securestring -string $o365pass -asplaintext -force
$mycred = new-object -typename System.Management.Automation.PSCredential -argumentlist $o365user,$pass
$O365Cred = Get-Credential $mycred
Connect-MsolService -Credential $O365Cred




function GetAllITGItems($Resource) {
    $array = @()
    
    $body = Invoke-RestMethod -Method get -Uri "$ITGbaseURI/$Resource" -Headers $headers -ContentType application/vnd.api+json
    $array += $body.data
    Write-Host "Retrieved $($array.Count) items"
        
    if ($body.links.next) {
        do {
            $body = Invoke-RestMethod -Method get -Uri $body.links.next -Headers $headers -ContentType application/vnd.api+json
            $array += $body.data
            Write-Host "Retrieved $($array.Count) items"
        } while ($body.links.next)
    }
    return $array
}
    
function CreateITGItem ($resource, $body) {
    $item = Invoke-RestMethod -Method POST -ContentType application/vnd.api+json -Uri $ITGbaseURI/$resource -Body $body -Headers $headers
    return $item
}
    
function UpdateITGItem ($resource, $existingItem, $newBody) {
    $updatedItem = Invoke-RestMethod -Method Patch -Uri "$ITGbaseUri/$Resource/$($existingItem.id)" -Headers $headers -ContentType application/vnd.api+json -Body $newBody
    return $updatedItem
}
    
function Build365TenantAsset ($tenantInfo) {
    
    $body = @{
        data = @{
            type       = "flexible-assets"
            attributes = @{
                "organization-id"        = $tenantInfo.OrganizationID
                "flexible-asset-type-id" = $assettypeID
                traits                   = @{
                    "tenant-name"      = $tenantInfo.TenantName
                    "tenant-id"        = $tenantInfo.TenantID
                    "initial-domain"   = $tenantInfo.InitialDomain
                    "verified-domains" = $tenantInfo.Domains
                    "licenses"         = $tenantInfo.Licenses
                    "licensed-users"   = $tenantInfo.LicensedUsers
                }
            }
        }
    }
    
    $tenantAsset = $body | ConvertTo-Json -Depth 10
    return $tenantAsset
}
    
    
    
$customer = Get-MsolCompanyInformation
    
$365domains = @()
    

    Write-Host "Getting domains for $($customer.DisplayName)" -ForegroundColor Green
    $companyInfo = Get-MSOLCompanyInformation | select objectID
    
    $customerDomains = Get-MsolDomain -TenantId $companyInfo.ObjectId | Where-Object {$_.status -contains "Verified"}
    $initialDomain = $customerDomains | Where-Object {$_.isInitial}
    $Licenses = $null
    $licenseTable = $null
    $Licenses = Get-MsolAccountSku -TenantId $customer.TenantId
    if ($licenses) {
        $licenseTableTop = "<br/><table class=`"table table-bordered table-hover`" style=`"width:600px`"><thead><tr><th>License Name</th><th>Active</th><th>Consumed</th><th>Unused</th></tr></thead><tbody><tr><td>"
        $licenseTableBottom = "</td></tr></tbody></table>"
        $licensesColl = @()
        foreach ($license in $licenses) {
            $licenseString = "$($license.SkuPartNumber)</td><td>$($license.ActiveUnits) active</td><td>$($license.ConsumedUnits) consumed</td><td>$($license.ActiveUnits - $license.ConsumedUnits) unused"
            $licensesColl += $licenseString
        }
        if ($licensesColl) {
            $licenseString = $licensesColl -join "</td></tr><tr><td>"
        }
        $licenseTable = "{0}{1}{2}" -f $licenseTableTop, $licenseString, $licenseTableBottom
    }
    $licensedUserTable = $null
    $licensedUsers = $null
    $licensedUsers = get-msoluser -TenantId $customer.TenantId -All | Where-Object {$_.islicensed} | Sort-Object UserPrincipalName
    if ($licensedUsers) {
        $licensedUsersTableTop = "<br/><table class=`"table table-bordered table-hover`" style=`"width:80%`"><thead><tr><th>Display Name</th><th>Addresses</th><th>Assigned Licenses</th></tr></thead><tbody><tr><td>"
        $licensedUsersTableBottom = "</td></tr></tbody></table>"
        $licensedUserColl = @()
        foreach ($user in $licensedUsers) {
           
            $aliases = (($user.ProxyAddresses | Where-Object {$_ -cnotmatch "SMTP" -and $_ -notmatch ".onmicrosoft.com"}) -replace "SMTP:", " ") -join "<br/>"
            $licensedUserString = "$($user.DisplayName)</td><td><strong>$($user.UserPrincipalName)</strong><br/>$aliases</td><td>$(($user.Licenses.accountsku.skupartnumber) -join "<br/>")"
            $licensedUserColl += $licensedUserString
        }
        if ($licensedUserColl) {
            $licensedUserString = $licensedUserColl -join "</td></tr><tr><td>"
        }
        $licensedUserTable = "{0}{1}{2}" -f $licensedUsersTableTop, $licensedUserString, $licensedUsersTableBottom
    
    
    }
        
        
    $hash = [ordered]@{
        TenantName        = $customer.displayname
        Domains           = $customerDomains.name
        TenantId          = $customer.TenantId
        InitialDomain     = $initialDomain.name
        Licenses          = $licenseTable
        LicensedUsers     = $licensedUserTable
    }
    $object = New-Object psobject -Property $hash
    $365domains += $object
        

    
# Get all organisations
#$orgs = GetAllITGItems -Resource organizations
    
# Get all Contacts
$itgcontacts = GetAllITGItems -Resource contacts
    
$itgEmailRecords = @()
foreach ($contact in $itgcontacts) {
    foreach ($email in $contact.attributes."contact-emails") {
        $hash = @{
            Domain         = ($email.value -split "@")[1]
            OrganizationID = $contact.attributes.'organization-id'
        }
        $object = New-Object psobject -Property $hash
        $itgEmailRecords += $object
    }
}
    
$allMatches = @()
foreach ($365tenant in $365domains) {
    foreach ($domain in $365tenant.Domains) {
        $itgContactMatches = $itgEmailRecords | Where-Object {$_.domain -contains $domain}
        foreach ($match in $itgContactMatches) {
            $hash = [ordered]@{
                Key            = "$($365tenant.TenantId)-$($match.OrganizationID)"
                TenantName     = $365tenant.TenantName
                Domains        = ($365tenant.domains -join ", ")
                TenantId       = $365tenant.TenantId
                InitialDomain  = $365tenant.InitialDomain
                OrganizationID = $match.OrganizationID
                Licenses       = $365tenant.Licenses
                LicensedUsers  = $365tenant.LicensedUsers
            }
            $object = New-Object psobject -Property $hash
            $allMatches += $object
        }
    }
}
    
$uniqueMatches = $allMatches | Sort-Object key -Unique
    
foreach ($match in $uniqueMatches) {
    $existingAssets = @()
    $existingAssets += GetAllITGItems -Resource "flexible_assets?filter[organization_id]=$($match.OrganizationID)&filter[flexible_asset_type_id]=$assetTypeID"
    $matchingAsset = $existingAssets | Where-Object {$_.attributes.traits.'tenant-id' -contains $match.TenantId}
        
    if ($matchingAsset) {
        Write-Host "Updating Office 365 tenant for $($match.tenantName)"
        $UpdatedBody = Build365TenantAsset -tenantInfo $match
        $updatedItem = UpdateITGItem -resource flexible_assets -existingItem $matchingAsset -newBody $UpdatedBody
    }
    else {
        Write-Host "Creating Office 365 tenant for $($match.tenantName)"
        $newBody = Build365TenantAsset -tenantInfo $match
        $newItem = CreateITGItem -resource flexible_assets -body $newBody
    }
}

