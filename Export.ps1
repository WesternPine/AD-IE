###############################################################################
# Export ADLDS Data via Direct LDAP Connection (No ADWS)
###############################################################################

# Test Data
$Target = "LDAP://localhost:636/"
$Path   = "CN=root,DC=dc,DC=local"  # Base container we know exists

# Define LDAP Connection (LDAP:// instead of ADWS-dependent Server parameter)
$Target = Read-Host -Promp "Server & Port (ex: localhost:636)"
$Secure = Read-Host -Promp "Use SSL/TLS? (ex: T/F)"
$Path = Read-Host -Promp "Directory Path (ex: CN=root,DC=dc,DC=local)"

$Protocol = "LDAP://"
If(-not $Target.ToUpper().StartsWith($Protocol)) {
	$Target = "$Protocol$Target"
}

$Dir = "/"
If(-not $Target.EndsWith($Dir)) {
	$Target = "$Target$Dir"
}

$BaseContainer = "$Target$Path"


# Gather Credentials (for authentication)
$Username   = Read-Host -Prompt "Username"
$Password   = Read-Host -Prompt "Password"

if($Secure.ToUpper().StartsWith("T")) {
	$authType = [System.DirectoryServices.AuthenticationTypes]::SecureSocketsLayer -bor [System.DirectoryServices.AuthenticationTypes]::Secure
} else {
	$authType = [System.DirectoryServices.AuthenticationTypes]::Secure
}

Write-Host "Using:"
Write-Host " - Target: $BaseContainer"
Write-Host " - Username: $Username"
Write-Host " - Password: $Password"
Write-Host " - AuthType: $authType"
Write-Host ""
Write-Host "This may take a minute, Connecting..."

# Test Base container connection
$ldapConnection = New-Object System.DirectoryServices.DirectoryEntry($BaseContainer, $Username, $Password, $authType)
if($ldapConnection -eq $null -or $ldapConnection.Children -eq $null) {
    Write-Host "Unable to connect to container."
    return
}

Write-Host "Successfully connected to container!"
$Export = Read-Host -Prompt "Begin Export? (Y/N)"

if($Export.ToUpper().StartsWith("N")) {
	Write-Host "Canceling Export process."
	return
}
Write-Host "Beginning Export..."
Start-Sleep -Milliseconds 3000

$ExcludedAttributes = @(
    # "objectGUID", "objectSid", "whenCreated", "whenChanged", "uSNChanged",
    # "nTSecurityDescriptor", "thumbnailPhoto", "pwdLastSet", "lastLogonTimestamp",
    # "badPasswordTime", "uSNCreated", "CanonicalName", "ProtectedFromAccidentalDeletion",
    # "instanceType", 
	"ObjectClass",
	# "ObjectCategory", "msDS-UserAccountDisabled",
    # "userPrincipalName", "badPwdCount", "DisplayName", "givenName", "sDRightsEffective",
    # "Name", "modifyTimeStamp", "Modified", "lockoutTime", "Created",
    # "dSCorePropagationData", "CN", "createTimeStamp", 
	"DistinguishedName"
	# ,
    # "PropertyCount", "PropertyNames"
)

# Function to retrieve LDAP objects
function Get-LdapObjects {
    param (
        [System.DirectoryServices.DirectorySearcher]$Searcher
    )
    
    $results = $Searcher.FindAll()
    $objects = @()
    
    foreach ($result in $results) {
        $entry = $result.GetDirectoryEntry()
        
        $obj = New-Object PSObject -Property @{
            DistinguishedName = $entry.distinguishedName[0]
            ObjectClass = $entry.objectClass[$entry.objectClass.Count-1]
        }
        
        foreach ($prop in $entry.Properties.PropertyNames) {
            if ($ExcludedAttributes -notcontains $prop) {
                $obj | Add-Member -MemberType NoteProperty -Name $prop -Value $entry.Properties[$prop].Value -Force
            }
        }
        
        $objects += $obj
    }
    return $objects
}

# Create DirectorySearcher instance
try {
    $Searcher = New-Object System.DirectoryServices.DirectorySearcher($ldapConnection)
    $Searcher.SearchScope = [System.DirectoryServices.SearchScope]::Subtree
    $Searcher.Filter = "(objectClass=*)"
    Write-Host "Retrieving objects..."
    
    $ExportedObjects = Get-LdapObjects -Searcher $Searcher
    
    if ($ExportedObjects.Count -eq 0) {
        Write-Host "No objects found. Exiting."
        return
    }
    
    Write-Host "Objects retrieved: $($ExportedObjects.Count)"
    
    # Convert to JSON and export
    Write-Host "Converting to JSON..."
    $JSON = $ExportedObjects | ConvertTo-Json -Depth 10
    
    Write-Host "Writing to file..."
    $JSON | Out-File "ADLDS_Export.json"
    
    Write-Host "Export completed successfully!"
} catch {
    Write-Host "[ERROR] Failed to retrieve objects: $($_.Exception.Message)"
}
