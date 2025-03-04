###############################################################################
# Import ADLDS Data from JSON and Create Objects via Direct LDAP
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
$import = Read-Host -Prompt "Begin import? (Y/N)"

if($import.ToUpper().StartsWith("N")) {
	Write-Host "Canceling import process."
	return
}
Write-Host "Beginning Import..."
Start-Sleep -Milliseconds 3000

# Read JSON
$ImportedObjects = Get-Content "ADLDS_Export.json" | ConvertFrom-Json

###############################################################################
# Function: retrieve everything after the first comma in DN
# (Simple approach; watch out for escaped commas if that might happen)
###############################################################################
function Get-ParentDN($distinguishedName) {
    $idx = $distinguishedName.IndexOf(',')
    if ($idx -lt 0) {
        return $null
    }
    return $distinguishedName.Substring($idx + 1)
}

###############################################################################
# Function: Attempt to create any objects that can be created (parents exist)
# Returns: [int] how many objects got created on this call
###############################################################################
function TryCreateObjects {
    param(
        [System.Collections.Generic.HashSet[string]]$CreatedDNs,
        [System.Object[]]$ObjectsToCreate,
        [string]$Target,
        [string]$Username,
        [string]$Password,
        [System.DirectoryServices.AuthenticationTypes]$AuthType,
        [string]$BaseDN   # e.g. $Path
    )

    $createdCount = 0

    foreach($obj in $ObjectsToCreate) {
		Write-Host "Checking Object:" $obj.DistinguishedName
		
        # Skip if already created or if it is the base container we know is present
        if ($CreatedDNs.Contains($obj.DistinguishedName)) {
			Write-Host "   - (Skipping)" $obj.DistinguishedName "was previously created."
            continue
        }
		
        # We skip the base container (already exists)
        if ($obj.DistinguishedName -eq $BaseDN) {
			Write-Host "   - (Skipping)" $obj.DistinguishedName "Is the base DN: $BaseDN"
            continue
        }

        # Figure out the parent's DN
        $parentDn = Get-ParentDN $obj.DistinguishedName
		Write-Host "   - Parent DN:" $parentDn
        if (-not $parentDn) {
            # If there's no parent portion, we can still attempt to create it
            # (uncommon scenario unless the DN is something like "CN=someTopLevel")
            $canCreate = $true
			Write-Host "   - (CanCreate =" $canCreate ") -not parentDn."
        }
        elseif ($parentDn -eq $BaseDN) {
            # If the parent is exactly the known base container
            $canCreate = $true
			Write-Host "   - (CanCreate =" $canCreate ") known base container."
        }
		
		
        # Attempt to bind to parent so we can add the child
		$parent = $null
        $parentPath = "$Target$parentDn"
		try {
			
            $parentEntry = New-Object System.DirectoryServices.DirectoryEntry(
                $parentPath, $Username, $Password, $AuthType
            )
			
            if($parentEntry -eq $null -or $parentEntry.Children -eq $null) {
                throw "Parent container unbindable: $parentDn"
            }
			
			$parent = $parentEntry
            $canCreate = $true
			Write-Host "   - (CanCreate =" $canCreate ") parent is created."
		} catch {
            Write-Host "   [ERROR] Could not bind to parent [$($parentDn)]: $($_.Exception.Message)"
			$canCreate = false
        }

        if (-not $canCreate) {
            continue
        }
        try {
            # RDN is everything up to the first comma
            # e.g., "CN=Child,OU=Whatever,DC=x,DC=y" => "CN=Child"
            $splitIndex = $obj.DistinguishedName.IndexOf(',')
            $rdn = $obj.DistinguishedName.Substring(0, $splitIndex)

            Write-Host "Creating object: $($obj.DistinguishedName)($rdn)"

            # Now add child
            $newAdObj = $parent.Children.Add($rdn, $obj.ObjectClass)
            $newAdObj.CommitChanges()

            # Mark as created
            $CreatedDNs.Add($obj.DistinguishedName) | Out-Null
            $createdCount++
            Write-Host "   -> Created."
        }
        catch {
            Write-Host "   [ERROR] Could not create [$($obj.DistinguishedName)]: $($_.Exception.Message)"
        }
    }

    return $createdCount
}

###############################################################################
# Build the objects list, skipping the base container from the main set
###############################################################################
$ObjectsToCreate = $ImportedObjects
$CreatedDNs      = New-Object System.Collections.Generic.HashSet[string]

###############################################################################
# Keep calling TryCreateObjects until it doesn't create anything new
###############################################################################
do {
    $created = TryCreateObjects -CreatedDNs $CreatedDNs `
                                -ObjectsToCreate $ObjectsToCreate `
                                -Target $Target `
                                -Username $Username `
                                -Password $Password `
                                -AuthType $authType `
                                -BaseDN $Path

    if($created -gt 0) {
        Write-Host "$created objects created in this iteration."
    }
} while($created -gt 0)

# Any objects left not in $CreatedDNs are considered uncreatable
$Remaining = $ObjectsToCreate | Where-Object {
    -not $CreatedDNs.Contains($_.DistinguishedName) `
    -and $_.DistinguishedName -ne $Path
}
if($Remaining.Count -gt 0) {
    Write-Host "Some objects could not be created because their parents never became available. Skipping these:"
    foreach($r in $Remaining) {
        Write-Host "   -> $($r.DistinguishedName)"
    }
}
Write-Host "Object creation phase complete!"
Write-Host "Successfully created: $($CreatedDNs.Count) object(s)."

###############################################################################
# Prompt user about attribute assignment
###############################################################################
$validChoices = 'N','A','X'
$choice = $null
while($validChoices -notcontains $choice) {
    $choice = Read-Host "Would you like to set attributes for (N)ewly-created objects only, (A)ll objects from import, or e(X)it/skip? (N/A/X)"
}

if($choice -eq 'X') {
    Write-Host "Skipping attribute assignment as requested."
    return
}

# Determine which objects to assign attributes to
if($choice -eq 'N') {
    $ObjectsToAssign = $ImportedObjects | Where-Object {
        $CreatedDNs.Contains($_.DistinguishedName)
    }
}
else {
    # 'A' => All objects in the import (except the base container itself)
    $ObjectsToAssign = $ImportedObjects | Where-Object {
        $_.DistinguishedName -ne $Path
    }
}

###############################################################################
# Assign attributes in a separate loop
###############################################################################


# Exclude read-only attributes
$ExcludedAttributes = @(
    "objectGUID",
    "objectSid",
    "whenCreated",
    "whenChanged",
    "uSNChanged",
    "pwdLastSet",
    "lastLogonTimestamp",
    "badPasswordTime",
    "uSNCreated",
    "CanonicalName",
    "ProtectedFromAccidentalDeletion",
    "instanceType",
    "ObjectClass",
    "ObjectCategory",
    "msDS-UserAccountDisabled",
    "userPrincipalName",
    "badPwdCount",
    "DisplayName",
    "givenName",
    "sDRightsEffective",
    "Name",
    "modifyTimeStamp",
    "Modified",
    "lockoutTime",
    "Created",
    "dSCorePropagationData",
    "CN",
    "createTimeStamp",
    "DistinguishedName",
    "PropertyCount",
    "PropertyNames"
)

foreach($Object in $ObjectsToAssign) {
    # Attempt to bind
    $fullPath = "$Target$($Object.DistinguishedName)"
    try {
        $entry = New-Object System.DirectoryServices.DirectoryEntry(
            $fullPath, $Username, $Password, $authType
        )
        if($entry -eq $null -or $entry.Properties -eq $null) {
            throw "Binding returned null for: $($Object.DistinguishedName)"
        }
    }
    catch {
        Write-Host "Cannot bind to $($Object.DistinguishedName). Skipping attribute assignment. Error: $($_.Exception.Message)"
        continue
    }
	
	Write-Host "Bound to:" $fullPath

    # Build a hashtable of non-excluded, non-empty attributes
    $NewAttributes = @{}
    foreach ($attr in $Object.PSObject.Properties) {
        if ($ExcludedAttributes -notcontains $attr.Name) {
            $val = $attr.Value
            if (
                $val -ne $null -and
                -not (
                    ($val -is [array]  -and $val.Count -eq 0) -or
                    ($val -is [string] -and $val.Trim() -eq "")
                )
            ) {
                $NewAttributes[$attr.Name] = $val
            }
        }
    }

    Write-Host "Assigning attributes for: $($Object.DistinguishedName)"
    foreach($AttrKey in $NewAttributes.Keys) {
		Write-Host "Trying Attribute: $AttrKey - $NewAttributes[$AttrKey].Value"
        try {
            $entry.Properties[$AttrKey].Value = $NewAttributes[$AttrKey]
        }
        catch {
            Write-Host "   [WARN] Could not set '$AttrKey' on $($Object.DistinguishedName): $($_.Exception.Message)"
        }
		
		try {
			$entry.CommitChanges()
			Write-Host "   -> Attributes updated for $($Object.DistinguishedName)"
		}
		catch {
			Write-Host "   [ERROR] Committing changes for $($Object.DistinguishedName): $($_.Exception.Message)"
		}
    }
}

Write-Host "Import and attribute assignment completed successfully!"
