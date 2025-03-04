#Requires -Version 5.1

# Script Variables
$script:logPath = "$env:USERPROFILE\NetworkDriveMapping"
$script:driveMapping = @(
    @{
        DriveLetter = "I"
        Path = "\\fs01.server.pri\fs\Public"
    },
    @{
        DriveLetter = "J"
        Path = "\\fs01.server.pri\fs\Documents"
    }
)

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter()]
        [ValidateSet('Information', 'Warning', 'Error')]
        [string]$Level = 'Information'
    )
    
    $logMessage = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    [void](Add-Content -Path $script:logPath -Value $logMessage)
    Write-Output $logMessage
}

function Test-NetworkConnectivity {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Server
    )
    
    try {
        $serverName = ($Server -split '\\')[2]
        $result = Test-NetConnection -ComputerName $serverName -Port 445 -WarningAction SilentlyContinue
        return $result.TcpTestSucceeded
    }
    catch {
        Write-Log -Message "Failed to test network connectivity to $Server. Error: $($_.Exception.Message)" -Level Error
        return $false
    }
}

function Connect-RequiredVpn {
    try {
        Write-Log -Message "Checking VPN connection status" -Level Information
        
        # Try to find the VPN connection in current user context
        $vpnConnection = Get-VpnConnection -Name "VPN Connection" -ErrorAction SilentlyContinue
        
        # If not found, check all user connections
        if (-not $vpnConnection) {
            Write-Log -Message "VPN not found in current user context, checking all user connections" -Level Information
            $vpnConnection = Get-VpnConnection -AllUserConnection -Name "VPN Connection" -ErrorAction SilentlyContinue
        }
        
        # If VPN connection found
        if ($vpnConnection) {
            Write-Log -Message "Found VPN connection 'VPN Connection' with status: $($vpnConnection.ConnectionStatus)" -Level Information
            
            # If disconnected, attempt to connect
            if ($vpnConnection.ConnectionStatus -eq "Disconnected") {
                Write-Log -Message "Attempting to connect to VPN" -Level Information
                
                # Use rasdial command to connect to the VPN
                $result = rasdial "VPN Connection" 2>&1
                
                # Check if connection successful
                if ($LASTEXITCODE -eq 0) {
                    Write-Log -Message "Successfully connected to VPN" -Level Information
                    Start-Sleep -Seconds 5  # Give the connection time to establish fully
                    return $true
                }
                else {
                    Write-Log -Message "Failed to connect to VPN: $result" -Level Warning
                    return $false
                }
            }
            elseif ($vpnConnection.ConnectionStatus -eq "Connected") {
                Write-Log -Message "VPN is already connected" -Level Information
                return $true
            }
            else {
                Write-Log -Message "VPN is in state: $($vpnConnection.ConnectionStatus), not attempting connection" -Level Warning
                return $false
            }
        }
        else {
            Write-Log -Message "VPN connection 'VPN Connection' not found" -Level Warning
            return $false
        }
    }
    catch {
        Write-Log -Message "Error checking/connecting to VPN: $($_.Exception.Message)" -Level Error
        return $false
    }
}

function Test-EntraGroupMembership {
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupName
    )
    
    try {
        Write-Log -Message "Starting Entra ID group membership check for group: $GroupName" -Level Information
        
        # Try to get Microsoft Graph PowerShell module
        $graphModule = Get-Module -Name Microsoft.Graph.Groups -ListAvailable
        if (-not $graphModule) {
            Write-Log -Message "Microsoft.Graph.Groups module not found. Skipping Entra ID check." -Level Warning
            return $false
        }
        Write-Log -Message "Microsoft Graph module is available" -Level Information

        # Connect to Microsoft Graph if not already connected
        try {
            $context = Get-MgContext -ErrorAction Stop
            if (-not $context) {
                throw "No existing connection"
            }
            Write-Log -Message "Using existing Microsoft Graph connection" -Level Information
        }
        catch {
            Write-Log -Message "No existing Microsoft Graph connection, attempting to connect..." -Level Information
            try {
                Connect-MgGraph -Scopes "User.Read.All", "Group.Read.All" -ErrorAction Stop
                Write-Log -Message "Successfully connected to Microsoft Graph" -Level Information
            }
            catch {
                Write-Log -Message "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -Level Error
                return $false
            }
        }

        # Get current user's UPN
        $usernamePrefix = $env:USERNAME
        $userUPN = "$usernamePrefix@contoso.com"
        Write-Log -Message "Looking up user in Entra ID with UPN: $userUPN" -Level Information

        # Get current user
        try {
            $userFilter = "userPrincipalName eq '$userUPN'"
            Write-Log -Message "Executing Entra ID user query with filter: $userFilter" -Level Information
            $user = Get-MgUser -Filter $userFilter -ErrorAction Stop
            
            if (-not $user) {
                Write-Log -Message "User not found in Entra ID with UPN: $userUPN" -Level Warning
                
                # Try alternative lookup by display name
                Write-Log -Message "Trying alternative user lookup by display name" -Level Information
                $user = Get-MgUser -Filter "displayName eq '$usernamePrefix'" -ErrorAction Stop
                
                if (-not $user) {
                    Write-Log -Message "User not found in Entra ID by display name either" -Level Warning
                    return $false
                }
                Write-Log -Message "Found user by display name: $($user.DisplayName) (ID: $($user.Id))" -Level Information
            }
            else {
                Write-Log -Message "Found user in Entra ID: $($user.DisplayName) (ID: $($user.Id))" -Level Information
            }
        }
        catch {
            Write-Log -Message "Error searching for user in Entra ID: $($_.Exception.Message)" -Level Error
            return $false
        }

        # Get group and check membership
        try {
            $groupFilter = "displayName eq '$GroupName'"
            Write-Log -Message "Looking up group in Entra ID with filter: $groupFilter" -Level Information
            $group = Get-MgGroup -Filter $groupFilter -ErrorAction Stop
            
            if (-not $group) {
                Write-Log -Message "Group '$GroupName' not found in Entra ID" -Level Warning
                return $false
            }
            
            Write-Log -Message "Found group: $($group.DisplayName) (ID: $($group.Id))" -Level Information
            
            # Check if user is a member of the group
            try {
                Write-Log -Message "Checking if user is member of the group" -Level Information
                $members = Get-MgGroupMember -GroupId $group.Id -ErrorAction Stop
                $isMember = $members | Where-Object { $_.Id -eq $user.Id }
                
                if ($null -ne $isMember) {
                    Write-Log -Message "User is a member of Entra ID group $GroupName" -Level Information
                    return $true
                }
                
                Write-Log -Message "User is NOT a member of Entra ID group $GroupName" -Level Information
                return $false
            }
            catch {
                Write-Log -Message "Error checking group membership: $($_.Exception.Message)" -Level Error
                return $false
            }
        }
        catch {
            Write-Log -Message "Error looking up group in Entra ID: $($_.Exception.Message)" -Level Error
            return $false
        }
    }
    catch {
        Write-Log -Message "Failed to check Entra ID group membership for $GroupName. Error: $($_.Exception.Message)" -Level Error
        return $false
    }
}

function Test-SecurityGroupMembership {
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupName,
        
        [Parameter()]
        [string]$EntraGroupName
    )
    
    # Try AD group check first, in its own try-catch
    try {
        Write-Log -Message "Attempting AD group membership check for $GroupName" -Level Information
        $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        $groups = $identity.Groups | ForEach-Object {
            try { $_.Translate([System.Security.Principal.NTAccount]).Value } catch { }
        }
        
        if ($groups -contains $GroupName) {
            Write-Log -Message "User is a member of AD group $GroupName" -Level Information
            return @{ Success = $true; Source = 'AD' }
        }
        
        # If we get here, AD check ran but user is not a member
        Write-Log -Message "AD check completed, user is not a member of $GroupName" -Level Information
    }
    catch {
        Write-Log -Message "AD group check failed with error: $($_.Exception.Message)" -Level Warning
        # Don't return, continue to Entra check
    }
    
    # If Entra group is specified, try that regardless of whether AD check succeeded
    if ($EntraGroupName) {
        Write-Log -Message "Checking Entra ID group membership for $EntraGroupName" -Level Information
        
        # Skip Entra check if Graph modules aren't available
        $graphModule = Get-Module -Name Microsoft.Graph.Groups -ListAvailable
        if (-not $graphModule) {
            Write-Log -Message "Microsoft.Graph.Groups module not found. Skipping Entra ID check." -Level Warning
            Write-Log -Message "Assuming permission granted based on AD check (Entra check skipped)" -Level Information
            return @{ Success = $true; Source = 'Assumed' }
        }
        
        # Try Entra check with timeout
        try {
            # Create a script block for the Entra check
            $entraCheckScript = {
                param($EntraGroupName)
                Test-EntraGroupMembership -GroupName $EntraGroupName
            }
            
            # Start job with timeout
            $job = Start-Job -ScriptBlock $entraCheckScript -ArgumentList $EntraGroupName
            
            # Wait for job with timeout (10 seconds)
            $completed = Wait-Job -Job $job -Timeout 10
            
            if ($completed -and $completed.State -eq 'Completed') {
                $entraResult = Receive-Job -Job $job
                if ($entraResult) {
                    Write-Log -Message "Entra check succeeded, user is a member of $EntraGroupName" -Level Information
                    return @{ Success = $true; Source = 'Entra' }
                }
                Write-Log -Message "Entra check completed, user is not a member of $EntraGroupName" -Level Information
            } 
            else {
                Stop-Job -Job $job -ErrorAction SilentlyContinue
                Write-Log -Message "Entra check timed out after 10 seconds" -Level Warning
                Write-Log -Message "Assuming permission granted (Entra check timeout)" -Level Information
                return @{ Success = $true; Source = 'Assumed' }
            }
        }
        catch {
            Write-Log -Message "Entra group check failed with error: $($_.Exception.Message)" -Level Error
            Write-Log -Message "Assuming permission granted (Entra check error)" -Level Information
            return @{ Success = $true; Source = 'Assumed' }
        }
        finally {
            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
        }
    }
    else {
        Write-Log -Message "No Entra group specified, skipping Entra check" -Level Information
    }
    
    # If we get here, both checks failed or weren't applicable
    return @{ Success = $false; Source = $null }
}

function Remove-StaleDriveMapping {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DriveLetter
    )
    
    try {
        [void](Write-Log -Message "Checking for existing mappings for drive $($DriveLetter):" -Level Information)
        
        # First try to force delete the mapping regardless of what net use shows
        # This helps with "remembered connections" that might not show in net use
        $deleteResult = net use "$($DriveLetter):" /delete /y 2>&1
        
        # If the delete worked or if there was no mapping, we're good
        if ($LASTEXITCODE -eq 0) {
            [void](Write-Log -Message "Successfully removed any existing mapping for $($DriveLetter):" -Level Information)
        }
        else {
            # If we get specific errors like "not mapped" that's fine
            if ($deleteResult -match "not mapped" -or $deleteResult -match "not found") {
                [void](Write-Log -Message "No existing mapping found for $($DriveLetter):" -Level Information)
            }
            else {
                # For other errors, log but continue
                [void](Write-Log -Message "Warning while removing $($DriveLetter): - $deleteResult" -Level Warning)
                
                # Try alternative method to clear remembered connections
                try {
                    # Try clearing the remembered connection more aggressively
                    $clearResult = cmd /c "net use $($DriveLetter): /delete" 2>&1
                    
                    if ($clearResult -match "was deleted successfully") {
                        [void](Write-Log -Message "Successfully cleared remembered connection for $($DriveLetter): via CMD" -Level Information)
                    }
                    elseif ($clearResult -match "not found" -or $clearResult -match "does not exist") {
                        [void](Write-Log -Message "No remembered connection found for $($DriveLetter): via CMD" -Level Information)
                    }
                    else {
                        [void](Write-Log -Message "Result of alternative method for $($DriveLetter): - $clearResult" -Level Warning)
                    }
                }
                catch {
                    [void](Write-Log -Message "Failed alternative method to clear $($DriveLetter): - $($_.Exception.Message)" -Level Warning)
                }
            }
        }
        
        # Always return false to force recreation of the mapping
        return $false
    }
    catch {
        [void](Write-Log -Message "Failed to check/remove drive mapping $($DriveLetter):. Error: $($_.Exception.Message)" -Level Error)
        return $false  # Error occurred, proceed with creating a new mapping
    }
}

# Main execution block

# Clear existing log files
if (-not (Test-Path -Path (Split-Path -Path $script:logPath -Parent))) {
    New-Item -ItemType Directory -Path (Split-Path -Path $script:logPath -Parent) -Force | Out-Null
}

if (Test-Path -Path $script:logPath) {
    Remove-Item -Path $script:logPath -Force
}
if (Test-Path -Path "$($script:logPath).log") {
    Remove-Item -Path "$($script:logPath).log" -Force
}

# Start logging
Start-Transcript -Path "$($script:logPath).log" -ErrorAction SilentlyContinue
Write-Log -Message "Starting network drive mapping script in $($ExecutionContext.SessionState.LanguageMode) mode"

# Try to clear all network drive mappings first
try {
    Write-Log -Message "Clearing all existing network drive mappings to prevent remembered connection issues" -Level Information
    $clearResult = net use * /delete /y 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Log -Message "Successfully cleared all network drive mappings" -Level Information
    }
    else {
        # If there are no drives to delete, that's fine
        if ($clearResult -match "no entries") {
            Write-Log -Message "No existing mappings to clear" -Level Information
        }
        else {
            Write-Log -Message "Warning while clearing mappings: $clearResult" -Level Warning
        }
    }
}
catch {
    Write-Log -Message "Error while clearing existing mappings: $($_.Exception.Message)" -Level Warning
}

# Main execution block
try {
    # Check server connectivity but don't exit if it fails
    $serverConnectivity = Test-NetworkConnectivity -Server $driveMapping[0].Path
    
    # Add direct ping test
    $serverName = ($driveMapping[0].Path -split '\\')[2]
    Write-Log -Message "Performing direct ping test to server $serverName" -Level Information
    $pingResult = Test-Connection -ComputerName $serverName -Count 1 -ErrorAction SilentlyContinue
    if ($pingResult) {
        Write-Log -Message "Ping test to $serverName successful: $($pingResult.IPv4Address)" -Level Information
    } else {
        Write-Log -Message "Ping test to $serverName failed!" -Level Warning
    }
    
    # Try connecting to port 445 (SMB) directly using Test-NetConnection
    Write-Log -Message "Testing SMB connectivity to $serverName on port 445" -Level Information
    $tcpTest = Test-NetConnection -ComputerName $serverName -Port 445 -InformationLevel Detailed -WarningAction SilentlyContinue
    Write-Log -Message "TCP test to $serverName:445 result - TcpTestSucceeded: $($tcpTest.TcpTestSucceeded)" -Level Information
    
    # If connectivity fails, try connecting to VPN and test again
    if (-not $serverConnectivity) {
        Write-Log -Message "Cannot connect to file server. Attempting to establish VPN connection." -Level Warning
        
        $vpnConnected = Connect-RequiredVpn
        if ($vpnConnected) {
            # Retest connectivity after VPN connection
            Start-Sleep -Seconds 5  # Allow network to stabilize
            $serverConnectivity = Test-NetworkConnectivity -Server $driveMapping[0].Path
            
            if ($serverConnectivity) {
                Write-Log -Message "Successfully connected to file server after VPN connection" -Level Information
            }
            else {
                Write-Log -Message "Still cannot connect to file server after VPN connection. Will attempt to map drives anyway." -Level Warning
            }
        }
        else {
            Write-Log -Message "Failed to establish VPN connection. Will attempt to map drives anyway." -Level Warning
        }
    }

    foreach ($drive in $driveMapping) {
        try {
            Write-Log -Message "Processing drive $($drive.DriveLetter):"
            
            # Check security group requirements
            if ($drive.SecurityGroup) {
                Write-Log -Message "Drive $($drive.DriveLetter): requires security group membership check"
                $accessCheck = Test-SecurityGroupMembership -GroupName $drive.SecurityGroup -EntraGroupName $drive.EntraGroup
                if (-not $accessCheck.Success) {
                    Write-Log -Message "Access denied for drive $($drive.DriveLetter): - User is not a member of required groups" -Level Warning
                    continue
                }
                Write-Log -Message "Access granted for drive $($drive.DriveLetter): - Security check passed ($($accessCheck.Source) group membership)" -Level Information
            }

            # Check if drive already exists with correct path
            [bool]$keepExisting = Remove-StaleDriveMapping -DriveLetter $drive.DriveLetter

            # Create new mapping only if needed
            if (-not $keepExisting) {
                Write-Log -Message "Attempting to map drive $($drive.DriveLetter): to $($drive.Path)" -Level Information
                
                # Try direct net use approach instead of Invoke-Expression
                try {
                    # Execute the command directly
                    $result = net use "$($drive.DriveLetter):" "$($drive.Path)" /PERSISTENT:YES 2>&1
                    $exitCode = $LASTEXITCODE
                    
                    Write-Log -Message "Command result: $($result -join ' ')" -Level Information
                    
                    # Check if command succeeded
                    if ($exitCode -eq 0) {
                        Write-Log -Message "Successfully mapped drive $($drive.DriveLetter): to $($drive.Path)" -Level Information
                        
                        # Verify the drive is actually accessible
                        if (Test-Path -Path "$($drive.DriveLetter):\" -ErrorAction SilentlyContinue) {
                            Write-Log -Message "Confirmed drive $($drive.DriveLetter): is accessible" -Level Information
                        }
                        else {
                            Write-Log -Message "WARNING: Drive $($drive.DriveLetter): was mapped but is not accessible" -Level Warning
                        }
                    }
                    else {
                        # Extract error code if present in the result
                        $errorCode = if ($result -match "System error (\d+)") { $matches[1] } else { "unknown" }
                        
                        # Format the result for better readability
                        $formattedResult = ($result -join " ").Trim()
                        
                        # Log with error code details
                        Write-Log -Message "ERROR mapping drive $($drive.DriveLetter): - System error $errorCode" -Level Error
                        Write-Log -Message "ERROR details: $formattedResult" -Level Error
                        
                        # Provide additional context for common error codes
                        switch ($errorCode) {
                            "53"  { Write-Log -Message "This error indicates the network path was not found." -Level Information }
                            "67"  { Write-Log -Message "This error indicates the network name cannot be found." -Level Information }
                            "86"  { Write-Log -Message "This error indicates the specified network password is not correct." -Level Information }
                            "1202" { Write-Log -Message "This error indicates the device is already assigned to a different network resource." -Level Information }
                            "1208" { Write-Log -Message "This error indicates an extended error has occurred." -Level Information }
                            "1219" { Write-Log -Message "This error indicates multiple connections to a server or shared resource by the same user, using more than one user name, are not allowed." -Level Information }
                            "1326" { Write-Log -Message "This error indicates logon failure: unknown user name or bad password." -Level Information }
                            "1909" { Write-Log -Message "This error indicates the referenced account is currently locked out and may not be logged on to." -Level Information }
                        }
                    }
                }
                catch {
                    Write-Log -Message "EXCEPTION mapping drive $($drive.DriveLetter): - $($_.Exception.Message)" -Level Error
                    Write-Log -Message "Exception details: $($_ | Out-String)" -Level Error
                }
                
                # Add visual separator for readability
                Write-Log -Message "------------------------------------------------------------" -Level Information
            }
        }
        catch {
            Write-Log -Message "Failed to map drive $($drive.DriveLetter):. Error: $($_.Exception.Message)" -Level Warning
            # Continue with next drive
            continue
        }
    }

    # Verify drive mappings exist after completion
    Write-Log -Message "Verifying drive mappings after completion:" -Level Information
    try {
        # Store the raw output for debugging
        $netUseRaw = net use 
        Write-Log -Message "Raw net use output available in log file" -Level Information
        
        # Dump the raw output as a string for debug purposes
        $netUseRawStr = $netUseRaw -join " "
        Write-Log -Message "Net use output summary: $netUseRawStr" -Level Information
        
        # Check for each expected drive mapping using a simpler approach
        $drivesFound = 0
        $drivesNotFound = 0
        
        foreach ($drive in $driveMapping) {
            $driveLetter = "$($drive.DriveLetter):"
            
            # Use a simpler method - just check if the drive letter exists in the output
            if ($netUseRaw -match $driveLetter) {
                Write-Log -Message "Verified mapping exists: $driveLetter" -Level Information
                $drivesFound++
            } else {
                Write-Log -Message "WARNING: Mapping not found for $driveLetter" -Level Warning
                $drivesNotFound++
            }
        }
        
        Write-Log -Message "Drive mapping verification complete. Found: $drivesFound, Not found: $drivesNotFound" -Level Information
        
        # Evaluate the success rate of drive mappings
        $totalDrives = $driveMapping.Count
        
        if ($drivesFound -eq 0) {
            # No drives were mapped - this is a severe error
            Write-Log -Message "No drives were successfully mapped ($drivesFound/$totalDrives)" -Level Error
            exit 3
        }
        elseif ($drivesFound -lt $totalDrives) {
            # Some drives were mapped, but not all - this is a warning
            Write-Log -Message "Only some drives were successfully mapped ($drivesFound/$totalDrives)" -Level Warning
            exit 3
        }
        else {
            # All drives were mapped successfully
            Write-Log -Message "SUCCESS: All drives were successfully mapped ($drivesFound/$totalDrives)" -Level Information
            # Continue to successful completion
        }
    }
    catch {
        Write-Log -Message "Error verifying drive mappings: $($_.Exception.Message)" -Level Error
    }

    # If initial server connectivity failed, exit with retry code
    if (-not $serverConnectivity) {
        Write-Log -Message "Some drives may not have mapped successfully. Exiting with retry code to attempt again later." -Level Warning
        exit 3
    }

    Write-Log -Message "Network drive mapping completed successfully"
    exit 0
}
catch {
    Write-Log -Message "Script failed with error: $($_.Exception.Message)" -Level Error
    exit 4
}
finally {
    Stop-Transcript -ErrorAction SilentlyContinue
} 