# Script for automatic build of 1C extensions and processings
# Encoding: UTF-8 with BOM
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Function to read and parse ibases.v8i file
function Get-1CBases {
    param(
        [string]$V8iPath = "$env:APPDATA\1C\1CEStart\ibases.v8i"
    )
    
    if (-not (Test-Path $V8iPath)) {
        Write-Host "Error: file ibases.v8i not found at path: $V8iPath" -ForegroundColor Red
        return $null
    }
    
    $bases = @()
    
    # Parse file line by line, avoiding regex issues
    $lines = Get-Content -Path $V8iPath -Encoding UTF8
    $currentBase = $null
    $currentConnect = ""
    
    foreach ($line in $lines) {
        $line = $line.Trim()
        
        # Check for new section start [BaseName]
        $openBracketIndex = $line.IndexOf([char]91)
        $closeBracketIndex = $line.IndexOf([char]93)
        if ($openBracketIndex -eq 0 -and $closeBracketIndex -gt 0) {
            # Save previous base if exists
            if ($currentBase -and $currentConnect) {
                $bases += [PSCustomObject]@{
                    Name = $currentBase
                    Connect = $currentConnect
                }
            }
            
            # Extract base name from line like [BaseName]
            $startIndex = $openBracketIndex
            $endIndex = $closeBracketIndex
            if ($startIndex -ge 0 -and $endIndex -gt $startIndex) {
                $currentBase = $line.Substring($startIndex + 1, $endIndex - $startIndex - 1)
                $currentConnect = ""
            }
        }
        # Check connection string
        elseif ($line -like 'Connect=*') {
            $currentConnect = $line.Substring(7).Trim()
        }
    }
    
    # Save last base
    if ($currentBase -and $currentConnect) {
        $bases += [PSCustomObject]@{
            Name = $currentBase
            Connect = $currentConnect
        }
    }
    
    return $bases
}

# Function to select base from list
function Select-1CBase {
    param(
        [array]$Bases
    )
    
    if ($Bases.Count -eq 0) {
        Write-Host "Error: base list is empty" -ForegroundColor Red
        return $null
    }
    
    Write-Host "`nSelect 1C database:" -ForegroundColor Cyan
    Write-Host "=" * 50 -ForegroundColor Cyan
    
    for ($i = 0; $i -lt $Bases.Count; $i++) {
        Write-Host "$($i + 1). $($Bases[$i].Name)" -ForegroundColor Yellow
    }
    
    Write-Host "=" * 50 -ForegroundColor Cyan
    
    do {
        $selection = Read-Host "Enter base number (1-$($Bases.Count))"
        $selectedIndex = [int]$selection - 1
    } while ($selectedIndex -lt 0 -or $selectedIndex -ge $Bases.Count)
    
    return $Bases[$selectedIndex]
}

# Function to determine project type (extension or processing)
function Get-ProjectType {
    param(
        [string]$ProjectPath = "."
    )
    
    $configXml = Join-Path $ProjectPath "Configuration.xml"
    
    # Check for Configuration.xml (Extension)
    if (Test-Path $configXml) {
        return [PSCustomObject]@{
            Type = "Extension"
            XmlFile = $configXml
        }
    }
    
    # Search for processing
    $xmlFiles = Get-ChildItem -Path $ProjectPath -Filter "*.xml" -File
    
    foreach ($xmlFile in $xmlFiles) {
        try {
            [xml]$xmlContent = Get-Content -Path $xmlFile.FullName -Encoding UTF8
            
            # Check for ExternalDataProcessor tag
            if ($xmlContent.MetaDataObject.ExternalDataProcessor) {
                return [PSCustomObject]@{
                    Type = "Processing"
                    XmlFile = $xmlFile.FullName
                }
            }
        }
        catch {
            # Skip files with parsing errors
            continue
        }
    }
    
    return $null
}

# Function to extract project name from XML
function Get-ProjectName {
    param(
        [string]$XmlFile,
        [string]$ProjectType
    )
    
    $result = $null
    try
    {
        [xml]$xmlContent = Get-Content -Path $XmlFile -Encoding UTF8
        
        if ($ProjectType -eq "Extension")
        {
            $nameNode = $xmlContent.MetaDataObject.Configuration.Properties.Name
            if ($nameNode)
            {
                $result = $nameNode
            }
        }
        elseif ($ProjectType -eq "Processing")
        {
            $nameNode = $xmlContent.MetaDataObject.ExternalDataProcessor.Properties.Name
            if ($nameNode)
            {
                $result = $nameNode
            }
        }
        
        if (-not $result)
        {
            Write-Host "Error: Name tag not found in file $XmlFile" -ForegroundColor Red
        }
    }
    catch
    {
        Write-Host "Error parsing XML file: $_" -ForegroundColor Red
    }
    
    return $result
}

# Function to convert connection string from ibases.v8i format to command line format
function Convert-1CConnectionString {
    param(
        [string]$ConnectString
    )
    
    # Parse connection string like: Srvr="DNA-DEVAPPS-1S0.dna-tankers.com";Ref="Storage1";
    # Convert to: /S "DNA-DEVAPPS-1S0.dna-tankers.com\Storage1"
    
    $server = ""
    $database = ""
    
    # Extract server name
    if ($ConnectString -match 'Srvr="([^"]+)"') {
        $server = $matches[1]
    }
    
    # Extract database name
    if ($ConnectString -match 'Ref="([^"]+)"') {
        $database = $matches[1]
    }
    
    if ($server -and $database) {
        return "/S `"$server\$database`""
    }
    elseif ($server) {
        # If only server is specified, return just server
        return "/S `"$server`""
    }
    else {
        # Fallback to original format if parsing fails
        return "/F `"$ConnectString`""
    }
}

# Function to check if built-in OpenSSH is available
function Ensure-OpenSSHClient {
    $sshPath = $null
    $possiblePaths = @(
        "$env:ProgramFiles\OpenSSH\ssh.exe",
        "$env:ProgramFiles(x86)\OpenSSH\ssh.exe",
        "C:\Windows\System32\OpenSSH\ssh.exe"
    )
    
    foreach ($path in $possiblePaths) {
        if (Test-Path $path) {
            $sshPath = $path
            break
        }
    }
    
    # Also check if ssh is in PATH
    if (-not $sshPath) {
        try {
            $sshCheck = Get-Command ssh -ErrorAction Stop
            $sshPath = $sshCheck.Source
        }
        catch {
            # SSH not found in PATH
        }
    }
    
    if ($sshPath) {
        Write-Host "Using built-in OpenSSH: $sshPath" -ForegroundColor Gray
        return $sshPath
    }
    else {
        Write-Host "Error: OpenSSH client not found. Please install OpenSSH client feature in Windows." -ForegroundColor Red
        Write-Host "You can install it with: Add-WindowsCapability -Online -Name OpenSSH.Client~~~~0.0.1.0" -ForegroundColor Yellow
        return $null
    }
}

# Function to parse JSON response from agent
function Parse-AgentResponse {
    param(
        [string]$Response
    )
    
    if ([string]::IsNullOrWhiteSpace($Response)) {
        return $null
    }
    
    try {
        # Remove prompt and clean response
        $cleanResponse = $Response -replace 'designer>', '' -replace '^\s+', '' -replace '\s+$', ''
        
        # Try to find JSON array in response (starts with [)
        if ($cleanResponse -match '\[.*\]') {
            $jsonMatch = $matches[0]
            $jsonResponse = $jsonMatch | ConvertFrom-Json
            
            if ($jsonResponse -is [System.Array]) {
                return $jsonResponse
            }
            elseif ($jsonResponse) {
                return @($jsonResponse)
            }
        }
        else {
            # Try to parse entire response as JSON
            $jsonResponse = $cleanResponse | ConvertFrom-Json
            
            if ($jsonResponse -is [System.Array]) {
                return $jsonResponse
            }
            elseif ($jsonResponse) {
                return @($jsonResponse)
            }
        }
        
        return $null
    }
    catch {
        Write-Host "Warning: Failed to parse JSON response: $_" -ForegroundColor Yellow
        Write-Host "Response preview: $($Response.Substring(0, [Math]::Min(200, $Response.Length)))" -ForegroundColor Gray
        return $null
    }
}

# Function to check if agent command was successful
function Test-AgentCommandSuccess {
    param(
        [array]$JsonResponse
    )
    
    if (-not $JsonResponse) {
        return $false
    }
    
    foreach ($item in $JsonResponse) {
        if ($item.type -eq "error") {
            Write-Host "Agent error: $($item.message)" -ForegroundColor Red
            if ($item.'error-type') {
                Write-Host "Error type: $($item.'error-type')" -ForegroundColor Red
            }
            return $false
        }
        if ($item.type -eq "canceled") {
            Write-Host "Operation was canceled" -ForegroundColor Yellow
            return $false
        }
    }
    
    # Check for success in the last item
    $lastItem = $JsonResponse[-1]
    if ($lastItem.type -eq "success") {
        return $true
    }
    
    return $false
}

# Function to generate SSH host key if needed
function Generate-SSHHostKey {
    param(
        [string]$KeysDirectory = "$env:APPDATA\1C\SSHKeys"
    )
    
    if (-not (Test-Path $KeysDirectory)) {
        New-Item -ItemType Directory -Path $KeysDirectory -Force | Out-Null
    }
    
    $rsaKeyFile = Join-Path $KeysDirectory "ssh_host_rsa_key"
    $ed25519KeyFile = Join-Path $KeysDirectory "ssh_host_ed25519_key"
    
    # Check if OpenSSH is available (Windows 10+)
    $sshKeygenPath = $null
    $possiblePaths = @(
        "$env:ProgramFiles\OpenSSH\ssh-keygen.exe",
        "$env:ProgramFiles(x86)\OpenSSH\ssh-keygen.exe",
        "C:\Windows\System32\OpenSSH\ssh-keygen.exe"
    )
    
    foreach ($path in $possiblePaths) {
        if (Test-Path $path) {
            $sshKeygenPath = $path
            break
        }
    }
    
    # Generate ED25519 key if it doesn't exist (preferred)
    if (-not (Test-Path $ed25519KeyFile)) {
        if ($sshKeygenPath) {
            Write-Host "Generating ED25519 SSH host key..." -ForegroundColor Cyan
            try {
                $process = Start-Process -FilePath $sshKeygenPath -ArgumentList @(
                    "-t", "ed25519",
                    "-f", "`"$ed25519KeyFile`"",
                    "-N", "`"`"",
                    "-C", "`"1C Agent Host Key`""
                ) -Wait -PassThru -NoNewWindow
                
                if ($process.ExitCode -eq 0 -and (Test-Path $ed25519KeyFile)) {
                    Write-Host "ED25519 key generated successfully" -ForegroundColor Green
                    # Ensure public key file exists
                    $publicKeyFile = "$ed25519KeyFile.pub"
                    if (-not (Test-Path $publicKeyFile)) {
                        Write-Host "Note: Public key file not found, but private key exists" -ForegroundColor Yellow
                    }
                    return $ed25519KeyFile
                }
            }
            catch {
                Write-Host "Warning: Failed to generate ED25519 key: $_" -ForegroundColor Yellow
            }
        }
    }
    elseif (Test-Path $ed25519KeyFile) {
        return $ed25519KeyFile
    }
    
    # Generate RSA key if ED25519 failed or doesn't exist
    if (-not (Test-Path $rsaKeyFile)) {
        if ($sshKeygenPath) {
            Write-Host "Generating RSA SSH host key..." -ForegroundColor Cyan
            try {
                $process = Start-Process -FilePath $sshKeygenPath -ArgumentList @(
                    "-t", "rsa",
                    "-b", "2048",
                    "-f", "`"$rsaKeyFile`"",
                    "-N", "`"`"",
                    "-C", "`"1C Agent Host Key`""
                ) -Wait -PassThru -NoNewWindow
                
                if ($process.ExitCode -eq 0 -and (Test-Path $rsaKeyFile)) {
                    Write-Host "RSA key generated successfully" -ForegroundColor Green
                    # Ensure public key file exists
                    $publicKeyFile = "$rsaKeyFile.pub"
                    if (-not (Test-Path $publicKeyFile)) {
                        Write-Host "Note: Public key file not found, but private key exists" -ForegroundColor Yellow
                    }
                    return $rsaKeyFile
                }
            }
            catch {
                Write-Host "Warning: Failed to generate RSA key: $_" -ForegroundColor Yellow
            }
        }
    }
    elseif (Test-Path $rsaKeyFile) {
        return $rsaKeyFile
    }
    
    Write-Host "Warning: Could not generate SSH host keys. Agent may fail to start." -ForegroundColor Yellow
    Write-Host "Please generate SSH keys manually or install OpenSSH." -ForegroundColor Yellow
    return $null
}

# Function to start 1C configurator in agent mode
function Start-1CAgent {
    param(
        [string]$BaseConnect,
        [int]$Port = 1543
    )
    
    $configurator = "C:\Program Files\1cv8\common\1cestart.exe"
    
    if (-not (Test-Path $configurator)) {
        Write-Host "Error: 1C configurator not found at path: $configurator" -ForegroundColor Red
        return $null
    }
    
    # Convert connection string to command line format
    $connectionParam = Convert-1CConnectionString -ConnectString $BaseConnect
    
    # Build agent command - always use /AgentSSHHostKeyAuto for automatic key generation
    # According to 1C documentation:
    # /AgentSSHHostKeyAuto - автоматическая генерация ключа
    # Port will use default value if not specified
    
    $agentCommand = "DESIGNER $connectionParam /AgentMode /AgentSSHHostKeyAuto /Visible"
    
    Write-Host "Starting 1C agent (using default port)..." -ForegroundColor Cyan
    Write-Host "Using automatic SSH host key generation (/AgentSSHHostKeyAuto)" -ForegroundColor Gray
    Write-Host "Command: $configurator $agentCommand" -ForegroundColor Gray
    Write-Host ""
    
    try {
        $process = Start-Process -FilePath $configurator -ArgumentList $agentCommand -PassThru -WindowStyle Normal
        
        Write-Host "Waiting for agent to start..." -ForegroundColor Cyan
        
        # Wait longer for agent to start (agent may need time to initialize)
        Start-Sleep -Seconds 5
        
        # Check if process is still running or if agent is accessible
        # Note: Process may exit but agent continues running in another process
        $processRunning = -not $process.HasExited
        
        # Also check if there are any 1cv8 processes running (agent might be in separate process)
        $agentProcesses = Get-Process -Name "1cv8" -ErrorAction SilentlyContinue
        
        if ($processRunning -or $agentProcesses) {
            Write-Host "Agent process detected" -ForegroundColor Green
            if ($processRunning) {
                Write-Host "Main process PID: $($process.Id)" -ForegroundColor Gray
            }
            if ($agentProcesses) {
                Write-Host "Found $($agentProcesses.Count) 1cv8 process(es) running" -ForegroundColor Gray
            }
            
            # Try to verify agent is actually listening on the port
            Write-Host "Verifying agent is accessible on port $Port..." -ForegroundColor Cyan
            Start-Sleep -Seconds 2
            
            # Check if port is listening (simple check)
            try {
                $tcpClient = New-Object System.Net.Sockets.TcpClient
                $connect = $tcpClient.BeginConnect("localhost", $Port, $null, $null)
                $wait = $connect.AsyncWaitHandle.WaitOne(2000, $false)
                if ($wait) {
                    $tcpClient.EndConnect($connect)
                    $tcpClient.Close()
                    Write-Host "Agent is listening on port $Port" -ForegroundColor Green
                    return $process
                }
                else {
                    $tcpClient.Close()
                    Write-Host "Warning: Port $Port is not responding yet, but agent may still be starting..." -ForegroundColor Yellow
                    Write-Host "Continuing anyway - agent window should be visible if it started successfully." -ForegroundColor Yellow
                    return $process
                }
            }
            catch {
                Write-Host "Warning: Could not verify port connection: $_" -ForegroundColor Yellow
                Write-Host "Agent may still be starting. Check the agent window." -ForegroundColor Yellow
                # Return process anyway - user can see if agent window appeared
                return $process
            }
        }
        else {
            Write-Host "Warning: Initial process exited, but agent may have started in a separate process." -ForegroundColor Yellow
            Write-Host "Please check if the agent window appeared on screen." -ForegroundColor Yellow
            Write-Host ""
            Write-Host "If agent window is visible, you can continue. Otherwise, check for errors in the window." -ForegroundColor Cyan
            Write-Host ""
            Write-Host "Command used:" -ForegroundColor Gray
            Write-Host "  $configurator $agentCommand" -ForegroundColor Gray
            
            # Still return the process object (even if exited) so caller can proceed
            # The actual agent may be running in a different process
            return $process
        }
    }
    catch {
        Write-Host "Error starting agent: $_" -ForegroundColor Red
        return $null
    }
}

# Function to connect to 1C agent via SSH using built-in OpenSSH
# Configured according to 1C recommendations for SSH clients:
# - Local echo: Force on
# - Local line ending: Force on
# - Remote character set: UTF-8
# - Don't allocate a pseudo-terminal
function Connect-1CAgentSSH {
    param(
        [string]$AgentHost = "localhost",
        [int]$Port = 1543,
        [string]$Username,
        [SecureString]$Password
    )
    
    $sshPath = Ensure-OpenSSHClient
    if (-not $sshPath) {
        return $null
    }
    
    $credential = $null
    if ($Username -and $Password) {
        $credential = New-Object System.Management.Automation.PSCredential($Username, $Password)
    }
    else {
        # Prompt for credentials
        # Note: These should be 1C database credentials, not SSH credentials
        Write-Host "Enter credentials for 1C information base (not SSH credentials)" -ForegroundColor Cyan
        Write-Host "These are the same credentials you use to connect to the 1C database" -ForegroundColor Gray
        $credential = Get-Credential -Message "Enter 1C database credentials for agent connection" -UserName ""
    }
    
    Write-Host "Connecting to agent at ${AgentHost}:${Port}..." -ForegroundColor Cyan
    Write-Host "Using built-in OpenSSH client" -ForegroundColor Gray
    Write-Host "Username: $($credential.UserName)" -ForegroundColor Gray
    
    # Store credentials for use in command execution
    # Return a custom object with connection info instead of Posh-SSH session
    return [PSCustomObject]@{
        Host = $AgentHost
        Port = $Port
        Username = $credential.UserName
        Password = $credential.Password
        SshPath = $sshPath
        Connected = $true
    }
}

# Function to execute command on agent via SSH using built-in OpenSSH
# Configured according to 1C recommendations:
# - Local echo: Force on (handled by reading output)
# - Local line ending: Force on (CRLF for Windows)
# - Remote character set: UTF-8
# - Don't allocate a pseudo-terminal (using -T parameter)
function Invoke-1CAgentCommand {
    param(
        [object]$Session,
        [string]$Command,
        [int]$TimeoutSeconds = 300
    )
    
    if (-not $Session -or -not $Session.Connected) {
        Write-Host "Error: SSH session is not established" -ForegroundColor Red
        return $null
    }
    
    Write-Host "Executing command: $Command" -ForegroundColor Gray
    
    try {
        # Convert SecureString password to plain text
        $passwordPlain = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Session.Password)
        )
        
        # Build SSH command with 1C recommended settings
        # -T: Disable pseudo-terminal allocation (recommended for 1C agent)
        # -o StrictHostKeyChecking=no: Accept host key automatically
        # -o UserKnownHostsFile=NUL: Don't save host key
        # -o LogLevel=ERROR: Reduce verbosity
        # -o PreferredAuthentications=password: Use password authentication
        $sshArgs = @(
            "-T",
            "-o", "StrictHostKeyChecking=no",
            "-o", "UserKnownHostsFile=NUL",
            "-o", "LogLevel=ERROR",
            "-o", "PreferredAuthentications=password",
            "-p", $Session.Port.ToString(),
            "$($Session.Username)@$($Session.Host)",
            $Command
        )
        
        Write-Host "Running SSH command..." -ForegroundColor Gray
        
        # Create a PowerShell script that will handle password input interactively
        $tempScript = [System.IO.Path]::GetTempFileName() + ".ps1"
        $scriptContent = @"
`$password = '$passwordPlain'
`$sshPath = '$($Session.SshPath)'
`$sshArgs = @(
    '-T',
    '-o', 'StrictHostKeyChecking=no',
    '-o', 'UserKnownHostsFile=NUL',
    '-o', 'LogLevel=ERROR',
    '-o', 'PreferredAuthentications=password',
    '-p', '$($Session.Port)',
    '$($Session.Username)@$($Session.Host)',
    '$($Command -replace "'", "''")'
)

`$process = Start-Process -FilePath `$sshPath -ArgumentList `$sshArgs -NoNewWindow -Wait -PassThru -RedirectStandardOutput (New-TemporaryFile).FullName -RedirectStandardError (New-TemporaryFile).FullName -RedirectStandardInput
"@
        
        # Use a simpler approach: create a batch file that uses echo to pipe password
        # But OpenSSH doesn't accept password from stdin directly
        # Better approach: use plink.exe from PuTTY if available, or use expect-like script
        
        # For now, use direct process execution with manual password handling
        # We'll use a workaround: create a temporary VBScript or use PowerShell's Start-Process with input
        
        # Alternative: Use ssh with -o BatchMode=no and handle password prompt manually
        # But this is complex. Let's try using plink.exe from PuTTY as fallback if available
        
        $plinkPath = "${env:ProgramFiles}\PuTTY\plink.exe"
        if (Test-Path $plinkPath) {
            Write-Host "Using PuTTY plink.exe (supports password parameter)" -ForegroundColor Gray
            $plinkArgs = @(
                "-ssh",
                "-P", $Session.Port.ToString(),
                "-l", $Session.Username,
                "-pw", $passwordPlain,
                "-batch",
                "-T",
                $Session.Host,
                $Command
            )
            
            $processInfo = New-Object System.Diagnostics.ProcessStartInfo
            $processInfo.FileName = $plinkPath
            $processInfo.Arguments = ($plinkArgs -join ' ')
            $processInfo.UseShellExecute = $false
            $processInfo.RedirectStandardOutput = $true
            $processInfo.RedirectStandardError = $true
            $processInfo.CreateNoWindow = $true
            $processInfo.StandardOutputEncoding = [System.Text.Encoding]::UTF8
            $processInfo.StandardErrorEncoding = [System.Text.Encoding]::UTF8
            
            $process = New-Object System.Diagnostics.Process
            $process.StartInfo = $processInfo
            $process.Start() | Out-Null
            $completed = $process.WaitForExit($TimeoutSeconds * 1000)
            
            if (-not $completed) {
                $process.Kill()
                Write-Host "Command timeout after $TimeoutSeconds seconds" -ForegroundColor Yellow
                $passwordPlain = $null
                return $null
            }
            
            $output = $process.StandardOutput.ReadToEnd()
            $errorOutput = $process.StandardError.ReadToEnd()
        }
        else {
            # Use OpenSSH with expect-like approach using PowerShell
            # Create a script that handles password prompt
            Write-Host "Using OpenSSH with interactive password handling" -ForegroundColor Gray
            
            $tempScript = [System.IO.Path]::GetTempFileName() + ".ps1"
            $timeoutMs = $TimeoutSeconds * 1000
            $scriptContent = @"
`$ErrorActionPreference = 'Stop'
`$password = '$passwordPlain'
`$sshPath = '$($Session.SshPath)'
`$sshArgs = @('-T', '-o', 'StrictHostKeyChecking=no', '-o', 'UserKnownHostsFile=NUL', '-o', 'LogLevel=ERROR', '-p', '$($Session.Port)', '$($Session.Username)@$($Session.Host)', '$($Command -replace "'", "''")')

`$psi = New-Object System.Diagnostics.ProcessStartInfo
`$psi.FileName = `$sshPath
`$psi.Arguments = `$sshArgs -join ' '
`$psi.UseShellExecute = `$false
`$psi.RedirectStandardOutput = `$true
`$psi.RedirectStandardError = `$true
`$psi.RedirectStandardInput = `$true
`$psi.CreateNoWindow = `$true
`$psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
`$psi.StandardErrorEncoding = [System.Text.Encoding]::UTF8

`$proc = New-Object System.Diagnostics.Process
`$proc.StartInfo = `$psi

# Enable asynchronous output reading
`$outputBuilder = New-Object System.Text.StringBuilder
`$errorBuilder = New-Object System.Text.StringBuilder

`$outputEvent = Register-ObjectEvent -InputObject `$proc -EventName OutputDataReceived -Action {
    if (`$EventArgs.Data) {
        [void]`$Event.MessageData.AppendLine(`$EventArgs.Data)
    }
} -MessageData `$outputBuilder

`$errorEvent = Register-ObjectEvent -InputObject `$proc -EventName ErrorDataReceived -Action {
    if (`$EventArgs.Data) {
        [void]`$Event.MessageData.AppendLine(`$EventArgs.Data)
    }
} -MessageData `$errorBuilder

`$proc.Start() | Out-Null
`$proc.BeginOutputReadLine()
`$proc.BeginErrorReadLine()

# Wait a bit for password prompt
Start-Sleep -Milliseconds 500

# Send password
`$proc.StandardInput.WriteLine(`$password)
`$proc.StandardInput.Close()

# Wait for completion with timeout
`$timeout = $timeoutMs
`$completed = `$proc.WaitForExit(`$timeout)

if (-not `$completed) {
    `$proc.Kill()
    Write-Error "Command timeout"
    exit 1
}

# Wait a bit more for async output to be captured
Start-Sleep -Milliseconds 500

# Get output
`$output = `$outputBuilder.ToString()
`$error = `$errorBuilder.ToString()

# Clean up events
Unregister-Event -SourceIdentifier `$outputEvent.Name
Unregister-Event -SourceIdentifier `$errorEvent.Name
Remove-Event -SourceIdentifier `$outputEvent.Name -ErrorAction SilentlyContinue
Remove-Event -SourceIdentifier `$errorEvent.Name -ErrorAction SilentlyContinue

# Combine output and error
if ([string]::IsNullOrWhiteSpace(`$output)) {
    `$output = `$error
}
elseif (-not [string]::IsNullOrWhiteSpace(`$error)) {
    `$output += "`n" + `$error
}

Write-Output `$output
"@
            Set-Content -Path $tempScript -Value $scriptContent -Encoding UTF8
            
            try {
                $output = & powershell.exe -ExecutionPolicy Bypass -File $tempScript
                $errorOutput = ""
            }
            finally {
                Remove-Item $tempScript -ErrorAction SilentlyContinue
            }
        }
        
        # Debug: show raw output before cleaning
        if ($output) {
            Write-Host "Raw SSH output (first 500 chars): $($output.Substring(0, [Math]::Min(500, $output.Length)))" -ForegroundColor DarkGray
        }
        
        # Clean output - remove command echo and prompt
        # Remove password prompt lines if present
        $cleanOutput = $output -replace ".*password:.*", ''
        $cleanOutput = $cleanOutput -replace ".*Password:.*", ''
        $cleanOutput = $cleanOutput -replace ".*пароль:.*", ''
        
        # Remove command echo (handle both with and without newline)
        $commandEscaped = [regex]::Escape($Command)
        $cleanOutput = $cleanOutput -replace "^$commandEscaped(`r?`n|`n)", ''
        $cleanOutput = $cleanOutput -replace "^$commandEscaped\s*", ''
        
        # Remove prompt (designer>) - handle multiple occurrences
        $cleanOutput = $cleanOutput -replace 'designer>\s*', ''
        $cleanOutput = $cleanOutput -replace 'designer>', ''
        
        # Remove any leading/trailing whitespace and newlines
        $cleanOutput = $cleanOutput.Trim()
        
        # Debug: show cleaned output
        if ($cleanOutput) {
            Write-Host "Cleaned output (first 500 chars): $($cleanOutput.Substring(0, [Math]::Min(500, $cleanOutput.Length)))" -ForegroundColor DarkGray
        }
        
        # Clear password from memory
        $passwordPlain = $null
        [System.GC]::Collect()
        
        return $cleanOutput
    }
    catch {
        Write-Host "Error executing command: $_" -ForegroundColor Red
        return $null
    }
}

# Function to stop 1C agent
function Stop-1CAgent {
    param(
        [object]$Session,
        [object]$Process
    )
    
    if ($Session -and $Session.Connected) {
        Write-Host "Shutting down agent..." -ForegroundColor Cyan
        try {
            $result = Invoke-1CAgentCommand -Session $Session -Command "common shutdown" -TimeoutSeconds 10
            Start-Sleep -Seconds 2
        }
        catch {
            Write-Host "Warning: Error sending shutdown command: $_" -ForegroundColor Yellow
        }
        
        # Clear password from session object
        if ($Session.Password) {
            $Session.Password.Dispose()
        }
    }
    
    if ($Process -and -not $Process.HasExited) {
        Write-Host "Terminating agent process..." -ForegroundColor Cyan
        try {
            Stop-Process -Id $Process.Id -Force -ErrorAction SilentlyContinue
        }
        catch {
            Write-Host "Warning: Error terminating process: $_" -ForegroundColor Yellow
        }
    }
}

# Function to build 1C project
function Build-1CProject {
    param(
        [string]$BaseConnect,
        [string]$ProjectPath,
        [string]$ProjectName,
        [string]$ProjectType,
        [string]$XmlFile,
        [switch]$UseAgent,
        [int]$AgentPort = 1543
    )
    
    $configurator = "C:\Program Files\1cv8\common\1cestart.exe"
    
    if (-not (Test-Path $configurator)) {
        Write-Host "Error: 1C configurator not found at path: $configurator" -ForegroundColor Red
        return $false
    }
    
    $outputFile = ""
    $fullProjectPath = if ([System.IO.Path]::IsPathRooted($ProjectPath)) { $ProjectPath } else { (Resolve-Path $ProjectPath).Path }
    
    if ($ProjectType -eq "Extension") {
        $outputFile = Join-Path $fullProjectPath "$ProjectName.cfe"
    }
    elseif ($ProjectType -eq "Processing") {
        $outputFile = Join-Path $fullProjectPath "$ProjectName.epf"
    }
    else {
        Write-Host "Error: unknown project type: $ProjectType" -ForegroundColor Red
        return $false
    }
    
    Write-Host "`nExecuting build..." -ForegroundColor Cyan
    Write-Host "Project type: $ProjectType" -ForegroundColor Gray
    Write-Host "Project name: $ProjectName" -ForegroundColor Gray
    Write-Host "Output file: $outputFile" -ForegroundColor Gray
    Write-Host "Mode: $(if ($UseAgent) { 'Agent' } else { 'Standard' })" -ForegroundColor Gray
    Write-Host ""
    
    if ($UseAgent) {
        return Build-1CProjectViaAgent -BaseConnect $BaseConnect -ProjectPath $fullProjectPath -ProjectName $ProjectName -ProjectType $ProjectType -XmlFile $XmlFile -OutputFile $outputFile -AgentPort $AgentPort
    }
    else {
        return Build-1CProjectStandard -BaseConnect $BaseConnect -ProjectPath $fullProjectPath -ProjectName $ProjectName -ProjectType $ProjectType -XmlFile $XmlFile -OutputFile $outputFile -Configurator $configurator
    }
}

# Function to build 1C project using standard mode
function Build-1CProjectStandard {
    param(
        [string]$BaseConnect,
        [string]$ProjectPath,
        [string]$ProjectName,
        [string]$ProjectType,
        [string]$XmlFile,
        [string]$OutputFile,
        [string]$Configurator
    )
    
    $command = ""
    $connectionParam = Convert-1CConnectionString -ConnectString $BaseConnect
    
    if ($ProjectType -eq "Extension") {
        $command = "DESIGNER $connectionParam /LoadConfigFromFiles `"$ProjectPath`" /Extension `"$ProjectName`" /DumpCfg `"$OutputFile`" /Exit"
    }
    elseif ($ProjectType -eq "Processing") {
        $xmlFilePath = if ([System.IO.Path]::IsPathRooted($XmlFile)) { $XmlFile } else { (Resolve-Path $XmlFile).Path }
        $command = "DESIGNER $connectionParam /LoadExternalDataProcessorOrReportFromFiles `"$xmlFilePath`" `"$OutputFile`" /Exit"
    }
    
    Write-Host "Command: $Configurator $command" -ForegroundColor Gray
    Write-Host ""
    
    try {
        $process = Start-Process -FilePath $Configurator -ArgumentList $command -Wait -PassThru -NoNewWindow
        
        if ($process.ExitCode -eq 0) {
            if (Test-Path $OutputFile) {
                Write-Host "`nBuild completed successfully!" -ForegroundColor Green
                Write-Host "Result saved to: $OutputFile" -ForegroundColor Green
                return $true
            }
            else {
                Write-Host "`nError: output file not created" -ForegroundColor Red
                return $false
            }
        }
        else {
            Write-Host "`nError executing command. Exit code: $($process.ExitCode)" -ForegroundColor Red
            return $false
        }
    }
    catch {
        Write-Host "`nError executing command: $_" -ForegroundColor Red
        return $false
    }
}

# Function to build 1C project using agent mode
function Build-1CProjectViaAgent {
    param(
        [string]$BaseConnect,
        [string]$ProjectPath,
        [string]$ProjectName,
        [string]$ProjectType,
        [string]$XmlFile,
        [string]$OutputFile,
        [int]$AgentPort = 1543
    )
    
    $agentProcess = $null
    $sshSession = $null
    
    try {
        # Start agent
        $agentProcess = Start-1CAgent -BaseConnect $BaseConnect -Port $AgentPort
        if (-not $agentProcess) {
            Write-Host "Failed to start agent" -ForegroundColor Red
            return $false
        }
        
        # Wait for agent to be ready and verify port is listening
        Write-Host "Waiting for agent to be ready..." -ForegroundColor Cyan
        $portReady = $false
        $maxWaitTime = 10
        $waitInterval = 1
        
        for ($i = 0; $i -lt $maxWaitTime; $i++) {
            Start-Sleep -Seconds $waitInterval
            
            # Check if port is listening
            try {
                $tcpClient = New-Object System.Net.Sockets.TcpClient
                $connect = $tcpClient.BeginConnect("localhost", $AgentPort, $null, $null)
                $wait = $connect.AsyncWaitHandle.WaitOne(500, $false)
                if ($wait) {
                    $tcpClient.EndConnect($connect)
                    $tcpClient.Close()
                    $portReady = $true
                    Write-Host "Port $AgentPort is listening" -ForegroundColor Green
                    break
                }
                else {
                    $tcpClient.Close()
                }
            }
            catch {
                # Port not ready yet
            }
        }
        
        if (-not $portReady) {
            Write-Host "Warning: Port $AgentPort is not responding after $maxWaitTime seconds" -ForegroundColor Yellow
            Write-Host "This may indicate:" -ForegroundColor Yellow
            Write-Host "  - Agent failed to start (check agent window for errors)" -ForegroundColor Yellow
            Write-Host "  - Agent is using a different port" -ForegroundColor Yellow
            Write-Host "  - Firewall is blocking the port" -ForegroundColor Yellow
            Write-Host "Attempting to connect anyway..." -ForegroundColor Yellow
        }
        
        # Connect via SSH (with one retry if needed)
        $sshSession = Connect-1CAgentSSH -AgentHost "localhost" -Port $AgentPort
        if (-not $sshSession) {
            Write-Host "Retrying connection..." -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            $sshSession = Connect-1CAgentSSH -AgentHost "localhost" -Port $AgentPort
        }
        
        if (-not $sshSession) {
            Write-Host "Failed to connect to agent after multiple attempts" -ForegroundColor Red
            Write-Host ""
            Write-Host "Diagnostics:" -ForegroundColor Cyan
            Write-Host "  - Port checked: $AgentPort" -ForegroundColor Gray
            Write-Host "  - Host: localhost" -ForegroundColor Gray
            
            # Check if any 1cv8 processes are running
            $processes = Get-Process -Name "1cv8" -ErrorAction SilentlyContinue
            if ($processes) {
                Write-Host "  - Found $($processes.Count) 1cv8 process(es) running" -ForegroundColor Gray
            }
            else {
                Write-Host "  - No 1cv8 processes found" -ForegroundColor Yellow
            }
            
            Write-Host ""
            Write-Host "Please check:" -ForegroundColor Yellow
            Write-Host "  1. Agent window is visible and shows no errors" -ForegroundColor Yellow
            Write-Host "  2. Port $AgentPort is not blocked by firewall" -ForegroundColor Yellow
            Write-Host "  3. Agent was started with correct port: /AgentSSHPort=$AgentPort" -ForegroundColor Yellow
            Write-Host "  4. Check agent window for any error messages" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "You can also try connecting manually with PuTTY or another SSH client:" -ForegroundColor Cyan
            Write-Host "  Host: localhost" -ForegroundColor Gray
            Write-Host "  Port: $AgentPort" -ForegroundColor Gray
            Write-Host "  Username: admin (or your 1C username)" -ForegroundColor Gray
            
            Stop-1CAgent -Session $null -Process $agentProcess
            return $false
        }
        
        # Set JSON output format
        Write-Host "Setting JSON output format..." -ForegroundColor Cyan
        $result = Invoke-1CAgentCommand -Session $sshSession -Command "options set --output-format json"
        
        if ([string]::IsNullOrWhiteSpace($result)) {
            Write-Host "Warning: Empty response from options set command" -ForegroundColor Yellow
            Write-Host "This may be normal - continuing anyway..." -ForegroundColor Gray
        }
        else {
            Write-Host "Response from options set: $($result.Substring(0, [Math]::Min(200, $result.Length)))" -ForegroundColor Gray
            $jsonResponse = Parse-AgentResponse -Response $result
            
            if ($jsonResponse) {
                if (-not (Test-AgentCommandSuccess -JsonResponse $jsonResponse)) {
                    Write-Host "Failed to set JSON output format" -ForegroundColor Red
                    Write-Host "Response: $result" -ForegroundColor Gray
                    Stop-1CAgent -Session $sshSession -Process $agentProcess
                    return $false
                }
                else {
                    Write-Host "JSON output format set successfully" -ForegroundColor Green
                }
            }
            else {
                # If response is not JSON, it might be that the command succeeded but response is in text format
                # Check if response contains success indicators
                if ($result -match "success|ok|успешно" -or [string]::IsNullOrWhiteSpace($result)) {
                    Write-Host "JSON format set (response in text format, which is expected before JSON is enabled)" -ForegroundColor Green
                }
                else {
                    Write-Host "Warning: Could not parse response as JSON, but continuing..." -ForegroundColor Yellow
                    Write-Host "Response: $result" -ForegroundColor Gray
                }
            }
        }
        
        # Connect to information base
        Write-Host "Connecting to information base..." -ForegroundColor Cyan
        $result = Invoke-1CAgentCommand -Session $sshSession -Command "common connect-ib"
        $jsonResponse = Parse-AgentResponse -Response $result
        if (-not (Test-AgentCommandSuccess -JsonResponse $jsonResponse)) {
            Write-Host "Failed to connect to information base" -ForegroundColor Red
            Stop-1CAgent -Session $sshSession -Process $agentProcess
            return $false
        }
        
        # Build processing from files
        if ($ProjectType -eq "Processing") {
            $xmlFilePath = if ([System.IO.Path]::IsPathRooted($XmlFile)) { $XmlFile } else { (Resolve-Path $XmlFile).Path }
            
            Write-Host "Building processing from XML files..." -ForegroundColor Cyan
            Write-Host "XML file: $xmlFilePath" -ForegroundColor Gray
            Write-Host "Output EPF file: $OutputFile" -ForegroundColor Gray
            
            # Load processing from XML files to EPF file
            # This command converts XML to EPF format
            $loadCommand = "config load-external-data-processor-or-report-from-files --file `"$xmlFilePath`" --ext-file `"$OutputFile`""
            $result = Invoke-1CAgentCommand -Session $sshSession -Command $loadCommand -TimeoutSeconds 600
            $jsonResponse = Parse-AgentResponse -Response $result
            if (-not (Test-AgentCommandSuccess -JsonResponse $jsonResponse)) {
                Write-Host "Failed to build processing from files" -ForegroundColor Red
                Stop-1CAgent -Session $sshSession -Process $agentProcess
                return $false
            }
        }
        elseif ($ProjectType -eq "Extension") {
            Write-Host "Extension build via agent is not yet fully implemented" -ForegroundColor Yellow
            # TODO: Implement extension build via agent
            Stop-1CAgent -Session $sshSession -Process $agentProcess
            return $false
        }
        
        # Disconnect from information base
        Write-Host "Disconnecting from information base..." -ForegroundColor Cyan
        $result = Invoke-1CAgentCommand -Session $sshSession -Command "common disconnect-ib"
        $jsonResponse = Parse-AgentResponse -Response $result
        # Don't fail if disconnect has issues
        
        # Check if output file was created
        if (Test-Path $OutputFile) {
            Write-Host "`nBuild completed successfully!" -ForegroundColor Green
            Write-Host "Result saved to: $OutputFile" -ForegroundColor Green
            Stop-1CAgent -Session $sshSession -Process $agentProcess
            return $true
        }
        else {
            Write-Host "`nError: output file not created" -ForegroundColor Red
            Stop-1CAgent -Session $sshSession -Process $agentProcess
            return $false
        }
    }
    catch {
        Write-Host "`nError during agent build: $_" -ForegroundColor Red
        Stop-1CAgent -Session $sshSession -Process $agentProcess
        return $false
    }
}

# Main function
function Main {
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host "   Script for building 1C extensions and processings      " -ForegroundColor Cyan
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host ""
    
    # 1. Get base list
    Write-Host "Reading base list from ibases.v8i..." -ForegroundColor Cyan
    $bases = Get-1CBases
    
    if (-not $bases -or $bases.Count -eq 0) {
        Write-Host "Failed to get base list. Exiting." -ForegroundColor Red
        exit 1
    }
    
    Write-Host "Found bases: $($bases.Count)" -ForegroundColor Green
    
    # 2. Select base
    $selectedBase = Select-1CBase -Bases $bases
    
    if (-not $selectedBase) {
        Write-Host "Base not selected. Exiting." -ForegroundColor Red
        exit 1
    }
    
    Write-Host "`nSelected base: $($selectedBase.Name)" -ForegroundColor Green
    Write-Host "Connection string: $($selectedBase.Connect)" -ForegroundColor Gray
    
    # 3. Select build mode
    Write-Host "`nSelect build mode:" -ForegroundColor Cyan
    Write-Host "1. Standard mode (direct configurator)" -ForegroundColor Yellow
    Write-Host "2. Agent mode (via SSH)" -ForegroundColor Yellow
    
    do {
        $modeSelection = Read-Host "Enter mode number (1-2)"
        $modeIndex = [int]$modeSelection - 1
    } while ($modeIndex -lt 0 -or $modeIndex -gt 1)
    
    $useAgent = ($modeIndex -eq 1)
    
    if ($useAgent) {
        Write-Host "Selected mode: Agent" -ForegroundColor Green
    }
    else {
        Write-Host "Selected mode: Standard" -ForegroundColor Green
    }
    
    # 4. Determine project type
    Write-Host "`nDetermining project type..." -ForegroundColor Cyan
    $projectInfo = Get-ProjectType -ProjectPath "."
    
    if (-not $projectInfo) {
        Write-Host "Error: failed to determine project type. Make sure current directory contains 1C project." -ForegroundColor Red
        exit 1
    }
    
    Write-Host "Project type: $($projectInfo.Type)" -ForegroundColor Green
    Write-Host "XML file: $($projectInfo.XmlFile)" -ForegroundColor Gray
    
    # 5. Extract project name
    Write-Host "`nExtracting project name..." -ForegroundColor Cyan
    $projectName = Get-ProjectName -XmlFile $projectInfo.XmlFile -ProjectType $projectInfo.Type
    
    if (-not $projectName) {
        Write-Host "Error: failed to extract project name." -ForegroundColor Red
        exit 1
    }
    
    Write-Host "Project name: $projectName" -ForegroundColor Green
    
    # 6. Execute build
    if ($useAgent) {
        $result = Build-1CProject -BaseConnect $selectedBase.Connect -ProjectPath "." -ProjectName $projectName -ProjectType $projectInfo.Type -XmlFile $projectInfo.XmlFile -UseAgent -AgentPort 1543
    }
    else {
        $result = Build-1CProject -BaseConnect $selectedBase.Connect -ProjectPath "." -ProjectName $projectName -ProjectType $projectInfo.Type -XmlFile $projectInfo.XmlFile
    }
    
    if ($result) {
        Write-Host "`n==================================================" -ForegroundColor Green
        Write-Host "         Build completed successfully!              " -ForegroundColor Green
        Write-Host "==================================================" -ForegroundColor Green
        exit 0
    }
    else {
        Write-Host "`n==================================================" -ForegroundColor Red
        Write-Host "         Error during project build              " -ForegroundColor Red
        Write-Host "==================================================" -ForegroundColor Red
        exit 1
    }
}

# Run main function
Main

