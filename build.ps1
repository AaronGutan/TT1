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

# Function to build 1C project
function Build-1CProject {
    param(
        [string]$BaseConnect,
        [string]$ProjectPath,
        [string]$ProjectName,
        [string]$ProjectType,
        [string]$XmlFile
    )
    
    $configurator = "C:\Program Files\1cv8\common\1cestart.exe"
    
    if (-not (Test-Path $configurator)) {
        Write-Host "Error: 1C configurator not found at path: $configurator" -ForegroundColor Red
        return $false
    }
    
    $outputFile = ""
    $command = ""
    
    $fullProjectPath = if ([System.IO.Path]::IsPathRooted($ProjectPath)) { $ProjectPath } else { (Resolve-Path $ProjectPath).Path }
    
    # Convert connection string to command line format
    $connectionParam = Convert-1CConnectionString -ConnectString $BaseConnect
    
    if ($ProjectType -eq "Extension") {
        $outputFile = Join-Path $fullProjectPath "$ProjectName.cfe"
        $command = "DESIGNER $connectionParam /LoadConfigFromFiles `"$fullProjectPath`" /Extension `"$ProjectName`" /DumpCfg `"$outputFile`" /Exit"
    }
    elseif ($ProjectType -eq "Processing") {
        $outputFile = Join-Path $fullProjectPath "$ProjectName.epf"
        $xmlFilePath = if ([System.IO.Path]::IsPathRooted($XmlFile)) { $XmlFile } else { (Resolve-Path $XmlFile).Path }
        $command = "DESIGNER $connectionParam /LoadExternalDataProcessorOrReportFromFiles `"$xmlFilePath`" `"$outputFile`" /Exit"
    }
    else {
        Write-Host "Error: unknown project type: $ProjectType" -ForegroundColor Red
        return $false
    }
    
    Write-Host "`nExecuting build..." -ForegroundColor Cyan
    Write-Host "Project type: $ProjectType" -ForegroundColor Gray
    Write-Host "Project name: $ProjectName" -ForegroundColor Gray
    Write-Host "Output file: $outputFile" -ForegroundColor Gray
    Write-Host "`nCommand: $configurator $command" -ForegroundColor Gray
    Write-Host ""
    
    try
    {
        $process = Start-Process -FilePath $configurator -ArgumentList $command -Wait -PassThru -NoNewWindow
        
        if ($process.ExitCode -eq 0)
        {
            if (Test-Path $outputFile)
            {
                Write-Host "`nBuild completed successfully!" -ForegroundColor Green
                Write-Host "Result saved to: $outputFile" -ForegroundColor Green
                return $true
            }
            else
            {
                Write-Host "`nError: output file not created" -ForegroundColor Red
                return $false
            }
        }
        else
        {
            Write-Host "`nError executing command. Exit code: $($process.ExitCode)" -ForegroundColor Red
            return $false
        }
    }
    catch
    {
        Write-Host "`nError executing command: $_" -ForegroundColor Red
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
    
    # 3. Determine project type
    Write-Host "`nDetermining project type..." -ForegroundColor Cyan
    $projectInfo = Get-ProjectType -ProjectPath "."
    
    if (-not $projectInfo) {
        Write-Host "Error: failed to determine project type. Make sure current directory contains 1C project." -ForegroundColor Red
        exit 1
    }
    
    Write-Host "Project type: $($projectInfo.Type)" -ForegroundColor Green
    Write-Host "XML file: $($projectInfo.XmlFile)" -ForegroundColor Gray
    
    # 4. Extract project name
    Write-Host "`nExtracting project name..." -ForegroundColor Cyan
    $projectName = Get-ProjectName -XmlFile $projectInfo.XmlFile -ProjectType $projectInfo.Type
    
    if (-not $projectName) {
        Write-Host "Error: failed to extract project name." -ForegroundColor Red
        exit 1
    }
    
    Write-Host "Project name: $projectName" -ForegroundColor Green
    
    # 5. Execute build
    $result = Build-1CProject -BaseConnect $selectedBase.Connect -ProjectPath "." -ProjectName $projectName -ProjectType $projectInfo.Type -XmlFile $projectInfo.XmlFile
    
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
