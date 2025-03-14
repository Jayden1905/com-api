# PowerShell script to help setup Windows Server 2019 for PowerPoint API containerization
# Run with administrative privileges

# Check if running as administrator
Write-Host "Checking for administrative rights..." -ForegroundColor Yellow
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "This script must be run as Administrator. Please restart with elevated privileges." -ForegroundColor Red
    Write-Host "Right-click on PowerShell and select 'Run as Administrator'" -ForegroundColor Red
    exit 1
}
Write-Host "Administrative rights confirmed." -ForegroundColor Green

# Create a marker file to track script progress through server restart
$progressMarker = "$env:TEMP\docker_setup_progress.txt"
$restartNeeded = $false

# Check if we're in the post-restart phase
if (Test-Path $progressMarker) {
    Write-Host "Resuming setup after restart..." -ForegroundColor Cyan
    $phase = Get-Content $progressMarker

    if ($phase -eq "post-restart") {
        # Skip to the Docker installation steps
        Write-Host "Continuing with Docker installation..." -ForegroundColor Cyan
    }
    else {
        Write-Host "Unknown progress state. Starting from beginning." -ForegroundColor Yellow
        Remove-Item $progressMarker -Force
    }
}
else {
    # First run - install prerequisites
    Write-Host "Starting initial setup..." -ForegroundColor Cyan

    # Install Docker features
    Write-Host "Checking and installing Docker prerequisites..." -ForegroundColor Green

    # Check if features are already installed
    $containersFeature = Get-WindowsFeature -Name Containers
    $hypervFeature = Get-WindowsFeature -Name Hyper-V

    if (-not $containersFeature.Installed) {
        Write-Host "Installing Containers feature..." -ForegroundColor Yellow
        Install-WindowsFeature -Name Containers
        $restartNeeded = $true
    } else {
        Write-Host "Containers feature is already installed." -ForegroundColor Green
    }

    if (-not $hypervFeature.Installed) {
        Write-Host "Installing Hyper-V feature..." -ForegroundColor Yellow
        Install-WindowsFeature -Name Hyper-V
        $restartNeeded = $true
    } else {
        Write-Host "Hyper-V feature is already installed." -ForegroundColor Green
    }

    # Create progress marker for post-restart
    if ($restartNeeded) {
        "post-restart" | Out-File $progressMarker -Force

        # Restart is needed after installing features
        Write-Host "Server needs to be restarted before continuing with Docker installation." -ForegroundColor Yellow
        Write-Host "Run this script again after restarting to complete the installation." -ForegroundColor Yellow
        Write-Host "Press Enter to restart now, or Ctrl+C to cancel..." -ForegroundColor Red
        $null = Read-Host
        Restart-Computer -Force
        exit 0
    } else {
        Write-Host "No restart required, continuing with Docker installation..." -ForegroundColor Green
    }
}

# The following runs either after restart or if no restart was needed

# Check if Docker is already installed
$dockerInstalled = $false
try {
    $dockerVer = docker version
    Write-Host "Docker is already installed." -ForegroundColor Green
    $dockerInstalled = $true
} catch {
    Write-Host "Docker is not installed. Installing now..." -ForegroundColor Yellow
}

if (-not $dockerInstalled) {
    # Install Docker
    try {
        Write-Host "Downloading Docker..." -ForegroundColor Green
        Invoke-WebRequest -UseBasicParsing "https://download.docker.com/win/static/stable/x86_64/docker-20.10.22.zip" -OutFile "$env:TEMP\docker.zip"

        Write-Host "Extracting Docker..." -ForegroundColor Green
        Expand-Archive -Path "$env:TEMP\docker.zip" -DestinationPath "$env:ProgramFiles" -Force

        Write-Host "Adding Docker to PATH..." -ForegroundColor Green
        $env:path += ";$env:ProgramFiles\docker"
        [Environment]::SetEnvironmentVariable("Path", $env:Path, [EnvironmentVariableTarget]::Machine)

        # Register the Docker service
        Write-Host "Registering Docker as a service..." -ForegroundColor Green
        & "$env:ProgramFiles\docker\dockerd.exe" --register-service

        # Start Docker service
        Write-Host "Starting Docker service..." -ForegroundColor Green
        Start-Service docker

        # Wait for Docker to start
        Write-Host "Waiting for Docker to initialize..." -ForegroundColor Yellow
        Start-Sleep -Seconds 10
    }
    catch {
        Write-Host "Error installing Docker: $_" -ForegroundColor Red
        exit 1
    }
}

# Check if Docker Compose is already installed
$composeInstalled = $false
try {
    $composeVer = docker-compose --version
    Write-Host "Docker Compose is already installed." -ForegroundColor Green
    $composeInstalled = $true
} catch {
    Write-Host "Docker Compose is not installed. Installing now..." -ForegroundColor Yellow
}

if (-not $composeInstalled) {
    # Install Docker Compose
    try {
        Write-Host "Installing Docker Compose..." -ForegroundColor Green
        Invoke-WebRequest -UseBasicParsing "https://github.com/docker/compose/releases/download/v2.17.2/docker-compose-Windows-x86_64.exe" -OutFile "$env:ProgramFiles\docker\docker-compose.exe"
    }
    catch {
        Write-Host "Error installing Docker Compose: $_" -ForegroundColor Red
        exit 1
    }
}

# Switch to Windows containers if needed
try {
    Write-Host "Ensuring Docker is using Windows containers..." -ForegroundColor Green
    & "$env:ProgramFiles\Docker\Docker\DockerCli.exe" -SwitchWindowsEngine
    Start-Sleep -Seconds 5
} catch {
    Write-Host "Note: Docker Desktop is not installed, continuing with Docker Engine..." -ForegroundColor Yellow
}

# Verify Docker is running
try {
    Write-Host "Verifying Docker installation..." -ForegroundColor Green
    $dockerVersion = docker version
    Write-Host "Docker is properly installed!" -ForegroundColor Green
} catch {
    Write-Host "Error verifying Docker: $_" -ForegroundColor Red
    Write-Host "Please check Docker installation manually." -ForegroundColor Red
    exit 1
}

try {
    Write-Host "Verifying Docker Compose installation..." -ForegroundColor Green
    $composeVersion = docker-compose --version
    Write-Host "Docker Compose is properly installed!" -ForegroundColor Green
} catch {
    Write-Host "Error verifying Docker Compose: $_" -ForegroundColor Red
    Write-Host "Please check Docker Compose installation manually." -ForegroundColor Red
    exit 1
}

# Create presentation folder for mounting
$presentationsFolder = "./presentations"
If (-not (Test-Path $presentationsFolder)) {
    Write-Host "Creating presentations directory..." -ForegroundColor Green
    New-Item -ItemType Directory -Path $presentationsFolder -Force
}

# Clean up progress marker
if (Test-Path $progressMarker) {
    Remove-Item $progressMarker -Force
}

# Check if Office is installed
$officeInstalled = $false
$officePaths = @(
    "C:\Program Files\Microsoft Office",
    "C:\Program Files (x86)\Microsoft Office"
)

foreach ($path in $officePaths) {
    if (Test-Path $path) {
        $officeInstalled = $true
        break
    }
}

if ($officeInstalled) {
    Write-Host "Microsoft Office appears to be installed on this system." -ForegroundColor Green
} else {
    Write-Host "WARNING: Microsoft Office does not appear to be installed!" -ForegroundColor Red
    Write-Host "You must install Microsoft Office/PowerPoint for the COM automation to work." -ForegroundColor Red
}

Write-Host "`n=== Setup Complete! ===" -ForegroundColor Cyan
Write-Host "Docker and Docker Compose have been successfully installed." -ForegroundColor Green
Write-Host "`nIMPORTANT REMINDERS:" -ForegroundColor Yellow
Write-Host "1. Microsoft Office / PowerPoint must be installed on the host for COM automation to work." -ForegroundColor Yellow
Write-Host "2. Configure PowerPoint Trust Center settings to allow VBA project access." -ForegroundColor Yellow
Write-Host "3. Place your PowerPoint files in the '$presentationsFolder' directory." -ForegroundColor Yellow
Write-Host "`nTo start the PowerPoint API service, run:" -ForegroundColor Cyan
Write-Host "docker-compose up -d" -ForegroundColor White