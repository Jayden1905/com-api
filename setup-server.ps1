# PowerShell script to help setup Windows Server 2019 for PowerPoint API containerization
# Run with administrative privileges

# Install Docker features
Write-Host "Installing Docker prerequisites..." -ForegroundColor Green
Install-WindowsFeature -Name Containers
Install-WindowsFeature -Name Hyper-V

# Restart is needed after installing features
Write-Host "Server needs to be restarted before continuing with Docker installation." -ForegroundColor Yellow
Write-Host "Run the rest of this script after restarting." -ForegroundColor Yellow
Write-Host "Press Enter to restart now, or Ctrl+C to cancel..." -ForegroundColor Red
$null = Read-Host
Restart-Computer -Force

# The following should be run after reboot

# Install Docker
Write-Host "Installing Docker..." -ForegroundColor Green
Invoke-WebRequest -UseBasicParsing "https://download.docker.com/win/static/stable/x86_64/docker-20.10.22.zip" -OutFile "$env:TEMP\docker.zip"
Expand-Archive -Path "$env:TEMP\docker.zip" -DestinationPath "$env:ProgramFiles" -Force
$env:path += ";$env:ProgramFiles\docker"
[Environment]::SetEnvironmentVariable("Path", $env:Path, [EnvironmentVariableTarget]::Machine)

# Register the Docker service
Write-Host "Registering Docker as a service..." -ForegroundColor Green
& "$env:ProgramFiles\docker\dockerd.exe" --register-service

# Start Docker service
Start-Service docker

# Install Docker Compose
Write-Host "Installing Docker Compose..." -ForegroundColor Green
Invoke-WebRequest -UseBasicParsing "https://github.com/docker/compose/releases/download/v2.17.2/docker-compose-Windows-x86_64.exe" -OutFile "$env:ProgramFiles\docker\docker-compose.exe"

# Verify Docker is running
docker version
docker-compose --version

# Create presentation folder for mounting
If (-not (Test-Path "./presentations")) {
    Write-Host "Creating presentations directory..." -ForegroundColor Green
    New-Item -ItemType Directory -Path "./presentations" -Force
}

Write-Host "Done! Docker and Docker Compose have been installed." -ForegroundColor Green
Write-Host "IMPORTANT: Microsoft Office / PowerPoint must be installed on the host for COM automation to work." -ForegroundColor Yellow
Write-Host "Run 'docker-compose up -d' to start the PowerPoint API service." -ForegroundColor Cyan