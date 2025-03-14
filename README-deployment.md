# PowerPoint API Deployment Guide for Windows Server 2019

This guide explains how to deploy the PowerPoint API service on Windows Server 2019 using Docker containers.

## Important Considerations

**⚠️ IMPORTANT:** This application uses COM automation to control PowerPoint, which requires special considerations for containerization:

1. Microsoft Office (PowerPoint) must be installed on the host server
2. The container must have access to the host's COM services
3. This setup requires Windows containers (not Linux containers)
4. The container needs appropriate permissions to interact with Office applications

## Prerequisites

- Windows Server 2019 with latest updates
- PowerShell 5.1 or higher (included with Windows Server 2019)
- Microsoft PowerPoint installed on the host server
- Administrative access to the server

## Installation Steps

### 1. Prepare the Environment

1. Clone this repository to your Windows Server 2019 machine
2. Run the included setup script with administrative privileges:

```powershell
# Open PowerShell as Administrator
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
# Then run the setup script
.\setup-server.ps1
```

3. The server will restart during the setup process. After restart, run the script again to complete the installation.

### 2. Configure Microsoft Office

1. Ensure Microsoft Office (particularly PowerPoint) is installed on the host server
2. Set Office to allow automation:
   - Open PowerPoint
   - Go to File > Options > Trust Center > Trust Center Settings > Macro Settings
   - Check "Trust access to the VBA project object model"
   - Click OK to save the settings

### 3. Prepare Your Presentations

1. Place your PowerPoint presentations in the `./presentations` folder
2. Make sure the presentations are in `.pptx` or `.pptm` format (for macros)

### 4. Build and Start the Container

```powershell
# Build and start the container in detached mode
docker-compose up -d
```

### 5. Verify Deployment

1. The API will be available at `http://localhost:8080`
2. Test the API with an endpoint such as:
   ```
   http://localhost:8080/powerpoint/hasmacros
   ```

## Troubleshooting

### COM Automation Issues

If the API cannot control PowerPoint, check:

1. Microsoft Office is properly installed on the host
2. The container is using process isolation (default in the docker-compose.yml)
3. The API logs for specific error messages
4. PowerPoint Trust Center settings are configured correctly

### Container Access Issues

If the container cannot access PowerPoint:

1. Make sure you're using Windows containers (not Linux)
2. Check Docker is running in Windows container mode
3. Verify the service account has appropriate permissions

### Permission Issues

COM automation may require elevated permissions:

1. Run the Docker service as an account with access to PowerPoint
2. Consider using a domain account with appropriate permissions

## Security Considerations

This setup has security implications:

1. The container has access to COM services on the host
2. Office automation can potentially execute code
3. Consider network isolation for production deployments
4. Use read-only access for presentations whenever possible

## Need Help?

Contact the development team for assistance with deployment issues.

## License

Refer to the project license for usage terms and conditions.
