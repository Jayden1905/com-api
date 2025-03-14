# Use Windows Server Core 2019 with .NET SDK for building
FROM mcr.microsoft.com/dotnet/sdk:8.0-windowsservercore-ltsc2019 AS build
WORKDIR /source

# Copy csproj and restore dependencies
# Note: Make sure your project targets .NET 8.0 to match this container's SDK version
COPY *.csproj .
RUN dotnet restore

# Copy everything else and build the application
COPY . .
RUN dotnet publish -c Release -o /app

# Build the runtime image
FROM mcr.microsoft.com/dotnet/aspnet:8.0-windowsservercore-ltsc2019

# Install PowerPoint requirements - note this requires Office installation
# This is where we'd normally install Office, but it requires licensing and complex install
# In production, consider using a custom image with Office pre-installed

WORKDIR /app
COPY --from=build /app .

# Set environment variables
ENV ASPNETCORE_URLS=http://+:80
ENV ASPNETCORE_ENVIRONMENT=Production

# Expose port 80
EXPOSE 80

# Start the app with the correct filename
ENTRYPOINT ["dotnet", "..dll"]