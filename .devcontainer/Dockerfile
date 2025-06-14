# Use PowerShell base image (Debian-based)
FROM mcr.microsoft.com/powershell:latest

# Use bash for apt operations
SHELL ["/bin/bash", "-c"]
   
# Install Git
RUN apt-get update && \
    apt-get install -y git

# Install .NET SDK
RUN apt-get update && \
    apt-get install -y wget apt-transport-https && \
    wget https://packages.microsoft.com/config/debian/12/packages-microsoft-prod.deb -O packages-microsoft-prod.deb && \
    dpkg -i packages-microsoft-prod.deb && \
    apt-get update && \
    apt-get install -y dotnet-sdk-8.0 && \
    rm -rf packages-microsoft-prod.deb

# Delete the apt cache
RUN apt-get clean && \
    rm -rf /var/lib/apt/lists/*

# Switch shell back to pwsh (optional)
SHELL ["pwsh", "-Command"]