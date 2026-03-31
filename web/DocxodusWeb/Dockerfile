FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src

# Copy solution and project files for restore
COPY Docxodus/Docxodus.csproj Docxodus/
COPY web/DocxodusWeb/DocxodusWeb.csproj web/DocxodusWeb/
RUN dotnet restore web/DocxodusWeb/DocxodusWeb.csproj

# Copy everything and build
COPY Docxodus/ Docxodus/
COPY web/DocxodusWeb/ web/DocxodusWeb/
RUN dotnet publish web/DocxodusWeb/DocxodusWeb.csproj -c Release -o /app/publish

# Runtime image
FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS runtime
WORKDIR /app

# Install native dependencies for SkiaSharp
RUN apt-get update && apt-get install -y --no-install-recommends \
    libfontconfig1 \
    && rm -rf /var/lib/apt/lists/*

COPY --from=build /app/publish .

# Cloud Run uses PORT env var
ENV ASPNETCORE_URLS=http://+:8080
ENV DATA_DIR=/data
EXPOSE 8080

ENTRYPOINT ["dotnet", "DocxodusWeb.dll"]
