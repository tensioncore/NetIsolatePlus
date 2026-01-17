# NetIsolate+ — build_readme

## Requirements
- Windows 10/11
- .NET SDK 8.x
- Git (optional)
- Visual Studio 2022 (optional) or VS Code + C# extension

NuGet dependency:
- `System.Management` (restores automatically)

## Build commands (PowerShell)

From the repository root:

### 1) Clean

dotnet clean .\NetIsolate.csproj -c Release

### 2) Restore

dotnet restore .\NetIsolate.csproj

### 3) Publish (single-file, self-contained)

dotnet publish .\NetIsolate.csproj -c Release -r win-x64

### Output

The published EXE will be here:
.\bin\Release\net8.0-windows\win-x64\publish\NetIsolatePlus.exe

### Notes

The app manifest is set to requireAdministrator, so Windows will prompt for UAC when launching.

Release publish settings (single-file, self-contained, compression) are defined in NetIsolate.csproj.