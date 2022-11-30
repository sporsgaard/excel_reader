# Basic AlarmPeople Net6 Services structure

## construct this using powershell

```powershell
dotnet new classlib -n APService
dotnet new console  -n APConsole
dotnet new worker   -n APWinService
dotnet new xunit    -n APUnitTests
dotnet new sln      -n APSolution -o .

dotnet sln add APService\APService.csproj
dotnet sln add APConsole\APConsole.csproj
dotnet sln add APWinService\APWinService.csproj
dotnet sln add APUnitTests\APUnitTests.csproj

pushd APConsole
dotnet add reference ..\APService\APService.csproj
popd

pushd APWinService
dotnet add reference ..\APService\APService.csproj
popd

pushd APUnitTests
dotnet add reference ..\APService\APService.csproj
popd

# check it all works
dotnet build
```

## Upgrading to newer versions of packages

To ease upgrading packages in projects via VS Code - install the extension:  **nuget package manager gui** extension

Then perform the following:

1. Open your project workspace in VSCode
2. Open the Command Palette (Ctrl+Shift+P)
3. Select > Nuget Package Manager GUI
4. Click Load Package Versions
5. Click Update All Packages
