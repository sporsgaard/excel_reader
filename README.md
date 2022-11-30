# Basic AlarmPeople Excel file reader

## construct this using powershell

```powershell
dotnet new console  -n excel_reader -o .

dotnet add package NPOI

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
