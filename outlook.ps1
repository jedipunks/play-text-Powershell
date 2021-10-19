try
{
    Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
}
catch { Throw "Could not add the assembly." }

get-process OUTLOOK -ErrorAction SilentlyContinue | Stop-Process

$Outlook = New-Object -comobject Outlook.Application

$namespace = $Outlook.GetNameSpace("MAPI")

$folder = $namespace.getDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderCalendar)

$folder.items | Sort-Object -Property start -Descending | Select-Object -First 5 -Property Subject, Start, Duration, Location

get-process OUTLOOK -ErrorAction SilentlyContinue | Stop-Process
