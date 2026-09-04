[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$AddinsRoot,
    [Parameter(Mandatory = $true)][string]$ExcelOptionsKey,
    [Parameter(Mandatory = $true)][string]$AddinManagerKey,
    [Parameter(Mandatory = $true)][string]$ExcelProcessName
)

Set-ItemProperty -Path $ExcelOptionsKey -Name "OPEN" -Value '"C:\\Broken\invSys.Operations.xlam"' -Type String
throw "Injected registration failure."
