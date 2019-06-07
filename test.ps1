[CmdletBinding()]
param(
    [string]$Program,
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$ArgumentList = @()
)

$args = @(Join-Path $PSScriptRoot 'sstart.vbs')
if ($program) {
    $args += @($program)
}
$args += $argumentList

try {
    $process = Start-Process -Wait -PassThru wscript.exe -ArgumentList $args
} finally {
    if ($process) {
        Write-Verbose "ExitCode = $($process.ExitCode)";
        $process.Dispose()
    }
}
