param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('x86', 'x64')]
    [string] $Architecture,
    [Parameter(Mandatory = $true)]
    [string] $OutputDirectory
)

$ErrorActionPreference = 'Stop'
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
if (-not (Test-Path -LiteralPath $vswhere)) {
    throw 'Native lifetime tests require Visual Studio C++ Build Tools and a Windows SDK.'
}

$installation = & $vswhere -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath
if (-not $installation) {
    throw 'Install the Desktop development with C++ workload to build the native lifetime test fixture.'
}

$msbuild = Join-Path $installation 'MSBuild\Current\Bin\MSBuild.exe'
$platform = if ($Architecture -eq 'x86') { 'Win32' } else { 'x64' }
$output = [IO.Path]::GetFullPath($OutputDirectory).TrimEnd('\') + '\'
# Isolated by managed target framework and architecture, including intermediate files.
& $msbuild (Join-Path $PSScriptRoot 'NativeLifetimeProbe.vcxproj') /nologo /verbosity:minimal `
    /p:Configuration=Release "/p:Platform=$platform" "/p:NativeProbeOutputDirectory=${output}."
if ($LASTEXITCODE -ne 0) {
    throw "Native lifetime fixture build failed with exit code $LASTEXITCODE."
}
