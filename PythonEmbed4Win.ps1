#!powershell
#
# PythonEmbed4Win.ps1
#
# To run:
#
#     PS> Invoke-WebRequest -Uri https://raw.githubusercontent.com/jtmoon79/PythonEmbed4Win/main/PythonEmbed4Win.ps1 -OutFile PythonEmbed4Win.ps1
#     PS> Get-Help .\PythonEmbed4Win.ps1 -full
#

<#
.SYNOPSIS
    Quickly setup a portable python environment for Windows using an embed.zip.
.DESCRIPTION
    Quickly setup a portable self-referential python environment for Windows. No
    prior python installation is needed. The python code is downloaded
    from the web (https://www.python.org/ftp/python/).

    If no -Version is passed then the latest Python version is chosen.
    If only Version major.minor passed then chooses latest major.minor.micro version.
    e.g. passing "-Version 3.8" will choose Python 3.8.12.

    The installation uses the Windows "Embedded" distribution zip file,
    e.g. python-3.8.12-embed-amd64.zip
    That zip file distribution requires tedious and non-obvious steps.
    This script adjusts the installation to be runnable and isolated (removes
    references to python paths outside it's own directory).
    This script also installs latest pip.

    Python 3.6 and later is supported.

    -UriCheck is merely a self-test to see which URIs for embed.zip files
    are valid.

    The installed Python distribution is like a Python Virtual Environment but
    technically is not. It does not set environment variable VIRTUAL_ENV nor
    modify the PATH.
    Users of this installation must call the `python.exe` executable. Do not
    call other modules by their script entrypoint,
    e.g. for using pip, do not
        C:/path/to/embed/Scripts/pip.exe install ...
    do
        C:/path/to/embed/Scripts/python.exe -m pip install ...

    Inspired by this stackoverflow.com question:
    https://stackoverflow.com/questions/68958635/python-windows-embeddable-package-fails-to-run-no-module-named-pip-the-system/68958636

    BUG: Some embed.zip for a few releases are hardcoded to look for other
         Windows Python installations.
         For example, if you install Python-3.8.6.msi and then run this script to
         install python-3.8.4-embed-amd64.zip, the embed Python 3.8.4 sys.path
         will be the confusing paths:
             C:\python-embed-3.8.4\python38.zip
             C:\python-msi-install-3.8.6\Lib
             C:\python-msi-install-3.8.6\DLLs
             C:\python-embed-3.8.4
             C:\python-msi-install-3.8.6
             C:\python-msi-install-3.8.6\lib\site-packages
          The python._pth and sitecustomize.py seem to have no affect.
          As of November 2023, the latest version of supported Pythons,
          3.6 to 3.12, appear to behave correctly. It only affects a few
          intermediate releases.

.PARAMETER Version
    Version of Python to install. Leave blank to fetch the latest Python.
    Can pass major.minor.micro or just major.minor, e.g. "3.8.2" or "3.8".
    If passed only major.minor then the latest major.minor.micro will be chosen.
    Python 3.6 and later is supported.
.PARAMETER Path
    Install to this path. Defaults to a descriptive name.
.PARAMETER Arch
    Architecture: win32 or amd64. Defaults to the current architecture.
.PARAMETER SkipExec
    Do not execute python.exe after installation.
    This skips the python.exe self-test and the run of `get-pip.py`.
.PARAMETER UriCheck
    Only check pre-filled URIs (script self-test). Does not install Python.
.LINK
    https://github.com/jtmoon79/PythonEmbed4Win
.NOTES
    Author: James Thomas Moon
#>
[Cmdletbinding(DefaultParameterSetName = 'Install')]
Param (
    [Parameter(ParameterSetName = 'Install')]
    [System.IO.FileInfo] $Path,
    [Parameter(ParameterSetName = 'Install')]
    [String] $Version,
    # TODO: how to set a script Param to custom Enum type `Archs`?
    #       placing the definition of `Archs` before this Param declaration
    #       will cause an error.
    [Parameter(ParameterSetName = 'Install')]
    [Parameter(ParameterSetName = 'UriCheck')]
    [ValidateSet('win32','amd64')]
    [String] $Arch,
    [Parameter(ParameterSetName = 'Install')]
    [switch] $SkipExec,
    [Parameter(ParameterSetName = 'UriCheck')]
    [switch] $UriCheck
)

$stopWatch = New-Object -TypeName System.Diagnostics.Stopwatch
$stopWatch.Start()

New-Variable -Name SCRIPT_NAME -Value "PythonEmbed4Win.ps1" -Option ReadOnly -Force
$SecurityProtocolType = [Net.SecurityProtocolType]::Tls12

Enum Archs {
    # enum names are the literal substrings within the named .zip file
    # e.g. "win32" from "python-3.8.12-embed-win32.zip"
    win32
    amd64
}
$arch_default = ${env:PROCESSOR_ARCHITECTURE}.ToLower()

New-Variable -Name URI_GETPIP -Option ReadOnly -Force -Value ([URI] "https://bootstrap.pypa.io/get-pip.py")
New-Variable -Name URI_GETPIP36 -Option ReadOnly -Force -Value ([URI] "https://bootstrap.pypa.io/pip/3.6/get-pip.py")
New-Variable -Name URI_PYTHON_VERSIONS -Option ReadOnly -Force -Value ([URI] "https://www.python.org/ftp/python")

function URI-Combine
{
    [OutputType([URI])]
    Param(
        [Parameter(Mandatory=$true)][URI]$uri,
        [Parameter(Mandatory=$true)][String]$append
    )
    return [URI]($uri.ToString() + $append.ToString())
}

# Pre-fill known URIs that exist as of Dec. 2023, i.e. return HTTP 200.
# This allows skipping some work of scraping URIs from the root page and makes
# this script a little faster.
#
# These can be reviewed by passing `-UriCheck` to this script.
$URIs_200 = @(
    # *win32.zip
    URI-Combine $URI_PYTHON_VERSIONS '/3.5.0/python-3.5.0-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.5.1/python-3.5.1-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.5.2/python-3.5.2-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.5.3/python-3.5.3-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.5.4/python-3.5.4-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.6.0/python-3.6.0-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.6.1/python-3.6.1-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.6.2/python-3.6.2-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.6.3/python-3.6.3-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.6.4/python-3.6.4-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.6.5/python-3.6.5-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.6.6/python-3.6.6-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.6.7/python-3.6.7-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.6.8/python-3.6.8-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.7.0/python-3.7.0-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.7.1/python-3.7.1-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.7.3/python-3.7.3-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.7.4/python-3.7.4-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.7.5/python-3.7.5-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.7.6/python-3.7.6-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.7.7/python-3.7.7-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.7.8/python-3.7.8-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.7.9/python-3.7.9-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.8.0/python-3.8.0-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.8.1/python-3.8.1-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.8.2/python-3.8.2-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.8.3/python-3.8.3-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.8.4/python-3.8.4-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.8.5/python-3.8.5-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.8.6/python-3.8.6-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.8.7/python-3.8.7-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.8.8/python-3.8.8-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.8.9/python-3.8.9-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.8.10/python-3.8.10-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.9.0/python-3.9.0-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.9.1/python-3.9.1-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.9.2/python-3.9.2-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.9.3/python-3.9.3-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.9.4/python-3.9.4-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.9.5/python-3.9.5-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.9.6/python-3.9.6-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.9.7/python-3.9.7-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.9.8/python-3.9.8-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.9.9/python-3.9.9-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.9.10/python-3.9.10-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.10.0/python-3.10.0-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.10.1/python-3.10.1-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.10.2/python-3.10.2-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.10.3/python-3.10.3-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.10.4/python-3.10.4-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.10.5/python-3.10.5-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.10.6/python-3.10.6-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.10.7/python-3.10.7-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.10.8/python-3.10.8-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.10.9/python-3.10.9-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.10.10/python-3.10.10-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.10.11/python-3.10.11-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.11.0/python-3.11.0-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.11.1/python-3.11.1-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.11.2/python-3.11.2-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.11.3/python-3.11.3-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.11.4/python-3.11.4-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.11.5/python-3.11.5-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.11.6/python-3.11.6-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.11.7/python-3.11.7-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.11.8/python-3.11.8-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.11.9/python-3.11.9-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.12.0/python-3.12.0-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.12.1/python-3.12.1-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.12.2/python-3.12.2-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.12.3/python-3.12.3-embed-win32.zip'
    # *amd64.zip
    URI-Combine $URI_PYTHON_VERSIONS '/3.5.0/python-3.5.0-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.5.1/python-3.5.1-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.5.2/python-3.5.2-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.5.3/python-3.5.3-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.5.4/python-3.5.4-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.6.0/python-3.6.0-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.6.1/python-3.6.1-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.6.2/python-3.6.2-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.6.3/python-3.6.3-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.6.4/python-3.6.4-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.6.5/python-3.6.5-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.6.6/python-3.6.6-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.6.7/python-3.6.7-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.6.8/python-3.6.8-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.7.0/python-3.7.0-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.7.1/python-3.7.1-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.7.3/python-3.7.3-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.7.4/python-3.7.4-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.7.5/python-3.7.5-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.7.6/python-3.7.6-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.7.7/python-3.7.7-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.7.8/python-3.7.8-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.7.9/python-3.7.9-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.8.0/python-3.8.0-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.8.1/python-3.8.1-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.8.2/python-3.8.2-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.8.3/python-3.8.3-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.8.4/python-3.8.4-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.8.5/python-3.8.5-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.8.6/python-3.8.6-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.8.7/python-3.8.7-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.8.8/python-3.8.8-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.8.9/python-3.8.9-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.8.10/python-3.8.10-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.9.0/python-3.9.0-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.9.1/python-3.9.1-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.9.2/python-3.9.2-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.9.3/python-3.9.3-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.9.4/python-3.9.4-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.9.5/python-3.9.5-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.9.6/python-3.9.6-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.9.7/python-3.9.7-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.9.8/python-3.9.8-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.9.9/python-3.9.9-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.9.11/python-3.9.11-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.9.12/python-3.9.12-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.9.13/python-3.9.13-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.9.10/python-3.9.10-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.10.0/python-3.10.0-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.10.1/python-3.10.1-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.10.2/python-3.10.2-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.10.3/python-3.10.3-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.10.4/python-3.10.4-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.10.5/python-3.10.5-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.10.6/python-3.10.6-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.10.7/python-3.10.7-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.10.8/python-3.10.8-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.10.9/python-3.10.9-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.10.10/python-3.10.10-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.10.11/python-3.10.11-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.11.0/python-3.11.0-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.11.1/python-3.11.1-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.11.2/python-3.11.2-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.11.3/python-3.11.3-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.11.4/python-3.11.4-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.11.5/python-3.11.5-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.11.6/python-3.11.6-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.11.7/python-3.11.7-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.11.8/python-3.11.8-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.11.9/python-3.11.9-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.12.0/python-3.12.0-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.12.1/python-3.12.1-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.12.2/python-3.12.2-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.12.3/python-3.12.3-embed-amd64.zip'
)

# Pre-fill known URIs that do not exist as of Dec. 2023, i.e. return HTTP 503.
$URIs_503 = @(
    # *win32.zip
    URI-Combine $URI_PYTHON_VERSIONS '/3.0.1/python-3.0.1-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.1.1/python-3.1.1-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.1.2/python-3.1.2-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.1.3/python-3.1.3-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.1.4/python-3.1.4-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.1.5/python-3.1.5-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.2.1/python-3.2.1-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.2.2/python-3.2.2-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.2.3/python-3.2.3-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.2.4/python-3.2.4-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.2.5/python-3.2.5-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.2.6/python-3.2.6-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.3.0/python-3.3.0-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.3.1/python-3.3.1-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.3.2/python-3.3.2-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.3.3/python-3.3.3-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.3.4/python-3.3.4-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.3.5/python-3.3.5-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.3.6/python-3.3.6-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.3.7/python-3.3.7-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.4.0/python-3.4.0-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.4.1/python-3.4.1-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.4.2/python-3.4.2-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.4.3/python-3.4.3-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.4.4/python-3.4.4-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.4.5/python-3.4.5-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.4.6/python-3.4.6-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.4.7/python-3.4.7-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.4.8/python-3.4.8-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.4.9/python-3.4.9-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.4.10/python-3.4.10-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.5.5/python-3.5.5-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.5.6/python-3.5.6-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.5.7/python-3.5.7-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.5.8/python-3.5.8-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.5.9/python-3.5.9-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.5.10/python-3.5.10-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.6.9/python-3.6.9-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.6.10/python-3.6.10-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.6.11/python-3.6.11-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.6.12/python-3.6.12-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.6.13/python-3.6.13-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.6.14/python-3.6.14-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.6.15/python-3.6.15-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.7.2/python-3.7.2-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.7.10/python-3.7.10-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.7.11/python-3.7.11-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.7.12/python-3.7.12-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.8.11/python-3.8.11-embed-win32.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.8.12/python-3.8.12-embed-win32.zip'
    # *amd64.zip
    URI-Combine $URI_PYTHON_VERSIONS '/3.0.1/python-3.0.1-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.1.1/python-3.1.1-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.1.2/python-3.1.2-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.1.3/python-3.1.3-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.1.4/python-3.1.4-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.1.5/python-3.1.5-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.2.1/python-3.2.1-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.2.2/python-3.2.2-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.2.3/python-3.2.3-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.2.4/python-3.2.4-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.2.5/python-3.2.5-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.2.6/python-3.2.6-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.3.0/python-3.3.0-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.3.1/python-3.3.1-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.3.2/python-3.3.2-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.3.3/python-3.3.3-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.3.4/python-3.3.4-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.3.5/python-3.3.5-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.3.6/python-3.3.6-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.3.7/python-3.3.7-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.4.0/python-3.4.0-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.4.1/python-3.4.1-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.4.2/python-3.4.2-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.4.3/python-3.4.3-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.4.4/python-3.4.4-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.4.5/python-3.4.5-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.4.6/python-3.4.6-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.4.7/python-3.4.7-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.4.8/python-3.4.8-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.4.9/python-3.4.9-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.4.10/python-3.4.10-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.5.5/python-3.5.5-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.5.6/python-3.5.6-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.5.7/python-3.5.7-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.5.8/python-3.5.8-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.5.9/python-3.5.9-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.5.10/python-3.5.10-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.6.9/python-3.6.9-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.6.10/python-3.6.10-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.6.11/python-3.6.11-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.6.12/python-3.6.12-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.6.13/python-3.6.13-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.6.14/python-3.6.14-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.6.15/python-3.6.15-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.7.2/python-3.7.2-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.7.10/python-3.7.10-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.7.11/python-3.7.11-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.7.12/python-3.7.12-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.7.13/python-3.7.13-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.7.14/python-3.7.14-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.7.15/python-3.7.15-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.7.16/python-3.7.16-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.7.17/python-3.7.17-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.8.11/python-3.8.11-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.8.12/python-3.8.12-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.8.13/python-3.8.13-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.8.14/python-3.8.14-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.8.15/python-3.8.15-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.8.16/python-3.8.16-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.8.17/python-3.8.17-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.8.18/python-3.8.18-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.9.14/python-3.9.14-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.9.15/python-3.9.15-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.9.16/python-3.9.16-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.9.17/python-3.9.17-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.9.18/python-3.9.18-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.10.12/python-3.10.12-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.10.13/python-3.10.13-embed-amd64.zip'
)

function New-TemporaryDirectory {
    [OutputType([System.IO.FileInfo])]
    Param([String]$extra)
    Join-Path -Path $([System.IO.Path]::GetTempPath()) -ChildPath $extra.ToString() `
        | ForEach-Object { New-Item -Path $_ -ItemType Directory -Force }
}

function Print-File-Nicely {
    <#
    .SYNOPSIS
    Print a file for easy grok.
    #>
    Param([System.IO.FileInfo]$path)

    Write-Host -ForegroundColor Yellow `
        "File:" $path "`n"
    $text = Get-Content $path -Raw
    $textt = "`t" + $text.Replace("`n", "`n`t").TrimEnd("`t")
    Write-Host -ForegroundColor Yellow $textt
}

function Invoke-WebRequest-Head {
    <#
    .SYNOPSIS
    Wrapper to call `Invoke-WebRequest` using HTTP Method HEAD,
    adjusted to the running version of PowerShell.
    #>
    Param(
        [Parameter(Mandatory=$true)][URI]$uri
    )

    if (7 -le $PSVersionTable.PSVersion.Major) {
        # XXX: not entirely sure precisely which Powershell version introduced
        #      the `-SkipHTTPErrorCheck` parameter.
        $result = Invoke-WebRequest -Uri $uri -Method Head -SkipHTTPErrorCheck
    } else {
        try {
            # if the HTTP Status Code is >=400 then this always raises an error
            # even if passed `-ErrorAction SilentlyContinue`. So it must be
            # handled specially.
            $result = Invoke-WebRequest -Uri $uri -Method Head
        } catch {
            $result = [PSCustomObject]@{
                StatusCode = 500
                StatusDescription = "BAD"
            }
        }
    }
    return $result
}

function Download
{
    <#
    .SYNOPSIS
    download file at $uri to $dest, note the time taken
    #>
    Param(
        [Parameter(Mandatory=$true)][URI]$uri,
        [Parameter(Mandatory=$true)][System.IO.FileInfo]$dest
    )

    $start_time = Get-Date
    [Net.ServicePointManager]::SecurityProtocol = $SecurityProtocolType
    $ProgressPreference = 'SilentlyContinue'
    $wr1 = Invoke-WebRequest -Uri $uri -OutFile $dest
    # BUG: why does Invoke-WebRequest sometimes return no object!?
    $sc = "URI "
    if ($wr1) {
        $sc = "URI " + $wr1.StatusCode.ToString() + " "
    }
    Write-Host ($sc + $uri.ToString() + " downloaded to temporary directory " + $dest)

    Write-Verbose "Downloaded time: $((Get-Date).Subtract($start_time).Seconds) second(s)"
}

function Confirm-URI
{
    <#
    .SYNOPSIS
    confirm URI exists by sending HTTP HEAD request
    return $True if HTTP StatusCode == 200 else $False
    #>
    [OutputType([Bool])]
    Param(
        [Parameter(Mandatory=$true)][URI]$uri,
        [Parameter(Mandatory=$false)][Bool]$printResult=$false
    )

    [Net.ServicePointManager]::SecurityProtocol = $SecurityProtocolType
    $ProgressPreference = 'SilentlyContinue'
    $wr1 = Invoke-WebRequest-Head -Uri $uri
    if ($wr1.StatusCode -eq 200) {
        if ($printResult) {
            Write-Host ("URI " + $uri.ToString() + " returned " + $wr1.StatusCode.ToString()) -ForegroundColor Green
        }
        return $True
    }
    if ($printResult) {
        Write-Host ("URI " + $uri.ToString() + " returned " + $wr1.StatusCode.ToString()) -ForegroundColor Red
    }
    return $False
}

function Confirm-URI-Python-Version
{
    <#
    .SYNOPSIS
    confirm the URI for a python zip file exists
    first check known hardcoded URIs then do Confirm-URI $uri
    if -onlyLive then skip known hardcoded URIs (check them online)
    #>
    [OutputType([Bool])]
    Param(
        [Parameter(Mandatory=$true)][URI]$uri,
        [Parameter(Mandatory=$false)][Bool]$onlyLive=$false
    )

    # first check hardcoded known URIs
    if (($URIs_200.Contains($uri.ToString())) -and (-not $onlyLive)) {
        Write-Verbose "URI known to exist ${uri}"
        return $True
    } elseif (($URIs_503.Contains($uri.ToString())) -and (-not $onlyLive)) {
        Write-Verbose "URI known to not exist ${uri}"
        return $False
    }
    # check live URI
    Write-Verbose "URI must be checked ${uri}"
    return Confirm-URI $uri $onlyLive
}

function Create-Python-Zip-Name
{
    <#
    .SYNOPSIS
    return String of python embed.zip file name
    e.g. "python-3.8.2-embed-amd64.zip"
    #>
    [OutputType([String])]
    Param
    (
        [Parameter(Mandatory=$true)][System.Version]$version,
        [Parameter(Mandatory=$true)][Archs]$arch
    )
    return "python-" + $version.ToString() + "-embed-" + $arch.ToString() + ".zip"
}

function Create-Python-Zip-URI
{
    <#
    .SYNOPSIS
    return URI for the Python embed.zip
    example URI of embed .zip
        https://www.python.org/ftp/python/3.8.2/python-3.8.2-embed-amd64.zip
    #>
    [OutputType([URI])]
    Param
    (
        [Parameter(Mandatory=$true)][URI]$base_uri,
        [Parameter(Mandatory=$true)][System.Version]$version,
        [Parameter(Mandatory=$true)][Archs]$arch
    )
    # $version_scraped the version number, e.g. '3.8.10'
    $filename = Create-Python-Zip-Name $version $arch
    # $filename e.g. 'python-3.8.2-embed-amd64.zip'
    # XXX: [URI] does not have an append method? disappoint.
    $uri_file = [URI]($base_uri.ToString() + '/' + $version.ToString() + '/' + $filename)
    # e.g. 'https://www.python.org/ftp/python/3.8.2/python-3.8.2-embed-amd64.zip'
    return $uri_file
}

function Scrape-Python-Versions
{
    <#
    .SYNOPSIS
    Scrape the Python versions HTML page for all archived Python builds among
    versions, do HTTP HEAD to check if the associated embed.zip file exists.

    return Hashtable{[System.Version]$version, [URI]$URI_zip}
    #>
    [OutputType([System.Collections.Hashtable])]
    Param
    (
        [Parameter(Mandatory=$true)][URI]$uri,
        [Parameter(Mandatory=$true)][System.IO.DirectoryInfo]$path_tmp,
        [Parameter(Mandatory=$true)][Archs]$arch_install,
        [Parameter(Mandatory=$false)][Bool]$onlyLive=$false
    )
    Write-Verbose ("Scraping all available versions of Python at " + $uri.ToString())

    $python_versions_html = [System.IO.FileInfo] (Join-Path -Path $path_tmp -ChildPath "python_versions.html")
    if(-not (Test-Path $python_versions_html))
    {
        Download $uri $python_versions_html
    }
    # example snippet of the HTML with listed versions (https://www.python.org/ftp/python/)
    #
    #     <a href="2.3.3/">2.3.3/</a>      20-Mar-2014 21:57 -
    #     <a href="2.3.4/">2.3.4/</a>      20-Mar-2014 21:57 -
    #     <a href="2.3.5/">2.3.5/</a>      20-Mar-2014 21:58 -
    #     <a href="2.3.6/">2.3.6/</a>      01-Nov-2006 07:25 -
    #     <a href="2.3.7/">2.3.7/</a>      20-Mar-2014 21:58 -
    #     ...
    #     <a href="3.9.2/">3.9.2/</a>      19-Feb-2021 14:44 -
    #     <a href="3.9.3/">3.9.3/</a>      02-Apr-2021 13:10 -
    #
    # TODO: consider using built-in HTML parser instead of regexp scraping
    #       https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/invoke-webrequest?view=powershell-7.2#example-3--get-links-from-a-web-page
    #

    # scrape the versions
    $versions_scraped = New-Object Collections.Generic.List[System.Version]
    foreach(
        $m1 in Select-String -CaseSensitive -Pattern '\<a href="[345]\.\d+\.\d+/"\>' -Path $python_versions_html
    ) {
        $m2 = $m1.Matches.value | Select-String -Pattern '"(.*)"'
        if ( $m2.Matches.Groups.Length -lt 2 ) {
            continue
        }
        $subdir = $m2.Matches.Groups[1].value
        # $subdir e.g. '3.8.10/'
        $m3 = $subdir | Select-String -Pattern '[345]\.\d+\.\d+'
        if ($m3.Matches.Length -lt 1) {
            continue
        }
        $version_scraped = $m3.Matches.Value
        # $version_scraped the version number, e.g. '3.8.10'
        $py_version = [System.Version]$version_scraped
        Write-Verbose "scraped version '$version_scraped'"
        $versions_scraped.Add($py_version)
    }
    $versions_scraped.sort()

    # check which scraped versions are valid URIs
    $links = [System.Collections.Hashtable]::new()
    foreach ($py_version in $versions_scraped) {
        $uri_file = Create-Python-Zip-URI $URI_PYTHON_VERSIONS $py_version $arch_install
        if (Confirm-URI-Python-Version $uri_file -onlyLive $onlyLive) {
            Write-Verbose "Add link '$uri_file'"
            $links.Add($py_version, $uri_file)
        }
    }

    return $links
}

function Check-Premade-Uris
{
    <#
    .SYNOPSIS
    Check built-in URIs for expected HTTP Status Codes
    Only meant to aid self-testing this script.
    #>
    Param
    (
        [Parameter(Mandatory=$true)][Archs]$arch_install
    )
    Write-Host "Check expected good URIs for" $arch_install.ToString()
    foreach ($uri in $URIs_200) {
        if (-not ($uri.ToString().Contains($arch_install.ToString())))
        {
            continue
        }
        if ($wr1 = Invoke-WebRequest-Head -Uri $uri) {
            if ($wr1.StatusCode -eq 200) {
                Write-Host ("URI " + $uri.ToString() + " returned " + $wr1.StatusCode.ToString() + " (expected 200)") -ForegroundColor Green
            } else {
                Write-Host ("URI " + $uri.ToString() + " returned " + $wr1.StatusCode.ToString() + " (not expected)") -ForegroundColor Red
            }
        }
    }

    Write-Host "Check expected bad URIs for" $arch_install.ToString() "(scraped URIs that lead to invalid data)"
    foreach ($uri in $URIs_503) {
        if (-not ($uri.ToString().Contains($arch_install.ToString())))
        {
            continue
        }
        if ($wr1 = Invoke-WebRequest-Head -Uri $uri) {
            if ($wr1.StatusCode -ge 400) {
                Write-Host ("URI " + $uri.ToString() + " returned " + $wr1.StatusCode.ToString() + " (expected ≥400)") -ForegroundColor Green
            } else {
                Write-Host ("URI " + $uri.ToString() + " returned " + $wr1.StatusCode.ToString() + " (not expected)") -ForegroundColor Red
            }
        }
    }
}

function Process-Python-Zip
{
    <#
    .SYNOPSIS
    Given the downloaded python zip file $python_zip, install it to
    $path_install.

    BUG: interleaved Write-host and python.exe stdout occurs here
    #>

    Param
    (
        [Parameter(Mandatory=$true)][System.IO.FileInfo]$path_zip,
        [Parameter(Mandatory=$true)][System.IO.FileInfo]$path_install,
        [Parameter(Mandatory=$true)][System.Version]$ver,
        [Parameter(Mandatory=$true)][bool]$skip_exec
    )

    # if $path_install does not exist this will raise
    $path_install = Join-Path $path_install -ChildPath "" -Resolve

    Expand-Archive $path_zip_tmp -DestinationPath $path_install

    #
    # do tedious operations within the $PWD of the recently unzipped Python environment
    #

    Push-Location -Path $path_install

    # 1a. not all versions of embed.zip have this directory
    # e.g. https://www.python.org/ftp/python/3.8.4/python-3.8.4-embed-amd64.zip
    New-Item -type Directory "Lib/site-packages"

    # 1b. the downloaded zip file contains a zip file, python39.zip. Unzip that
    # to under `Lib/`.
    $pythonzip = Get-ChildItem -File -Filter "python*.zip" -Depth 1
    Write-Host -ForegroundColor Yellow "Unzip" $pythonzip
    Expand-Archive $pythonzip -DestinationPath "Lib/python_zip"
    Remove-Item -Path $pythonzip

    # 2. set python._pth file
    $pythonpth = Get-ChildItem -File -Filter "python*._pth" -Depth 1
    if($null -eq $pythonpth) {
        $pythonpth = ".\python._pth"
    }
    $content_pythonpth = "# python._pth
#
# this file was added by PythonEmbed4Win.ps1
.\Lib\python_zip
.\DLLs
.\Lib
.\Scripts
.
.\Lib\site-packages
".Replace("`r`n", "`n")
    # use -Encoding 'ascii'; 'utf8' will prepend UTF8 BOM which is seen by Python
    # as a path (and Python then adds a junk path to `sys.path`)
    $content_pythonpth | Out-File -Force -FilePath $pythonpth -Encoding "ascii"
    Print-File-Nicely $pythonpth

    # 3. set sitecustomize.py
    $python_site_path = Join-Path -Path "." -ChildPath "sitecustomize.py"
    $content_sitecustomize = "# sitecustomize.py
#
# this file was added by PythonEmbed4Win.ps1

import os
import site
import sys

# do not use user-wide site.USER_SITE path; it refers to a path location
# outside of this embed installation
site.ENABLE_USER_SITE = False

# remove site.USER_SITE and the realpath variation from sys.path
# XXX: somewhat time consuming to do on every startup but thorough
__sys_path_index_del = list()
`"`"`"index to delete from sys.path`"`"`"
__user_site_resolve = os.path.realpath(site.USER_SITE)
for __i, __path in enumerate(sys.path):
    __path_resolve = os.path.realpath(__path)
    if site.USER_SITE in (__path, __path_resolve):
        __sys_path_index_del.append(__i)
        continue
    if __user_site_resolve in (__path, __path_resolve):
        __sys_path_index_del.append(__i)
for __index_del in reversed(__sys_path_index_del):
    sys.path.pop(__index_del)
del __sys_path_index_del
del __user_site_resolve
".Replace("`r`n", "`n")
    # use 'ascii' encoding, see above
    $content_sitecustomize | Out-File -FilePath $python_site_path -Encoding "ascii"
    Print-File-Nicely $python_site_path

    # 5. set `pip.ini`
    $python_pip_ini = Join-Path -Path "." -ChildPath "pip.ini"
    $content_pip_ini = "# pip.ini
#
# this file was added by PythonEmbed4Win.ps1

[install]
# this embed installation does not add itself to the shell environment PATH; do not warn about that
no-warn-script-location = true
".Replace("`r`n", "`n")
    # use 'ascii' encoding, see above
    $content_pip_ini | Out-File -FilePath $python_pip_ini -Encoding "ascii"
    Print-File-Nicely $python_pip_ini

    # 6. create empty directory DLLs
    $python_DLL_path = Join-Path -Path "." -ChildPath "DLLs"
    Write-Host -ForegroundColor Yellow "Create Directory" $python_DLL_path
    New-Item -type directory $python_DLL_path

    # 7. basic tests that python can run
    #
    # path to newly installed python.exe
    $python_exe = Join-Path -Path $path_install -ChildPath "python.exe" -Resolve
    # message to print before function returns
    $message1 = "`n`n`nNew self-contained Python executable is at "
    $message2 = "Note that this installation can only run when python.exe is the first command argument.
Do not try run pip.exe directly from the Scripts directory nor any other program there.
Run python.exe and import the required program module, e.g.
$python_exe -m pip

Also, this installation cannot create new virtual environments.
"
    if ($skip_exec) {
        Pop-Location
        Write-Host -ForegroundColor Yellow -NoNewline $message1
        Write-Host -ForegroundColor Yellow -BackgroundColor Blue $python_exe
        Write-Host ""
        Write-Host -ForegroundColor Yellow -NoNewline $message2
        return
    }

    Push-Location ~

    Write-Host -ForegroundColor Yellow "`nTest python can run:"
    Write-Host -ForegroundColor Green "${python_exe} --version`n"
    & $python_exe --version
    if ($LastExitCode -ne 0) {
        Write-Error "Python version test failed"
    }

    $print_sys_path_script = 'import pprint, shutil, sys; w = shutil.get_terminal_size((80, 20)).columns or 80; pprint.pprint(sys.path, width=w)'
    Write-Host -ForegroundColor Yellow "`nPython print sys.path:"
    Write-Host -ForegroundColor Green "${python_exe} -c '$print_sys_path_script'`n"
    & $python_exe -O -c "$print_sys_path_script"
    if ($LastExitCode -ne 0) {
        Write-Error "Python sys.path test failed"
    }

    Pop-Location

    # 8. get pip
    # XXX: oddly, the `ensurepip` module is not available in the Windows embed version of Python
    $path_getpip = ".\get-pip.py"
    Write-Host -ForegroundColor Yellow "`n`nInstall pip:"
    Write-Host -ForegroundColor Green "${python_exe} -O ${path_getpip} --no-warn-script-location`n"
    $uri_getpip1 = $URI_GETPIP
    if ($ver -lt [System.Version]"3.7") {
        $uri_getpip1 = $URI_GETPIP36
    }
    Download $uri_getpip1 $path_getpip
    & $python_exe -O $path_getpip --no-warn-script-location
    if ($LastExitCode -ne 0) {
        Write-Error "Python get-pip.py failed"
    }

    Pop-Location

    Write-Host -ForegroundColor Yellow "`nList installed packages:"
    Write-Host -ForegroundColor Green "${python_exe} -B -m pip list -vv --disable-pip-version-check --no-python-version-warning`n"
    & $python_exe -B -m pip list -vv --disable-pip-version-check --no-python-version-warning
    if ($LastExitCode -ne 0) {
        Write-Error "python -m pip list failed"
    }
    Write-Host -ForegroundColor Yellow -NoNewline $message1
    Write-Host -ForegroundColor Yellow -BackgroundColor Blue $python_exe
    Write-Host ""
    Write-Host -ForegroundColor Yellow -NoNewline $message2
}

function Install-Python
{
    <#
    .SYNOPSIS
    Install python from zipped install `$uri_zip` to `$path_install`.
    Use temporary path `$path_tmp` for intermediate steps.
    #>
    Param
    (
        [Parameter(Mandatory=$true)][System.IO.DirectoryInfo]$path_tmp,
        [Parameter(Mandatory=$true)][System.IO.FileInfo]$path_install,
        [Parameter(Mandatory=$true)][URI]$uri_zip,
        [Parameter(Mandatory=$true)][System.Version]$ver,
        [Parameter(Mandatory=$true)][bool]$skip_exec
    )
    $name_zip = $uri_zip.Segments[-1]
    $path_zip_tmp = [System.IO.FileInfo] (Join-Path -Path $path_tmp -ChildPath $name_zip)

    Download $uri_zip $path_zip_tmp

    if (-not (Test-Path $path_install)) {
        New-Item -Path $path_install -ItemType Directory
    } else {
        Write-Warning "Path already exists, remove it before continuing? ${path_install}"
        Remove-Item -Recurse $path_install -Confirm
        # TODO: Test the user agreed with Remove-item?
        New-Item -Path $path_install -ItemType Directory
    }

    Write-Information "Installing Python to ${path_install}"
    Process-Python-Zip $path_zip_tmp $path_install $ver $skip_exec
}

function Process-Version {
    <#
    .SYNOPSIS
    Resolve major or major.minor to latest major.minor.micro
    #>
    [OutputType([System.Version])]
    Param
    (
        [Parameter(Mandatory=$true)][System.Version]$version,
        [Parameter(Mandatory=$true)][Collections.Generic.List[System.Version]]$versions
    )
    $versions_s = $versions | Sort-Object
    if ($version.Minor -eq -1){
        return (
            $versions_s `
            | Where-Object { $_.Major -eq $version.Major } `
            | Select-Object -Last 1
        )
    }
    elseif ($version.Build -eq -1){
        return (
            $versions_s `
            | Where-Object { $_.Major -eq $version.Major -And $_.Minor -eq $version.Minor } `
            | Select-Object -Last 1
        )
    }
    return $version
}

try {
    Set-StrictMode -Version 3.0
    # save current values
    # TODO: is there a way to push and pop context like this?
    $erroractionpreference_ = $ErrorActionPreference
    $ErrorActionPreference = "Stop"
    $startLocation = Get-Location

    if (-not $Arch) {
        $Arch = $arch_default
    }
    if (-not [Archs].GetEnumNames().Contains($Arch)) {
        Write-Error ("Unknown -arch '${arch}', must be one of " + [Archs].GetEnumNames())
    }
    $archs_ = [Archs]$Arch

    if ($UriCheck) {
        Check-Premade-Uris $archs_
        Write-Host "Check live scraped URIs for" $archs_.ToString()
        Write-Host "(These should match the previous predefined URI settings)"
        $path_tmp1 = New-TemporaryDirectory -extra ("python-latest-" + $archs_.ToString())
        $version_links = Scrape-Python-Versions $URI_PYTHON_VERSIONS $path_tmp1 $archs_ $True
        return
    }

    if ([String]::IsNullOrEmpty($Version)) {
        $path_tmp1 = New-TemporaryDirectory -extra ("python-latest-" + $archs_.ToString())
        Write-Verbose "Temporary Directory ${path_tmp1}"
        $version_links = Scrape-Python-Versions $URI_PYTHON_VERSIONS $path_tmp1 $archs_
        $ver = [System.Version]($version_links.keys | Sort-Object | Select-Object -Last 1)
        Write-Verbose "Version set to latest ${ver}"
        Write-Information ("Determined the latest version of Python to be " + $ver.ToString())
        $url_zip = $version_links[$ver]
        $uri_zip = [URI]$url_zip
    } else {
        $ver = [System.Version] $Version
        if ($ver -le [System.Version]"3.5") {
            Write-Error "Python 3.5 and prior can not run from an embed.zip installation; given $ver"
        }
        $path_tmp1 = New-TemporaryDirectory -extra ("python-" + $Version + "-" + $archs_.ToString())
        Write-Verbose "Temporary Directory ${path_tmp1}"
        if ($ver.Minor -eq -1 -or $ver.Build -eq -1) {
            $version_links = Scrape-Python-Versions $URI_PYTHON_VERSIONS $path_tmp1 $archs_
            $versions = @($version_links.keys)
            $ver = Process-Version ([System.Version] $ver) $versions
            Write-Verbose "Version set to ${ver}"
            if ((-not $ver) -or ($ver.Major -eq -1)) {
                Write-Error "Unable to process given version ${Version}"
            }
        } else {
            $ver = [System.Version]$Version
            Write-Verbose "Version is ${ver}"
        }
        $uri_zip = Create-Python-Zip-URI $URI_PYTHON_VERSIONS $ver $archs_
        if(-not (Confirm-URI $uri_zip))
        {
            Write-Error "Python version ${Version} for arch ${archs_} was not found at ${URI_PYTHON_VERSIONS}"
        }
    }

    if (-not $Path) {
        # user did not pass -Path so create a sensible one
        $pyDist = "python-" + $ver.ToString() + "-embed-" + $archs_.ToString()
        $Path = [System.IO.FileInfo] (Join-Path -Path "." -ChildPath $pyDist)
    }
    Install-Python $path_tmp1 $Path $uri_zip $ver $SkipExec
    Write-Host -ForegroundColor Yellow "`nInstalled from" $uri_zip

    Write-Host "`nCompleted in" $stopWatch.Elapsed.ToString()
} catch {
    $ErrorActionPreference = "Continue"
    Write-Error $_.ScriptStackTrace
    Write-Error -Message $_.Exception.Message
} finally {
    Set-StrictMode -Off
    if ($null -ne $path_tmp1) {
        Remove-Item -Recurse $path_tmp1 -ErrorAction Continue
    }
    $ErrorActionPreference = $erroractionpreference_
    Set-Location -Path $startLocation
}
