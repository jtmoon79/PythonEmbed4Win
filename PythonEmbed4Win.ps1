#!powershell
#
# PythonEmbed4Win.ps1
#
# BUG: how to flush all streams? currently getting interleaved text in console
#      hack workarounds are Write-Host `n but it's not perfect
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

    The installation uses the Windows "Embedded" distribution zip file.
    That zip file distribution requires tedious and non-obvious steps.
    This script adjusts installation to be runnable and isolated (removes
    references to python paths outside it's own directory). Also installs latest
    pip. See https://gist.github.com/jtmoon79/ce63fe655b2f544462e70d8e5ec30ff5

    Only Python 3.6 and later releases will function correctly.

    This installed Python distribution a Python Virtual Environment but
    technically is not. It does not set environment variable VIRTUAL_ENV nor
    modify the PATH.

    Inspired by this stackoverflow.com question:
    https://stackoverflow.com/questions/68958635/python-windows-embeddable-package-fails-to-run-no-module-named-pip-the-system/68958636

    BUG: some embed.zip are hardcoded to look for other Windows Python
         installations, like Python 3.8 series.
         For example, install Python-3.8.6.msi. Then run this script to
         install python-3.8.4-embed-amd64.zip.
         The Python 3.8.4 sys.path will be the confusing:
             C:\python-embed-3.8.4\python38.zip
             C:\python-msi-install-3.8.6\Lib
             C:\python-msi-install-3.8.6\DLLs
             C:\python-embed-3.8.4
             C:\python-msi-install-3.8.6
             C:\python-msi-install-3.8.6\lib\site-packages
          The python._pth and sitecustomize.py seem to have no affect.
.PARAMETER Version
    Version of Python to install. Leave blank to fetch the latest Python.
    Can pass major.minor.micro or just major.minor, e.g. "3.8.2" or "3.8".
    If passed only major.minor then the latest major.minor.micro will be chosen.
.PARAMETER Path
    Install to this path. Defaults to a descriptive name.
.PARAMETER Arch
    Architecture: win32 or amd64. Defatults to the current architecture.
.PARAMETER Help
    Print the help message and return.
.LINK
    https://github.com/jtmoon79/PythonEmbed4Win
.NOTES
    Author: James Thomas Moon
    Date: 2022
#>
[Cmdletbinding()]
Param (
    [string] $Path,
    [string] $Version,
    [string] $Arch,
    [switch] $help
)

New-Variable -Name SCRIPT_NAME -Value "PythonEmbed4Win.ps1" -Option ReadOnly -Force
$SecurityProtocolType = [Net.SecurityProtocolType]::Tls12

if ($help) {
    Get-Help "${PSScriptRoot}\${SCRIPT_NAME}"
    Return
}

# save current values
# TODO: is there a way to push and pop context like this?
$erroractionpreference_ = $ErrorActionPreference
$ErrorActionPreference = "Stop"
$startLocation = Get-Location

Enum Archs {
    # enum names are the literal substrings within the named .zip file
    win32
    amd64
}
$arch_default = ${env:PROCESSOR_ARCHITECTURE}.ToLower()

New-Variable -Name URI_GETPIP -Option ReadOnly -Force -Value ([URI] "https://bootstrap.pypa.io/get-pip.py")
New-Variable -Name URI_PYTHON_VERSIONS -Option ReadOnly -Force -Value ([URI] "https://www.python.org/ftp/python")

function URI-Combine
{
    Param(
        [Parameter(Mandatory=$true)][URI]$uri,
        [Parameter(Mandatory=$true)][string]$append
    )
    return [URI]($uri.ToString() + $append.ToString())
}

# Pre-fill known URIs that exist as of Feb. 2022, i.e. return HTTP 200
$URIs_200 = @(
    #
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
    #
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
    URI-Combine $URI_PYTHON_VERSIONS '/3.9.10/python-3.9.10-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.10.0/python-3.10.0-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.10.1/python-3.10.1-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.10.2/python-3.10.2-embed-amd64.zip'
)
# Pre-fill known URIs that do not exist as of Feb. 2022, i.e. return HTTP 503
$URIs_503 = @(
    #
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
    URI-Combine $URI_PYTHON_VERSIONS '/3.11.0/python-3.11.0-embed-win32.zip'
    #
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
    URI-Combine $URI_PYTHON_VERSIONS '/3.8.11/python-3.8.11-embed-amd64.zip'
    URI-Combine $URI_PYTHON_VERSIONS '/3.8.12/python-3.8.12-embed-amd64.zip'
    #URI-Combine $URI_PYTHON_VERSIONS '/3.11.0/python-3.11.0-embed-amd64.zip'
)

function New-TemporaryDirectory {
    Param([string]$extra)
    Join-Path -Path $([System.IO.Path]::GetTempPath()) -ChildPath $extra.ToString() `
        | ForEach-Object { New-Item -Path $_ -ItemType Directory -Force }
}

function Download
{
    # download file at $uri to $dest, note the time taken
    Param(
        [Parameter(Mandatory=$true)][URI]$uri,
        [Parameter(Mandatory=$true)][string]$dest
    )

    $start_time = Get-Date
    [Net.ServicePointManager]::SecurityProtocol = $SecurityProtocolType
    $wr1 = Invoke-WebRequest -Uri $uri.ToString() -OutFile $dest
    # BUG: why is Invoke-WebRequest sometimes returning no object!?
    $sc = "URI "
    if ($wr1) {
        $sc = "URI " + $wr1.StatusCode.ToString() + " "
    }
    Write-Host ($sc + $uri.ToString() + " downloaded to " + $dest)

    Write-Verbose "Downloaded time: $((Get-Date).Subtract($start_time).Seconds) second(s)"
}

function Confirm-URI
{
    # confirm URI exists by sending HTTP HEAD request
    # return $True if HTTP StatusCode == 200 else $False
    Param(
        [Parameter(Mandatory=$true)][URI]$uri
    )

    [Net.ServicePointManager]::SecurityProtocol = $SecurityProtocolType
    $wr1 = Invoke-WebRequest -Uri $uri.ToString() -Method Head -SkipHTTPErrorCheck
    Write-Host ("URI " + $uri.ToString() + " returned " + $wr1.StatusCode.ToString())
    if ($wr1.StatusCode -eq 200) {
        return $True
    }
    return $False
}

function Confirm-URI-Python-Version
{
    # confirm the URI for a python zip file exists
    # first check known hardcoded URIs then do Confirm-URI $uri
    # return boolean
    Param(
        [Parameter(Mandatory=$true)][URI]$uri
    )

    # first check hardcoded known URIs
    if ($URIs_200.Contains($uri.ToString())) {
        Write-Verbose "URI known to exist ${uri}"
        return $True
    } elseif ($URIs_503.Contains($uri.ToString())) {
        Write-Verbose "URI known to not exist ${uri}"
        return $False
    }
    # check live URI
    Write-Verbose "URI must be checked ${uri}"
    return Confirm-URI $uri
}

function Create-Python-Zip-Name
{
    # return String of python embed.zip file name
    # e.g. "python-3.8.2-embed-amd64.zip"
    Param
    (
        [Parameter(Mandatory=$true)][System.Version]$version,
        [Parameter(Mandatory=$true)][Archs]$arch
    )
    return "python-" + $version.ToString() + "-embed-" + $arch.ToString() + ".zip"
}

function Create-Python-Zip-URI
{
    # return URI for the Python embed.zip
    #
    # example URI of embed .zip
    #     https://www.python.org/ftp/python/3.8.2/python-3.8.2-embed-amd64.zip
    #
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
    # scrape the HTML page for all archived Python builds
    # among versions, do HTTP HEAD to check if the associated embed.zip file
    # exists.
    #
    # return Hashtable{[System.Version]$version, [URI]$URI_zip}
    Param
    (
        [Parameter(Mandatory=$true)][URI]$uri,
        [Parameter(Mandatory=$true)][string]$path_tmp,
        [Parameter(Mandatory=$true)][Archs]$arch_install
    )
    Write-Verbose ("Scraping all available versions of Python at " + $uri.ToString())

    $python_versions_html = Join-Path -Path $path_tmp -ChildPath "python_versions.html"
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
    $links = [Collections.Hashtable]::new()
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
        $sys_version = [System.Version]$version_scraped
        $uri_file = Create-Python-Zip-URI $URI_PYTHON_VERSIONS $sys_version $arch_install
        if (Confirm-URI-Python-Version $uri_file) {
            $links.Add($sys_version, $uri_file)
        }
    }
    return $links
}

function Process-Python-Zip
{
    # given the downloaded python zip file $python_zip, install it to $path_install
    # with needed tweaks
    #
    # BUG: interleaved Write-host and $python_exe stdout occurs
    Param
    (
        [Parameter(Mandatory=$true)][string]$path_zip,
        [Parameter(Mandatory=$true)][string]$path_install
    )

    # if $path_install does not exist this will raise
    $path_install = Join-Path $path_install -ChildPath "" -Resolve

    Expand-Archive $path_zip_tmp -DestinationPath $path_install

    #
    # do tedious operations within the $PWD of the recently unzipped Python environment
    #

    Push-Location -Path $path_install

    # 1. the downloaded zip file contains a zip file, python39.zip. Unzip that
    $pythonzip = Get-ChildItem -File -Filter "python*.zip" -Depth 1
    Write-Host -ForegroundColor Yellow "Unzip" $pythonzip
    Expand-Archive $pythonzip -DestinationPath "python_zip"
    Remove-Item -Path $pythonzip

    # not all versions of embed.zip have this directory
    # e.g. https://www.python.org/ftp/python/3.8.4/python-3.8.4-embed-amd64.zip
    New-Item -type directory "Lib/site-packages"

    # 2. set python._pth file
    $pythonpth = Get-ChildItem -File -Filter "python*._pth" -Depth 1
    if($pythonpth -eq $null) {
        $pythonpth = ".\python._pth"
    }
    $content_pythonpth = "# python._pth
.\Scripts
.
# standard libraries
.\python_zip
# importing site will run sitecustomize.py
import site
".ReplaceLineEndings()
    Write-Host -ForegroundColor Yellow `
        "Set" $pythonpth "Contents:`n`n`t" $content_pythonpth.ReplaceLineEndings("`n`t") "`n"
    $content_pythonpth | Out-File -Force -FilePath $pythonpth -Encoding "utf8"

    # 3. set sitecustomize.py
    $python_site_path = Join-Path -Path "." -ChildPath "sitecustomize.py"
    $content_sitecustomize = "# sitecustomize.py
import sys
import site
# do not use user-wide site.USER_SITE path, it refers to location outside this embed installation
# this hack was added by ${SCRIPT_NAME}
site.ENABLE_USER_SITE = False
if site.USER_SITE in sys.path:
    sys.path.remove(site.USER_SITE)".ReplaceLineEndings()
    Write-Host -ForegroundColor Yellow `
        "Set" $python_site_path "Contents:`n`n`t" $content_sitecustomize.ReplaceLineEndings("`n`t") "`n"
    $content_sitecustomize | Out-File -FilePath $python_site_path -Encoding "utf8"

    # 4. create empty directory DLLs
    $python_DLL_path = Join-Path -Path "." -ChildPath "DLLs"
    Write-Host -ForegroundColor Yellow "Create Directory" $python_DLL_path
    New-Item -type directory $python_DLL_path

    # 5. basic tests that python can run
    #
    # path to newly installed python.exe
    $python_exe = Join-Path -Path $path_install -ChildPath "python.exe" -Resolve
    Push-Location ~

    Write-Host -ForegroundColor Yellow "`nTest python can run:`n${python_exe} --version`n"
    & $python_exe --version
    if ($LastExitCode -ne 0) {
        Write-Error "Python version test failed"
    }

    Write-Host -ForegroundColor Yellow "`nPython print sys.path:"
    & $python_exe -O -c "import sys, pprint
pprint.pprint(sys.path)
print()
"
    if ($LastExitCode -ne 0) {
        Write-Error "Python sys.path test failed"
    }

    Pop-Location

    # 6. get pip
    Write-Host -ForegroundColor Yellow "`n`nInstall pip:`n${python_exe} -O ${path_getpip} --no-warn-script-location`n"
    $path_getpip = ".\get-pip.py"
    Download $URI_GETPIP $path_getpip
    & $python_exe -O $path_getpip --no-warn-script-location
    if ($LastExitCode -ne 0) {
        Write-Error "Python get-pip.py failed"
    }

    Pop-Location

    Write-Host -ForegroundColor Yellow "`nList installed packages:`n${python_exe} -B -m pip list -vv --disable-pip-version-check --no-python-version-warning`n"
    & $python_exe -B -m pip list -vv --disable-pip-version-check --no-python-version-warning
    if ($LastExitCode -ne 0) {
        Write-Error "python -m pip list failed"
    }
    Write-Host -ForegroundColor Yellow "`n`n`nNew self-contained Python executable is at" $python_exe
}

function Install-Python
{
    # install python from an embed.zip URI
    Param
    (
        [Parameter(Mandatory=$true)][string]$path_tmp,
        [Parameter(Mandatory=$true)][string]$path_install,
        [Parameter(Mandatory=$true)][URI]$uri_zip
    )
    $name_zip = $uri_zip.Segments[-1]
    $path_zip_tmp = Join-Path -Path $path_tmp -ChildPath $name_zip

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
    Process-Python-Zip $path_zip_tmp $path_install
}

function Process-Version {
    # resolve major or major.minor to latest major.minor.micro
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
    if (-not $Arch) {
        $Arch = $arch_default
    }
    if (-not [Archs].GetEnumNames().Contains($Arch)) {
        Write-Error ("Unknown -arch '${arch}', must be one of " + [Archs].GetEnumNames())
    }
    $archs_ = [Archs]$Arch

    if ([string]::IsNullOrEmpty($Version)) {
        $path_tmp1 = New-TemporaryDirectory -extra ("python-latest-" + $archs_.ToString())
        Write-Verbose "Temporary Directory ${path_tmp1}"
        $version_links = Scrape-Python-Versions $URI_PYTHON_VERSIONS $path_tmp1 $archs_
        $ver = $version_links.keys | Sort-Object | Select-Object -Last 1
        Write-Verbose "Version set to latest ${ver}"
        Write-Information ("Determined the latest version of Python to be " + $ver.ToString())
        $url_zip = $version_links[$ver]
        $uri_zip = [URI]$url_zip
    } else {
        $path_tmp1 = New-TemporaryDirectory -extra ("python-" + $Version + "-" + $archs_.ToString())
        Write-Verbose "Temporary Directory ${path_tmp1}"
        $v1 = [System.Version] $Version
        if ($v1.Minor -eq -1 -or $v1.Build -eq -1) {
            $version_links = Scrape-Python-Versions $URI_PYTHON_VERSIONS $path_tmp1 $archs_
            $versions = @($version_links.keys)
            $ver = Process-Version ([System.Version] $v1) $versions
            Write-Verbose "Version set to ${ver}"
            if ((-not $ver) -or ($ver.Major -eq -1)) {
                Write-Error "Unable to process given version ${Version}"
            }
        } else {
            $ver = $Version
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
        $Path = Join-Path -Path "." -ChildPath $pyDist
    }
    Install-Python $path_tmp1 $Path $uri_zip
    Write-Host -ForegroundColor Yellow "Installed from" $uri_zip
} catch {
    $ErrorActionPreference = "Continue"
    Write-Error $_.ScriptStackTrace
    Write-Error -Message $_.Exception.Message
} finally {
    Remove-Item -Recurse $path_tmp1
    $ErrorActionPreference = $erroractionpreference_
    Set-Location -Path $startLocation
}
