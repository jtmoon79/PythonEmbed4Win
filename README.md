# PythonEmbed4Win

![PowerShell 5](https://img.shields.io/badge/5-blue?logo=Powershell&logoColor=blue&label=PowerShell&labelColor=white&color=blue) ![PowerShell 7](https://img.shields.io/badge/7-blue?logo=Powershell&logoColor=purple&label=PowerShell&labelColor=white&color=purple)
![Python Versions](https://img.shields.io/badge/3.6%20%7C%203.7%20%7C%203.8%20%7C%203.9%20%7C%203.10%20%7C%203.11%20%7C%203.12%20%7C%203.13-blue?logo=Python&logoColor=yellow&label=Python&labelColor=blue&color=white)

A [single PowerShell script](PythonEmbed4Win.ps1) to easily and quickly
create a standalone Python local environment for Windows by downloading the requested `embed.zip`
distributed file. No prior Python installation is required.

## Download and run

```powershell
Invoke-WebRequest `
  -Uri "https://raw.githubusercontent.com/jtmoon79/PythonEmbed4Win/main/PythonEmbed4Win.ps1" `
  -OutFile "PythonEmbed4Win.ps1"

.\PythonEmbed4Win.ps1
```

If you get the error:

`PythonEmbed4Win.ps1 cannot be loaded because running scripts is disabled on this system`<br/>

then run:

```powershell
Set-ExecutionPolicy `
  -ExecutionPolicy Unrestricted `
  -Scope Process
```

## About

Installing the Python for Windows embedded zip file requires some tedious tweaks.
See this [gist](https://gist.github.com/jtmoon79/ce63fe655b2f544462e70d8e5ec30ff5).
This script will handle the tedious tweaks and updates so the new Python
installation will run correctly in an isolated manner.

This is similar to a Python Virtual Environment but technically is not.
It does not require an _activate_ script to set environment variable `VIRTUAL_ENV`
or modify the `PATH`. It will run isolated without environment modifications.

One disadvantage is that a Windows embed Python cannot create a functioning
virtual environment. They will be created but `virtualenv` and `venv`
selectively copy files from the source and do not copy necessary library files
unique to an Windows embed Python.
Additionally some pypi libraries with complex external C module dependencies may crash during initialization. 
The good news is you'll know immediately if the Python embed installation created by this script will work for you.

Derived from [this StackOverflow answer](https://stackoverflow.com/a/68958636/471376).

## help

```plain-text
PS> Get-Help ./PythonEmbed4Win.ps1 -full

NAME
    ./PythonEmbed4Win.ps1
    
SYNOPSIS
    Quickly setup a portable python environment for Windows using an embed.zip.
    
    
SYNTAX
    ./PythonEmbed4Win.ps1 [-Path <String>] [-Version <String>] [-Arch <String>] [-SkipExec] [-trace] [<CommonParameters>]

    ./PythonEmbed4Win.ps1 -ZipFile <String> [-Path <String>] [-SkipExec] [-trace] [<CommonParameters>]
    
    ./PythonEmbed4Win.ps1 -UriCheck [[-Path] <String>] [-Version <String>] [-Arch <String>] [-trace] [<CommonParameters>]
    
    
DESCRIPTION
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
    
    -ZipFile allows installing a local .zip file.
     The .zip file must have an embedded version string within the basename.
    
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
    

PARAMETERS
    -ZipFile <String>
        Install this local .zip file. Does not download. The .zip file must have a version string
        embedded in the basename.
        
        Required?                    true
        Position?                    named
        Default value                
        Accept pipeline input?       false
        Accept wildcard characters?  false
        
    -UriCheck [<SwitchParameter>]
        Only check pre-filled URIs (script self-test). Does not install Python.
        
        Required?                    true
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters?  false
        
    -Path <String>
        Install to this path. Defaults to a descriptive name.
        Default pipeline value and default argument value.
        
        Required?                    false
        Position?                    1
        Default value                
        Accept pipeline input?       true (ByValue)
        Accept wildcard characters?  false
        
    -Version <String>
        Version of Python to install. Leave blank to fetch the latest Python.
        Can pass major.minor.micro or just major.minor, e.g. "3.8.2" or "3.8".
        If passed only major.minor then the latest major.minor.micro will be chosen.
        Python 3.6 and later is supported.
        
        Required?                    false
        Position?                    named
        Default value                
        Accept pipeline input?       false
        Accept wildcard characters?  false
        
    -Arch <String>
        Architecture: win32 or amd64 or arm64. Defaults to the current architecture.
        
        Required?                    false
        Position?                    named
        Default value                
        Accept pipeline input?       false
        Accept wildcard characters?  false
        
    -SkipExec [<SwitchParameter>]
        Do not execute python.exe after installation.
        This skips the python.exe self-test and the run of `get-pip.py`.
        
        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters?  false
        
    -trace [<SwitchParameter>]
        Turn on debug tracing.
        
        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters?  false
        
    <CommonParameters>
        This cmdlet supports the common parameters: Verbose, Debug,
        ErrorAction, ErrorVariable, WarningAction, WarningVariable,
        OutBuffer, PipelineVariable, and OutVariable. For more information, see
        about_CommonParameters (https://go.microsoft.com/fwlink/?LinkID=113216). 
    
INPUTS
    
OUTPUTS
    
NOTES
    
    
        Author: James Thomas Moon
    
    
RELATED LINKS
    https://github.com/jtmoon79/PythonEmbed4Win

```