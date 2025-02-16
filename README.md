# PythonEmbed4Win

![PowerShell 5](https://img.shields.io/badge/5-blue?logo=Powershell&logoColor=blue&label=PowerShell&labelColor=white&color=blue) ![PowerShell 7](https://img.shields.io/badge/7-blue?logo=Powershell&logoColor=purple&label=PowerShell&labelColor=white&color=purple)
![Python Versions](https://img.shields.io/badge/3.6%20%7C%203.7%20%7C%203.8%20%7C%203.9%20%7C%203.10%20%7C%203.11%20%7C%203.12%20%7C%203.13-blue?logo=Python&logoColor=yellow&label=Python&labelColor=blue&color=white)

A [single PowerShell script](PythonEmbed4Win.ps1) to easily and quickly
create a standalone Python local environment for Windows by downloading the requested `embed.zip`
distributed file. No prior Python installation is required.

To run:

```powershell
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/jtmoon79/PythonEmbed4Win/main/PythonEmbed4Win.ps1" -OutFile "PythonEmbed4Win.ps1"
.\PythonEmbed4Win.ps1
```

For detailed help:

```powershell
Get-Help .\PythonEmbed4Win.ps1 -full
```

If you get the error<br/> `PythonEmbed4Win.ps1 cannot be loaded because running scripts is disabled on this system`<br/>
then run:

```powershell
Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope Process
```

<br/>

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

Derived from [this StackOverflow answer](https://stackoverflow.com/a/68958636/471376).
