# PythonEmbed4Win

A [single PowerShell script](PythonEmbed4Win.ps1) to easily and quickly
create a standalone Python local environment from a Windows embed.zip. No prior
Python installation is required.

To run:

```powershell
PS> Invoke-WebRequest -Uri "https://raw.githubusercontent.com/jtmoon79/PythonEmbed4Win/main/PythonEmbed4Win.ps1" -OutFile "PythonEmbed4Win.ps1"
PS> .\PythonEmbed4Win.ps1 -?
```

<br/>

Installing the Python for Windows Embed zip file requires some tedious tweaks.
See this [gist](https://gist.github.com/jtmoon79/ce63fe655b2f544462e70d8e5ec30ff5).
This script will handle the tedious tweaks so the new Python installation will
run in an isolated manner.

This is similar to a Python Virtual Environment but technically is not.
It does not require an _activate_ script to set environment variable `VIRTUAL_ENV`
or modify the `PATH`. It will run isolated without environment modifications.

Derived from [this StackOverflow answer](https://stackoverflow.com/a/68958636/471376).
