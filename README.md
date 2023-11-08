# PythonEmbed4Win

![PowerShell 5](https://img.shields.io/badge/5-blue?logo=Powershell&logoColor=blue&label=PowerShell&labelColor=white&color=blue) ![PowerShell 7](https://img.shields.io/badge/7-blue?logo=Powershell&logoColor=purple&label=PowerShell&labelColor=white&color=purple)
![Python Versions](https://img.shields.io/badge/3.5%20%7C%203.6%20%7C%203.7%20%7C%203.8%20%7C%203.9%20%7C%203.10%20%7C%203.11%20%7C%203.12-blue?logo=Python&logoColor=yellow&label=Python&labelColor=blue&color=white)


A [single PowerShell script](PythonEmbed4Win.ps1) to easily and quickly
create a standalone Python local environment from a Windows `embed.zip`
distributed file. No prior Python installation is required.

To run:

```powershell
PS> Invoke-WebRequest -Uri "https://raw.githubusercontent.com/jtmoon79/PythonEmbed4Win/main/PythonEmbed4Win.ps1" -OutFile "PythonEmbed4Win.ps1"
PS> .\PythonEmbed4Win.ps1 -?
```

<br/>

Installing the Python for Windows embedded zip file requires some tedious tweaks.
See this [gist](https://gist.github.com/jtmoon79/ce63fe655b2f544462e70d8e5ec30ff5).
This script will handle the tedious tweaks so the new Python installation will
run in an isolated manner.

This is similar to a Python Virtual Environment but technically is not.
It does not require an _activate_ script to set environment variable `VIRTUAL_ENV`
or modify the `PATH`. It will run isolated without environment modifications.

Derived from [this StackOverflow answer](https://stackoverflow.com/a/68958636/471376).
