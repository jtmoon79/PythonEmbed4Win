
# PythonEmbed4Win

A single PowerShell script to easily and quickly create a standalone Python
embed.zip local environment.

This is like a Python Virtual Environment but technically is not. It does not
set environment variable `VIRTUAL_ENV` nor modify the `PATH`.

Installing the Windows Embed zip file requires some tedious tweaks.
See this [gist](https://gist.github.com/jtmoon79/ce63fe655b2f544462e70d8e5ec30ff5).

This script will handle the tedious tweaks so the new Python installation will
run in an isolated manner.
