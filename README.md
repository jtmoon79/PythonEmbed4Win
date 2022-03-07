# PythonEmbed4Win

A [single PowerShell script](PythonEmbed4Win.ps1) to easily and quickly
create a standalone Python local environment from a Windows embed.zip. No prior
Python installation is required.

Installing the Python for Windows Embed zip file requires some tedious tweaks.
See this [gist](https://gist.github.com/jtmoon79/ce63fe655b2f544462e70d8e5ec30ff5).
This script will handle the tedious tweaks so the new Python installation will
run in an isolated manner.

This is similar to a Python Virtual Environment but technically is not.
It does not add an _activate_ script to set environment variable `VIRTUAL_ENV`
or modify the `PATH`.
