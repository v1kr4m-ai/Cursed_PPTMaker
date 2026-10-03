@echo off
rem Double-click to pick files, or drag documents (or a spec.json) onto this file.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0make_ppt_launcher.ps1" %*
