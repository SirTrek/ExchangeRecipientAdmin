@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Start-Process powershell.exe -Verb RunAs -ArgumentList '-ExecutionPolicy Bypass -File \"%~dp0Start-ExchangeRecipientAdminCenter.ps1\"'"
