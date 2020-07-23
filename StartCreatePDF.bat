echo off

REM PowerShellよりスクリプトを起動する。

powershell -executionpolicy remotesigned .\ChangePdf.ps1

REM 実行結果確認のため、一時停止
PAUSE
