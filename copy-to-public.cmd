@echo off
chcp 65001>nul
For /f "tokens=1-4 delims=. " %%a in ('date /t') do (
set mydate=%%c-%%b-%%a
)

FOR /L %%G IN (1,1,12) DO (
	echo [92m %%G группа.xlsx ---- \Public\%%G\%%G группа %mydate%.xlsx [0m
	echo
	copy /Y "%%G группа.xlsx" "..\Public\%%G\%%G группа %mydate%.xlsx" >nul
)                                                                    
