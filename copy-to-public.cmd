@echo off
chcp 65001>nul
For /f "tokens=1-4 delims=. " %%a in ('date /t') do (
set mydate=%%c-%%b-%%a
)

echo [92mКопирование файлов групп для публичного доступа [0m
echo [93mВведите номер группы, для которой нужно скопировать файл (1-12, 0 - если нужно скопировать все): [0m
set /p group=

IF "%group%"=="0" (

		FOR /L %%G IN (1,1,12) DO (
			echo [92m%%G группа.xlsx ---- \Public\%%G\%%G группа %mydate%.xlsx [0m
			copy /Y "%%G группа.xlsx" "..\Public\%%G\%%G группа %mydate%.xlsx" >nul
		)
) ELSE (

echo [92m%group% группа.xlsx ---- \Public\%group%\%group% группа %mydate%.xlsx [0m
copy /Y "%group% группа.xlsx" "..\Public\%group%\%group% группа %mydate%.xlsx" >nul

)
pause >nul