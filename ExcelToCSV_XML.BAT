
REM @echo off

REM set curdir=%~dp0
REM set exepath="%curdir%EXE\ReadExcel.exe"
REM set savepath="%curdir%CSV"
 
REM EXE路径
set exepath="G:\MyProject\GameDemoDocument\ExcelToCSV_XML\ExcelToCSV_XML\ExcelToCSV_XML\bin\Debug\ExcelToCSV_XML.exe"


REM 表路径
set filepath="G:\MyProject\GameDemoDocument\Table"


REM 导出路径
set savepath="G:\MyProject\GameDemoDocument\Export"

cd %filepath%

for /f "delims=" %%i in ('dir /s/a-d /b *.xlsx') do %exepath% %savepath% %%i



pause