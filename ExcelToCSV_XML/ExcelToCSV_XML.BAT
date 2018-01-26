
REM @echo off

REM set curdir=%~dp0
REM set exepath="%curdir%EXE\ReadExcel.exe"
REM set savepath="%curdir%CSV"
 
REM EXE路径
set exepath="G:\MyProject\GameDemoTools\ExcelToCSV_XML\ExcelToCSV_XML\ExcelToCSV_XML\bin\Debug\ExcelToCSV_XML.exe"


REM 表路径
set filepath="G:\MyProject\GameDemoTools\ExcelToCSV_XML\Table"


REM 导出路径
set savepath="G:\MyProject\GameDemoTools\ExcelToCSV_XML\Export"

cd %filepath%

for /f "delims=" %%i in ('dir /s/a-d /b *.xlsx') do %exepath% %savepath% %%i



pause