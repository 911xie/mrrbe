@echo off

rem tokens=3 代表第三列，所以必须配置文件=两边要有空格
rem skip=1 代表越过第一行



set path=""
for /f "tokens=3" %%a in (D:\eclipse-workspace\mrrbe\script-path.txt) do (
    set path=%%a
    goto :Show
)
:Show
ECHO %path%
C:\Windows\explorer.exe %path%