title LineChanger Test
color 17

@echo off
@cd/d "%~dp0"

set file=%cd%\hello_text.txt
set line=Array(4,5)
set txt=Bye World

@cd ..


:BaseTest0
    for /l %%i in (1, 1, 8) do echo %%i >>%file%
    call wscript "%cd%\LineChanger.vbs" %file% %line% %txt%
    
    type %file%
    del  %file%
    echo ------------------ Great ! ------------------
    
    pause
    
    