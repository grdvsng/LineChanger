@echo off
title LineChanger Help
mode con:cols=60 lines=30
Setlocal EnableDelayedExpansion

set colX=0
set colY=9
set load=------
 
call :def_color_animation
call :def_run

:def_run
    echo ----------------------- LineChanger ------------------------
    echo author grdvsng@gmail.com
    echo ------------------------------------------------------------
    echo.
    
    echo 1 How to configure the script to replace 1 line.
    echo 2 How to configure the script to replace many lines.
    
    CHOICE /C 12 /M "Press key> "
    if %errorlevel%==1 cls & call :def_line_help1
    if %errorlevel%==2 cls & call :def_line_help2
    
    exit /b
    
    
:def_color_animation

    color %colX%%colY%
    set /a colX+=1
    set /a colY-=1
    set load=%load%------
    
    timeout /t 1 /nobreak >nul
    cls & echo %load%
    
    if %colX% NEQ 9 call :def_color_animation
    cls & color 37 & exit /b
    
    

:def_line_help1
    set line3=write like that: ChangeLine "your_file_fullpath", "line", "new text"
    set line4=------------------------------------------------------------
    set line5=exsample: 
    set line6=    ChangeLine "C:\Temp\test.txt", 2, "Hello World!"
    
    call :def_engine 6 4
    exit /b
    
    
:def_line_help2
    set line3=write like that: ChangeLine "your_file_fullpath", Array("lines1", ...), Array("new text1", ...)
    set line4=------------------------------------------------------------
    set line5=exsample: 
    set line6=    ChangeLine "C:\Temp\test.txt", Array(2, 3), Array("HW!", "WH!")
    set line7=%line4%
    set line8= Where Array(lines) == Array(texts)
    
     
    call :def_engine 8 4
    exit /b


:def_engine
    set line1=open file LineChanger.vbs
    set line2=down on bottom line:
    
    echo -----------------------    HELP    -------------------------
    
    for /l %%i in (1, 1, %1) do (
        if %%i LSS %2 echo Step:%%i 
        echo !line%%i!
        echo.
        timeout /t 2 /nobreak >nul
    )
    
    pause
    cls & exit /b