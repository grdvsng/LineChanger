dim fso, newfile, env, err

set fso = CreateObject("Scripting.FileSystemObject")
set osh = CreateObject("WScript.Shell")
set env = osh.Environment("Process")
err     = array("Errors occurred while copying.", _
                "Text array != line array.")


function is_same(base, current)
    ' function is_same
    ' parameters  base, current
    ' base - base file
    ' current - file than was created via script
    ' if file not same return false
    ' else return true 
    
    dim result, f1, f2
    set file1 = fso.OpenTextFile(base)
    set file2 = fso.OpenTextFile(current)
    
    result = 0
    
    do until file1.AtEndOfStream
        if file1.ReadLine <> file2.ReadLine then result = result + 1
    loop
    
    file1.close
    file2.close
    is_same = result <= 1
    
end function


function exists(file)
    ' function exists
    ' parameters file
    ' check exists if not return MsgBox
    ' else return true
    
    if not fso.FileExists(file) then 
        MsgBox "File: '" & file & "' not found!", 0, "Error!"
        exists = false
    else
        exists = true
    end if
    
end function


function Array_loop(file, line, text)
    dim msg
    
    for i=0 to uBound(line)
        if IsArray(text) then 
            if uBound(text) <> uBound(line) then 
                MsgBox err(1), 0, "Error"
                WScript.quit()
            end if
            
            msg = text(i)

        else 
            msg = text
        end if
        ChangeLine file, line(i), msg
    next
    WScript.quit()
    
end function


function ChangeLine(file, line, text)
    ' function ChangeLine
    ' parameters file, line, text
    ' file - file obj.
    ' line int or array for change.
    ' text - new str in file line.
    
    line = eval(line)
    if IsArray(line) then call Array_loop(file, line, text)
    
    dim format, equ, txt, newfile, temp_path
    if not exists(file) then WScript.Quit
    
    equ         = 0
    temp_path   = env("temp") & "\__temp.txt"
    
    set newfile = fso.CreateTextFile(temp_path, True)
    set format  = fso.OpenTextFile(file, 1)

    do until format.AtEndOfStream
        equ = equ+1
        txt = format.ReadLine
        
        if equ = line then txt = text
        newfile.WriteLine txt
    loop
    
    newfile.close
    format.close
    
    if not is_same(file, temp_path) then 
        msgbox err(0) 
        WScript.Quit
    end if
    
    fso.DeleteFile file
    fso.MoveFile   temp_path, file
    
end function
 
 
if WScript.Arguments.Count < 3 then 
    if fso.GetFileName(WScript.ScriptFullName) <> "LineChanger.vbs" then 
        MsgBox "Missing start parameters", 0, "Error"
        WScript.quit()
    end if
    
else 
    ChangeLine WScript.Arguments(0), WScript.Arguments(1), WScript.Arguments(2)
end if
ChangeLine "C:\Temp\projects\topline\linechanger\tests\t.txt", Array(5,6, 7), Array("Hello", "World")