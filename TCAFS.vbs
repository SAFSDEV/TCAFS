Option Explicit

'******************************************************************************
'*
'* Optional script parameters:
'*     
'*     -safs.project.config fullpath(s)
'*         
'*         Optional: One or more fullpaths separated by semi-colons pointing 
'*         to the SAFS INI files containing configuration information and options.
'*         Internally, there is no default value.  However, the SAFSTC Engine
'*         Class will attempt to provide this information if the SAFS is 
'*         configured to AUTOLAUNCH the SAFSTC Engine.
'*         
'*         Examples: 
'*         
'*             -safs.project.config "C:\SAFS\Project\tcafs.ini"
'*             -safs.project.config "C:\SAFS\Project\tcafs.ini;C:\SAFS\Project\safstid.ini"
'*         
'*     -suitename
'*         
'*         Optional: An alternate Project Suite to use instead of the default.
'*         The SAFS default is "C:\SAFS\TCAFS\TCAFS.pjs".
'*         
'*         Example: 
'*         
'*             -suitename "C:\Some\Other\Suite.pjs"
'*         
'*     -projectname
'*
'*         Optional: An alternate Project name to use instead of the default.
'*         The SAFS default is "TCAFS".
'*         
'*         Example: 
'*         
'*             -projectname "MyProject"
'*         
'*     -scriptname
'*
'*         Optional: An alternate Script name to use instead of the default.
'*         The SAFS default is "StepDriver".
'*         
'*         Example: 
'*         
'*             -scriptname "MyScript"
'*         
'*     -passthru
'*
'*         Optional: An argument that allows you to pass-thru one or more other
'*         custom Test Complete command-line arguments.
'*         Note: It appears that Test Complete will only passthru arguments 
'*         preceded with a forward-slash (/). So the value for this argument 
'*         must contain a "/" prefix in order for it to be passed along by
'*         Test Complete.
'*         
'*         Example: 
'*         
'*             -passthru "/customArg:value /AnotherOne"
'*         
'*     
'* Copyright (C) SAS Institute
'* General Public License: http://www.opensource.org/licenses/gpl-license.php
'******************************************************************************

Dim status
Dim message
Dim details
Dim projectname
Dim suitename
Dim scriptname
Dim command
Dim shell, exec, env, fso
Dim i,args,arg,lcarg,safsconfig
Dim passthru
Dim binpath, executable

projectname = "TCAFS"
suitename = "C:\SAFS\TCAFS\TCAFS.pjs"
scriptname = "StepDriver"
safsconfig = ""
passthru = ""
binpath = "%TESTCOMPLETE_HOME%\bin\"
executable = binpath & "TestComplete.exe"

Set shell = WScript.CreateObject("WScript.Shell")
Set env   = shell.Environment("SYSTEM")
Set args  = WScript.Arguments

' loop thru all args
'======================
For i = 0 to args.Count -1
    arg = args(i)
        
    'remove any trailing '\' or '/'
    '====================================================
    if ((Right(arg,1)="\")or(Right(arg,1)="/")) then
        arg = Left(arg, Len(arg)-1)
    end if
    
    lcarg = lcase(arg)
    
    'check safsdir and stafdir alternate install locations
    '====================================================
    if (arg = "-safs.project.config") then
        if ( i < args.Count -1) then
            arg = args(i+1)
            if ((Right(arg,1)="\")or(Right(arg,1)="/")) then
                arg = Left(arg, Len(arg)-1)
            end if
            if(Len(arg)>0) then
            	'leading space below is REQUIRED
            	safsconfig=" /safs.project.config:"& arg
            end if
        end if
    elseif (arg = "-suitename") then
        if ( i < args.Count -1) then
            arg = args(i+1)
            if ((Right(arg,1)="\")or(Right(arg,1)="/")) then
                arg = Left(arg, Len(arg)-1)
            end if
            if(Len(arg)>0) then
            	suitename = arg
            end if
        end if
    elseif (arg = "-projectname") then
        if ( i < args.Count -1) then
            arg = args(i+1)
            if(Len(arg)>0) then
            	projectname = arg
            end if
        end if
    elseif (arg = "-scriptname") then
        if ( i < args.Count -1) then
            arg = args(i+1)
            if(Len(arg)>0) then
            	scriptname = arg
            end if
        end if
    elseif (arg = "-passthru") then
        if ( i < args.Count -1) then
            arg = args(i+1)
            if(Len(arg)>0) then            
            	'add space prefix to separate arg, if not already present
            	if (StrComp(Left(arg,1)," ",0) <> 0) then arg = " "& arg
            	passthru = arg
            end if
        end if
    end if    
Next

status = env("TESTCOMPLETE_EXE")
if Len(status) > 0 then executable = binpath & status

command = executable &" "& suitename & " /r /p:" & projectname & " /u:" & scriptname & " /rt:Main /e /SilentMode /ns"& safsconfig & passthru

message = "Command"
details = command
On Error Resume Next
Set exec = shell.Exec(command)
If Err.Number <> 0 Then
    'If we get here, Test Complete is already running
    msgbox "Test Complete is running, Stop all test processes including STAF and Try again"
End if