Attribute VB_Name = "NetzVerbIni"
Option Explicit
Public oEnvSystem As New System
Public ErrNumber&, ErrDescription$, FNr&
Public FPos& ' Fehlerposition
Type OSVERSIONINFO
       dwOSVersionInfoSize As Long
       dwMajorVersion As Long
       dwMinorVersion As Long
       dwBuildNumber As Long
       dwPlatformId As Long
       szCSDVersion As String * 128 ' Service Pack
End Type
Type OSVERSIONINFOEX
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128 ' Service Pack
        wServicePackMajor As Integer
        wServicePackMinor As Integer
        wSuiteMask As Integer
        wProductType As Byte
        wReserved As Byte
End Type
Public userprof$
Public WV As WindowsVersion
Declare Function GetVersionEx1& Lib "kernel32.dll" Alias "GetVersionExA" (ByRef LpVersionInformation As OSVERSIONINFO)
Declare Function GetVersionEx2& Lib "kernel32.dll" Alias "GetVersionExA" (ByRef LpVersionInformation As OSVERSIONINFOEX)
Declare Function GetShortPathName& Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath$, ByVal lpszShortPath$, ByVal cchBuffer&)
Declare Function OpenProcess& Lib "kernel32.dll" (ByVal dwDesiredAccess&, ByVal bInheritHandle&, ByVal dwProcId&)
Declare Function timeGetTime& Lib "winmm.dll" ()
Declare Function GetExitCodeProcess& Lib "kernel32" (ByVal hProcess&, lpExitCode&)

Const PROCESS_QUERY_INFORMATION As Long = 1024 ' &H400
Private Const STILL_ACTIVE = &H103
Public FSO As New FileSystemObject
Declare Function GetVersionEx& Lib "kernel32" Alias "GetVersionExA" (ByRef LpVersionInformation As OSVERSIONINFO)
Const VER_PLATFORM_WIN32_NT     As Long = 2

Private Const MAX_PATH                  As Long = 260
Private Const TH32CS_SNAPPROCESS        As Long = &H2&
Private Type PROCESSENTRY32 ' Prozesseintrag
   dwSize As Long
   cntUsage As Long
   th32ProcessID As Long
   th32DefaultHeapID As Long
   th32ModuleID     As Long
   cntThreads As Long
   th32ParentProcessID As Long
   pcPriClassBase As Long
   dwFlags As Long
   szExeFile As String * MAX_PATH ' = 260
End Type
Private Declare Function CreateToolhelp32Snapshot& Lib "kernel32" (ByVal dwFlags&, ByVal th32ProcessID&)
Private Declare Function Process32First& Lib "kernel32" (ByVal hSnapShot&, ByRef lppe As PROCESSENTRY32)
Declare Function FindWindow& Lib "user32" Alias "FindWindowA" (ByVal lpClassName$, ByVal lpWindowName$)
Declare Function GetParent& Lib "user32" (ByVal hwnd&)
Declare Function GetWindowThreadProcessId& Lib "user32" (ByVal hwnd&, lpdwProcessId&)
Private Const GW_HWNDNEXT = 2
Private Declare Function GetWindow& Lib "user32" (ByVal hwnd&, ByVal wCmd&)
Private Const SW_RESTORE = 9 ' Restore window
Private Declare Function ShowWindow& Lib "user32.dll" (ByVal hwnd&, ByVal nCmdShow&)
Private Declare Sub SetForegroundWindow Lib "user32" (ByVal hwnd&)
Public Const WM_CLOSE = &H10
Public Declare Function PostMessage& Lib "user32" Alias "PostMessageA" (ByVal hwnd&, ByVal wMsg&, ByVal wParam&, lParam As Any)
Private Const PROCESS_TERMINATE         As Long = &H1
Declare Function TerminateProcess& Lib "kernel32" (ByVal hProcess&, ByVal uExitCode&)
Declare Function CloseHandle& Lib "kernel32.dll" (ByVal Handle&)
Const SYNCHRONIZE               As Long = &H100000
Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type
Private Declare Function CreateJobObject Lib "kernel32.dll" Alias "CreateJobObjectA" (lpJobAttributes As SECURITY_ATTRIBUTES, lpName As String) As Long
Public Declare Function AssignProcessToJobObject Lib "kernel32" (ByVal hJob As Long, ByVal hProcess As Long) As Long
Public Declare Function TerminateJobObject Lib "kernel32" (ByVal hJob As Long, ByVal hProcess As Long) As Long
Private Const PROCESS_VM_READ           As Long = 16
Private Declare Function Process32Next& Lib "kernel32" (ByVal hSnapShot&, ByRef lppe As PROCESSENTRY32)
Private Declare Function EnumProcesses& Lib "psapi.dll" (ByRef lpidProcess&, ByVal cb&, ByRef cbNeeded&)
Private Declare Function EnumProcessModules& Lib "psapi.dll" (ByVal hProcess&, ByRef lphModule&, ByVal cb&, ByRef cbNeeded&)
Private Declare Function GetModuleFileNameEx& _
  Lib "psapi.dll" Alias "GetModuleFileNameExA" ( _
  ByVal hProcess&, _
  ByVal hModule&, _
  ByVal ModuleName$, _
  ByVal nSize& _
  )

Public uVerz$, vVerz$, pVerz$, plzVz$


Sub Main()
 Const nverb$ = "nverb.exe"
 Const nn$ = "\nachricht.exe"
 Const br$ = "\berechtigungen.exe"
 Const Pw$ = "\Programmierung\Netzverbind"
' Const Pi$ = "\Programmierung\NVIni"
' Const PiExe$ = "\NVIni.exe.lnk"
 Dim pu$, pz$, runde%, ird%, pf$(7), pp$, Du$, ngef%, ausgStr$ ', piakt$
 oEnvSystem.Environment("NVIni") = 1
 'C:\Windows\System32\cmd.exe /k %windir%\System32\reg.exe ADD HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System /v EnableLUA /t REG_DWORD /d 0 /f
 'C:\Windows\System32\cmd.exe /k %windir%\System32\reg.exe ADD HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System /v EnableLUA /t REG_DWORD /d 1 /f
 pf(1) = "u:"
 pf(2) = "\\linux1\daten\eigene Dateien"
 pf(3) = "\\linux\daten\eigene Dateien"
 pf(4) = "\\mitte1\u"
 pf(5) = "\\anmeldl\u"
 pf(6) = "\\anmeldr\eigeneDateien"
 pf(7) = "\\linserv\eigene Dateien"
 On Error GoTo fehler
 FPos = 1
 pp = Environ("localappdata")
 If LenB(pp) = 0 Then pp = Environ("appdata") ' Windows XP
' pp = Environ("ProgramFiles(x86)")
' If LenB(pp) = 0 Then pp = Environ("programfiles") ' 32-bit-Architektur
 FPos = 2
 On Error Resume Next
 Err.Clear
 For runde = 1 To 2 ' UBound(pf)
   Err.Clear
   pu = pf(runde) & Pw & "\" & nverb
'   piakt = pf(runde) & Pi & PiExe
   Du = Dir(pu)
   If Err.Number = 0 And LenB(Du) <> 0 Then
    If runde > 2 Then
     Call Shell(pp & nn & " " & "Laufwerk u nicht gefunden. Nehme `" & pf(runde) & "`")
    End If
    Exit For
   End If
   If runde = UBound(pf) Then
    ausgStr = "Folgende Laufwerke nicht gefunden:"
    For ird = 1 To UBound(pf)
     ausgStr = ausgStr & vbCrLf & pf(ird)
    Next ird
    Call Shell(pp & nn & " " & ausgStr)
    ngef = True
   End If
 Next runde
 On Error GoTo fehler
 FPos = 3
 pz = pp & "\" & nverb
 If ngef = 0 Then
 KWn nverb, pf(runde) & Pw, pp
 KWn "nachricht.exe", pf(runde) & Pw, pp
 KWn "berechtigungen.exe", pf(runde) & Pw, pp
'  Call KopierWennNeuer(pu, pz)
'  Call KopierWennNeuer(Replace(pu, nv, nn), pp & nn)
'  Call KopierWennNeuer(Replace(pu, nv, br), pp & br)
''  Call KopierWennNeuer(piakt, Environ("allusersprofile") & "\Startmenü\Programme\Autostart" & PiExe)
 End If
 On Error Resume Next
 runde = 1
 Do
  Err.Clear
  If LenB(Dir(pz)) <> 0 Then
'   Call Shell(pz, vbMinimizedFocus)
   rufauf pz, , , , 0, 1
   Exit Do
  End If
  runde = runde + 1
  If runde > UBound(pf) Then Exit Do
  KopierWennNeuer pf(runde) & Pw & "\" & nverb, pz
 Loop
 ProgEnde
 Exit Sub
fehler:
 Select Case MsgBox("Fpos: " & FPos & " ErrNr: " & CStr(Err.Number) & vbCrLf & "LastDLLError: " & CStr(Err.LastDllError) & vbCrLf & "Source: " & IIf(IsNull(Err.Source), vbNullString, CStr(Err.Source)) & vbCrLf & "Description: " & Err.Description & vbCrLf, vbAbortRetryIgnore, "Aufgefangener Fehler in Main/" & App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' main

Function ProgEnde()
 oEnvSystem.Environment("NVIni") = vbNullString
 End
End Function ' ProgEnde

#If False Then
Public Function GetOSVersion() As WindowsVersion
' Konstanten
  Const VER_PLATFORM_WIN32s As Long = 0&
  Const VER_PLATFORM_WIN32_WINDOWS As Long = 1&
  Const VER_PLATFORM_WIN32_NT As Long = 2&

' Um zu testen, ob XP Home oder Professional verwendet wird.
' Weitere Informationen gibt es unter
' http://msdn.microsoft.com/library/en-us/sysinfo/base/getversionex.asp
' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/ ->
'     sysinfo/base/osversioninfoex_str.asp

  Const VER_SUITE_PERSONAL As Long = &H200&
 
' Private Variablen
  Static m_bAlreadyGot As Boolean
  Static m_OsVersion As WindowsVersion

' WinAPI

    Dim OsVersInfoEx As OSVERSIONINFOEX
    Dim OsVersInfo As OSVERSIONINFO
    
    On Error GoTo fehler
    
    If m_bAlreadyGot Then
        GetOSVersion = m_OsVersion
        Exit Function
    End If
    
    ' Zuerst nehmen wir nur die kleinere Struktur um sicherzugehen,
    ' dass Windows 95 und Konsorten auch damit klarkommen
    OsVersInfo.dwOSVersionInfoSize = Len(OsVersInfo)
    
    If GetVersionEx1(OsVersInfo) = 0 Then
'       oboffenlassen = 0    ' 27.9.09 für Wien
'       Ende                 ' 27.9.09 für Wien
        MsgBox "Das Betriebssystem konnte nicht korrekt erkannt " & _
        "werden:" & _
                vbCrLf & "Fehler im API-Aufruf"
        
        m_OsVersion = WIN_OLD
        Exit Function
    End If
        
    With OsVersInfo
        Select Case .dwPlatformId
            Case VER_PLATFORM_WIN32s
                m_OsVersion = WIN_OLD
            Case VER_PLATFORM_WIN32_WINDOWS
                Select Case .dwMinorVersion
                    Case 0
                        m_OsVersion = WIN_95
                    Case 10
                        m_OsVersion = WIN_98
                    Case 90
                        m_OsVersion = WIN_ME
                End Select
            Case VER_PLATFORM_WIN32_NT
                Select Case .dwMajorVersion
                    Case 3
                        m_OsVersion = WIN_NT_3x
                    Case 4
                        m_OsVersion = win_nt_4x
                    Case 5
                        Select Case .dwMinorVersion
                            Case 0
                                m_OsVersion = win_2k
                            Case 1
                                
' Es handelt sich um Windows XP. Um zu erfahren, ob das verwendete
' Produkt eine Home-Edition ist, erfragen wir die Version erneut und
' empfangen dieses Mal die komplette Liste
                                
                                OsVersInfoEx.dwOSVersionInfoSize = Len(OsVersInfoEx)
                                If GetVersionEx2(OsVersInfoEx) = 0 Then
'       oboffenlassen = 0    ' 27.9.09 für Wien
'       Ende                 ' 27.9.09 für Wien
                                    MsgBox "Das Betriebssystem konnte nicht korrekt erkannt werden:" & _
                                        vbCrLf & "Fehler im API-Aufruf"
                                    m_OsVersion = win_xp
                                    Exit Function
                                End If
                                If (OsVersInfoEx.wSuiteMask And VER_SUITE_PERSONAL) = VER_SUITE_PERSONAL Then
                                    m_OsVersion = win_xP_home
                                Else
                                    m_OsVersion = win_xp
                                End If
                            Case 2
                                m_OsVersion = WIN_2003
                        End Select
                    Case Else
                        m_OsVersion = .dwMajorVersion + 6 ' getestet: Windows 8
                End Select
        End Select
    End With
    GetOSVersion = m_OsVersion
    m_bAlreadyGot = True
    Exit Function
fehler:
Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in GetOSVersion/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' GetOSVersion
#End If

Public Function ShellaW( _
        sShell As String, _
        Optional ByVal eWindowStyle As VBA.VbAppWinStyle = vbNormalFocus, _
        Optional ByRef sError As String, _
        Optional ByVal lTimeOut As Long = 2000000000 _
    ) As Boolean
Dim hProcess As Long
Dim lR As Long
Dim lTimeStart As Long
Dim bSuccess As Boolean
    
On Error GoTo ShellAndWaitForTerminationError
    
    ' This is v2 which is somewhat more reliable:
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(sShell, eWindowStyle))
    If (hProcess = 0) Then
        sError = "This program could not determine whether the process started." & _
             "Please watch the program and check it completes."
        ' Only fail if there is an error - this can happen
        ' when the program completes too quickly.
    Else
        bSuccess = True
        lTimeStart = timeGetTime()
        Do
            ' Get the status of the process
            GetExitCodeProcess hProcess, lR
            ' Sleep during wait to ensure the other process gets
            ' processor slice:
            DoEvents: Sleep 100
            If (timeGetTime() - lTimeStart > lTimeOut) Then
                ' Too long!
                sError = "The process has timed out."
                lR = 0
                bSuccess = False
            End If
        Loop While lR = STILL_ACTIVE
    End If
    ShellaW = bSuccess
        
    Exit Function

ShellAndWaitForTerminationError:
    sError = Err.Description
    Exit Function
End Function ' ShellaW

Public Sub machOrdner(tStr$)
 On Error GoTo fehler
  If LenB(tStr) <> 0 Then
    If Right(tStr, 1) = "\" Then tStr = Left(tStr, Len(tStr) - 1)
    If WV <= win_vista Then
     If Not FSO.FolderExists(tStr) And InStrB(Mid$(tStr, 3), "\") <> 0 Then 'And Not left(TStr, 2) = "\\" And InStr(mid$(TStr, 3), "\") = 1 Then
      Call FSO.CreateFolder(tStr)
     End If
    Else
     If Dir(tStr, vbDirectory) = "" Then
      If AdminPwd = "" Then AdminPwd = holap("Administrator")
      Shell ("\\linux1\daten\down\pstools\psexec -u administrator -p " & AdminPwd & " cmd /e:on /c mkdir " & Chr(34) & tStr & Chr(34))
     End If
    End If
  End If
  Exit Sub
fehler:
Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in machOrdner/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' machOrdner

Function VerzPrüf(ByVal Verz$)
 Dim Bstd$(), i%, j%, k%, tStr$
 On Error GoTo fehler
' If FSO Is Nothing Then Set FSO = CreateObject("Scripting.FileSystemObject")
 VerzPrüf = FSO.GetAbsolutePathName(Verz)
 Bstd = Split(VerzPrüf, "\")
 For j = 0 To UBound(Bstd)
  tStr = ""
  For i = 0 To j
   tStr = tStr + IIf(i = 0, "", "\") + Bstd(i)
  Next i
  machOrdner tStr
 Next
 Exit Function
fehler:
Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in VerzPrüf/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' VerzPrüf(ByVal Verz$)

Private Function IsWindowsNT() As Boolean
' Gibt True für Windows NT (und 2000, XP, 2003, Vista) zurück
Dim OSInfo As OSVERSIONINFO
  On Error GoTo fehler
  With OSInfo
    .dwOSVersionInfoSize = Len(OSInfo)  ' Angabe der Größe dieser Struktur
    .szCSDVersion = Space$(128)         ' Speicherreservierung für Angabe des Service Packs
    GetVersionEx OSInfo                 ' OS-Informationen ermitteln
    IsWindowsNT = (.dwPlatformId = VER_PLATFORM_WIN32_NT) ' für Windows NT
  End With
  Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescription = Err.Description
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vbNullString, CStr(Err.Source)) + vbCrLf + "Description: " + ErrDescription + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in IsWindowsNT/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' IsWindowsNT() As Boolean

Private Function TrimNullChar(ByVal s As String) As String
' Kürzt einen String s bis zum Zeichen vor einem vbNullChar
  Dim Pos1 As Long
  On Error GoTo fehler
  ' vbNullChar = Chr$(0) im String suchen
  Pos1 = InStr(s, vbNullChar)
  ' Falls vorhanden, den String entsprechend kürzen
  If (Pos1 > 0) Then
    TrimNullChar = Left$(s, Pos1 - 1)
  Else
    TrimNullChar = s
  End If
  Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescription = Err.Description
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vbNullString, CStr(Err.Source)) + vbCrLf + "Description: " + ErrDescription + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in TrimNullChar/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' TrimNullChar(ByVal s As String) As String

Function ProcIDFromWnd(ByVal hwnd As Long) As Long
   Dim idProc As Long
   On Error GoTo fehler
   ' Get PID for this HWnd
   GetWindowThreadProcessId hwnd, idProc

   ' Return PID
   ProcIDFromWnd = idProc
  Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescription = Err.Description
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vbNullString, CStr(Err.Source)) + vbCrLf + "Description: " + ErrDescription + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in ProcIDFromWnd/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' ProcIDFromWnd(ByVal hwnd As Long) As Long

Function GetWinHandle(hInstance As Long) As Long
   Dim tempHwnd As Long
   On Error GoTo fehler
   ' Grab the first window handle that Windows finds:
   tempHwnd = FindWindow(vbNullString, vbNullString)

   ' Loop until you find a match or there are no more window handles:
   Do Until tempHwnd = 0
      ' Check if no parent for this window
      If GetParent(tempHwnd) = 0 Then
         ' Check for PID match
         If hInstance = ProcIDFromWnd(tempHwnd) Then
            ' Return found handle
            GetWinHandle = tempHwnd
            ' Exit search loop
            Exit Do
         End If
      End If

      ' Get the next window handle
      tempHwnd = GetWindow(tempHwnd, GW_HWNDNEXT)
   Loop
  Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescription = Err.Description
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vbNullString, CStr(Err.Source)) + vbCrLf + "Description: " + ErrDescription + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in GetWinHandle/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' GetWinHandle(hInstance As Long) As Long

Public Function KillProcessByPID(ByVal pid As Long) As Boolean
' Versucht auf Basis einer Prozess-ID, den zugehörigen
' Prozess zu terminieren. Im Erfolgsfall wird True zurückgegeben.
  Dim hProcess As Long, tpid&, p&, nRet
  On Error GoTo fehler
  ' Öffnen des Prozesses über seine Prozess-ID
  hProcess = OpenProcess(PROCESS_TERMINATE, False, pid)
  If hProcess = 0 Then
   If True Then
' hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, pid)
    hProcess = OpenProcess(0, False, pid)
    GetExitCodeProcess hProcess, nRet
    Call TerminateProcess(hProcess, nRet)
    Call CloseHandle(hProcess)
   ElseIf True Then
    hProcess = OpenProcess(SYNCHRONIZE, False, pid)
    Dim jh&, n As SECURITY_ATTRIBUTES
    jh = CreateJobObject(n, "kill")
    GetWindowThreadProcessId pid, tpid
    p = OpenProcess(2035711, 0, pid)
    p = OpenProcess(PROCESS_QUERY_INFORMATION, 0, pid)
    AssignProcessToJobObject jh, p
    TerminateJobObject jh, 0
   End If
  End If
  ' Gibt es ein Handle, wird der Prozess darüber abgeschossen
  If (hProcess <> 0) Or True Then
    KillProcessByPID = TerminateProcess(hProcess, 1&) <> 0
    CloseHandle hProcess
  Else
   hProcess = OpenProcess(PROCESS_VM_READ, False, pid)
   hProcess = OpenProcess(&H1, False, pid)
   Stop
  End If
  Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescription = Err.Description
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vbNullString, CStr(Err.Source)) + vbCrLf + "Description: " + ErrDescription + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in KillProcessByPID/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' KillProcessByPID(ByVal pid As Long) As Boolean

Public Function GetProcessCollection(Optional obkill%, Optional exe$) As Collection
' Ermittelt die abfragbaren laufenden Prozesse des lokalen
' Rechners. Jeder gefundene Prozess wird mit seiner ID
' als String in einem Element der Rückgabe-Collection
' gespeichert im Format "Prozessname|Prozess-ID".
  Dim collProcesses As New Collection
  Dim ProcID As Long
  Dim hdl&
  On Error GoTo fehler
  If (Not IsWindowsNT) Then
  
    ' WINDOWS 95 / 98 / Me
    ' --------------------
  
    Dim sName As String
    Dim hSnap As Long
    Dim pEntry As PROCESSENTRY32
  
    ' Einen Snapshot der Prozessinformationen erstellen
    hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If hSnap = 0 Then
      Exit Function ' Pech gehabt
    End If
  
    pEntry.dwSize = Len(pEntry) ' Größe der Struktur zur Verfügung stellen
  
    ' Den ersten Prozess im Snapshot ermitteln
    ProcID = Process32First(hSnap, pEntry)
  
    ' Mittels Process32Next über alle weiteren Prozesse iterieren
    Do While (ProcID <> 0) ' Gibt es eine gültige Prozess-ID?
      sName = TrimNullChar(pEntry.szExeFile)  ' Rückgabestring stutzen
      collProcesses.Add sName & "|" & CStr(ProcID) ' Collection-Eintrag
          If obkill <> 0 Then
           If InStrB(LCase$(sName), LCase$(exe)) <> 0 Then
            Select Case obkill
             Case 1
              hdl = GetWinHandle(ProcID)
              If hdl <> 0 Then
               ShowWindow hdl, SW_RESTORE
               SetForegroundWindow hdl
               PostMessage hdl, WM_CLOSE, 0&, 0&
               Exit Function
              End If
             Case 2
              If KillProcessByPID(ProcID) <> 0 Then Exit Function
            End Select
           End If
          End If
      ProcID = Process32Next(hSnap, pEntry)   ' Nächste PID des Snapshots
    Loop
  
    ' Handle zum Snapshot freigeben
    CloseHandle hSnap
  
  Else
    ' WINDOWS NT / 2000 / XP / 2003 / Vista
    ' -------------------------------------
    Dim cb As Long
    Dim cbNeeded As Long
    Dim RetVal As Long
    Dim NumElements As Long
    Dim ProcessIDs() As Long
    Dim cbNeeded2 As Long
    Dim NumElements2 As Long
    Dim Modules(1) As Long
    Dim ModuleName As String
    Dim LenName As Long
    Dim hProcess As Long
    Dim i As Long
  
    cb = 8         ' "CountBytes": Größe des Arrays (in Bytes)
    cbNeeded = 9   ' cbNeeded muss initial größer als cb sein
  
    ' Schrittweise an die passende Größe des Prozess-ID-Arrays
    ' heranarbeiten. Dazu vergößern wir das Array großzügig immer
    ' weiter, bis der zur Verfügung gestellte Speicherplatz (cb)
    ' den genutzten (cbNeeded) überschreitet:
    Do While cb <= cbNeeded ' Alle Bytes wurden belegt -
                            ' es könnten also noch mehr sein
      cb = cb * 2                      ' Speicherplatz verdoppeln
      ReDim ProcessIDs(cb / 4) As Long ' Long = 4 Bytes
      EnumProcesses ProcessIDs(1), cb, cbNeeded ' Array abholen
    Loop
  
    ' In cbNeeded steht der übergebene Speicherplatz in Bytes.
    ' Da jedes Element des Arrays als Long aus 4 Bytes besteht,
    ' ermitteln wir die Anzahl der tatsächlich übergebenen
    ' Elemente durch entsprechende Division:
    NumElements = cbNeeded / 4
  
    ' Jede gefundene Prozess-ID des Arrays abarbeiten
    For i = 1 To NumElements
  
      ' Versuchen, den Prozess zu öffnen und ein Handle zu erhalten
      hProcess = OpenProcess(PROCESS_QUERY_INFORMATION _
                          Or PROCESS_VM_READ, _
                             0, ProcessIDs(i))
  
      If (hProcess <> 0) Then ' OpenProcess war erfolgreich
  
        ' EnumProcessModules gibt die dem Prozess angehörenden
        ' Module in einem Array zurück.
        RetVal = EnumProcessModules(hProcess, Modules(1), 1, cbNeeded2)
  
        If (RetVal <> 0) Then ' EnumProcessModules war erfolgreich
          ModuleName = Space$(MAX_PATH) ' Speicher reservieren
          ' Den Pfadnamen für das erste gefundene Modul bestimmen
          LenName = GetModuleFileNameEx(hProcess, Modules(1), ModuleName, Len(ModuleName))
          ' Den gefundenen Pfad und die Prozess-ID unserer
          ' ProcessCollection hinzufügen (Trennzeichen "|")
          collProcesses.Add Left$(ModuleName, LenName) & "|" & CStr(ProcessIDs(i))

          If obkill <> 0 Then
           Debug.Print ModuleName
           If InStrB(LCase$(ModuleName), LCase$(exe)) <> 0 Then
            Select Case obkill
             Case 1, -1
              hdl = GetWinHandle(ProcessIDs(i))
              If hdl <> 0 Then
               ShowWindow hdl, SW_RESTORE
               SetForegroundWindow hdl
               PostMessage hdl, WM_CLOSE, 0&, 0&
               CloseHandle hProcess ' Offenes Handle schließen
               Exit Function
              End If
             Case 2
              If KillProcessByPID(ProcessIDs(i)) <> 0 Then
               Exit Function
              Else
               Call Shell("v:\pcwkill.exe " & exe, vbMaximizedFocus)
               CloseHandle hProcess ' Offenes Handle schließen
               Exit Function
              End If
            End Select
           End If
          End If
        End If
      End If
      CloseHandle hProcess ' Offenes Handle schließen
    Next i
  End If
  ' Zusammengestellte Collection übergeben
  Set GetProcessCollection = collProcesses
 Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescription = Err.Description
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vbNullString, CStr(Err.Source)) + vbCrLf + "Description: " + ErrDescription + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in GetProcessCollection /" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' GetProcessCollection(Optional obKill%, Optional exe$) As Collection

Sub KWn(D$, ByVal v1$, ByVal v2$) ' Kopiere wenn neuer
 'Dim FSO As New FileSystemObject
 Dim D1, D2, voll1$, voll2$, obKopier%
 On Error GoTo fehler
 If WV = 0 Then WV = GetOSVersion
 If userprof = "" Then userprof = Environ("userprofile")
 Select Case WV
    Case Is < win_vista ' win_xp, win_xP_home
    Case Else
     Dim Result&, AA$
     AA = Space$(255)
     Result = GetShortPathName(v1, AA, Len(AA))
'     If WV < win_vista Then v1 = Mid$(AA, 1, Result)
     AA = Space$(255)
     Result = GetShortPathName(v2, AA, Len(AA))
     If Result = 0 Then
'      Shell "runas /user:administrator ""cmd /c md """ & v2 & """ "" "
      Shell "\\linux1\daten\down\pstools\psexec -u administrator -p sonne cmd /e:on /c md """ & v2 & """", vbHide
      AA = Space$(255)
      Result = GetShortPathName(v2, AA, Len(AA))
     End If
     If Result = 0 Then
      Dim werklverz$, wvlen%
      werklverz = userprof & "\werkl"
      wvlen = Len(werklverz)
      Shell "cmd /c md """ & werklverz & """"
      Shell "cmd /c del """ & werklverz & "\aktpfad.txt"""
      Shell "cmd /c echo @echo %~s1 ^> " & werklverz & "\aktpfad.txt > " & werklverz & "\zeigkurz.bat"
      ShellaW ("cmd /c " & werklverz & "\zeigkurz.bat """ & v2 & """"), vbHide, , 10000
      Open werklverz & "\aktpfad.txt" For Input As #98
      Dim Text$
      While Not EOF(98)
       Line Input #98, Text
       If Result = 0 Then Result = Len(Text)
      Wend
      Close #98
      If Left$(Text, wvlen) = werklverz Then
       MsgBox "Falscher Pfad '" & v2 & "' in NVerb"
       Exit Sub
      Else
      
      End If
     Else
'    If WV < win_vista Then v2 = Mid$(AA, 1, Result)
     End If
 End Select
 obKopier = 0
 voll1 = v1 & IIf(Right(v1, 1) = "\", "", "\") & D
 voll2 = v2 & IIf(Right(v2, 1) = "\", "", "\") & D
 If WV < win_vista Then Call VerzPrüf(v2 & IIf(Right(v2, 1) = "\", "", "\"))
 If FSO.FileExists(voll1) Then
  Set D1 = FSO.GetFile(voll1)
  If FSO.FileExists(voll2) Then
   If Not FSO.FileExists(voll2) Then
    obKopier = -1
   Else
    Set D2 = FSO.GetFile(voll2)
    On Error Resume Next
    If D1.DateLastModified > D2.DateLastModified Then obKopier = -1
    On Error GoTo fehler
   End If
  Else
   obKopier = -1
  End If
  If obKopier Then
   If InStrB(v1, "poetaktiv") <> 0 Then
    Call GetProcessCollection(2, "poetaktiv")
   End If
   Dim d1str$
   d1str = D1
   Call KopDat(d1str, voll2)
  End If
 End If
Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos) & vbCrLf & "Datei: " & D & vbCrLf & "Verzeichnis 1:" & v1 & vbCrLf & "Verzeichnis 2:" & v2, vbAbortRetryIgnore, "Aufgefangener Fehler in KWn/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' KWn

Function KopierWennNeuer(pu$, pz$)
#If False Then
 Dim Dz$, Du$
 Dim DatU As Date, DatZ As Date
 Dim FlU&, FlZ&
 Du = Dir(pu)
 Dz = Dir(pz)
 FPos = 4
 If Du <> vbNullString Then
  FPos = 5
  DatU = FileDateTime(pu)
  FPos = 6
  FlU = FileLen(pu)
  FPos = 7
  If Dz <> vbNullString Then
  FPos = 8
   DatZ = FileDateTime(pz)
  FPos = 9
   FlZ = FileLen(pz)
  FPos = 10
  End If
  If (DatU > DatZ) Or (DatZ = 0 And FlU <> 0) Then
  FPos = 11
   On Error Resume Next
   Call KopDat(pu, pz)
   On Error GoTo fehler
  End If
 End If
 Exit Function
fehler:
 Select Case MsgBox("Fpos: " & FPos & " ErrNr: " & CStr(Err.Number) & vbCrLf & "LastDLLError: " & CStr(Err.LastDllError) & vbCrLf & "Source: " & IIf(IsNull(Err.Source), vbNullString, CStr(Err.Source)) & vbCrLf & "Description: " & Err.Description & vbCrLf & pu & vbCrLf & pz, vbAbortRetryIgnore, "Aufgefangener Fehler in KopierWennNeuer/" & App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
#End If
End Function ' KopierWennNeuer

Sub adminaktiv()
 Dim Text$
 On Error GoTo fehler
' Shell ("cmd /c net user administrator > " & Environ("appdata") & "\obadmin.txt")
 rufauf "cmd", "/c net user administrator > """ & Environ("appdata") & "\obadmin.txt""", , , 0, 0
 Do While Dir(Environ("appdata") & "\obadmin.txt") = ""
 Loop
 Do While FileLen(Environ("appdata") & "\obadmin.txt") < 1000 ' 1044 ist die Zielgröße
 Loop
 Open Environ("appdata") & "\obadmin.txt" For Input As #193
 Do While Not EOF(193)
  Line Input #193, Text
  If Left(Text, 11) = "Konto aktiv" Then
   If Text Like "*Nein*" Then
    MsgBox "Bitte aktivieren Sie den Administrator mit 'net user Administrator * /active:yes' und starten Sie das Programm dann nochmal!"
    Unload FürIcon
    End
' geht nicht, da der Administrator nicht vom Nicht-Administrator aktiviert werden kann
'    Shell ("cmd.exe /c net user Administrator /active:yes")
   End If
   Exit Do
  End If
 Loop
 Close #193
 Do While LenB(Dir(Environ("appdata") & "\obadmin.txt")) = 0
  On Error Resume Next
  Kill Environ("appdata") & "\obadmin.txt"
 Loop
 Exit Sub
fehler:
 Select Case MsgBox("Fpos: " & FPos & " ErrNr: " & CStr(Err.Number) & vbCrLf & "LastDLLError: " & CStr(Err.LastDllError) & vbCrLf & "Source: " & IIf(IsNull(Err.Source), vbNullString, CStr(Err.Source)) & vbCrLf & "Description: " & Err.Description & vbCrLf, vbAbortRetryIgnore, "adminaktiv/" & App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' adminaktiv

Public Sub KopDat(q$, z$, Optional WV As WindowsVersion)
 Static iWV As WindowsVersion, ErrNr&
 Static pruefe%
 Dim Datei$
 On Error GoTo fehler
 Datei = Right$(q, Len(q) - InStrRev(q, "\"))
 On Error Resume Next
 Kill z & Datei
 FileCopy q, z
 If LenB(Dir(z)) = 0 And IIf(Right$(z, 1) = "\", LenB(Dir(z & Datei)) = 0, True) Then
  FSO.CopyFile q, z, True
  If LenB(Dir(z)) = 0 And IIf(Right$(z, 1) = "\", LenB(Dir(z & Datei)) = 0, True) Then
   If iWV = 0 Then If WV = 0 Then iWV = GetOSVersion Else iWV = WV
   If iWV >= win_vista Then
   On Error GoTo fehler
    If Not pruefe Then
     Call adminaktiv
     pruefe = True
    End If
'    FileCopy q, Environ("userprofile") & "\" & Datei
    rufauf "cmd", "/c copy """ & Environ("userprofile") & "\" & Datei & """ """ & z & """", 2, , , 0
    If LenB(Dir(Environ("userprofile") & "\" & Datei)) <> 0 Then
'    Shell ("\\linux1\daten\down\pstools\psexec -u administrator -p sonne cmd /e:on /c move """ & Environ("userprofile") & "\" & Datei & """ """ & z & """")
     rufauf "cmd", "/c move """ & Environ("userprofile") & "\" & Datei & """ """ & z & """", 2, , , 0
    End If
   End If
  End If
  If LenB(Dir(z)) = 0 Then
   Debug.Print "Fehler beim Erstellen von: " & z
  End If
 End If
 Exit Sub
fehler:
 Select Case MsgBox("Fpos: " & FPos & " ErrNr: " & CStr(Err.Number) & vbCrLf & "LastDLLError: " & CStr(Err.LastDllError) & vbCrLf & "Source: " & IIf(IsNull(Err.Source), vbNullString, CStr(Err.Source)) & vbCrLf & "Description: " & Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in KopDat/" & App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' KopDat

