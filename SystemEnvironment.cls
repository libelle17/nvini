VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "System"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mstrEnvironment As String
'Private Const REGISTRY_PATH = "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
Private Const REGISTRY_PATH = "SOFTWARE\GSProducts"
Private oRegistry As Registry

Private Declare Function SendMessageTimeout Lib "user32" _
    Alias "SendMessageTimeoutA" (ByVal hwnd As Long, _
    ByVal msg As Long, ByVal wParam As Long, _
    ByVal lParam As String, ByVal fuFlags As Long, _
    ByVal uTimeout As Long, lpdwResult As Long) As Long
    
Private Const HWND_BROADCAST As Long = &HFFFF&
Private Const SMTO_ABORTIFHUNG As Long = &H2
Private Const WM_SETTINGCHANGE As Long = &H1A

Private mValuePairs As cValuePairs
Private Sub EnumerateVariables()
    Dim sKeys() As String
    Dim iKeyCount As Long
    Dim x As Long
    Dim oValuePair As cValuePair
    
    Set mValuePairs = New cValuePairs
    
    oRegistry.EnumerateValues sKeys(), iKeyCount
    For x = 1 To iKeyCount
        oRegistry.ValueKey = sKeys(x)
        
        Set oValuePair = New cValuePair
        oValuePair.Variable = sKeys(x)
        oValuePair.Value = oRegistry.Value
        
        mValuePairs.Add oValuePair, oValuePair.Variable
        Set oValuePair = Nothing
    Next
    
End Sub

Public Property Get Environment(ByVal strName As String) As String
    oRegistry.ValueKey = strName
    Environment = oRegistry.Value
End Property

Public Property Let Environment(ByVal strName As String, ByVal strValue As String)
    oRegistry.ValueKey = strName
    oRegistry.ValueType = REG_SZ
    oRegistry.Value = strValue
    
'    BroadcastEnvironmentChange
End Property

Private Sub BroadcastEnvironmentChange()
    Dim lngReturnValue As Long

    SendMessageTimeout HWND_BROADCAST, WM_SETTINGCHANGE, 0&, _
        "Environment", SMTO_ABORTIFHUNG, 5000&, lngReturnValue
End Sub


Public Property Get List() As cValuePairs
    If mValuePairs Is Nothing Then
        EnumerateVariables
    End If
    
    Set List = mValuePairs
End Property

Private Sub Class_Initialize()
    Set oRegistry = New Registry
    oRegistry.ClassKey = HKEY_LOCAL_MACHINE
    oRegistry.SectionKey = REGISTRY_PATH
End Sub


Private Sub Class_Terminate()
    Set oRegistry = Nothing
    Set mValuePairs = Nothing
End Sub


