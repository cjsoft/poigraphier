VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScriptRT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public SC As Object


Private Sub Class_Initialize()
    Set SC = New MSScriptControl.ScriptControl
End Sub
Public Function CheckSyntax(ResCode As String) As String
    On Error GoTo ERRT
    CheckSyntax = ""
    Dim a As New ScriptRT
    a.SetScriptLanguage "VBS"
    a.SC.ExecuteStatement ResCode
    Exit Function
ERRT:
    CheckSyntax = a.SC.Error.Number & " ' " & a.SC.Error.Description & vbCrLf & "错误出现在第" & a.SC.Error.Line & "poi行、" & vbCrLf & vbCrLf & "  再出错就poi掉你= =、"
End Function

Public Sub SetScriptLanguage(Script_Name As String)
    SC.Language = Script_Name
End Sub

Public Sub AddCode(Function_ As String)
    Debug.Print Function_
    SC.AddCode Function_
End Sub

Public Function Run_LL(x_ As Long) As Long
    On Error Resume Next
    Run_LL = SC.Run("Main", x_)
End Function

Public Function Run_LD(x_ As Long) As Double
    On Error Resume Next
    Run_LD = SC.Run("Main", x_)
End Function

Public Function Run_DL(x_ As Double) As Long
    On Error Resume Next
    Run_DL = SC.Run("Main", x_)
End Function

Public Function Run_DD(x_ As Double) As Double
    On Error Resume Next
    Run_DD = SC.Run("Main", x_)
End Function


