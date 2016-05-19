VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PoiGraphier安装/反安装程序"
   ClientHeight    =   735
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3375
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   3375
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      Caption         =   "卸载"
      Height          =   735
      Left            =   1680
      TabIndex        =   1
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "安装"
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    On Error GoTo errort
    Dim RTB() As Byte
    If Dir("richtx32.ocx") = "" Then
        RTB = LoadResData(101, "CUSTOM")
        Open "richtx32.ocx" For Binary As #1
            Put #1, , RTB
        Close #1
        If Dir("C:\Windows\SysWOW64", vbDirectory) = "" Then
            Shell "regsvr32.exe """ & Replace(App.Path & "\RICHTX32.OCX""", "\\", "\"), vbNormalFocus
        Else
            Shell "C:\Windows\SysWOW64\regsvr32.exe """ & Replace(App.Path & "\RICHTX32.OCX""", "\\", "\"), vbNormalFocus
        End If
    End If
    Dim MEXE() As Byte
    If Dir("PoiGraphier.exe") = "" Then
        RTB = LoadResData(102, "CUSTOM")
        Open "PoiGraphier.exe" For Binary As #1
            Put #1, , RTB
        Close #1
    End If
    Command1.Caption = "安装成功"
    Command1.Enabled = False
    Shell "explorer.exe /select,""" & Replace(App.Path & "\PoiGraphier.exe""", "\\", "\"), vbNormalFocus
    Exit Sub
errort:
    MsgBox Err.Description
    
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    If Dir("PoiGraphier.exe") <> "" Then Kill ("PoiGraphier.exe")
    If Dir("skin.she") <> "" Then Kill ("skin.she")
    If Dir("bass.dll") <> "" Then Kill ("bass.dll")
    If Dir("SkinH_VB6.dll") <> "" Then Kill ("SkinH_VB6.dll")
    If Dir("RICHTX32.OCX") <> "" Then Kill ("RICHTX32.OCX")
    Command2.Caption = "卸载成功"
    Command2.Enabled = False
End Sub
