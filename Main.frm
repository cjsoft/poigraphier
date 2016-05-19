VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form MainFrm 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CJSoft PoiGraphier"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7575
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   7575
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox t3 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   270
      Left            =   3600
      TabIndex        =   14
      Text            =   "1"
      Top             =   6840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox t2 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   270
      Left            =   3600
      TabIndex        =   13
      Text            =   "40"
      Top             =   6600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox t1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   270
      Left            =   3600
      TabIndex        =   12
      Text            =   "-40"
      Top             =   6360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdClear 
      Appearance      =   0  'Flat
      Caption         =   "清除图像/更新坐标"
      Height          =   735
      Left            =   4920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6360
      Width           =   2295
   End
   Begin VB.Frame Frame5 
      Caption         =   "坐标设置"
      Height          =   1095
      Left            =   360
      TabIndex        =   21
      Top             =   5160
      Width           =   6855
      Begin VB.TextBox Oy 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   4800
         TabIndex        =   11
         Text            =   "15"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Ox 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   3720
         TabIndex        =   10
         Text            =   "20"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Ky 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   1080
         TabIndex        =   9
         Text            =   "200"
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox Kx 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   1080
         TabIndex        =   8
         Text            =   "200"
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "，"
         Height          =   180
         Left            =   4560
         TabIndex        =   27
         Top             =   720
         Width           =   180
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "坐标比例下的零点位置"
         Height          =   180
         Left            =   3720
         TabIndex        =   26
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "y坐标比例"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   810
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "x坐标比例"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   810
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "样式设置"
      Height          =   2055
      Left            =   360
      TabIndex        =   17
      Top             =   3000
      Width           =   6855
      Begin VB.CheckBox ShowMark 
         Appearance      =   0  'Flat
         Caption         =   "显示标记值"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox ScaleMark 
         Appearance      =   0  'Flat
         Caption         =   "坐标标记"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.TextBox DeltaY 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   4800
         TabIndex        =   7
         Text            =   "2"
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox DeltaX 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   4800
         TabIndex        =   6
         Text            =   "2"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CommandButton ColorSelector 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Height          =   735
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
      Begin VB.ComboBox SetWidth 
         Appearance      =   0  'Flat
         Height          =   300
         ItemData        =   "Main.frx":030A
         Left            =   720
         List            =   "Main.frx":032C
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Frame Frame4 
         Caption         =   "Sample"
         Height          =   855
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   2295
         Begin VB.Line LSamp 
            X1              =   360
            X2              =   1920
            Y1              =   480
            Y2              =   480
         End
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "比例下y坐标标记间隔"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3000
         TabIndex        =   25
         Top             =   1560
         Width           =   1710
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "比例下x坐标标记间隔"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3000
         TabIndex        =   24
         Top             =   1200
         Width           =   1710
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "粗细"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   20
         Top             =   1200
         Width           =   360
      End
      Begin VB.Label lbl1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "点击左侧按钮以更换颜色"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4320
         TabIndex        =   19
         Top             =   960
         Width           =   1980
      End
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "添加函数线"
      Height          =   735
      Left            =   360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6360
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "公式设置"
      Height          =   2655
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   6855
      Begin RichTextLib.RichTextBox CodeRes 
         Height          =   1935
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Visible         =   0   'False
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   3413
         _Version        =   393217
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"Main.frx":034F
      End
      Begin VB.CheckBox EMode 
         Appearance      =   0  'Flat
         Caption         =   "简单模式"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Value           =   1  'Checked
         Width           =   6255
      End
      Begin VB.TextBox FX 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   32
         Text            =   "x"
         Top             =   600
         Width           =   5055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "f(x)="
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   36
         Top             =   600
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nico Poi Duang Nico Poi Duang"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   360
         TabIndex        =   35
         Top             =   1320
         Width           =   4350
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "More Poi More Health"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   34
         Top             =   1680
         Width           =   3000
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "不要吐槽、小心我把你poi掉= =、"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   33
         Top             =   2160
         Width           =   4320
      End
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "步长"
      Height          =   180
      Left            =   3000
      TabIndex        =   30
      Top             =   6840
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "终止X"
      Height          =   180
      Left            =   3000
      TabIndex        =   29
      Top             =   6600
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "起始X"
      Height          =   180
      Left            =   3000
      TabIndex        =   28
      Top             =   6360
      Visible         =   0   'False
      Width           =   450
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const RN As String = vbCrLf
Dim FS As Long


Public Sub cmdClear_Click()
    frmHDC.Cls
    DrawScale
End Sub

Private Sub cmdOK_Click()
    On Error Resume Next
    Dim ErrD As String
    Dim poipoipoi() As Byte
    poipoipoi = LoadResData(101, "CUSTOM")
    Dim Calc As New ScriptRT
    Calc.SetScriptLanguage "VBS"
    If EMode.value = 1 Then
        ErrD = Calc.CheckSyntax("Function Main(x)" & RN & "    Main=" & FX.Text & RN & "End Function")
        If ErrD <> "" Then
            
            
            FS = BASS_StreamCreateFile(BASSTRUE, VarPtr(poipoipoi(0)), 0, UBound(poipoipoi), 0)
            BASS_ChannelPlay FS, BASSTRUE
            MsgBox "错误的公式poi、", vbCritical, "语法错误"
            Exit Sub
        End If
        Calc.AddCode "Function Main(x)" & RN & "    Main=" & FX.Text & RN & "End Function"
    Else
        ErrD = Calc.CheckSyntax(CodeRes.Text)
        If ErrD <> "" Then
            FS = BASS_StreamCreateFile(BASSTRUE, VarPtr(poipoipoi(0)), 0, UBound(poipoipoi), 0)
            BASS_ChannelPlay FS, BASSTRUE
            MsgBox ErrD, vbCritical, "语法错误"
            Exit Sub
        End If
        Calc.AddCode CodeRes.Text
    End If
    frmHDC.Show
    'Do While frmHDC.ReadyState <> 1
    'Loop
    DrawScale
    Dim i As Long
    Dim Yy As Long
    Yy = CLng(Ky.Text) * CLng(Oy.Text)
    Dim Xx As Long
    Xx = CLng(Kx.Text) * CLng(Ox.Text)
    Dim Dx As Long
    Dx = CLng(DeltaX.Text)
    Dim Dy As Long
    Dy = CLng(DeltaY.Text)
    Dim KkX As Long
    KkX = CLng(Kx.Text)
    Dim Kky As Long
    Kky = CLng(Ky.Text)
    Oxx = Xx: Oyy = Yy
    Debug.Print Oxx, Oyy
    frmHDC.ForeColor = ColorSelector.BackColor: frmHDC.DrawWidth = CInt(IIf(SetWidth.Text = "", "1", SetWidth.Text))
    Dim LastX As Double, LastY As Double, CurY As Double, Sx As Long, Ex As Long, Dd As Long
    Sx = CLng(t1.Text): Ex = CLng(t2.Text): Dd = CLng(t3.Text)
    Dim Temps As String
    'LastX = Sx * KkX: LastY = Calc.Run_LL(Sx) * Kky:
    LastX = -Xx - 100: LastY = Calc.Run_LL((-Xx - 100) / KkX) * Kky:
    Dim T As Double: T = Val(Dd)
    'For i = Sx * KkX + 1 To Ex * KkX Step KkX / 100
    For i = -Xx To frmHDC.ScaleWidth - Xx Step IIf(frmHDC.ScaleWidth / 2000 < 1, 100, frmHDC.ScaleWidth / 2000)
        CurY = Calc.Run_DD(i / KkX) * Kky
        'Debug.Print RescaleY(LastY)
        frmHDC.Line (RescaleX(CLng(LastX)), RescaleY(CLng(LastY)))-(RescaleX(i), RescaleY(CLng(CurY)))
        LastX = i: LastY = CurY
    Next
    Kx_LostFocus
    Ky_LostFocus
    Ox_LostFocus
    Oy_LostFocus
End Sub



Private Sub CodeRes_GotFocus()
    Dim C As Object
    On Error Resume Next
    For Each C In Me.Controls
        C.TabStop = False
    Next
End Sub

Private Sub CodeRes_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRT
    Dim i As Long, CNT As Long
    CNT = 0
    If KeyAscii = 9 Then
        KeyAscii = 0
        CodeRes.SelText = Space(4)
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        For i = InStrRev(CodeRes.Text, vbCrLf, CodeRes.SelStart) + 2 To CodeRes.SelStart
            If (Mid(CodeRes.Text, i, 1) = " ") Then
                CNT = CNT + 1
            Else
                Exit For
            End If
        Next
        CodeRes.SelText = vbCrLf & Space(Int(CNT / 4) * 4)
        hll CodeRes.SelStart
    End If
    Exit Sub
ERRT:
    If Err.Number = 5 Then CodeRes.SelText = vbCrLf
End Sub

Private Sub CodeRes_LostFocus()
    Dim C As Object
    On Error Resume Next
    For Each C In Me.Controls
        C.TabStop = True
    Next
End Sub

Private Sub ColorSelector_Click()
    Dim a As Integer, b As Integer, C As Integer
    Randomize
    a = CInt(256 * Rnd)
    Randomize
    b = CInt(256 * Rnd)
    Randomize
    C = CInt(256 * Rnd)
    ColorSelector.BackColor = RGB(a, b, C)
    LSamp.BorderColor = ColorSelector.BackColor
End Sub




Private Sub DeltaX_GotFocus()
    With DeltaX
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub DeltaX_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then DeltaY.SetFocus
End Sub


Private Sub DeltaX_LostFocus()
    If Trim(DeltaX.Text) = "" Then DeltaX.Text = "2"
End Sub

Private Sub DeltaY_GotFocus()
With DeltaY
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub DeltaY_LostFocus()
    If Trim(DeltaY.Text) = "" Then DeltaY.Text = "2"
End Sub
Private Sub DeltaY_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Kx.SetFocus
End Sub

Private Sub EMode_Click()
    If (EMode.value = 0) Then
        CodeRes.Visible = True
        CodeRes.Enabled = True
        FX.Enabled = False
    Else
        CodeRes.Visible = False
        CodeRes.Enabled = False
        FX.Enabled = True
    End If
End Sub

Private Sub EMode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdOK_Click
End Sub



Private Sub Form_DblClick()
    frmHDC.Form_Load
End Sub

Private Sub Form_Load()
    CodeRes = "Function Main(x)" & RN & "    '请在此书写函数代码（VBS），Main函数返回结果" & RN & RN & "End Function"
    Me.DrawWidth = 10
    Me.FillColor = vbBlack
    hll CodeRes.SelStart
    Dim SKHDLL() As Byte, SKINSHE() As Byte, BASSDLL() As Byte
    If Dir("SkinH_VB6.dll") = "" Then
        SKHDLL = LoadResData(103, "CUSTOM")
        Open "SkinH_VB6.dll" For Binary As #1
            Put #1, , SKHDLL
        Close #1
    End If
    If Dir("skin.she") = "" Then
        SKINSHE = LoadResData(102, "CUSTOM")
        Open "skin.she" For Binary As #1
            Put #1, , SKINSHE
        Close #1
    End If
    If Dir("bass.dll") = "" Then
        BASSDLL = LoadResData(104, "CUSTOM")
        Open "bass.dll" For Binary As #1
            Put #1, , BASSDLL
        Close #1
    End If
    
    SkinH_AttachEx "skin.she", ""
    If HiWord(BASS_GetVersion) <> BASSVERSION Then
    MsgBox "BASS不匹配"
     Unload Me
    End If
    If BASS_Init(-1, 44100, 0, Me.hWnd, 0) = BASSFALSE Then
        MsgBox "初始化失败"
        Unload Me
    End If
    Dim poipoipoi() As Byte
    
    poipoipoi = LoadResData(101, "CUSTOM")
    FS = BASS_StreamCreateFile(BASSTRUE, VarPtr(poipoipoi(0)), 0, UBound(poipoipoi), 0)
    BASS_ChannelPlay FS, BASSTRUE
    'SkinH_SetAero 1
    If Me.WindowState <> 1 And frmHDC.WindowState <> 1 Then
        Me.Move Screen.Width - (Me.Width + frmHDC.Width) / 2
    End If
    '====ReadSettings====
    DeltaX.Text = GetSetting("PoiGraphier", "Style", "DeltaX", "2")
    DeltaY.Text = GetSetting("PoiGraphier", "Style", "DeltaY", "2")
    Kx.Text = GetSetting("PoiGraphier", "Scale", "KX", "200")
    Ky.Text = GetSetting("PoiGraphier", "Scale", "KY", "200")
    Ox.Text = GetSetting("PoiGraphier", "Scale", "OX", "20")
    Oy.Text = GetSetting("PoiGraphier", "Scale", "OY", "15")
End Sub

Public Sub DrawScale()
    On Error Resume Next
    Const HLMark As Integer = 70
    Dim tempa As Integer, tempb
    Dim i As Long, Temps As String
    Dim Yy As Long
    Yy = CLng(Ky.Text) * CLng(Oy.Text)
    Dim Xx As Long
    Xx = CLng(Kx.Text) * CLng(Ox.Text)
    Dim Dx As Long
    Dx = CLng(DeltaX.Text)
    Dim Dy As Long
    Dy = CLng(DeltaY.Text)
    Dim KkX As Long
    KkX = CLng(Kx.Text)
    Dim Kky As Long
    Kky = CLng(Ky.Text)
    With frmHDC
        tempa = .DrawWidth
        tempb = .ForeColor
        .DrawWidth = 1
        .ForeColor = vbBlack
        If ScaleMark.value = 1 Then
            frmHDC.Line (0, Yy)-(.ScaleWidth, Yy)
            frmHDC.Line (Xx, 0)-(Xx, .ScaleHeight)
            For i = -Xx / Dx / KkX To (.ScaleWidth - Xx) / Dx / KkX
                frmHDC.Line (i * Dx * KkX + Xx, Yy - HLMark)-(i * Dx * KkX + Xx, Yy + HLMark)
            Next
            For i = -Yy / Dy / Kky To (.ScaleHeight - Yy) / Dy / Kky
                frmHDC.Line (Xx - HLMark, i * Dy * Kky + Yy)-(Xx + HLMark, i * Dy * Kky + Yy)
            Next
            If ShowMark.value = 1 Then
                If Dx = 0 Then GoTo JMPX
                For i = -Xx / Dx / KkX To (.ScaleWidth - Xx) / Dx / KkX
                    Temps = i * Dx
                    Temps = Trim(Temps)
                    If i = 0 Then Temps = ""
                    .CurrentX = i * Dx * KkX + Xx - .TextWidth(Temps) / 2
                    .CurrentY = Yy + HLMark
                    frmHDC.Print Temps
                Next
JMPX:
                If Dy = 0 Then GoTo JMPY
                For i = -Yy / Dy / Kky To (.ScaleHeight - Yy) / Dy / Kky
                    Temps = -i * Dy
                    Temps = Trim(Temps)
                    If i = 0 Then Temps = ""
                    frmHDC.Line (Xx - HLMark, i * Dy * Kky + Yy)-(Xx + HLMark, i * Dy * Kky + Yy)
                    .CurrentX = Xx - HLMark - .TextWidth(Temps)
                    .CurrentY = i * Dy * Kky + Yy - .TextHeight(Temps) / 2
                    frmHDC.Print Temps
                Next
JMPY:
            End If
        End If
    End With
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveSetting "PoiGraphier", "Style", "DeltaX", DeltaX.Text
    SaveSetting "PoiGraphier", "Style", "DeltaY", DeltaY.Text
    SaveSetting "PoiGraphier", "Scale", "KX", Kx.Text
    SaveSetting "PoiGraphier", "Scale", "KY", Ky.Text
    SaveSetting "PoiGraphier", "Scale", "OX", Ox.Text
    SaveSetting "PoiGraphier", "Scale", "OY", Oy.Text
    BASS_Free
End Sub

Private Sub FX_GotFocus()
    With FX
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub FX_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdOK_Click
End Sub

Public Sub Kx_Change()

End Sub

Private Sub Kx_GotFocus()
    With Kx
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Kx_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Ky.SetFocus
End Sub

Public Sub Kx_LostFocus()
    On Error GoTo ET
    
    Gkx = CLng(Kx.Text)
    Gxx = Gkx * CLng(Ox.Text)
    Exit Sub
ET:
    Kx.Text = "200"
    Gkx = CLng(Kx.Text)
    Gxx = Gkx * CLng(Ox.Text)
End Sub

Public Sub Ky_Change()

End Sub

Private Sub Ky_GotFocus()
    With Ky
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Ky_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Ox.SetFocus
End Sub

Public Sub Ky_LostFocus()
    On Error GoTo ET
    
    Gky = CLng(Ky.Text)
    Gyy = Gky * CLng(Oy.Text)
    Exit Sub
ET:
    Ky.Text = "200"
    Gky = CLng(Ky.Text)
    Gyy = Gky * CLng(Oy.Text)
End Sub



Private Sub Ox_GotFocus()
    With Ox
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Ox_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Oy.SetFocus
    
End Sub

Public Sub Ox_LostFocus()
    On Error GoTo ET
    Kx_LostFocus
    Gxx = CLng(Ox.Text) * Gkx
    Exit Sub
ET:
    Ox.Text = "20"
    Kx_LostFocus
    Gxx = CLng(Ox.Text) * Gkx
End Sub

Private Sub Oy_GotFocus()
    With Oy
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Oy_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdOK_Click
End Sub

Public Sub Oy_LostFocus()
    On Error GoTo ET
    Ky_LostFocus
    Gyy = CLng(Oy.Text) * Gky
    Exit Sub
ET:
    Oy.Text = "15"
    Ky_LostFocus
    Gyy = CLng(Oy.Text) * Gky
End Sub

Private Sub ScaleMark_Click()
    If ScaleMark.value = 1 Then
        ShowMark.Enabled = True
        DeltaX.Enabled = True
        DeltaY.Enabled = True
    Else
        ShowMark.Enabled = False
        DeltaX.Enabled = False
        DeltaY.Enabled = False
    End If
End Sub

Private Sub ScaleMark_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then cmdOK_Click
End Sub

Private Sub SetWidth_Click()
    LSamp.BorderWidth = CInt(IIf(SetWidth.Text = "", "1", SetWidth.Text))
End Sub

Private Sub SetWidth_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then cmdOK_Click
End Sub

Private Sub ShowMark_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdOK_Click
End Sub

Private Sub t1_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then t2.SetFocus
End Sub

Private Sub t2_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then t3.SetFocus
End Sub

Private Sub t3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdOK_Click
End Sub

Private Sub HighLight(ByVal find As String, ByVal NewColor As Long, ByVal OldSStart As Long)
    Dim i As Long
    i = 1
    Do While InStr(i, CodeRes.Text, find, vbTextCompare) <> 0
        With CodeRes
            .SelStart = InStr(i, CodeRes.Text, find, vbTextCompare) - 1
            .SelLength = Len(find)
            .SelColor = NewColor
            i = .SelStart + 1 + Len(find)
        End With
    Loop
    CodeRes.SelStart = OldSStart
    CodeRes.SelLength = 0
End Sub


Private Sub hll(oss As Long)
    Dim DarkYellow As Long
    Dim Pink As Long
    Pink = RGB(211, 54, 105)
    DarkYellow = RGB(113, 153, 0)
    SHL CodeRes.SelStart
End Sub


Private Sub SHL(oss As Long)
    Dim DarkYellow As Long
    Dim Pink As Long
    Pink = RGB(211, 54, 105)
    DarkYellow = RGB(113, 153, 0)
    CodeRes.SelStart = 0
    CodeRes.SelLength = Len(CodeRes.Text)
    CodeRes.SelColor = vbBlack
    Dim temp As String, T As String, tag As Boolean
    Dim C As New Collection, D As New Collection
    Dim i As Long, j
    tag = False
    With C
        .Add "function": .Add "sub"
        .Add "end": .Add "exit"
    End With
    With D
        .Add "dim": .Add "as": .Add "nico": .Add "poi": .Add "duang"
         .Add "if": .Add "then": .Add "do": .Add "while": .Add "loop": .Add "nothing"
        .Add "next": .Add "for": .Add "with": .Add "is": .Add "each": .Add "new"
    End With
    For i = 1 To Len(CodeRes.Text)
        T = Mid(CodeRes.Text, i, 1)
        If ((Asc(T) >= Asc("a") And Asc(T) <= Asc("z")) Or ((Asc(T) >= Asc("A") And Asc(T) <= Asc("Z")))) And Not tag Then
            temp = temp & T
        ElseIf T = """" Or tag Then
            If tag And T = """" Then
                CodeRes.SelLength = CodeRes.SelLength + 1
                CodeRes.SelColor = vbRed
                tag = False
            Else
                If T = """" Then
                    CodeRes.SelStart = i - 1
                    CodeRes.SelLength = 1
                End If
                CodeRes.SelLength = CodeRes.SelLength + 1
                tag = True
            End If
        Else
            For Each j In C
                If j = Format(temp, "<") Then
                    CodeRes.SelStart = i - Len(temp) - 1
                    CodeRes.SelLength = Len(temp)
                    CodeRes.SelColor = vbBlue
                    Exit For
                End If
            Next j
            For Each j In D
                If j = Format(temp, "<") Then
                    CodeRes.SelStart = i - Len(temp) - 1
                    CodeRes.SelLength = Len(temp)
                    CodeRes.SelColor = DarkYellow
                    Exit For
                End If
            Next j
            temp = ""
        End If
    Next i
    For Each j In C
            If j = Format(temp, "<") Then
                 CodeRes.SelStart = Len(CodeRes.Text) - Len(temp)
                 CodeRes.SelLength = Len(temp)
                 CodeRes.SelColor = vbBlue
                temp = ""
            Exit For
         End If
    Next j
    For Each j In C
            If j = Format(temp, "<") Then
                 CodeRes.SelStart = Len(CodeRes.Text) - Len(temp)
                 CodeRes.SelLength = Len(temp)
                 CodeRes.SelColor = DarkYellow
                temp = ""
            Exit For
         End If
    Next j
    CodeRes.SelStart = oss
    CodeRes.SelLength = 0
End Sub

Private Sub Timer1_Timer()
    frmHDC.Show
    If Me.WindowState <> 1 And frmHDC.WindowState <> 1 Then
        Me.Move Screen.Width / 2 - (Me.Width + frmHDC.Width) / 2
        Form_DblClick
    End If
    Timer1.Enabled = False
End Sub
