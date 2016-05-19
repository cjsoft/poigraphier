VERSION 5.00
Begin VB.Form frmHDC 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "绘图窗口"
   ClientHeight    =   5430
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   ScaleHeight     =   5430
   ScaleMode       =   0  'User
   ScaleWidth      =   8000
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox PSave 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4200
      ScaleHeight     =   855
      ScaleWidth      =   975
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Line LineY 
      BorderStyle     =   3  'Dot
      X1              =   -100.132
      X2              =   -100.132
      Y1              =   360
      Y2              =   2400
   End
   Begin VB.Line LineX 
      BorderStyle     =   3  'Dot
      X1              =   -3288.538
      X2              =   -100.132
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Menu cmdSave 
      Caption         =   "保存"
   End
   Begin VB.Menu AidScaleEnabled 
      Caption         =   "参考线:开"
   End
   Begin VB.Menu CurPos 
      Caption         =   "X: Y:"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "frmHDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ReadyState As Integer

Private Sub AidScaleEnabled_Click()
    If AidScaleEnabled.Caption = "参考线:开" Then
        LineX.Visible = False
        LineY.Visible = False
        AidScaleEnabled.Caption = "参考线:关"
    Else
        LineX.Visible = True
        LineY.Visible = True
        AidScaleEnabled.Caption = "参考线:开"
    End If
End Sub

Private Sub cmdSave_Click()
    Dim T As String
    T = Format(Now, "yyyy_mm_dd h_m_s")
    If Dir("Saves", vbDirectory) = "" Then MkDir "Saves"
    PSave.Cls
    PSave.Width = Me.ScaleWidth
    PSave.Height = Me.ScaleHeight
    PSave.PaintPicture Me.Image, 0, 0
    SavePicture PSave.Image, Replace(App.Path & "\Saves\" & T & ".bmp", "\\", "\")
    
    Shell "explorer.exe /select,""" & Replace(App.Path & "\Saves\" & T & ".bmp", "\\", "\") & """", vbNormalFocus
End Sub

Private Sub Form_DblClick()
    If MainFrm.WindowState <> 1 Then
        MainFrm.Move Me.left - MainFrm.Width, Me.top
    End If
End Sub

Public Sub Form_Load()
    ReadyState = 0
    MainFrm.Kx_LostFocus
    MainFrm.Ky_LostFocus
    MainFrm.Ox_LostFocus
    MainFrm.Oy_LostFocus
    If Me.WindowState <> 1 Then
        Me.left = MainFrm.left + MainFrm.Width
        Me.top = MainFrm.top
    End If
    MainFrm.cmdClear_Click
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With LineX
        .X1 = 0: .Y1 = Y
        .X2 = Me.ScaleWidth: .Y2 = Y
        .Refresh
    End With
    With LineY
        .X1 = X: .Y1 = 0
        .X2 = X: .Y2 = Me.ScaleWidth
        .Refresh
    End With
    Debug.Print RescaleY((Gyy - Y))
    CurPos.Caption = "X:" & CDbl(X - Gxx) / CDbl(Gkx) & " Y:" & CDbl(Gyy - Y) / CDbl(Gky)
End Sub

Private Sub Form_Paint()
    ReadyState = 1
End Sub
