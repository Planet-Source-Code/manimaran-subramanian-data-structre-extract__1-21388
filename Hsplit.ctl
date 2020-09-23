VERSION 5.00
Begin VB.UserControl HSplit 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4995
   ControlContainer=   -1  'True
   KeyPreview      =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4995
   ToolboxBitmap   =   "Hsplit.ctx":0000
   Begin VB.PictureBox PicHold 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3405
      Left            =   15
      ScaleHeight     =   3405
      ScaleWidth      =   4920
      TabIndex        =   0
      Top             =   60
      Width           =   4920
      Begin VB.PictureBox HPicSplit 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   45
         Left            =   435
         MousePointer    =   7  'Size N S
         ScaleHeight     =   45
         ScaleWidth      =   4065
         TabIndex        =   1
         Top             =   2595
         Width           =   4065
      End
   End
End
Attribute VB_Name = "HSplit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim moHSplitter As CHorizontalSplitter
Event BarResized(TTop As Single, THeight As Single, BTop As Single, BHeight As Single)
Private Sub HPicSplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    moHSplitter.MouseDown
    RaiseEvent BarResized(moHSplitter.TopChild.Top, moHSplitter.TopChild.Height, moHSplitter.BottomChild.Top, moHSplitter.BottomChild.Height)
End Sub
Private Sub HPicSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    moHSplitter.MouseMove Y
    RaiseEvent BarResized(moHSplitter.TopChild.Top, moHSplitter.TopChild.Height, moHSplitter.BottomChild.Top, moHSplitter.BottomChild.Height)
End Sub
Private Sub HPicSplit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    moHSplitter.MouseUp
    RaiseEvent BarResized(moHSplitter.TopChild.Top, moHSplitter.TopChild.Height, moHSplitter.BottomChild.Top, moHSplitter.BottomChild.Height)
End Sub
Private Sub UserControl_Initialize()
    PicHold.Left = 0
    PicHold.Top = 0
    HPicSplit.Enabled = False
End Sub
Private Sub UserControl_Resize()
    PicHold.Left = 0
    PicHold.Top = 0
    PicHold.Width = UserControl.Width
    PicHold.Height = UserControl.Height
End Sub
Public Sub ActivateHSControl(TopCtl As Variant, BotCtl As Variant)
    HPicSplit.Enabled = True
    HPicSplit.Top = PicHold.Height / 2
   Set moHSplitter = New CHorizontalSplitter
   moHSplitter.Init PicHold, HPicSplit, TopCtl, BotCtl
End Sub
Public Sub RefreshControl()
    On Error Resume Next
    moHSplitter.Refresh
End Sub
Private Sub UserControl_Terminate()
   Set moHSplitter = Nothing
End Sub
