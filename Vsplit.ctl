VERSION 5.00
Begin VB.UserControl VSplit 
   ClientHeight    =   3450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4950
   ControlContainer=   -1  'True
   KeyPreview      =   -1  'True
   PropertyPages   =   "Vsplit.ctx":0000
   ScaleHeight     =   3450
   ScaleWidth      =   4950
   ToolboxBitmap   =   "Vsplit.ctx":000E
   Begin VB.PictureBox PicHold 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3405
      Left            =   -15
      ScaleHeight     =   3405
      ScaleWidth      =   4920
      TabIndex        =   0
      Top             =   0
      Width           =   4920
      Begin VB.PictureBox vPicSplit 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2925
         Left            =   1710
         MousePointer    =   9  'Size W E
         ScaleHeight     =   2925
         ScaleWidth      =   45
         TabIndex        =   1
         Top             =   150
         Width           =   45
      End
   End
End
Attribute VB_Name = "VSplit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim moVSplitter As CVerticalSplitter
Event BarResized(LLeft As Single, LWidth As Single, RLeft As Single, RWidth As Single)
Private Sub UserControl_Initialize()
    Pichold.Left = 0
    Pichold.Top = 0
    vPicSplit.Enabled = False
End Sub
Private Sub UserControl_Resize()
    Pichold.Left = 0
    Pichold.Top = 0
    Pichold.Width = UserControl.Width
    Pichold.Height = UserControl.Height
End Sub
Public Sub ActivateVSControl(LefCtl As Variant, RigCtl As Variant)
   Set moVSplitter = New CVerticalSplitter
   moVSplitter.Init Pichold, vPicSplit, LefCtl, RigCtl
   vPicSplit.Enabled = True
End Sub
Private Sub UserControl_Terminate()
   Set moVSplitter = Nothing
End Sub
Private Sub vPicSplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    moVSplitter.MouseDown
    RaiseEvent BarResized(moVSplitter.LeftChild.Left, moVSplitter.LeftChild.Width, moVSplitter.RightChild.Left, moVSplitter.RightChild.Width)
End Sub
Private Sub vPicSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    moVSplitter.MouseMove X
    RaiseEvent BarResized(moVSplitter.LeftChild.Left, moVSplitter.LeftChild.Width, moVSplitter.RightChild.Left, moVSplitter.RightChild.Width)
End Sub
Private Sub vPicSplit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    moVSplitter.MouseUp
    RaiseEvent BarResized(moVSplitter.LeftChild.Left, moVSplitter.LeftChild.Width, moVSplitter.RightChild.Left, moVSplitter.RightChild.Width)
End Sub
Public Sub RefreshControl()
    On Error Resume Next
    moVSplitter.Refresh
End Sub
