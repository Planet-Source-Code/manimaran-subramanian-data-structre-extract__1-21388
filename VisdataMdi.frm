VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm VisdataMdi 
   BackColor       =   &H8000000C&
   Caption         =   "Visdata (Trial)"
   ClientHeight    =   4320
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   3345
   Icon            =   "VisdataMdi.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog Cmdlg 
      Left            =   1830
      Top             =   1740
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Menu MNUFILE 
      Caption         =   "&File"
      Begin VB.Menu mnuopn 
         Caption         =   "&Open Database"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuprint 
         Caption         =   "Print Structure"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuwith 
         Caption         =   "Import Table With Data"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu MNUFILELIST 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu MNUFILELIST 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu MNUFILELIST 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu MNUFILELIST 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu MNUSEP2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuext 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuwin 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnucas 
         Caption         =   "Cascade"
      End
      Begin VB.Menu mnutil 
         Caption         =   "Tile Horizantal"
      End
      Begin VB.Menu mnutilver 
         Caption         =   "Tile Vertical"
      End
      Begin VB.Menu mnuico 
         Caption         =   "Iconic"
      End
   End
End
Attribute VB_Name = "VisdataMdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CmdlgFileName As String
Public SrvStr As String
Private Sub MDIForm_Load()
    Dim I As Integer
    I = Val(GetSetting("MYVISDATA", "OPTION", "CHECK", "1"))
    If I = 0 Then
        mnuwith.Checked = False
    Else
        mnuwith.Checked = True
    End If
End Sub
Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu MNUFILE, , X, Y
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
    If mnuwith.Checked Then
        SaveSetting "MYVISDATA", "OPTION", "CHECK", 1
    Else
        SaveSetting "MYVISDATA", "OPTION", "CHECK", 0
    End If
End Sub
Private Sub mnucas_Click()
    Me.Arrange 0
End Sub
Private Sub mnuext_Click()
    Unload Me
End Sub
Private Sub MNUFILELIST_Click(Index As Integer)
'    Dim FRM As New TreeForm
'    If UCase(Right(MNUFILELIST(Index).Caption, 4)) = "NONE" Then Exit Sub
'    VisdataMdi.Cmdlg.FileName = Right(MNUFILELIST(Index).Caption, Len(MNUFILELIST(Index).Caption) - 3)
'    Load FRM
'    FRM.Show
End Sub
Private Sub mnuico_Click()
    Me.Arrange 3
End Sub
Private Sub mnuopn_Click()
'    Cmdlg.Filter = "Access Database(*.mdb)|*.mdb"
'    Dim DD As rdoConnection
    On Error GoTo aa:
'    Cmdlg.ShowOpen
'    Set DD = OpenDatabase(Cmdlg.FileName)
'    DD.Close
    CmdlgFileName = ""
    SrvStr = ""
    frmtabs.Show 1
    If Len(CmdlgFileName) = 0 Then Exit Sub
    Dim Temp As New TreeForm
    Load Temp
    Temp.Show
aa:
End Sub
Private Sub mnuprint_Click()
    PrintTabStructure Screen.ActiveForm
End Sub
Private Sub mnutil_Click()
    Me.Arrange 1
End Sub
Private Sub mnutilver_Click()
    Me.Arrange 2
End Sub
Public Sub PrintTabStructure(FRM As Form)
    Dim I As Integer
    Dim PQ As Boolean
    Dim RES As Integer
    If FRM.Caption = Me.Caption Or Left(UCase(FRM.Caption), 6) = "RESULT" Then Exit Sub
    If FRM.TVW.Nodes.Count = 0 Then Exit Sub
    RES = MsgBox("DO YOU WANT PRINT QUERIES ?", vbYesNoCancel + vbQuestion, "VISDATA")
    If RES = vbYes Then
        PQ = True
    ElseIf RES = vbNo Then
        PQ = False
    Else
        Exit Sub
    End If
    
    FRM.rt.Text = FRM.Caption & vbCrLf
    frmprog.Show
    For I = 2 To FRM.TVW.Nodes.Count
        DoEvents
        If PQ Then
            Select Case FRM.TVW.Nodes(I).Image
                Case 4
                    If FRM.TVW.Nodes(I).Checked = True Then FRM.rt.Text = FRM.rt.Text & "</TABLE><BR><TABLE BORDER=1><TR><TD>" & UCase(FRM.TVW.Nodes(I).Text) & "</TD></TR>": frmprog.lblprog.Caption = FRM.TV.Nodes(I).Text
                Case 6
                    If FRM.TVW.Nodes(I).Checked = True Then FRM.rt.Text = FRM.rt.Text & "</TABLE><BR><TABLE BORDER=1><TR><TD> Query : " & FRM.TVW.Nodes(I).Text & "</TD></TR>": frmprog.lblprog.Caption = FRM.TV.Nodes(I).Text
                Case 5, 2, 1
                    If FRM.TVW.Nodes(I).Checked = True Then FRM.rt.Text = FRM.rt.Text & FRM.TVW.Nodes(I).Text: frmprog.lblprog.Caption = FRM.TV.Nodes(I).Text
            End Select
        Else
            Select Case FRM.TVW.Nodes(I).Image
                Case 4
                    If FRM.TVW.Nodes(I).Checked = True Then FRM.rt.Text = FRM.rt.Text & "</TABLE><BR><TABLE BORDER=1><TR><TD>" & UCase(FRM.TVW.Nodes(I).Text) & "</TD></TR>": frmprog.lblprog.Caption = FRM.TV.Nodes(I).Text
                Case 5, 2
                    If FRM.TVW.Nodes(I).Checked = True Then FRM.rt.Text = FRM.rt.Text & FRM.TVW.Nodes(I).Text & vbCrLf: frmprog.lblprog.Caption = FRM.TV.Nodes(I).Text
            End Select
        End If
    Next
    frmprog.lblprog.Caption = "Structure Generation Completed"
    FRM.rt.SaveFile FRM.Caption & ".HTM", 1
    frmprog.cmdclose.Enabled = True
End Sub
Private Sub mnuwith_Click()
    mnuwith.Checked = Not mnuwith.Checked
End Sub
