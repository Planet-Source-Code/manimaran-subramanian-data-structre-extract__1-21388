VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmtabs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open Database"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   Icon            =   "frmtabs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   645
      Top             =   1170
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtabs.frx":030A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LV 
      Height          =   2685
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   4736
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmtabs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim cn As New rdoConnection
    On Error GoTo Ennd:
    cn.Connect = "uid=sa;pwd=;database=master;driver={sql server}"
    cn.CursorDriver = rdUseOdbc
    cn.EstablishConnection , rdDriverPrompt
    VisdataMdi.SrvStr = cn.Connect
    Collecttabs cn
    Exit Sub
Ennd:
    MsgBox Err.Number & " : " & Err.Description, vbCritical, "VISDATA"
    Unload Me
End Sub
Private Sub Collecttabs(CNN As rdoConnection)
    Dim rs As rdoResultset
    Set rs = CNN.OpenResultset("select name from sysdatabases", 1, 3)
    While Not rs.EOF
        'If UCase(rs(0)) <> "MASTER" Then
        lv.ListItems.Add , , UCase(rs(0)), 1, 1
        rs.MoveNext
    Wend
End Sub
Private Sub LV_DblClick()
    If lv.SelectedItem Is Nothing Then Exit Sub
    VisdataMdi.CmdlgFileName = lv.SelectedItem.Text
    Unload Me
End Sub
Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then LV_DblClick
    If KeyCode = 27 Then Unload Me
End Sub
