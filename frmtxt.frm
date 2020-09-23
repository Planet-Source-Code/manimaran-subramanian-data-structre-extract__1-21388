VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmtxt 
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5745
   Icon            =   "frmtxt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5085
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   Begin sqlproject.HSplit HSp 
      Height          =   4170
      Left            =   -90
      TabIndex        =   1
      Top             =   -15
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   7355
      Begin MSComctlLib.ListView lv 
         Height          =   975
         Left            =   0
         TabIndex        =   3
         Top             =   945
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   1720
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin RichTextLib.RichTextBox TXTSQL 
         Height          =   645
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   1138
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmtxt.frx":030A
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6990
      Top             =   660
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtxt.frx":03B8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar STB 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   4815
      Width           =   5745
      _ExtentX        =   10134
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2090
            MinWidth        =   776
            Picture         =   "frmtxt.frx":080C
            Text            =   "Previous"
            TextSave        =   "Previous"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   1588
            MinWidth        =   776
            Picture         =   "frmtxt.frx":0C60
            Text            =   "Next"
            TextSave        =   "Next"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmtxt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Expr As Boolean
Dim tLdb As rdoConnection
Dim Trst As rdoResultset
Dim Tnd As Node
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Expr = True
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    HSp.Left = 0
    HSp.Top = 0
    HSp.Width = Me.ScaleWidth
    HSp.Height = Me.ScaleHeight - STB.Height
    HSp.ActivateHSControl TXTSQL, lv
End Sub
Public Sub CollectLV(rsT As rdoResultset, LDB As rdoConnection, NDTXT As Node)
    Dim I As Integer
    Dim litem As ListItem
'    Dim Td As TableDef
    Dim TXT As String
    Dim pos As Integer
    lv.ListItems.Clear
    lv.ColumnHeaders.Clear
    Me.Caption = "RESULT OF " & NDTXT.Text
    Set Tnd = NDTXT
    Set Trst = rsT
    STB.Panels(3).Text = LDB.Name
    Set tLdb = LDB
    For I = 0 To rsT.rdoColumns.Count - 1
        lv.ColumnHeaders.Add , , UCase(rsT.rdoColumns(I).Name)
    Next
    If NDTXT.Children = 2 Then
        For I = 1 To lv.ColumnHeaders.Count
            pos = InStr(1, UCase(NDTXT.Child.Text), UCase(lv.ColumnHeaders(I).Text))
            If pos > 0 Then
                lv.ColumnHeaders(I).Icon = 1
            End If
        Next
    End If
    While Not rsT.EOF
        DoEvents
        If Expr Then
            If MsgBox("Are you sure to stop Collection?", vbYesNo + vbQuestion, "VISDATA") = vbYes Then
                Expr = False
                Exit Sub
            End If
            Expr = False
        End If
        TXT = IIf(IsNull(rsT(0)), "", rsT(0))
        Set litem = lv.ListItems.Add(, , TXT)
        For I = 1 To rsT.rdoColumns.Count - 1
            On Error GoTo NNd:
            If Not IsNull(rsT(I)) Then
                litem.SubItems(I) = rsT(I)
            Else
                litem.SubItems(I) = ""
            End If
                
'            TXT = IIf(IsNull(rsT(I)), "", rsT(I))
'            litem.SubItems(I) = TXT
        Next
        rsT.MoveNext
   Wend
    If lv.ListItems.Count > 0 Then lv.ListItems(1).Selected = True
    STB.Panels(2).Text = "NO OF FIELDS " & lv.ColumnHeaders.Count
    If lv.ListItems.Count > 0 Then
        STB.Panels(4).Text = "RECORDS " & lv.SelectedItem.Index & " OF " & lv.ListItems.Count
    Else
        STB.Panels(4).Text = "RECORDS 0 OF " & lv.ListItems.Count
    End If
    Exit Sub
NNd:
    Resume Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmtxt = Nothing
End Sub
Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    STB.Panels(4).Text = "RECORDS " & lv.SelectedItem.Index & " OF " & lv.ListItems.Count
End Sub
Private Sub STB_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If lv.ListItems.Count = 0 Then Exit Sub
    On Error Resume Next
    If Panel.Index = 1 Then
        lv.ListItems(lv.SelectedItem.Index - 1).Selected = True
    ElseIf Panel.Index = 5 Then
        lv.ListItems(lv.SelectedItem.Index + 1).Selected = True
    End If
    lv.ListItems(lv.SelectedItem.Index).EnsureVisible
End Sub
Public Sub TXTSQL_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim lrc As rdoResultset
    On Error GoTo Ennd:
    If Shift = 2 And KeyCode = vbKeyF5 Then
        If Len(Trim(TXTSQL.SelText)) = 0 Then
            Set lrc = tLdb.OpenResultset(TXTSQL.Text, 1, 3)
        Else
            Set lrc = tLdb.OpenResultset(TXTSQL.SelText, 1, 3)
        End If
        CollectLV lrc, tLdb, Tnd
    End If
    Exit Sub
Ennd:
    MsgBox Err.Number & " : " & Err.Description, vbCritical, "VISDATA"
End Sub

