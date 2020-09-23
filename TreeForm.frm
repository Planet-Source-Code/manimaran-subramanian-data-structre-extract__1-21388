VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form TreeForm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4065
   Icon            =   "TreeForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.TreeView TVW 
      Height          =   2100
      Left            =   570
      TabIndex        =   2
      Top             =   3255
      Visible         =   0   'False
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   3704
      _Version        =   393217
      Style           =   7
      Checkboxes      =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1845
      Top             =   4155
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TreeForm.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TreeForm.frx":0896
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1845
      Top             =   3060
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TreeForm.frx":0BBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TreeForm.frx":0ED6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TreeForm.frx":11F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TreeForm.frx":150E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TreeForm.frx":182A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TreeForm.frx":1B46
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TreeForm.frx":2612
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TreeForm.frx":292E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TV 
      Height          =   2760
      Left            =   45
      TabIndex        =   1
      Top             =   105
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4868
      _Version        =   393217
      Indentation     =   176
      LabelEdit       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   0
      OLEDropMode     =   1
   End
   Begin RichTextLib.RichTextBox rt 
      Height          =   570
      Left            =   690
      TabIndex        =   0
      Top             =   3075
      Visible         =   0   'False
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1005
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"TreeForm.frx":2C4A
   End
End
Attribute VB_Name = "TreeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Db As New rdoConnection
Dim QUSTR As String
Public Sub Form_Load()
    'On Error GoTo ENND:
    Me.Caption = VisdataMdi.CmdlgFileName
    COLLECT_TREE_NODES
Ennd:
End Sub
Private Sub COLLECT_TREE_NODES()
'    Dim Td As TableDef
'    Dim Qd As QueryDef
'    Dim Pro As Property
    Dim rsT As rdoResultset
    Dim rsF As rdoResultset
    Dim RSi As rdoResultset
    Dim des As String
    Dim f1 As Integer
    Dim f2 As Integer
    Dim Fcn As String
    Dim Scn As String
    On Error GoTo Ennd:
    Screen.MousePointer = 11
    If Not Db Is Nothing Then Db.Close
    des = VisdataMdi.SrvStr ' "Server=SREENI;uid=sa;pwd=;database=" & Me.Caption & ";driver={sql server}"
    f1 = InStr(1, UCase(des), "DATABASE=MASTER")
    Fcn = Left(des, f1 - 1)
    Scn = Right(des, (Len(des) - f1 - 14))
    des = Fcn & "database=" & Me.Caption & Scn
    Db.Connect = des
    Db.CursorDriver = rdUseOdbc
    Db.EstablishConnection , rdDriverNoPrompt
    des = ""
    TV.Nodes.Clear
    TV.Nodes.Add , , Db.Name, Db.Name, 8, 8
    TV.Nodes(1).Expanded = True
    TVW.Nodes.Clear
    TVW.Nodes.Add , , Db.Name, Db.Name, 8, 8
    TVW.Nodes(1).Expanded = True
    Set rsT = Db.OpenResultset("select distinct name from sysobjects where upper(type)='U' and status>=0", 1, 3)
    While Not rsT.EOF
        DoEvents
            TV.Nodes.Add Db.Name, tvwChild, "Tab" & rsT(0), rsT(0), 4, 4
            TVW.Nodes.Add Db.Name, tvwChild, "Tab" & rsT(0), rsT(0), 4, 4
                des = "select convert(sysname,c.name),convert(sysname,i.name) From " & _
                 " sysindexes i, syscolumns c, sysobjects o, syscolumns c1 Where o.id = OBJECT_ID('" & rsT(0) & "') and o.id = c.id and o.id = i.id " & _
                " and (i.status & 0x800) = 0x800 and c.name = index_col ('" & rsT(0) & "', i.indid, c1.colid) and c1.colid <= i.keycnt and c1.id = OBJECT_ID('" & rsT(0) & "') order by 1 "
                Set RSi = Db.OpenResultset(des, 1, 3)
                des = ""
                While Not RSi.EOF
                    des = des & "+" & RSi(0) & ";"
                    RSi.MoveNext
                Wend
                If Len(des) > 0 Then
                    TV.Nodes.Add "Tab" & rsT(0), tvwChild, , des, 1
                    TVW.Nodes.Add "Tab" & rsT(0), tvwChild, , "<tr><td>" & des & "</td></tr>", 1
                End If
            Set rsF = Db.OpenResultset("select syscolumns.xtype, syscolumns.name,systypes.name,syscolumns.length,syscolumns.prec,syscolumns.cdefault,syscolumns.isnullable from syscolumns,systypes  where id = object_id('" & _
                rsT(0) & "') and syscolumns.xtype = systypes.xtype order by colid", 1, 3)
            TV.Nodes.Add "Tab" & rsT(0), tvwChild, "Tab" & rsT(0) & "Fie", "Fields", 3, 3
            TVW.Nodes.Add "Tab" & rsT(0), tvwChild, "Tab" & rsT(0) & "Fie", "Fields", 3, 3
            While Not rsF.EOF
                des = ""
                If UCase(rsF(2)) = "VARCHAR" Or UCase(rsF(2)) = "NVARCHAR" Or UCase(rsF(2)) = "CHAR" Then
                    TV.Nodes.Add "Tab" & rsT(0) & "Fie", tvwChild, , rsF(1) & " " & rsF(2) & " (" & rsF(3) & ")", 2
                    TVW.Nodes.Add "Tab" & rsT(0) & "Fie", tvwChild, , "<TR><TD>" & rsF(1) & "</TD><TD>" & rsF(2) & "</TD><TD>(" & rsF(3) & ")</TD></TR>", 2
                Else
                    TV.Nodes.Add "Tab" & rsT(0) & "Fie", tvwChild, , rsF(1) & " " & rsF(2), 2
                    TVW.Nodes.Add "Tab" & rsT(0) & "Fie", tvwChild, , "<TR><TD>" & rsF(1) & "</TD><TD>" & rsF(2) & "</TD><TD>...</TD></TR>", 2
                End If
                rsF.MoveNext
            Wend
            rsT.MoveNext
    Wend
    Set rsT = Db.OpenResultset("select distinct name from sysobjects where upper(type)='V' and status>=0", 1, 3)
    While Not rsT.EOF
        TV.Nodes.Add Db.Name, tvwChild, "Views" & rsT(0), rsT(0), 6, 6
        TVW.Nodes.Add Db.Name, tvwChild, "Views" & rsT(0), rsT(0), 6, 6
        Set rsF = Db.OpenResultset("select distinct text from syscomments where id=object_id('" & rsT(0) & "')", 1, 3)
        If Not rsF.BOF Then
            TV.Nodes.Add "Views" & rsT(0), tvwChild, "Text" & Right(rsF(0), Len(rsF(0)) - Len("CREATE VIEW dbo. " & rsT(0) & "  as ")), Right(rsF(0), Len(rsF(0)) - Len(" CREATE VIEW dbo. " & rsT(0) & "  as ") + 2), 1, 1
            TVW.Nodes.Add "Views" & rsT(0), tvwChild, "Text" & Right(rsF(0), Len(rsF(0)) - Len("CREATE VIEW dbo. " & rsT(0) & "  as ")), "<TR><TD> " & Right(rsF(0), Len(rsF(0)) - Len(" CREATE VIEW dbo. " & rsT(0) & "  as ") + 2) & "</TD></TR>", 1, 1
        End If
        rsT.MoveNext
    Wend
    Screen.MousePointer = 0
    Exit Sub
Ennd:
    MsgBox Err.Description, vbCritical, "VISDATA SQL"
    Screen.MousePointer = 0
End Sub
Private Sub Form_Resize()
    TV.Left = 0
    TV.Top = 0
    TV.Width = Me.ScaleWidth
    TV.Height = Me.ScaleHeight
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
  '  ASSIGNLIST Me.Caption
    Db.Close
End Sub
Private Sub TV_DblClick()
    TV_MouseU
End Sub
Private Sub TV_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim NNd As Node
    If TV.SelectedItem Is Nothing Then Exit Sub
    Set NNd = TV.SelectedItem
    If (KeyCode = vbKeyDelete Or KeyCode = 110) And NNd.Image = 4 Then
        If MsgBox("Are you sure to Delete Table " & NNd.Text & "?", vbYesNo + vbQuestion, "VISDATA") = vbNo Then Exit Sub
        Db.Execute "drop table [" & NNd.Text & "]"
        TV.Nodes.Remove NNd.Index
    End If
End Sub
Private Sub TV_KeyUp(KeyCode As Integer, Shift As Integer)
    If Shift <> 2 Then Exit Sub
    Dim NNd As Node
    If TV.SelectedItem Is Nothing Then Exit Sub
    Set NNd = TV.SelectedItem
    If NNd.Image = 4 And KeyCode = vbKeyC Then
        Clipboard.SetText CreateQueryStr(NNd)
    ElseIf KeyCode = vbKeyV Then
        ADD_TABLE_NODES Clipboard.GetText(1)
        Clipboard.Clear
    End If
End Sub
Private Sub TV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim NNd As Node
    Set NNd = TV.HitTest(X, Y)
    If NNd Is Nothing Then Exit Sub
    NNd.Selected = True
    TVW.Nodes(NNd.Index).Selected = True
End Sub
Private Sub TV_MouseU()
    Dim lrc As rdoResultset
    Dim NNd As Node
    On Error GoTo Ennd:
    If TV.SelectedItem Is Nothing Then Exit Sub
    Set NNd = TV.SelectedItem
    Dim txtfrm As New frmtxt
    Load txtfrm
    If NNd.Image = 4 Then
        txtfrm.TXTSQL.Text = "SELECT * FROM [" & NNd.Text & "]"
        Set lrc = Db.OpenResultset("SELECT * FROM [" & NNd.Text & "]", 1, 3)
        txtfrm.CollectLV lrc, Db, NNd
    ElseIf NNd.Image = 1 And Left(NNd.Text, 1) <> "+" Then
        txtfrm.TXTSQL.Text = NNd.Text
        Set lrc = Db.OpenResultset(NNd.Text, 1, 3)
        txtfrm.CollectLV lrc, Db, NNd.Parent
    ElseIf NNd.Image = 6 Then
        txtfrm.TXTSQL.Text = NNd.Child.Text
        Set lrc = Db.OpenResultset(NNd.Child.Text, 1, 3)
        txtfrm.CollectLV lrc, Db, NNd
    Else
        Unload txtfrm
        Exit Sub
    End If
    txtfrm.Show
    Exit Sub
Ennd:
    MsgBox Err.Description, vbCritical, "VISDATA"
End Sub
Private Sub TV_NodeCheck(ByVal Node As MSComctlLib.Node)
    Node.Bold = Node.Checked
    TVW.Nodes(Node.Index).Checked = Node.Checked
    TVW.Nodes(Node.Index).Bold = TVW.Nodes(Node.Index).Checked
    If Node.Checked = True Then
        If Node.Child Is Nothing Then Exit Sub
        For I = Node.Child.FirstSibling.Index To Node.Child.LastSibling.Index
            'If TV.Nodes(I).Checked = True Then
                TV.Nodes(I).Checked = True
                TV_NodeCheck TV.Nodes(I)
            'End If
        Next
    Else
        If Node.Child Is Nothing Then Exit Sub
        For I = Node.Child.FirstSibling.Index To Node.Child.LastSibling.Index
            'If TV.Nodes(I).Checked = True Then
                TV.Nodes(I).Checked = False
                TV_NodeCheck TV.Nodes(I)
            'End If
        Next
    End If
End Sub
'Private Sub TV_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Dim DBnaM As String
'   DBnaM = Data.GetData(1)
'   ADD_TABLE_NODES DBnaM
'Ennd:
'    If Err.Number > 0 Then
'        MsgBox Err.Number & " : " & Err.Description, vbInformation, "VISDATA"
'    End If
'End Sub
Private Sub ADD_TABLE_NODES(SStr As String)
    Dim f1 As Integer
    Dim QRY As String
    Dim TNMA As String
ENNW:
   On Error GoTo Ennd:
   f1 = InStr(1, SStr, "!")
   TNMA = Right(SStr, Len(SStr) - f1)
   QRY = Left(SStr, f1 - 1)
   'MsgBox QRY
   Db.Execute QRY
   If Not VisdataMdi.mnuwith.Checked Then
        Db.Execute "DELETE  FROM " & TNMA
   End If
   COLLECT_TREE_NODES
Ennd:
    If Err.Number > 0 Then
        If Err.Number = 3010 Then
            If MsgBox(Err.Description & vbCrLf & "Overwrite Existing Table ?", vbYesNo + vbQuestion, "VisData") = vbYes Then
               Db.Execute Find_Table_Name(Err.Description)
               GoTo ENNW:
            End If
        ElseIf Err.Number = 3290 Then
            MsgBox Err.Number & " : " & Err.Description & SStr, vbInformation, "VISDATA"
        ElseIf UCase(Err.Description) = UCase("S0001: [Microsoft][ODBC SQL Server Driver][SQL Server]There is already an object named '" & Right(Left(TNMA, Len(TNMA) - 1), Len(Left(TNMA, Len(TNMA) - 1)) - 1) & "' in the database.") Then
            If MsgBox("Do you want to Create Copy of " & TNMA & "?", vbYesNo + vbQuestion, "VISDATA") = vbYes Then Dulicate_Table QRY, TNMA
        Else
            MsgBox Err.Number & " : " & Err.Description, vbInformation, "VISDATA"
        End If
    End If
End Sub
Private Sub Dulicate_Table(qury As String, tm As String)
    Dim f1 As Integer
    Dim f2 As Integer
    Dim Fqry As String
    Dim Sqry As String
    On Error GoTo Ennd:
    f1 = InStr(1, qury, "[")
    Fqry = Left(qury, f1 - 1)
    f2 = InStr(1, qury, "]")
    tm = Right(tm, Len(tm) - 1)
    tm = Left(tm, Len(tm) - 1)
    Sqry = Right(qury, Len(qury) - f2)
    qury = Fqry & "[Copy Of " & tm & "] " & Sqry
   Db.Execute qury
   If Not VisdataMdi.mnuwith.Checked Then
        Db.Execute "DELETE  FROM " & tm
   End If
   COLLECT_TREE_NODES
   Exit Sub
Ennd:
   MsgBox Err.Number & " : " & Err.Description, vbInformation, "VISDATA"
End Sub
'Private Sub TV_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
'    If TV.SelectedItem Is Nothing Then Exit Sub
'    If TV.SelectedItem.Image = 4 Then
'        Data.SetData CreateQueryStr(TV.SelectedItem) '& ":" & Me.Caption
'    Else
'        Data.Clear
'    End If
'End Sub
Private Function Find_Table_Name(StRR As String) As String
    Dim f1 As Integer
    Dim tnm As String
    tnm = Right(StRR, Len(StRR) - 7)
    f1 = InStr(1, tnm, "'")
    tnm = "[" & Left(tnm, f1 - 1) & "]"
    Find_Table_Name = "DROP TABLE " & tnm
End Function
Private Function CreateQueryStrR(nnnd As Node) As String
    Dim str As String
    Dim I As Integer
    str = "CREATE TABLE " & nnnd.Text & "("
    For I = 1 To TV.Nodes.Count
        If Not TV.Nodes(I).Parent Is Nothing Then
            If UCase(TV.Nodes(I).Parent.Key) = "TAB" & UCase(nnnd.Text) & "FIE" Then
               str = str & TV.Nodes(I).Text & ","
            End If
        End If
    Next
    str = Left(str, Len(str) - 1) & ")"
    CreateQueryStrR = str
End Function
Private Function CreateQueryStr(nnnd As Node) As String
    Dim str As String
    Dim I As Integer
'    str = "SELECT * INTO [" & nnnd.Text & "] FROM " & TV.Nodes(1).Text & "." & Me.Caption & ".DBO." & "[" & nnnd.Text & "]![" & nnnd.Text & "]"
    str = "SELECT * INTO [" & nnnd.Text & "] FROM " & Me.Caption & ".DBO." & "[" & nnnd.Text & "]![" & nnnd.Text & "]"
    CreateQueryStr = str
End Function

