VERSION 5.00
Begin VB.Form frmprog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generate Structure"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   Icon            =   "frmprog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdclose 
      Caption         =   "&Close"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2595
      TabIndex        =   1
      Top             =   1485
      Width           =   1770
   End
   Begin VB.Label lblprog 
      Alignment       =   2  'Center
      Height          =   990
      Left            =   420
      TabIndex        =   0
      Top             =   150
      Width           =   6450
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmprog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cmdclose.Enabled = False
End Sub
