VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVerticalSplitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
#If Win32 Then
   Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
   Private Declare Function ReleaseCapture Lib "user32" () As Long
#Else
   Private Declare Function SetCapture Lib "User" (ByVal hwnd As Integer) As Integer
   Private Declare Function ReleaseCapture Lib "User" () As Integer
#End If
Const THREED_PAD = 0
Const DARKGRAY = &H808080
Const LIGHTGRAY = &HC0C0C0
Private mpicBar As PictureBox
Private mvContainer As Variant
Private mvLeftChild As Variant
Private mvRightChild As Variant
Private mnBarWidth As Integer
Private mbMouseDown As Boolean
Public Property Get BarWidth() As Integer
   BarWidth = mpicBar.Width
End Property
Public Property Let BarWidth(ByVal vnBarWidth As Integer)
   If vnBarWidth > 0 Then
      mpicBar.Width = vnBarWidth
      mnBarWidth = vnBarWidth
      RefreshWidth
   End If
End Property
Public Property Get Height() As Integer
   Height = mvContainer.Height
End Property
Public Property Let Height(ByVal vnHeight As Integer)
   If vnHeight >= 0 Then
      mvContainer.Height = vnHeight
      RefreshHeight
   End If
End Property
Public Property Get Left() As Integer
   Left = mvContainer.Left
End Property
Public Property Let Left(ByVal vnLeft As Integer)
   If vnLeft >= 0 Then
      mvContainer.Left = vnLeft
   End If
End Property
Public Property Get LeftChild() As Variant
   Set LeftChild = mvLeftChild
End Property
Public Property Set LeftChild(Child As Variant)
   Set mvLeftChild = Child
   If Not mvLeftChild Is Nothing Then _
      Set mvLeftChild.Container = mvContainer
   Me.Refresh
End Property
Public Property Get RightChild() As Variant
   Set RightChild = mvRightChild
End Property
Public Property Set RightChild(Child As Variant)
   Set mvRightChild = Child
   If Not mvRightChild Is Nothing Then _
      Set mvRightChild.Container = mvContainer
   Me.Refresh
End Property
Public Property Get Top() As Integer
   Top = mvContainer.Top
End Property
Public Property Let Top(ByVal vnTop As Integer)
   If vnTop >= 0 Then
      mvContainer.Top = vnTop
   End If
End Property
Public Property Get Visible() As Boolean
   Visible = mvContainer.Visible
End Property
Public Property Let Visible(ByVal vbVisible As Boolean)
   mvContainer.Visible = vbVisible
End Property
Public Property Get Width() As Integer
   Width = mvContainer.Width
End Property
Public Property Let Width(ByVal vnWidth As Integer)
   If vnWidth >= 0 Then
      mvContainer.Width = vnWidth
      RefreshWidth
   End If
End Property
Public Sub Init(Container As Variant, Bar As PictureBox, LeftChild As Variant, RightChild As Variant)
    On Error Resume Next
   ' do this once
   Static sbBeenHere As Boolean
   If Not sbBeenHere Then
      sbBeenHere = True
      Set mvContainer = Container
      Set mpicBar = Bar
      Set mvLeftChild = LeftChild
      If Not mvLeftChild Is Nothing Then _
         Set mvLeftChild.Container = mvContainer
      Set mvRightChild = RightChild
      If Not mvRightChild Is Nothing Then _
         Set mvRightChild.Container = mvContainer
      mnBarWidth = mpicBar.Width
      mpicBar.ZOrder vbBringToFront
      Me.Refresh
   End If
End Sub
Public Sub MouseDown()
   mbMouseDown = True
   SetCapture mpicBar.hwnd
   mpicBar.BackColor = DARKGRAY
   mpicBar.Width = mnBarWidth + 15
   mpicBar.ZOrder vbBringToFront
End Sub
Public Sub MouseMove(ByVal X As Single)
   If mbMouseDown Then mpicBar.Left = mpicBar.Left + X
End Sub
Public Sub MouseUp()
   mbMouseDown = False
   ReleaseCapture
   mpicBar.Width = mnBarWidth
   mpicBar.BackColor = LIGHTGRAY
   Dim nMax As Integer
   nMax = mvContainer.ScaleWidth - THREED_PAD
   Dim nMin As Integer
   nMin = 0
   If mpicBar.Left < nMin Then
      mpicBar.Left = nMin
   ElseIf mpicBar.Left > nMax Then
      mpicBar.Left = nMax
   End If
   RefreshWidth
End Sub
Public Sub Move(ByVal Left As Integer, _
   Optional ByVal Top As Variant, _
   Optional ByVal Width As Variant, _
   Optional ByVal Height As Variant)
   Me.Left = Left
   If Not VarType(Top) = vbError Then _
      Me.Top = Top
   If Not VarType(Width) = vbError Then
      Me.Width = Width
   End If
   If Not VarType(Height) = vbError Then
      Me.Height = Height
   End If
End Sub
Public Sub Refresh()
   RefreshHeight
   RefreshWidth
End Sub
Private Sub RefreshHeight()
   mpicBar.Top = 0
   mpicBar.Height = _
      mvContainer.ScaleHeight
   If Not mvLeftChild Is Nothing Then
      mvLeftChild.Top = 0
      mvLeftChild.Height = _
         mvContainer.ScaleHeight
   End If
   If Not mvRightChild Is Nothing Then
      mvRightChild.Top = 0
      mvRightChild.Height = _
         mvContainer.ScaleHeight
   End If
End Sub
Private Sub RefreshWidth()
   If Not mvLeftChild Is Nothing Then
      mvLeftChild.Move 0, 0
      If mpicBar.Left > 0 Then
         mvLeftChild.Width = mpicBar.Left
         mvLeftChild.Visible = True
      Else
         mvLeftChild.Visible = False
      End If
   End If
   If Not mvRightChild Is Nothing Then
      mvRightChild.Move _
         mpicBar.Left + mpicBar.Width, 0
      Dim nRightWidth As Integer
      nRightWidth = mvContainer.ScaleWidth - _
         mpicBar.Left - mpicBar.Width - THREED_PAD
      If nRightWidth > 0 Then
         mvRightChild.Width = nRightWidth
         mvRightChild.Visible = True
      Else
         mvRightChild.Visible = False
      End If
   End If
End Sub
