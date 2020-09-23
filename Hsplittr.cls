VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHorizontalSplitter"
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
Const THREED_PAD = 30
Const DARKGRAY = &H808080
Const LIGHTGRAY = &HC0C0C0
Private mpicBar As PictureBox
Private mvContainer As Variant
Private mvTopChild As Variant
Private mvBottomChild As Variant
Private mnBarHeight As Integer
Private mbMouseDown As Boolean
Public Property Get BarHeight() As Integer
   BarHeight = mpicBar.Height
End Property
Public Property Let BarHeight(ByVal vnBarHeight As Integer)
   If vnBarHeight > 0 Then
      mpicBar.Height = vnBarHeight
      mnBarHeight = vnBarHeight
      RefreshHeight
   End If
End Property
Public Property Get BottomChild() As Variant
   Set BottomChild = mvBottomChild
End Property
Public Property Set BottomChild(Child As Variant)
   Set mvBottomChild = Child
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
Public Property Get TopChild() As Variant
   Set TopChild = mvTopChild
End Property
Public Property Set TopChild(Child As Variant)
   Set mvTopChild = Child
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
Public Sub Init(Container As Variant, _
   Bar As PictureBox, _
   TopChild As Variant, _
   BottomChild As Variant)
    On Error Resume Next
   ' do this once
   Static sbBeenHere As Boolean
   If Not sbBeenHere Then
      sbBeenHere = True
      Set mvContainer = Container
      Set mpicBar = Bar
      Set mvTopChild = TopChild
      If Not mvTopChild Is Nothing Then _
         Set mvTopChild.Container = mvContainer
      Set mvBottomChild = BottomChild
      If Not mvBottomChild Is Nothing Then _
         Set mvBottomChild.Container = mvContainer
      mnBarHeight = mpicBar.Height
      mpicBar.ZOrder vbBringToFront
      Me.Refresh
   End If
End Sub
Public Sub MouseDown()
   mbMouseDown = True
   SetCapture mpicBar.hwnd
   mpicBar.BackColor = DARKGRAY
   mpicBar.Height = mnBarHeight + 15
   mpicBar.ZOrder vbBringToFront
End Sub
Public Sub MouseMove(ByVal Y As Single)
   If mbMouseDown Then mpicBar.Top = mpicBar.Top + Y
End Sub
Public Sub MouseUp()
   mbMouseDown = False
   ReleaseCapture
   mpicBar.Height = mnBarHeight
   mpicBar.BackColor = LIGHTGRAY
   Dim nMax As Integer
   nMax = mvContainer.ScaleHeight - THREED_PAD
   Dim nMin As Integer
   nMin = 0
   If mpicBar.Top < nMin Then
      mpicBar.Top = nMin
   ElseIf mpicBar.Top > nMax Then
      mpicBar.Top = nMax
   End If
   RefreshHeight
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
   If Not mvTopChild Is Nothing Then
      mvTopChild.Top = 0
      If mpicBar.Top > 0 Then
         mvTopChild.Height = mpicBar.Top
         mvTopChild.Visible = True
      Else
         mvTopChild.Visible = False
      End If
   End If
   If Not mvBottomChild Is Nothing Then
      mvBottomChild.Top = _
         mpicBar.Top + mpicBar.Height
      Dim nBottomHeight As Integer
      nBottomHeight = mvContainer.ScaleHeight - _
         mpicBar.Top - mpicBar.Height - THREED_PAD
      If nBottomHeight > 0 Then
         mvBottomChild.Height = nBottomHeight
         mvBottomChild.Visible = True
      Else
         mvBottomChild.Visible = False
      End If
   End If
End Sub
Private Sub RefreshWidth()
   mpicBar.Left = 0
   mpicBar.Width = _
      mvContainer.ScaleWidth
   If Not mvTopChild Is Nothing Then
      mvTopChild.Left = 0
      mvTopChild.Width = _
         mvContainer.ScaleWidth
   End If
   If Not mvBottomChild Is Nothing Then
      mvBottomChild.Left = 0
      mvBottomChild.Width = _
         mvContainer.ScaleWidth
   End If
End Sub
