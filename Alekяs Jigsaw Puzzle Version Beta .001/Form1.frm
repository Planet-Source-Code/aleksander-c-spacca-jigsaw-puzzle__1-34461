VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000C&
   Caption         =   "Jigsaw Puzzle - Beta Version .001 - by Aleksander C. Spacca"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7365
   ClipControls    =   0   'False
   FontTransparent =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   349
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   491
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2760
      Index           =   0
      ItemData        =   "Form1.frx":0000
      Left            =   -450
      List            =   "Form1.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   180
      Width           =   345
   End
   Begin VB.PictureBox pictSource 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   900
      Left            =   -2055
      ScaleHeight     =   58.065
      ScaleMode       =   0  'User
      ScaleWidth      =   58.065
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   8205
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   900
      Index           =   0
      Left            =   -1050
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   8220
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   7200
      Left            =   -10155
      Picture         =   "Form1.frx":0004
      Top             =   -6960
      Visible         =   0   'False
      Width           =   9600
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu s1 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu s5 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu s7 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Const TileSize = 44
Const PieceSize = TileSize + 16

Private Type PuzlePiece
  x As Single
  y As Single
End Type

Dim Clips() As PuzlePiece, SnapLayer()
Dim StartX As Single, StartY As Single
Dim EndX As Single, EndY As Single
Dim i As Integer, ii As Integer
Dim draggingSnap As Boolean


Private Sub Form_Load()
  
  Randomize
  ReDim SnapLayer(0): draggingSnap = False
  Picture1(0).Width = PieceSize: Picture1(0).Height = PieceSize
  Set ShapeTheControls = New clsTransForm
  ReDim Clips(Int(Image1.Width / TileSize), Int(Image1.Height / TileSize))
  Me.Width = Screen.Width: Me.Height = Screen.Height - 400
  Me.ScaleMode = 3: Me.Top = 0: Me.Left = 0
  Form2.TileBkgd Me, Form2.Picture1
  
  xCount = 0
  For xCount = 1 To Int(Image1.Width / TileSize) * Int(Image1.Height / TileSize) - 1
    Load Picture1(xCount)
    Picture1(xCount).Top = Int(RndRange(0, Me.ScaleHeight - TileSize))
    Picture1(xCount).Left = Int(RndRange(0, Me.ScaleWidth - TileSize))
    Picture1(xCount).Visible = True
  Next xCount
  
  xCount = 0
  For i = 0 To Int(Image1.Width / TileSize) - 1
    For ii = 0 To Int(Image1.Height / TileSize) - 1
      Clips(i, ii).x = i * TileSize
      Clips(i, ii).y = ii * TileSize
      Picture1(xCount).Tag = 0
      Picture1(xCount).PaintPicture Image1.Picture, 0, 0, PieceSize, PieceSize, Clips(i, ii).x, Clips(i, ii).y, 60, 60
      ShapeTheControls.ShapeMe Picture1(xCount), RGB(255, 255, 255), True, App.Path & "\piece.dat"
      xCount = xCount + 1
    Next
  Next
  
  Set ShapeTheControls = Nothing
  Me.WindowState = vbMaximized

  Unload SplashScreen: Me.Show

End Sub

Private Function RndRange(ByVal Min As Single, ByVal Max As Single) As Single
  Randomize
  RndRange = (Rnd * (Max - Min + 1)) + Min
End Function

Private Sub Form_Unload(Cancel As Integer)
  Unload Form3
  Unload Form2
  Unload Form1
End Sub

Private Sub mnuLoad_Click()
  Me.MousePointer = vbHourglass
  Form3.Show vbModal
  Me.MousePointer = vbNormal
End Sub

Private Sub mnuSounds_Click()
  If mnuSounds.Checked = False Then
    mnuSounds.Checked = True
  Else
    mnuSounds.Checked = False
  End If
End Sub

Private Sub Picture1_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  EndX = x: EndY = y
  If Picture1(index).Tag <> 0 Then
    For xCount = 0 To List1(Picture1(index).Tag).ListCount - 1
      Picture1(List1(Picture1(index).Tag).List(xCount)).ZOrder 0
    Next
  End If
  Picture1(index).ZOrder 0
  Me.MousePointer = 5
End Sub

Private Sub Picture1_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 1 Then
    If Picture1(index).Tag <> 0 Then
      Me.AutoRedraw = True
      For xCount = 0 To List1(Picture1(index).Tag).ListCount - 1
        Sx = Picture1(List1(Picture1(index).Tag).List(xCount)).Left - EndX + x
        Sy = Picture1(List1(Picture1(index).Tag).List(xCount)).Top - EndY + y
        Picture1(List1(Picture1(index).Tag).List(xCount)).Move Int(Sx), Int(Sy)
      Next
      Me.AutoRedraw = False
    Else
      Sx = Picture1(index).Left - EndX + x: Sy = Picture1(index).Top - EndY + y
      Picture1(index).Move Int(Sx), Int(Sy)
    End If
    StartX = Picture1(index).Left - EndX
    StartY = Picture1(index).Top - EndY
  End If
End Sub

Private Sub Picture1_MouseUp(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  
  On Error Resume Next
  
  Dim PieceSnap: PieceSnap = 0
  P0Number = index: P0PosX = Picture1(P0Number).Left: P0PosY = Picture1(P0Number).Top
  P1Number = P0Number - ii: P1PosX = Picture1(P1Number).Left: P1PosY = Picture1(P1Number).Top
  P2Number = P0Number - 1:  P2PosX = Picture1(P2Number).Left: P2PosY = Picture1(P2Number).Top
  P3Number = P0Number + ii: P3PosX = Picture1(P3Number).Left: P3PosY = Picture1(P3Number).Top
  P4Number = P0Number + 1:  P4PosX = Picture1(P4Number).Left: P4PosY = Picture1(P4Number).Top
  
  If P0PosX > P1PosX + 30 And P0PosX < P1PosX + 50 And P0PosY > P1PosY - 10 And P0PosY < P1PosY + 10 Then
    PieceSnap = P1Number
    Picture1(P0Number).Move P1PosX + TileSize - 1, P1PosY
  ElseIf P0PosX < P3PosX - 30 And P0PosX > P3PosX - 50 And P0PosY > P3PosY - 10 And P0PosY < P3PosY + 10 Then
    PieceSnap = P3Number
    Picture1(P0Number).Move P3PosX - TileSize + 1, P3PosY
  ElseIf P0PosY > P2PosY + 30 And P0PosY < P2PosY + 50 And P0PosX > P2PosX - 10 And P0PosX < P2PosX + 10 Then
    PieceSnap = P2Number
    Picture1(P0Number).Move P2PosX, P2PosY + TileSize - 2
  ElseIf P0PosY < P4PosY - 30 And P0PosY > P4PosY - 50 And P0PosX > P4PosX - 10 And P0PosX < P4PosX + 10 Then
    PieceSnap = P4Number
    Picture1(P0Number).Move P4PosX, P4PosY - TileSize + 2
  End If
  
  If Picture1(P0Number).Tag <> 0 And Picture1(P0Number).Tag = Picture1(PieceSnap).Tag Then GoTo errHandler
  
  If PieceSnap <> 0 Then
    PieceSnapLayer = 0
    xLay = Picture1(P0Number).Tag
    If xLay = 0 Then xLay = Picture1(PieceSnap).Tag
    If xLay = 0 Then
      xLay = UBound(SnapLayer) + 1: ReDim Preserve SnapLayer(xLay)
      SnapLayer(xLay) = "x": Load List1(xLay)
      List1(xLay).Visible = True
    End If
    PieceSnapLayer = xLay
    Picture1(PieceSnap).Tag = PieceSnapLayer: Picture1(P0Number).Tag = PieceSnapLayer
    If CheckGroupDupe(PieceSnapLayer, P0Number) = False Then List1(PieceSnapLayer).AddItem P0Number
    If CheckGroupDupe(PieceSnapLayer, PieceSnap) = False Then List1(PieceSnapLayer).AddItem PieceSnap
  End If
  
errHandler:
  
  Me.MousePointer = 0

End Sub

Private Function CheckGroupDupe(List, Piece) As Boolean
  CheckGroupDupe = False
  For xCount = 0 To List1(List).ListCount - 1
    If List1(List).List(xCount) = Piece Then
      CheckGroupDupe = True: Exit For
    End If
  Next
End Function
