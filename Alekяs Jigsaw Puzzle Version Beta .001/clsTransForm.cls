VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTransForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetRegionData Lib "gdi32" (ByVal hRgn As Long, ByVal dwCount As Long, lpRgnData As Any) As Long
Private Declare Function ExtCreateRegion Lib "gdi32" (lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long

Private Const RGN_DIFF = 4
Private objName As Object
Private lngHeight As Long, lngWidth As Long

Dim CurRgn As Long, TempRgn As Long

Public Sub ShapeMe(Name As Object, Color As Long, Optional Load As Boolean = True, Optional FileName As String = vbNullString)

  Dim x As Long, y As Long
  Dim booMiddleOfSet As Boolean
  Dim colPoints As Collection
  Set colPoints = New Collection
  Dim Z As Variant
  Dim lngTransY As Long
  Dim lngTransStartX As Long
  Dim lngTransEndX As Long
  Dim lngHDC As Long

  Set objName = Name
  
  With objName
    .AutoRedraw = True
    .ScaleMode = 3
    .ClipControls = False
    lngHeight = .ScaleHeight
    lngWidth = .ScaleWidth
  End With
    
  If FileName <> vbNullString And Load = True Then
    LoadRegionDataFromFile FileName
  Else
    booMiddleOfSet = False
    For y = 0 To lngHeight
      lngTransY = y
      For x = 0 To lngWidth
        If objName.Point(x, y) = Color Then
          If booMiddleOfSet = False Then
            lngTransStartX = x
            lngTransEndX = x
            booMiddleOfSet = True
          Else
            lngTransEndX = x
          End If
        Else
          If booMiddleOfSet Then
            colPoints.Add Array(lngTransY, lngTransStartX, lngTransEndX)
            booMiddleOfSet = False
          End If
        End If
      Next x
    Next y
    CurRgn = CreateRectRgn(0, 0, lngWidth, lngHeight)
  End If
    
 If FileName <> vbNullString And Load = True Then
 Else
    For Each Z In colPoints
      TempRgn = CreateRectRgn(Z(1), Z(0), Z(2) + 1, Z(0) + 1)
      CombineRgn CurRgn, CurRgn, TempRgn, RGN_DIFF
      DeleteObject (TempRgn)
    Next
  End If

  If FileName <> vbNullString And Load = False Then
    SaveRegionDataToFile FileName
  End If

  SetWindowRgn objName.hWnd, CurRgn, True
  ReleaseDC objName.hWnd, lngHDC
  Set colPoints = Nothing

End Sub

Private Sub SaveRegionDataToFile(ByVal sPath As String)
  Dim iFile As Long
  Dim nBytes As Long
  Dim b() As Byte
  On Error GoTo ErrorHandler
  nBytes = GetRegionData(CurRgn, 0, ByVal 0&)
  If nBytes > 0 Then
     ReDim b(0 To nBytes - 1) As Byte
     If nBytes = GetRegionData(CurRgn, nBytes, b(0)) Then
        On Error Resume Next
        Kill sPath
        On Error GoTo ErrorHandler
        iFile = FreeFile
        Open sPath For Binary Access Write Lock Read As #iFile
        Put #iFile, , b
        Close #iFile
     End If
  End If
ErrorHandler:
End Sub

Private Sub LoadRegionDataFromFile(ByVal sFileName As String)
  Dim iFile As Long
  Dim b() As Byte
  Dim dwCount As Long
  On Error GoTo ErrorHandler
  iFile = FreeFile
  Open sFileName For Binary Access Read Lock Write As #iFile
  ReDim b(0 To LOF(iFile) - 1) As Byte
  Get #iFile, , b
  Close #iFile
  dwCount = UBound(b) - LBound(b) + 1
  CurRgn = ExtCreateRegion(ByVal 0&, dwCount, b(0))
ErrorHandler:
End Sub
