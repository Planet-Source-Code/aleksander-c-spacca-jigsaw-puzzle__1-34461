Attribute VB_Name = "Module1"

Public Sub FadeFormBlue(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 0, 255 - intLoop), B
    Next intLoop
End Sub

Public Sub TileBkgd(frm As Form, picholder As PictureBox)
  Dim ScWidth%, ScHeight%, ScMode%, n%, o%
  ScMode% = frm.ScaleMode
  picholder.ScaleMode = 3
  frm.ScaleMode = 3
  picholder.ScaleMode = 3
  For n% = 0 To frm.Height Step picholder.ScaleHeight
    For o% = 0 To frm.Width Step picholder.ScaleWidth
      frm.PaintPicture picholder.Picture, o%, n%
    Next o%
  Next n%
  frm.ScaleMode = ScMode%
End Sub

Public Sub TileMDIBkgd(MDIForm As Form, bkgdtiler As Form)
  Dim ScWidth%, ScHeight%
  ScWidth% = Screen.Width / Screen.TwipsPerPixelX
  ScHeight% = Screen.Height / Screen.TwipsPerPixelY
  Load bkgdtiler
  bkgdtiler.Height = Screen.Height
  bkgdtiler.Width = Screen.Width
  bkgdtiler.ScaleMode = 3
  bkgdtiler!Picture1.Top = 0
  bkgdtiler!Picture1.Left = 0
  bkgdtiler!Picture1.ScaleMode = 3
  For n% = 0 To ScHeight% Step bkgdtiler!Picture1.ScaleHeight
    For o% = 0 To ScWidth% Step bkgdtiler!Picture1.ScaleWidth
      bkgdtiler.PaintPicture bkgdtiler!Picture1.Picture, o%, n%
    Next o%
  Next n%
  MDIForm.Picture = bkgdtiler.Image
  Unload bkgdtiler
End Sub
