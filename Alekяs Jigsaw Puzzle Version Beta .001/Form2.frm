VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   3630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6705
   ControlBox      =   0   'False
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3630
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   60
      Picture         =   "Form2.frx":000C
      ScaleHeight     =   4455
      ScaleWidth      =   7290
      TabIndex        =   0
      Top             =   -225
      Visible         =   0   'False
      Width           =   7290
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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




