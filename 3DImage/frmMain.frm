VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "3D Image"
   ClientHeight    =   2670
   ClientLeft      =   150
   ClientTop       =   315
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   3765
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox rightImg 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   60
      ScaleHeight     =   169
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   241
      TabIndex        =   2
      Top             =   60
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.PictureBox img3D 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   60
      ScaleHeight     =   169
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   241
      TabIndex        =   1
      Top             =   60
      Width           =   3615
   End
   Begin VB.PictureBox leftImg 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   60
      ScaleHeight     =   169
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   241
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Menu CHGIMG 
      Caption         =   "Load Image Set"
      Begin VB.Menu IS 
         Caption         =   "Image Set 1"
         Index           =   1
      End
      Begin VB.Menu IS 
         Caption         =   "Image Set 2"
         Index           =   2
      End
      Begin VB.Menu IS 
         Caption         =   "Image Set 3"
         Index           =   3
      End
      Begin VB.Menu IS 
         Caption         =   "Image Set 4"
         Index           =   4
      End
      Begin VB.Menu IS 
         Caption         =   "Image Set 5"
         Index           =   5
      End
      Begin VB.Menu IS 
         Caption         =   "Image Set 6"
         Index           =   6
      End
      Begin VB.Menu IS 
         Caption         =   "Image Set 7"
         Index           =   7
      End
      Begin VB.Menu IS 
         Caption         =   "Image Set 8"
         Index           =   8
      End
      Begin VB.Menu IS 
         Caption         =   "Image Set 9"
         Index           =   9
      End
      Begin VB.Menu IS 
         Caption         =   "Image Set 10"
         Index           =   10
      End
      Begin VB.Menu IS 
         Caption         =   "Image Set 11"
         Index           =   11
      End
      Begin VB.Menu IS 
         Caption         =   "Image Set 12"
         Index           =   12
      End
      Begin VB.Menu IS 
         Caption         =   "Image Set 13"
         Index           =   13
      End
      Begin VB.Menu IS 
         Caption         =   "Image Set 14"
         Index           =   14
      End
      Begin VB.Menu IS 
         Caption         =   "Image Set 15"
         Index           =   15
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Byte
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Sub Create3D()
  Dim X As Long, Y As Long
  Dim Color1 As Long, Color2 As Long
  Dim R As Integer, G As Integer, B As Integer
  
  'set the mousepointer to "hourglass"
  img3D.MousePointer = 11
 
  'resizes the form to the size of the output image
  frmMain.Width = img3D.Width
  frmMain.Height = img3D.Height + 650
  frmMain.Refresh
  
  'allow the form to refresh
  DoEvents
  
  'formulas for extracting RGB from a pixel
  'r = (Color1 Mod 256)
  'b = (Int(Color1 \ 65536))
  'g = ((Color1 - (b1 * 65536) - r1) \ 256)

  For Y = 0 To img3D.ScaleHeight
    For X = 0 To img3D.ScaleWidth
      
      'get the current pixel's color values from both "left" and "right" pictures
      Color1 = GetPixel(leftImg.hDC, X, Y)
      Color2 = GetPixel(rightImg.hDC, X, Y)

      'blue channel from "right" image is mapped to the red output channel
      R = (Int(Color2 \ 65536))

      'blue channel from "left" image is mapped to the blue output channel
      B = (Int(Color1 \ 65536))
      
      'green channel from "left" image is mapped to the green output channel
      G = ((Color1 - ((Int(Color1 \ 65536)) * 65536) - (Color1 Mod 256)) \ 256)

      'paints the output pixel with the derived RGB values
      SetPixelV img3D.hDC, X, Y, RGB(R, G, B)
    Next X
    
    'refresh the ouput image every row (just for effect)
    img3D.Refresh
   Next Y
   
   'one final refresh
   img3D.Refresh
   
   'set the mouse pointer back to "normal"
   img3D.MousePointer = 1
End Sub
Private Sub Form_Load()

  'moves the visible output image to the top left corner of the form
  img3D.Move 0, 0
  
  'set scalemode to "pixel"
  leftImg.ScaleMode = 3
  rightImg.ScaleMode = 3
  img3D.ScaleMode = 3
  
  'resizes the form to the size of the output image
  frmMain.Width = img3D.Width
  frmMain.Height = img3D.Height + 650
End Sub
Private Sub IS_Click(Index As Integer)
  
  'loads left image from defined set
  Set leftImg.Picture = LoadPicture(App.Path & "\3DImages\" & CStr(Index) & "_l.jpg")
  
  'loads right image from defined set
  Set rightImg.Picture = LoadPicture(App.Path & "\3DImages\" & CStr(Index) & "_r.jpg")
  
  'sets the output image to the left image (just for effect)
  Set img3D = leftImg
  
  'calls routine to generate the 3D image
  Create3D
End Sub
