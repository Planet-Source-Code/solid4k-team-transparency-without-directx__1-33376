VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transperancy without DirectX"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   3300
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   735
      Left            =   1680
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   1
      Top             =   0
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' VB Transperancy without DirectX!
' by: Solid - SOLID4K, inc.
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Sub Form_Load()
Picture1.Picture = VB.LoadPicture(App.Path & "\BG.bmp") ' Load the background pic
Picture2.Picture = VB.LoadPicture(App.Path & "\BALL.bmp") ' Load the ball pic
Me.Height = Picture1.ScaleHeight + 460 ' Set height to fit picture boxes
Me.Show ' Show the form so the picture's HDC can be read
Call PixelShit ' Call the PixelShit function to the bottom
End Sub
Public Sub PixelShit()
Dim CurPix As Long ' The current pixel's color
For xx = 0 To Picture2.ScaleWidth ' Picture2's X loop
    DoEvents ' Lets the computer do what it needs to do
    For yy = 0 To Picture2.ScaleHeight ' Picture2's Y loop
        DoEvents ' Lets the computer do what it needs to do
        Let CurPix = GetPixel(Picture2.hdc, xx, yy) ' This gets the current pixel's color
        If CurPix = vbWhite Then GoTo jumpz: ' Transperency color white, if it is white, goto next pixel
        If CurPix = vbBlack Then GoTo jumpz: ' This sometimes happens due to size difference
        Call SetPixel(Picture1.hdc, xx, yy, CurPix) ' Puts the current pixel's color onto the BG
        
jumpz: ' Used to goto the next Y pixel
    Next yy
Next xx
End Sub
