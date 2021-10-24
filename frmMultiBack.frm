VERSION 5.00
Begin VB.Form frmMultiBack 
   AutoRedraw      =   -1  'True
   Caption         =   "Multiple Backgrounds"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   305
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerScroll 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   2760
      Top             =   3840
   End
End
Attribute VB_Name = "frmMultiBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Chapter 1
'Multiple backgrounds
'

Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'Constants for the GenerateDC function
'**LoadImage Constants**
Const IMAGE_BITMAP As Long = 0
Const LR_LOADFROMFILE As Long = &H10
Const LR_CREATEDIBSECTION As Long = &H2000
'****************************************

'Display window size
Const ScrollWidth As Long = 250

'Sprite position
Dim XSprite As Long

'Sizes of the bitmaps
Const SpriteWidth As Long = 64
Const SpriteHeight As Long = 64

Const ABBackHeight As Long = 250
Const ABBAckWidth As Long = 750

Const Back1Height As Long = 100
Const Back1Width As Long = 600

Const ForeHeight As Long = 50
Const ForeWidth As Long = 400


'Sprite DC
Dim DCSprite As Long
Dim DCSpriteM As Long

'Absolute back
Dim DCABBAck As Long

'First Background DCs
Dim DCBack1 As Long
Dim DCBack1M As Long

'Foreground DCs
Dim DCFore As Long
Dim DCForeM As Long

Private Sub cmdExit_Click()



End Sub

Private Sub cmdStart_KeyPress(KeyAscii As Integer)
MsgBox KeyAscii
End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler

Dim AppPath As String

AppPath = App.Path & "\"

'Generate the DCs

DCSprite = GenerateDC(AppPath & "sprite.bmp")
DCSpriteM = GenerateDC(AppPath & "mask.bmp")

DCABBAck = GenerateDC(AppPath & "abback.bmp")

DCBack1 = GenerateDC(AppPath & "1back.bmp")
DCBack1M = GenerateDC(AppPath & "1backm.bmp")

DCFore = GenerateDC(AppPath & "fore.bmp")
DCForeM = GenerateDC(AppPath & "forem.bmp")

'Resize the form
Me.Move Me.Left, Me.Top, ScrollWidth * Screen.TwipsPerPixelX, Me.Height


'Set the sprite position on the center of the display
XSprite = ScrollWidth / 2 - SpriteWidth / 2


ErrorHandler:

    Select Case Err
        
        Case 0 'No errors
        
        Case Else
            
            MsgBox "Failure in loading graphics"
            'clean up
            CleanUp
    End Select
    TimerScroll.Enabled = True
End Sub
Private Sub CleanUp()

DeleteGeneratedDC DCSprite
DeleteGeneratedDC DCSpriteM
DeleteGeneratedDC DCABBAck
DeleteGeneratedDC DCBack1
DeleteGeneratedDC DCBack1M
DeleteGeneratedDC DCFore
DeleteGeneratedDC DCForeM

Unload Me
Set frmMultiBack = Nothing

End Sub

'IN: FileName: The file name of the graphics
'OUT: The Generated DC
Public Function GenerateDC(FileName As String) As Long
Dim DC As Long
Dim hBitmap As Long

'Create a Device Context, compatible with the screen
DC = CreateCompatibleDC(0)

If DC < 1 Then
    GenerateDC = 0
    'Raise error
    Err.Raise vbObjectError + 1
    Exit Function
End If

'Load the image....BIG NOTE: This function is not supported under NT, there you can not
'specify the LR_LOADFROMFILE flag
hBitmap = LoadImage(0, FileName, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE Or LR_CREATEDIBSECTION)

If hBitmap = 0 Then 'Failure in loading bitmap
    DeleteDC DC
    GenerateDC = 0
    'Raise error
    Err.Raise vbObjectError + 2
    Exit Function
End If

'Throw the Bitmap into the Device Context
SelectObject DC, hBitmap

'Return the device context
GenerateDC = DC

'Delte the bitmap handle object
DeleteObject hBitmap

End Function
'Deletes a generated DC
Private Function DeleteGeneratedDC(DC As Long) As Long

If DC > 0 Then
    DeleteGeneratedDC = DeleteDC(DC)
Else
    DeleteGeneratedDC = 0
End If

End Function

Private Sub Form_Resize()
XSprite = ScrollWidth / 2 - SpriteWidth / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
CleanUp
End Sub

Private Sub TimerScroll_Timer()
Static X As Long, XBack1 As Long, XFore As Long
Dim GlueWidth As Long, EndScroll As Long

'Draw the absolute background
If X + ScrollWidth > ABBAckWidth Then 'We ned to glue at the beginnig again
    'Calculate the remaining width
    GlueWidth = X + ScrollWidth - ABBAckWidth
    EndScroll = ScrollWidth - GlueWidth
    
    'Blit the first part
    BitBlt Me.hdc, 0, 0, EndScroll, ABBackHeight, DCABBAck, X, 0, vbSrcCopy
    'Now draw from the beginning again
    BitBlt Me.hdc, EndScroll, 0, GlueWidth, ABBackHeight, DCABBAck, 0, 0, vbSrcCopy
Else
    BitBlt Me.hdc, 0, 0, ScrollWidth, ABBackHeight, DCABBAck, X, 0, vbSrcCopy
End If

'Draw the first back ground
If XBack1 + ScrollWidth > Back1Width Then 'We ned to glue at the beginnig again
    'Calculate the remaining width
    GlueWidth = XBack1 + ScrollWidth - Back1Width
    EndScroll = ScrollWidth - GlueWidth
    'Blit the first part
    BitBlt Me.hdc, 0, ABBackHeight - Back1Height, EndScroll, Back1Height, DCBack1M, XBack1, 0, vbSrcAnd
    BitBlt Me.hdc, 0, ABBackHeight - Back1Height, EndScroll, Back1Height, DCBack1, XBack1, 0, vbSrcPaint
    'Now draw from the beginning again
    BitBlt Me.hdc, EndScroll, ABBackHeight - Back1Height, GlueWidth, Back1Height, DCBack1M, 0, 0, vbSrcAnd
    BitBlt Me.hdc, EndScroll, ABBackHeight - Back1Height, GlueWidth, Back1Height, DCBack1, 0, 0, vbSrcPaint
Else
    BitBlt Me.hdc, 0, ABBackHeight - Back1Height, ScrollWidth, Back1Height, DCBack1M, XBack1, 0, vbSrcAnd
    BitBlt Me.hdc, 0, ABBackHeight - Back1Height, ScrollWidth, Back1Height, DCBack1, XBack1, 0, vbSrcPaint
End If

'Draw the sprite
BitBlt Me.hdc, XSprite, ABBackHeight - SpriteHeight, SpriteWidth, SpriteHeight, DCSpriteM, 0, 0, vbSrcAnd
BitBlt Me.hdc, XSprite, ABBackHeight - SpriteHeight, SpriteWidth, SpriteHeight, DCSprite, 0, 0, vbSrcPaint

'Draw the fore ground
If XFore + ScrollWidth > ForeWidth Then 'We ned to glue at the beginnig again
    'Calculate the remaining width
    GlueWidth = XFore + ScrollWidth - ForeWidth
    EndScroll = ScrollWidth - GlueWidth
    'Blit the first part
    BitBlt Me.hdc, 0, ABBackHeight - ForeHeight, EndScroll, ForeHeight, DCForeM, XFore, 0, vbSrcAnd
    BitBlt Me.hdc, 0, ABBackHeight - ForeHeight, EndScroll, ForeHeight, DCFore, XFore, 0, vbSrcPaint
    'Now draw from the beginning again
    BitBlt Me.hdc, EndScroll, ABBackHeight - ForeHeight, GlueWidth, ForeHeight, DCForeM, 0, 0, vbSrcAnd
    BitBlt Me.hdc, EndScroll, ABBackHeight - ForeHeight, GlueWidth, ForeHeight, DCFore, 0, 0, vbSrcPaint
Else
    BitBlt Me.hdc, 0, ABBackHeight - ForeHeight, ScrollWidth, ForeHeight, DCForeM, XFore, 0, vbSrcAnd
    BitBlt Me.hdc, 0, ABBackHeight - ForeHeight, ScrollWidth, ForeHeight, DCFore, XFore, 0, vbSrcPaint
End If


Me.Refresh

'Modify the positions.
X = (X Mod ABBAckWidth) + 1
XBack1 = (XBack1 Mod Back1Width) + 5
XFore = (XFore Mod ForeWidth) + 25

End Sub
