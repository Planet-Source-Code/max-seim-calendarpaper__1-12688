VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTake 
   AutoRedraw      =   -1  'True
   Caption         =   "Auto Screen-Shot Taker"
   ClientHeight    =   930
   ClientLeft      =   10275
   ClientTop       =   2160
   ClientWidth     =   2040
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTake.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   62
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   136
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   1440
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save Screen Shot"
      Filter          =   "Bitmap (*.BMP)|*.BMP"
   End
   Begin VB.PictureBox picScreenShot 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   1
      Top             =   0
      Width           =   1335
      Begin VB.Label lblCaption 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   2775
      End
   End
   Begin VB.PictureBox FinishedProduct 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   1335
      TabIndex        =   0
      Top             =   0
      Width           =   1365
   End
   Begin VB.Menu mnuOpts 
      Caption         =   "Options"
      Visible         =   0   'False
      Begin VB.Menu mnuGrab 
         Caption         =   "Grab Screen Shot"
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddCaption 
         Caption         =   "Add Caption"
      End
      Begin VB.Menu mnuHideCaption 
         Caption         =   "Hide Caption"
      End
      Begin VB.Menu mnuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveBitmap 
         Caption         =   "Save as Bitmap"
      End
   End
End
Attribute VB_Name = "frmTake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As String, ByVal fuWinIni As Long) As Long
    Const SPIF_UPDATEINIFILE = &H1
    Const SPI_SETDESKWALLPAPER = 20
    Const SPIF_SENDWININICHANGE = &H2
Private Function SetWallpaper(sFileName As String) As Long
  SetWallpaper = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, sFileName, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)
End Function

Private Sub Form_Load()
Dim TheHWND As Long, Capt As String
Calendar.Show
'Ask the user to input the caption of the window
'If they click Cancel, or leave it blank, don't continue
Capt = "Calendar"
If Capt = "" Then Exit Sub

'Find the window the user specified
TheHWND = FindWindow(vbNullString, Capt)

'Exit the sub and notify the user if we couldn't
'find the window
If TheHWND = 0 Then MsgBox "Window could not be located.", vbExclamation + vbOKOnly: Exit Sub

'Reposition the window we're taking a screen shot of
'so it is at the very top-left of the screen
Window_Reposition TheHWND

'Display the Bright-Blue backdrop form
frmBackDrop.Show

'Pause briefly
Pause 0.2

'Set the focus to the window we're taking a screen shot
'of
Window_SetFocus TheHWND

'Pause briefly
Pause 0.2

'Take a snapshot of the entire screen
GrabScreenShot frmBackDrop.hdc

'Pause briefly
Pause 0.1

'Copy JUST the screen shot of the window into
'the picturebox on this form
Copy_Form frmBackDrop, picScreenShot, BACKDROPCOLOR

'Unload the Bright-Blue backdrop form
Unload frmBackDrop

'Resize the border picturebox
FinishedProduct.Width = picScreenShot.Width + 2
FinishedProduct.Height = picScreenShot.Height + 2

'Resize the form
Me.Height = FinishedProduct.ScaleHeight + 1170
Me.Width = FinishedProduct.ScaleWidth + 360
'Calendar.Hide
Dim Y As Integer
    If lblCaption.Visible = True Then
    
        'Draw over any old captions
        For Y = lblCaption.Top To (lblCaption.Top + 15)
            picScreenShot.ForeColor = vbBlack
            picScreenShot.Line (0, Y)-(picScreenShot.Width, Y)
        Next Y
        
        'Change forecolor
        picScreenShot.ForeColor = vbWhite
        
        'Tell the picturebox where to type the caption
        picScreenShot.CurrentX = 1
        picScreenShot.CurrentY = lblCaption.Top
        
        'Print the caption onto the picturebox
        picScreenShot.Print lblCaption.Caption
    End If
    
    'Save the picture using VB's built in image
    'saving sub
    SavePicture picScreenShot.Image, "c:\calendar.bmp"
SetWallpaper ("c:\calendar.bmp")
endall

End Sub
