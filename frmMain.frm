VERSION 5.00
Begin VB.Form Calendar 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   Caption         =   "Calendar"
   ClientHeight    =   10260
   ClientLeft      =   495
   ClientTop       =   690
   ClientWidth     =   4545
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10260
   ScaleWidth      =   4545
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox DaySlot 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   3000
      TabIndex        =   5
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   3000
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   3000
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Make Wallpaper"
      Height          =   615
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Labelss 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   3000
      TabIndex        =   0
      Top             =   720
      Width           =   855
   End
End
Attribute VB_Name = "Calendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public M As Integer
Public Y As Integer
Dim XX As Integer
Dim offset As Integer
Dim YY As Integer
Dim X As Integer
Dim daytext As String
Dim spaces As Integer
Dim temp As Integer
Option Explicit
Private Sub Command1_Click()
' end everything
endall
End Sub
Private Sub Command2_Click()
On Error Resume Next
' open up the two .ini files - text for each day.
' save the changes made.
Open "calendar1.ini" For Output As 1
Open "calendar2.ini" For Output As 2
For X = 0 To (DaysIn(M, Y) + (spaces - 1))
Print #1, Text1(X).Text
Print #2, Text2(X).Text
Next X
10 '
Close
' make the buttons on the screen invisible.
Command1.Visible = False
Command2.Visible = False
' create the wallpaper
Load frmTake
End Sub
Private Sub Form_Load()
    Calendar.Height = Screen.Height - 20
    Calendar.Left = 0
    Calendar.Top = 0
    M = Month(Date)
    Y = Year(Date)
' set where the calendar will be placed
' on the screen.
    offset = 1500
    SetupDays
End Sub
Public Sub SetupDays()
  Dim X As Integer
  Dim H As Integer
  Dim W As Integer
  Dim D As String
  Dim TD As Date
    On Error Resume Next
    For X = 0 To 40
        DaySlot(X).Visible = False
        Text1(X).Visible = False
        Text2(X).Visible = False
    Next X
    DaySlot(0).Left = DaySlot(0).Width + offset
    H = DaySlot(0).Height + Text1(0).Height + Text2(0).Height
    W = DaySlot(0).Width
    YY = DaySlot(0).Top
    XX = DaySlot(0).Left + DaySlot(0).Width
    TD = M & "/01/" & Y
    D = Format(TD, "ddd")
    DaySlot(0).Text = 1
    spaces = GetBlanks(D)
    For X = 1 To (DaysIn(M, Y) + (spaces - 1))
        If X = 6 Or X = 12 Then
            'Beep
        End If
        Load DaySlot(X)
        DaySlot(X).Text = X - (spaces - 1)
        If Val(DaySlot(X).Text) < 1 Then
            DaySlot(X).Text = ""
            DaySlot(X).Visible = False
          Else
            DaySlot(X).Visible = True
            DaySlot(X).Tag = M & "/" & X - (spaces - 1) & "/" & Format(Y, "00")
            DaySlot(X).ToolTipText = DaySlot(X).Tag
        End If
        DaySlot(X).Left = XX + offset
        DaySlot(X).Top = YY
        DaySlot(X).Height = DaySlot(0).Height
        DaySlot(X).Width = DaySlot(0).Width
       
        XX = XX + W
        If X Mod 7 = 6 Then
            YY = YY + H
            XX = DaySlot(0).Left
        End If
    Next X
      
    
Open "calendar1.ini" For Input As 1
    Text1(0).Left = Text1(0).Width + offset
    H = Text1(0).Height + DaySlot(0).Height + Text2(0).Height
    W = Text1(0).Width
    YY = Text1(0).Top
    XX = Text1(0).Left + Text1(0).Width
    TD = M & "/01/" & Y
    D = Format(TD, "ddd")
    Line Input #1, daytext
    Text1(0).Text = "."
    spaces = GetBlanks(D)
    For X = 1 To (DaysIn(M, Y) + (spaces - 1))
    Line Input #1, daytext
   
        Load Text1(X)
        'temp = X - (spaces - 1)
        If Val(DaySlot(X).Text) < 1 Then
            Text1(X).Text = ""
            Text1(X).Visible = False
          Else
        Text1(X).Text = daytext
        Text1(X).Visible = True
        End If
        Text1(X).Left = XX + offset
        Text1(X).Top = YY
        Text1(X).Height = Text1(0).Height
        Text1(X).Width = Text1(0).Width
       
        XX = XX + W
        If X Mod 7 = 6 Then
            YY = YY + H
            XX = Text1(0).Left
        End If
      Next X
    Close
    
    Open "calendar2.ini" For Input As 1
    Text2(0).Left = Text2(0).Width + offset
    H = Text2(0).Height + DaySlot(0).Height + Text1(0).Height
    W = Text2(0).Width
    YY = Text2(0).Top
    XX = Text2(0).Left + Text2(0).Width
    TD = M & "/01/" & Y
    D = Format(TD, "ddd")
    Line Input #1, daytext
    Text2(0).Text = daytext
    spaces = GetBlanks(D)
    For X = 1 To (DaysIn(M, Y) + (spaces - 1))
    Line Input #1, daytext
   
        Load Text2(X)
        'temp = X - (spaces - 1)
        If Val(DaySlot(X).Text) < 1 Then
            Text2(X).Text = ""
            Text2(X).Visible = False
          Else
        Text2(X).Text = daytext
        Text2(X).Visible = True
        End If
        Text2(X).Left = XX + offset
        Text2(X).Top = YY
        Text2(X).Height = Text2(0).Height
        Text2(X).Width = Text2(0).Width
       
        XX = XX + W
        If X Mod 7 = 6 Then
            YY = YY + H
            XX = Text2(0).Left
        End If
    Next X
    Close
    
    For X = 0 To 6
        Load Labelss(X)
        Labelss(X).Left = DaySlot(X).Left
        Labelss(0).Left = DaySlot(0).Left + offset
        Labelss(X).Top = DaySlot(X).Top - Labelss(X).Height
        Labelss(X).Width = DaySlot(X).Width
        Labelss(X).Visible = True
    Next X
    
    Labelss(0) = "Sun"
    Labelss(1) = "Mon"
    Labelss(2) = "Tue"
    Labelss(3) = "Wed"
    Labelss(4) = "Thu"
    Labelss(5) = "Fri"
    Labelss(6) = "Sat"
    Width = Labelss(6).Left + (Labelss(6).Width * 2)
    
    lblTitle.Alignment = vbCenter
    lblTitle.Left = Labelss(0).Left
    lblTitle.Width = Labelss(6).Left + Labelss(6).Width - Labelss(0).Left
    lblTitle.Top = Labelss(0).Top - (lblTitle.Height + 60)
    
    TD = M & "/01/" & Y
    lblTitle = Format(TD, "mmmm") & ", " & Format(TD, "yyyy")
End Sub
Function DaysIn(Mnth As Integer, Yeer As Integer) As Integer
  Dim ThisDate As Date
    ThisDate = Mnth & "/01/" & Yeer
    Do Until Month(DateAdd("D", 1, ThisDate)) <> Mnth
        ThisDate = DateAdd("D", 1, ThisDate)
    Loop
    DaysIn = Day(ThisDate)
End Function

Function GetBlanks(D As String) As Integer
    Select Case D
        Case "Sun"
            GetBlanks = 0
        Case "Mon"
            GetBlanks = 1
        Case "Tue"
            GetBlanks = 2
        Case "Wed"
            GetBlanks = 3
        Case "Thu"
            GetBlanks = 4
        Case "Fri"
            GetBlanks = 5
        Case "Sat"
            GetBlanks = 6
    End Select
End Function
Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub


