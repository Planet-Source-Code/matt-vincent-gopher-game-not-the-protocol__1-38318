VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGopher 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gopher"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGopher.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   6540
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Level Select"
      Height          =   825
      Left            =   1470
      MouseIcon       =   "frmGopher.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   1530
      Width           =   4965
      Begin VB.Image imgButtons 
         Height          =   405
         Index           =   1
         Left            =   180
         Top             =   270
         Width           =   405
      End
      Begin VB.Image imgButtons 
         Height          =   405
         Index           =   2
         Left            =   720
         Top             =   270
         Width           =   405
      End
      Begin VB.Image imgButtons 
         Height          =   405
         Index           =   3
         Left            =   1230
         Top             =   270
         Width           =   405
      End
      Begin VB.Image imgButtons 
         Height          =   405
         Index           =   4
         Left            =   1770
         Top             =   270
         Width           =   405
      End
      Begin VB.Image imgButtons 
         Height          =   405
         Index           =   5
         Left            =   2280
         Top             =   270
         Width           =   405
      End
      Begin VB.Image imgButtons 
         Height          =   405
         Index           =   6
         Left            =   2790
         Top             =   270
         Width           =   405
      End
      Begin VB.Image imgButtons 
         Height          =   405
         Index           =   7
         Left            =   3300
         Top             =   270
         Width           =   405
      End
      Begin VB.Image imgButtons 
         Height          =   405
         Index           =   8
         Left            =   3810
         Top             =   270
         Width           =   405
      End
      Begin VB.Image imgButtons 
         Height          =   405
         Index           =   9
         Left            =   4350
         Top             =   270
         Width           =   405
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Game Control"
      Height          =   825
      Left            =   90
      MouseIcon       =   "frmGopher.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   1530
      Width           =   1305
      Begin VB.Image imgButtons 
         Height          =   405
         Index           =   10
         Left            =   180
         Top             =   270
         Width           =   405
      End
      Begin VB.Image imgButtons 
         Height          =   405
         Index           =   11
         Left            =   720
         Top             =   270
         Width           =   405
      End
   End
   Begin MSComctlLib.ImageList imgButtonsUp 
      Left            =   6660
      Top             =   1500
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   27
      ImageHeight     =   27
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGopher.frx":091E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGopher.frx":0F18
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGopher.frx":1513
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGopher.frx":1B0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGopher.frx":210B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGopher.frx":2700
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGopher.frx":2D00
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGopher.frx":32FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGopher.frx":38F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGopher.frx":3EEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGopher.frx":44BF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgButtonsDown 
      Left            =   6660
      Top             =   870
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   27
      ImageHeight     =   27
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGopher.frx":4A4B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGopher.frx":5036
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGopher.frx":5621
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGopher.frx":5C15
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGopher.frx":61FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGopher.frx":67EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGopher.frx":6DDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGopher.frx":73CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGopher.frx":79B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGopher.frx":7F9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGopher.frx":855F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Game"
      Height          =   3435
      Left            =   90
      MouseIcon       =   "frmGopher.frx":8AD8
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2430
      Width           =   6345
      Begin VB.Timer timBetweenGopher 
         Enabled         =   0   'False
         Interval        =   750
         Left            =   5880
         Top             =   2940
      End
      Begin VB.Timer timGameTime 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   5430
         Top             =   2940
      End
      Begin VB.Timer timUpFor 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   4980
         Top             =   2940
      End
      Begin MSComctlLib.ImageList imgList 
         Left            =   5700
         Top             =   2340
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   71
         ImageHeight     =   23
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGopher.frx":93A2
               Key             =   "Up"
               Object.Tag             =   "Down"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGopher.frx":A75C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGopher.frx":D2B6
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblTimer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "30"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   180
         TabIndex        =   8
         Top             =   480
         Width           =   525
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Clock"
         ForeColor       =   &H000000FF&
         Height          =   165
         Left            =   270
         TabIndex        =   7
         Top             =   270
         Width           =   375
      End
      Begin VB.Shape Shape2 
         Height          =   555
         Left            =   120
         Top             =   240
         Width           =   645
      End
      Begin VB.Label lblScore 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5700
         TabIndex        =   4
         Top             =   480
         Width           =   435
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Score"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   5730
         TabIndex        =   3
         Top             =   270
         Width           =   435
      End
      Begin VB.Shape Shape1 
         Height          =   555
         Left            =   5640
         Top             =   240
         Width           =   555
      End
      Begin VB.Image imgGDown 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   6
         Left            =   2700
         Top             =   1830
         Width           =   1095
      End
      Begin VB.Image imgGDown 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   5
         Left            =   540
         Top             =   1830
         Width           =   1095
      End
      Begin VB.Image imgGDown 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   4
         Left            =   1650
         Top             =   2910
         Width           =   1095
      End
      Begin VB.Image imgGDown 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   3
         Left            =   3720
         Top             =   2910
         Width           =   1095
      End
      Begin VB.Image imgGDown 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   2
         Left            =   4830
         Top             =   1830
         Width           =   1095
      End
      Begin VB.Image imgGDown 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   1
         Left            =   3720
         Top             =   660
         Width           =   1095
      End
      Begin VB.Image imgGUp 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   795
         Index           =   6
         Left            =   2700
         Top             =   1410
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image imgGUp 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   795
         Index           =   5
         Left            =   540
         Top             =   1410
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image imgGUp 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   795
         Index           =   4
         Left            =   1650
         Top             =   2490
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image imgGUp 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   795
         Index           =   3
         Left            =   3720
         Top             =   2490
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image imgGUp 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   795
         Index           =   2
         Left            =   4830
         Top             =   1410
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image imgGUp 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   795
         Index           =   1
         Left            =   3720
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image imgGDown 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   0
         Left            =   1650
         Top             =   660
         Width           =   1095
      End
      Begin VB.Image imgGUp 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   795
         Index           =   0
         Left            =   1650
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Instructions"
      Height          =   1395
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6375
      Begin VB.Image imgAbout 
         Height          =   480
         Left            =   5760
         Picture         =   "frmGopher.frx":FE10
         Top             =   210
         Width           =   480
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   $"frmGopher.frx":10252
         Height          =   795
         Left            =   360
         TabIndex        =   1
         Top             =   390
         Width           =   5205
      End
   End
End
Attribute VB_Name = "frmGopher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lngSelectedGopher As Long
Private lngNumberHit As Long
Private lngGameTime As Long
Private blnGopherHit As Boolean

Const LENGTHOFGAME = 30

Private Sub GopherUp(IsUp As Boolean, WhichGopher As Long)

imgGUp(WhichGopher).Visible = IsUp
imgGDown(WhichGopher).Visible = Not IsUp

End Sub

Private Sub Form_Load()

LoadButtons 'Load all the button images
InitializeImages 'Initialize the gopher pics
lngSelectedGopher = 0 'Gopher 0 is selected
lngNumberHit = 0 'No gophers hit
lngGameTime = LENGTHOFGAME 'This is the counter that keeps track of game time
SwitchLevel 1 'Set level 1 button
SetLevel 1 'Level 1 is default
timGameTime.Enabled = False 'Disable all timers
timBetweenGopher.Enabled = False
timUpFor.Enabled = False
blnGopherHit = False

End Sub

Private Sub InitializeImages()

Dim i As Long

For i = 0 To imgGDown.Count - 1
    imgGDown.Item(i).Picture = imgList.ListImages(1).Picture
    imgGDown.Item(i).BorderStyle = 0
    
    imgGUp.Item(i).Picture = imgList.ListImages(2).Picture
    imgGUp.Item(i).BorderStyle = 0
Next i

End Sub

Private Sub GetRandomGopher()

Dim lngNewGopher As Long

Do
    lngNewGopher = CLng(Rnd() * 6)
Loop While lngNewGopher = lngSelectedGopher

lngSelectedGopher = lngNewGopher

End Sub

Private Sub StartGame()

timGameTime.Enabled = True
timBetweenGopher.Enabled = True

lblTimer.Caption = "30"
lblScore.Caption = 0

lngGameTime = LENGTHOFGAME
lngNumberHit = 0

blnGopherHit = False

End Sub

Private Sub StopGame()

timGameTime.Enabled = False
timUpFor.Enabled = False
timBetweenGopher.Enabled = False
GopherUp False, lngSelectedGopher
imgButtons.Item(10).Picture = imgButtonsUp.ListImages(10).Picture

End Sub

Private Sub imgAbout_Click()

frmSplash.Show

End Sub

Private Sub imgButtons_Click(Index As Integer)

    Select Case Index
        Case 10
            If Not timGameTime.Enabled = True Then
                imgButtons.Item(10).Picture = imgButtonsDown.ListImages(10).Picture
                StartGame
            End If
        Case 11
            StopGame
        Case 1 To 9
            If Not timGameTime.Enabled = True Then
                SwitchLevel Index
                SetLevel Index
            End If
    End Select
    
End Sub

Private Sub imgButtons_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Select Case Index
    Case 11:
        imgButtons.Item(11).Picture = imgButtonsDown.ListImages(11).Picture
End Select

End Sub

Private Sub imgButtons_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Select Case Index
    Case 11:
        imgButtons.Item(11).Picture = imgButtonsUp.ListImages(11).Picture
End Select

End Sub

Private Sub imgGDown_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If imgGUp(Index).Visible = True And blnGopherHit = False Then
    'Show that the gopher has been hit
    imgGUp.Item(lngSelectedGopher).Picture = imgList.ListImages(3).Picture
    lngNumberHit = lngNumberHit + 1
    lblScore.Caption = lngNumberHit
    blnGopherHit = True
End If

End Sub

Private Sub imgGUp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If imgGUp(Index).Visible = True And blnGopherHit = False Then
    'Show that the gopher has been hit
    imgGUp.Item(lngSelectedGopher).Picture = imgList.ListImages(3).Picture
    lngNumberHit = lngNumberHit + 1
    lblScore.Caption = lngNumberHit
    blnGopherHit = True
End If

End Sub

Private Sub timBetweenGopher_Timer()

timBetweenGopher.Enabled = False
GetRandomGopher
GopherUp True, lngSelectedGopher
timUpFor.Enabled = True
DoEvents

End Sub

Private Sub timGameTime_Timer()

lngGameTime = lngGameTime - 1
lblTimer.Caption = lngGameTime

If lngGameTime = 0 Then
    StopGame
End If

DoEvents

End Sub

Private Sub timUpFor_Timer()

blnGopherHit = False
timBetweenGopher.Enabled = True
imgGUp.Item(lngSelectedGopher).Picture = imgList.ListImages(2).Picture
GopherUp False, lngSelectedGopher
timUpFor.Enabled = False
DoEvents

End Sub

Private Sub LoadButtons()

Dim i As Long

For i = 1 To imgButtons.Count
    imgButtons.Item(i).Picture = imgButtonsUp.ListImages(i).Picture
Next i

End Sub

Private Sub SwitchLevel(Index As Integer)

LoadButtons
imgButtons.Item(Index).Picture = imgButtonsDown.ListImages(Index).Picture

End Sub

Private Sub SetLevel(Index As Integer)

Select Case Index
    Case 1
        timBetweenGopher.Interval = 1000 '600
        timUpFor.Interval = 1000 '600
    Case 2
        timBetweenGopher.Interval = 600
        timUpFor.Interval = 550
    Case 3
        timBetweenGopher.Interval = 600
        timUpFor.Interval = 500
    Case 4
        timBetweenGopher.Interval = 500
        timUpFor.Interval = 500
    Case 5
        timBetweenGopher.Interval = 500
        timUpFor.Interval = 450
    Case 6
        timBetweenGopher.Interval = 500
        timUpFor.Interval = 400
    Case 7
        timBetweenGopher.Interval = 400
        timUpFor.Interval = 400
    Case 8
        timBetweenGopher.Interval = 400
        timUpFor.Interval = 350
    Case 9
        timBetweenGopher.Interval = 400
        timUpFor.Interval = 300
End Select

End Sub
