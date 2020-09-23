VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1995
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   3480
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "V 1.0"
      Height          =   225
      Left            =   3000
      TabIndex        =   3
      Top             =   1740
      Width           =   435
   End
   Begin VB.Image Image1 
      Height          =   765
      Left            =   60
      Picture         =   "frmSplash.frx":000C
      Top             =   360
      Width           =   1065
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "mattvincent@hotmail.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   60
      TabIndex        =   2
      Top             =   1710
      Width           =   1905
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Question or comments email me..."
      Height          =   225
      Left            =   60
      TabIndex        =   1
      Top             =   1440
      Width           =   2565
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Gopher"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   1320
      TabIndex        =   0
      Top             =   330
      Width           =   1665
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1

Private Sub Form_Click()
    Unload Me
    frmGopher.Show
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
    frmGopher.Show
End Sub

Private Sub Form_Load()
    frmGopher.Hide
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label3.ForeColor = vbBlue
End Sub

Private Sub Image1_Click()
    Unload Me
    frmGopher.Show
End Sub

Private Sub Label1_Click()
    Unload Me
    frmGopher.Show
End Sub

Private Sub Label2_Click()
    Unload Me
    frmGopher.Show
End Sub

Private Sub Label3_Click()
    ShellExecute Me.hwnd, vbNullString, "mailto:mattvincent@hotmail.com", vbNullString, "C:\", SW_SHOWNORMAL
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label3.ForeColor = vbRed
End Sub

Private Sub Label4_Click()
    Unload Me
    frmGopher.Show
End Sub
