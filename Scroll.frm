VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ScrollBar Changer"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7560
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Scroll.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   7560
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text8 
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   15
      Top             =   3000
      Width           =   4215
   End
   Begin VB.TextBox Text7 
      Height          =   360
      Left            =   2160
      TabIndex        =   14
      Top             =   2600
      Width           =   2175
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      Height          =   360
      Left            =   2160
      TabIndex        =   10
      Top             =   1760
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   920
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   2160
      TabIndex        =   2
      Top             =   80
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4320
      Left            =   4440
      Picture         =   "Scroll.frx":0442
      ScaleHeight     =   4260
      ScaleWidth      =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   3060
      Begin VB.CommandButton Command1 
         Caption         =   "Generate Code"
         Height          =   375
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   1815
      End
   End
   Begin VB.Label Label7 
      Caption         =   "&3D Light Color:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "&DarkShadow Color"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2220
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Highlight Color:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "&Face Color:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1400
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "&Shadow Color:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "&Track Color:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   520
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "&Arrow Color:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Code As String

Private Sub Command1_Click()
Code = "<STYLE TYPE=""text/css""><!--" & vbCrLf _
& "BODY {" & vbCrLf _
& "scrollbar-arrow-color:" & Text1.Text & ";" & vbCrLf _
& "scrollbar-track-color:" & Text2.Text & ";" & vbCrLf _
& "scrollbar-shadow-color:" & Text3.Text & ";" & vbCrLf _
& "scrollbar-face-color:" & Text4.Text & ";" & vbCrLf _
& "scrollbar-highlight-color;" & Text5.Text & ";" & vbCrLf _
& "scrollbar-darkshadow-color:" & Text6.Text & ";" & vbCrLf _
& "scrollbar-3dlight-color:" & Text7.Text & ";" & vbCrLf _
& "}" & vbCrLf _
& "//--></STYLE>"

Text8.Text = ""
Text8.Text = Code
End Sub

Private Sub Form_Load()
MsgBox "Works Only in IE 5.5 or Higher", vbInformation, "Works"
End Sub

