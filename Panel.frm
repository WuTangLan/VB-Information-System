VERSION 5.00
Begin VB.Form Panel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dobrodosli"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   11535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "UPOSLENICI"
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8640
      TabIndex        =   4
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "PROFESORI"
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6000
      TabIndex        =   3
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ASISTENTI"
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3360
      TabIndex        =   2
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "STUDENTI"
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      TabIndex        =   1
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Dobrodosli!"
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   10695
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F1F0EC&
      BackStyle       =   1  'Opaque
      Height          =   4095
      Left            =   120
      Top             =   120
      Width           =   11295
   End
End
Attribute VB_Name = "Panel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Studenti.Show
Panel.Hide
End Sub

Private Sub Command2_Click()
Asistenti.Show
End Sub

Private Sub Command3_Click()
Profesori.Show
End Sub

Private Sub Command4_Click()
Uposlenici.Show
End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
End Sub
