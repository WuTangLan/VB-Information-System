VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Uposlenici 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Uposlenici"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command10 
      Caption         =   "Trazi"
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   28
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Ocisti"
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   27
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000004&
      Caption         =   "DODAJ"
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   26
      Top             =   3960
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SPREMI"
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   25
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "AZURIRAJ"
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   24
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "OBRISI"
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   23
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Prvi"
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   22
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   21
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   20
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Zadnji"
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   19
      Top             =   3960
      Width           =   615
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   6120
      TabIndex        =   18
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   7
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Top             =   840
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Musko"
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   1920
      Width           =   255
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Zensko"
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Top             =   1920
      Width           =   255
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   1560
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   114098177
      CurrentDate     =   43807
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Radno mjesto"
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4800
      TabIndex        =   17
      Top             =   480
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ime"
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   16
      Top             =   480
      Width           =   285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Prezime"
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   15
      Top             =   840
      Width           =   600
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Datum rodjenja"
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   14
      Top             =   1200
      Width           =   1200
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "JMBG"
      Height          =   195
      Left            =   360
      TabIndex        =   13
      Top             =   1560
      Width           =   435
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Spol"
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   12
      Top             =   1920
      Width           =   345
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Telefon"
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   11
      Top             =   2640
      Width           =   555
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Musko"
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1680
      TabIndex        =   10
      Top             =   1920
      Width           =   510
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Zensko"
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2760
      TabIndex        =   9
      Top             =   1920
      Width           =   540
   End
   Begin VB.Label Adresa 
      AutoSize        =   -1  'True
      Caption         =   "Adresa"
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   2280
      Width           =   525
   End
End
Attribute VB_Name = "Uposlenici"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String

Sub clear()

Text1.Text = ""
Text2.Text = ""
DTPicker1.Value = (12 / 12 / 19)
Text6.Text = ""
Option1.Value = False
Option2.Value = False
Text4.Text = ""
Text3.Text = ""
Text5.Text = ""

End Sub

Private Sub Command1_Click()
clear
rs.AddNew
Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Command9.Enabled = False
End Sub

Private Sub Command10_Click()
rs.Close
rs.Open "Select * from UposleniciTBL where Radno_mjesto='" + Text5.Text + "'", con, adOpenDynamic, adLockPessimistic
If Not rs.EOF Then
display
reload
Else
MsgBox "Nema trazenih rezultata...!!", vbInformation, ""
End If
End Sub

Sub reload()
rs.Close
rs.Open "Select * from UposleniciTBL", con, odOpenStatic, adLockPessimistic
End Sub

Private Sub Command11_Click()
clear
End Sub

Private Sub Command2_Click()
rs.Fields("Ime").Value = Text1.Text
rs.Fields("Prezime").Value = Text2.Text
rs.Fields("Datum_rodjenja").Value = DTPicker1.Value
rs.Fields("JMBG").Value = Text6.Text
If Option1.Value = True Then
rs.Fields("Spol") = Option1.Caption
Else
rs.Fields("Spol") = Option2.Caption
End If
rs.Fields("Adresa") = Text3.Text
rs.Fields("Telefon") = Text4.Text
rs.Fields("Radno_mjesto") = Text5.Text
MsgBox "Podaci su uspjesno snimljeni...!!!", vbInformation
rs.Update

Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Command9.Enabled = True
End Sub

Private Sub Command3_Click()
rs.Fields("Ime").Value = Text1.Text
rs.Fields("Prezime").Value = Text2.Text
rs.Fields("Datum_rodjenja").Value = DTPicker1.Value
rs.Fields("JMBG").Value = Text6.Text
If Option1.Value = True Then
rs.Fields("Spol") = Option1.Caption
Else
rs.Fields("Spol") = Option2.Caption
End If
rs.Fields("Adresa") = Text3.Text
rs.Fields("Telefon") = Text4.Text
rs.Fields("Radno_mjesto") = Text5.Text
MsgBox "Podaci su uspjesno azurirani...!!!", vbInformation
rs.Update
End Sub

Private Sub Command4_Click()
confirm = MsgBox("Da li zelite obrisati uposlenika?", vbYesNo + vbCritical, "Potvrda Brisanja")
If confirm = vbYes Then
rs.Delete adAffectCurrent
MsgBox "Student je uspjesno obrisan!", vbInformation, " Upozorenje"
rs.Update
refreshdata
Else
MsgBox " Brisanje je otkazano..!!!", vbInformation, "Upozorenje"
End If
End Sub

Sub refreshdata()

rs.Close
rs.Open "Select * from UposleniciTBL", con, odOpenStatic, adLockPessimistic
If Not rs.EOF Then
rs.MoveNext
display
Else
MsgBox "Nema rezultata pretrazivanja!!!"
End If

End Sub

Private Sub Command6_Click()
rs.MoveFirst
display
End Sub


Private Sub Command7_Click()
rs.MovePrevious
If rs.BOF Then
rs.MoveLast
display
Else
display
End If
End Sub

Private Sub Command8_Click()
rs.MoveNext
If rs.EOF Then
rs.MoveFirst
display
Else
display
End If
End Sub

Private Sub Command9_Click()
rs.MoveLast
display
End Sub


Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2

con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Asus\Desktop\Projekat\ProfileDB1.mdb;Persist Security Info=False"
rs.Open "Select * from UposleniciTBL", con, adOpenDynamic, adLockPessimistic

display
End Sub

Sub display()
Text1.Text = rs!Ime
Text2.Text = rs!Prezime
DTPicker1.Value = rs!Datum_rodjenja
Text6.Text = rs!JMBG
If rs!Spol = "Musko" Then
Option1.Value = True
Else
Option2.Value = True
End If
Text3.Text = rs!Adresa
Text4.Text = rs!Telefon
Text5.Text = rs!Radno_mjesto
Command2.Enabled = False
End Sub
