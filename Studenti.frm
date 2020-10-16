VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Studenti 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Student"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8415
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command12 
      Caption         =   "Command12"
      Height          =   495
      Left            =   7680
      TabIndex        =   38
      Top             =   120
      Width           =   495
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
      Left            =   4200
      TabIndex        =   37
      Top             =   1080
      Width           =   1335
   End
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
      Left            =   4200
      TabIndex        =   36
      Top             =   480
      Width           =   1335
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
      Left            =   7560
      TabIndex        =   34
      Top             =   5160
      Width           =   615
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
      Left            =   6960
      TabIndex        =   33
      Top             =   5160
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
      Left            =   6360
      TabIndex        =   32
      Top             =   5160
      Width           =   495
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
      Left            =   5760
      TabIndex        =   31
      Top             =   5160
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7560
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      TabIndex        =   29
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "UCITAJ SLIKU"
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
      Left            =   6240
      TabIndex        =   28
      Top             =   2880
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   255
      Left            =   1680
      TabIndex        =   27
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
      Format          =   113377281
      CurrentDate     =   43807
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      TabIndex        =   26
      Text            =   "Odaberi semestar"
      Top             =   4320
      Width           =   1935
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      TabIndex        =   25
      Text            =   "Odaberi smjer"
      Top             =   3960
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      TabIndex        =   24
      Text            =   "Odaberi fakultet"
      Top             =   3600
      Width           =   1935
   End
   Begin VB.TextBox Text5 
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
      TabIndex        =   23
      Top             =   3240
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
      TabIndex        =   22
      Top             =   2640
      Width           =   1935
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
      TabIndex        =   21
      Top             =   2280
      Width           =   1935
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
      TabIndex        =   20
      Top             =   1920
      Width           =   255
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
      TabIndex        =   17
      Top             =   1920
      Width           =   255
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
      TabIndex        =   16
      Top             =   840
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
      TabIndex        =   15
      Top             =   480
      Width           =   1935
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
      Left            =   3840
      TabIndex        =   14
      Top             =   5160
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
      Left            =   2640
      TabIndex        =   13
      Top             =   5160
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
      Left            =   1440
      TabIndex        =   12
      Top             =   5160
      Width           =   1095
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
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   11
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   6240
      ScaleHeight     =   2.937
      ScaleMode       =   0  'User
      ScaleWidth      =   2.937
      TabIndex        =   2
      Top             =   840
      Width           =   1725
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Slika"
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
      Left            =   6240
      TabIndex        =   35
      Top             =   480
      Width           =   360
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
      TabIndex        =   30
      Top             =   2280
      Width           =   525
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
      TabIndex        =   19
      Top             =   1920
      Width           =   540
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
      TabIndex        =   18
      Top             =   1920
      Width           =   510
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Semestar"
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
      TabIndex        =   10
      Top             =   4320
      Width           =   705
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Smjer"
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
      TabIndex        =   9
      Top             =   3960
      Width           =   435
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Fakultet"
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
      Top             =   3600
      Width           =   600
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Broj indexa"
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
      TabIndex        =   7
      Top             =   3240
      Width           =   855
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
      TabIndex        =   6
      Top             =   2640
      Width           =   555
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
      TabIndex        =   5
      Top             =   1920
      Width           =   345
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "JMBG"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   435
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
      TabIndex        =   3
      Top             =   1200
      Width           =   1200
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
      TabIndex        =   1
      Top             =   840
      Width           =   600
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
      TabIndex        =   0
      Top             =   480
      Width           =   285
   End
End
Attribute VB_Name = "Studenti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String

Private Sub Combo1_Click()

Combo2.clear
If Combo1.Text = "FIT" Then
Combo2.AddItem "Softverski inzenjering"
Combo2.AddItem "Racunarski sistemi i mreze"
Combo2.AddItem "Informacione tehnologije"

ElseIf Combo1.Text = "Politehnika" Then
Combo2.AddItem "Masinstvo"
Combo2.AddItem "Telekomunikacije"
Combo2.AddItem "Energetika"

ElseIf Combo1.Text = "Ekonomija" Then
Combo2.AddItem "Bankarstvo"
Combo2.AddItem "Finansije"
Combo2.AddItem "Racunovodstvo"

ElseIf Combo1.Text = "Pravo" Then
Combo2.AddItem "Opce pravo"
Combo2.AddItem "Rimsko pravo"
Combo2.AddItem "Sudsko odlucivanje"

ElseIf Combo1.Text = "Mediji i komunikacije" Then
Combo2.AddItem "Novinarstvo"

Else

End If

End Sub

Private Sub Command1_Click()

clear
rs.AddNew
' Ako se klikne na tipku DODAJ vise od jednom dodje do greske zato se na ostale tipke ne moze
' kliknuti nego samo na SPREMI
Command2.Enabled = True
Command1.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Command9.Enabled = False
Command3.Enabled = False
Command4.Enabled = False

End Sub

Sub clear()

Text5.Text = ""
Text1.Text = ""
Text2.Text = ""
DTPicker1.Value = (21 / 12 / 19)
Text6.Text = ""
Option1.Value = False
Option2.Value = False
Combo1.Text = "Odaberi fakultet"
Combo2.Text = "Odaberi smjer"
Combo3.Text = "Odaberi semestar"
Text3.Text = ""
Text4.Text = ""
Picture1.Picture = LoadPicture("")

End Sub

Private Sub Command10_Click()
rs.Close
rs.Open "Select * from StudentiTBL where Broj_indeksa='" + Text5.Text + "'", con, adOpenDynamic, adLockPessimistic
If Not rs.EOF Then
display
reload
Else
MsgBox "Nema trazenih rezultata...!!", vbInformation, ""
End If
End Sub

Sub reload()
rs.Close
rs.Open "Select * from StudentiTBL", con, odOpenStatic, adLockPessimistic
End Sub


Private Sub Command11_Click()
clear
End Sub

Private Sub Command12_Click()
Unload Me
Panel.Show
End Sub

Private Sub Command2_Click()

rs.Fields("Broj_indeksa").Value = Text5.Text
rs.Fields("Ime").Value = Text1.Text
rs.Fields("Prezime").Value = Text2.Text
rs.Fields("Datum_rodjenja").Value = DTPicker1.Value
rs.Fields("JMBG").Value = Text6.Text

If Option1.Value = True Then

rs.Fields("Spol") = Option1.Caption

Else

rs.Fields("Spol") = Option2.Caption

End If

rs.Fields("Fakultet").Value = Combo1.Text
rs.Fields("Smjer").Value = Combo2.Text
rs.Fields("Semestar").Value = Combo3.Text
rs.Fields("Adresa").Value = Text3.Text
rs.Fields("Telefon").Value = Text4.Text
rs.Fields("Slika").Value = str
MsgBox "Podaci su uspjesno snimljeni...!!!", vbInformation
rs.Update
' Kad se podaci spreme onda se vise ne moze kliknuti na SPREMI
Command1.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Command9.Enabled = True
Command3.Enabled = True
Command4.Enabled = True

End Sub

Private Sub Command3_Click()
rs.Fields("Broj_indeksa").Value = Text5.Text
rs.Fields("Ime").Value = Text1.Text
rs.Fields("Prezime").Value = Text2.Text
rs.Fields("Datum_rodjenja").Value = DTPicker1.Value
rs.Fields("JMBG").Value = Text6.Text

If Option1.Value = True Then

rs.Fields("Spol") = Option1.Caption

Else

rs.Fields("Spol") = Option2.Caption

End If

rs.Fields("Fakultet").Value = Combo1.Text
rs.Fields("Smjer").Value = Combo2.Text
rs.Fields("Semestar").Value = Combo3.Text
rs.Fields("Adresa").Value = Text3.Text
rs.Fields("Telefon").Value = Text4.Text
' Ako varijabla str nije prazna to znaci da je slika izabrana i stavlja se u bazu,
' u suprotnom se nece mijenjati polje u bazi
If str <> "" Then
rs.Fields("Slika").Value = str
End If
MsgBox "Podaci su uspjesno azurirani...!!!", vbInformation
rs.Update
' Svaki put kad se azuriraju podaci str varijabla je prazna
str = ""
End Sub

Private Sub Command4_Click()
confirm = MsgBox("Da li zelite obrisati studenta?", vbYesNo + vbCritical, "Potvrda Brisanja")
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
rs.Open "Select * from StudentiTBL", con, odOpenStatic, adLockPessimistic
If Not rs.EOF Then
rs.MoveNext
display
Else
MsgBox "Nema rezultata pretrazivanja!!!"
End If

End Sub

Private Sub Command5_Click()

CommonDialog1.ShowOpen
CommonDialog1.Filter = "Jpeg|*.jpg"
str = CommonDialog1.FileName
Picture1.Picture = LoadPicture(str)

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
If Not rs.EOF Then
display
Else
rs.MoveFirst
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
rs.Open "Select * from StudentiTBL", con, adOpenDynamic, adLockPessimistic

Combo1.AddItem "FIT"
Combo1.AddItem "Politehnika"
Combo1.AddItem "Ekonomija"
Combo1.AddItem "Pravo"
Combo1.AddItem "Mediji i komunikacije"
Combo3.AddItem "Semestar I"
Combo3.AddItem "Semestar II"
Combo3.AddItem "Semestar III"
Combo3.AddItem "Semestar IV"
Combo3.AddItem "Semestar V"
Combo3.AddItem "Semestar VI"
Combo3.AddItem "Semestar VII"
Combo3.AddItem "Semestar VIII"

display
End Sub

Sub display()
Text5.Text = rs!Broj_indeksa
Text1.Text = rs!Ime
Text2.Text = rs!Prezime
DTPicker1.Value = rs!Datum_rodjenja
If rs!Spol = "Musko" Then
Option1.Value = True
Else
Option2.Value = True
End If
Combo1.Text = rs!Fakultet
Combo2.Text = rs!Smjer
Combo3.Text = rs!Semestar
Text3.Text = rs!Adresa
Text4.Text = rs!Telefon
Text6.Text = rs!JMBG
Picture1.Picture = LoadPicture(rs!Slika)
' Prilikom prikazivanja recorda iz baze tipka SPREMI je disabled zato sto se ti podaci sad mogu samo
' azurirati ili izbrisati
Command2.Enabled = False
End Sub
