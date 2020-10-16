VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Profesori 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Profesori"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8850
   BeginProperty Font 
      Name            =   "Malgun Gothic"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   8850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command10 
      Caption         =   "Trazi"
      Height          =   495
      Left            =   6840
      TabIndex        =   32
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Ocisti"
      Height          =   495
      Left            =   5400
      TabIndex        =   31
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Prvi"
      Height          =   495
      Left            =   6120
      TabIndex        =   30
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "<"
      Height          =   495
      Left            =   6720
      TabIndex        =   29
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Caption         =   ">"
      Height          =   495
      Left            =   7320
      TabIndex        =   28
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Zadnji"
      Height          =   495
      Left            =   7920
      TabIndex        =   27
      Top             =   3480
      Width           =   615
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   5520
      TabIndex        =   26
      Text            =   "Odaberi predmet"
      Top             =   1320
      Width           =   3135
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   5520
      TabIndex        =   25
      Text            =   "Odaberi fakultet"
      Top             =   960
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "OBRISI"
      Height          =   495
      Left            =   4200
      TabIndex        =   23
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "AZURIRAJ"
      Height          =   495
      Left            =   2880
      TabIndex        =   22
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SPREMI"
      Height          =   495
      Left            =   1560
      TabIndex        =   21
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000004&
      Caption         =   "DODAJ"
      Height          =   495
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   20
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   5520
      TabIndex        =   19
      Top             =   240
      Width           =   3135
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5520
      TabIndex        =   18
      Text            =   "Odaberi titulu"
      Top             =   600
      Width           =   3135
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1680
      TabIndex        =   17
      Top             =   2040
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Musko"
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   1680
      Width           =   255
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Zensko"
      Height          =   255
      Left            =   3480
      TabIndex        =   13
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1680
      TabIndex        =   12
      Top             =   1320
      Width           =   2055
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   255
      Left            =   1680
      TabIndex        =   11
      Top             =   960
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      _Version        =   393216
      Format          =   113115137
      CurrentDate     =   43818
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1680
      TabIndex        =   10
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1680
      TabIndex        =   9
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Predmeti"
      Height          =   195
      Left            =   4680
      TabIndex        =   24
      Top             =   1320
      Width           =   690
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Musko"
      Height          =   195
      Left            =   1680
      TabIndex        =   16
      Top             =   1680
      Width           =   510
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Zensko"
      Height          =   195
      Left            =   2760
      TabIndex        =   15
      Top             =   1680
      Width           =   540
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Fakultet"
      Height          =   195
      Left            =   4680
      TabIndex        =   8
      Top             =   960
      Width           =   600
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Titula"
      Height          =   195
      Left            =   4680
      TabIndex        =   7
      Top             =   600
      Width           =   420
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Mjesto"
      Height          =   195
      Left            =   4680
      TabIndex        =   6
      Top             =   240
      Width           =   525
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Adresa"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   525
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Spol"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   345
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "JMBG"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Datum rodjenja"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Prezime"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ime"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   285
   End
End
Attribute VB_Name = "Profesori"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String

Private Sub Combo2_Click()
Combo3.clear
If Combo2.Text = "FIT" Then
Combo3.AddItem "Digitalna ekonomija"
Combo3.AddItem "Arhitektura informacionih sistema"
Combo3.AddItem "Informacione tehnologije"
Combo3.AddItem "Programiranje I"
ElseIf Combo2.Text = "Ekonomija" Then
Combo3.AddItem "Osnove ekonomije"
Combo3.AddItem "Ekonomija preduzeca"
Combo3.AddItem "Poslovna informatika"
ElseIf Combo2.Text = "Pravo" Then
Combo3.AddItem "Rimsko pravo"
Combo3.AddItem "Nasljedno pravo"
Combo3.AddItem "Teorija drzave i prava"
Combo3.AddItem "Sociologija"
ElseIf Combo2.Text = "Politehnika" Then
Combo3.AddItem "Matematika I"
Combo3.AddItem "Osnove elektrotehnike"
Combo3.AddItem "Tehnike programiranja"
Combo3.AddItem "Engleski jezik"
ElseIf Combo2.Text = "Mediji i komunikacije" Then
Combo3.AddItem "Film i fotografija"
Combo3.AddItem "Medijska pismenost i kultura"
Combo3.AddItem "Engleski jezik"
End If

End Sub

Sub clear()

Text1.Text = ""
Text2.Text = ""
DTPicker1.Value = (12 / 12 / 19)
Text3.Text = ""
Option1.Value = False
Option2.Value = False
Text4.Text = ""
Text5.Text = ""
Combo1.Text = "Odaberi titulu"
Combo2.Text = "Odaberi fakultet"
Combo3.Text = "Odaberi predmet"

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
rs.Open "Select * from ProfesoriTBL where Ime='" + Text1.Text + "' and Prezime='" + Text2.Text + "'", con, adOpenDynamic, adLockPessimistic
If Not rs.EOF Then
display
reload
Else
MsgBox "Nema trazenih rezultata...!!", vbInformation, ""
End If
End Sub

Sub reload()
rs.Close
rs.Open "Select * from ProfesoriTBL", con, odOpenStatic, adLockPessimistic
End Sub

Private Sub Command11_Click()
clear
End Sub

Private Sub Command2_Click()

rs.Fields("Ime").Value = Text1.Text
rs.Fields("Prezime").Value = Text2.Text
rs.Fields("Datum_rodjenja").Value = DTPicker1.Value
rs.Fields("JMBG").Value = Text3.Text

If Option1.Value = True Then

rs.Fields("Spol") = Option1.Caption

Else

rs.Fields("Spol") = Option2.Caption

End If

rs.Fields("Adresa").Value = Text4.Text
rs.Fields("Mjesto").Value = Text5.Text
rs.Fields("Titula").Value = Combo1.Text
rs.Fields("Predmeti").Value = Combo3.Text
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

Sub refreshdata()
rs.Close
rs.Open "Select * from ProfesoriTBL", con, adOpenStatic, adLockPessimistic
If Not rs.EOF Then
rs.MoveNext
display
Else
MsgBox "Nema rezultata pretrazivanja!!!"
End If
End Sub

Private Sub Command3_Click()
rs.Fields("Ime").Value = Text1.Text
rs.Fields("Prezime").Value = Text2.Text
rs.Fields("Datum_rodjenja").Value = DTPicker1.Value
rs.Fields("JMBG").Value = Text3.Text

If Option1.Value = True Then

rs.Fields("Spol") = Option1.Caption

Else

rs.Fields("Spol") = Option2.Caption

End If

rs.Fields("Adresa").Value = Text4.Text
rs.Fields("Mjesto").Value = Text5.Text
rs.Fields("Titula").Value = Combo1.Text
rs.Fields("Predmeti").Value = Combo3.Text
MsgBox "Podaci su uspjesno azurirani...!!!", vbInformation
rs.Update
End Sub

Private Sub Command4_Click()
confirm = MsgBox("Da li ste sigurni da zelite obrisati profesora?", vbYesNo + vbCritical, "Potvrda Brisanja")
If confirm = vbYes Then
rs.Delete adAffectCurrent
MsgBox "Profesor je uspjesno obrrisan!", vbInformation, "Upozorenje"
rs.Update
refreshdata
Else
MsgBox "Brisanje je otkazano..!!!", vbInformation, "Upozorenje"
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
rs.Open "Select * from ProfesoriTBL", con, adOpenDynamic, adLockPessimistic

Combo1.AddItem "Prof.dr."
Combo1.AddItem "Doc.dr."
Combo2.AddItem "FIT"
Combo2.AddItem "Politehnika"
Combo2.AddItem "Ekonomija"
Combo2.AddItem "Pravo"
Combo2.AddItem "Mediji i komunikacije"

display
End Sub

Sub display()
Text1.Text = rs!Ime
Text2.Text = rs!Prezime
DTPicker1.Value = rs!Datum_rodjenja
Text3.Text = rs!JMBG
If rs!Spol = "Musko" Then
Option1.Value = True
Else
Option2.Value = True
End If
Text4.Text = rs!Adresa
Text5.Text = rs!Mjesto
Combo1.Text = rs!Titula
Combo3.Text = rs!Predmeti
Command2.Enabled = False

End Sub
