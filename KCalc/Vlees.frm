VERSION 5.00
Begin VB.Form VleesVenster 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vlees"
   ClientHeight    =   2175
   ClientLeft      =   3510
   ClientTop       =   2100
   ClientWidth     =   4215
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9.063
   ScaleMode       =   4  'Character
   ScaleWidth      =   35.125
   Begin VB.TextBox HoeveelheidVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      LinkTimeout     =   0
      MaxLength       =   10
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox PrijsVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      LinkTimeout     =   0
      MaxLength       =   10
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
   Begin VB.ComboBox ProductenLijst 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox SnijverliesVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      LinkTimeout     =   0
      MaxLength       =   10
      TabIndex        =   3
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Line Scheidingsbalk2 
      X1              =   1
      X2              =   34
      Y1              =   7
      Y2              =   7
   End
   Begin VB.Label KostprijsVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2160
      LinkTimeout     =   0
      TabIndex        =   9
      Top             =   1800
      UseMnemonic     =   0   'False
      Width           =   1935
   End
   Begin VB.Line Scheidingsbalk1 
      X1              =   1
      X2              =   34
      Y1              =   2
      Y2              =   2
   End
   Begin VB.Label ProductLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Product:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1995
   End
   Begin VB.Label SnijverliesLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Snijverlies %:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      LinkTimeout     =   0
      TabIndex        =   7
      Top             =   1320
      UseMnemonic     =   0   'False
      Width           =   1995
   End
   Begin VB.Label KostprijsLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Kostprijs:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      LinkTimeout     =   0
      TabIndex        =   6
      Top             =   1800
      UseMnemonic     =   0   'False
      Width           =   1995
   End
   Begin VB.Label HoeveelheidLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Hoeveelheid:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      LinkTimeout     =   0
      TabIndex        =   5
      Top             =   960
      UseMnemonic     =   0   'False
      Width           =   1995
   End
   Begin VB.Label PrijsLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Prijs:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      LinkTimeout     =   0
      TabIndex        =   4
      Top             =   600
      UseMnemonic     =   0   'False
      Width           =   1995
   End
End
Attribute VB_Name = "VleesVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze module bevat een productenvenster.
Option Explicit

'Deze procedure stelt dit venster in.
Private Sub Form_Load()
On Error GoTo Fout
   
   VVStelVensterIn Vlees, Me, Me.Name
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure past de status van het huidige gerecht aan.
Private Sub HoeveelheidVeld_Change()
On Error GoTo Fout

   Gerecht.WijzigingenBewaard = False
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure verwerkt de invoer van de gebruiker.
Private Sub HoeveelheidVeld_KeyPress(KeyAscii As Integer)
On Error GoTo Fout

   BewerkGetalVeld HoeveelheidVeld, KeyAscii, MAX_HOEVEELHEID
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure verwijdert indien gewenst de inhoud van het geselecteerde veld.
Private Sub HoeveelheidVeld_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Fout

   If KeyCode = vbKeyDelete Then HoeveelheidVeld.Text = vbNullString

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure legt de opgegeven invoer vast wanneer het veld niet langer geselecteerd is.
Private Sub HoeveelheidVeld_LostFocus()
On Error GoTo Fout

   VVVerwerkInvoer Vlees, Me
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure toont het meeteenhedenmenu.
Private Sub PrijsVeld_DblClick()
On Error GoTo Fout

   Vlees.Meeteenheid = SelecteerMeeteenheid(Vlees.Meeteenheid, VleesVenster)
   
   VVVerwerkInvoer Vlees, Me
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure verwerkt de invoer van de gebruiker.
Private Sub PrijsVeld_KeyPress(KeyAscii As Integer)
On Error GoTo Fout

   BewerkGetalVeld PrijsVeld, KeyAscii, MAX_PRIJS, IsBedrag:=True
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure verwijdert indien gewenst de inhoud van het geselecteerde veld.
Private Sub PrijsVeld_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Fout

   If KeyCode = vbKeyDelete Then PrijsVeld.Text = vbNullString

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure legt de opgegeven invoer vast wanneer het veld niet langer geselecteerd is.
Private Sub PrijsVeld_LostFocus()
On Error GoTo Fout

   VVVerwerkInvoer Vlees, Me
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure past de status van het huidige gerecht aan.
Private Sub ProductenLijst_Change()
On Error GoTo Fout
      
   If Vlees.Selectie = 0 Then ProductenLijst.Text = GEEN_PRODUCT_TEKST
   Gerecht.WijzigingenBewaard = False
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure werkt de invoervelden bij wanneer een product geselecteerd wordt.
Private Sub ProductenLijst_Click()
On Error GoTo Fout

   HoeveelheidVeld.Enabled = (ProductenLijst.ListIndex > 0)
   PrijsVeld.Enabled = HoeveelheidVeld.Enabled
   SnijverliesVeld.Enabled = HoeveelheidVeld.Enabled
   
   With Vlees
      .Selectie = ProductenLijst.ListIndex
      PrijsVeld.Text = .Prijs(.Selectie)
      HoeveelheidVeld.Text = .Hoeveelheid(.Selectie)
      SnijverliesVeld.Text = .Snijverlies(.Selectie)
   End With
   
   VVVerwerkInvoer Vlees, Me
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure verwerkt de invoer van de gebruiker.
Private Sub ProductenLijst_KeyPress(KeyAscii As Integer)
On Error GoTo Fout

   BewerkProductVeld ProductenLijst, KeyAscii, MAX_PRODUCTNAAM_LENGTE

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure verwijdert indien gewenst de inhoud van het geselecteerde veld.
Private Sub ProductenLijst_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Fout

   If KeyCode = vbKeyDelete Then ProductenLijst.Text = vbNullString

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure legt de opgegeven invoer vast wanneer het veld niet langer geselecteerd is.
Private Sub ProductenLijst_LostFocus()
On Error GoTo Fout

   With Vlees
      .Product(.Selectie) = BBDGMSTVVControleerProduct(ProductenLijst, .Selectie)
      ProductenLijst.List(.Selectie) = .Product(.Selectie)
      ProductenLijst.Text = .Product(.Selectie)
   End With
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure past de status van het huidige gerecht aan.
Private Sub SnijverliesVeld_Change()
On Error GoTo Fout

   Gerecht.WijzigingenBewaard = False
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure verwerkt de invoer van de gebruiker.
Private Sub SnijverliesVeld_KeyPress(KeyAscii As Integer)
On Error GoTo Fout

   BewerkGetalVeld SnijverliesVeld, KeyAscii, MAX_PERCENTAGE

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure verwijdert indien gewenst de inhoud van het geselecteerde veld.
Private Sub SnijverliesVeld_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Fout

   If KeyCode = vbKeyDelete Then SnijverliesVeld.Text = vbNullString

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure legt de opgegeven invoer vast wanneer het veld niet langer geselecteerd is.
Private Sub SnijverliesVeld_LostFocus()
On Error GoTo Fout

   VVVerwerkInvoer Vlees, Me
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

