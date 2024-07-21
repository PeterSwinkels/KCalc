VERSION 5.00
Begin VB.Form DoorverkoopproductenVenster 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Doorverkoopproducten"
   ClientHeight    =   2175
   ClientLeft      =   1545
   ClientTop       =   2130
   ClientWidth     =   6615
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
   ScaleWidth      =   55.125
   Begin VB.TextBox PrijsVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   2280
      LinkTimeout     =   0
      MaxLength       =   10
      TabIndex        =   9
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox PrijsVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   2280
      LinkTimeout     =   0
      MaxLength       =   10
      TabIndex        =   5
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox PrijsVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   2280
      LinkTimeout     =   0
      MaxLength       =   10
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox HoeveelheidVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   3720
      LinkTimeout     =   0
      MaxLength       =   10
      TabIndex        =   10
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox HoeveelheidVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   3720
      LinkTimeout     =   0
      MaxLength       =   10
      TabIndex        =   6
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox HoeveelheidVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   3720
      LinkTimeout     =   0
      MaxLength       =   10
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.ComboBox ProductenLijst 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   2055
   End
   Begin VB.ComboBox ProductenLijst 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   2055
   End
   Begin VB.ComboBox ProductenLijst 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label KostPrijsVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   2
      Left            =   5160
      LinkTimeout     =   0
      TabIndex        =   11
      Top             =   1320
      UseMnemonic     =   0   'False
      Width           =   1335
   End
   Begin VB.Label KostPrijsVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   1
      Left            =   5160
      LinkTimeout     =   0
      TabIndex        =   7
      Top             =   960
      UseMnemonic     =   0   'False
      Width           =   1335
   End
   Begin VB.Label KostPrijsVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   0
      Left            =   5160
      LinkTimeout     =   0
      TabIndex        =   3
      Top             =   600
      UseMnemonic     =   0   'False
      Width           =   1335
   End
   Begin VB.Label TotaalVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2280
      LinkTimeout     =   0
      TabIndex        =   12
      Top             =   1800
      UseMnemonic     =   0   'False
      Width           =   1335
   End
   Begin VB.Label TotaalLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Totaal:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      LinkTimeout     =   0
      TabIndex        =   17
      Top             =   1800
      UseMnemonic     =   0   'False
      Width           =   615
   End
   Begin VB.Label KostprijsLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Kostprijs:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5160
      LinkTimeout     =   0
      TabIndex        =   13
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   1335
   End
   Begin VB.Label PrijsLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Prijs:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2280
      LinkTimeout     =   0
      TabIndex        =   14
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   1335
   End
   Begin VB.Label HoeveelheidLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Hoeveelheid:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3720
      LinkTimeout     =   0
      TabIndex        =   15
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   1335
   End
   Begin VB.Line Scheidingsbalk2 
      X1              =   1
      X2              =   54
      Y1              =   7
      Y2              =   7
   End
   Begin VB.Line Scheidingsbalk1 
      X1              =   1
      X2              =   54
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
      LinkTimeout     =   0
      TabIndex        =   16
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   2055
   End
End
Attribute VB_Name = "DoorverkoopproductenVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze module bevat een productenvenster.
Option Explicit

'Deze procedure stelt dit venster in.
Private Sub Form_Load()
On Error GoTo Fout

   BBDGMSTStelVensterIn Doorverkoopproducten, Me, Me.Name

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure past de status van het huidige gerecht aan.
Private Sub HoeveelheidVeld_Change(Index As Integer)
On Error GoTo Fout
   
   Gerecht.WijzigingenBewaard = False

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure legt de geselecteerde productrij vast.
Private Sub HoeveelheidVeld_GotFocus(Index As Integer)
On Error GoTo Fout

   Doorverkoopproducten.Productrij = Index
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure verwerkt de invoer van de gebruiker.
Private Sub HoeveelheidVeld_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo Fout

   BewerkGetalVeld HoeveelheidVeld(Doorverkoopproducten.Productrij), KeyAscii, MAX_HOEVEELHEID
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure verwijdert indien gewenst de inhoud van het geselecteerde veld.
Private Sub HoeveelheidVeld_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo Fout

   If KeyCode = vbKeyDelete Then HoeveelheidVeld(Doorverkoopproducten.Productrij).Text = vbNullString

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure legt de opgegeven invoer vast wanneer het veld niet langer geselecteerd is.
Private Sub HoeveelheidVeld_LostFocus(Index As Integer)
On Error GoTo Fout

   BBDGMSTVerwerkInvoer Doorverkoopproducten, Me

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure toont het meeteenhedenmenu.
Private Sub PrijsVeld_DblClick(Index As Integer)
On Error GoTo Fout

   Doorverkoopproducten.Meeteenheid(Doorverkoopproducten.Productrij) = SelecteerMeeteenheid(Doorverkoopproducten.Meeteenheid(Doorverkoopproducten.Productrij), DoorverkoopproductenVenster)
      
   BBDGMSTVerwerkInvoer Doorverkoopproducten, Me

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure legt de geselecteerde productrij vast.
Private Sub PrijsVeld_GotFocus(Index As Integer)
On Error GoTo Fout

   Doorverkoopproducten.Productrij = Index

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure verwerkt de invoer van de gebruiker.
Private Sub PrijsVeld_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo Fout

   BewerkGetalVeld PrijsVeld(Doorverkoopproducten.Productrij), KeyAscii, MAX_PRIJS, IsBedrag:=True

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure verwijdert indien gewenst de inhoud van het geselecteerde veld.
Private Sub PrijsVeld_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo Fout

   If KeyCode = vbKeyDelete Then PrijsVeld(Doorverkoopproducten.Productrij).Text = vbNullString

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure legt de opgegeven invoer vast wanneer het veld niet langer geselecteerd is.
Private Sub PrijsVeld_LostFocus(Index As Integer)
On Error GoTo Fout

   BBDGMSTVerwerkInvoer Doorverkoopproducten, Me

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure blokkeert invoer van er geen product geselecteerd is.
Private Sub ProductenLijst_Change(Index As Integer)
On Error GoTo Fout
   
   If Doorverkoopproducten.Selectie(Doorverkoopproducten.Productrij) = 0 Then ProductenLijst(Doorverkoopproducten.Productrij).Text = GEEN_PRODUCT_TEKST
   Gerecht.WijzigingenBewaard = False
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure werkt de invoervelden bij wanneer een product geselecteerd wordt.
Private Sub ProductenLijst_Click(Index As Integer)
On Error GoTo Fout

   With Doorverkoopproducten
      .Productrij = Index
      HoeveelheidVeld(.Productrij).Enabled = (ProductenLijst(.Productrij).ListIndex > 0)
      PrijsVeld(.Productrij).Enabled = HoeveelheidVeld(.Productrij).Enabled
      .Selectie(.Productrij) = ProductenLijst(.Productrij).ListIndex
      HoeveelheidVeld(.Productrij).Text = .Hoeveelheid(.Selectie(.Productrij), .Productrij)
      PrijsVeld(.Productrij).Text = .Prijs(.Selectie(.Productrij), .Productrij)
   End With
   
   BBDGMSTVerwerkInvoer Doorverkoopproducten, Me

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure legt de geselecteerde productrij vast.
Private Sub ProductenLijst_GotFocus(Index As Integer)
On Error GoTo Fout

   Doorverkoopproducten.Productrij = Index

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure verwerkt de invoer van de gebruiker.
Private Sub ProductenLijst_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo Fout

   BewerkProductVeld ProductenLijst(Doorverkoopproducten.Productrij), KeyAscii, MAX_PRODUCTNAAM_LENGTE

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure verwijdert indien gewenst de inhoud van het geselecteerde veld.
Private Sub ProductenLijst_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo Fout

   If KeyCode = vbKeyDelete Then ProductenLijst(Doorverkoopproducten.Productrij).Text = vbNullString

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure legt de opgegeven invoer vast wanneer het veld niet langer geselecteerd is.
Private Sub ProductenLijst_LostFocus(Index As Integer)
On Error GoTo Fout

   With Doorverkoopproducten
      .Product(.Selectie(.Productrij), .Productrij) = BBDGMSTVVControleerProduct(ProductenLijst(.Productrij), .Selectie(.Productrij))
      ProductenLijst(.Productrij).List(.Selectie(.Productrij)) = .Product(.Selectie(.Productrij), .Productrij)
      ProductenLijst(.Productrij).Text = .Product(.Selectie(.Productrij), .Productrij)
   End With
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

