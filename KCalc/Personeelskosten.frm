VERSION 5.00
Begin VB.Form PersoneelskostenVenster 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Personeelskosten"
   ClientHeight    =   3975
   ClientLeft      =   2235
   ClientTop       =   1860
   ClientWidth     =   5775
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   16.563
   ScaleMode       =   4  'Character
   ScaleWidth      =   48.125
   Begin VB.TextBox LoonVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   7
      Left            =   2160
      LinkTimeout     =   0
      MaxLength       =   10
      TabIndex        =   14
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox TijdVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   7
      Left            =   3960
      LinkTimeout     =   0
      MaxLength       =   10
      TabIndex        =   15
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox LoonVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   6
      Left            =   2160
      LinkTimeout     =   0
      MaxLength       =   10
      TabIndex        =   12
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox TijdVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   6
      Left            =   3960
      LinkTimeout     =   0
      MaxLength       =   10
      TabIndex        =   13
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox LoonVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   5
      Left            =   2160
      LinkTimeout     =   0
      MaxLength       =   10
      TabIndex        =   10
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox TijdVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   5
      Left            =   3960
      LinkTimeout     =   0
      MaxLength       =   10
      TabIndex        =   11
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox LoonVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   2160
      LinkTimeout     =   0
      MaxLength       =   10
      TabIndex        =   8
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox TijdVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   3960
      LinkTimeout     =   0
      MaxLength       =   10
      TabIndex        =   9
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox LoonVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   2160
      LinkTimeout     =   0
      MaxLength       =   10
      TabIndex        =   6
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox TijdVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   3960
      LinkTimeout     =   0
      MaxLength       =   10
      TabIndex        =   7
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox LoonVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   2160
      LinkTimeout     =   0
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox TijdVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   3960
      LinkTimeout     =   0
      MaxLength       =   10
      TabIndex        =   5
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox LoonVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   2160
      LinkTimeout     =   0
      MaxLength       =   10
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox TijdVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   3960
      LinkTimeout     =   0
      MaxLength       =   10
      TabIndex        =   3
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox TijdVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   3960
      LinkTimeout     =   0
      MaxLength       =   10
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox LoonVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   2160
      LinkTimeout     =   0
      MaxLength       =   10
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label ExternIILabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Extern II:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      LinkTimeout     =   0
      TabIndex        =   22
      Top             =   3120
      UseMnemonic     =   0   'False
      Width           =   1995
   End
   Begin VB.Label ExternILabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Extern I:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      LinkTimeout     =   0
      TabIndex        =   28
      Top             =   2760
      UseMnemonic     =   0   'False
      Width           =   1995
   End
   Begin VB.Label KeukenhulpLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Keukenhulp:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      LinkTimeout     =   0
      TabIndex        =   27
      Top             =   2400
      UseMnemonic     =   0   'False
      Width           =   1995
   End
   Begin VB.Label LeerlingLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Leerling:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      LinkTimeout     =   0
      TabIndex        =   26
      Top             =   2040
      UseMnemonic     =   0   'False
      Width           =   1995
   End
   Begin VB.Label ProductiekokLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Productiekok:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      LinkTimeout     =   0
      TabIndex        =   25
      Top             =   1680
      UseMnemonic     =   0   'False
      Width           =   1995
   End
   Begin VB.Label ChefDePartiLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Chef de Parti:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      LinkTimeout     =   0
      TabIndex        =   24
      Top             =   1320
      UseMnemonic     =   0   'False
      Width           =   1995
   End
   Begin VB.Label SausChefLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Saus Chef:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      LinkTimeout     =   0
      TabIndex        =   23
      Top             =   960
      UseMnemonic     =   0   'False
      Width           =   1995
   End
   Begin VB.Label TotaalLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Totaal:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      LinkTimeout     =   0
      TabIndex        =   21
      Top             =   3600
      UseMnemonic     =   0   'False
      Width           =   615
   End
   Begin VB.Label TotaalVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2160
      LinkTimeout     =   0
      TabIndex        =   20
      Top             =   3600
      UseMnemonic     =   0   'False
      Width           =   1695
   End
   Begin VB.Line Scheidingsbalk2 
      X1              =   1
      X2              =   47
      Y1              =   14.5
      Y2              =   14.5
   End
   Begin VB.Label TijdInMinutenLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Tijd in minuten:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3960
      LinkTimeout     =   0
      TabIndex        =   19
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   1695
   End
   Begin VB.Label LoonPerUurLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Loon per uur:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2160
      LinkTimeout     =   0
      TabIndex        =   18
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   1695
   End
   Begin VB.Line Scheidingsbalk1 
      X1              =   1
      X2              =   47
      Y1              =   2
      Y2              =   2
   End
   Begin VB.Label WerknemerLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Werknemer:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      LinkTimeout     =   0
      TabIndex        =   17
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   2000
   End
   Begin VB.Label ChefKokLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Chef-kok:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      LinkTimeout     =   0
      TabIndex        =   16
      Top             =   600
      UseMnemonic     =   0   'False
      Width           =   2000
   End
End
Attribute VB_Name = "PersoneelskostenVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
On Error GoTo Fout

   VensterNaastMenu Me, Me.Name
   PersoneelskostenStelVensterIn
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

Private Sub LoonVeld_Change(Index As Integer)
On Error GoTo Fout

   Gerecht.WijzigingenBewaard = False
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

Private Sub LoonVeld_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo Fout

   BewerkGetalVeld LoonVeld(Index), KeyAscii, MAX_LOON, IsBedrag:=True
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

Private Sub LoonVeld_LostFocus(Index As Integer)
On Error GoTo Fout
   
   PersoneelskostenVerwerkInvoer
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

Private Sub TijdVeld_Change(Index As Integer)
On Error GoTo Fout

   Gerecht.WijzigingenBewaard = False
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

Private Sub TijdVeld_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo Fout

   BewerkGetalVeld TijdVeld(Index), KeyAscii, MAX_TIJD
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

Private Sub TijdVeld_LostFocus(Index As Integer)
On Error GoTo Fout

   PersoneelskostenVerwerkInvoer
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

