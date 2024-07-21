VERSION 5.00
Begin VB.Form KostenVenster 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kosten"
   ClientHeight    =   3612
   ClientLeft      =   1596
   ClientTop       =   1476
   ClientWidth     =   3612
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
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
   ScaleHeight     =   15.05
   ScaleMode       =   4  'Character
   ScaleWidth      =   30.1
   Begin VB.Label TafelgarnituurTotaalVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2280
      LinkTimeout     =   0
      TabIndex        =   20
      Top             =   2520
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label TafelgarnituurLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Tafelgarnituur:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      LinkTimeout     =   0
      TabIndex        =   21
      Top             =   2520
      UseMnemonic     =   0   'False
      Width           =   2055
   End
   Begin VB.Label BittergarnituurLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Bittergarnituur:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      LinkTimeout     =   0
      TabIndex        =   23
      Top             =   2280
      UseMnemonic     =   0   'False
      Width           =   2055
   End
   Begin VB.Label BittergarnituurTotaalVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2280
      LinkTimeout     =   0
      TabIndex        =   22
      Top             =   2280
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label DoorverkoopproductenTotaalVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2280
      LinkTimeout     =   0
      TabIndex        =   19
      Top             =   2040
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label DoorverkoopproductenLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Doorverkoopproducten:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      LinkTimeout     =   0
      TabIndex        =   18
      Top             =   2040
      UseMnemonic     =   0   'False
      Width           =   2055
   End
   Begin VB.Label SoepenEnSausenLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Soepen en Sausen:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      LinkTimeout     =   0
      TabIndex        =   17
      Top             =   1560
      UseMnemonic     =   0   'False
      Width           =   2055
   End
   Begin VB.Label SoepenEnSausenTotaalVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2280
      LinkTimeout     =   0
      TabIndex        =   16
      Top             =   1560
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label VisLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Vis:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      LinkTimeout     =   0
      TabIndex        =   15
      Top             =   840
      UseMnemonic     =   0   'False
      Width           =   2055
   End
   Begin VB.Label VisKostprijsVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2280
      LinkTimeout     =   0
      TabIndex        =   14
      Top             =   840
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label TotaalVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2280
      LinkTimeout     =   0
      TabIndex        =   8
      Top             =   3240
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label PersoneelskostenTotaalVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2280
      LinkTimeout     =   0
      TabIndex        =   13
      Top             =   2760
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label BijproductenTotaalVeld 
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
      Width           =   1215
   End
   Begin VB.Label MeelspijzenTotaalVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2280
      LinkTimeout     =   0
      TabIndex        =   11
      Top             =   1320
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label GroentenTotaalVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2280
      LinkTimeout     =   0
      TabIndex        =   10
      Top             =   1080
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label VleesKostprijsVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2280
      LinkTimeout     =   0
      TabIndex        =   9
      Top             =   600
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label TotaalLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Totaal:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      LinkTimeout     =   0
      TabIndex        =   7
      Top             =   3240
      UseMnemonic     =   0   'False
      Width           =   2055
   End
   Begin VB.Line Scheidingsbalk2 
      X1              =   1
      X2              =   29
      Y1              =   13
      Y2              =   13
   End
   Begin VB.Line Scheidingsbalk1 
      X1              =   1
      X2              =   29
      Y1              =   2
      Y2              =   2
   End
   Begin VB.Label PersoneelskostenLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Personeelskosten:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      LinkTimeout     =   0
      TabIndex        =   6
      Top             =   2760
      UseMnemonic     =   0   'False
      Width           =   2055
   End
   Begin VB.Label BijproductenLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Bijproducten:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      LinkTimeout     =   0
      TabIndex        =   5
      Top             =   1800
      UseMnemonic     =   0   'False
      Width           =   2055
   End
   Begin VB.Label MeelspijzenLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Meelspijzen:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      LinkTimeout     =   0
      TabIndex        =   4
      Top             =   1320
      UseMnemonic     =   0   'False
      Width           =   2055
   End
   Begin VB.Label GroentenLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Groenten:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      LinkTimeout     =   0
      TabIndex        =   3
      Top             =   1080
      UseMnemonic     =   0   'False
      Width           =   2055
   End
   Begin VB.Label VleesLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Vlees:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      LinkTimeout     =   0
      TabIndex        =   2
      Top             =   600
      UseMnemonic     =   0   'False
      Width           =   2055
   End
   Begin VB.Label GerechtVeld 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      LinkTimeout     =   0
      TabIndex        =   1
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   2535
   End
   Begin VB.Label GerechtLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Gerecht:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      LinkTimeout     =   0
      TabIndex        =   0
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "KostenVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze module bevat het kostenvenster.
Option Explicit
'Deze procedure geeft de opdracht om de gerechtkosten te berekenen.
Private Sub Form_Activate()
On Error GoTo Fout

   GerechtBerekenKosten
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure stelt dit venster in.
Private Sub Form_Load()
On Error GoTo Fout

   VensterNaastMenu Me, Me.Name

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

