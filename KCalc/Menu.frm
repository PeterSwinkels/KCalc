VERSION 5.00
Begin VB.Form MenuVenster 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu"
   ClientHeight    =   5052
   ClientLeft      =   312
   ClientTop       =   1560
   ClientWidth     =   2784
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
   ScaleHeight     =   21.05
   ScaleMode       =   4  'Character
   ScaleWidth      =   23.2
   Begin VB.TextBox GerechtnaamVeld 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      LinkTimeout     =   0
      MaxLength       =   255
      MousePointer    =   1  'Arrow
      TabIndex        =   16
      Top             =   4680
      Width           =   2535
   End
   Begin VB.CommandButton CategorieKnop 
      Appearance      =   0  'Flat
      Caption         =   "&Vlees"
      Height          =   246
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2555
   End
   Begin VB.CommandButton CategorieKnop 
      Appearance      =   0  'Flat
      Caption         =   "&Groenten"
      Height          =   246
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2555
   End
   Begin VB.CommandButton CategorieKnop 
      Appearance      =   0  'Flat
      Caption         =   "&Meelspijzen"
      Height          =   246
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   2555
   End
   Begin VB.CommandButton CategorieKnop 
      Appearance      =   0  'Flat
      Caption         =   "&Bijproducten"
      Height          =   246
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   2555
   End
   Begin VB.CommandButton CategorieKnop 
      Appearance      =   0  'Flat
      Caption         =   "&Personeelskosten"
      Height          =   246
      HelpContextID   =   4
      Index           =   9
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   2555
   End
   Begin VB.CommandButton LaadGerechtKnop 
      Appearance      =   0  'Flat
      Caption         =   "&Laad Gerecht"
      Height          =   246
      Left            =   120
      TabIndex        =   12
      Top             =   3360
      Width           =   2555
   End
   Begin VB.CommandButton BewaarGerechtKnop 
      Appearance      =   0  'Flat
      Caption         =   "&Bewaar Gerecht"
      Height          =   246
      Left            =   120
      TabIndex        =   13
      Top             =   3600
      Width           =   2555
   End
   Begin VB.CommandButton BerekenKostenKnop 
      Appearance      =   0  'Flat
      Caption         =   "&Bereken Kosten"
      Height          =   246
      Left            =   120
      TabIndex        =   15
      Top             =   4080
      Width           =   2555
   End
   Begin VB.CommandButton VerwijderGerechtKnop 
      Appearance      =   0  'Flat
      Caption         =   "&Verwijder Gerecht"
      Height          =   246
      Left            =   120
      TabIndex        =   14
      Top             =   3840
      Width           =   2555
   End
   Begin VB.CommandButton NieuwGerechtKnop 
      Appearance      =   0  'Flat
      Caption         =   "&Nieuw Gerecht"
      Height          =   246
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   2555
   End
   Begin VB.CommandButton CategorieKnop 
      Appearance      =   0  'Flat
      Caption         =   "&Vis"
      Height          =   246
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2555
   End
   Begin VB.CommandButton CategorieKnop 
      Appearance      =   0  'Flat
      Caption         =   "&Soepen/Sausen"
      Height          =   246
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   2555
   End
   Begin VB.CommandButton CategorieKnop 
      Appearance      =   0  'Flat
      Caption         =   "&Doorverkoopproducten"
      Height          =   246
      Index           =   6
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   2555
   End
   Begin VB.CommandButton CategorieKnop 
      Appearance      =   0  'Flat
      Caption         =   "&Buffet"
      Height          =   246
      Index           =   10
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   2550
   End
   Begin VB.CommandButton CategorieKnop 
      Appearance      =   0  'Flat
      Caption         =   "&Bittergarnituur"
      Height          =   246
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   2550
   End
   Begin VB.CommandButton CategorieKnop 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Tafelgarnituur"
      Height          =   246
      Index           =   8
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   2550
   End
   Begin VB.Label KiesEenCategorieLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Kies een categorie:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      LinkTimeout     =   0
      TabIndex        =   18
      Top             =   0
      Width           =   1815
   End
   Begin VB.Line Scheidingsbalk1 
      X1              =   0
      X2              =   23
      Y1              =   12.5
      Y2              =   12.5
   End
   Begin VB.Label GerechtLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Gerecht:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      LinkTimeout     =   0
      TabIndex        =   17
      Top             =   4440
      Width           =   855
   End
End
Attribute VB_Name = "MenuVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze module bevat het menuvenster.
Option Explicit

'Deze procedure geeft de opdracht om het kostenvenster te tonen.
Private Sub BerekenKostenKnop_Click()
On Error GoTo Fout

   KostenVenster.WindowState = vbNormal
   KostenVenster.ZOrder
   KostenVenster.Show
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht om het huidige gerecht te bewaren.
Private Sub BewaarGerechtKnop_Click()
On Error GoTo Fout

   BewaarGerecht Gerecht.HuidigBestand
   WerkGerechtenlijstBij
   
   MenuVenster.GerechtNaamVeld.Text = Gerecht.Naam
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht om de geselecteerde productcategorie te tonen.
Private Sub CategorieKnop_Click(Index As Integer)
On Error GoTo Fout
Dim Venster As Form

   Set Venster = Nothing
   
   Select Case Index
      Case 0
         Set Venster = VleesVenster
      Case 1
         Set Venster = VisVenster
      Case 2
         Set Venster = GroentenVenster
      Case 3
         Set Venster = MeelspijzenVenster
      Case 4
         Set Venster = SoepenEnSausenVenster
      Case 5
         Set Venster = BijproductenVenster
      Case 6
         Set Venster = DoorverkoopproductenVenster
      Case 7
         Set Venster = BittergarnituurVenster
      Case 8
         Set Venster = TafelgarnituurVenster
      Case 9
         Set Venster = PersoneelskostenVenster
      Case 10
         Set Venster = BuffetVenster
   End Select
   
   If Not Venster Is Nothing Then
      Venster.WindowState = vbNormal
      Venster.ZOrder
      Venster.Show
   End If
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure stelt dit venster in.
Private Sub Form_Load()
On Error GoTo Fout

   Me.Left = 50
   Me.Top = 50
   
   GerechtNaamVeld.Text = Gerecht.Naam

EindeProcedure:
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure verzoekt de gebruiker indien nodig om aan te geven of het huidige gerecht bewaard moet worden.
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Fout

   If Not Gerecht.WijzigingenBewaard Then
      If BewaarGerechtDialoog() = vbYes Then
         BewaarGerecht Gerecht.HuidigBestand
         WerkGerechtenlijstBij
      End If
   End If
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure legt de door de gebruiker ingevoerde gerechtnaam vast.
Private Sub GerechtnaamVeld_LostFocus()
On Error GoTo Fout

   GerechtNaamVeld.Text = Trim$(GerechtNaamVeld.Text)
   If GerechtNaamVeld.Text = vbNullString Then GerechtNaamVeld.Text = "gerecht " & Gerecht.HuidigNr
   Gerecht.Naam = GerechtNaamVeld.Text
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht om het venster waarin een te laden gerecht geselecteerd kan worden te tonen.
Private Sub LaadGerechtKnop_Click()
On Error GoTo Fout
   
   LaadGerechtVenster.WindowState = vbNormal
   LaadGerechtVenster.ZOrder
   LaadGerechtVenster.Show

EindeProcedure:
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure controleert of het huidige gerecht bewaard moet worden en maakt een nieuw gerecht aan.
Private Sub NieuwGerechtKnop_Click()
On Error GoTo Fout

   If Not Gerecht.WijzigingenBewaard Then
      If BewaarGerechtDialoog() = vbYes Then
         BewaarGerecht Gerecht.HuidigBestand
         WerkGerechtenlijstBij
      End If
   End If

   StelGerechtIn IsNieuwGerecht:=True
   StelGerechtvenstersIn
   
   MenuVenster.GerechtNaamVeld.Text = Gerecht.Naam
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht om het venster waarin een te verwijderen gerecht geselecteerd kan worden te tonen.
Private Sub VerwijderGerechtKnop_Click()
On Error GoTo Fout
   
   VerwijderGerechtVenster.WindowState = vbNormal
   VerwijderGerechtVenster.ZOrder
   VerwijderGerechtVenster.Show

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

