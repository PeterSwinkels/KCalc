VERSION 5.00
Begin VB.Form VerwijderGerechtVenster 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerecht Verwijderen"
   ClientHeight    =   3372
   ClientLeft      =   1848
   ClientTop       =   2280
   ClientWidth     =   4452
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
   ScaleHeight     =   14.05
   ScaleMode       =   4  'Character
   ScaleWidth      =   37.1
   Begin VB.CommandButton VerwijderGerechtKnop 
      Appearance      =   0  'Flat
      Caption         =   "&Verwijder Gerecht"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.ListBox GerechtLijst 
      Appearance      =   0  'Flat
      Height          =   2520
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label DatumLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Datum:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label DatumVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "VerwijderGerechtVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze module bevat het venster waarin een te verwijderen gerecht geselecteerd wordt.
Option Explicit

'Deze procedure stelt dit venster in.
Private Sub Form_Load()
On Error GoTo Fout
   
   VensterNaastMenu Me, Me.Name
   WerkGerechtenlijstBij
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure werkt de voor het geselecteerde gerechtbestand getoonde informatie bij.
Private Sub GerechtLijst_Click()
On Error GoTo Fout

   DatumVeld.Caption = GerechtBestanden(GerechtLijst.ListIndex).Datum
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht om het geselecteerde gerechtbestand te verwijderen.
Private Sub VerwijderGerechtKnop_Click()
On Error GoTo Fout
Dim Bestand As String

   With GerechtLijst
      If .ListIndex = GEEN_GERECHTNR Then
         MsgBox "Kies een gerecht uit de lijst.", vbInformation
      Else
         If MsgBox("Gerecht verwijderen?", vbQuestion Or vbYesNo Or vbDefaultButton2) = vbYes Then
            Bestand = GerechtBestanden(.ListIndex).Naam
            VerwijderGerecht Bestand, .ListIndex
            DatumVeld.Caption = vbNullString
         End If
      End If
   End With
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

