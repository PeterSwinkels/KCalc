VERSION 5.00
Begin VB.Form BackupVenster 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup"
   ClientHeight    =   1212
   ClientLeft      =   2976
   ClientTop       =   1848
   ClientWidth     =   3972
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
   ScaleHeight     =   5.05
   ScaleMode       =   4  'Character
   ScaleWidth      =   33.1
   Begin VB.CommandButton ZetBackupTerugKnop 
      Appearance      =   0  'Flat
      Caption         =   "&Zet Backup Terug"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox BackuplocatieVeld 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3735
   End
   Begin VB.CommandButton MaakBackupKnop 
      Appearance      =   0  'Flat
      Caption         =   "&Maak Backup"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label BackuplocatieLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Backuplocatie:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      LinkTimeout     =   0
      TabIndex        =   3
      Top             =   600
      UseMnemonic     =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "BackupVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze module bevat het backupvenster.
Option Explicit

'Deze procedure stelt het huidige station in als backuplocatie indien er geen locatie is opgegeven.
Private Sub BackuplocatieVeld_LostFocus()
On Error GoTo Fout

   If BackuplocatieVeld.Text = vbNullString Then
      BackuplocatieVeld.Text = VoegBackslashToe(Stationsletter(CurDir$()))
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

   CentreerVenster Me, Me.Name
   BackuplocatieVeld.Text = Backuplocatie()

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht om een backup te maken.
Private Sub MaakBackupKnop_Click()
On Error GoTo Fout
Dim DataDoel As String
Dim DataLocatie As String

   DataDoel = VoegBackslashToe(BackuplocatieVeld.Text)
   DataLocatie = VoegBackslashToe(App.Path) & Mid$(DATA_MAP, 3)
      
   BewaarBackupLocatie DataDoel
   BewaarProducten
   BewaarBuffet
   
   KopieerBackupbestanden DataLocatie, DataDoel

EindeProcedure:
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht om een backup terug te zetten.
Private Sub ZetBackupTerugKnop_Click()
On Error GoTo Fout
Dim DataDoel As String
Dim DataLocatie As String
Dim Keuze As Long

   Keuze = MsgBox("De huidige gegevens zullen worden verwijderd." & vbCr & "Backup terug zetten?", vbQuestion Or vbYesNo Or vbDefaultButton1)
   If Keuze = vbYes Then
      DataDoel = VoegBackslashToe(App.Path) & Mid$(DATA_MAP, 3)
      DataLocatie = BackuplocatieVeld.Text
      
      KopieerBackupbestanden DataLocatie, DataDoel
      
      StelProgrammaIn
      StelGerechtIn
      GerechtBerekenKosten
      LaadProducten
      LaadBuffet
      WerkGerechtenlijstBij
   End If

EindeProcedure:
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

