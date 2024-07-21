VERSION 5.00
Begin VB.Form InformatieVenster 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informatie"
   ClientHeight    =   1332
   ClientLeft      =   4440
   ClientTop       =   2088
   ClientWidth     =   2292
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
   ScaleHeight     =   5.55
   ScaleMode       =   4  'Character
   ScaleWidth      =   19.1
   Begin VB.CommandButton SluitenKnop 
      Appearance      =   0  'Flat
      Caption         =   "&Sluiten"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label InformatieLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      LinkTimeout     =   0
      TabIndex        =   1
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   2055
   End
End
Attribute VB_Name = "InformatieVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze module bevat het programmainformatie venster.
Option Explicit

'Deze procedure stelt dit venster in.
Private Sub Form_Load()
On Error GoTo Fout
 
   CentreerVenster Me, Me.Name
   InformatieLabel.Caption = Programmainformatie()
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure sluit dit venster.
Private Sub SluitenKnop_Click()
On Error GoTo Fout

   Unload Me

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

