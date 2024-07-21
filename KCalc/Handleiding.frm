VERSION 5.00
Begin VB.Form HandleidingVenster 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keuken Calculator - Handleiding "
   ClientHeight    =   3600
   ClientLeft      =   1800
   ClientTop       =   2040
   ClientWidth     =   6000
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
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   15
   ScaleMode       =   4  'Character
   ScaleWidth      =   50
   Begin VB.PictureBox HulptekstVeld 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3600
      Left            =   0
      LinkTimeout     =   0
      ScaleHeight     =   298
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   498
      TabIndex        =   0
      Top             =   0
      Width           =   6000
   End
   Begin VB.Menu OnderwerpenHoofdmenu 
      Caption         =   "&Onderwerpen"
      Begin VB.Menu OnderwerpenMenu 
         Caption         =   ""
         Index           =   1
      End
   End
End
Attribute VB_Name = "HandleidingVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze module bevat het handleidingvenster.
Option Explicit

'Deze procedure stelt dit venster in.
Private Sub Form_Load()
On Error GoTo Fout

   CentreerVenster Me, Me.Name, YAfwijking:=3
   LaadHandleidingonderwerpen

   HulptekstVeld.Cls
   HulptekstVeld.Print "Geen onderwerp gekozen."
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht om het geselecteerde handleidingonderwerp te laden en te tonen.
Private Sub OnderwerpenMenu_Click(Index As Integer)
On Error GoTo Fout

   LaadHandleidingonderwerp CLng(Index), HulptekstVeld
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

