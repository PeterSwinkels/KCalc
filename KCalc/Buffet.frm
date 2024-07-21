VERSION 5.00
Begin VB.Form BuffetVenster 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buffet"
   ClientHeight    =   2175
   ClientLeft      =   2055
   ClientTop       =   2310
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
   Begin VB.PictureBox Tabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      ScaleHeight     =   4.063
      ScaleMode       =   4  'Character
      ScaleWidth      =   53.125
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   600
      Width           =   6375
      Begin VB.PictureBox TabelVelden 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         ScaleHeight     =   1.063
         ScaleMode       =   4  'Character
         ScaleWidth      =   51.125
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   0
         Width           =   6135
         Begin VB.Label TabelVeld 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   4800
            LinkTimeout     =   0
            TabIndex        =   12
            Top             =   0
            UseMnemonic     =   0   'False
            Width           =   1335
         End
         Begin VB.Label TabelVeld 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   3480
            LinkTimeout     =   0
            TabIndex        =   11
            Top             =   0
            UseMnemonic     =   0   'False
            Width           =   1335
         End
         Begin VB.Label TabelVeld 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   2160
            LinkTimeout     =   0
            TabIndex        =   10
            Top             =   0
            UseMnemonic     =   0   'False
            Width           =   1335
         End
         Begin VB.Label TabelVeld 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   0
            LinkTimeout     =   0
            TabIndex        =   9
            Top             =   0
            UseMnemonic     =   0   'False
            Width           =   2175
         End
      End
      Begin VB.VScrollBar ZichtbareRijSchuifbalk 
         Height          =   975
         Left            =   6120
         Max             =   0
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Label ProductLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Product:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      LinkTimeout     =   0
      TabIndex        =   5
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   2055
   End
   Begin VB.Line Scheidingsbalk1 
      X1              =   1
      X2              =   54
      Y1              =   2
      Y2              =   2
   End
   Begin VB.Line Scheidingsbalk2 
      X1              =   1
      X2              =   54
      Y1              =   7
      Y2              =   7
   End
   Begin VB.Label HoeveelheidLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Hoeveelheid:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3600
      LinkTimeout     =   0
      TabIndex        =   4
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label PrijsLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Prijs:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2280
      LinkTimeout     =   0
      TabIndex        =   3
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label KostprijsLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Kostprijs:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4920
      LinkTimeout     =   0
      TabIndex        =   2
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   1575
   End
   Begin VB.Label TotaalLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Totaal:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4200
      LinkTimeout     =   0
      TabIndex        =   1
      Top             =   1800
      UseMnemonic     =   0   'False
      Width           =   615
   End
   Begin VB.Label TotaalVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4920
      LinkTimeout     =   0
      TabIndex        =   0
      Top             =   1800
      UseMnemonic     =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "BuffetVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze module bevat het buffetvenster.
Option Explicit

'Deze procedure verwerkt de invoer van de gebruiker.
Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo Fout

   If KeyAscii = vbKeyTab Then
      BuffettabelNavigatietoets vbKeyRight
   Else
      With Buffet
         Select Case TabelVeld(.ActiefVeld).Tag
            Case "Product"
               .Product(.Productrij) = BewerkProduct(.Product(.Productrij), KeyAscii, MAX_PRODUCTNAAM_LENGTE)
               
               TabelVeld(.ActiefVeld).Caption = .Product(.Productrij)
               Gerecht.WijzigingenBewaard = False
            Case "Prijs"
               .Prijs(.Productrij) = BewerkGetal(.Prijs(.Productrij), KeyAscii, MAX_PRIJS, IsBedrag:=True)
               
               TabelVeld(.ActiefVeld).Caption = .Prijs(.Productrij)
               Gerecht.WijzigingenBewaard = False
            Case "Hoeveelheid"
               .Hoeveelheid(.Productrij) = BewerkGetal(.Hoeveelheid(.Productrij), KeyAscii, MAX_HOEVEELHEID)
   
               TabelVeld(.ActiefVeld).Caption = .Hoeveelheid(.Productrij)
               Gerecht.WijzigingenBewaard = False
         End Select
      End With
   End If
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure verwerkt de invoer van de gebruiker.
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Fout

   With Buffet
      Select Case KeyCode
         Case vbKeyDelete
            If Not TabelVeld(.ActiefVeld).Tag = "Kostprijs" Then
               TabelVeld(.ActiefVeld).Caption = vbNullString
            End If
            
            Select Case TabelVeld(.ActiefVeld).Tag
               Case "Product"
                  .Product(.Productrij) = vbNullString
               Case "Prijs"
                  .Prijs(.Productrij) = "0.00"
               Case "Hoeveelheid"
                  .Hoeveelheid(.Productrij) = "0"
            End Select
         Case Else
            BuffettabelNavigatietoets KeyCode
      End Select
   End With

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
   
   BuffetStelTabelIn
   BuffetStelVensterIn

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure selecteert het aangeklikte tabelveld.
Private Sub TabelVeld_Click(Index As Integer)
On Error GoTo Fout

   BuffetVerwerkIngevoerdProduct
   
   TabelVeld(Buffet.ActiefVeld).BackColor = vbWhite
   Buffet.ActiefVeld = Index
   TabelVeld(Buffet.ActiefVeld).BackColor = GRIJS
   Buffet.Productrij = BuffetGeselecteerdeRij()
   
   BuffetScrollTabel

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure toont het meeteenhedenmenu na een dubbele klik in een prijsveld.
Private Sub TabelVeld_DblClick(Index As Integer)
On Error GoTo Fout

   If TabelVeld(Buffet.ActiefVeld).Tag = "Prijs" Then
      Buffet.Meeteenheid(Buffet.Productrij) = SelecteerMeeteenheid(Buffet.Meeteenheid(Buffet.Productrij), BuffetVenster, DecilitersZichtbaar:=True)
      
      BuffetVerwerkInvoer
   End If

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure past aan welke tabelrijen zichtbaar zijn.
Private Sub ZichtbareRijSchuifbalk_Change()
On Error GoTo Fout
Dim Rij As Long
Dim TabelVeldenY As Long

   Rij = ZichtbareRijSchuifbalk.Value
   TabelVeldenY = Rij
   TabelVelden.Top = -TabelVeldenY
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

