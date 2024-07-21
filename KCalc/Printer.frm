VERSION 5.00
Begin VB.Form PrinterinstellingenVenster 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Printerinstellingen"
   ClientHeight    =   2775
   ClientLeft      =   2055
   ClientTop       =   1725
   ClientWidth     =   3615
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
   ScaleHeight     =   11.563
   ScaleMode       =   4  'Character
   ScaleWidth      =   30.125
   Begin VB.VScrollBar FormaatSchuifbalk 
      Height          =   255
      Left            =   2160
      Max             =   1
      Min             =   40
      TabIndex        =   4
      Top             =   1200
      Value           =   1
      Width           =   255
   End
   Begin VB.CheckBox VetVeld 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "vet"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.CheckBox OnderstreeptVeld 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "onderstreept"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.CheckBox CursiefVeld 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "cursief"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Sluiten 
      Appearance      =   0  'Flat
      Caption         =   "&Sluiten"
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
   End
   Begin VB.ListBox LettertypeLijst 
      Appearance      =   0  'Flat
      Height          =   1590
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label FormaatVeld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2400
      LinkTimeout     =   0
      TabIndex        =   7
      Top             =   1200
      UseMnemonic     =   0   'False
      Width           =   1095
   End
   Begin VB.Label FormaatLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Formaat:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2160
      LinkTimeout     =   0
      TabIndex        =   8
      Top             =   960
      UseMnemonic     =   0   'False
      Width           =   1335
   End
   Begin VB.Label VoorbeeldVeld 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Voorbeeld"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      LinkTimeout     =   0
      TabIndex        =   6
      Top             =   1920
      UseMnemonic     =   0   'False
      Width           =   3375
   End
End
Attribute VB_Name = "PrinterinstellingenVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze module bevat het printerinstellingenvenster.
Option Explicit
'Deze procedure werkt het voorbeeldveld bij.
Private Sub CursiefVeld_Click()
On Error GoTo Fout

   VoorbeeldVeld.FontItalic = (CursiefVeld.Value = vbChecked)
   
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
           
   ToonPrinterlettertypen LettertypeLijst
   
   With Printerlettertype
      CursiefVeld.Value = Abs(.Cursief)
      OnderstreeptVeld.Value = Abs(.Onderstreept)
      VetVeld.Value = Abs(.Vet)
      FormaatSchuifbalk.Value = CInt(.Formaat)
      SelecteerLijstitem LettertypeLijst, .Naam
   End With
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure legt de geselecteerde printerinstellingen vast.
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Fout

   With Printerlettertype
      .Cursief = (CursiefVeld.Value = vbChecked)
      .Onderstreept = (OnderstreeptVeld.Value = vbChecked)
      .Vet = (VetVeld.Value = vbChecked)
      .Formaat = FormaatSchuifbalk.Value
      If LettertypeLijst.ListIndex >= 0 Then .Naam = LettertypeLijst.List(LettertypeLijst.ListIndex)
   
      BewaarPrinterinstelingen
      StelPrinterIn
   End With
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure werkt het voorbeeldveld bij.
Private Sub FormaatSchuifbalk_Change()
On Error GoTo Fout

   FormaatVeld.Caption = FormaatSchuifbalk.Value
   VoorbeeldVeld.FontSize = FormaatSchuifbalk.Value
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure werkt het voorbeeldveld bij.
Private Sub LettertypeLijst_Click()
On Error GoTo Fout

   VoorbeeldVeld.FontName = LettertypeLijst.List(LettertypeLijst.ListIndex)
   VoorbeeldVeld.FontItalic = (CursiefVeld.Value = vbChecked)
   VoorbeeldVeld.FontUnderline = (OnderstreeptVeld.Value = vbChecked)
   VoorbeeldVeld.FontBold = (VetVeld.Value = vbChecked)

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure werkt het voorbeeldveld bij.
Private Sub OnderstreeptVeld_Click()
On Error GoTo Fout

   VoorbeeldVeld.FontUnderline = (OnderstreeptVeld.Value = vbChecked)
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure sluit dit venster in.
Private Sub Sluiten_Click()
On Error GoTo Fout

   Unload Me
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure werkt het voorbeeldveld bij.
Private Sub VetVeld_Click()
On Error GoTo Fout

   VoorbeeldVeld.FontBold = (VetVeld.Value = vbChecked)
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

