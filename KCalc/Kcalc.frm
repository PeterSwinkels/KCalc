VERSION 5.00
Begin VB.MDIForm KCalcVenster 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "Keuken Calculator"
   ClientHeight    =   3795
   ClientLeft      =   3975
   ClientTop       =   2925
   ClientWidth     =   6390
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu EindeMenu 
      Caption         =   "&Einde"
   End
   Begin VB.Menu PrintenHoofdmenu 
      Caption         =   "&Printen"
      Begin VB.Menu PrintGerechtMenu 
         Caption         =   "Print &Gerecht"
      End
      Begin VB.Menu PrintBuffetMenu 
         Caption         =   "Print &Buffet"
      End
      Begin VB.Menu PrinterinstellingenMenu 
         Caption         =   "&Printerinstellingen"
      End
   End
   Begin VB.Menu BackupMenu 
      Caption         =   "&Backup"
   End
   Begin VB.Menu HandleidingMenu 
      Caption         =   "&Handleiding"
   End
   Begin VB.Menu InformatieMenu 
      Caption         =   "&Informatie"
   End
   Begin VB.Menu MeeteenhedenHoofdmenu 
      Caption         =   "&Meeteenheden"
      Visible         =   0   'False
      Begin VB.Menu MeeteenheidMenu 
         Caption         =   "&Kilo's"
         Index           =   0
      End
      Begin VB.Menu MeeteenheidMenu 
         Caption         =   "&Stuks"
         Index           =   1
      End
      Begin VB.Menu MeeteenheidMenu 
         Caption         =   "&Deciliters"
         Index           =   2
      End
   End
End
Attribute VB_Name = "KCalcVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze module bevat het hoofdvenster.
Option Explicit

'Deze procudure geeft de opdracht om het backupvenster te tonen.
Private Sub BackupMenu_Click()
On Error GoTo Fout

   BackupVenster.Show
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure sluit dit venster.
Private Sub EindeMenu_Click()
On Error GoTo Fout

   Unload Me
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procudure geeft de opdracht om het handleidingvenster te tonen.
Private Sub HandleidingMenu_Click()
On Error GoTo Fout

   HandleidingVenster.Show
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procudure geeft de opdracht om het programmainformatievenster te tonen.
Private Sub InformatieMenu_Click()
On Error GoTo Fout

   InformatieVenster.Show
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure stelt dit venster in.
Private Sub MDIForm_Load()
On Error GoTo Fout

   Me.Width = Screen.Width / 2
   Me.Height = Screen.Height / 2
   Me.WindowState = vbMaximized
   
   Load MenuVenster
   
EindeProcedure:
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure verwerkt de geselecteerde meeteenheid.
Private Sub MeeteenheidMenu_Click(Index As Integer)
On Error GoTo Fout

   WerkMeeteenhedenmenuBij GeselecteerdeMeeteenheid:=CLng(Index)
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht het buffet af te drukken.
Private Sub PrintBuffetMenu_Click()
On Error GoTo Fout

   BuffetDrukAf

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procudure geeft de opdracht om het printerinstellingenvenster te tonen.
Private Sub PrinterinstellingenMenu_Click()
On Error GoTo Fout

   PrinterinstellingenVenster.Show
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procudure geeft de opdracht om het gerecht af te drukken.
Private Sub PrintGerechtMenu_Click()
On Error GoTo Fout

   DrukGerechtAf

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

