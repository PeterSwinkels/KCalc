Attribute VB_Name = "KernModule"
'Deze module bevat de kernprocudures voor dit programma.
Option Explicit

'Deze opsomming bevat the buffettabelkolommen.
Private Enum BuffetkolommenE
   BKProduct      'De productnaam.
   BKPrijs        'De productprijs.
   BKHoeveelheid  'De producthoeveelheid.
   BKKostprijs    'De productkostprijs.
End Enum

'Deze opsomming bevat de ondersteunde meeteenhedenlijst.
Private Enum MeeteenhedenE
   MEKilos        'De kilo meeteenheid.
   MEStuks        'De stuks meeteenheid.
   MEDeciliters   'De deciliters meeteenheid.
   MEProcent      'De percentage meeteenheid.
   METijd         'De tijd in minuten meeteenheid.
End Enum

'Deze structuur definieert productencategorieen met meerdere productrijen.
Private Type BBDGMSTStr
   Product() As String               'Definieert de productenlijst.
   Prijs() As String                 'Definieert de productprijzenlijst.
   Hoeveelheid() As String           'Definieert de producthoeveelhedenlijst.
   Kostprijs() As String             'Definieert de productkostprijs gebaseerd op de productprijs en hoeveelheid.
   Totaalkosten As String            'Definieert de productentotaalprijs in een categorie.
   Meeteenheid() As MeeteenhedenE    'Definieert de geselecteerde meeteenheid voor een productrij.
   Selectie() As Long                'Definieert het geselecteerde product voor een productrij.
   Productrij As Long                'Definieert de geselecteerde productrij in een categorie.
End Type

'Deze structuur definieert de buffetcategorie.
Private Type BuffetStr
   Product() As String               'Definieert de productenlijst.
   Prijs() As String                 'Definieert de productprijzenlijst.
   Hoeveelheid() As String           'Definieert de producthoeveelhedenlijst.
   Kostprijs() As String             'Definieert de productkostprijs gebaseerd op de productprijs en hoeveelheid.
   Totaalkosten As String            'Definieert de productentotaalprijs in een categorie.
   Meeteenheid() As MeeteenhedenE    'Definieert de geselecteerde meeteenheid voor een productrij.
   ActiefVeld As Long                'Definieert het geselecteerde veld in de buffettabel.
   Productrij As Long                'Definieert de geselecteerde productrij.
End Type

'Deze structuur definieert de gerechtbestanden.
Public Type GerechtbestandStr
   Naam As String    'Defineert de gerechtsbestandsnaam.
   Datum As String   'Definieert de gerechtsbestandsdatum.
End Type

'Deze structuur definieert een gerecht.
Private Type GerechtStr
   Naam As String                  'Definieert de gerechtnaam.
   Datum As String                 'Definieert de datum waarop een gerecht bewaard is.
   HuidigBestand As String         'Definieert de gerechtbestandsnaam.
   HuidigNr As Long                'Definieert het gerechtnummer.
   Totaalkosten As String          'Definieert de gerechttotaalkosten.
   WijzigingenBewaard As Boolean   'Geeft aan of wijzigingen binnen een gerecht bewaard zijn.
End Type

'Deze structuur definieert de personeelskosten.
Private Type PersoneelskostenStr
   Loon() As String         'Definieert het loon.
   Tijd() As String         'Definieert de tijd in minuten.
   Totaalkosten As String   'Definieert de totaalkosten van alle lonen en tijden.
End Type

'Deze structuur defineert het printerlettertype.
Private Type PrinterlettertypeStr
   Naam As String            'Definieert de lettertypenaam.
   Cursief As Boolean        'Geeft aan of het lettertype cursief is.
   Onderstreept As Boolean   'Geeft aan of het lettertype onderstreept is.
   Vet As Boolean            'Geeft aan of het lettertype vet is.
   Formaat As Long           'Definieert het lettertypeformaat.
End Type

   
Private Type VVStr
   Product() As String               'Definieert de productenlijst.
   Prijs() As String                 'Definieert de productprijzenlijst.
   Hoeveelheid() As String           'Definieert de producthoeveelhedenlijst.
   Snijverlies() As String           'Definieert de productsnijverliezenlijst.
   Kostprijs As String               'Definieert de productkostprijs gebaseerd op de prijs, hoeveelheid, en het snijverlies.
   Meeteenheid As MeeteenhedenE      'Definieert de geselecteerde meeteenheid.
   Selectie As Long                  'Definieert het geselecteerde product.
End Type

Private Const BUFFET_AANTAL_KOLOMMEN As Long = 4            'Definieert het aantal kolommen in de buffettabel.
Private Const BUFFET_AANTAL_VELDEN As Long = 1200           'Definieert het aantal velden in de buffettabel.
Private Const BUFFET_AANTAL_ZICHTBARE_RIJEN As Long = 4     'Definieert het aantal rijen dat tegelijk zichtbaar is in de buffettabel.
Private Const E_TREMA As Long = &HEB                        'Definieert de letter "e" met een trema.
Private Const GEEN_AANTAL As Long = -1                      'Definieert dat er geen productaantal opgegeven wordt.
Private Const GEEN_MEETEENHEID As Long = -1                 'Definieert dat er geen meeteenheid geselecteerd is.
Private Const MAX_AANTAL_GERECHTEN As Long = 100            'Definieert het maximale aantal gerechten dat bewaard kan worden.
Public Const DATA_MAP As String = ".\Data\"                 'Definieert het datamappad.
Public Const GEEN_GERECHTNR As Long = -1                    'Definieert het ontbreken van een gerechtnummer.
Public Const GEEN_PRODUCT_TEKST As String = "geen"          'Definieert de tekst die wordt weergegeven wanneer er geen product geselecteerd is.
Public Const GRIJS As Long = &HC0C0C0                       'Definieert de kleur lichtgrijs.
Public Const MAX_HOEVEELHEID As Long = 10000                'Definieert de maximale hoeveelheid voor producten.
Public Const MAX_LOON As Double = 1000                      'Definieert de maximale loon.
Public Const MAX_PERCENTAGE As Long = 100                   'Definieert het maximale percentage.
Public Const MAX_PRIJS As Double = 1000                     'Definieert de maximale prijs voor producten.
Public Const MAX_PRODUCTNAAM_LENGTE As Long = 255           'Definieert de maximale productnaam lengte.
Public Const MAX_TIJD As Long = 1440                        'Definieert de maximale tijd.

Public Bijproducten As BBDGMSTStr                  'Definieert de bijproducten categorie.
Public Bittergarnituur As BBDGMSTStr               'Definieert de bittergarnituur categorie.
Public Buffet As BuffetStr                         'Definieert de buffet categorie.
Public Doorverkoopproducten As BBDGMSTStr          'Definieert de doorverkoopproducten categorie.
Public Gerecht As GerechtStr                       'Definieert het gerecht.
Public GerechtBestanden() As GerechtbestandStr     'Definieert de gerechtbestanden.
Public Groenten As BBDGMSTStr                      'Definieert de groenten categorie.
Public Meelspijzen As BBDGMSTStr                   'Definieert de meelspijzen categorie.
Public Personeelskosten As PersoneelskostenStr     'Definieert de personeelskosten.
Public Printerlettertype As PrinterlettertypeStr   'Definieert het printerlettertype.
Public SoepenEnSausen As BBDGMSTStr                'Definieert de soepen en sausen categorie.
Public Tafelgarnituur As BBDGMSTStr                'Definieert de tafelgarnituur categorie.
Public Vlees As VVStr                              'Definieert de vlees categorie.
Public Vis As VVStr                                'Definieert de vis categorie.

'Deze procedure stuurt de backuplocatie terug.
Public Function Backuplocatie() As String
On Error GoTo Fout
Dim BestandH As Integer
Dim Locatie As String

   Screen.MousePointer = vbHourglass
   
   If Not Dir$("Backup.loc", vbArchive Or vbHidden Or vbSystem) = vbNullString Then
      BestandH = FreeFile()
      Open "Backup.loc" For Binary Lock Read Write As BestandH
         Locatie = Input$(LOF(BestandH), BestandH)
      Close BestandH
   Else
      Locatie = Stationsletter(CurDir$()) & "\Backup\"
   End If
   
EindeProcedure:
   Screen.MousePointer = vbDefault
   Backuplocatie = Locatie
   Exit Function

Fout:
   If HandelFoutAf(, ActiefBestand:="het backuplocatiebestand") = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Function

'Deze procedure berekent het totaal van de opgegeven producten.
Private Sub BBDGMSTBerekenTotaal(Producten As BBDGMSTStr)
On Error GoTo Fout
Dim Productrij As Long

   With Producten
      .Totaalkosten = "0.00"
      
      For Productrij = LBound(.Product(), 2) To UBound(.Product(), 2)
         .Totaalkosten = TelBedragenOp(.Totaalkosten, .Kostprijs(Productrij))
      Next Productrij
      
      .Totaalkosten = RondBedragAf(.Totaalkosten)
   End With
   
EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure bewaart de opgegeven producten in het opgegeven gerechtbestand.
Private Sub BBDGMSTBewaarGerecht(Producten As BBDGMSTStr, BestandH As Integer)
On Error GoTo Fout
Dim Productnr As Long
Dim Productrij As Long

   With Producten
      For Productrij = LBound(.Product(), 2) To UBound(.Product(), 2)
         Print #BestandH, Chr$(.Selectie(Productrij));
         For Productnr = LBound(.Product(), 1) To UBound(.Product(), 1)
            Print #BestandH, Chr$(Len(.Hoeveelheid(Productnr, Productrij))); .Hoeveelheid(Productnr, Productrij);
         Next Productnr
      Next Productrij
   End With

EindeProcedure:
   Exit Sub
   
Fout:
   If HandelFoutAf(, ActiefBestand:="het gerechtbestand") = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure bewaart de opgegeven producten in het opgegeven productenbestand.
Private Sub BBDGMSTBewaarProducten(Producten As BBDGMSTStr, BestandH As Integer)
On Error GoTo Fout
Dim Productnr As Long
Dim Productrij As Long

   With Producten
      For Productrij = LBound(.Product(), 2) To UBound(.Product(), 2)
         Print #BestandH, Chr$(.Meeteenheid(Productrij));
         For Productnr = LBound(.Product(), 1) + 1 To UBound(.Product(), 1)
            Print #BestandH, Chr$(Len(.Product(Productnr, Productrij))); .Product(Productnr, Productrij);
            Print #BestandH, Chr$(Len(.Prijs(Productnr, Productrij))); .Prijs(Productnr, Productrij);
         Next Productnr
      Next Productrij
   End With
      
EindeProcedure:
   Exit Sub
   
Fout:
   If HandelFoutAf(, ActiefBestand:="het productenbestand") = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure drukt producten uit de opgegeven categorie af.
Private Sub BBDGMSTDrukAf(Producten As BBDGMSTStr, Categorie As String)
On Error GoTo Fout
Dim Productrij As Long

   With Producten
      Printer.Print
      Printer.Print Tab(3); Categorie; ":"
      Printer.Print String(75, "-")
      Printer.Print Tab(3); "Product:"; Tab(23); "Prijs:"; Tab(35); "Hoeveelheid:"; Tab(49); "Kostprijs:"
      Printer.Print String(75, "-")
      For Productrij = LBound(.Product(), 2) To UBound(.Product(), 2)
         Printer.Print Tab(3); Left$(.Product(.Selectie(Productrij), Productrij), 20);
         Printer.Print Tab(31 - Len(.Prijs(.Selectie(Productrij), Productrij))); .Prijs(.Selectie(Productrij), Productrij);
         Printer.Print Tab(47 - Len(.Hoeveelheid(.Selectie(Productrij), Productrij))); .Hoeveelheid(.Selectie(Productrij), Productrij);
         Printer.Print Tab(57 - Len(.Kostprijs(Productrij))); .Kostprijs(Productrij)
      Next Productrij
      Printer.Print String(75, "-")
      Printer.Print Tab(31 - Len(.Totaalkosten)); .Totaalkosten
   End With

EindeProcedure:
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procudure laadt een gerecht uit het opgegeven bestand.
Private Sub BBDGMSTLaadGerecht(Producten As BBDGMSTStr, BestandH As Integer)
On Error GoTo Fout
Dim Lengte As Long
Dim Productnr As Long
Dim Productrij As Long

   With Producten
      For Productrij = LBound(.Product(), 2) To UBound(.Product(), 2)
         .Selectie(Productrij) = Asc(Input$(1, BestandH))
         For Productnr = LBound(Groenten.Product(), 1) To UBound(.Product(), 1)
            Lengte = Asc(Input$(1, BestandH)): .Hoeveelheid(Productnr, Productrij) = Val(Input$(Lengte, BestandH))
            .Kostprijs(Productrij) = BerekenBedrag(.Prijs(.Selectie(Productrij), Productrij), .Hoeveelheid(.Selectie(Productrij), Productrij), .Meeteenheid(Productrij))
         Next Productnr
      Next Productrij
   End With
         
EindeProcedure:
   Exit Sub
   
Fout:
   If HandelFoutAf(, ActiefBestand:="het gerechtbestand") = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procudure laadt de producten uit het opgegeven bestand.
Private Sub BBDGMSTLaadProducten(Producten As BBDGMSTStr, BestandH As Integer)
On Error GoTo Fout
Dim Lengte As Long
Dim Productnr As Long
Dim Productrij As Long

   With Producten
      For Productrij = LBound(.Product(), 2) To UBound(.Product(), 2)
         .Meeteenheid(Productrij) = Asc(Input$(1, BestandH))
         For Productnr = LBound(.Product(), 1) + 1 To UBound(.Product(), 1)
            Lengte = Asc(Input$(1, BestandH)): .Product(Productnr, Productrij) = Input$(Lengte, BestandH)
            Lengte = Asc(Input$(1, BestandH)): .Prijs(Productnr, Productrij) = Input$(Lengte, BestandH)
         Next Productnr
      Next Productrij
   End With
         
EindeProcedure:
   Exit Sub
   
Fout:
   If HandelFoutAf(, ActiefBestand:="het productenbestand") = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure stelt de opgegeven product categorie in.
Private Sub BBDGMSTStelIn(Producten As BBDGMSTStr, Optional AantalProducten As Long = GEEN_AANTAL)
5 On Error GoTo Fout
Dim Productnr As Long
Dim Productrij As Long

   With Producten
      If Not AantalProducten = GEEN_AANTAL Then
         ReDim .Product(0 To AantalProducten, 0 To 2) As String
         ReDim .Prijs(0 To AantalProducten, 0 To 2) As String
         ReDim .Meeteenheid(0 To 2) As MeeteenhedenE
         ReDim .Hoeveelheid(0 To AantalProducten, 0 To 2) As String
         ReDim .Kostprijs(0 To 2) As String
         ReDim .Selectie(0 To 2) As Long
      End If

      .Productrij = 0
      .Totaalkosten = "0.00"
      For Productrij = LBound(.Product(), 2) To UBound(.Product(), 2)
         If Not AantalProducten = GEEN_AANTAL Then
            .Meeteenheid(Productrij) = MEKilos
            .Product(LBound(.Product(), 1), Productrij) = GEEN_PRODUCT_TEKST
         End If
         
         .Kostprijs(Productrij) = "0.00"
         
         For Productnr = LBound(.Product(), 1) To UBound(.Product(), 1)
            If Not AantalProducten = GEEN_AANTAL Then
               If Productnr > LBound(.Product(), 1) Then .Product(Productnr, Productrij) = "product " & Productnr
               .Prijs(Productnr, Productrij) = "0.00"
            End If
            .Hoeveelheid(Productnr, Productrij) = "0"
         Next Productnr
      Next Productrij
   End With
   
EindeProcedure:
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure stelt het opgegeven venster in met de opgegeven producten.
Public Sub BBDGMSTStelVensterIn(Producten As BBDGMSTStr, Venster As Form, VensterNaam As String)
On Error GoTo Fout
Dim Productnr As Long
Dim Productrij As Long
  
   With Producten
      Screen.MousePointer = vbHourglass
      
      If VensterActief(VensterNaam) Then
         VensterNaastMenu Venster, VensterNaam
         
         For Productrij = LBound(.Product(), 2) To UBound(.Product(), 2)
            Venster.ProductenLijst(Productrij).Clear
            
            For Productnr = LBound(.Product(), 1) To UBound(.Product(), 1)
               Venster.ProductenLijst(Productrij).AddItem .Product(Productnr, Productrij)
            Next Productnr
           
            .Productrij = Productrij
            Venster.ProductenLijst(Productrij).Text = .Product(.Selectie(Productrij), Productrij)
            Venster.HoeveelheidVeld(Productrij).Text = .Hoeveelheid(.Selectie(Productrij), Productrij)
            Venster.PrijsVeld(Productrij).Text = RondBedragAf(.Prijs(.Selectie(Productrij), Productrij))
         Next Productrij
         
         For Productrij = LBound(.Product(), 2) To UBound(.Product(), 2)
            Venster.ProductenLijst(Productrij).ListIndex = .Selectie(Productrij)
         Next Productrij
      End If
      
      BBDGMSTBerekenTotaal Producten
   End With
   
EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure verwerkt de productinformatieinvoer.
Public Sub BBDGMSTVerwerkInvoer(Producten As BBDGMSTStr, Venster As Form)
On Error GoTo Fout
Dim Productrij As Long

   With Producten
      For Productrij = LBound(.Product(), 2) To UBound(.Product(), 2)
         Venster.PrijsVeld(Productrij).Text = RondBedragAf(Venster.PrijsVeld(Productrij).Text)
         Venster.HoeveelheidVeld(Productrij).Text = CStr(Val(Venster.HoeveelheidVeld(Productrij).Text))
         
         .Prijs(.Selectie(Productrij), Productrij) = Venster.PrijsVeld(Productrij).Text
         .Hoeveelheid(.Selectie(Productrij), Productrij) = Venster.HoeveelheidVeld(Productrij).Text
         .Kostprijs(Productrij) = BerekenBedrag(.Prijs(.Selectie(Productrij), Productrij), .Hoeveelheid(.Selectie(Productrij), Productrij), .Meeteenheid(Productrij))
         
         Venster.KostPrijsVeld(Productrij).Caption = .Kostprijs(Productrij)
      Next Productrij
      
      BBDGMSTBerekenTotaal Producten
      
      Venster.TotaalVeld.Caption = .Totaalkosten
   End With
   
EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht om te controleren of een product al is ingevoerd en wijzigt eventueel de productnaam.
Public Function BBDGMSTVVControleerProduct(ProductenLijst As ComboBox, GeselecteerdProduct As Long) As String
On Error GoTo Fout
Dim Product As String

   If BBDGMSTVVProductAlIngevoerd(ProductenLijst, GeselecteerdProduct) Then
       Product = "product " & GeselecteerdProduct
   Else
       Product = ProductenLijst.Text
   End If

EindeProcedure:
   BBDGMSTVVControleerProduct = Product
   Exit Function
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Function

'Deze procedure geeft aan of het opgegeven product zich al in de opgegeven productenlijst bevindt.
Private Function BBDGMSTVVProductAlIngevoerd(ProductenLijst As ComboBox, GeselecteerdProductNr As Long) As Boolean
On Error GoTo Fout
Dim AlIngevoerd As Boolean
Dim Productnr As Long

   Screen.MousePointer = vbHourglass

   AlIngevoerd = False
   For Productnr = 1 To ProductenLijst.ListCount - 1
      If Not Productnr = GeselecteerdProductNr Then
         If Not (Trim$(LCase$(ProductenLijst.Text)) = vbNullString Or Trim$(LCase$(ProductenLijst.List(Productnr))) = vbNullString) Then
            If LCase$(ProductenLijst.Text) = LCase$(ProductenLijst.List(Productnr)) Then
               AlIngevoerd = True
               Exit For
            End If
         End If
      End If
   Next Productnr
   
   If AlIngevoerd Then MsgBox "Dat product is al ingevoerd.", vbExclamation
   
EindeProcedure:
   Screen.MousePointer = vbDefault
   BBDGMSTVVProductAlIngevoerd = AlIngevoerd
   Exit Function
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Function

'Deze procedure stuurt een bedrag berekend met opgegeven parameters terug.
Private Function BerekenBedrag(Bedrag As String, Hoeveelheid As String, Meeteenheid As MeeteenhedenE) As String
On Error GoTo Fout
Dim NieuwBedrag As String

   NieuwBedrag = Replace(Bedrag, ",", ".")
      
   Select Case Meeteenheid
      Case MEKilos
         NieuwBedrag = CStr((Val(NieuwBedrag) / 1000) * Val(Hoeveelheid))
      Case MEStuks
         NieuwBedrag = CStr(Val(NieuwBedrag) * Val(Hoeveelheid))
      Case MEDeciliters
         NieuwBedrag = CStr((Val(NieuwBedrag) / 10) * Val(Hoeveelheid))
      Case MEProcent
         NieuwBedrag = CStr((Val(NieuwBedrag) / 100) * Val(Hoeveelheid))
      Case METijd
         NieuwBedrag = CStr((Val(NieuwBedrag) / 60) * Val(Hoeveelheid))
   End Select
   
   NieuwBedrag = RondBedragAf(Replace(NieuwBedrag, ",", "."))

EindeProcedure:
   BerekenBedrag = NieuwBedrag
   Exit Function

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Function

'Deze procedure bewaart de opgegeven backuplocatie.
Public Sub BewaarBackupLocatie(Locatie As String)
On Error GoTo Fout
Dim BestandH As Integer

   Screen.MousePointer = vbHourglass
   
   BestandH = FreeFile()
   Open "Backup.loc" For Output Lock Read Write As BestandH
      Print #BestandH, Locatie;
   Close BestandH
   
EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub

Fout:
   If HandelFoutAf(, ActiefBestand:="het backuplocatiebestand") = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure bewaart het buffet.
Public Sub BewaarBuffet()
On Error GoTo Fout
Dim BestandH As Integer
Dim Productnr As Long

   Screen.MousePointer = vbHourglass
   
   BestandH = FreeFile()
   Open DATA_MAP & "Buffet.kcc" For Output Lock Read Write As BestandH
      With Buffet
         Print #BestandH, Chr$(Len(CStr(.Productrij))); CStr(.Productrij);
         For Productnr = LBound(.Product()) To UBound(.Product())
            Print #BestandH, Chr$(Len(.Product(Productnr))); .Product(Productnr);
            Print #BestandH, Chr$(Len(.Prijs(Productnr))); .Prijs(Productnr);
            Print #BestandH, Chr$(Len(.Hoeveelheid(Productnr))); .Hoeveelheid(Productnr);
            Print #BestandH, Chr$(.Meeteenheid(Productnr));
         Next Productnr
      End With
   Close BestandH

EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub

Fout:
   If HandelFoutAf(, ActiefBestand:="het buffetbestand") = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure bewaart het huidige gerecht in het opgegeven bestand.
Public Sub BewaarGerecht(Gerechtbestand As String)
On Error GoTo Fout
Dim BestandH As Integer
Dim Datum As String
Dim Personeelskostenrij As Long

   Screen.MousePointer = vbHourglass
   
   BestandH = FreeFile()
   Open DATA_MAP & Gerechtbestand For Output Lock Read Write As BestandH
      With Gerecht
         Print #BestandH, Chr$(Len(.Naam)); .Naam;
         Print #BestandH, Chr$(Len(.Datum)); .Datum;
      End With
      
      VVBewaarGerecht Vlees, BestandH
      VVBewaarGerecht Vis, BestandH
      BBDGMSTBewaarGerecht Groenten, BestandH
      BBDGMSTBewaarGerecht Meelspijzen, BestandH
      BBDGMSTBewaarGerecht SoepenEnSausen, BestandH
      BBDGMSTBewaarGerecht Bijproducten, BestandH
      BBDGMSTBewaarGerecht Doorverkoopproducten, BestandH
      BBDGMSTBewaarGerecht Bittergarnituur, BestandH
      BBDGMSTBewaarGerecht Tafelgarnituur, BestandH
      
      With Personeelskosten
         For Personeelskostenrij = LBound(.Tijd()) To UBound(.Tijd())
            Print #BestandH, Chr$(Len(.Tijd(Personeelskostenrij))); .Tijd(Personeelskostenrij);
         Next Personeelskostenrij
      End With
   Close BestandH

   Gerecht.WijzigingenBewaard = True

EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub

Fout:
   If HandelFoutAf(, ActiefBestand:="het gerechtbestand") = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure geeft aan of de gebruiker verzocht heeft het huidige gerecht te bewaren.
Public Function BewaarGerechtDialoog() As Long
On Error GoTo Fout
Dim Antwoord As Long
   
   Antwoord = MsgBox("Gerecht bewaren?", vbQuestion Or vbYesNo Or vbDefaultButton1)

EindeProcedure:
   BewaarGerechtDialoog = Antwoord
   Exit Function
   
Fout:
   If HandelFoutAf(, ActiefBestand:="het gerechtbestand") = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Function

'Deze procedure bewaart de personeelskosten.
Private Sub BewaarPersoneelskosten()
On Error GoTo Fout
Dim BestandH As Integer
Dim Personeelskostenrij As Long

   Screen.MousePointer = vbHourglass
   
   BestandH = FreeFile()
   Open DATA_MAP & "Personeel.kcc" For Output Lock Read Write As BestandH
      With Personeelskosten
         For Personeelskostenrij = LBound(.Loon()) To UBound(.Loon())
            Print #BestandH, Chr$(Len(.Loon(Personeelskostenrij))); .Loon(Personeelskostenrij);
         Next Personeelskostenrij
      End With
   Close BestandH
   
EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub
   
Fout:
   If HandelFoutAf(, ActiefBestand:="het personeelskostenbestand") = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure bewaart de printerinstellingen.
Public Sub BewaarPrinterinstelingen()
On Error GoTo Fout
Dim BestandH As Integer

   Screen.MousePointer = vbHourglass
 
   BestandH = FreeFile()
   Open "Printer.kcc" For Output Lock Read Write As BestandH
      With Printerlettertype
         Print #BestandH, Chr$(Abs(.Cursief));
         Print #BestandH, Chr$(Abs(.Onderstreept));
         Print #BestandH, Chr$(Abs(.Vet));
         Print #BestandH, Chr$(.Formaat);
         Print #BestandH, Chr$(Len(.Naam)); .Naam;
      End With
   Close BestandH

EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub

Fout:
   If HandelFoutAf(, ActiefBestand:="het printerinstellingenbestand") = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure bewaart de huidige producten.
Public Sub BewaarProducten()
On Error GoTo Fout
Dim BestandH As Integer

   Screen.MousePointer = vbHourglass
   
   BestandH = FreeFile()
   Open DATA_MAP & "Producten.kcc" For Output Lock Read Write As BestandH
      VVBewaarProducten Vlees, BestandH
      VVBewaarProducten Vis, BestandH
      BBDGMSTBewaarProducten Groenten, BestandH
      BBDGMSTBewaarProducten Meelspijzen, BestandH
      BBDGMSTBewaarProducten SoepenEnSausen, BestandH
      BBDGMSTBewaarProducten Bijproducten, BestandH
      BBDGMSTBewaarProducten Doorverkoopproducten, BestandH
      BBDGMSTBewaarProducten Bittergarnituur, BestandH
      BBDGMSTBewaarProducten Tafelgarnituur, BestandH
   Close BestandH
   
EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub
   
Fout:
   If HandelFoutAf(, ActiefBestand:="het productenbestand") = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure stuurt het opgegeven getal terug na het uitvoeren van de opgegeven bewerking.
Public Function BewerkGetal(Getal As String, ByRef Toets As Integer, MaxWaarde As Double, Optional IsBedrag As Boolean = False) As String
On Error GoTo Fout
Dim BewerktGetal As String
Dim Lengte As Long
Dim Teken As String

   Teken = Chr$(Toets)
   Toets = Empty
   BewerktGetal = Getal
   
   If IsBedrag And BewerktGetal = "0.00" Then
      BewerktGetal = vbNullString
   ElseIf BewerktGetal = "0" Then
      BewerktGetal = vbNullString
   End If
   
   If Teken = "," Then Teken = "."
  
   Select Case Teken
      Case "0" To "9"
         If Val(BewerktGetal & Teken) <= MaxWaarde Then BewerktGetal = BewerktGetal & Teken
      Case "."
         If IsBedrag Then
            If InStr(BewerktGetal, Teken) = 0 Then BewerktGetal = BewerktGetal & "."
         End If
      Case Chr$(vbKeyBack)
         Lengte = Len(BewerktGetal)
         If Lengte > 0 Then BewerktGetal = Left$(BewerktGetal, Lengte - 1)
         If Lengte = 0 Then BewerktGetal = "0"
   End Select
   
EindeProcedure:
   BewerkGetal = BewerktGetal
   Exit Function
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Function

'Deze procedure geeft opdracht om de opgegeven bewerking op het opgegeven getalveld uit te voeren.
Public Sub BewerkGetalVeld(GetalVeld As TextBox, ByRef Toets As Integer, MaxWaarde As Double, Optional IsBedrag As Boolean = False)
On Error GoTo Fout
   
   GetalVeld.Text = BewerkGetal(GetalVeld.Text, Toets, MaxWaarde, IsBedrag:=IsBedrag)
   GetalVeld.SelStart = 0
   Toets = Empty

EindeProcedure:
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure stuurt het opgegeven product terug na het uitvoeren van de opgegeven bewerking
Public Function BewerkProduct(Product As String, ByRef Toets As Integer, MaxLengte As Long) As String
On Error GoTo Fout
Dim BewerktProduct As String
Dim Lengte As Long
Dim Teken As String

   Teken = Chr$(Toets)
   Toets = Empty
   BewerktProduct = Product
   
   Lengte = Len(BewerktProduct)
     
   Select Case Teken
      Case Chr$(vbKeyBack)
         If Lengte > 0 Then BewerktProduct = Left$(BewerktProduct, Lengte - 1)
      Case Is >= " "
         If Lengte < MaxLengte Then BewerktProduct = BewerktProduct & Teken
   End Select
   
EindeProcedure:
   BewerkProduct = BewerktProduct
   Exit Function
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Function

'Deze procedure geeft opdracht om de opgegeven bewerking op het opgegeven productveld uit te voeren.
Public Sub BewerkProductVeld(ProductVeld As ComboBox, ByRef Toets As Integer, MaxLengte As Long)
On Error GoTo Fout
   
   ProductVeld.Text = BewerkProduct(ProductVeld.Text, Toets, MaxLengte)
   ProductVeld.SelStart = Len(ProductVeld.Text)
   Toets = Empty

EindeProcedure:
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure berekent de totaalkosten voor het buffet.
Private Sub BuffetBerekenTotaal()
On Error GoTo Fout
Dim Productnr As Long

   Screen.MousePointer = vbHourglass
   
   With Buffet
      .Totaalkosten = "0.00"
      
      For Productnr = LBound(.Product()) To UBound(.Product())
         .Totaalkosten = TelBedragenOp(.Totaalkosten, .Kostprijs(Productnr))
      Next Productnr
       
      .Totaalkosten = RondBedragAf(.Totaalkosten)
   End With
   
EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure drukt het buffet af.
Public Sub BuffetDrukAf()
On Error GoTo Fout
Dim Productrij As Long

   Screen.MousePointer = vbHourglass
   
   With Buffet
      Printer.Print
      Printer.Print Tab(3); "Buffet:"
      Printer.Print String(75, "-")
      Printer.Print Tab(3); "Product:"; Tab(23); "Prijs:"; Tab(35); "Hoeveelheid:"; Tab(49); "Kostprijs:"
      Printer.Print String(75, "-")
      For Productrij = LBound(.Product()) To UBound(.Product())
         Printer.Print Tab(3); Left$(.Product(Productrij), 20);
         Printer.Print Tab(31 - Len(.Prijs(Productrij))); .Prijs(Productrij);
         Printer.Print Tab(47 - Len(.Hoeveelheid(Productrij))); .Hoeveelheid(Productrij);
         Printer.Print Tab(57 - Len(.Kostprijs(Productrij))); .Kostprijs(Productrij)
      Next Productrij
      Printer.Print String(75, "-")
      Printer.Print Tab(31 - Len(.Totaalkosten)); .Totaalkosten
      Printer.EndDoc
   End With
   
EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure stuurt de geselecteerde buffettabelrij terug.
Public Function BuffetGeselecteerdeRij() As Long
On Error GoTo Fout
Dim GeselecteerdeRij As Long

   GeselecteerdeRij = Buffet.ActiefVeld \ BUFFET_AANTAL_KOLOMMEN
   
EindeProcedure:
   BuffetGeselecteerdeRij = GeselecteerdeRij
   Exit Function
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Function

'Deze procedure geeft aan of het geselecteerde buffetproduct al eerder is ingevoerd.
Private Function BuffetProductAlIngevoerd() As Boolean
On Error GoTo Fout
Dim AlIngevoerd As Boolean
Dim GeselecteerdProductNr As Long
Dim Productnr As Long

   Screen.MousePointer = vbHourglass

   AlIngevoerd = False
   GeselecteerdProductNr = BuffetGeselecteerdeRij() + 1
   
   For Productnr = LBound(Buffet.Product()) + 1 To UBound(Buffet.Product()) + 1
      If Not Productnr = GeselecteerdProductNr Then
         If LCase$(BuffetVenster.TabelVeld(Buffet.ActiefVeld)) = LCase$(BuffetVenster.TabelVeld((Productnr - 1) * 4)) Then
            AlIngevoerd = True
            Exit For
         End If
      End If
   Next Productnr
   
   If AlIngevoerd Then MsgBox "Dat product is al ingevoerd.", vbExclamation
   
EindeProcedure:
   Screen.MousePointer = vbDefault
   BuffetProductAlIngevoerd = AlIngevoerd
   Exit Function
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Function

'Deze procedure scrollt de buffettabel.
Public Sub BuffetScrollTabel()
On Error GoTo Fout
Dim Rij As Long
   
   With BuffetVenster
      Rij = BuffetGeselecteerdeRij()
      If .ZichtbareRijSchuifbalk.Value > Rij Or .ZichtbareRijSchuifbalk.Value <= Rij - BUFFET_AANTAL_ZICHTBARE_RIJEN Then
         .ZichtbareRijSchuifbalk.Value = Rij
      End If
   End With

EindeProcedure:
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure stelt de buffetproducten in.
Private Sub BuffetStelIn(AantalProducten As Long)
On Error GoTo Fout
Dim Productnr As Long

   With Buffet
      ReDim .Product(0 To AantalProducten) As String
      ReDim .Prijs(0 To AantalProducten) As String
      ReDim .Hoeveelheid(0 To AantalProducten) As String
      ReDim .Kostprijs(0 To AantalProducten) As String
      ReDim .Meeteenheid(0 To AantalProducten) As MeeteenhedenE

      .Totaalkosten = "0.00"
      For Productnr = LBound(.Product()) To UBound(.Product())
         .Product(Productnr) = "product " & Productnr
         .Prijs(Productnr) = "0.00"
         .Hoeveelheid(Productnr) = "0"
         .Kostprijs(Productnr) = "0.00"
         .Meeteenheid(Productnr) = MEKilos
      Next Productnr
   End With

EindeProcedure:
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure stelt de buffettabel in.
Public Sub BuffetStelTabelIn()
On Error GoTo Fout
Dim Subveld As Long
Dim Veld As Long
 
   With BuffetVenster
      Screen.MousePointer = vbHourglass
      
      For Veld = LBound(Buffet.Product()) To UBound(Buffet.Product()) * BUFFET_AANTAL_KOLOMMEN Step BUFFET_AANTAL_KOLOMMEN
         For Subveld = Veld To Veld + BUFFET_AANTAL_KOLOMMEN - 1
            If Veld \ BUFFET_AANTAL_KOLOMMEN > 0 Then
               Load .TabelVeld(Subveld)
               .TabelVeld(Subveld).Alignment = .TabelVeld(Subveld - BUFFET_AANTAL_KOLOMMEN).Alignment
               .TabelVeld(Subveld).BorderStyle = vbFixedSingle
               .TabelVeld(Subveld).Visible = True
               .TabelVeld(Subveld).Left = .TabelVeld(Subveld - BUFFET_AANTAL_KOLOMMEN).Left
               .TabelVeld(Subveld).Top = .TabelVeld(Subveld - BUFFET_AANTAL_KOLOMMEN).Top + 1
               .TabelVeld(Subveld).Width = .TabelVeld(Subveld - BUFFET_AANTAL_KOLOMMEN).Width
            End If
         
            Select Case Subveld - Veld
               Case BKProduct
                  .TabelVeld(Subveld).Tag = "Product"
               Case BKPrijs
                  .TabelVeld(Subveld).Tag = "Prijs"
               Case BKHoeveelheid
                  .TabelVeld(Subveld).Tag = "Hoeveelheid"
               Case BKKostprijs
                  .TabelVeld(Subveld).Tag = "Kostprijs"
            End Select
         Next Subveld
         
         .TabelVelden.Height = .TabelVelden.Height + 1
      Next Veld
      
      .ZichtbareRijSchuifbalk.Max = (.TabelVelden.Height \ 1)
      
      Buffet.ActiefVeld = 0
      .TabelVeld(Buffet.ActiefVeld).BackColor = GRIJS
   End With

EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure stelt het buffetvenster in.
Public Sub BuffetStelVensterIn()
On Error GoTo Fout
Dim Productnr As Long
Dim Veld As Long

   With Buffet
      Screen.MousePointer = vbHourglass
      
      If VensterActief("BuffetVenster") Then
         For Veld = 0 To BUFFET_AANTAL_VELDEN - BUFFET_AANTAL_KOLOMMEN Step BUFFET_AANTAL_KOLOMMEN
            Productnr = Veld \ BUFFET_AANTAL_KOLOMMEN
            BuffetVenster.TabelVeld(Veld + BKProduct).Caption = .Product(Productnr)
            BuffetVenster.TabelVeld(Veld + BKPrijs).Caption = RondBedragAf(.Prijs(Productnr))
            BuffetVenster.TabelVeld(Veld + BKHoeveelheid).Caption = .Hoeveelheid(Productnr)
            BuffetVenster.TabelVeld(Veld + BKKostprijs).Caption = .Kostprijs(Productnr)
         Next Veld
      End If
      
      BuffetBerekenTotaal
      
      BuffetVenster.TotaalVeld.Caption = .Totaalkosten
   End With

EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure verwerkt de opgegeven navigatietoets aanslag voor de buffettabel.
Public Sub BuffettabelNavigatietoets(Toets As Integer)
On Error GoTo Fout

   With BuffetVenster
      Select Case Toets
         Case vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyUp
            .TabelVeld(Buffet.ActiefVeld).BackColor = vbWhite
            BuffetVerwerkIngevoerdProduct
      End Select
   
      Select Case Toets
         Case vbKeyDown
            If Buffet.ActiefVeld + BUFFET_AANTAL_KOLOMMEN < BUFFET_AANTAL_VELDEN Then Buffet.ActiefVeld = Buffet.ActiefVeld + BUFFET_AANTAL_KOLOMMEN
         Case vbKeyLeft
            If Buffet.ActiefVeld > 0 Then Buffet.ActiefVeld = Buffet.ActiefVeld - 1
         Case vbKeyRight
            If Buffet.ActiefVeld < BUFFET_AANTAL_VELDEN - 1 Then Buffet.ActiefVeld = Buffet.ActiefVeld + 1
         Case vbKeyUp
            If Buffet.ActiefVeld - BUFFET_AANTAL_KOLOMMEN >= 0 Then Buffet.ActiefVeld = Buffet.ActiefVeld - BUFFET_AANTAL_KOLOMMEN
      End Select
      
      Select Case Toets
         Case vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyUp
            BuffetScrollTabel
            Buffet.Productrij = BuffetGeselecteerdeRij()
            BuffetVerwerkIngevoerdProduct
            .TabelVeld(Buffet.ActiefVeld).BackColor = GRIJS
      End Select
   End With

EindeProcedure:
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure stuurt het buffettabelveld terug voor de opgegeven rij en kolom.
Private Function BuffettabelVeld(Rij As Long, Kolom As Long)
On Error GoTo Fout
Dim TabelVeld As Long
   
   TabelVeld = (Rij * BUFFET_AANTAL_KOLOMMEN) + Kolom

EindeProcedure:
   BuffettabelVeld = TabelVeld
   Exit Function
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Function

'Deze procedure verkt een ingevoerd buffetproduct.
Public Sub BuffetVerwerkIngevoerdProduct()
On Error GoTo Fout

   With BuffetVenster
      Select Case .TabelVeld(Buffet.ActiefVeld).Tag
         Case "Product"
            If BuffetProductAlIngevoerd() Then
               .TabelVeld(Buffet.ActiefVeld).Caption = "product " & BuffetGeselecteerdeRij()
               Buffet.Product(Buffet.Productrij) = .TabelVeld(Buffet.ActiefVeld).Caption
            End If
         Case "Prijs", "Hoeveelheid"
            BuffetVerwerkInvoer
      End Select
   End With

EindeProcedure:
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure verwerkt de productinformatieinvoer.
Public Sub BuffetVerwerkInvoer()
On Error GoTo Fout
Dim Productnr As Long

   Screen.MousePointer = vbHourglass
   
   With Buffet
      .Totaalkosten = "0.00"
      
      For Productnr = LBound(.Product()) To UBound(.Product())
         BuffetVenster.TabelVeld(BuffettabelVeld(Productnr, Kolom:=BKPrijs)).Caption = RondBedragAf(BuffetVenster.TabelVeld(BuffettabelVeld(Productnr, Kolom:=1)).Caption)
         BuffetVenster.TabelVeld(BuffettabelVeld(Productnr, Kolom:=BKHoeveelheid)).Caption = CStr(Val(BuffetVenster.TabelVeld(BuffettabelVeld(Productnr, Kolom:=2)).Caption))
                 
         .Prijs(Productnr) = BuffetVenster.TabelVeld(BuffettabelVeld(Productnr, Kolom:=BKPrijs)).Caption
         .Hoeveelheid(Productnr) = BuffetVenster.TabelVeld(BuffettabelVeld(Productnr, Kolom:=BKHoeveelheid)).Caption
         .Kostprijs(Productnr) = BerekenBedrag(.Prijs(Productnr), .Hoeveelheid(Productnr), .Meeteenheid(Productnr))
         
         BuffetVenster.TabelVeld(BuffettabelVeld(Productnr, Kolom:=BKKostprijs)).Caption = .Kostprijs(Productnr)
      Next Productnr
      
      BuffetBerekenTotaal
       
      BuffetVenster.TotaalVeld.Caption = .Totaalkosten
   End With
   
EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure centreert het opgegeven venster met de opgegeven verticale afwijking.
Public Sub CentreerVenster(Venster As Form, VensterNaam As String, Optional YAfwijking As Long = 4)
On Error GoTo Fout

   With Venster
      If .WindowState = vbNormal Then
         .Left = (KCalcVenster.Width / 2) - (.Width / 2)
         .Top = (KCalcVenster.Height / YAfwijking) - (.Height / 2)
      End If
   End With
   
EindeProcedure:
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure controleert of de datamap aanwezig is.
Private Sub ControleerDataMap()
On Error GoTo DataMapFout
Dim Foutcode As Long
Dim Omschrijving As String

   ChDir DATA_MAP
   Exit Sub

DataMapFout:
   Foutcode = Err.Number
   Omschrijving = Err.Description
   
   If Foutcode = 76 Then
      MaakDatamap
   Else
      HandelDataMapFoutAf Foutcode, Omschrijving
   End If
End Sub

'Deze procedure drukt het huidige gerecht af.
Public Sub DrukGerechtAf()
On Error GoTo Fout
Dim Personeelskostenrij As Long
   
   Screen.MousePointer = vbHourglass
   
   Printer.Print Gerecht.Datum
   
   VVDrukAf Vlees, "Vlees"
   VVDrukAf Vis, "Vis"
   BBDGMSTDrukAf Groenten, "Groenten"
   BBDGMSTDrukAf Meelspijzen, "Meelspijzen"
   BBDGMSTDrukAf SoepenEnSausen, "SoepenEnSausen"
   BBDGMSTDrukAf Bijproducten, "Bijproducten"
   BBDGMSTDrukAf Doorverkoopproducten, "Doorverkoopproducten"
   BBDGMSTDrukAf Bittergarnituur, "Bittergarnituur"
   BBDGMSTDrukAf Tafelgarnituur, "Tafelgarnituur"
   
   With Personeelskosten
      Printer.Print
      Printer.Print Tab(3); "Personeelskosten:"
      Printer.Print String(75, "-")
      Printer.Print Tab(3); "Loon:"; Tab(10); "Tijd:"
      Printer.Print String(75, "-")
      For Personeelskostenrij = LBound(.Loon()) To UBound(.Loon())
         Printer.Print Tab(7 - Len(.Loon(Personeelskostenrij))); .Loon(Personeelskostenrij);
         Printer.Print Tab(14 - Len(.Tijd(Personeelskostenrij))); .Tijd(Personeelskostenrij)
      Next Personeelskostenrij
      Printer.Print String(75, "-")
      Printer.Print Tab(8 - Len(.Totaalkosten)); .Totaalkosten
   End With
   
   Printer.Print
   Printer.Print Tab(3); "Kosten:"
   Printer.Print String(75, "-")
   Printer.Print Tab(3); "Gerecht:"; Tab(13); Gerecht.Naam
   Printer.Print String(75, "-")
   Printer.Print Tab(3); "Vlees:"; Tab(30 - Len(Vlees.Kostprijs)); Vlees.Kostprijs
   Printer.Print Tab(3); "Vis:"; Tab(30 - Len(Vis.Kostprijs)); Vis.Kostprijs
   Printer.Print Tab(3); "Groenten:"; Tab(30 - Len(Groenten.Totaalkosten)); Groenten.Totaalkosten
   Printer.Print Tab(3); "Meelspijzen:"; Tab(30 - Len(Meelspijzen.Totaalkosten)); Meelspijzen.Totaalkosten
   Printer.Print Tab(3); "Soepen en Sausen:"; Tab(30 - Len(SoepenEnSausen.Totaalkosten)); SoepenEnSausen.Totaalkosten
   Printer.Print Tab(3); "Bijproducten:"; Tab(30 - Len(Bijproducten.Totaalkosten)); Bijproducten.Totaalkosten
   Printer.Print Tab(3); "Doorverkoopproducten:"; Tab(30 - Len(Doorverkoopproducten.Totaalkosten)); Doorverkoopproducten.Totaalkosten
   Printer.Print Tab(3); "Bittergarnituur:"; Tab(30 - Len(Bittergarnituur.Totaalkosten)); Bittergarnituur.Totaalkosten
   Printer.Print Tab(3); "Tafelgarnituur:"; Tab(30 - Len(Tafelgarnituur.Totaalkosten)); Tafelgarnituur.Totaalkosten
   Printer.Print Tab(3); "Personeelskosten:"; Tab(30 - Len(Personeelskosten.Totaalkosten)); Personeelskosten.Totaalkosten
   Printer.Print String(75, "-")
   Printer.Print Tab(3); "Totaal:"; Tab(30 - Len(Gerecht.Totaalkosten)); Gerecht.Totaalkosten
   
   Printer.EndDoc
   
EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure berekent de kosten voor het huidige gerecht.
Public Sub GerechtBerekenKosten()
On Error GoTo Fout

   Screen.MousePointer = vbHourglass
   
   VVBerekenKostprijs Vlees
   VVBerekenKostprijs Vis
   BBDGMSTBerekenTotaal Groenten
   BBDGMSTBerekenTotaal Meelspijzen
   BBDGMSTBerekenTotaal SoepenEnSausen
   BBDGMSTBerekenTotaal Bijproducten
   BBDGMSTBerekenTotaal Doorverkoopproducten
   BBDGMSTBerekenTotaal Bittergarnituur
   BBDGMSTBerekenTotaal Tafelgarnituur
   PersoneelskostenBerekenTotaal
   
   With Gerecht
      .Totaalkosten = "0.00"
      .Totaalkosten = TelBedragenOp(.Totaalkosten, Vlees.Kostprijs)
      .Totaalkosten = TelBedragenOp(.Totaalkosten, Vis.Kostprijs)
      .Totaalkosten = TelBedragenOp(.Totaalkosten, Groenten.Totaalkosten)
      .Totaalkosten = TelBedragenOp(.Totaalkosten, Meelspijzen.Totaalkosten)
      .Totaalkosten = TelBedragenOp(.Totaalkosten, SoepenEnSausen.Totaalkosten)
      .Totaalkosten = TelBedragenOp(.Totaalkosten, Bijproducten.Totaalkosten)
      .Totaalkosten = TelBedragenOp(.Totaalkosten, Doorverkoopproducten.Totaalkosten)
      .Totaalkosten = TelBedragenOp(.Totaalkosten, Bittergarnituur.Totaalkosten)
      .Totaalkosten = TelBedragenOp(.Totaalkosten, Tafelgarnituur.Totaalkosten)
      .Totaalkosten = TelBedragenOp(.Totaalkosten, Personeelskosten.Totaalkosten)
      .Totaalkosten = RondBedragAf(.Totaalkosten)
   End With
   
   If VensterActief("KostenVenster") Then
      With KostenVenster
         .GerechtVeld.Caption = Gerecht.Naam
         .VleesKostprijsVeld.Caption = RondBedragAf(Vlees.Kostprijs)
         .VisKostprijsVeld.Caption = RondBedragAf(Vis.Kostprijs)
         .GroentenTotaalVeld.Caption = RondBedragAf(Groenten.Totaalkosten)
         .MeelspijzenTotaalVeld.Caption = RondBedragAf(Meelspijzen.Totaalkosten)
         .SoepenEnSausenTotaalVeld.Caption = RondBedragAf(SoepenEnSausen.Totaalkosten)
         .BijproductenTotaalVeld.Caption = RondBedragAf(Bijproducten.Totaalkosten)
         .DoorverkoopproductenTotaalVeld.Caption = RondBedragAf(Doorverkoopproducten.Totaalkosten)
         .BittergarnituurTotaalVeld.Caption = RondBedragAf(Bittergarnituur.Totaalkosten)
         .TafelgarnituurTotaalVeld.Caption = RondBedragAf(Tafelgarnituur.Totaalkosten)
         .PersoneelskostenTotaalVeld.Caption = RondBedragAf(Personeelskosten.Totaalkosten)
         .TotaalVeld.Caption = Gerecht.Totaalkosten
      End With
   End If
   
EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure handelt eventuele datamap gerelateerde fouten af.
Private Sub HandelDataMapFoutAf(Foutcode As Long, Omschrijving As String)
On Error GoTo Fout
Dim Bericht As String
   
   On Error Resume Next
   Bericht = "Datamap fout:" & vbCr
   Bericht = Bericht & Omschrijving & "." & vbCr
   Bericht = Bericht & "Foutcode: " & CStr(Foutcode)

   Select Case MsgBox(Bericht, vbExclamation Or vbOKCancel Or vbDefaultButton1)
      Case vbCancel
         End
      Case vbOK
         Resume
   End Select
   
EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure handelt eventuele fouten af en stuurt de door de gebruiker gekozen actie terug.
Public Function HandelFoutAf(Optional VraagVorigeKeuzeOp As Boolean = False, Optional ActiefBestand As String = vbNullString) As Long
Dim Bericht As String
Dim Foutcode As Long
Dim Muisaanwijzerstatus As Integer
Dim Omschrijving As String
Static Keuze As Long

   Foutcode = Err.Number
   Omschrijving = Err.Description
   
   On Error Resume Next
   If Not VraagVorigeKeuzeOp Then
      Muisaanwijzerstatus = Screen.MousePointer
      Screen.MousePointer = vbDefault
      
      Bericht = Omschrijving & "." & vbCr
      If Not ActiefBestand = vbNullString Then Bericht = Bericht & "Actief bestand: """ & ActiefBestand & """" & vbCr
      Bericht = Bericht & "Foutcode: " & CStr(Foutcode)
      
      Keuze = MsgBox(Bericht, vbExclamation Or vbAbortRetryIgnore Or vbDefaultButton2)
      Select Case Keuze
         Case vbAbort
            End
         Case vbIgnore
            On Error Resume Next
            Close
            Reset
         Case vbRetry
            Screen.MousePointer = Muisaanwijzerstatus
      End Select
   End If
   
   HandelFoutAf = Keuze
End Function

'Deze procedure handelt eventuele printerfouten af.
Private Function HandelPrinterFoutAf() As Long
Dim Bericht As String
Dim Foutcode As Long
Dim Keuze As Long
Dim Omschrijving As String

   Foutcode = Err.Number
   Omschrijving = Err.Description
   
   On Error Resume Next
   Bericht = "Printer fout:" & vbCr
   Bericht = Bericht & Omschrijving & "." & vbCr
   Bericht = Bericht & vbCr & "Foutcode: " & CStr(Foutcode)
   
   Keuze = MsgBox(Bericht, vbExclamation Or vbAbortRetryIgnore Or vbDefaultButton2)
      
   HandelPrinterFoutAf = Keuze
End Function

'Deze procedure kopieeert backbestanden van de opgegeven locatie naar het opgegeven doel.
Public Sub KopieerBackupbestanden(DataLocatie As String, DataDoel As String)
On Error GoTo Fout
Dim Bericht As String
Dim Bestand As String
   
   Screen.MousePointer = vbHourglass

   If LCase$(DataLocatie) = LCase$(DataDoel) Then
      Bericht = App.Title & " bevindt zich op de opgegeven backuplocatie." & vbCr
      Bericht = Bericht & "Kies een andere backuplocatie."
      MsgBox Bericht, vbExclamation
      Screen.MousePointer = vbDefault
   Else
      If Not Dir$(DataDoel & "*.kcc", vbArchive Or vbHidden Or vbSystem) = vbNullString Then
         Kill DataDoel & "*.kcc"
      End If
   
      Bestand = Dir$(DataLocatie & "*.kcc", vbArchive Or vbHidden Or vbSystem)
      If Bestand = vbNullString Then
         MsgBox "Geen bestanden om te kopi" & E_TREMA & "ren aanwezig.", vbExclamation
      Else
         Do Until Bestand = vbNullString
            FileCopy DataLocatie & Bestand, DataDoel & Bestand
            Bestand = Dir$()
         Loop
      End If
   
      MsgBox "De bestanden zijn gekopieerd.", vbInformation
   End If
   
EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub
   
Fout:
   If HandelFoutAf(, ActiefBestand:=Bestand) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure laadt het buffet.
Public Sub LaadBuffet()
On Error GoTo Fout
Dim BestandH As Integer
Dim Lengte As Long
Dim Productnr As Long
   
   Screen.MousePointer = vbHourglass
   
   If Not Dir$(DATA_MAP & "Buffet.kcc", vbArchive Or vbHidden Or vbSystem) = vbNullString Then
      BestandH = FreeFile()
      Open DATA_MAP & "Buffet.kcc" For Binary Lock Read Write As BestandH
         With Buffet
            Lengte = Asc(Input$(1, BestandH)): .Productrij = Asc(Input$(Lengte, BestandH))
            For Productnr = LBound(.Product()) To UBound(.Product())
               Lengte = Asc(Input$(1, BestandH)): .Product(Productnr) = Input$(Lengte, BestandH)
               Lengte = Asc(Input$(1, BestandH)): .Prijs(Productnr) = Input$(Lengte, BestandH)
               Lengte = Asc(Input$(1, BestandH)): .Hoeveelheid(Productnr) = Input$(Lengte, BestandH)
               .Meeteenheid(Productnr) = Asc(Input$(1, BestandH))
                .Kostprijs(Productnr) = BerekenBedrag(.Prijs(Productnr), .Hoeveelheid(Productnr), .Meeteenheid(Productnr))
            Next Productnr
         End With
      Close BestandH
   End If

EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub

Fout:
   If HandelFoutAf(, ActiefBestand:="het buffetbestand") = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure laadt het opgegeven gerechtbestand.
Public Sub LaadGerecht(Gerechtbestand As String)
On Error GoTo Fout
Dim BestandH As Integer
Dim Lengte As Long
Dim Personeelskostenrij As Long

   Screen.MousePointer = vbHourglass

   If Dir$(DATA_MAP & Gerechtbestand, vbArchive Or vbHidden Or vbSystem) = vbNullString Then
      Error 53
   Else
      BestandH = FreeFile()
      Open DATA_MAP & Gerechtbestand For Binary Lock Read Write As BestandH
         With Gerecht
            Lengte = Asc(Input$(1, BestandH)): .Naam = Input$(Lengte, BestandH)
            Lengte = Asc(Input$(1, BestandH)): .Datum = Input$(Lengte, BestandH)
         End With
         
         VVLaadGerecht Vlees, BestandH
         VVLaadGerecht Vis, BestandH
         BBDGMSTLaadGerecht Groenten, BestandH
         BBDGMSTLaadGerecht Meelspijzen, BestandH
         BBDGMSTLaadGerecht SoepenEnSausen, BestandH
         BBDGMSTLaadGerecht Bijproducten, BestandH
         BBDGMSTLaadGerecht Doorverkoopproducten, BestandH
         BBDGMSTLaadGerecht Bittergarnituur, BestandH
         BBDGMSTLaadGerecht Tafelgarnituur, BestandH
         
         With Personeelskosten
            For Personeelskostenrij = LBound(.Tijd()) To UBound(.Tijd())
               Lengte = Asc(Input$(1, BestandH)): .Tijd(Personeelskostenrij) = Val(Input$(Lengte, BestandH))
            Next Personeelskostenrij
         End With
      Close BestandH
      
      With Gerecht
         .HuidigBestand = Gerechtbestand
         .WijzigingenBewaard = True
      End With
   End If
   
EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub
   
Fout:
   If HandelFoutAf(, ActiefBestand:="het gerechtbestand") = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure laadt het opgegeven handleidingonderwerp en toont deze in het opgegeven veld.
Public Sub LaadHandleidingonderwerp(GeselecteerdOnderwerpnr As Long, HulptekstVeld As PictureBox)
On Error GoTo Fout
Dim BestandH As Integer
Dim Onderwerpnr As Long
Dim Tekst As String

   Screen.MousePointer = vbHourglass
   
   HulptekstVeld.FontBold = False
   HulptekstVeld.FontItalic = False
   HulptekstVeld.FontUnderline = False
   
   HulptekstVeld.Cls
   Onderwerpnr = 0
   
   BestandH = FreeFile()
   Open "Handleiding.kcc" For Input Lock Read Write As BestandH
      Do Until EOF(BestandH)
         Line Input #BestandH, Tekst
         If Left$(Tekst, 1) = "-" Then Onderwerpnr = Onderwerpnr + 1
         
         If Onderwerpnr > GeselecteerdOnderwerpnr Then
            Exit Do
         ElseIf Onderwerpnr = GeselecteerdOnderwerpnr Then
            ToonTekstMetOpMaak Unescape(Mid$(Tekst, 2), EscapeTeken:="\"), HulptekstVeld
         End If
      Loop
   Close BestandH
   
EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub

Fout:
   If HandelFoutAf(, ActiefBestand:="het handleidingbestand") = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure laadt de handleidingonderwerpen.
Public Sub LaadHandleidingonderwerpen()
On Error GoTo Fout
Dim BestandH As Integer
Dim Onderwerpnr As Long
Dim Regel As String

   Screen.MousePointer = vbHourglass
   
   BestandH = FreeFile()
   Open "Handleiding.kcc" For Input Lock Read Write As BestandH
      Onderwerpnr = 1
      Do While DoEvents() > 0
         Line Input #BestandH, Regel
         If Regel = vbNullString Then Exit Do
         If Onderwerpnr > 1 Then Load HandleidingVenster.OnderwerpenMenu(Onderwerpnr)
         HandleidingVenster.OnderwerpenMenu(Onderwerpnr).Caption = Regel
         Onderwerpnr = Onderwerpnr + 1
      Loop
   Close BestandH
   
EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub

Fout:
   If HandelFoutAf(, ActiefBestand:="het handleidingbestand") = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure laadt de personeelskosten.
Private Sub LaadPersoneelskosten()
On Error GoTo Fout
Dim BestandH As Integer
Dim Lengte As Long
Dim Personeelskostenrij As Long

   Screen.MousePointer = vbHourglass
   
   If Not Dir$(DATA_MAP & "Personeel.kcc", vbArchive Or vbHidden Or vbSystem) = vbNullString Then
      BestandH = FreeFile()
      Open DATA_MAP & "Personeel.kcc" For Binary Lock Read Write As BestandH
         With Personeelskosten
            For Personeelskostenrij = LBound(.Loon()) To UBound(.Loon())
               Lengte = Asc(Input$(1, BestandH)): .Loon(Personeelskostenrij) = Val(Input$(Lengte, BestandH))
            Next Personeelskostenrij
         End With
      Close BestandH
   End If
   
EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub
   
Fout:
   If HandelFoutAf(, ActiefBestand:="het personeelskostenbestand") = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure laadt en stuurt de printerinstellingen terug.
Private Function LaadPrinterinstellingen() As PrinterlettertypeStr
On Error GoTo Fout
Dim BestandH As Integer
Dim Lengte As Long
Dim Lettertype As PrinterlettertypeStr

   Screen.MousePointer = vbHourglass
   
   If Not Dir$("Printer.kcc", vbArchive Or vbHidden Or vbSystem) = vbNullString Then
      BestandH = FreeFile()
      Open "Printer.kcc" For Binary Lock Read Write As BestandH
         With Lettertype
            .Cursief = Asc(Input$(1, BestandH))
            .Onderstreept = Asc(Input$(1, BestandH))
            .Vet = Asc(Input$(1, BestandH))
            .Formaat = Asc(Input$(1, BestandH))
            Lengte = Asc(Input$(1, BestandH)): .Naam = Input$(Lengte, BestandH)
         End With
      Close BestandH
   End If
   
EindeProcedure:
   Screen.MousePointer = vbDefault
   LaadPrinterinstellingen = Lettertype
   Exit Function

Fout:
   If HandelFoutAf(, ActiefBestand:="het printerinstellingenbestand") = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Function

'Deze procedure laadt het productenbestand.
Public Sub LaadProducten()
On Error GoTo Fout
Dim BestandH As Integer

   Screen.MousePointer = vbHourglass
   
   If Not Dir$(DATA_MAP & "Producten.kcc", vbArchive Or vbHidden Or vbSystem) = vbNullString Then
      BestandH = FreeFile()
      Open DATA_MAP & "Producten.kcc" For Binary Lock Read Write As BestandH
         VVLaadProducten Vlees, BestandH
         VVLaadProducten Vis, BestandH
         BBDGMSTLaadProducten Groenten, BestandH
         BBDGMSTLaadProducten Meelspijzen, BestandH
         BBDGMSTLaadProducten SoepenEnSausen, BestandH
         BBDGMSTLaadProducten Bijproducten, BestandH
         BBDGMSTLaadProducten Doorverkoopproducten, BestandH
         BBDGMSTLaadProducten Bittergarnituur, BestandH
         BBDGMSTLaadProducten Tafelgarnituur, BestandH
      Close BestandH
   End If
   
EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub
   
Fout:
   If HandelFoutAf(, ActiefBestand:="het productenbestand") = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure maakt de datamap.
Private Sub MaakDatamap()
On Error GoTo DataMapFout
Dim Bericht As String
Dim Foutcode As Long
Dim Omschrijving As String

   MkDir DATA_MAP
   Exit Sub
   
DataMapFout:
   Foutcode = Err.Number
   Omschrijving = Err.Description

   On Error Resume Next
   Bericht = "Kan de datamap """ & DATA_MAP & """ niet maken." & vbCr
   Bericht = Bericht & Omschrijving & "." & vbCr
   Bericht = Bericht & "Foutcode: " & CStr(Foutcode)
   
   Select Case MsgBox(Bericht, vbExclamation Or vbOKCancel Or vbDefaultButton1)
      Case vbCancel
         End
      Case vbOK
         Resume
   End Select
End Sub

'Deze procedure stelt dit programma in wanneer dit program gestart wordt.
Private Sub Main()
On Error GoTo Fout
   
   ChDrive Left$(App.Path, InStr(App.Path, ":"))
   ChDir App.Path
   
   ControleerDataMap
   ChDir App.Path
   
   StelProgrammaIn
   StelGerechtIn IsNieuwGerecht:=True
   GerechtBerekenKosten
   StelPrinterIn
   LaadPersoneelskosten
   LaadProducten
   LaadBuffet
   
   KCalcVenster.Show
   WerkMeeteenhedenmenuBij GeselecteerdeMeeteenheid:=0, DecilitersZichtbaar:=False
   
   Do While DoEvents() > 0
      If Not VensterActief("MenuVenster") Then Unload KCalcVenster
   Loop

   BewaarPersoneelskosten
   BewaarProducten
   BewaarBuffet

EindeProcedure:
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure stuurt de in het meeteenhedenmenu geselecteerde meeteenheid terug.
Private Function Meeteenhedenmenuselectie() As Long
On Error GoTo Fout
Dim Meeteenheidnr As Long
Dim Selectie As Long

   Selectie = GEEN_MEETEENHEID
   For Meeteenheidnr = KCalcVenster.MeeteenheidMenu.LBound To KCalcVenster.MeeteenheidMenu.UBound
      If KCalcVenster.MeeteenheidMenu(Meeteenheidnr).Checked Then
         Selectie = Meeteenheidnr
         Exit For
      End If
   Next Meeteenheidnr
  
EindeProcedure:
   Meeteenhedenmenuselectie = Selectie
   Exit Function
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Function

'Deze procedure stuurt een nieuw gerechtnummer terug.
Private Function NieuwGerechtnr() As Long
On Error GoTo Fout
Dim Bestand As String
Dim Gerechtnr As Long
Dim NieuwNr As Long

   Screen.MousePointer = vbHourglass
   
   NieuwNr = GEEN_GERECHTNR
   For Gerechtnr = 0 To MAX_AANTAL_GERECHTEN - 1
      Bestand = Dir$(DATA_MAP & "Gerecht" & Gerechtnr & ".kcc", vbArchive Or vbHidden Or vbSystem)
      
      If Bestand = vbNullString Then
         NieuwNr = Gerechtnr
         Exit For
      End If
   Next Gerechtnr
   
   If NieuwNr = GEEN_GERECHTNR Then
      MsgBox "Het maximale aantal gerechten is bereikt." & vbCr & "Het laatst aangemaakte gerecht wordt geladen.", vbExclamation
      NieuwNr = MAX_AANTAL_GERECHTEN - 1
   End If
   
EindeProcedure:
   Screen.MousePointer = vbDefault
   NieuwGerechtnr = NieuwNr
   Exit Function

Fout:
   If HandelFoutAf(, ActiefBestand:="het gerechtbestand") = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Function

'Deze procedure berekent de personeelskostentotalen.
Private Sub PersoneelskostenBerekenTotaal()
On Error GoTo Fout
Dim Personeelskostenrij As Long

   With Personeelskosten
      .Totaalkosten = "0.00"
      
      For Personeelskostenrij = LBound(.Loon()) To UBound(.Loon())
         .Totaalkosten = TelBedragenOp(.Totaalkosten, BerekenBedrag(.Loon(Personeelskostenrij), .Tijd(Personeelskostenrij), METijd))
      Next Personeelskostenrij
   End With

EindeProcedure:
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure stelt de personeelskosten in.
Private Sub PersoneelskostenStelIn(Optional ResetAlleenTijden As Boolean = True)
On Error GoTo Fout
Dim Personeelskostenrij As Long

   With Personeelskosten
      If Not ResetAlleenTijden Then ReDim .Loon(0 To 7) As String
      ReDim .Tijd(0 To 7) As String
   
      .Totaalkosten = "0.00"
      For Personeelskostenrij = LBound(.Tijd()) To UBound(.Tijd())
         If Not ResetAlleenTijden Then .Loon(Personeelskostenrij) = "0.00"
         .Tijd(Personeelskostenrij) = "0"
      Next Personeelskostenrij
   End With

EindeProcedure:
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure stelt het personeelskosten venster in.
Public Sub PersoneelskostenStelVensterIn()
On Error GoTo Fout
Dim Personeelskostenrij As Long

   Screen.MousePointer = vbHourglass
 
   If VensterActief("PersoneelskostenVenster") Then
      With Personeelskosten
         For Personeelskostenrij = LBound(.Loon()) To UBound(.Loon())
            PersoneelskostenVenster.LoonVeld(Personeelskostenrij).Text = RondBedragAf(.Loon(Personeelskostenrij))
            PersoneelskostenVenster.TijdVeld(Personeelskostenrij).Text = .Tijd(Personeelskostenrij)
         Next Personeelskostenrij
      
         PersoneelskostenBerekenTotaal
         
         PersoneelskostenVenster.TotaalVeld = Personeelskosten.Totaalkosten
      End With
   End If
   
EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure verwerkt de personeelskosteninvoer.
Public Sub PersoneelskostenVerwerkInvoer()
On Error GoTo Fout
Dim Personeelskostenrij As Long

   With Personeelskosten
      For Personeelskostenrij = LBound(.Loon()) To UBound(.Loon())
         PersoneelskostenVenster.LoonVeld(Personeelskostenrij).Text = RondBedragAf(PersoneelskostenVenster.LoonVeld(Personeelskostenrij).Text)
         PersoneelskostenVenster.TijdVeld(Personeelskostenrij).Text = CStr(Val(PersoneelskostenVenster.TijdVeld(Personeelskostenrij).Text))
         
         .Loon(Personeelskostenrij) = PersoneelskostenVenster.LoonVeld(Personeelskostenrij).Text
         .Tijd(Personeelskostenrij) = PersoneelskostenVenster.TijdVeld(Personeelskostenrij).Text
      Next Personeelskostenrij
         
      PersoneelskostenBerekenTotaal
      
      PersoneelskostenVenster.TotaalVeld.Caption = .Totaalkosten
   End With

EindeProcedure:
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure stuurt de programmainformatie terug.
Public Function Programmainformatie() As String
On Error GoTo Fout
Dim Informatie As String

   With App
      Informatie = .Title & vbCr & "Versie " & .Major & "." & .Minor & .Revision & vbCr & "Door: " & .CompanyName & vbCr & "***1999***"
   End With
   
EindeProcedure:
   Programmainformatie = Informatie
   Exit Function

Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Function

'Deze procedure rondt het opgegeven bedrag af en stuurt deze terug.
Private Function RondBedragAf(Bedrag As String) As String
On Error GoTo Fout
Dim AfgerondBedrag As String

   AfgerondBedrag = Replace(Format$(Round(Val(Replace(Bedrag, ",", ".")), 2), "####0.00"), ",", ".")
   
EindeProcedure:
   RondBedragAf = AfgerondBedrag
   Exit Function
   
Fout:
   AfgerondBedrag = "0.00"
   Resume EindeProcedure
End Function

'Deze procedure selecteert het opgegeven item in de opgegeven lijst.
Public Sub SelecteerLijstitem(Lijst As ListBox, Item As String)
On Error GoTo Fout
Dim Index As Long
   
   Screen.MousePointer = vbHourglass
   
   For Index = 0 To Lijst.ListCount - 1
      If Lijst.List(Index) = Item Then
         Lijst.ListIndex = Index
         Exit For
      End If
   Next Index
   
EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure toont het meeteenhedenmenu en stuurt de keuze van de gebruiker terug.
Public Function SelecteerMeeteenheid(HuidigeProductmeeteenheid As Long, Venster As Form, Optional DecilitersZichtbaar As Boolean = False) As Long
On Error GoTo Fout
Dim NieuweMeeteenheid As Long

   WerkMeeteenhedenmenuBij HuidigeProductmeeteenheid, DecilitersZichtbaar:=DecilitersZichtbaar
   
   Venster.PopupMenu KCalcVenster.MeeteenhedenHoofdmenu
   NieuweMeeteenheid = Meeteenhedenmenuselectie()
   
   WerkMeeteenhedenmenuBij NieuweMeeteenheid, DecilitersZichtbaar:=DecilitersZichtbaar
   
EindeProcedure:
   Screen.MousePointer = vbDefault
   SelecteerMeeteenheid = NieuweMeeteenheid
   Exit Function
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Function

'Deze procedure stuurt de eventuele stationsletter in het opgegeven pad terug.
Public Function Stationsletter(Pad As String) As String
On Error GoTo Fout
Dim Letter As String
Dim Positie As Long

   Letter = vbNullString
   Positie = InStr(Pad, ":")
   If Positie > 0 Then Letter = Left$(Pad, Positie)
   
EindeProcedure:
   Stationsletter = Letter
   Exit Function
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Function

'Deze procedure stelt het huidige gerecht in.
Public Sub StelGerechtIn(Optional IsNieuwGerecht As Boolean = False)
On Error GoTo Fout
   
   With Gerecht
      .WijzigingenBewaard = True
      
      If IsNieuwGerecht Then
         .Datum = CStr(Day(Date)) & "/" & CStr(Month(Date)) & "/" & CStr(Year(Date))
         .HuidigNr = NieuwGerechtnr()
         .HuidigBestand = "Gerecht" & .HuidigNr & ".kcc"
         .Naam = "gerecht " & .HuidigNr
      End If
   End With
          
   VVStelIn Vlees
   VVStelIn Vis
   BBDGMSTStelIn Groenten
   BBDGMSTStelIn Meelspijzen
   BBDGMSTStelIn SoepenEnSausen
   BBDGMSTStelIn Bijproducten
   BBDGMSTStelIn Doorverkoopproducten
   BBDGMSTStelIn Bittergarnituur
   BBDGMSTStelIn Tafelgarnituur
   PersoneelskostenStelIn
  
EindeProcedure:
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht om de gerechtsvensters in te stellen.
Public Sub StelGerechtvenstersIn()
On Error GoTo Fout

   VVStelVensterIn Vlees, VleesVenster, "VleesVenster"
   VVStelVensterIn Vis, VisVenster, "VisVenster"
   BBDGMSTStelVensterIn Groenten, GroentenVenster, "GroentenVenster"
   BBDGMSTStelVensterIn Meelspijzen, MeelspijzenVenster, "MeelspijzenVenster"
   BBDGMSTStelVensterIn SoepenEnSausen, SoepenEnSausenVenster, "SoepenEnSausenVenster"
   BBDGMSTStelVensterIn Bijproducten, BijproductenVenster, "BijproductenVenster"
   BBDGMSTStelVensterIn Doorverkoopproducten, DoorverkoopproductenVenster, "DoorverkoopproductenVenster"
   BBDGMSTStelVensterIn Bittergarnituur, BittergarnituurVenster, "BittergarnituurVenster"
   BBDGMSTStelVensterIn Tafelgarnituur, TafelgarnituurVenster, "TafelgarnituurVenster"
   PersoneelskostenStelVensterIn
   GerechtBerekenKosten

EindeProcedure:
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure stelt de printer in.
Public Sub StelPrinterIn()
On Error GoTo Fout
Dim Bericht As String
Dim Lettertype As PrinterlettertypeStr
   
   If Printers.Count = 0 Then
      MsgBox "Geen printer aanwezig.", vbExclamation
   Else
      Screen.MousePointer = vbHourglass
      
      Lettertype = LaadPrinterinstellingen()
      
      If Lettertype.Naam = vbNullString Then
         MsgBox "Er zijn geen opgeslagen printerinstellingen gevonden." & vbCr & "De huidige printerinstellingen worden gebruikt.", vbInformation
         BewaarPrinterinstelingen
      Else
         On Error GoTo PrinterFout
         
         Printerlettertype = Lettertype
         
         Printer.FontItalic = Printerlettertype.Cursief
         Printer.FontUnderline = Printerlettertype.Onderstreept
         Printer.FontBold = Printerlettertype.Vet
         Printer.FontSize = Printerlettertype.Formaat
         Printer.FontName = Printerlettertype.Naam
            
         On Error GoTo Fout
      End If
   End If

EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub

Fout:
   If HandelFoutAf(, ActiefBestand:="het printerinstellingenbestand") = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
   Exit Sub

PrinterFout:
   Screen.MousePointer = vbDefault
   Select Case HandelPrinterFoutAf()
      Case vbAbort
         End
      Case vbRetry
         Screen.MousePointer = vbHourglass
         Resume
      Case vbIgnore
         Resume EindeProcedure
   End Select
End Sub

'Deze procedure stelt dit programma in.
Public Sub StelProgrammaIn()
On Error GoTo Fout

   Screen.MousePointer = vbHourglass
    
   VVStelIn Vlees, AantalProducten:=100
   VVStelIn Vis, AantalProducten:=100
   BBDGMSTStelIn Groenten, AantalProducten:=50
   BBDGMSTStelIn Meelspijzen, AantalProducten:=50
   BBDGMSTStelIn SoepenEnSausen, AantalProducten:=50
   BBDGMSTStelIn Bijproducten, AantalProducten:=50
   BBDGMSTStelIn Doorverkoopproducten, AantalProducten:=25
   BBDGMSTStelIn Bittergarnituur, AantalProducten:=50
   BBDGMSTStelIn Tafelgarnituur, AantalProducten:=25
   BuffetStelIn AantalProducten:=299
   PersoneelskostenStelIn ResetAlleenTijden:=False
   
   With Printerlettertype
      .Cursief = Printer.FontItalic
      .Formaat = Printer.FontSize
      .Naam = Printer.FontName
      .Onderstreept = Printer.FontUnderline
      .Vet = Printer.FontBold
   End With
   
   ReDim GerechtBestanden(0 To 0) As GerechtbestandStr
   
EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure telt de twee opgegeven bedragen bij elkaar op en stuurt het resultaat terug.
Private Function TelBedragenOp(Bedrag1 As String, Bedrag2 As String) As String
On Error GoTo Fout
Dim NieuwBedrag As String

   NieuwBedrag = RondBedragAf(Replace(CStr(Val(Replace(Bedrag1, ",", ".")) + Val(Replace(Bedrag2, ",", "."))), ",", "."))
   
EindeProcedure:
   TelBedragenOp = NieuwBedrag
   Exit Function
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Function

'Deze procedure toont de printerlettertypen in de opgegeven lijst.
Public Sub ToonPrinterlettertypen(Lijst As ListBox)
On Error GoTo Fout
Dim Lettertypenr As Long

   Screen.MousePointer = vbHourglass
   
   With Printer
      Lijst.Clear
      For Lettertypenr = 0 To .FontCount - 1
         Lijst.AddItem .Fonts(Lettertypenr)
      Next Lettertypenr
   End With
   
EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure toont tekst met eventuele opmaak in het opgegeven veld.
Private Sub ToonTekstMetOpMaak(Tekst As String, Veld As PictureBox)
On Error GoTo Fout
Dim Teken As String
Dim Tekennr As Long
Dim VolgendTeken As String

   Screen.MousePointer = vbHourglass

   With Veld
      Tekennr = 1
      Do Until Tekennr > Len(Tekst) And DoEvents() > 0
         Teken = Mid(Tekst, Tekennr, 1)
         VolgendTeken = Mid(Tekst, Tekennr + 1, 1)
         Tekennr = Tekennr + 1
         
         Select Case Teken
            Case vbTab
               Veld.Print ,
            Case "*", "/", "_"
               If Teken = VolgendTeken Then
                  Veld.Print Teken;
                  Tekennr = Tekennr + 1
               Else
                  Select Case Teken
                     Case "*"
                        .FontBold = Not .FontBold
                     Case "/"
                        .FontItalic = Not .FontItalic
                     Case "_"
                        .FontUnderline = Not .FontUnderline
                  End Select
               End If
            Case Else
               Veld.Print Teken;
         End Select
      Loop
      
      Veld.Print
   End With
   
EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure unescapet de opgegeven tekst en stuurt deze terug.
Private Function Unescape(Tekst As String, Optional EscapeTeken As String = "/", Optional Foutpositie As Long = 0) As String
On Error GoTo Fout
Dim Hexadecimalen As String
Dim Index As Long
Dim Teken As String
Dim Unescaped As String
Dim VolgendTeken As String

   Foutpositie = 0
   Index = 1
   Unescaped = vbNullString
   Do Until Index > Len(Tekst)
      Teken = Mid$(Tekst, Index, 1)
      VolgendTeken = Mid$(Tekst, Index + 1, 1)
   
      If Teken = EscapeTeken Then
         If VolgendTeken = EscapeTeken Then
            Unescaped = Unescaped & Teken
            Index = Index + 1
         Else
            Hexadecimalen = UCase$(Mid$(Tekst, Index + 1, 2))
            If Len(Hexadecimalen) = 2 Then
               If Left$(Hexadecimalen, 1) = "0" Then Hexadecimalen = Right$(Hexadecimalen, 1)
      
               If UCase$(Hex$(CLng(Val("&H" & Hexadecimalen & "&")))) = Hexadecimalen Then
                  Unescaped = Unescaped & Chr$(CLng(Val("&H" & Hexadecimalen & "&")))
                  Index = Index + 2
               Else
                  Foutpositie = Index
                  Exit Do
               End If
            Else
               Foutpositie = Index
               Exit Do
            End If
         End If
      Else
         Unescaped = Unescaped & Teken
      End If
      Index = Index + 1
   Loop
   
EindeProcedure:
   Unescape = Unescaped
   Exit Function
  
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Function

'Deze procedure geeft aan of het opgegeven venster actief is.
Private Function VensterActief(VensterNaam As String) As Boolean
On Error GoTo Fout
Dim Actief As Boolean
Dim Venster As Form

   Actief = False
   For Each Venster In Forms
      If LCase$(Venster.Name) = LCase$(VensterNaam) Then
         Actief = True
         Exit For
      End If
   Next Venster

EindeProcedure:
   VensterActief = Actief
   Exit Function
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Function

'Deze procedure plaatst het opgegeven venster naast het programmamenu.
Public Sub VensterNaastMenu(Venster As Form, VensterNaam As String)
On Error GoTo Fout

   With Venster
      If .WindowState = vbNormal Then
         .Left = MenuVenster.Left + MenuVenster.Width + 50
         .Top = MenuVenster.Top
      End If
   End With
   
EindeProcedure:
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure verwijdert het opgegeven gerechtbestand.
Public Sub VerwijderGerecht(VerwijderdGerechtbestand As String, VerwijderdGerechtnr As Long)
On Error GoTo Fout
   Screen.MousePointer = vbHourglass
   
   Kill DATA_MAP & VerwijderdGerechtbestand
    
   If VerwijderdGerechtnr = Gerecht.HuidigNr Then
      StelGerechtIn
      GerechtBerekenKosten
   End If
    
   WerkGerechtenlijstBij
   
EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub
   
Fout:
   If HandelFoutAf(, ActiefBestand:="het gerechtbestand") = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure voegt een backslash toe aan het opgegeven pad indien deze niet al aanwezig is en stuurt het resultaat terug.
Public Function VoegBackslashToe(Pad As String) As String
On Error GoTo Fout
Dim AangepastPad As String

   AangepastPad = Trim$(Pad)
   If Not Right$(AangepastPad, 1) = "\" Then AangepastPad = AangepastPad & "\"
   
EindeProcedure:
   VoegBackslashToe = AangepastPad
   Exit Function
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Function

'Deze procedure berekent de kostprijs voor de opgegeven producten.
Private Sub VVBerekenKostprijs(Producten As VVStr)
On Error GoTo Fout

   With Producten
      .Kostprijs = BerekenBedrag(.Prijs(.Selectie), .Hoeveelheid(.Selectie), .Meeteenheid)
      .Kostprijs = TelBedragenOp(.Kostprijs, BerekenBedrag(.Kostprijs, .Snijverlies(.Selectie), MEProcent))
   End With

EindeProcedure:
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure bewaart de opgegeven producten in het opgegeven gerechtbestand.
Private Sub VVBewaarGerecht(Producten As VVStr, BestandH As Integer)
On Error GoTo Fout
Dim Productnr As Long

   With Producten
      Print #BestandH, Chr$(.Selectie);
      For Productnr = LBound(.Product()) To UBound(.Product())
         Print #BestandH, Chr$(Len(.Hoeveelheid(Productnr))); .Hoeveelheid(Productnr);
         Print #BestandH, Chr$(Len(.Snijverlies(Productnr))); .Snijverlies(Productnr);
      Next Productnr
   End With

EindeProcedure:
   Exit Sub
   
Fout:
   If HandelFoutAf(, ActiefBestand:="het gerechtbestand") = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure bewaart de opgegeven producten in het opgegeven productenbestand.
Private Sub VVBewaarProducten(Producten As VVStr, BestandH As Integer)
On Error GoTo Fout
Dim Productnr As Long

   With Producten
      Print #BestandH, Chr$(.Meeteenheid);
      For Productnr = LBound(.Product()) + 1 To UBound(.Product())
         Print #BestandH, Chr$(Len(.Product(Productnr))); .Product(Productnr);
         Print #BestandH, Chr$(Len(.Prijs(Productnr))); .Prijs(Productnr);
      Next Productnr
   End With
      
EindeProcedure:
   Exit Sub
   
Fout:
   If HandelFoutAf(, ActiefBestand:="het productenbestand") = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure drukt de opgegeven producten af.
Private Sub VVDrukAf(Producten As VVStr, Categorie As String)
On Error GoTo Fout

   With Producten
      Printer.Print Tab(3); Categorie; ":"
      Printer.Print String(75, "-")
      Printer.Print Tab(3); "Product:"; Tab(22); Left$(.Product(.Selectie), 20)
      Printer.Print Tab(3); "Prijs:"; Tab(32 - Len(.Prijs(.Selectie))); .Prijs(.Selectie)
      Printer.Print Tab(3); "Hoeveelheid:"; Tab(32 - Len(.Hoeveelheid(.Selectie))); .Hoeveelheid(.Selectie)
      Printer.Print Tab(3); "Snijverlies %:"; Tab(32 - Len(.Snijverlies(.Selectie))); .Snijverlies(.Selectie)
      Printer.Print String(75, "-")
      Printer.Print Tab(3); "Kostprijs:"; Tab(32 - Len(.Kostprijs)); .Kostprijs
   End With

EindeProcedure:
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procudure laadt producten uit het opgegeven bestand.
Private Sub VVLaadGerecht(Producten As VVStr, BestandH As Integer)
On Error GoTo Fout
Dim Lengte As Long
Dim Productnr As Long

   With Producten
      .Selectie = Asc(Input$(1, BestandH))
      For Productnr = LBound(Vlees.Product()) To UBound(.Product())
         Lengte = Asc(Input$(1, BestandH)): .Hoeveelheid(Productnr) = Val(Input$(Lengte, BestandH))
         Lengte = Asc(Input$(1, BestandH)): .Snijverlies(Productnr) = Val(Input$(Lengte, BestandH))
         .Kostprijs = BerekenBedrag(.Prijs(.Selectie), .Hoeveelheid(.Selectie), .Meeteenheid)
      Next Productnr
   End With
         
EindeProcedure:
   Exit Sub
   
Fout:
   If HandelFoutAf(, ActiefBestand:="het gerechtbestand") = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procudure laadt de producten uit het opgegeven bestand.
Private Sub VVLaadProducten(Producten As VVStr, BestandH As Integer)
On Error GoTo Fout
Dim Lengte As Long
Dim Productnr As Long

   With Producten
      .Meeteenheid = Asc(Input$(1, BestandH))
      For Productnr = LBound(.Product()) + 1 To UBound(.Product())
         Lengte = Asc(Input$(1, BestandH)): .Product(Productnr) = Input$(Lengte, BestandH)
         Lengte = Asc(Input$(1, BestandH)): .Prijs(Productnr) = Input$(Lengte, BestandH)
      Next Productnr
   End With
         
EindeProcedure:
   Exit Sub
   
Fout:
   If HandelFoutAf(, ActiefBestand:="het productenbestand") = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure stelt de opgegeven product categorie in.
Private Sub VVStelIn(Producten As VVStr, Optional AantalProducten As Long = GEEN_AANTAL)
On Error GoTo Fout
Dim Productnr As Long

   With Producten
      If Not AantalProducten = GEEN_AANTAL Then
         ReDim .Product(0 To AantalProducten) As String
         ReDim .Prijs(0 To AantalProducten) As String
         ReDim .Hoeveelheid(0 To AantalProducten) As String
         ReDim .Snijverlies(0 To AantalProducten) As String
      End If
   
      .Kostprijs = "0.00"
      .Meeteenheid = MEKilos
      .Product(LBound(.Product())) = GEEN_PRODUCT_TEKST
      For Productnr = LBound(.Product()) To UBound(.Product())
         If Not AantalProducten = GEEN_AANTAL Then
            If Productnr > LBound(.Product()) Then .Product(Productnr) = "product " & Productnr
            .Prijs(Productnr) = "0.00"
         End If
         
         .Hoeveelheid(Productnr) = "0"
         .Snijverlies(Productnr) = "0"
      Next Productnr
   
      .Selectie = 0
   End With
   
EindeProcedure:
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure stelt het opgegeven venster in met de opgegeven producten.
Public Sub VVStelVensterIn(Producten As VVStr, Venster As Form, VensterNaam As String)
On Error GoTo Fout
Dim Productnr As Long

   With Producten
      Screen.MousePointer = vbHourglass
      
      If VensterActief(VensterNaam) Then
         VensterNaastMenu Venster, VensterNaam
         
         Venster.ProductenLijst.Clear
      
         For Productnr = LBound(.Product()) To UBound(.Product())
            If VensterActief(VensterNaam) Then Venster.ProductenLijst.AddItem .Product(Productnr)
         Next Productnr
      
         Venster.ProductenLijst.ListIndex = .Selectie
         Venster.ProductenLijst.Text = .Product(.Selectie)
         Venster.PrijsVeld.Text = RondBedragAf(.Prijs(.Selectie))
         Venster.HoeveelheidVeld.Text = .Hoeveelheid(.Selectie)
         Venster.SnijverliesVeld.Text = .Snijverlies(.Selectie)
         
         VVBerekenKostprijs Producten
         
         Venster.KostPrijsVeld.Caption = .Kostprijs
      End If
      
      Screen.MousePointer = vbDefault
   End With

EindeProcedure:
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure verwerkt de productinformatieinvoer.
Public Sub VVVerwerkInvoer(Producten As VVStr, Venster As Form)
On Error GoTo Fout

   With Producten
      Venster.PrijsVeld.Text = RondBedragAf(Venster.PrijsVeld.Text)
      Venster.HoeveelheidVeld.Text = CStr(Val(Venster.HoeveelheidVeld.Text))
      Venster.SnijverliesVeld.Text = CStr(Val(Venster.SnijverliesVeld.Text))

      .Prijs(.Selectie) = Venster.PrijsVeld.Text
      .Hoeveelheid(.Selectie) = Venster.HoeveelheidVeld.Text
      .Snijverlies(.Selectie) = Venster.SnijverliesVeld.Text
            
      VVBerekenKostprijs Producten
              
      Venster.KostPrijsVeld.Caption = .Kostprijs
   End With

EindeProcedure:
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure werkt de bewaarde gerechtenlijst bij.
Public Sub WerkGerechtenlijstBij()
On Error GoTo Fout
Dim Bestand As String
Dim BestandH As Integer
Dim GerechtDatum As String
Dim GerechtNaam As String
Dim Gerechtnr As Long
Dim Lengte As Long

   Screen.MousePointer = vbHourglass

   If VensterActief("LaadGerechtVenster") Or VensterActief("VerwijderGerechtVenster") Then
      ReDim GerechtBestanden(0 To 0) As GerechtbestandStr
      
      If VensterActief("LaadGerechtVenster") Then LaadGerechtVenster.GerechtLijst.Clear
      If VensterActief("VerwijderGerechtVenster") Then VerwijderGerechtVenster.GerechtLijst.Clear
      
      For Gerechtnr = 0 To MAX_AANTAL_GERECHTEN - 1
         Bestand = Dir$(DATA_MAP & "Gerecht" & Gerechtnr & ".kcc", vbArchive Or vbHidden Or vbSystem)
         If Not Bestand = vbNullString Then
            BestandH = FreeFile()
            Open DATA_MAP & Bestand For Binary Lock Read Write As BestandH
               Lengte = Asc(Input$(1, BestandH)): GerechtNaam = Input$(Lengte, BestandH)
               Lengte = Asc(Input$(1, BestandH)): GerechtDatum = Input$(Lengte, BestandH)
            Close BestandH
            
            GerechtBestanden(UBound(GerechtBestanden())).Naam = "Gerecht" & Gerechtnr & ".kcc"
            GerechtBestanden(UBound(GerechtBestanden())).Datum = GerechtDatum
            
            If VensterActief("LaadGerechtVenster") Then LaadGerechtVenster.GerechtLijst.AddItem GerechtNaam
            If VensterActief("VerwijderGerechtVenster") Then VerwijderGerechtVenster.GerechtLijst.AddItem GerechtNaam
            
            ReDim Preserve GerechtBestanden(LBound(GerechtBestanden()) To UBound(GerechtBestanden()) + 1) As GerechtbestandStr
         End If
      Next Gerechtnr
      
   End If
EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub
   
Fout:
   If HandelFoutAf(, ActiefBestand:="het gerechtbestand") = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

'Deze procedure werkt het meeteenhedenmenu bij.
Public Sub WerkMeeteenhedenmenuBij(GeselecteerdeMeeteenheid As Long, Optional DecilitersZichtbaar As Boolean = False)
On Error GoTo Fout
Dim Meeteenheidnr As Long
   
   With KCalcVenster
      For Meeteenheidnr = .MeeteenheidMenu.LBound To .MeeteenheidMenu.UBound
         .MeeteenheidMenu(Meeteenheidnr).Checked = (Meeteenheidnr = GeselecteerdeMeeteenheid)
      Next Meeteenheidnr
     
      .MeeteenheidMenu(MEDeciliters).Visible = DecilitersZichtbaar
   End With
   
EindeProcedure:
   Exit Sub
   
Fout:
   If HandelFoutAf() = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf(VraagVorigeKeuzeOp:=True) = vbRetry Then Resume
End Sub

