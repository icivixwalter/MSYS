ROUTINE_DA_SALVARE_IN_DB_FINALE
   DA SALVARE IN c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSYS_ATTIVA_GEST_OGGETTI\MDB\MSYS_ATTIVA_GEST_OGGETTI.mdb
   MODULO UTILITA_MsysDF14Mdl05_IMPORTAZIONE_SEMPLICE_DI_OGGETTI

    '//***************************************************************************************************************//
'//*'
'//*'   MODULO      :   UTILITA_MsysDF14Mdl05_IMPORTAZIONE_SEMPLICE_DI_OGGETTI
'//*'   DEFINIZIONE :   MODULO PER L'IMPORTAZIONE SEMPLICE DEGLI OGGETTI.
'//*'   Note        :   Modulo di collegamento o link semplice.
'//*'               :   Esegue il link delle tabelle, importa oggetti query, form report e macro senza effettuare
'//*'                   controlli sul DATABASE ESTERNO. Non controllo infatti la corrispodenza degli oggetti
'//*'                   tra il DATABASE ESTERNO e il DATABASE INTERNO.
'//*'                   EFFETTUA SOLO IL CONTROLLO DI CORRISPONDENZA  tra la collection oggetti e l'oggetto
'//*'                   da importare se esiste nel DATABASE CORRENTE
'//*'                   Per ora EFFETTUA IL LINK DELLE SOLE TABELLE DLL
'//*'
'//*'
'//*'   codice      : @importazione@semplice_(importa tabella o collega tabelle SENZA CONTROLLO ESTERNO)
'//*'                 @link@tabelle@dll_(importazione semplice senza controllo sul db esterno delle tabelle DLL)
'//*'
'//*'
'//*'   parametro   : il parametro interno è la path ed il database sui cui prelevare gli oggetti, eventualmente
'//*'               è possibile cambiare la path e db esterno:
'//*'                   PathDb_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\MSYS\MDB\MSYS_N01_OGGETTI.mdb"
'//*'
'//*'  attivazione  : si attiva mediante una macro che esegue il codice che importa tutto o puo attivare singolarmente.
'//*'
'//*'                   ImportaTuttiGliOggetti          = importo tutto
'//*'                   ImportaModuli()                 = importo solo i moduli
'//*'                   ImportaTabelel()                = importo solo le TABELLE ecc..
'//*'
'//*'
'//*'   ESTENSIONE : l'Importazione semplice puo essere estesa anche alle query, report, form macro e moduli sempre
'//*'               con lo stesso criterio: solo CONTROLLO INTERNO e NESSUN CONTROLLO ESTERNO
'//*'



'//* Nome funzione/routine          | Descrizione sintetica
'//* -------------------------------|-------------------------------------------------------------
'//* InizializzaPath                | Inizializza la path del database
'//* InizializzaArrayTABELLE        | Inizializza gli array DELLE TABELLE con i nomi degli oggetti
'//* InizializzaArrayQUERY          | Inizializza gli array DELLE QUERY con i nomi degli oggetti
'//* InizializzaArrayMACRO          | Inizializza gli array DELLE MACRO con i nomi degli oggetti
'//* InizializzaArrayMODULI         | Inizializza gli array DEI MODULI con i nomi degli oggetti
'//* ATTIVA_ImportaMacro            | Attiva la funzione per l’importazione delle macro
'//* ImportaMacro                   | Attiva la funzione per l’importazione delle macro
'//* ATTIVA_CancellaMacro           | Cancella le macro importate (errore gestito)
'//* CancellaMacro                  | Cancella le macro importate (errore gestito)
'//* ATTIVA_ImportaQuery            | Attiva la funzione per l’importazione delle query
'//* ImportaQuery                   | Attiva la funzione per l’importazione delle query
'//* ATTIVA_CancellaQuery           | Cancella le query importate (errore gestito)
'//* CancellaQuery                  | Cancella le query importate (errore gestito)
'//* ATTIVA_ImportaTabelle          | Attiva la funzione per l’importazione delle tabelle
'//* ImportaTabelle                 | Attiva la funzione per l’importazione delle tabelle
'//* ATTIVA_CancellaTabelle         | Attiva la funzione per la cancellazione delle tabelle
'//* CancellaTabelle                | Cancella le tabelle importate
'//* ATTIVA_ImportaModuli           | Attiva la funzione per l’importazione dei moduli
'//* ATTIVA_CancellaModuli          | Attiva la funzione per la cancellazione dei moduli
'//* ImportaModuli                  | Importa i moduli (descrizione non esplicita nel codice)
'//* CancellaModuli                 | Cancella i moduli importati
'//* ImportaTuttiGliOggetti         | Importa tutti gli oggetti (moduli, query, tabelle, ecc.)
'//* CancellaTuttiGliOggetti        | Cancella tutti gli oggetti importati


'//***************************************************************************************************************//


Option Compare Database

' // DIM variabili procedura importa e collega DLL
Dim PathDb_s As String

' // Array globali per i nomi degli oggetti
Dim queryArray(7) As String
Dim TableArray(3) As String
Dim ModuleArray(1) As String
Dim MacroArray(1) As String

' // >>> MODIFICA: dichiarazione globale Collection per contenere i riferimenti alle tabelle
Dim TabellaCollection As Collection



'// DIM COLLECTION TABELLE
Private Type TabellaInfo
    Nome As String
    PercorsoCompleto As String
End Type




'//2024.08.05_IMPORTAZIONE_SEMPLICE _
                modifico il modulo UTILITA_Msys_Mdl05_IMPORTAZIONE_SEMPLICE_DI_OGGETTI. Qui effetto l'importazione degli oggetti _
                in modo semplice cioè senza controllo del DATABASE ESTERNO. _
                Il controllo è solo SUGLI OGGETTI INTERNI.





' // Inizializza la path del database
Private Sub InizializzaPath()
    PathDb_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\MSYS\MDB\MSYS_N01_OGGETTI.mdb"
End Sub



' // Inizializza gli array DELLE QUERY con i nomi degli oggetti
Private Sub InizializzaArrayQUERY()
    On Error GoTo InizializzaArrayQUERY_Err

    queryArray(0) = "Msys_DLL_Qry01_81_}--------------------------------------------@"
    queryArray(1) = "Msys_DF14Qry01_01_DELETE_LIBRERIE"
    queryArray(2) = "Msys_DF13Qry01_00_}-----------------------------------------@DLL"
    queryArray(3) = "Msys_DF13Qry01_01_DLL_PROGETTO"
    queryArray(4) = "Msys_DF14Qry01_02_DLL_LIBRERIE"
    queryArray(5) = "Msys_DLLQry01_00_}---------------------------------------------@"
    queryArray(6) = "Msys_DLLQry01_01_LIBRERIE_ATTIVE"
    queryArray(7) = "Msys_DLLQry01_10_GROUP_DLL_Correnti"

    
    Exit Sub

InizializzaArrayQUERY_Err:
    MsgBox "Errore durante l'inizializzazione degli array: " & Err.Description
End Sub



' // Inizializza gli array DELLE MACRO con i nomi degli oggetti
Private Sub InizializzaArrayMACRO()
    On Error GoTo InizializzaArrayMACRO_Err

    MacroArray(0) = "UTILITA_MSys_Mcr00_}----------------------------------------@DLL"
    MacroArray(1) = "UTILITA_Msys_Mcr01_STAMPA_LIBRERIE_ATTIVE_DLL"

    Exit Sub

InizializzaArrayMACRO_Err:
    MsgBox "Errore durante l'inizializzazione degli array: " & Err.Description
End Sub



' // Inizializza gli array DEI MODULI con i nomi degli oggetti
Private Sub InizializzaArrayMODULI()
    On Error GoTo InizializzaArrayMODULI_Err
  
    ModuleArray(0) = "UTILITA_Msys_Mdl00_}----------------------------------------@DLL"
    ModuleArray(1) = "UTILITA_Msys_Mdl03_DLL_REFERENZIA_LE_LIBRERIE_DLL"

    Exit Sub

InizializzaArrayMODULI_Err:
    MsgBox "Errore durante l'inizializzazione degli array MODULI: " & Err.Description
End Sub



'//==================================================================================================//
'//         OGGETTO QUERY           *** INIZIO ***
'//==================================================================================================//







'//==================================================================================================//
'//         OGGETTO MACRO           *** INIZIO ***
'//==================================================================================================//

'//ATTIVO LA FUNZIONE
Private Sub ATTIVA_ImportaMacro()
    Bool1 = ImportaMacro
End Sub

' // Funzione per importare le Macro
Public Function ImportaMacro() As Boolean
    On Error GoTo ImportaMacro_Err

    ' Inizializza gli array delle Macro e il percorso del database
    InizializzaArrayMACRO
    InizializzaPath

    ' Importa le Macro
    For i = LBound(MacroArray) To UBound(MacroArray)
        ' Verifica se la Macro esiste facendo una Macro nella tabella di sistema ed estraendo il nome ricercato _
        se esiste conteggia il rs > 0 allo esiste e quindi CANCELLA PRIMA LA Macro. Se il rs = 0 non esiste la Macro e quindi non _
        cancella ma importa direttamente
        If DCount("*", "MSysObjects", "Name='" & MacroArray(i) & "' AND Type=5") > 0 Then
            DoCmd.DeleteObject acMacro, MacroArray(i)
        End If
        DoCmd.TransferDatabase acImport, "Microsoft Access", PathDb_s, acMacro, MacroArray(i), MacroArray(i), False
    Next i

    ImportaMacro = True
    Exit Function

ImportaMacro_Err:
    MsgBox "Errore durante l'importazione delle Macro: " & Err.Description
    ImportaMacro = False
End Function



Private Sub ATTIVA_CancellaMacro()
    Bool1 = CancellaMacro
End Sub

' // Funzione per cancellare tutte le Macro
Public Function CancellaMacro() As Boolean
    On Error GoTo CancellaMacro_Err

    'InizializzaArray           ' BLOCCATO PERCHE CONTENTEVA TUTTI GLI OGGETTI
    
    InizializzaArrayMACRO
    
    ' Cancellazione delle Macro
    ' Verifica se la Macro esiste facendo una Macro nella tabella di sistema ed estraendo il nome ricercato _
    se esiste conteggia il rs > 0 allo esiste e quindi CANCELLA PRIMA LA Macro. Se il rs = 0 non esiste la Macro e quindi non _
    cancella ma importa direttamente
    
    '//TODO: qui non cancella le macro importate forse il type è errato
    For i = LBound(MacroArray) To UBound(MacroArray)
        If DCount("[Name]", "MSysObjects", "[Type] = 5 AND [Name] = '" & MacroArray(i) & "'") > 0 Then
            DoCmd.DeleteObject acMacro, MacroArray(i)
        End If
    Next i

    CancellaMacro = True
    Exit Function

CancellaMacro_Err:
    MsgBox "Errore durante la cancellazione delle Macro: " & Err.Description
    CancellaMacro = False
End Function

'//==================================================================================================//
'//         OGGETTO MACRO           *** FINE ***
'//==================================================================================================//

'//ATTIVO LA FUNZIONE
Private Sub ATTIVA_ImportaQuery()
    Bool1 = ImportaQuery
End Sub

' // Funzione per importare le query
Public Function ImportaQuery() As Boolean
    On Error GoTo ImportaQuery_Err

    ' Inizializza gli array delle query e il percorso del database
    InizializzaArrayQUERY
    InizializzaPath

    ' Importa le query
    For i = LBound(queryArray) To UBound(queryArray)
        ' Verifica se la query esiste facendo una query nella tabella di sistema ed estraendo il nome ricercato _
        se esiste conteggia il rs > 0 allo esiste e quindi CANCELLA PRIMA LA QUERY. Se il rs = 0 non esiste la query e quindi non _
        cancella ma importa direttamente
        If DCount("*", "MSysObjects", "Name='" & queryArray(i) & "' AND Type=5") > 0 Then
            DoCmd.DeleteObject acQuery, queryArray(i)
        End If
        DoCmd.TransferDatabase acImport, "Microsoft Access", PathDb_s, acQuery, queryArray(i), queryArray(i), False
    Next i

    ImportaQuery = True
    Exit Function

ImportaQuery_Err:
    MsgBox "Errore durante l'importazione delle query: " & Err.Description
    ImportaQuery = False
End Function



Private Sub ATTIVA_CancellaQuery()
    Bool1 = CancellaQuery
End Sub

' // Funzione per cancellare tutte le query
Public Function CancellaQuery() As Boolean
    On Error GoTo CancellaQuery_Err

    'InizializzaArray           ' BLOCCATO PERCHE CONTENTEVA TUTTI GLI OGGETTI
    
    InizializzaArrayQUERY
    
    ' Cancellazione delle query
    ' Verifica se la query esiste facendo una query nella tabella di sistema ed estraendo il nome ricercato _
    se esiste conteggia il rs > 0 allo esiste e quindi CANCELLA PRIMA LA QUERY. Se il rs = 0 non esiste la query e quindi non _
    cancella ma importa direttamente
 
    For i = LBound(queryArray) To UBound(queryArray)
        If DCount("[Name]", "MSysObjects", "[Type] = 5 AND [Name] = '" & queryArray(i) & "'") > 0 Then
            DoCmd.DeleteObject acQuery, queryArray(i)
        End If
    Next i

    CancellaQuery = True
    Exit Function

CancellaQuery_Err:
    MsgBox "Errore durante la cancellazione delle query: " & Err.Description
    CancellaQuery = False
End Function

'//==================================================================================================//
'//         OGGETTO QUERY           *** FINE ***
'//==================================================================================================//






'//==================================================================================================//
'//         OGGETTO TABLE           *** INIZIO ***
'//==================================================================================================//

'// ATTIVA LA FUNZIONE DI COLLEGAMENTO
Public Sub ATTIVA_ImportaTabelle()
    Bool1 = ImportaTabelle
End Sub

'// Inizializza la Collection DELLE TABELLE con nome e percorso completo
'// TODO: costruisco la procedura di collegamento delle tabelle per tutto il db
'// codice----> MSYS_ATTIVA_GEST_OGGETTI_Link_Tables=(collego tutte le tabelle necessarie per la gestione del db)
Private Sub InizializzaCollectionTABELLE()
        
    On Error GoTo InizializzaArrayTABELLE_Err

    ' >>> MODIFICA: inizializza la collection
    Set TabellaCollection = New Collection

    Dim PercorsoBase_s As String
    Dim tbl As Object  ' >>> MODIFICA: uso un Dictionary per salvare Nome e PercorsoCompleto

    '01)_DF01
    ' -------------------------------------------------------------------------------------
    PercorsoBase_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_DF01\MDB\MSys_DF01.mdb"

        ' Tabella 0
        Set tbl = CreateObject("Scripting.Dictionary")
        tbl.Add "Nome", "MSys_DF01_}----------------------------------------------------@"
        tbl.Add "PercorsoCompleto", PercorsoBase_s
        TabellaCollection.Add tbl
    
        ' Tabella 1
        Set tbl = CreateObject("Scripting.Dictionary")
        tbl.Add "Nome", "MSys_DF01_COMANDI"
        tbl.Add "PercorsoCompleto", PercorsoBase_s
        TabellaCollection.Add tbl

    '02)_DF02
    ' -------------------------------------------------------------------------------------
    PercorsoBase_s = "c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_DF02\MDB\MSys_DF02.mdb"
    
        ' Tabella 2
        Set tbl = CreateObject("Scripting.Dictionary")
        tbl.Add "Nome", "MSys_DF02_}----------------------------------------------------@"
        tbl.Add "PercorsoCompleto", PercorsoBase_s
        TabellaCollection.Add tbl
    
        ' Tabella 3
        Set tbl = CreateObject("Scripting.Dictionary")
        tbl.Add "Nome", "MSys_DF02_TIPO_OGGETTO"
        tbl.Add "PercorsoCompleto", PercorsoBase_s
        TabellaCollection.Add tbl

    '04)_DF05
    ' -------------------------------------------------------------------------------------
        PercorsoBase_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_DF05\MDB\MSys_DF05.mdb"
    
        ' Tabella 4
        Set tbl = CreateObject("Scripting.Dictionary")
        tbl.Add "Nome", "MSys_DF05_}----------------------------------------------------@"
        tbl.Add "PercorsoCompleto", PercorsoBase_s
        TabellaCollection.Add tbl
    
    
       ' Tabella 5
        Set tbl = CreateObject("Scripting.Dictionary")
        tbl.Add "Nome", "MSys_DF05_ATTRIBUTI_TABELLA"
        tbl.Add "PercorsoCompleto", PercorsoBase_s
        TabellaCollection.Add tbl
    
    '05)_DF11
    ' -------------------------------------------------------------------------------------
        PercorsoBase_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_DF11\MDB\MSys_DF11.mdb"
    
        ' Tabella 6
        Set tbl = CreateObject("Scripting.Dictionary")
        tbl.Add "Nome", "MSys_DF11_{@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@}_Tab"
        tbl.Add "PercorsoCompleto", PercorsoBase_s
        TabellaCollection.Add tbl
    
    
       ' Tabella 7
        Set tbl = CreateObject("Scripting.Dictionary")
        tbl.Add "Nome", "MSys_DF11_}----------------------------------------------------@"
        tbl.Add "PercorsoCompleto", PercorsoBase_s
        TabellaCollection.Add tbl
    
    
    
    
       ' Tabella 8
        Set tbl = CreateObject("Scripting.Dictionary")
        tbl.Add "Nome", "Msys_DF11_PROGETTI"
        tbl.Add "PercorsoCompleto", PercorsoBase_s
        TabellaCollection.Add tbl
    
    '06)_DF12
    ' -------------------------------------------------------------------------------------
       PercorsoBase_s = "c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_DF12\MDB\MSys_DF12.mdb"

       ' Tabella 9
       Set tbl = CreateObject("Scripting.Dictionary")
       tbl.Add "Nome", "Msys_DF12_}----------------------------------------------------@"
       tbl.Add "PercorsoCompleto", PercorsoBase_s
       TabellaCollection.Add tbl


      ' Tabella 10
       Set tbl = CreateObject("Scripting.Dictionary")
       tbl.Add "Nome", "Msys_DF12_GE_ObjProgetti"
       tbl.Add "PercorsoCompleto", PercorsoBase_s
       TabellaCollection.Add tbl

    '07)_DF13
    ' -------------------------------------------------------------------------------------
       PercorsoBase_s = "c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_DF13\MDB\MSys_DF13.mdb"

       ' Tabella 11
       Set tbl = CreateObject("Scripting.Dictionary")
       tbl.Add "Nome", "Msys_DF13_}----------------------------------------------------@"
       tbl.Add "PercorsoCompleto", PercorsoBase_s
       TabellaCollection.Add tbl


      ' Tabella 12
       Set tbl = CreateObject("Scripting.Dictionary")
       tbl.Add "Nome", "Msys_DF13_DLL_PROGETTO"
       tbl.Add "PercorsoCompleto", PercorsoBase_s
       TabellaCollection.Add tbl


    '08)_DF14
    ' -------------------------------------------------------------------------------------
       PercorsoBase_s = "c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_DF14\MDB\MSys_DF14.mdb"

       ' Tabella 13
       Set tbl = CreateObject("Scripting.Dictionary")
       tbl.Add "Nome", "Msys_DF14_}----------------------------------------------------@"
       tbl.Add "PercorsoCompleto", PercorsoBase_s
       TabellaCollection.Add tbl


      ' Tabella 14
       Set tbl = CreateObject("Scripting.Dictionary")
       tbl.Add "Nome", "Msys_DF14_DLL_LIBRERIE"
       tbl.Add "PercorsoCompleto", PercorsoBase_s
       TabellaCollection.Add tbl








  
    
    





    Exit Sub

InizializzaArrayTABELLE_Err:
    MsgBox "Errore in InizializzaCollectionTABELLE: " & Err.Description, vbExclamation
End Sub

'// Funzione per collegare (linkare) le tabelle
Public Function ImportaTabelle() As Boolean
    On Error GoTo ImportaTabelle_Err

    Dim ImportedTables() As String
    Dim TableCount As Integer
    Dim i As Long
    Dim tbl As Object

    TableCount = 0

    InizializzaCollectionTABELLE

    ReDim ImportedTables(0 To TabellaCollection.Count - 1)

    ' Loop di collegamento delle tabelle
    For i = 1 To TabellaCollection.Count
        Set tbl = TabellaCollection(i)

        ' Se esiste una tabella locale con lo stesso nome, la elimino prima
        If DCount("*", "MSysObjects", "Name='" & Replace(tbl("Nome"), "'", "''") & "' AND Type=1") > 0 Then
            DoCmd.DeleteObject acTable, tbl("Nome")
        End If

        ' DEBUG: Visualizzo nome tabella e percorso
        Debug.Print "Tabella da collegare: " & tbl("Nome")
        Debug.Print "Database sorgente: " & tbl("PercorsoCompleto")

        ' >>> MODIFICA: collego la tabella invece di importarla ' <<< CAMBIATO da acImport a acLink
        DoCmd.TransferDatabase _
            TransferType:=acLink, _
            DatabaseType:="Microsoft Access", _
            DatabaseName:=tbl("PercorsoCompleto"), _
            objectType:=acTable, _
            Source:=tbl("Nome"), _
            Destination:=tbl("Nome"), _
            StructureOnly:=False

        ' Memorizzo nome tabella collegata
        ImportedTables(TableCount) = tbl("Nome")
        Debug.Print tbl("nome")
        TableCount = TableCount + 1
    Next i

    ' Riepilogo
    Debug.Print "Statistiche riepilogative delle tabelle collegate:"
    Debug.Print "Numero totale di tabelle collegate: " & TableCount
    For i = 0 To TableCount - 1
        Debug.Print "Tabella " & (i + 1) & ": " & ImportedTables(i)
    Next i

    ImportaTabelle = True
    Exit Function

ImportaTabelle_Err:
    MsgBox "Errore durante il collegamento delle tabelle: " & Err.Description
    ImportaTabelle = False
End Function

'// ATTIVA LA FUNZIONE DI CANCELLAZIONE
Private Sub ATTIVA_CancellaTabelle()
    Bool1 = CancellaTabelle
End Sub

' // Funzione per cancellare tutte le tabelle
Public Function CancellaTabelle() As Boolean
    On Error GoTo CancellaTabelle_Err

    InizializzaCollectionTABELLE

    Dim tabelleCancellate As Integer
    Dim tabelleNonTrovate As Integer
    Dim nomiTabelleCancellate As String
    Dim nomiTabelleNonTrovate As String
    Dim i As Long
    Dim tbl As Object

    tabelleCancellate = 0
    tabelleNonTrovate = 0
    nomiTabelleCancellate = ""
    nomiTabelleNonTrovate = ""

    For i = 1 To TabellaCollection.Count
        Set tbl = TabellaCollection(i)

        ' Verifica se la tabella esiste (locale o collegata)
        If DCount("*", "MSysObjects", "Name='" & Replace(tbl("Nome"), "'", "''") & "' AND Type IN (1, 4, 6)") > 0 Then
            DoCmd.DeleteObject acTable, tbl("Nome")
            Debug.Print "Tabella cancellata: " & tbl("Nome")
            tabelleCancellate = tabelleCancellate + 1
            nomiTabelleCancellate = nomiTabelleCancellate & tbl("Nome") & ", "
        Else
            Debug.Print "Tabella non trovata: " & tbl("Nome")
            tabelleNonTrovate = tabelleNonTrovate + 1
            nomiTabelleNonTrovate = nomiTabelleNonTrovate & tbl("Nome") & ", "
        End If
    Next i

    ' Rimuovo ultima virgola
    If tabelleCancellate > 0 Then
        nomiTabelleCancellate = Left(nomiTabelleCancellate, Len(nomiTabelleCancellate) - 2)
    End If
    If tabelleNonTrovate > 0 Then
        nomiTabelleNonTrovate = Left(nomiTabelleNonTrovate, Len(nomiTabelleNonTrovate) - 2)
    End If

    ' Output finale
    Debug.Print "Totale tabelle cancellate: " & tabelleCancellate
    Debug.Print "Nomi delle tabelle cancellate: " & nomiTabelleCancellate
    Debug.Print "Totale tabelle non trovate: " & tabelleNonTrovate
    Debug.Print "Nomi delle tabelle non trovate: " & nomiTabelleNonTrovate

    CancellaTabelle = True
    Exit Function

CancellaTabelle_Err:
    MsgBox "Errore durante la cancellazione delle tabelle: " & Err.Description
    CancellaTabelle = False
End Function







'//==================================================================================================//
'//         OGGETTO TABLE           *** FINE ***
'//==================================================================================================//




'//==================================================================================================//
'//         OGGETTO MODULI           *** INIZIO ***
'//==================================================================================================//



'//ATTIVO LA FUNZIONE
Private Sub ATTIVA_ImportaModuli()
    Bool1 = ImportaModuli
End Sub


'//ATTIVO LA FUNZIONE
Private Sub ATTIVA_CancellaModuli()
    Bool1 = CancellaModuli
End Sub





' // Funzione per importare i moduli
Public Function ImportaModuli() As Boolean
    On Error GoTo ImportaModuli_Err

    InizializzaArrayMODULI           ' RECUPERO L'ARRAY DEI MODULI
    InizializzaPath
    
    Dim moduliImportati As Integer
    Dim moduliGiaEsistenti As Integer
    Dim nomiModuliImportati As String
    Dim nomiModuliGiaEsistenti As String
    
    moduliImportati = 0
    moduliGiaEsistenti = 0
    nomiModuliImportati = ""
    nomiModuliGiaEsistenti = ""
    
    ' Importa i moduli
    For i = LBound(ModuleArray) To UBound(ModuleArray)
        ' Verifica se il modulo esiste nel database corrente
        If DCount("*", "MSysObjects", "Name='" & ModuleArray(i) & "' AND Type=-32761") > 0 Then
            Debug.Print "Modulo già esistente: " & ModuleArray(i)
            moduliGiaEsistenti = moduliGiaEsistenti + 1
            nomiModuliGiaEsistenti = nomiModuliGiaEsistenti & ModuleArray(i) & ", "
        Else
            DoCmd.TransferDatabase acImport, "Microsoft Access", PathDb_s, acModule, ModuleArray(i), ModuleArray(i), False
            Debug.Print "Modulo importato: " & ModuleArray(i)
            moduliImportati = moduliImportati + 1
            nomiModuliImportati = nomiModuliImportati & ModuleArray(i) & ", "
        End If
    Next i
    
    ' Rimuove l'ultima virgola e spazio
    If moduliImportati > 0 Then
        nomiModuliImportati = Left(nomiModuliImportati, Len(nomiModuliImportati) - 2)
    End If
    
    If moduliGiaEsistenti > 0 Then
        nomiModuliGiaEsistenti = Left(nomiModuliGiaEsistenti, Len(nomiModuliGiaEsistenti) - 2)
    End If
    
    ' Stampa della statistica riepilogativa
    Debug.Print "Totale moduli importati: " & moduliImportati
    Debug.Print "Nomi dei moduli importati: " & nomiModuliImportati
    Debug.Print "Totale moduli già esistenti: " & moduliGiaEsistenti
    Debug.Print "Nomi dei moduli già esistenti: " & nomiModuliGiaEsistenti

    ImportaModuli = True
    Exit Function

ImportaModuli_Err:
    MsgBox "Errore durante l'importazione dei moduli: " & Err.Description
    ImportaModuli = False
    
End Function


' // Funzione per cancellare tutti i moduli
Public Function CancellaModuli() As Boolean
    On Error GoTo CancellaModuli_Err

    InizializzaArrayMODULI           ' RECUPERO L'ARRAY DEI MODULI
    
    Dim moduliCancellati As Integer
    Dim moduliNonTrovati As Integer
    Dim nomiModuliCancellati As String
    Dim nomiModuliNonTrovati As String
    
    moduliCancellati = 0
    moduliNonTrovati = 0
    nomiModuliCancellati = ""
    nomiModuliNonTrovati = ""
    
    ' Cancellazione dei moduli
    For i = LBound(ModuleArray) To UBound(ModuleArray)
        ' Verifica se il modulo esiste
        If DCount("*", "MSysObjects", "Name='" & ModuleArray(i) & "' AND Type=-32761") > 0 Then
            DoCmd.DeleteObject acModule, ModuleArray(i)
            Debug.Print "Modulo cancellato: " & ModuleArray(i)
            moduliCancellati = moduliCancellati + 1
            nomiModuliCancellati = nomiModuliCancellati & ModuleArray(i) & ", "
        Else
            Debug.Print "Modulo non trovato: " & ModuleArray(i)
            moduliNonTrovati = moduliNonTrovati + 1
            nomiModuliNonTrovati = nomiModuliNonTrovati & ModuleArray(i) & ", "
        End If
    Next i
    
    ' Rimuove l'ultima virgola e spazio
    If moduliCancellati > 0 Then
        nomiModuliCancellati = Left(nomiModuliCancellati, Len(nomiModuliCancellati) - 2)
    End If
    
    If moduliNonTrovati > 0 Then
        nomiModuliNonTrovati = Left(nomiModuliNonTrovati, Len(nomiModuliNonTrovati) - 2)
    End If
    
    ' Stampa della statistica riepilogativa
    Debug.Print "Totale moduli cancellati: " & moduliCancellati
    Debug.Print "Nomi dei moduli cancellati: " & nomiModuliCancellati
    Debug.Print "Totale moduli non trovati: " & moduliNonTrovati
    Debug.Print "Nomi dei moduli non trovati: " & nomiModuliNonTrovati

    CancellaModuli = True
    Exit Function

CancellaModuli_Err:
    MsgBox "Errore durante la cancellazione dei moduli: " & Err.Description
    CancellaModuli = False
    
End Function



'//==================================================================================================//
'//         OGGETTO MODULI           *** FINE ***
'//==================================================================================================//



' // Funzione generale per importare tutti gli oggetti
Public Function ImportaTuttiGliOggetti() As Boolean
    On Error GoTo ImportaTuttiGliOggetti_Err

    ImportaQuery
    ImportaTabelle
    ImportaModuli
    ImportaMacro

    ImportaTuttiGliOggetti = True
    Exit Function

ImportaTuttiGliOggetti_Err:
    MsgBox "Errore durante l'importazione di tutti gli oggetti: " & Err.Description
    ImportaTuttiGliOggetti = False
End Function

' // Funzione generale per cancellare tutti gli oggetti
Public Function CancellaTuttiGliOggetti() As Boolean
    On Error GoTo CancellaTuttiGliOggetti_Err

    CancellaQuery
    CancellaTabelle
    CancellaModuli
    CancellaMacro
    MsgBox "Tutti gli oggetti importati o collegati sono stati cancellati."

    CancellaTuttiGliOggetti = True
    Exit Function

CancellaTuttiGliOggetti_Err:
    MsgBox "Errore durante la cancellazione di tutti gli oggetti: " & Err.Description
    CancellaTuttiGliOggetti = False
End Function


'//********************************************************************************//
'//     ESECUZIONI DI PROVA









