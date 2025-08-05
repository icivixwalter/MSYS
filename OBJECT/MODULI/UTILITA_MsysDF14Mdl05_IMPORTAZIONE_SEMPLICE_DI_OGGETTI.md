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
'//*'                   ImportaModuli_PFunct()                 = importo solo i moduli
'//*'                   ImportaTabelel()                = importo solo le TABELLE ecc..
'//*'
'//*'
'//*'   ESTENSIONE : l'Importazione semplice puo essere estesa anche alle query, report, form macro e moduli sempre
'//*'               con lo stesso criterio: solo CONTROLLO INTERNO e NESSUN CONTROLLO ESTERNO
'//*'
'//*'   codice ---> 2025.08.05_MSys_MODULI.mdb_(modifiche e miglioramenti EFFETTUATI NEL MODULO)


'//* Nome funzione/routine          | Descrizione sintetica
'//* -------------------------------|-------------------------------------------------------------
'//* InizializzaPath                | Inizializza la path del database
'//* InizializzaArrayTABELLE        | Inizializza gli array DELLE TABELLE con i nomi degli oggetti
'//* InizializzaArrayQUERY          | Inizializza gli array DELLE QUERY con i nomi degli oggetti
'//* InizializzaArrayMACRO          | Inizializza gli array DELLE MACRO con i nomi degli oggetti
'//* InizializzaArrayMODULI         | Inizializza gli array DEI MODULI con i nomi degli oggetti
'//* ATTIVA_ImportaMacro_PFunct            | Attiva la funzione per l’importazione delle macro
'//* ImportaMacro_PFunct                   | Attiva la funzione per l’importazione delle macro
'//* ATTIVA_CancellaMacro_Pfunct           | Cancella le macro importate (errore gestito)
'//* CancellaMacro_Pfunct                  | Cancella le macro importate (errore gestito)
'//* ATTIVA_ImportaQuery_PFunct            | Attiva la funzione per l’importazione delle query
'//* ImportaQuery_PFunct                   | Attiva la funzione per l’importazione delle query
'//* ATTIVA_CancellaQuery_Funct           | Cancella le query importate (errore gestito)
'//* CancellaQuery_Funct                  | Cancella le query importate (errore gestito)
'//* ATTIVA_ImportaTabelle_PFunct          | Attiva la funzione per l’importazione delle tabelle
'//* ImportaTabelle_Pfunct                 | Attiva la funzione per l’importazione delle tabelle
'//* ATTIVA_CancellaTabelle_PFunct         | Attiva la funzione per la cancellazione delle tabelle
'//* CancellaTabelle_PFunct                | Cancella le tabelle importate
'//* ATTIVA_ImportaModuli_PFunct           | Attiva la funzione per l’importazione dei moduli
'//* ATTIVA_CancellaModuli_PFunct          | Attiva la funzione per la cancellazione dei moduli
'//* ImportaModuli_PFunct                  | Importa i moduli (descrizione non esplicita nel codice)
'//* CancellaModuli_PFunct                 | Cancella i moduli importati
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
Dim ModuloCollection As Collection


'// DIM COLLECTION TABELLE
Private Type TabellaInfo
    Nome As String
    percorsoCompleto As String
End Type




'//2024.08.05_IMPORTAZIONE_SEMPLICE _
                modifico il modulo UTILITA_Msys_Mdl05_IMPORTAZIONE_SEMPLICE_DI_OGGETTI. Qui effetto l'importazione degli oggetti _
                in modo semplice cioè senza controllo del DATABASE ESTERNO. _
                Il controllo è solo SUGLI OGGETTI INTERNI.




'// #TASK_01 - InizializzaPath
'// Inizializza la path del database
Private Sub InizializzaPath()
    PathDb_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\MSYS\MDB\MSYS_N01_OGGETTI.mdb"
End Sub


'// #TASK_02 - InizializzaArrayQUERY
'// Inizializza gli array DELLE QUERY con i nomi degli oggetti
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


'// #TASK_03 - InizializzaArrayMACRO
'// Inizializza gli array DELLE MACRO con i nomi degli oggetti
Private Sub InizializzaArrayMACRO()
    On Error GoTo InizializzaArrayMACRO_Err

    MacroArray(0) = "UTILITA_MSys_Mcr00_}----------------------------------------@DLL"
    MacroArray(1) = "UTILITA_Msys_Mcr01_STAMPA_LIBRERIE_ATTIVE_DLL"

    Exit Sub

InizializzaArrayMACRO_Err:
    MsgBox "Errore durante l'inizializzazione degli array: " & Err.Description
End Sub


'//==================================================================================================//
'//         OGGETTO QUERY           *** INIZIO ***
'//==================================================================================================//







'//==================================================================================================//
'//         OGGETTO MACRO           *** INIZIO ***
'//==================================================================================================//

'// #TASK_05 - ATTIVA_ImportaMacro_PFunct
'// ATTIVO LA FUNZIONE
Private Function ATTIVA_ImportaMacro_PFunct()
    Bool1 = ImportaMacro_PFunct
End Function

'// #TASK_06 - ImportaMacro_PFunct
'// Funzione per importare le Macro
Public Function ImportaMacro_PFunct() As Boolean
    On Error GoTo ImportaMacro_PFunct_Err

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

    ImportaMacro_PFunct = True
    Exit Function

ImportaMacro_PFunct_Err:
    MsgBox "Errore durante l'importazione delle Macro: " & Err.Description
    ImportaMacro_PFunct = False
End Function


'// #TASK_07 - ATTIVA_CancellaMacro_Pfunct
Private Function ATTIVA_CancellaMacro_Pfunct()
    Bool1 = CancellaMacro_Pfunct
End Function

'// #TASK_08 - CancellaMacro_Pfunct
'// Funzione per cancellare tutte le Macro
Public Function CancellaMacro_Pfunct() As Boolean
    On Error GoTo CancellaMacro_Pfunct_Err

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

    CancellaMacro_Pfunct = True
    Exit Function

CancellaMacro_Pfunct_Err:
    MsgBox "Errore durante la cancellazione delle Macro: " & Err.Description
    CancellaMacro_Pfunct = False
End Function

'//==================================================================================================//
'//         OGGETTO MACRO           *** FINE ***
'//==================================================================================================//

'// #TASK_09 - ATTIVA_ImportaQuery_PFunct
'// ATTIVO LA FUNZIONE
Private Function ATTIVA_ImportaQuery_PFunct()
    Bool1 = ImportaQuery_PFunct
End Function

'// #TASK_10 - ImportaQuery_PFunct
'// Funzione per importare le query
Public Function ImportaQuery_PFunct() As Boolean
    On Error GoTo ImportaQuery_PFunct_Err

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

    ImportaQuery_PFunct = True
    Exit Function

ImportaQuery_PFunct_Err:
    MsgBox "Errore durante l'importazione delle query: " & Err.Description
    ImportaQuery_PFunct = False
End Function



'// #TASK_11 - ATTIVA_CancellaQuery_Funct
Private Function ATTIVA_CancellaQuery_Funct()
    Bool1 = CancellaQuery_Funct
End Function

'// #TASK_13 - ATTIVA_ImportaTabelle_PFunct
' // Funzione per cancellare tutte le query
Public Function CancellaQuery_Funct() As Boolean
    On Error GoTo CancellaQuery_Funct_Err

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

    CancellaQuery_Funct = True
    Exit Function

CancellaQuery_Funct_Err:
    MsgBox "Errore durante la cancellazione delle query: " & Err.Description
    CancellaQuery_Funct = False
End Function

'//==================================================================================================//
'//         OGGETTO QUERY           *** FINE ***
'//==================================================================================================//






'//==================================================================================================//
'//         OGGETTO TABLE           *** INIZIO ***
'//==================================================================================================//

'// #TASK_13 - ATTIVA_ImportaTabelle_PFunct
'// ATTIVA LA FUNZIONE DI COLLEGAMENTO
Public Function ATTIVA_ImportaTabelle_PFunct()
    MsgBox "ATTIVO LA FUNZIONE IMPORTA TABELLE", vbExclamation, "MACRO IMPORTA TABELLE"
    Bool1 = ImportaTabelle_Pfunct
End Function

'// #TASK_14 - InizializzaCollectionTABELLE
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


'// #TASK_15 - ImportaTabelle_Pfunct
'// Funzione per collegare (linkare) le tabelle
Public Function ImportaTabelle_Pfunct() As Boolean
    On Error GoTo ImportaTabelle_Pfunct_Err

    Dim ImportedTables() As String
    Dim TableCount As Integer
    Dim i As Long
    Dim tbl As Object

    TableCount = 0

    InizializzaCollectionTABELLE

    ReDim ImportedTables(0 To TabellaCollection.count - 1)

    ' Loop di collegamento delle tabelle
    For i = 1 To TabellaCollection.count
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
        TableCount = TableCount + 1
    Next i

    ' Riepilogo
    Debug.Print "Statistiche riepilogative delle tabelle collegate:"
    Debug.Print "Numero totale di tabelle collegate: " & TableCount
    For i = 0 To TableCount - 1
        Debug.Print "Tabella " & (i + 1) & ": " & ImportedTables(i)
    Next i

    ImportaTabelle_Pfunct = True
    Exit Function

ImportaTabelle_Pfunct_Err:
    MsgBox "Errore durante il collegamento delle tabelle: " & Err.Description
    ImportaTabelle_Pfunct = False
End Function

'// #TASK_16 - ATTIVA_CancellaTabelle_PFunct
'// ATTIVA LA FUNZIONE DI CANCELLAZIONE
Public Function ATTIVA_CancellaTabelle_PFunct()
    MsgBox "ATTIVO LA FUNZIONE CANCELLA TABELLE", vbExclamation, "MACRO CANCELLA TABELLE"
    Bool1 = CancellaTabelle_PFunct
End Function

'// #TASK_17 - CancellaTabelle_PFunct
'// Funzione per cancellare tutte le tabelle

Public Function CancellaTabelle_PFunct() As Boolean
    On Error GoTo CancellaTabelle_PFunct_Err

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

    For i = 1 To TabellaCollection.count
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

    CancellaTabelle_PFunct = True
    Exit Function

CancellaTabelle_PFunct_Err:
    MsgBox "Errore durante la cancellazione delle tabelle: " & Err.Description
    CancellaTabelle_PFunct = False
End Function







'//==================================================================================================//
'//         OGGETTO TABLE           *** FINE ***
'//==================================================================================================//




'//==================================================================================================//
'//         OGGETTO MODULI           *** INIZIO ***
'//==================================================================================================//


'// #TASK_18 - ATTIVA_ImportaModuli_PFunct
'// ATTIVO LA FUNZIONE
Public Function ATTIVA_ImportaModuli_PFunct()
    MsgBox "ATTIVO LA FUNZIONE IMPORTA I MODULI", vbExclamation, "MACRO IMPORTA MODULI"
    Bool1 = ImportaModuli_PFunct
End Function

'// #TASK_19 - ATTIVA_CancellaModuli_PFunct
'// ATTIVO LA FUNZIONE
Public Function ATTIVA_CancellaModuli_PFunct()
    MsgBox "ATTIVO LA FUNZIONE CANCELLA MODULI", vbExclamation, "MACRO CANCELLA MODULI"
    Bool1 = CancellaModuli_PFunct
End Function


'// #TASK_19.bis - InizializzaCollectionMODULI   @colletion@moduli
'// Inizializza la Collection DEI MODULI con nome e percorso completo per l'importazione
'// 2025.08.05_MSys_MODULI.mdb_(modifiche e miglioramenti) vedi progetto sublimetext
'//...........................................................................................................//

Private Sub InizializzaCollectionMODULI()
    On Error GoTo InizializzaCollectionMODULI_Err

    ' >>> Inizializza la collection che conterrà i moduli da importare
    Set ModuloCollection = New Collection

    ' >>> Variabili per costruzione del dizionario dei moduli
    Dim PercorsoBase_s As String                   ' Percorso base dei moduli da cui importare
    Dim modItem As Object                          ' Oggetto Dictionary che conterrà Nome e PercorsoCompleto
    Dim nomeModulo As String                       ' Nome del modulo da importare
    Dim percorsoCompleto As String                 ' Percorso completo del modulo da importare

    '01)_DLL
    ' -------------------------------------------------------------------------------------
    ' >>> Sezione dei moduli relativi alla DLL
        
        '//SUL DB MSys_DF13.mdb  @collection@moduli
        '//--------------------------------------------------------------------------------------//
                ' --- Modulo_00 ---
                ' Imposta il percorso di origine dei moduli
                    PercorsoBase_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_MODULI\MDB\MSys_MODULI.mdb"
        
                ' Nome simbolico e decorativo per il modulo
                nomeModulo = "CLASSE_GESTIONE_OGGETTI_DATABASE"
                percorsoCompleto = PercorsoBase_s
        
                ' Crea un nuovo dizionario per il modulo e aggiungilo alla collection
                Set modItem = CreateObject("Scripting.Dictionary")
                modItem.Add "Nome", nomeModulo
                modItem.Add "PercorsoCompleto", percorsoCompleto
                ModuloCollection.Add modItem
        
                ' --- Modulo_01 ---
                ' Nome del modulo effettivo da importare
                nomeModulo = "UTIL_MDL40_N08_RICERCA_TUTTI_GLI_OGGETTI_NEL_DB"
                percorsoCompleto = PercorsoBase_s
        
                ' Crea un nuovo dizionario per il secondo modulo e aggiungilo alla collection
                Set modItem = CreateObject("Scripting.Dictionary")
                modItem.Add "Nome", nomeModulo
                modItem.Add "PercorsoCompleto", percorsoCompleto
                ModuloCollection.Add modItem



                ' --- Modulo_02 ---
                ' Nome del modulo effettivo da importare
                nomeModulo = "UTIL_Nrx1003_n01_RECUPERO_PARAMETRI_TABELLA_OGGETTI"
                percorsoCompleto = PercorsoBase_s
        
                ' Crea un nuovo dizionario per il secondo modulo e aggiungilo alla collection
                Set modItem = CreateObject("Scripting.Dictionary")
                modItem.Add "Nome", nomeModulo
                modItem.Add "PercorsoCompleto", percorsoCompleto
                ModuloCollection.Add modItem

        '//--------------------------------------------------------------------------------------//
        
        
    ' >>> Tutti i moduli sono stati aggiunti alla collection
    Exit Sub

' >>> Gestione degli errori
InizializzaCollectionMODULI_Err:
    MsgBox "Errore durante l'inizializzazione della Collection MODULI: " & Err.Description
End Sub

'//...........................................................................................................//




'// #TASK_20 - ImportaModuli_PFunct
'//...........................................................................................................//
'// Funzione per importare i moduli da Collection (nome + path), cancellando eventuali moduli esistenti.
'// Registra in un file di log l'esito per ciascun modulo (importato o fallito).
            
            '//Comportamento corretto: _
            Per ogni modulo: _
            Se il modulo esiste nel DB corrente: _
            Lo cancella. _
            Poi tenta l’importazione. _
            Logga sia successi che errori. _
            Se il modulo NON esiste nel DB corrente: _
            Tenta l’importazione direttamente. _
            Logga comunque il risultato (successo o errore). _
            In entrambi i casi: _
            Scrive il risultato nel log: nome, percorso, eventuale errore.

'// #TASK_20 - ImportaModuli_PFunct
'// Funzione che importa moduli da file esterni.
'// Se il modulo esiste nel DB attuale, lo cancella prima dell'importazione.
'// Registra il risultato (successo/errore) nel file log "ImportaModuli_PFunct_log.txt".

Public Function ImportaModuli_PFunct() As Boolean
    On Error GoTo ImportaModuli_PFunct_Err

    ' >>> Blocco importazione se eseguita sul database di origine
    Dim percorsoDbCorrente As String
    percorsoDbCorrente = CurrentDb.Name

    Const percorsoDbOrigine As String = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_MODULI\MDB\MSys_MODULI.mdb"

    If StrComp(percorsoDbCorrente, percorsoDbOrigine, vbTextCompare) = 0 Then
        MsgBox "ERRORE: Non è consentito importare moduli nel database di origine." & vbCrLf & _
               "Percorso corrente: " & percorsoDbCorrente, vbCritical, "Importazione bloccata"
        ImportaModuli_PFunct = False
        Exit Function
    End If

    ' Inizializza la collection dei moduli
    InizializzaCollectionMODULI

    ' Variabili di tracciamento
    Dim moduliImportati As Integer
    Dim moduliNonImportati As Integer
    Dim nomiModuliImportati As String
    Dim nomiModuliNonImportati As String

    moduliImportati = 0
    moduliNonImportati = 0
    nomiModuliImportati = ""
    nomiModuliNonImportati = ""

    ' Per log
    Dim modItem As Object
    Dim nomeModulo As String
    Dim pathModulo As String
    Dim logFile As String
    Dim fnum As Integer
    Dim errDescrizione As String

    ' Percorso del file log sul desktop
    logFile = CurrentProject.Path & "\ImportaModuli_PFunct_log.txt"

    fnum = FreeFile

    ' Crea file log
    Open logFile For Output As #fnum
    Print #fnum, "LOG IMPORTAZIONE MODULI - " & Now
    Print #fnum, String(80, "-")

    ' Ciclo sui moduli da importare
    For Each modItem In ModuloCollection
        nomeModulo = modItem("Nome")
        pathModulo = modItem("PercorsoCompleto")

        ' Cancella modulo se già presente nel DB
        On Error Resume Next
        If DCount("*", "MSysObjects", "Name='" & nomeModulo & "' AND Type=-32761") > 0 Then
            DoCmd.DeleteObject acModule, nomeModulo
            If Err.Number <> 0 Then
                ' Errore nella cancellazione
                errDescrizione = "Errore cancellazione: " & Err.Description
                moduliNonImportati = moduliNonImportati + 1
                nomiModuliNonImportati = nomiModuliNonImportati & nomeModulo & ", "
                Print #fnum, "ERRORE: " & nomeModulo
                Print #fnum, "  Percorso: " & pathModulo
                Print #fnum, "  " & errDescrizione
                Print #fnum, String(40, "-")
                Err.Clear
                GoTo SaltaImportazione
            End If
        End If
        Err.Clear

        ' Tenta l’importazione
        DoCmd.TransferDatabase acImport, "Microsoft Access", pathModulo, acModule, nomeModulo, nomeModulo, False

        If Err.Number = 0 Then
            ' Successo
            moduliImportati = moduliImportati + 1
            nomiModuliImportati = nomiModuliImportati & nomeModulo & ", "
            Print #fnum, "IMPORTATO: " & nomeModulo
            Print #fnum, "  Percorso: " & pathModulo
        Else
            ' Errore nell'importazione
            errDescrizione = "Errore importazione: " & Err.Description
            moduliNonImportati = moduliNonImportati + 1
            nomiModuliNonImportati = nomiModuliNonImportati & nomeModulo & ", "
            Print #fnum, "ERRORE: " & nomeModulo
            Print #fnum, "  Percorso: " & pathModulo
            Print #fnum, "  " & errDescrizione
            Print #fnum, String(40, "-")
            Err.Clear
        End If

SaltaImportazione:
        On Error GoTo ImportaModuli_PFunct_Err
    Next modItem

    ' Chiudi log
    Close #fnum

    ' Rimuove virgole in eccesso
    If moduliImportati > 0 Then nomiModuliImportati = Left(nomiModuliImportati, Len(nomiModuliImportati) - 2)
    If moduliNonImportati > 0 Then nomiModuliNonImportati = Left(nomiModuliNonImportati, Len(nomiModuliNonImportati) - 2)

    ' Output riepilogo
    Debug.Print "Totale moduli importati: " & moduliImportati
    Debug.Print "Nomi dei moduli importati: " & nomiModuliImportati
    Debug.Print "Totale moduli NON importati: " & moduliNonImportati
    Debug.Print "Nomi dei moduli NON importati: " & nomiModuliNonImportati

    MsgBox "Importazione completata." & vbCrLf & _
           "Moduli importati: " & moduliImportati & vbCrLf & _
           "Moduli non importati: " & moduliNonImportati & vbCrLf & _
           "Vedi il file log su Desktop: ImportaModuli_PFunct_log.txt", vbInformation

    ImportaModuli_PFunct = True
    Exit Function

' Gestione errore globale
ImportaModuli_PFunct_Err:
    MsgBox "Errore durante l'importazione: " & Err.Description, vbCritical
    On Error Resume Next
    Close #fnum
    ImportaModuli_PFunct = False
End Function
'//...........................................................................................................//




'// #TASK_21 - CancellaModuli_PFunct
'//...........................................................................................................//
'// Funzione per cancellare tutti i moduli _
        cancellazione solo nel DB corrente (come già faceva, perché DoCmd.DeleteObject agisce sul DB aperto), _
        controllo esistenza prima di cancellare, _
        contabilizzazione di moduli cancellati e non trovati, _
        messaggio finale con riepilogo. _




Public Function CancellaModuli_PFunct() As Boolean
    On Error GoTo CancellaModuli_PFunct_Err

    ' Inizializza la collection dei nomi dei moduli da cancellare (solo nomi, senza percorso)
    InizializzaCollectionMODULI

    Dim moduliCancellati As Integer
    Dim moduliNonTrovati As Integer
    Dim nomiModuliCancellati As String
    Dim nomiModuliNonTrovati As String
    Dim modItem As Object
    Dim nomeModulo As String

    moduliCancellati = 0
    moduliNonTrovati = 0
    nomiModuliCancellati = ""
    nomiModuliNonTrovati = ""
    
    
    ' >>> Blocco importazione se eseguita sul database di origine
    Dim percorsoDbCorrente As String
    percorsoDbCorrente = CurrentDb.Name

    Const percorsoDbOrigine As String = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_MODULI\MDB\MSys_MODULI.mdb"
    
    If StrComp(percorsoDbCorrente, percorsoDbOrigine, vbTextCompare) = 0 Then
        MsgBox "ERRORE: Non è consentito importare moduli nel database di origine." & vbCrLf & _
               "Percorso corrente: " & percorsoDbCorrente, vbCritical, "Importazione bloccata"
        CancellaModuli_PFunct = False
        Exit Function
    End If

    ' Ciclo su ogni modulo nella collection per verificarne la presenza e cancellarlo
    For Each modItem In ModuloCollection
        nomeModulo = modItem("Nome")

        ' Controllo se il modulo esiste nel db corrente tramite MSysObjects
        If DCount("*", "MSysObjects", "Name='" & nomeModulo & "' AND Type=-32761") > 0 Then
            ' Cancella il modulo
            DoCmd.DeleteObject acModule, nomeModulo
            Debug.Print "Modulo cancellato: " & nomeModulo
            moduliCancellati = moduliCancellati + 1
            nomiModuliCancellati = nomiModuliCancellati & nomeModulo & ", "
        Else
            Debug.Print "Modulo non trovato: " & nomeModulo
            moduliNonTrovati = moduliNonTrovati + 1
            nomiModuliNonTrovati = nomiModuliNonTrovati & nomeModulo & ", "
        End If
    Next modItem

    ' Rimuove l'ultima virgola e spazio
    If moduliCancellati > 0 Then
        nomiModuliCancellati = Left(nomiModuliCancellati, Len(nomiModuliCancellati) - 2)
    End If

    If moduliNonTrovati > 0 Then
        nomiModuliNonTrovati = Left(nomiModuliNonTrovati, Len(nomiModuliNonTrovati) - 2)
    End If

    ' Stampa riepilogo
    Debug.Print "Totale moduli cancellati: " & moduliCancellati
    Debug.Print "Nomi moduli cancellati: " & nomiModuliCancellati
    Debug.Print "Totale moduli non trovati: " & moduliNonTrovati
    Debug.Print "Nomi moduli non trovati: " & nomiModuliNonTrovati

    MsgBox "Cancellazione moduli completata." & vbCrLf & _
           "Moduli cancellati: " & moduliCancellati & vbCrLf & _
           "Moduli non trovati: " & moduliNonTrovati, vbInformation, "Risultato cancellazione moduli"

    CancellaModuli_PFunct = True
    Exit Function

CancellaModuli_PFunct_Err:
    MsgBox "Errore durante la cancellazione dei moduli: " & Err.Description, vbCritical
    CancellaModuli_PFunct = False
End Function



'//...........................................................................................................//

'//==================================================================================================//
'//         OGGETTO MODULI           *** FINE ***
'//==================================================================================================//


'// #TASK_22 - ImportaTuttiGliOggetti
'// Funzione generale per importare tutti gli oggetti
Public Function ImportaTuttiGliOggetti() As Boolean
    On Error GoTo ImportaTuttiGliOggetti_Err

    ImportaQuery_PFunct
    ImportaTabelle_Pfunct
    ImportaModuli_PFunct
    ImportaMacro_PFunct

    ImportaTuttiGliOggetti = True
    Exit Function

ImportaTuttiGliOggetti_Err:
    MsgBox "Errore durante l'importazione di tutti gli oggetti: " & Err.Description
    ImportaTuttiGliOggetti = False
End Function


'// #TASK_23 - CancellaTuttiGliOggetti
'// Funzione generale per cancellare tutti gli oggetti
Public Function CancellaTuttiGliOggetti() As Boolean
    On Error GoTo CancellaTuttiGliOggetti_Err

    CancellaQuery_Funct
    CancellaTabelle_PFunct
    CancellaModuli_PFunct
    CancellaMacro_Pfunct
    MsgBox "Tutti gli oggetti importati o collegati sono stati cancellati."

    CancellaTuttiGliOggetti = True
    Exit Function

CancellaTuttiGliOggetti_Err:
    MsgBox "Errore durante la cancellazione di tutti gli oggetti: " & Err.Description
    CancellaTuttiGliOggetti = False
End Function


'//********************************************************************************//









