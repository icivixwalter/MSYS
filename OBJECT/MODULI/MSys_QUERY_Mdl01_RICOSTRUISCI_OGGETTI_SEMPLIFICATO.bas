Attribute VB_Name = "MSys_QUERY_Mdl01_RICOSTRUISCI_OGGETTI_SEMPLIFICATO"
'//*********************************************************************************************//'
'//** DENOMINAZIONE DEL MODULO + DETTAGLI OPERATIVI + '


'-------------------------------------------------------------------------------------------'
'                                                                                           '
'                       DENOMINAZIONE                                                       '
' MODULO: MSys_QUERY_Mdly01_RICOSTRUISCI_OGGETTI_SEMPLIFICATO                               '
' @modulo@attiva@Macro_(Modulo vba che attiva i collegamenti degli @impegni)                '
'                                                                                           '
'-------------------------------------------------------------------------------------------'


'-------------------------------------------------------------------------------------------'
'                                                                                           '
'                       DETTAGLI OPERATIVI                                                  '
'                                                                                           '
'                                                                                           '
'                                                                                           '
'   Dettagli della Soluzione                                                                '
'   1. Variabili fisse per percorso e file:                                                 '
'                                                                                           '
'       dbPath, dbFile, e fullPath definiscono il percorso e il nome del file una           '
'       volta sola, rendendo il codice più leggibile e mantenibile.                         '
'   2. Collection per i nomi delle tabelle:                                                 '
'       I nomi delle tabelle vengono inseriti in una Collection. Questo ti consente         '
'       di gestire facilmente più tabelle aggiungendo o rimuovendo elementi                 '
'       dalla collection.                                                                   '
'   2Ciclo For Each:                                                                        '
'                                                                                           '
'       Un ciclo For Each semplifica l'iterazione sui nomi delle tabelle e applica          '
'       le stesse operazioni a ciascuna.                                                    '
'       Gestione degli errori:                                                              '
'                                                                                           '
'       On Error Resume Next evita che il codice si interrompa se una tabella non           '
'           esiste, mentre On Error GoTo 0 ripristina la gestione                           '
'           degli errori al termine del ciclo.                                              '
'   Vantaggi                                                                                '
'   1) Manutenibilità: Per aggiungere nuove tabelle, basta inserirne il nome                '
'       nella Collection senza duplicare il codice.                                         '
'   senza duplicare il codice.                                                              '
'   2) Leggibilità: Il codice è più compatto e facile da seguire.                           '
'   3) Efficienza: Elimina operazioni ripetitive, riducendo il rischio di errori manuali.   '
'                                                                                           '
'   FUNZIONE collegata alla macro:                                                          '
'       *** da creare ---> LLPP_IMPEGNI_Tb02Mcr09_RICOSTRUISCI_TUTTI_GLI_IMPEGNI            '
'                                                                                           '
'                                                                                           '
'-------------------------------------------------------------------------------------------'

    
'//OPERAZIONI ESEGUITE: _
01.LINK_TABELLE= Costruisci da zero il LINK delle tabelle per la gestione DEL DATABASE MSYS_QUERY_GE _
02.IMPORT_QUERY= Importa le query per la gestione DEL DATABASE MSYS_QUERY_GE _
03_LINK=FINE RISULTATO IN MSG

'//*********************************************************************************************//'



Option Compare Database
Option Explicit






'//01.LINK_TABELLE= Costruisci da zero il LINK delle tabelle per la gestione DEL DATABASE MSYS_QUERY_GE
Sub ATTIVA_LINK_MULTI_DB()
    Call IMPORTA_LINK_TABLE_Collection_pFunct
End Sub


'//LINK TABELLE MULTIDATABASE
'//================================================================================================================//
'//NOTE: Cosa Fa il Codice Aggiornato _
    Collections per Successi e Fallimenti: _
    collegate: Memorizza i nomi delle tabelle collegate con successo. _
    nonCollegate: Memorizza i nomi delle tabelle che non sono state collegate. _
    Controllo Errori: _
    Dopo ogni tentativo di collegamento (DoCmd.TransferDatabase), verifica se _l'errore (Err.Number) _
    è diverso da zero per decidere se aggiungere il nome _della tabella a nonCollegate. _
    Messaggio Dettagliato: Alla fine, costruisce un messaggio che elenca tutte le tabelle collegate e _quelle non collegate._

Public Function IMPORTA_LINK_TABLE_Collection_pFunct()

'//ATTENZIONE: per inserire un nuovo database occorre : _
    a) inserire una nuova collectio Dim tblCollectionXX As Collection _
    b) qualificare il nuovo database con la collectio : _
        ' Database 1: TABELLE DEL PROGETTO MSys_QUERY.mdb
        'Set tblCollectionXX = New Collection
        ' tblCollectionXX.Add "MSys_QUERY"   .... _
    c) inserire la nuova collection ed la nuova PATH + DATABSE + COLLECTION _
      dbCollection.Add Array("C:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\", _
                           "MSys_DF02.mdb", tblCollectionXX)

    On Error GoTo IMPORTA_LINK_TABLE_Collection_pFunct_Err
    
    '//dim e set
    '//.....................................................................//
        Dim dbCollection As Collection
        Dim dbItem As Variant
        Dim tblName As Variant
    
        Dim tblCollection As Collection
        Dim collegate As Collection
        Dim nonCollegate As Collection
    
        Dim dbPath As String
        Dim dbFile As String
        Dim fullPath As String
        Dim msg As String
    
        ' Inizializzo le collections principali
        Set dbCollection = New Collection
        Set collegate = New Collection
        Set nonCollegate = New Collection
    '//.....................................................................//
    
    '//01_LINK=INIZIALIZZO
    ' Richiamo la subroutine che inizializza dbCollection
    Call InizializzaCollezioneTabelle(dbCollection)
    
    '//02_LINK=ITERNO NELLE COLLECTION
    ' Iterazione sui database
    For Each dbItem In dbCollection
        dbPath = dbItem(0)              ' Percorso del database
        dbFile = dbItem(1)              ' Nome del file MDB
        fullPath = dbPath & dbFile      ' PATH COMPLETA
        
        '//Collections per Successi e Fallimenti: _
        collegate: Memorizza i nomi delle tabelle collegate con successo. _
        nonCollegate: Memorizza i nomi delle tabelle che non sono state collegate.
        '//.............................................................................................//
                    Set tblCollection = dbItem(2)  ' Collection delle tabelle da collegare
            
                    ' Collegamento delle tabelle
                    For Each tblName In tblCollection
                        On Error Resume Next    ' Ripristina la gestione degli errori
                        DoCmd.DeleteObject acTable, tblName ' Elimina tabella se già esistente
                        ' se il collegamento da errore vuol dire che nel db corrente non esiste e non viene cancellata _
                            passa direttamente al collegamento della nuova tabella
                        Err.Clear
                        
                        '//COLLEGA
                        DoCmd.TransferDatabase acLink, "Microsoft Access", fullPath, acTable, tblName, tblName, False
            
                        If Err.Number = 0 Then
                            ' Se l'operazione è riuscita, aggiungi alla lista collegate
                            collegate.Add tblName & " (" & dbFile & ")"
                        Else
                            ' Se c'è stato un errore, aggiungi alla lista non collegate
                            nonCollegate.Add tblName & " (" & dbFile & ")"
                        End If
                    Next tblName
                Next dbItem
            
        '//.............................................................................................//
    
    
    '//03_LINK=FINE RISULTATO IN MSG
    '//MESSAGGIO FINALE
    '//---------------------------------------------------------------------//
      ' Costruzione del messaggio finale
      msg = "TABELLE COLLEGATE CORRETTAMENTE:" & vbCrLf
      For Each tblName In collegate
          msg = msg & "- " & tblName & vbCrLf
      Next tblName
    
        If nonCollegate.Count > 0 Then
            msg = msg & vbCrLf & "TABELLE NON COLLEGATE (verifica il percorso o il nome):" & vbCrLf
            For Each tblName In nonCollegate
                msg = msg & "- " & tblName & vbCrLf
            Next tblName
        End If
    
        MsgBox msg, vbInformation, "Risultato Collegamento Tabelle"

    '//---------------------------------------------------------------------//

'// USCITA E CONTROLLO ERRORI
'//................................................................................................//
IMPORTA_LINK_TABLE_Collection_pFunct_Exit:
    ' Pulizia delle Collection
    Set dbCollection = Nothing
    Set tblCollection = Nothing
    Set collegate = Nothing
    Set nonCollegate = Nothing
    Exit Function

IMPORTA_LINK_TABLE_Collection_pFunct_Err:
    MsgBox "Errore: " & Err.Description, vbCritical, "Errore"
    Resume IMPORTA_LINK_TABLE_Collection_pFunct_Exit
End Function

'//INIZIALIZZO LE COLLECTION DELLE TABELLE
'//@LE@TABELLE@DA@COLLEGARE_(tabelle degli @IMPEGNI da collegare con elenco da aggiungere o ridurre)
Private Sub InizializzaCollezioneTabelle(dbCollection As Collection)
    Dim tblCollection1 As Collection
    Dim tblCollection2 As Collection
    Dim tblCollection3 As Collection

    ' Database 1: TABELLE DEL PROGETTO MSys_QUERY.mdb
    Set tblCollection1 = New Collection
    tblCollection1.Add "MSys_QUERY"
    tblCollection1.Add "MSys_QUERY_TMP"
    dbCollection.Add Array("C:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\", _
                           "MSys_QUERY.mdb", tblCollection1)

    ' Database 2: TABELLE DEL PROGETTO MSys_DF01.mdb
    Set tblCollection2 = New Collection
    tblCollection2.Add "MSys_DF01_}----------------------------------------------------@"
    tblCollection2.Add "MSys_DF01_COMANDI"
    dbCollection.Add Array("C:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\", _
                           "MSys_DF01.mdb", tblCollection2)

    ' Database 3: TABELLE DEL PROGETTO MSys_DF02.mdb
    Set tblCollection3 = New Collection
    tblCollection3.Add "MSys_{@===================================================@}_Tab"
    tblCollection3.Add "MSys_DF02_}----------------------------------------------------@"
    tblCollection3.Add "MSys_DF02_TIPO_OGGETTO"
    dbCollection.Add Array("C:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\", _
                           "MSys_DF02.mdb", tblCollection3)

End Sub


'//LINK TABELLE MULTIDATABASE  *** LINK FINE ***
'//================================================================================================================//




'//*** IMPORTA_OGGETTI_QUERY_Collection_pFunct ***
'//======================================================================================================//
'//NOTE:   @importa@query@semplice_(importa le query in modalita semplice ossia con una Array o Collection non tramite Db)



'//02.IMPORT_QUERY= Importa le query per la gestione DEL DATABASE MSYS_QUERY_GE
Sub ATTIVA_IMPORTA_QUERY()
    Call IMPORTA_OGGETTI_QUERY_Collection_pFunct
End Sub

Public Function IMPORTA_OGGETTI_QUERY_Collection_pFunct()
    On Error GoTo IMPORTA_OGGETTI_QUERY_Collection_pFunct_Err
    
    Dim dbCollection As Collection
    Dim dbItem As Variant
    Dim qryName As Variant
    
    Dim qryCollection As Collection
    Dim importate As Collection
    Dim nonImportate As Collection
    
    Dim dbPath As String
    Dim dbFile As String
    Dim fullPath As String
    Dim msg As String
    
    ' Inizializzo le collections
    Set dbCollection = New Collection ' Per i vari database e relative query
    Set importate = New Collection
    Set nonImportate = New Collection

    ' ** Aggiunta dei Database e relative query da importare **
    ' Ogni elemento della dbCollection è un array con Path, File e Collection di query
    dbCollection.Add Array("C:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\", _
                           "MSys_DF01.mdb", Array("MSys_DF01_COMANDI_Qry01_01_SELECT_TUTTI", _
                                                   "MSys_DF01_COMANDI_Qry01_02_GROUP_TUTTI"))

    dbCollection.Add Array("C:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\", _
                           "MSys_QUERY.mdb", Array("MSys_QUERY_Qry00_{@===========================================@}", _
                                                 "MSys_QUERY_Qry01_}---------------------------------------------@", _
                                                 "MSys_QUERY_Qry01_01_SELECT_TUTTE", _
                                                 "MSys_QUERY_Qry01_01_SELECT_TUTTE_TMP", _
                                                 "MSys_QUERY_Qry01_70_}------------------------------------@INSERT", _
                                                 "MSys_QUERY_Qry01_71_INSERT_toMsys_DF12", _
                                                 "MSys_QUERY_Qry01_80_}------------------------------------@DELETE", _
                                                 "MSys_QUERY_Qry01_81_DELETE_TUTTE" _
                                                 ))

    ' Iterazione su ogni database
    For Each dbItem In dbCollection
        dbPath = dbItem(0)
        dbFile = dbItem(1)
        fullPath = dbPath & dbFile
        Set qryCollection = New Collection
        
        ' Inizializza la collection delle query per questo database
        Dim i As Integer
        For i = LBound(dbItem(2)) To UBound(dbItem(2))
            qryCollection.Add dbItem(2)(i)
        Next i
        
        ' Importazione delle query
        For Each qryName In qryCollection
            On Error Resume Next
            DoCmd.DeleteObject acQuery, qryName ' Elimina query se già esistente
            Err.Clear
            
            DoCmd.TransferDatabase acImport, "Microsoft Access", fullPath, acQuery, qryName, qryName
            
            If Err.Number = 0 Then
                importate.Add qryName & " (" & dbFile & ")"
            Else
                nonImportate.Add qryName & " (" & dbFile & ")"
            End If
        Next qryName
    Next dbItem
    
    ' Costruzione del messaggio finale
    msg = "QUERY IMPORTATE CORRETTAMENTE:" & vbCrLf
    For Each qryName In importate
        msg = msg & "- " & qryName & vbCrLf
    Next qryName
    
    If nonImportate.Count > 0 Then
        msg = msg & vbCrLf & "QUERY NON IMPORTATE (verifica il percorso o il nome):" & vbCrLf
        For Each qryName In nonImportate
            msg = msg & "- " & qryName & vbCrLf
        Next qryName
    End If
    
    MsgBox msg, vbInformation, "Risultato Importazione Query"

IMPORTA_OGGETTI_QUERY_Collection_pFunct_Exit:
    Set dbCollection = Nothing
    Set qryCollection = Nothing
    Set importate = Nothing
    Set nonImportate = Nothing
    Exit Function

IMPORTA_OGGETTI_QUERY_Collection_pFunct_Err:
    MsgBox "Errore: " & Err.Description, vbCritical, "Errore"
    Resume IMPORTA_OGGETTI_QUERY_Collection_pFunct_Exit
End Function

'//*** IMPORTA_OGGETTI_QUERY_Collection_pFunct ***
'//======================================================================================================//

