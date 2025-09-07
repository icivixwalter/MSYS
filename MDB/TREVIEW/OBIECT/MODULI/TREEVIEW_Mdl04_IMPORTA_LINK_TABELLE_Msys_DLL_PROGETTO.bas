Attribute VB_Name = "TREEVIEW_Mdl04_IMPORTA_LINK_TABELLE_Msys_DLL_PROGETTO"
'//MODULO PER L'IMPORTAZIONE E LINK DELLE TABELLE MSYS: _
    TREEVIEW_Mdl04_IMPORTA_LINK_TABELLE_Msys_DLL_PROGETTO _
        Viene attivata con la macro: TREEVIEW_Mcr04_IMPORTA_LINK_TABELLE_Msys_DLL_PROGETTO
    
'//@collega@tabelle@dll_(@MODULO per collegare le tabelle relative alla librerie DLL)

Option Compare Database
Option Explicit

'//DIM LE VARIABILI GENERALI
Dim Bool1 As Boolean

' // DIM variabili procedura importa e collega DLL
Dim PathDb_s As String

' // Array globali per i nomi degli oggetti
Dim QueryArray(7) As String
Dim TableArray(3) As String
Dim ModuleArray(1) As String
Dim MacroArray(1) As String

' // Inizializza la path del database
Private Sub InizializzaPath()
    PathDb_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSYS_N01_OGGETTI.mdb"
End Sub

' // Inizializza gli array DELLE TABELLE con i nomi degli oggetti
Private Sub InizializzaArrayTABELLE()
    On Error GoTo InizializzaArrayTABELLE_Err

    TableArray(0) = "MSys_DF13_}----------------------------------------------------@"
    TableArray(1) = "Msys_DF13_DLL_PROGETTO"
    TableArray(2) = "Msys_DF14_DLL_LIBRERIE"
    TableArray(3) = "Msys_DLL_LIBRERIE"

    Exit Sub

InizializzaArrayTABELLE_Err:
    MsgBox "Errore durante l'inizializzazione degli array: " & Err.Description
End Sub

'//==================================================================================================//
'//         OGGETTO TABLE           *** INIZIO ***
'//==================================================================================================//

'//ATTIVO LA FUNZIONE
Private Sub ATTIVA_Importa_LINK_Tabelle_Msys_PFunct()
    Bool1 = ImportaLINKTabelle_Msys_PFunct
End Sub

' // Funzione per importare le tabelle
Public Function ImportaLINKTabelle_Msys_PFunct() As Boolean
    On Error GoTo ImportaLINKTabelle_Msys_PFunct_Err

    Dim ImportedTables() As String
    Dim TableCount As Integer
    Dim i As Integer
    
    '//reset
    TableCount = 0

    InizializzaArrayTABELLE
    InizializzaPath

    ' Redimensiona l'array ImportedTables per contenere il numero massimo di tabelle
    ReDim ImportedTables(LBound(TableArray) To UBound(TableArray))

    ' Importa le tabelle
    For i = LBound(TableArray) To UBound(TableArray)
        ' Verifica se la tabella esiste
        If TableExists(TableArray(i)) Then
            DoCmd.DeleteObject acTable, TableArray(i)
        End If
        
        ' Controllo tabella
        Debug.Print "tabella da importare: " & TableArray(i)
        Debug.Print "database : " & PathDb_s
        
        ' Copia la tabella dal database di origine al database corrente
        DoCmd.TransferDatabase acLink, "Microsoft Access", PathDb_s, acTable, TableArray(i), TableArray(i), False
        
        ' Aggiungi il nome della tabella importata all'array
        ImportedTables(TableCount) = TableArray(i)
        TableCount = TableCount + 1
        
    Next i

            'vado all'ultima tabella salvata
            DoCmd.SelectObject acTable, "Msys_DLL_LIBRERIE", True
            
    ' Stampa le statistiche riepilogative
    Debug.Print "Statistiche riepilogative delle tabelle importate:"
    Debug.Print "Numero totale di tabelle importate: " & TableCount
    For i = 0 To TableCount - 1
        Debug.Print "Tabella " & (i + 1) & ": " & ImportedTables(i)
    Next i

    ImportaLINKTabelle_Msys_PFunct = True
    Exit Function

ImportaLINKTabelle_Msys_PFunct_Err:
    MsgBox "Errore durante l'importazione delle tabelle: " & Err.Description
    ImportaLINKTabelle_Msys_PFunct = False
    
End Function

' Funzione per verificare se una tabella esiste
Private Function TableExists(TableName As String) As Boolean
    On Error Resume Next
    TableExists = (DCount("*", "MSysObjects", "Name='" & TableName & "' AND Type IN (1, 4, 6)") > 0)
    On Error GoTo 0
End Function

