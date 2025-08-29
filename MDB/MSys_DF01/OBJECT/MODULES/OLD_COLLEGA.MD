
Option Compare Database
Option Explicit

'//Manutenzione semplificata: per aggiungere o rimuovere tabelle, ti basta aggiornare l’array tabelle. _
    Codice più compatto e leggibile. _
    Separazione chiara tra logica (ciclo) e dati (path e nomi tabelle). _
    Aggiungiamo un controllo per verificare che entrambi i valori (tabName e mdbFile) siano valorizzati prima di procedere:

Function COLLEGA_MSys_DF01_COMANDI_PFunct()
    On Error GoTo COLLEGA_MSys_DF01_COMANDI_PFunct_Err

    Dim basePath As String
    basePath = "c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_DF01\MDB\"
                
    ' Array di oggetti con nome tabella e nome file MDB, i database possono essere diversi
    Dim tabelle As Variant
    tabelle = Array( _
        Array("MSys_DF01_}----------------------------------------------------@", "MSys_DF01.mdb"), _
        Array("MSys_DF01_COMANDI", "MSys_DF01.mdb"), _
        Array("", "MSys_DF01.mdb"), _
        Array("", "MSys_DF01.mdb"), _
        Array("", "MSys_DF01.mdb"), _
        Array("", "MSys_DF01.mdb"), _
        Array("", "MSys_DF01.mdb") _
    )

    Dim i As Integer
    For i = LBound(tabelle) To UBound(tabelle)
        Dim tabName As String
        Dim mdbFile As String
        Dim fullPath As String

        tabName = Trim(tabelle(i)(0))
        mdbFile = Trim(tabelle(i)(1))

        ' Verifica che entrambi i valori siano presenti
        If tabName <> "" And mdbFile <> "" Then
            fullPath = basePath & mdbFile

            ' Elimina la tabella se già esistente
            On Error Resume Next
            DoCmd.DeleteObject acTable, tabName
            On Error GoTo COLLEGA_MSys_DF01_COMANDI_PFunct_Err

            ' Collega la tabella
            DoCmd.TransferDatabase acLink, "Microsoft Access", fullPath, acTable, tabName, tabName, False
        End If
    Next i

COLLEGA_MSys_DF01_COMANDI_PFunct_Exit:
    Exit Function

COLLEGA_MSys_DF01_COMANDI_PFunct_Err:
    MsgBox "Errore: " & Err.Description, vbExclamation, "Errore di collegamento tabelle"
    Resume COLLEGA_MSys_DF01_COMANDI_PFunct_Exit
End Function



Option Compare Database

'//*************************************************************************************************//'
'//*
'//* MODULO : COLLEGA_Mcr_GEST_MENU_TB01_VOCI
'//*
'//* Per ricapitolare, il modulo rimane così strutturato:
'//* @01 GetTabellaPaths ? contiene solo i dati (nomi tabelle e percorsi MDB).
'//* @02 CollegaTabelle_PFunct ? collega le tabelle, crea log e messaggi.
'//* @03 CancellaTabelleCollegate ? elimina i collegamenti senza log né conferma.
'//* @04 EliminaTabellaSeEsiste ? verifica ed elimina una tabella se esiste.
'//* Questa versione è più leggera e diretta, adatta se vuoi una cancellazione
'//* rapida senza interazioni aggiuntive.
'//*

'//*************************************************************************************************//'

' ============================================================
' @01 FUNZIONE GetTabellaPaths
' Restituisce la collection di tabelle e percorsi MDB da collegare
' ============================================================
Public Function GetTabellaPaths() As Collection
    Dim TabellaPath As New Collection

    ' === Primo gruppo ===
    TabellaPath.Add Array("GEST_MENU_TB01_VOCI", "c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_DF01\MDB\MENU_TB01_VOCI.mdb")
    TabellaPath.Add Array("GEST_MENU_TB01_VOCI_TMP", "c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_DF01\MDB\MENU_TB01_VOCI.mdb")
    TabellaPath.Add Array("GEST_MENU_Tb01_}-----------------------------------------------@", "c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_DF01\MDB\MENU_TB01_VOCI.mdb")
    TabellaPath.Add Array("GEST_MENU_{@==================================================@}", "c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_DF01\MDB\MENU_TB01_VOCI.mdb")

    ' === Secondo gruppo ===
    TabellaPath.Add Array("GEST_MENU_TB02_FUNZIONI", "c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_DF01\MDB\MENU_TB02_FUNZIONI.mdb")
    TabellaPath.Add Array("GEST_MENU_TB02_FUNZIONI_TMP", "c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_DF01\MDB\MENU_TB02_FUNZIONI.mdb")
    TabellaPath.Add Array("GEST_MENU_Tb02_}-----------------------------------------------@", "c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_DF01\MDB\MENU_TB02_FUNZIONI.mdb")

    ' === Terzo gruppo ===
    TabellaPath.Add Array("MSys_DF00_}----------------------------------------------------@", "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSYS_N01_OGGETTI.mdb")
    TabellaPath.Add Array("MSys_DF01_COMANDI", "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSYS_N01_OGGETTI.mdb")
    TabellaPath.Add Array("MSys_DF01_COMANDI_TMP", "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSYS_N01_OGGETTI.mdb")
    TabellaPath.Add Array("MSys_DF02_TIPO_OGGETTO", "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSYS_N01_OGGETTI.mdb")

    Set GetTabellaPaths = TabellaPath
End Function


' ============================================================
' @02 FUNZIONE CollegaTabelle_PFunct
' Collega le tabelle del database corrente ai MDB esterni
' ============================================================
Public Function CollegaTabelle_PFunct()
    Dim TabellaPath As Collection
    Set TabellaPath = GetTabellaPaths()
    
    Dim NomeTabella As String
    Dim PercorsoFile As String
    Dim i As Long
    Dim LogFile As Integer
    Dim LogFileName As String
    Dim ConteggioOk As Long
    Dim ConteggioErr As Long

    ' @02.01 Creazione file log
    LogFileName = CurrentProject.Path & "\COLLEGA_Mcr_GEST_MENU_TB01_VOCI_LOG.txt"
    LogFile = FreeFile
    Open LogFileName For Output As #LogFile

    Print #LogFile, "Log esecuzione funzione: COLLEGA_Mcr_GEST_MENU_TB01_VOCI"
    Print #LogFile, "Data/Ora: " & Now()
    Print #LogFile, "--------------------------------------------"

    ' @02.02 Ciclo su tutte le tabelle
    For i = 1 To TabellaPath.Count
        Dim Item As Variant
        Item = TabellaPath(i)            ' Array contenente NomeTabella e PercorsoFile
        NomeTabella = Item(0)
        PercorsoFile = Item(1)

        On Error Resume Next
        EliminaTabellaSeEsiste NomeTabella
        DoCmd.TransferDatabase acLink, "Microsoft Access", PercorsoFile, acTable, NomeTabella, NomeTabella, False

        If Err.Number = 0 Then
            Print #LogFile, Now & " - OK: " & NomeTabella & " -> " & PercorsoFile
            ConteggioOk = ConteggioOk + 1
        Else
            Print #LogFile, Now & " - ERRORE: " & NomeTabella & " -> " & Err.Description
            ConteggioErr = ConteggioErr + 1
            Err.Clear
        End If
        On Error GoTo 0
    Next i

    ' @02.03 Riepilogo log
    Print #LogFile, "--------------------------------------------"
    Print #LogFile, "TOTALE TABELLE COLLEGATE: " & ConteggioOk
    Print #LogFile, "TOTALE ERRORI: " & ConteggioErr
    Close #LogFile

    ' @02.04 Messaggio a video
    MsgBox "Operazione completata!" & vbCrLf & _
           "Tabelle collegate: " & ConteggioOk & vbCrLf & _
           "Errori: " & ConteggioErr & vbCrLf & _
           "Log salvato in: " & LogFileName, vbInformation
End Function


' ============================================================
' @03 FUNZIONE CancellaTabelleCollegate
' Elimina i collegamenti alle tabelle presenti nel database corrente
' ============================================================
Public Sub CancellaTabelleCollegate()
    Dim TabellaPath As Collection
    Set TabellaPath = GetTabellaPaths()

    Dim i As Long
    For i = 1 To TabellaPath.Count
        Dim Item As Variant
        Item = TabellaPath(i)
        EliminaTabellaSeEsiste Item(0)
    Next i

    MsgBox "Cancellazione dei collegamenti completata!", vbInformation
End Sub



' ============================================================
' @04 FUNZIONE EliminaTabellaSeEsiste
' Controlla se la tabella esiste e la elimina
' ============================================================
Public Sub EliminaTabellaSeEsiste(NomeTabella As Variant)
    On Error Resume Next
    Dim tdf As DAO.TableDef
    For Each tdf In CurrentDb.TableDefs
        If tdf.Name = NomeTabella Then
            CurrentDb.TableDefs.Delete NomeTabella
            Exit For
        End If
    Next tdf
    On Error GoTo 0
End Sub





' ============================================================
' @05 FUNZIONE CancellaTutteTabelleCollegate_PFunct
' Cancella tutte le tabelle collegate definite in GetTabellaPaths
' solo se effettivamente esistono nel db corrente
' ============================================================
' ============================================================
' @05 FUNZIONE CancellaTutteTabelleCollegate_PFunct
' Cancella tutte le tabelle collegate definite in GetTabellaPaths
' solo se effettivamente esistono nel db corrente
' con log dei risultati su file
' ============================================================
Public Function CancellaTutteTabelleCollegate_PFunct()
    Dim TabellaPath As Collection
    Set TabellaPath = GetTabellaPaths()   ' Precarica la collection delle tabelle

    Dim i As Long
    Dim Item As Variant
    Dim ConteggioOk As Long
    Dim ConteggioErr As Long
    Dim tdf As DAO.TableDef
    Dim TabellaEsiste As Boolean
    Dim LogFile As Integer
    Dim LogFileName As String

    ' Inizializza log
    LogFileName = CurrentProject.Path & "\CANC_TutteTabelleCollegate_PFunct_LOG.txt"
    LogFile = FreeFile
    Open LogFileName For Output As #LogFile

    Print #LogFile, "Log esecuzione funzione: CancellaTutteTabelleCollegate_PFunct"
    Print #LogFile, "Data/Ora: " & Now()
    Print #LogFile, "--------------------------------------------"

    ConteggioOk = 0
    ConteggioErr = 0

    ' Ciclo per eliminare tutte le tabelle collegate
    For i = 1 To TabellaPath.Count
        Item = TabellaPath(i)
        TabellaEsiste = False

        ' Controlla se la tabella esiste effettivamente in CurrentDb
        For Each tdf In CurrentDb.TableDefs
            If tdf.Name = Item(0) Then
                TabellaEsiste = True
                Exit For
            End If
        Next tdf

        ' Se esiste, elimina e aggiorna il conteggio
        If TabellaEsiste Then
            On Error Resume Next
            EliminaTabellaSeEsiste Item(0)
            If Err.Number = 0 Then
                Print #LogFile, Now & " - OK: " & Item(0)
                ConteggioOk = ConteggioOk + 1
            Else
                Print #LogFile, Now & " - ERRORE: " & Item(0) & " -> " & Err.Description
                ConteggioErr = ConteggioErr + 1
                Err.Clear
            End If
            On Error GoTo 0
        Else
            ' Tabella nella collection ma non esiste effettivamente
            Print #LogFile, Now & " - NON TROVATA: " & Item(0)
            ConteggioErr = ConteggioErr + 1
        End If
    Next i

    ' Riepilogo log
    Print #LogFile, "--------------------------------------------"
    Print #LogFile, "TOTALE TABELLE ELIMINATE: " & ConteggioOk
    Print #LogFile, "TOTALE TABELLE NON TROVATE: " & ConteggioErr
    Close #LogFile

    ' Messaggio a video
    MsgBox "Cancellazione completata!" & vbCrLf & _
           "Tabelle eliminate: " & ConteggioOk & vbCrLf & _
           "Tabelle non trovate: " & ConteggioErr & vbCrLf & _
           "Log salvato in: " & LogFileName, vbInformation
End Function






