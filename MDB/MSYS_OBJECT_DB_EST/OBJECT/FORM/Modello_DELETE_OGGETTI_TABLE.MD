# Modello_DELETE_OGGETTI_TABLE.MD



'//CANCELLA GLI OGGETTI TABLE *** INIZIO *** '//DELETE TABLES
'//==========================================================================================================//

Private Sub Cmd_Delete_TABLE_Click()

    Dim dbCurrent As DAO.Database
    Dim tdf As DAO.TableDef
    Dim objectTables As Collection
    Dim varItem As Variant
    
'//TODO: controllare se tale routine funziona???
On Error GoTo Err_Cmd_Delete_TABLE_Click
    
    
    '//RESET
    deletedTablesCount_i = 0
    deletedTables_s = ""
   
        
              '//ATTIVO LA COLLEZIONE
              '//------------------------------------------------------------------------//
              '//Note: CHIAMO LA FUNZIONE COLLECTION TABLE per ottenere la collezione di TABELLE
                
                '//
                Set tablesCollection = CollectionTable_PFunct()
                
              '//------------------------------------------------------------------------//
        
    
    ' Apri il database corrente
    Set dbCurrent = CurrentDb
    
    ' Loop attraverso tutte le tabelle nel database corrente
    For Each tdf In dbCurrent.TableDefs
        ' Controlla se la tabella è nella collezione e passo tutta la collezione objectTables + il nome della tabella da controllare
        If InCollectionTables_Funct_b(tablesCollection, tdf.Name) Then
            ' Elimina la tabella
            On Error Resume Next
                deletedTablesCount_i = deletedTablesCount_i + 1
                DoCmd.DeleteObject acTable, tdf.Name
            On Error GoTo 0
        End If
    Next tdf
    
    
            ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
            If Me.ctr_AVVISO_ABILITATO_TXT = True Then
                MsgBox "LE TABELLE CANCELLATA DA QUESTO DATABASE SONO :  " & deletedTablesCount_i, vbExclamation, "MSG CHIAMATO LA FUNZIONE PATHFILE"
            End If


                '




'//USCITA  E GESTIONE ERRORI
'//..............................................................................................................

Exit_Cmd_Delete_TABLE_Click:
    Exit Sub

Err_Cmd_Delete_TABLE_Click:
        
    Set dbCurrent = Nothing
    
    MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & " Procedura -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Stop
    Resume Exit_Cmd_Delete_TABLE_Click

End Sub

'//*** FINE ***
'//SENZA_PARAMETRI
'//========================================================================================================================================//


'//DELETE TABLES        ***FINE ***
'//==========================================================================================================//
