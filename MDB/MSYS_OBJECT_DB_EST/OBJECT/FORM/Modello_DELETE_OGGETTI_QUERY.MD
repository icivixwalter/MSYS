# Modello_DELETE_OGGETTI_QUERY.MD







'//CANCELLA GLI OGGETTI QUERY *** INIZIO *** '//DELETE QUERIES
'//==========================================================================================================//

Private Sub Cmd_Delete_QUERY_Click()

    Dim dbCurrent As DAO.Database
    Dim qry As DAO.QueryDef
    Dim objectTables As Collection
    Dim varItem As Variant
    Dim MATRICE_queryNames_s() As String             'matrice di Query
    Dim tempCount As Integer
    Dim queryCount As Integer
    Dim qdf  As QueryDef
    

    
'//TODO: controllare se tale routine funziona???
On Error GoTo Err_Cmd_Delete_QUERY_Click
    
    
    '//RESET
    icount = 0
    deleteQueryesCount_i = 0
    deletedTablesCount_i = 0
    deletedTables_s = ""
   
        
              '//ATTIVO LA COLLEZIONE
              '//------------------------------------------------------------------------//
              '//Note: CHIAMO LA FUNZIONE COLLECTION QUERY per ottenere la collezione di QUERY
                
                
                
                '//
                Set queryesCollection = CollectionQuery_PFunct()
                
              '//------------------------------------------------------------------------//
                    
                    ' Loop attraverso tutte I REPORT nel database corrente
                     Set dbCurrent = CurrentDb
                     
                     '//CALCOLO INDICE QUERY
                     icount = dbCurrent.QueryDefs.Count
                   
                    Set dbCurrent = CurrentDb

                        
                    '//se ci sono query da cancellare
                    If icount > 0 Then
                    
                        '//reimposto la matrice sulle query esistenti nel db corrente
                        ReDim MATRICE_queryNames_s(icount - 1)
                        
                        ' Costruisci un array con i nomi dei Query
                        tempCount = 0
                        
                        '//CARICO LA MATRICE: ciclo for per caricare MATRICE_queryNames_s(tempCount)
                        For Each qdf In dbCurrent.QueryDefs
                            MATRICE_queryNames_s(tempCount) = qdf.Name
                            tempCount = tempCount + 1
                        Next qdf
                        
                        ' Cancella i Query usando l'array dei nomi
                        For i = 0 To tempCount - 1
                            On Error Resume Next  ' Aggiungi gestione errori specifica per ogni iterazione
                            Debug.Print "Cancellazione Query: " & MATRICE_queryNames_s(i)
                            
                            '//CONTROLLO MATRICE = COLLECTION
                            Bool1 = InCollectionQueryes_Funct_b(queryesCollection, MATRICE_queryNames_s(i))
                            
                            '//cancello solo se la matrice (i) esiste nella collection
                            If Bool1 = True Then
                                    Debug.Print
                                    Debug.Print "cancello la query trovata sia nella collezione che nella matrice : " & MATRICE_queryNames_s(i)
                                    DoCmd.DeleteObject acQuery, MATRICE_queryNames_s(i)
                                    
                                    If Err.Number <> 0 Then
                                        Debug.Print "Errore durante la cancellazione del Query: " & MATRICE_queryNames_s(i) & " - " & Err.Description
                                        Err.Clear  ' Pulisci l'errore per la prossima iterazione
                                    Else
                                        deleteQueryesCount_i = deleteQueryesCount_i + 1
                                    End If
                                    
                            End If '//If Bool1 = True Then
                            
                            On Error GoTo 0  ' Reimposta gestione errori
                        Next i
                    End If
    
            ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
            If Me.ctr_AVVISO_ABILITATO_TXT = True Then
                MsgBox "LE QUERY CANCELLATA DA QUESTO DATABASE SONO :  " & deleteQueryesCount_i, vbExclamation, "MSG CHIAMATO LA FUNZIONE PATHFILE"
            End If





'//USCITA  E GESTIONE ERRORI
'//..............................................................................................................

Exit_Cmd_Delete_QUERY_Click:
    Exit Sub

Err_Cmd_Delete_QUERY_Click:
        
    Set dbCurrent = Nothing
    
    MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & " Procedura -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Stop
    Resume Exit_Cmd_Delete_QUERY_Click

End Sub

'//*** FINE ***
'//SENZA_PARAMETRI
'//========================================================================================================================================//



' Funzione per verificare se un elemento esiste nella collezione
Function InCollectionQueryes_Funct_b(col As Collection, key As String) As Boolean
    Dim varItem As Variant
    On Error Resume Next
    '//ITERNO NELLA COLLECTION PASSATA COME PARAMETRO
    For Each varItem In col
        
        Debug.Print
        Debug.Print "CONTROLLO COLLEZIONE = MATRICE (I)"
        Debug.Print "collezione: " & varItem & " = " & "matrice key: " & key
        Debug.Print
        If varItem = key Then
            InCollectionQueryes_Funct_b = True
            Debug.Print "trovato restituisco : " & InCollectionQueryes_Funct_b
            Exit Function
        End If
    Next varItem
    InCollectionQueryes_Funct_b = False
    On Error GoTo 0
End Function



'//DELETE QUERIES        ***FINE ***
'//==========================================================================================================//
