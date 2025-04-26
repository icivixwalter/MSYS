STAMPA_OGGETTI_DB_CORRENTE.MD

   Note
      @stampa@oggetti@db_(stampa gli oggetti presenti nel database)


STAMPA_QUERY 

   '//STAMPA_TUTTI_GLI_OGGETTI_QUERY
   '//--------------------------------------------------------------------------------------------//
   '//note: nel db corrente stampa gli oggetti query


   Function Stampa_Oggetti_Query_Pfunct()
       On Error GoTo Stampa_Oggetti_Query_Pfunct_Err
       
       Dim db As DAO.Database
       Dim qryDef As DAO.QueryDef
       
       ' Apri il database corrente
       Set db = CurrentDb
       
       ' Scorri tutte le query nel database
       For Each qryDef In db.QueryDefs
           Debug.Print qryDef.Name ' Stampa il nome della query
       Next qryDef
       
       ' Pulisce la memoria
       Set qryDef = Nothing
       Set db = Nothing
       
   Stampa_Oggetti_Query_Pfunct_Exit:
       Exit Function
       
   Stampa_Oggetti_Query_Pfunct_Err:
       MsgBox "Errore: " & Err.Description
       Resume Stampa_Oggetti_Query_Pfunct_Exit
   End Function




STAMPA_TABELLE_COLLEGATE

   Sub StampaTabelleCollegate()
       Dim db As DAO.Database
       Dim tdf As DAO.TableDef
       
       Set db = CurrentDb
       
       Debug.Print "Tabelle Collegate:"
       
       For Each tdf In db.TableDefs
           If Len(tdf.Connect) > 0 Then ' Controlla se è una tabella collegata
               Debug.Print vbCrLf
               Debug.Print tdf.Name & "                --> " & tdf.Connect & vbCrLf;
           End If
       Next
       
       Set db = Nothing
   End Sub


    
