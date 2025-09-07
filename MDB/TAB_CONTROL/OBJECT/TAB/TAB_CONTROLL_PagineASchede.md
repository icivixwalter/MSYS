TAB_CONTROLL_paginaeASchede.md



	SCHEMA PER IL CONTROLLO PGINE EVENTO CHANGE


            
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        '                                       TAB CONTROL
        ' @TabControllo@Evento@change_(@gestione @eventi del tab Controll, @cambio@pagina)
        '//@Trova@Progetti_(comando DOS apri trova @progetti@tutti)

        '//START "apri il progetto dei comandi di TUTTI@PROGETTI" "c:\\CASA\\LINGUAGGI\\TROVA_PROGETTI\\TROVA_PROGETTI.sublime-project"
        
        
        
        '//@APRI@PROGETTO@TAB@CONTROLL_(la path ed i comando DOS, per la routine di gestione del TAB CONTROL e le pagine che vengono aggiornate con @page@requery)
        '//START "il progetto di gestione del TAB CONTROLL con la routine di gestione delle pagine" "c:\CASA\LINGUAGGI\TROVA_PROGETTI\TROVA_PROGETTI.sublime-project"
        
        
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<
        
        
        
        
        '@CONTROLLO @PAGINE A @SCHEDE @EVENTO@CHANGE_(del controllo Tab pagine a @Schede)
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        'NOTA:
        'Utilizzo una sub routine unica per l'impostazione dei vari oggetti form.
        'I parametri da passare per ogni oggetto a cui impostare le proprieta sono i
        'seguenti:
        '1) Codice Form; 2)Tipo di oggetto, 3) Proprieta dell'oggetto - Esempio
        'sxCODOGGETTO = "PG_02_PARMIND": sxTIPOGGETTO = "Form": sxPROPRIETA = "Caption"
            
                        
            Private Sub TabCtl_01_Change()
            
                Dim ixPage As Integer                                           'variabili indice delle paginae
                Dim bFormAperta As Boolean                                      'variabile che indica se la form principale è aperta
                Dim NameSourceObject_s As String                                'variabile dell'oggetto incluso nella sottoform
        
            
                On Error GoTo Err_TabCtl_01_Change
        
        
                            'CONTROLLO SULLA FORM PRINCIPALE @controllo@pagine_(@routine per il controllo delle @pagine@a@schede)
                            '......................................................................................
                            ' La proprieta @IsLoaded Determina se la FORM  esiste e se è
                            ' attualmente aperta impostando la variabile booleana.
                            
                                                    
                                    
                                    'CONTROLLO OGGETTI DEL DABATABASE
                                    '---------------------------------------------------------------------
                                    'nota : @form@aperta_(Controllo se la form base di gestione è aperta.)
                                    
                                        
                                            Dim obj As AccessObject, dbs As Object
                                               Set dbs = Application.CurrentProject
                                               ' Search for open AccessObject objects in AllForms collection.
                                               For Each obj In dbs.AllForms
                                                   If obj.IsLoaded = True Then
                                                       ' Print name of obj la form è aperta.
                                                       Debug.Print "la form aperta IsLoaded= True " & obj.NAME
                                                       
                                                       sxFrm_CHIAMANTE_GENERALE = obj.NAME
                                                       
                                                   End If
                                               Next obj
            
                                    
                                    '---------------------------------------------------------------------
                            
                            
                                                
                                            'Imposto la variabile dello stato della form generale
                                            If sxFrm_CHIAMANTE_GENERALE = "AAGESTFrm01_N01_GESTIONI" Then
                                            
                                                    'form aperta
                                                    bFormAperta = True
                                            Else
                                                    'form chiusa
                                                    bFormAperta = False
                                            
                                            End If
                                
                            
                            '......................................................................................
        
                                    
                    
                        
                        'PAGINE
                        '____________________________________________________________________________________________________________
                            
                            
                            'RECUPERO IL NOME paginaA CORRENTE
                            '............................................................................
                                '@reset@pagina@corrente_(reset le variabili e recupero l'@indice della @pagina@corrente)
                                sxFrm_CHIAMANTE_GENERALE_CORRENTE = Me.Form.NAME
                                
                                'ixPage = Me.TabCtl_01
                                
                                'chiamo la funzione
                                'sxPagCorrente_Caption = pfunc_Recupero_NOME_paginaA_CORRENTE(sxFrm_CHIAMANTE_GENERALE_CORRENTE , 1)
                            '............................................................................
                            
                            
                            
                            '//@tab@controll@indice@pagina_(controllo dell'indice delle pagine del controllo a schede mediante @select@case@page)
                            '//Note: Vengono gestite @diverse@proprieta del Tab controll. _
                                         Per prima cosa Viene individuata la pagine selezionata con l'indice della stessa confrontando la proprieta _
                                        .value del tab controll che è un indice da 0 all'ultima pagina. _
                                                @Proprieta@value_(dell'indice tab controllo per trovare la @pagina @selezionata) _
                                        .SourceObject =  viene controlla la proprieta oggetto sorgente per vedere se la sottoform esiste. _
                                                @Proprieta@Oggetto@Sorgente_(viene controllata la @proprieta@SourceObject per controllare se la sottoform è stata inclusa)
                
                            If Me.TabCtl_01.Value = 0 Then
                           
                            
                                    
                                    'PG_01 INDICE_PAGINA = 0
                                    '............................................................................
                                    'Aggiorno la pagina 1 con il requery
                                    
                                    
                                         'REQUERY DELLA SOTTOFORM_1
                                         Me.SottoForm_01.Requery


                                        '//CONTROLLO NOME OGGETTO SORGENTE = se esiste l'oggetto incluso nella proprieta .SourceObject esso _
                                            viene stampato e aggiornato con refresh se non esiste non è possibile fare il refresh perche _
                                            da errore di oggetto chiuso o eliminato.
                                            
                                         NameSourceObject_s = Me.SottoForm_01.SourceObject '// se esiste l'oggetto incluso lo salvo nella variabile
                                         
                                        '//se il nome dell'oggetto esiste nella variabile vuol dire che è stato incluso nella sottoform _
                                        quindi puo essere sottoposto a refresh
                                        If NameSourceObject_s > "" Then
                                            Debug.Print "se esiste a sottoform stampa il nome " & Me.SottoForm_01.SourceObject
                                            Me.SottoForm_01.Form.Refresh
                                        
                                        End If
     
                                         
                                        
                                    '............................................................................
                                    
                                            
                                            
                                        
                        
                        'Pag_02
                        '____________________________________________________________________________________________________________
                            ElseIf Me.TabCtl_01.Value = 1 Then
                        
                        
                                         
                                    
                                    'PG_02 INDICE_PAGINA = 1
                                    '............................................................................
                                    'Aggiorno la pagina 1 con il requery
                                    
                                    
                                        'REQUERY DELLA SOTTOFORM_2
                                            Me.SottoForm_02.Requery

                                        '//CONTROLLO NOME OGGETTO SORGENTE = se esiste l'oggetto incluso nella proprieta .SourceObject esso _
                                            viene stampato e aggiornato con refresh se non esiste non è possibile fare il refresh perche _
                                            da errore di oggetto chiuso o eliminato.
                                            
                                         NameSourceObject_s = Me.SottoForm_02.SourceObject '// se esiste l'oggetto incluso lo salvo nella variabile
                                         
                                        '//se il nome dell'oggetto esiste nella variabile vuol dire che è stato incluso nella sottoform _
                                        quindi puo essere sottoposto a refresh
                                        If NameSourceObject_s > "" Then
                                            Debug.Print "se esiste a sottoform stampa il nome " & Me.SottoForm_02.SourceObject
                                            Me.SottoForm_02.Form.Refresh
                                        
                                        End If
     


                                    '............................................................................
                                    
                                    
                                    
                        'Pag_03
                        '____________________________________________________________________________________________________________
                            ElseIf Me.TabCtl_01.Value = 2 Then
                        
                                    
                                    
                                    'PG_03 INDICE_PAGINA = 2
                                    '............................................................................
                                    'Aggiorno la pagina 3 con il requery
                                    
                                    
                                         'REQUERY DELLA SOTTOFORM_3
                                         Me.SottoForm_03.Requery

                                        '//CONTROLLO NOME OGGETTO SORGENTE = se esiste l'oggetto incluso nella proprieta .SourceObject esso _
                                            viene stampato e aggiornato con refresh se non esiste non è possibile fare il refresh perche _
                                            da errore di oggetto chiuso o eliminato.
                                            
                                         NameSourceObject_s = Me.SottoForm_03.SourceObject '// se esiste l'oggetto incluso lo salvo nella variabile
                                         
                                        '//se il nome dell'oggetto esiste nella variabile vuol dire che è stato incluso nella sottoform _
                                        quindi puo essere sottoposto a refresh
                                        If NameSourceObject_s > "" Then
                                            Debug.Print "se esiste a sottoform stampa il nome " & Me.SottoForm_03.SourceObject
                                            Me.SottoForm_03.Form.Refresh
                                        
                                        End If


                                        
                                    '............................................................................
                                        
                                    
                        'Pag_04
                        '____________________________________________________________________________________________________________
                            ElseIf Me.TabCtl_01.Value = 3 Then
                        
                                    
                                    'PG_04 INDICE_PAGINA = 3
                                    '............................................................................
                                    'Aggiorno la pagina 4 con il requery
                                    
                                    
                                         'REQUERY DELLA SOTTOFORM_4
                                         Me.SottoForm_04.Requery

                                        '//CONTROLLO NOME OGGETTO SORGENTE = se esiste l'oggetto incluso nella proprieta .SourceObject esso _
                                            viene stampato e aggiornato con refresh se non esiste non è possibile fare il refresh perche _
                                            da errore di oggetto chiuso o eliminato.
                                            
                                         NameSourceObject_s = Me.SottoForm_04.SourceObject '// se esiste l'oggetto incluso lo salvo nella variabile
                                         
                                        '//se il nome dell'oggetto esiste nella variabile vuol dire che è stato incluso nella sottoform _
                                        quindi puo essere sottoposto a refresh
                                        If NameSourceObject_s > "" Then
                                            Debug.Print "se esiste a sottoform stampa il nome " & Me.SottoForm_04.SourceObject
                                            Me.SottoForm_04.Form.Refresh
                                        
                                        End If

                                        
                                    '............................................................................
                                
                        '____________________________________________________________________________________________________________
                    
                                     
                        'Pag_05
                        '____________________________________________________________________________________________________________
                            ElseIf Me.TabCtl_01.Value = 4 Then
                        
                                
                                    
                                    'PG_04 INDICE_PAGINA = 4
                                    '............................................................................
                                    'Aggiorno la pagina 5 con il requery
                                    
                                    
                                         'REQUERY DELLA SOTTOFORM_5
                                         Me.SottoForm_05.Requery


                                        '//CONTROLLO NOME OGGETTO SORGENTE = se esiste l'oggetto incluso nella proprieta .SourceObject esso _
                                            viene stampato e aggiornato con refresh se non esiste non è possibile fare il refresh perche _
                                            da errore di oggetto chiuso o eliminato.
                                            
                                         NameSourceObject_s = Me.SottoForm_05.SourceObject '// se esiste l'oggetto incluso lo salvo nella variabile
                                         
                                        '//se il nome dell'oggetto esiste nella variabile vuol dire che è stato incluso nella sottoform _
                                        quindi puo essere sottoposto a refresh
                                        If NameSourceObject_s > "" Then
                                            Debug.Print "se esiste a sottoform stampa il nome " & Me.SottoForm_05.SourceObject
                                            Me.SottoForm_05.Form.Refresh
                                        
                                        End If

                                        
                                    '............................................................................                                
                        '____________________________________________________________________________________________________________
                    
                       
                    
                        'Pag_06
                        '____________________________________________________________________________________________________________
                            ElseIf Me.TabCtl_01.Value = 5 Then
                        
            
                                    'PG_06 INDICE_PAGINA = 5
                                    '............................................................................
                                    'Aggiorno la pagina 6 con il requery
                                    
                                    
                                         'REQUERY DELLA SOTTOFORM_6
                                         Me.SottoForm_06.Requery


                                        '//CONTROLLO NOME OGGETTO SORGENTE = se esiste l'oggetto incluso nella proprieta .SourceObject esso _
                                            viene stampato e aggiornato con refresh se non esiste non è possibile fare il refresh perche _
                                            da errore di oggetto chiuso o eliminato.
                                            
                                         NameSourceObject_s = Me.SottoForm_06.SourceObject '// se esiste l'oggetto incluso lo salvo nella variabile
                                         
                                        '//se il nome dell'oggetto esiste nella variabile vuol dire che è stato incluso nella sottoform _
                                        quindi puo essere sottoposto a refresh
                                        If NameSourceObject_s > "" Then
                                            Debug.Print "se esiste a sottoform stampa il nome " & Me.SottoForm_06.SourceObject
                                            Me.SottoForm_06.Form.Refresh
                                        
                                        End If

                                        
                                    '............................................................................

                        '____________________________________________________________________________________________________________
                    
                    
                        'Pag_07
                        '____________________________________________________________________________________________________________
                            ElseIf Me.TabCtl_01.Value = 6 Then
                        
                                    'PG_07 INDICE_PAGINA = 6
                                    '............................................................................
                                    'Aggiorno la pagina 7 con il requery
                                    
                                    
                                         'REQUERY DELLA SOTTOFORM_7
                                         Me.SottoForm_07.Requery

                                        '//CONTROLLO NOME OGGETTO SORGENTE = se esiste l'oggetto incluso nella proprieta .SourceObject esso _
                                            viene stampato e aggiornato con refresh se non esiste non è possibile fare il refresh perche _
                                            da errore di oggetto chiuso o eliminato.
                                            
                                         NameSourceObject_s = Me.SottoForm_07.SourceObject '// se esiste l'oggetto incluso lo salvo nella variabile
                                         
                                        '//se il nome dell'oggetto esiste nella variabile vuol dire che è stato incluso nella sottoform _
                                        quindi puo essere sottoposto a refresh
                                        If NameSourceObject_s > "" Then
                                            Debug.Print "se esiste a sottoform stampa il nome " & Me.SottoForm_07.SourceObject
                                            Me.SottoForm_07.Form.Refresh
                                        
                                        End If


                                    '............................................................................
                                    
                        '____________________________________________________________________________________________________________
                    
                    
                        'Pag_08
                        '____________________________________________________________________________________________________________
                            ElseIf Me.TabCtl_01.Value = 7 Then
                        
                                    
                                    'PG_08 INDICE_PAGINA = 7
                                    '............................................................................
                                    'Aggiorno la pagina 8 con il requery
                                    
                                    
                                         'REQUERY DELLA SOTTOFORM_8
                                         Me.SottoForm_08.Requery

                                        '//CONTROLLO NOME OGGETTO SORGENTE = se esiste l'oggetto incluso nella proprieta .SourceObject esso _
                                            viene stampato e aggiornato con refresh se non esiste non è possibile fare il refresh perche _
                                            da errore di oggetto chiuso o eliminato.
                                            
                                         NameSourceObject_s = Me.SottoForm_08.SourceObject '// se esiste l'oggetto incluso lo salvo nella variabile
                                         
                                        '//se il nome dell'oggetto esiste nella variabile vuol dire che è stato incluso nella sottoform _
                                        quindi puo essere sottoposto a refresh
                                        If NameSourceObject_s > "" Then
                                            Debug.Print "se esiste a sottoform stampa il nome " & Me.SottoForm_08.SourceObject
                                            Me.SottoForm_08.Form.Refresh
                                        
                                        End If


                                    '............................................................................
                        '____________________________________________________________________________________________________________
                    
                                     
                        'Pag_09
                        '____________________________________________________________________________________________________________
                            ElseIf Me.TabCtl_01.Value = 8 Then
                        
                                    'PG_09 INDICE_PAGINA = 8
                                    '............................................................................
                                    'Aggiorno la pagina 9 con il requery
                                    
                                    
                                         'REQUERY DELLA SOTTOFORM_9
                                         Me.SottoForm_09.Requery

                                        '//CONTROLLO NOME OGGETTO SORGENTE = se esiste l'oggetto incluso nella proprieta .SourceObject esso _
                                            viene stampato e aggiornato con refresh se non esiste non è possibile fare il refresh perche _
                                            da errore di oggetto chiuso o eliminato.
                                            
                                         NameSourceObject_s = Me.SottoForm_09.SourceObject '// se esiste l'oggetto incluso lo salvo nella variabile
                                         
                                        '//se il nome dell'oggetto esiste nella variabile vuol dire che è stato incluso nella sottoform _
                                        quindi puo essere sottoposto a refresh
                                        If NameSourceObject_s > "" Then
                                            Debug.Print "se esiste a sottoform stampa il nome " & Me.SottoForm_09.SourceObject
                                            Me.SottoForm_09.Form.Refresh
                                        
                                        End If


                                    '............................................................................
                                
                        '____________________________________________________________________________________________________________
                    
                    
                                     
                        'Pag_10
                        '____________________________________________________________________________________________________________
                            ElseIf Me.TabCtl_01.Value = 9 Then
                        
                                    'PG_10 INDICE_PAGINA = 9
                                    '............................................................................
                                    'Aggiorno la pagina 10 con il requery
                                    
                                    
                                         'REQUERY DELLA SOTTOFORM_10
                                         Me.SottoForm_10.Requery

                                        '//CONTROLLO NOME OGGETTO SORGENTE = se esiste l'oggetto incluso nella proprieta .SourceObject esso _
                                            viene stampato e aggiornato con refresh se non esiste non è possibile fare il refresh perche _
                                            da errore di oggetto chiuso o eliminato.
                                            
                                         NameSourceObject_s = Me.SottoForm_10.SourceObject '// se esiste l'oggetto incluso lo salvo nella variabile
                                         
                                        '//se il nome dell'oggetto esiste nella variabile vuol dire che è stato incluso nella sottoform _
                                        quindi puo essere sottoposto a refresh
                                        If NameSourceObject_s > "" Then
                                            Debug.Print "se esiste a sottoform stampa il nome " & Me.SottoForm_10.SourceObject
                                            Me.SottoForm_10.Form.Refresh
                                        
                                        End If


                                    '............................................................................
                                
                        '____________________________________________________________________________________________________________
                    
                            
                                     
                        'Pag_11
                        '____________________________________________________________________________________________________________
                            ElseIf Me.TabCtl_01.Value = 10 Then
                        
                                    'PG_11 INDICE_PAGINA = 10
                                    '............................................................................
                                    'Aggiorno la pagina 11 con il requery
                                    
                                    
                                         'REQUERY DELLA SOTTOFORM_11
                                         Me.SottoForm_11.Requery

                                        '//CONTROLLO NOME OGGETTO SORGENTE = se esiste l'oggetto incluso nella proprieta .SourceObject esso _
                                        viene stampato e aggiornato con refresh se non esiste non è possibile fare il refresh perche _
                                        da errore di oggetto chiuso o eliminato.
                                            
                                         NameSourceObject_s = Me.SottoForm_11.SourceObject '// se esiste l'oggetto incluso lo salvo nella variabile
                                         
                                        '//se il nome dell'oggetto esiste nella variabile vuol dire che è stato incluso nella sottoform _
                                        quindi puo essere sottoposto a refresh
                                        If NameSourceObject_s > "" Then
                                            Debug.Print "se esiste a sottoform stampa il nome " & Me.SottoForm_11.SourceObject
                                            Me.SottoForm_11.Form.Refresh
                                        
                                        End If

                                        
                                    '............................................................................
                    
                        '____________________________________________________________________________________________________________
                    

                        'Pag_12
                        '____________________________________________________________________________________________________________
                            ElseIf Me.TabCtl_01.Value = 11 Then
                        
                                    'PG_12 INDICE_PAGINA = 11
                                    '............................................................................
                                    'Aggiorno la pagina 12 con il requery
                                    
                                    
                                         'REQUERY DELLA SOTTOFORM_12
                                         Me.SottoForm_12.Requery

                                        '//CONTROLLO NOME OGGETTO SORGENTE = se esiste l'oggetto incluso nella proprieta .SourceObject esso _
                                        viene stampato e aggiornato con refresh se non esiste non è possibile fare il refresh perche _
                                        da errore di oggetto chiuso o eliminato.
                                            
                                         NameSourceObject_s = Me.SottoForm_12.SourceObject '// se esiste l'oggetto incluso lo salvo nella variabile
                                         
                                        '//se il nome dell'oggetto esiste nella variabile vuol dire che è stato incluso nella sottoform _
                                        quindi puo essere sottoposto a refresh
                                        If NameSourceObject_s > "" Then
                                            Debug.Print "se esiste a sottoform stampa il nome " & Me.SottoForm_12.SourceObject
                                            Me.SottoForm_12.Form.Refresh
                                        
                                        End If

                                        
                                    '............................................................................
                    
                        '____________________________________________________________________________________________________________
                    


                            
                        End If
                        
                        
                        
                        
                                    
                    
        
        'USCITA  E GESTIONE ERRORI
        '..............................................................................................................
        
        Exit_TabCtl_01_Change:
            Exit Sub
        
        Err_TabCtl_01_Change:
            MsgBox Err.Description
            
            Resume Exit_TabCtl_01_Change
        
        
        
        End Sub
        
        '*** FINE ***
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        '                                       TAB CONTROL
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<
        
