COLLEGA_Mdl01_PRES3000_DF00_TUTTE_LE_DEFINIZIONI.MD
    Note
        Modulo di collegamento di tutte le definizioni delle timbrature 

    schema

        Option Compare Database
            Dim BloccoProcedura_b As Boolean


            '------------------------------------------------------------
            ' moduli @per il @collegamento delle tabelle di definizione  
            '               COLLEGA_Mdl01_PRES3000_DF00_}----------------------------------@                            
            '   NOME    :   COLLEGA_PRES3000_TUTTE_LE_DEFINIZIONI_Pf.CLS
            '   
            '           N.b = per funzionare devi creare anche questa macro
            '                   COLLEGA_Mcr_PRES3000_DF00_EseguiCodice_COLLEGA_TIMBRATURE_TUTTE
            '
            '           Nella macro inserisci questo codice come nome della funzion:
            '                   COLLEGA_PRES3000_TUTTE_LE_DEFINIZIONI_Pf()
            '
            '           
            '
            '
            '  @muodulo@collega@timbrature_(@collega@tutto,  @importa il @collegamento di @tutte le tabelle di @definizione @delle @timbrature)
            '  il modulo di BASE si trova in questo progetto: 
            '  START "modulo" call "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\PRES3000_GESTIONE\ARCHIVI_PRES3000\PRES3000_N00_GESTIONE\GESTIONE_MDB\PRES3000_N00_GESTIONE.mdb"
            '------------------------------------------------------------
            Public Function COLLEGA_PRES3000_TUTTE_LE_DEFINIZIONI_Pf()

            Dim ProjectMdb_s As String
            Dim Path_ProjectMdb_s As String


            Dim Database_PATH_s As String
            Dim Database_s As String
            Dim Database_TABLE_01_s As String
            Dim Database_TABLE_02_s As String
            Dim Database_TABLE_03_s As String
            Dim Database_TABLE_04_s As String
            Dim Database_TABLE_05_s As String
            Dim Database_TABLE_06_s As String
            Dim Database_TABLE_07_s As String
            Dim Database_TABLE_08_s As String
            Dim Database_TABLE_09_s As String
            Dim Database_TABLE_10_s As String
            Dim Database_TABLE_11_s As String
            Dim Database_TABLE_12_s As String
            Dim Database_TABLE_13_s As String
            Dim Database_TABLE_14_s As String
            Dim Database_TABLE_15_s As String
            Dim Database_TABLE_16_s As String
            Dim Database_TABLE_17_s As String
            Dim Database_TABLE_18_s As String

            Dim icount as Integer                   '//contatore delle tabelle collegate'




            On Error GoTo COLLEGA_PRES3000_TUTTE_LE_DEFINIZIONI_Pf_Err


            '//reset
            ProjectMdb_s = "PRES3000_N05_DEFINIZIONI_ORAR_GIORN_TIMB"
            Path_ProjectMdb_s = "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\PRES3000_GESTIONE\ARCHIVI_PRES3000\"

                

                BloccoProcedura_b = False
                If BloccoProcedura_b = True Then
                    MsgBox "BLOCCO PROCEDURA DI COLLEGAMENTO LLPP_DF05_FALDONI per evitare la cancellazione della tabella dati! " & _
                        "Impostare a False la variabile BloccoProcedura_b", vbExclamation, "MSG BLOCCO COLLEGAMENTO"
                    GoTo COLLEGA_PRES3000_TUTTE_LE_DEFINIZIONI_Pf_Exit
                End If

                        '//MESSAGGIO INIZIO COLLEGAMENTO
                               
                            MsgBox "IMPORTO COLLEGAMENTI", vbExclamation, "COLLEGO TUTTE LE TABELLE DI DEFINIZIONE DELLE TIMBRATURE ---> " & Path_ProjectMdb_s & "\" & ProjectMdb_sProjectMdb_s
                            
                
                ' In caso di errore non viene eseguita l'istruzione elimina oggetto ma quella successiva
                '//------------------------------------------------------------------------------------------------------//
                    '//parametri per l'importazione
                    Database_PATH_s = "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\PRES3000_GESTIONE\ARCHIVI_PRES3000\PRES3000_N05_DEFINIZIONI_ORAR_GIORN_TIMB\"
                    Database_s = "PRES3000_N05_DEFINIZIONI_ORAR_GIORN_TIMB.mdb"
                    Database_TABLE_s = "PRES3000_DF00_{@==========================================@}_Tab"
                   
                
                '//modello
                '//On Error Resume Next
                ' viene prima cancellata con
                ' EliminaOggetto
                ' e poi riscotruito il collegamento con
                ' TrasferisciDatabase
                ' e cioè RICOSTUIRISCE TRUISCE IL COLELGAMENTO trasferisce database. Se invece esiste la tabella  ATTI DI IMPEGNO La stessa
                ' prima elimina la tabella collegata gia esistente
                '//DoCmd.DeleteObject acTable, Database_TABLE_s
                ' LLPP_DF05_FALDONI
                ' path della tabella  ------------> :c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\
                ' database -----------------------> LLPP_DF05_FALDONI.mdb
                ' collego la tabella LLPP_IMPEGNI_Tb02_ELENCO_DI_SPESA
                '//DoCmd.TransferDatabase acLink, "Microsoft Access", Database_PATH_s & Database_s, acTable, Database_TABLE_s, Database_TABLE_s, False

                        '//TABELLE DA COLLEGARE'

                                Database_TABLE_01_s="PRES3000_DF05_}------------------------------------------------@"
                                Database_TABLE_02_s="PRES3000_DF05_ORARI"
                                Database_TABLE_03_s="PRES3000_DF10_}------------------------------------------------@"
                                Database_TABLE_04_s="PRES3000_DF10_GIORNALIERE"
                                Database_TABLE_05_s="PRES3000_DF11_}------------------------------------------------@"
                                Database_TABLE_06_s="PRES3000_DF11_TIMBRATURE"
                                Database_TABLE_07_s="PRES3000_DF12_}------------------------------------------------@"
                                Database_TABLE_08_s="PRES3000_DF12_ORARI_TIMBRATURE"
                                Database_TABLE_09_s="PRES3000_DF19_}------------------------------------------------@"
                                Database_TABLE_10_s="PRES3000_DF19_GiorniCalendario"
                                Database_TABLE_12_s="PRES3000_DF20_}------------------------------------------------@"
                                Database_TABLE_13_s="PRES3000_DF20_Help_Calendario"
                                Database_TABLE_14_s="PRES3000_DF00_{@==========================================@}_Tab"
                                Database_TABLE_15_s="PRES3000_DF01_}------------------------------------------------@"
                                Database_TABLE_16_s="PRES3000_DF01_CalendarioFestivita"
                                Database_TABLE_17_s="PRES3000_DF22_}------------------------------------------------@"
                                Database_TABLE_18_s="PRES3000_DF22_Help_Giornaliere"
                              
                                '//accoda solo le tabelle
                                If Database_TABLE_01_s > "" Then
                                    On Error Resume Next        '//se non ci sono errori cancella
                                    DoCmd.DeleteObject acTable, Database_TABLE_01_s
                                    DoCmd.TransferDatabase acLink, "Microsoft Access", Database_PATH_s & Database_s, acTable, Database_TABLE_01_s, Database_TABLE_01_s, False
                                    icount=icount+1
                                End If
                               

                                '//accoda solo le tabelle
                                If Database_TABLE_02_s > "" Then
                                    On Error Resume Next        '//se non ci sono errori cancella
                                    DoCmd.DeleteObject acTable, Database_TABLE_02_s
                                    DoCmd.TransferDatabase acLink, "Microsoft Access", Database_PATH_s & Database_s, acTable, Database_TABLE_02_s, Database_TABLE_02_s, False
                                    icount=icount+1
                                End If



                                '//accoda solo le tabelle
                                If Database_TABLE_03_s > "" Then
                                    On Error Resume Next        '//se non ci sono errori cancella
                                    DoCmd.DeleteObject acTable, Database_TABLE_03_s
                                    DoCmd.TransferDatabase acLink, "Microsoft Access", Database_PATH_s & Database_s, acTable, Database_TABLE_03_s, Database_TABLE_03_s, False
                                    icount=icount+1
                                End If



                                '//accoda solo le tabelle
                                If Database_TABLE_04_s > "" Then
                                    On Error Resume Next        '//se non ci sono errori cancella
                                    DoCmd.DeleteObject acTable, Database_TABLE_04_s
                                    DoCmd.TransferDatabase acLink, "Microsoft Access", Database_PATH_s & Database_s, acTable, Database_TABLE_04_s, Database_TABLE_04_s, False
                                    icount=icount+1
                                End If


                                '//accoda solo le tabelle
                                If Database_TABLE_05_s > "" Then
                                    On Error Resume Next        '//se non ci sono errori cancella
                                    DoCmd.DeleteObject acTable, Database_TABLE_05_s
                                    DoCmd.TransferDatabase acLink, "Microsoft Access", Database_PATH_s & Database_s, acTable, Database_TABLE_05_s, Database_TABLE_05_s, False
                                    icount=icount+1
                                End If




                                '//accoda solo le tabelle
                                If Database_TABLE_06_s > "" Then
                                    On Error Resume Next        '//se non ci sono errori cancella
                                    DoCmd.DeleteObject acTable, Database_TABLE_06_s
                                    DoCmd.TransferDatabase acLink, "Microsoft Access", Database_PATH_s & Database_s, acTable, Database_TABLE_06_s, Database_TABLE_06_s, False
                                    icount=icount+1
                                End If




                                '//accoda solo le tabelle
                                If Database_TABLE_07_s > "" Then
                                    On Error Resume Next        '//se non ci sono errori cancella
                                    DoCmd.DeleteObject acTable, Database_TABLE_07_s
                                    DoCmd.TransferDatabase acLink, "Microsoft Access", Database_PATH_s & Database_s, acTable, Database_TABLE_07_s, Database_TABLE_07_s, False
                                    icount=icount+1
                                End If



                                '//accoda solo le tabelle
                                If Database_TABLE_08_s > "" Then
                                    On Error Resume Next        '//se non ci sono errori cancella
                                    DoCmd.DeleteObject acTable, Database_TABLE_08_s
                                    DoCmd.TransferDatabase acLink, "Microsoft Access", Database_PATH_s & Database_s, acTable, Database_TABLE_08_s, Database_TABLE_08_s, False
                                    icount=icount+1
                                End If



                                '//accoda solo le tabelle
                                If Database_TABLE_09_s > "" Then
                                    On Error Resume Next        '//se non ci sono errori cancella
                                    DoCmd.DeleteObject acTable, Database_TABLE_09_s
                                    DoCmd.TransferDatabase acLink, "Microsoft Access", Database_PATH_s & Database_s, acTable, Database_TABLE_09_s, Database_TABLE_09_s, False
                                    icount=icount+1
                                End If




                                '//accoda solo le tabelle
                                If Database_TABLE_10_s > "" Then
                                    On Error Resume Next        '//se non ci sono errori cancella
                                    DoCmd.DeleteObject acTable, Database_TABLE_10_s
                                    DoCmd.TransferDatabase acLink, "Microsoft Access", Database_PATH_s & Database_s, acTable, Database_TABLE_10_s, Database_TABLE_10_s, False
                                    icount=icount+1
                                End If




                                '//accoda solo le tabelle
                                If Database_TABLE_11_s > "" Then
                                    On Error Resume Next        '//se non ci sono errori cancella
                                    DoCmd.DeleteObject acTable, Database_TABLE_11_s
                                    DoCmd.TransferDatabase acLink, "Microsoft Access", Database_PATH_s & Database_s, acTable, Database_TABLE_11_s, Database_TABLE_11_s, False
                                    icount=icount+1
                                End If



                                '//accoda solo le tabelle
                                If Database_TABLE_12_s > "" Then
                                    On Error Resume Next        '//se non ci sono errori cancella
                                    DoCmd.DeleteObject acTable, Database_TABLE_12_s
                                    DoCmd.TransferDatabase acLink, "Microsoft Access", Database_PATH_s & Database_s, acTable, Database_TABLE_12_s, Database_TABLE_12_s, False
                                    icount=icount+1
                                End If


                                '//accoda solo le tabelle
                                If Database_TABLE_13_s > "" Then
                                    On Error Resume Next        '//se non ci sono errori cancella
                                    DoCmd.DeleteObject acTable, Database_TABLE_13_s
                                    DoCmd.TransferDatabase acLink, "Microsoft Access", Database_PATH_s & Database_s, acTable, Database_TABLE_13_s, Database_TABLE_13_s, False
                                    icount=icount+1
                                End If



                                '//accoda solo le tabelle
                                If Database_TABLE_14_s > "" Then
                                    On Error Resume Next        '//se non ci sono errori cancella
                                    DoCmd.DeleteObject acTable, Database_TABLE_14_s
                                    DoCmd.TransferDatabase acLink, "Microsoft Access", Database_PATH_s & Database_s, acTable, Database_TABLE_14_s, Database_TABLE_14_s, False
                                    icount=icount+1
                                End If



                                '//accoda solo le tabelle
                                If Database_TABLE_15_s > "" Then
                                    On Error Resume Next        '//se non ci sono errori cancella
                                    DoCmd.DeleteObject acTable, Database_TABLE_15_s
                                    DoCmd.TransferDatabase acLink, "Microsoft Access", Database_PATH_s & Database_s, acTable, Database_TABLE_15_s, Database_TABLE_15_s, False
                                    icount=icount+1
                                End If


                                '//accoda solo le tabelle
                                If Database_TABLE_16_s > "" Then
                                    On Error Resume Next        '//se non ci sono errori cancella
                                    DoCmd.DeleteObject acTable, Database_TABLE_16_s
                                    DoCmd.TransferDatabase acLink, "Microsoft Access", Database_PATH_s & Database_s, acTable, Database_TABLE_16_s, Database_TABLE_16_s, False
                                    icount=icount+1
                                End If



                                '//accoda solo le tabelle
                                If Database_TABLE_17_s > "" Then
                                    On Error Resume Next        '//se non ci sono errori cancella
                                    DoCmd.DeleteObject acTable, Database_TABLE_17_s
                                    DoCmd.TransferDatabase acLink, "Microsoft Access", Database_PATH_s & Database_s, acTable, Database_TABLE_17_s, Database_TABLE_17_s, False
                                    icount=icount+1
                                End If


                                '//accoda solo le tabelle
                                If Database_TABLE_18_s > "" Then
                                    On Error Resume Next        '//se non ci sono errori cancella
                                    DoCmd.DeleteObject acTable, Database_TABLE_18_s
                                    DoCmd.TransferDatabase acLink, "Microsoft Access", Database_PATH_s & Database_s, acTable, Database_TABLE_18_s, Database_TABLE_18_s, False
                                    icount=icount+1
                                End If




                                '//vado alla prima tabella inserita'
                                    ' tabella PRES3000_Tb01_Calendario
                                    DoCmd.SelectObject acTable, "PRES3000_Tb01_Calendario", True


                
                '//------------------------------------------------------------------------------------------------------//
                
                
                MsgBox "IMPORTAZIONE CONCLUSA tabelle collegate : "&icount, vbExclamation, "TITOLO FINALE " 
                
                

            COLLEGA_PRES3000_TUTTE_LE_DEFINIZIONI_Pf_Exit:
                Exit Function

            COLLEGA_PRES3000_TUTTE_LE_DEFINIZIONI_Pf_Err:
                MsgBox Error$
                Resume COLLEGA_PRES3000_TUTTE_LE_DEFINIZIONI_Pf_Exit

            End Function











