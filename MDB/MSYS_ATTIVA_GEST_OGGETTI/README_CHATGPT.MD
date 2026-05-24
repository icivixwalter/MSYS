README_CHATGPT.MD

01=CONTROLLO SE LA FORM MASTER AA_MSys_FORM_M01_GESTIONE_OGGETTI_MASTER E' APERTA
   
   con questo codice di controllo del codice progetto si deve aggiungre un controll per evitare
   errori che verifica se la FORM PRINCIPALE MASTER E' APERTA, definita nel titolo;
   dove puoi inserirla nel seguente codice? se la MASTER NON E' APERTA il resto delle attivita
   deve essere inibito con un messaggio "FORM MASTER : AA_MSys_FORM_M01_GESTIONE_OGGETTI_MASTER USCITA DALLA PROCEDURA"
   integra il seguente codice:
   
   '// COD_PROGETTO_s_TXT_DblClick, MenuGestito_s, CodiceMenu_s    *** inizio **
   '//========================================================================================================//
   '//codice ----> COD_PROGETTO_s_TXT_DblClick_(ATTIVO IL PROGEDTTO MASTER PER IL CARICAMENTO DELLE SOTTOFORM)
   '//
   '//ATTIVITA ESEGUITE:
   '// ========================================================================================================//
   '// LEGENDA STEP DELLA ROUTINE COD_PROGETTO_s_TXT_DblClick
   '// ========================================================================================================//
   '// STEP 01 = DICHIARAZIONI VARIABILI E ACQUISIZIONE CODICE PROGETTO DALLA FORM
   '// STEP 02 = VERIFICA ESISTENZA PROGETTO SELEZIONATO NELLA TABELLA CONTROL MASTER
   '// STEP 03 = SVUOTO TABELLA CONTROL MASTER TMP E INSERISCO IL PROGETTO SCELTO
   '// STEP 04 = SVUOTO TABELLA TMP DI VISUALIZZAZIONE PAGINE MASTER
   '// STEP 05 = INSERISCO LE PAGINE COLLEGATE AL PROGETTO SCELTO NELLA TABELLA TMP
   '// STEP 06 = INDIVIDUO LA FORM MASTER APERTA (NOMI PREDEFINITI O PERSONALIZZATI)
   '// STEP 07 = SE FORM MASTER APERTA: RESET + CARICA 12 PAGINE + RECALC
   '// STEP 08 = POPOLO TABELLA TMP MSys_TABELLE_Tmp CON IL CODICE PROGETTO
   '// ========================================================================================================//
   
   Public Sub COD_PROGETTO_s_TXT_DblClick(Cancel As Integer)
   
   
   '//---------------------------------------------------------------------------------------------------------//
   'STEP 01 = Dichiarazione variabili e acquisizione del codice progetto dalla form
   '//---------------------------------------------------------------------------------------------------------//
   On Error GoTo Err_COD_PROGETTO_s_Txt_DblClick
                   
                   Dim COD_PROGETTO_s As String
                   Dim DENOM_PROGETTO_s As String
                   Dim FormAperta_b As Boolean
                   Dim Valore_b As Boolean
                                       
                   COD_PROGETTO_s = Me.COD_PROGETTO_s_TXT
                           
   '//---------------------------------------------------------------------------------------------------------//
   'STEP 02 = Verifica se esiste un progetto selezionato nella tabella Control Master
   '//---------------------------------------------------------------------------------------------------------//
                                       sSql = ""
                                       sSql = sSql & "SELECT MSys_TAB401_Qry01_01_GE_TabContr_MASTER.COD_PROGETTO_s, MSys_TAB401_Qry01_01_GE_TabContr_MASTER.Scel_b "
                                       sSql = sSql & "FROM MSys_TAB401_Qry01_01_GE_TabContr_MASTER "
                                       sSql = sSql & "WHERE (((MSys_TAB401_Qry01_01_GE_TabContr_MASTER.Scel_b)=True));"
                                           
                                           Valore_b = False
                                   
                                           Set DaoRs = CurrentDb.OpenRecordset(sSql)
                                               Debug.Print sSql
                                                               
                                               With DaoRs
                                                   
                                                       While Not DaoRs.EOF
                           
                                                       If .Fields("Scel_b") = True Then
                                                               Valore_b = True
                                                               Vv1 = .Fields("COD_PROGETTO_s")
                                                       End If
                                                  
                                                      .MoveNext
                                                       Wend
                                                                           
                                               End With
                                                           
                                           DaoRs.Close
                                           Set DaoRs = Nothing
                                           
                                           If Valore_b = False Then
                                               MsgBox "ATTENZIONE DEVI IMPOSTARE UN PROGETTO!", vbCritical, "CONTROLLO PROGETTI IMPOSTATI"
                                               Me.Scel_b = True = True
                                               Me.Requery
                                               GoTo Exit_COD_PROGETTO_s_Txt_DblClick
                                           End If
                                           
   '//---------------------------------------------------------------------------------------------------------//
   'STEP 03 = Svuoto la tabella Control Master TMP e inserisco il progetto scelto
   '//---------------------------------------------------------------------------------------------------------//
                                     sSql = ""
                                     sSql = sSql & "DELETE MSys_TAB401_Qry01_01_GE_TabContr_MASTER_TMP.* FROM MSys_TAB401_Qry01_01_GE_TabContr_MASTER_TMP; "
                                                  
                                     Debug.Print sSql
                                     CurrentDb.Execute (sSql)
                                     
                                     DoCmd.OpenQuery "MSys_TAB401_Qry01_01_GE_TabContr_MASTER_Insert_toTMP"
                                               
   '//---------------------------------------------------------------------------------------------------------//
   'STEP 04 = Svuoto la tabella TMP di visualizzazione delle pagine master
   '//---------------------------------------------------------------------------------------------------------//
                                   DoCmd.OpenQuery "MSys_TAB402_Qry01_02_VISUAL_PAG_MASTER_Delete_TMP"
                                 
   '//---------------------------------------------------------------------------------------------------------//
   'STEP 05 = Inserisco le pagine collegate al progetto scelto nella tabella TMP
   '//---------------------------------------------------------------------------------------------------------//
                                     DoCmd.OpenQuery "MSys_TAB402_Qry01_02_VISUAL_PAG_MASTER_Insert_TMP"
                                        
   '//---------------------------------------------------------------------------------------------------------//
   'STEP 06 = Individuo la form master aperta tra quelle elencate o con nome personalizzato
   '//---------------------------------------------------------------------------------------------------------//
                   FormAperta_b = False
                   
                   For Each frm In forms
                         
                       Debug.Print "nome della master è " & Application.forms.Application.Name
                       
                       If frm.Name = "MSys_FORM_M01_GESTIONE_OGGETTI_MSys" Then
                           frmName_s = frm.Name
                               Debug.Print frmName_s
                               FormAperta_b = True
                       End If
                        
                       If frm.Name = "ANAG_FrmDF01_M01_GESTIONE" Then
                           frmName_s = frm.Name
                               Debug.Print frmName_s
                               FormAperta_b = True
                       End If
                       
                       If frm.Name = "AA_MSys_FORM_M01_GESTIONE_OGGETTI_MASTER" Then
                           frmName_s = frm.Name
                               Debug.Print frmName_s
                               FormAperta_b = True
                       End If
                       
                       If frmName_s = Null Or frmName_s = "" Then
                           frmName_s = frm.Name
                               Debug.Print frmName_s
                               FormAperta_b = True
                       End If
                               
                               Debug.Print "--------------------------------------------------"
                               Debug.Print "               form corrente                      "
                               Debug.Print frm.Name
                               Debug.Print
                               Debug.Print "--------------------------------------------------"
                       
                                   If FormAperta_b = True Then
                                   
                                       For Each ctl In frm.Controls
                                           Debug.Print ">>>"; ctl.Name
                                           If ctl.Name = "SottoForm_01" Then
                                           End If
                                       Next ctl
                                   End If
                   
                   Next frm
                           
   '//---------------------------------------------------------------------------------------------------------//
   'STEP 07 = Se la form master è aperta eseguo Reset, Carica 12 pagine e Recalc
   '//---------------------------------------------------------------------------------------------------------//
                           If FormAperta_b = True Then
                               
                               forms(frmName_s).Imposta_N21_RESETTA_LE_SOTTOFORM_Sub
                               forms(frmName_s).MSys_TUTTE_LE_PAGINE_Carica_pFunction
                               forms(frmName_s).Recalc
                           End If
                           
   '//---------------------------------------------------------------------------------------------------------//
   'STEP 08 = Popolo la tabella TMP MSys_TABELLE_Tmp con il codice progetto
   '//---------------------------------------------------------------------------------------------------------//
                               DoCmd.OpenQuery "MSys_TABELLE_TMP_Qry01_81_DELETE_TUTTE"
                               
                               sSql = ""
                               sSql = sSql & "SELECT MSys_TABELLE_Tmp.* "
                               sSql = sSql & "FROM MSys_TABELLE_Tmp;"
                               
                               Debug.Print sSql
                               Set DaoRs = CurrentDb.OpenRecordset(sSql)
                               
                               If DaoRs.EOF = True And DaoRs.BOF = True Then
                                   
                                   DaoRs.AddNew
                                       DaoRs.Fields("NRO_OGGETTO_i") = 1
                                       DaoRs.Fields("TIPOGGETTO_s") = "TABLES"
                                       DaoRs.Fields("COD_PROGETTO_s") = COD_PROGETTO_s
                                       
                                   DaoRs.Update
                                   DaoRs.Close
                                   Set DaoRs = Nothing
                               End If
   
   '//---------------------------------------------------------------------------------------------------------//
   '// exit ed errori
   '//---------------------------------------------------------------------------------------------------------//
   Exit_COD_PROGETTO_s_Txt_DblClick:
       Exit Sub
   
   Err_COD_PROGETTO_s_Txt_DblClick:
       MsgBox Err.Description
       Resume Exit_COD_PROGETTO_s_Txt_DblClick
   
   End Sub
   
   
   '//*** fine ***
   '//COD_PROGETTO_s_Txt_DblClick MenuGestito_s, CodiceMenu_s
   '//========================================================================================================//
   
