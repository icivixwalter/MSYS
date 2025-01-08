VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//***************************************************************************************************************//
'//form master : MSys_TUTTI_PROG_FrmM01_01_PROGETTI

'// MODIFICHE: _
      @2024_10_21_ATTIVO_HELP_FORM_MASTER: imposto l'help che spiega il FUNZIONAMENTO DEL MODO DI CARICARE LE _
          PAGINE NELLA FORM MASTER PRINCIPALE. Apro il file help in pdf con le spiegazioni di funzionamento _
          della FORM MASTER PRINCIPALE che utilizza una seconda tabella per caricare le 12 pagine base.

'//***************************************************************************************************************//

Option Compare Database
Option Explicit

'//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++//
'//       LE VARIABILI DI MODULO

'//LE VARIABILI DATABASE
'//....................................................................//
    Dim DaoDB As DAO.Database          '//occorre il riferimento alla libreria  Microsoft DAO 3.6 Object Libray
    
    
    
    Dim DaoWks As DAO.Workspace
    Dim DaoRs As DAO.Recordset

    Dim ADODB As Database
    Dim AdodaoRs As Recordset
    Dim sSql As String                              '//STRINGA SQL
    '//Dim Path_s As String                         '//la path


    '//Contatori
    Dim iCount As Integer
    Dim dbl_count As Double
    
   
    'Le variabili generiche
    Dim Vv1 As Variant
    Dim Dbl1 As Double
    Dim Int1 As Integer
    Dim Long1 As Long
    Dim Str1 As String
    
    '//Messaggi di errore
    Dim ProceduraMessaggioErrore_s As String    '//Errore procedura
    Dim ProceduraAttivaEseguita_s As String     '//Errore Attivita eseguita
    
    
    '//la form
        Dim frm As Form, ctl As Control
        Dim frmName_s As String




    




Private Sub COD_PROGETTO_s_Txt_BeforeUpdate(Cancel As Integer)
    '@aggiornamento, @prima.di.aggiornare
    Me.DATAAGG_d_TXT = Date
    Me.ORA_AGG_d_TXT = Now()
    
    
    
End Sub

Private Sub DENOM_PROGETTO_s_Txt_DblClick(Cancel As Integer)
    '//APRO LA TABELLA DEI PROGETTI DA CARICARE   @in@modifica@oggi
        DoCmd.OpenTable "MSys_TUTTI_TB01_FORM_MASTER_GE_PAGINE"  'old ----> "MSys_TUTTI_PROG"
        
        
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
   '@aggiornamento, @prima.di.aggiornare
    Me.DATAAGG_d_TXT = Date
    Me.ORA_AGG_d_TXT = Now()
 
    
End Sub


Private Sub Form_Open(Cancel As Integer)
    
Me.Recalc

End Sub


'//@evento@click@COD_PROGETTO_s_Txt_DblClick_(attivo la scelta del codice progetto per il caricamento delle pagine NEL TAB CONTROL MASTER)
'//========================================================================================================//
'//codice ----> GeTabControlMaster_COD_PROGETTO_s_Txt_DblClick
'//attivita eseguite: _
     01_GeTabControlSEMPLICE_COD_PROGETTO_s_Txt_DblClick= ESTRAGGO IL CODICE PROGETTO SCELTO = TRUE _
     02_GeTabControlSEMPLICE_COD_PROGETTO_s_Txt_DblClick = SVUOTO LA TAVBELLA CONTROL MASTER TMP _
     03_GeTabControlSEMPLICE_COD_PROGETTO_s_Txt_DblClick = SALVO IL CODICE PROGETTO NELLA TABELLA TMP MSys_TB0401_GE_TabContr_MASTER_TMP _
     04_GeTabControlSEMPLICE_COD_PROGETTO_s = RECUPERO IL CODICE PROGETTO SELEZIONATO DALLA TABELLA TB501 _
                                            con il codice soggetto selezionato apro un rs in cui il campo secl = TRUE e quindi trovo il progetto _
                                            scelto se TRUE salvando il codice progetto, altrimenti Msg. _
     05_GeTabControlSEMPLICE_SVUOTO_TB05_TMP = SVUOTO LA TABELLA TMP DOVE SALVARE IL CODICE PROGETTO TB501 TMP _
     06_GeTabControlSEMPLICE_INSERT_IN_TB05_TMP = INSERISCO LE PAGINE DEL TAB CONTROL SEMPLICE IN TB501 TMP _
            Inserisco nella tabella TB051_TMP tutte le pagine da visualizzare nel TAB CONTROL SEMPLICE. _
     07_GeTabControlSEMPLICE_ATTIVO_LA_MASTER_SEMPLICE = VIENE IMPOSTATA LA FORM DEL TAB CONTROL SEMPLICE


'//========================================================================================================//
Public Sub COD_PROGETTO_s_Txt_DblClick(Cancel As Integer)
 
'//Modello di impostazione della proprieta Caption nella form Master passata come parametro
'MODELLO DI PROVA FUNZIONANTE ---> Application.Forms(LA_FORM_MASTER_s)!Pag_04.Caption = "dentro la classe"
'//@CLASSE.MENU.GESTITI _
    @MENU.GESTITI.CLASSE _



'//TODO: @TB05@DA COSTRUIRE DA ZERO fare la procedura per la gestione del TAB CONTROL SEMPLICE

On Error GoTo Err_COD_PROGETTO_s_Txt_DblClick
 
            
                Dim COD_PROGETTO_s As String
                Dim DENOM_PROGETTO_s As String
                Dim FormAperta_b As Boolean                 '//se la form è aperta = True
                            
                
                COD_PROGETTO_s = Me.COD_PROGETTO_s_TXT
                DENOM_PROGETTO_s = Me.DENOM_PROGETTO_s_Txt
                            
                  
                     Dim tabCtl As TabControl
                  
                '//SOSPESA ...
                '//2024_10_22_PROJECT_TMP =SOSPESO L'ATTIVITA DI SALVATAGGIO DEL CODICE IN PROGETTI TMP
                '//2024_08_23_PROJECT_TMP = Attivazione delle procedura di salvataggio del codice progetto in tmp
                '//--------------------------------------------------------------------------------//---------------//
                '//nota: nella tabella PROGETTI_Msys_TB01Qry01_01_PROJECT_TMP viene salvato il codice del progetto _
                        scelto al fine di poter filtrare le varie form degli oggetti gestiti es: _
                            MSys_QUERY_Frm01_S01_TUTTE _
                            MSys_REPORT_Frm01_S01_TUTTI ..... ecc.
                '//devo inserire all'inizio dell'evento doppio click l'aggiornamento della tabella tmp perchè alla fine _
                    della routine vi è aggiornamento delle pagine e si perde il valore del codice progetto. Posto all'inizio _
                    posso eseeguire la cancellazione della tmp e l'aggiornamento della tabella tmp senza problemai di perdita _
                    di validita dell'area di memoria della variabile vv1
                            
                        '//01 svuoto la tabella tmp   **** TODO: controllare se cancellare oppure riutilizzare
                        'DoCmd.OpenQuery "PROGETTI_Msys_TB01Qry01_01_PROJECT_TMP_Delete"
                        
                        '//controllo e stampa dell'area di validita della casella di testo
                        'Vv1 = Me.COD_PROGETTO_s
                        'Debug.Print "valore della casella di testo in VV1 = " & Vv1
                        
                        '//**** TODO: controllare se cancellare oppure riutilizzare
                        '//02 inserisco con una query sql il codice del progetto selezionato nella tabella tmp
                        'sSql = ""
                        'sSql = sSql & "INSERT INTO PROGETTI_Msys_TB01_PROJECT_TMP ( COD_PROGETTO_s ) "
                        'sSql = sSql & "SELECT '" & Me.COD_PROGETTO_s_TXT.Value & "' AS COD_PROGETTO_s "
                        'sSql = sSql & "WITH OWNERACCESS OPTION;"
                        'Debug.Print sSql
                        'CurrentDb.Execute sSql
                '//--------------------------------------------------------------------------------//---------------//
             
                  
            
                        '//PRIMA CONTROLLO SE ESITE UN PROGETTTO SELEZIONATO A TRUE
                        '//01_GeTabControlSEMPLICE_COD_PROGETTO_s_Txt_DblClick= ESTRAGGO IL CODICE PROGETTO SCELTO = TRUE
                        '//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++//
                        '//@CONTROLLO@PROGETTO@NULLO_(se viene deselezionato tutto e non viene scelto nessun progetto rimane impostato su msys)
                        '//@PROGETTO@NON@SELEZIONATO@_(se tutti i progetti sono False allora avviso di impostare un flag)
                            
                            
                            
                            '//@IN@SOSTITUZIONE
                            '//controllo se la tabella TUTTI I PROGETTI HA un elemento selezionato
                              '//sql apro il rs
                                    Dim Valore_b As Boolean
                                    'old
                                    'sSql = ""
                                    'sSql = sSql & "SELECT MSys_TUTTI_PROG.COD_PROGETTO_s, MSys_TUTTI_PROG.Scel_b "
                                    'sSql = sSql & "FROM MSys_TUTTI_PROG;"
                                    'sSql = sSql & "WHERE (((MSys_TUTTI_PROG.Scel_b)=True));"
                                    
                                    
                                    '//@2024.10_22.NUOVO_CODICE
                                    '//@sostituito.nuovo
                                    '//codice estrazione MSys_TUTTI_TB01
                                    
                                    '//@02_@2024.10_30.MODIFICA= da vecchia MSys_TB0501_GE_TabContr_SEMPLICE a nuova tabela tb05 = MSys_TB0501_GE_TabContr_SEMPLICE
                                    
                                    
                                    '//04_GeTabControlSEMPLICE_COD_PROGETTO_s = RECUPERO IL CODICE PROGETTO SELEZIONATO DALLA TABELLA TB501 _
                                            con il codice soggetto selezionato apro un rs in cui il campo secl = TRUE e quindi trovo il progetto _
                                            scelto se TRUE salvando il codice progetto, altrimenti Msg.
                                    sSql = ""
                                    sSql = sSql & "SELECT MSys_TB0501_GE_TabContr_SEMPLICE.COD_PROGETTO_s, MSys_TB0501_GE_TabContr_SEMPLICE.Scel_b "
                                    sSql = sSql & "FROM MSys_TB0501_GE_TabContr_SEMPLICE "
                                    sSql = sSql & "WHERE (((MSys_TB0501_GE_TabContr_SEMPLICE.Scel_b)=True));"
                                    
                                    '//controllo
                                     Debug.Print sSql

                        
                                        Valore_b = False
                                
                                        'Apro rs estrazione Sql
                                        Set DaoRs = CurrentDb.OpenRecordset(sSql)
                                            '//controllo
                                            Debug.Print sSql
                                                            
                                            With DaoRs
                                                
                                                'CICLO RS
                                                '..................................................................................................
                                                    While Not DaoRs.EOF
                        
                        
                                                    'recupero il valore del campo
                                                    If .Fields("Scel_b") = True Then
                                                            Valore_b = True
                                                            
                                                            '//SALVO IL CODICE PROGETTO
                                                            Vv1 = .Fields("COD_PROGETTO_s")
                                                  
                                                    End If
                                               
                        
                                                   .MoveNext
                                                    Wend
                        
                                                'FINE CICLO RS
                                                '..................................................................................................
                                                                        
                                                                            
                                                
                                            End With
                                                        
                                                            
                                                        
                                        'Chiudo rs estrazione Sql
                                        DaoRs.Close
                                        Set DaoRs = Nothing
                                        
                                        '//@PROGETTO@NULLO_(avviso che devi selezionare un progetto) _
                                            @ERRORE@PROGETTO@NULLO_(avviso che non hai impostato il progetto)
                                        If Valore_b = False Then
                                            '//@CONTROLLO@PROGETTO_(SE NON IMPOSTATO @AVVISO DI @IMPOSTAZIONE)
                                            MsgBox "ATTENZIONE DEVI IMPOSTARE UN PROGETTO!", vbCritical, "CONTROLLO PROGETTI IMPOSTATI"
                                            Me.Scel_b = True
                                            Me.Requery
                                            GoTo Exit_COD_PROGETTO_s_Txt_DblClick
                                        End If
                                        
                        '//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++//
                
                
                '//SE IL PROGETTO E' STATO SCELTO ESEGUO L'ACCODAMENTO DEI DATI PROGETTO SULLA TABELLA TMP
                '//---------------------------------------------------------------------------------//
                '//codice --->  @INSERTI@CODICI@PROGETTO@TMP_(inserisco i codici del progetto nella casella PROGETTO TMP)
                '// @2024_03_11_CREATO_INSERIMENTO_CODICI_PROGETTO_IN_TMP
                
                 '//SOSPESA ---> prima cancello i dati nella tmp
                  
                     'sSql = ""
                     'sSql = sSql & "DELETE MSys_TUTTI_PROG_tmp.* FROM MSys_TUTTI_PROG_tmp; "
                     'CurrentDb.Execute (sSql)
                     
                     
                     '//@ma@inserisce@in@tmp ??? = da verificare
                    
                     '// SOSPESA ---> dopo accodo i codici del progetto
                     '//==================================================================//
                      'sSql = ""
                      
                      'sSql = sSql & "INSERT "
                      'sSql = sSql & "INTO MSys_TUTTI_PROG_TMP "
                      'sSql = sSql & "( NRO_OGGETTO_i, COD_PROGETTO_s, DENOM_PROGETTO_s ) "
                      'sSql = sSql & "SELECT "
                      'sSql = sSql & "MSys_TUTTI_PROG.NRO_OGGETTO_i, "
                      'sSql = sSql & "MSys_TUTTI_PROG.COD_PROGETTO_s, "
                      'sSql = sSql & "MSys_TUTTI_PROG.DENOM_PROGETTO_s "
                      'sSql = sSql & "FROM MSys_TUTTI_PROG "
                      'sSql = sSql & "WHERE (((MSys_TUTTI_PROG.COD_PROGETTO_s)='" & Me.COD_PROGETTO_s_TXT & "'));"
                      
                      
                      
                      '//CONTROLLO ED ESECUZIONE
                      'Debug.Print sSql
                      
                      'CurrentDb.Execute (sSql)
                      
                      
                     '//==================================================================//
                     
                     
                       
                     '//                                 LE PAGINE DA VISUALIZZARE NELLA TABELLA TB05 TMP  *** INIZIO ***
                     '//=============================================================================================================================//
                    
                     
                     '//@attenzione= INSERISCE nella tabella tmp TUTTI I PROGETTI ma bisogna verificare _
                        se nella tabella collegata inserisce anche le pagine ??? : MSys_TUTTI_PROG_VIS_PAGINE
                        
                        '//controllare le pagine inserite ??? : @controllare@salvataggio@da@MASTER---->detail
                        'todo: finire la procedura.... occorre salvare le pagine del codice progetto _
                        nella tabella tmp denominata MSys_TUTTI_PROG_VIS_PAGINE_TMP
                        
                        '//TODO: MODIFICARE IL SALVATAGGIO DEL CODICE PROGETTO E DELLE PAGINE : _
                          Le pagine da visualizzare nel TAB CONTROLL MASTER che viene utilizzato dalla FORM MASTER, _
                          deve avere questo funzionamento SECONDO IL PROGETTO SCELTO: _
                            01) una volta scelto il progetto, salvo il CODICE DI PROGETTO nella tabella, tmp : _
                                da MSys_TUTTI_TB01_FORM_MASTER_GE_PAGINE in ----> MSys_TUTTI_TB01_FORM_MASTER_GE_PAGINE_TMP _
                                                            qui viene salvatato :       il codice del progetto. _
                            02) le pagine da visualizzare nella FORM MASTER vengono salvate dalla tabella VISUALIZZA LE PAGINE alla tmp _
                                da MSys_TUTTI_TB01_FORM_MASTER_VISUAL_PAGINE in ----> MSys_TUTTI_TB01_FORM_MASTER_VISUAL_PAGINE_TMP.
                                
                            '//@2024.10_22.NUOVO_CODICE
                            
                            '//02_GeTabControlSEMPLICE_COD_PROGETTO_s_Txt_DblClick = SVUOTO LA TAVBELLA CONTROL MASTER TMP
                            '//ESECUZIONE DEL PROGETTO: _
                              01) PRIMA SOSTITUZIONE = svuoto la tabella MSys_TUTTI_TB01_FORM_MASTER_GE_PAGINE_TMP dove salvo _
                                  il codice del progetto nuovo.  CODICE OGGETTI ---> MSys_TUTTI_TB01
                                  
                                  
                                  '//@03_@2024.10_30.MODIFICA= da vecchia a nuova tabela MSys_TB0401_GE_TabContr_MASTER
                                  
                                  '//05_GeTabControlSEMPLICE_SVUOTO_TB05_TMP = SVUOTO LA TABELLA TMP DOVE SALVARE IL CODICE PROGETTO TB501 TMP
                                  sSql = ""
                                  sSql = sSql & "DELETE MSys_TB0501_GE_TabContr_SEMPLICE_TMP.* FROM MSys_TB0501_GE_TabContr_SEMPLICE_TMP; "
                                               
                                  '//CONTROLLO ED ESECUZIONE
                                  Debug.Print sSql
                                  CurrentDb.Execute (sSql)
                                  
                                  

                                    
                                    '//06_GeTabControlSEMPLICE_INSERT_IN_TB05_TMP = INSERISCO LE PAGINE DEL TAB CONTROL SEMPLICE IN TB501 TMP
                                            DoCmd.OpenQuery "MSys_TB0502Qry01_02_VISUAL_PAG_SEMPLICE_Insert_TMP"
                                            
                                            
                                  '//TODO: 04.11.2024 FINO QUI .... CONTINUARE ... = perche se il tab control semplice viene inserito nel _
                                          tab control MASTER, quando viene aggiornato il codice progetto nel Tab control SEMPLICE si attiva _
                                          per errore il TAB CONTROL MASTER ??? Gli eventi devono essere separati.
                                                            
                                            
                                          
                                    
                                             
                                
                                
                                  
                                '//                                 LE PAGINE DA VISUALIZZARE NELLA TABELLA TB05 TMP  *** INIZIO ***
                                '//=============================================================================================================================//
                               
                
                
                
                '//2022.06.26_AggiornoTabellaTmp,
                '//@aggiorno.codice.progetto.tmp,
                '//svuoto la tabella tmp  --
                '//--------------------------------------------------------------------------------//---------------//
                '//CON L'EVENTO DOPPIO CLICK attivo il controllo oggetti della form master che è stata modificata _
                    con in questa data : _
                    @2022.03.07.modifiche = Modifica del sistema del controllo della form master con questo procedimento _
                    attivando questi punti : _
                    01) CICLO OGGETTI FORM _
                    02) ATTIVO IL RESET _
                    03) CARICO LE  12 PAGINE _
                    04) RECALC  della form corrente
                 '//CICLO OGGETTI = Enumera insieme Forms individuo la form master ed attribuisco alla variabile _
                                    il nome della form master corrente per utilizzarla nell'accesso della routine pubblica _
                                    resetta le form
                
                '//IMPOSTO LA VARIABILE FORM APERTA A FALSE
                FormAperta_b = False
                
                '@riferimento@tab@control_(faccio riferimento alla pagina corrente aperta del tab control)
                    
                      'Dim tabCtl As TabControl
                       'im ctlCurrent As Control
                      
                      
                
                For Each frm In forms
                    ' 01) CICLO OGGETTI FORM = Stampa nome della maschera correntemente aperta, ma devi aggiungere il nome _
                      della pagina corrente
                      
                    '//@NOME@MASTER_(CONTROLLO IL NOME DELLA MASTER)_@FORM@MASTER
                    Debug.Print "nome della master è " & Application.forms.Application.Name & " " & frm.Name
                    
                    
                    '//SE CI TROVIAMO NEL FORM MASTER ALLORA SIGNIFICA CHE LA TAB CONTROL SEMPLICE SI TROVA AL SUO INTERNO _
                      e quindi devi recuperare la pagina in cui si trova ed a limite effettuare il requery della pagine che al momento _
                      è la pagina 8
                    If frm.Name = "AA_MSys_FORM_M01_GESTIONE_OGGETTI_MASTER" Then
                                  
                            
                            FormAperta_b = True
                            
                            '//modifica   @ATTIVATO@MASTER@TAB@CONTROL@SEMLICE
                            Me.Parent!TxtTit_01.Caption = "AGGIORNATO PAGINE NEL TAB CONTROL SEMPLICE - CODICE PROGETTO: " & Me.COD_PROGETTO_s_TXT
                            
                    End If
                     
                     '//todo: cambiare  l'aggiornamento in tab
                   ' If frm.Name = "MSys_TB0501Frm00_M01_MASTER_SEMPLICE" Then
                   '     frmName_s = frm.Name
                   '         Debug.Print frmName_s
                   '         FormAperta_b = True
                   ' End If
                    
                    
                    '//@LA FORM ?
                   ' If frm.Name = "" Then
                   '     frmName_s = frm.Name
                   '         Debug.Print frmName_s
                   '         FormAperta_b = True
                   ' End If
                    
                    
            
                    
                    
                    
                  
                    
                 '//NON ESISE LA FORM PREIMPOSTATA ALLORA ASSUME QUELLA APERTA ma deve essere master _
                    inserita il 2023_07_02 perche quando cambi il nome della form master essa non _
                    viene riconosciuta fra quelle preimpostate e la procedura si Blocca. In questo _
                    modo possiamo utilizzare a form master generale anche cambiando denominazione. _
                        @cambio@denominazione@form@master_(se cambi denominazione alla form master viene utilizzata ugualmente)
                    
                    If frmName_s = Null Or frmName_s = "" Then
                        '//MODIFICATO  inserito il true per la form aperta, in questo modo anche se non è indicato come nome nella stringa _
                                        precedente la form segnalata come aperta permette l'aggiornamento delle 12 sottoform.
                        frmName_s = frm.Name
                            Debug.Print frmName_s
                            FormAperta_b = True

                    End If
                                 
                            '//controllo form corrente
                            Debug.Print "--------------------------------------------------"
                            Debug.Print "               form corrente                      "
                            Debug.Print frm.Name
                            Debug.Print
                            Debug.Print "--------------------------------------------------"
                    
                        
                                '//se la form è aperta enumero i controlli
                                If FormAperta_b = True Then
                                
                                    ' Enumera insieme Controls di ciascuna maschera.
                                    For Each ctl In frm.Controls
                                        ' Stampa nome di ciascun controllo della form corrente.
                                        Debug.Print ">>>"; ctl.Name
                                        If ctl.Name = "SottoForm_01" Then
                                            'Stop
                                            
                                        End If
                                    
                                    Next ctl
                                End If   '//se la form è aperta
                
                Next frm
                        
                        
                        '//@CONTROLLO@FORM@SINGOLA_(CONTROLLO SE LA FORM APERTA SI TROVA SINGOLA O ANNIDATA)
                        '//@FORM@MASTER@CHIUSA_(Non carico le pagine ma salto.Vengono resettate le pagine e caricate di nuovo se è aperta la FORM MASTER)
                        '//se la form è aperta aggiorno le pagine
                        If FormAperta_b = True Then
                            
                            '//@aggiorna.le.pagine - della FORM MASTER, @aggiorna.progetto @chiamo@form@esterne
                            '//INSIEME FORMS _
                            Form è l'insieme che rappresenta tutte le maschere aperte a cui puoi fare riferimento con il nome _
                            ed attivare le sue proprieta come la funzione ..MSys_TUTTE_LE_PAGINE_Carica_pFunction; in questo caso _
                            l'oggetto form utilizza la variabile frmName_s per individuare quale form master  è aperta ed utilizzare _
                            la routine di reset delle form; esempio di modello : _
                                Forms![MSys_FORM_M01_GESTIONE_OGGETTI_MSys].MSys_TUTTE_LE_PAGINE_Carica_pFunction
                            '// con lo stesso sistema attivo quindi le tre condizioni : RESET + CARICA LE FORM + RECALC
                            
                            
                            '//RESETTO SE TAB CONTROL SEMPLICE SE SI TROVA DA SOLO OPPURE SE SI TROVA NELLA MASTER PRINCIPALE
                            '//===========================================================================================//
    
                                If frmName_s = "AA_MSys_FORM_M01_GESTIONE_OGGETTI_MASTER" Then
                                  'provo l'aggiornamento
                                      
                                      db_agg.Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE.Requery
                                      db_agg.Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE.Recalc
                                      
                                      db_agg.Form_MSys_TB0501Frm00_M01_MASTER_SEMPLICE.Recalc
                                  
                                      
                                     
                                      
                                      '//TODO: per gestire le pagine nel tab control semplice annidato in quello Master _
                                      occorre : 1) creare un rs che legge le pagine scelte
                                      'db_agg.Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE.Parent!PAG_02.Requery
                                      'db_agg.Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE.Parent!PAG_03.Requery
                                      'db_agg.Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE.Parent!Pag_04.Requery
                                      'db_agg.Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE.Parent!PAG_05.Requery
                                      'db_agg.Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE.Parent!Pag_06.Requery
                                      'db_agg.Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE.Parent!Pag_07.Requery
                                      'db_agg.Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE.Parent!Pag_08.Requery
                                      'db_agg.Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE.Parent!Pag_09.Requery
                                      'db_agg.Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE.Parent!Pag_10.Requery
                                      'db_agg.Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE.Parent!PAG_11.Requery
                                      
                                      
                                        '//TODO: COMPLETARE IL CONTROLLO DELLE PAGINE DELLA SOTTOFORM ANNIDATA'
                                        '//       CONTROLLO PAGINE DELLA SOTTOFORM TAB CONTROL SEMPLICE ANNIDATA
                                        '//===========================================================================================//
                                        '//NOTE: creo una query select tra le due tabelle ed estraggo tutte le pagine da visualizzare _
                                          individuando il nome della pagine salvata nel campo NOMEpagina_s che ha il valore VISIBLE_b = true _
                                          riesco a capire che la pagina ad es. Pag_01 = True deve essere visibile mentre la _
                                                                                 Pag_02 = False non deve essere visibile. In questo modo _
                                            posso attivare della classe della form che rende visibile o non visibile la pagina salvata : _
                                                 db_agg.Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE.Parent!PAG_01.Visible = True _
                                          Esempio la pagina 1 deve essere impostata a visibile. Visto che la struttura della query _
                                          puo essere utilzzata piu volte la stessa dovrebbe essere inserita in una funzione con la _
                                          possibilita di restituire la struttura sql per altri utilizzi.
                                        
                                        sSql = ""   '
                                        sSql = sSql & "SELECT "
                                        sSql = sSql & "MSys_TB0501_GE_TabContr_SEMPLICE.COD_PROGETTO_s, "
                                        sSql = sSql & "MSys_TB0501_GE_TabContr_SEMPLICE.Scel_b, "
                                        sSql = sSql & "MSys_TB0501_GE_TabContr_SEMPLICE.DENOM_PROGETTO_s, "
                                        sSql = sSql & "MSys_TB0502_VISUAL_PAG_SEMPLICE.TIPOGGETTO_s, "
                                        sSql = sSql & "MSys_TB0502_VISUAL_PAG_SEMPLICE.NOMEpagina_s, "
                                        sSql = sSql & "MSys_TB0502_VISUAL_PAG_SEMPLICE.NOMEOGGETTO_s, "
                                        sSql = sSql & "MSys_TB0502_VISUAL_PAG_SEMPLICE.VISIBLE_b "
                                        sSql = sSql & "FROM "
                                        sSql = sSql & "MSys_TB0501_GE_TabContr_SEMPLICE "
                                        sSql = sSql & "INNER JOIN MSys_TB0502_VISUAL_PAG_SEMPLICE "
                                        sSql = sSql & "ON MSys_TB0501_GE_TabContr_SEMPLICE.COD_PROGETTO_s = MSys_TB0502_VISUAL_PAG_SEMPLICE.COD_PROGETTO_s "
                                        sSql = sSql & "WITH OWNERACCESS OPTION;"
                                        
                                        
                                        '//controllo
                                        Debug.Print sSql
                                        
                                        
                                        Valore_b = False
                                        
                                        'Apro rs estrazione Sql
                                        Set DaoRs = CurrentDb.OpenRecordset(sSql)
                                        '//controllo
                                        Debug.Print sSql
                                        
                                          With DaoRs
                                        
                                            'CICLO RS   TODO: DA FINIRE ???
                                            '..................................................................................................
                                            While Not DaoRs.EOF
                                        
                                        
                                                
                                                
                                                Select Case .Fields("TIPOGGETTO_s")
                                                  Case "SottoForm_01"
                                                      '//PAG_01
                                                      If .Fields("VISIBLE_b") = True Then
                                                         db_agg.Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE.Parent!PAG_01.Visible = True
                                                      Else
                                                         db_agg.Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE.Parent!PAG_01.Visible = False
                                                      End If
                                                      
                                                      
                                                  
                                                  Case "SottoForm_02"
                                                      Case "SottoForm_02"
                                                      '//PAG_02
                                                      If .Fields("VISIBLE_b") = True Then
                                                         db_agg.Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE.Parent!PAG_02.Visible = True
                                                      Else
                                                         db_agg.Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE.Parent!PAG_02.Visible = False
                                                      End If
    
                                                  
                                                  Case "SottoForm_03"
                                                      '//PAG_03
                                                      If .Fields("VISIBLE_b") = True Then
                                                         db_agg.Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE.Parent!PAG_03.Visible = True
                                                      Else
                                                         db_agg.Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE.Parent!PAG_03.Visible = False
                                                      End If
    
                                                  Case "SottoForm_04"
                                                      '//PAG_04
                                                      If .Fields("VISIBLE_b") = True Then
                                                         db_agg.Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE.Parent!Pag_04.Visible = True
                                                      Else
                                                         db_agg.Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE.Parent!Pag_04.Visible = False
                                                      End If
                                                      
                                                  Case "SottoForm_05"
                                                      '//PAG_05
                                                      If .Fields("VISIBLE_b") = True Then
                                                         db_agg.Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE.Parent!PAG_05.Visible = True
                                                      Else
                                                         db_agg.Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE.Parent!PAG_05.Visible = False
                                                      End If
    
                                                 Case "SottoForm_06"
                                                      '//PAG_06
                                                      If .Fields("VISIBLE_b") = True Then
                                                         db_agg.Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE.Parent!Pag_06.Visible = True
                                                      Else
                                                         db_agg.Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE.Parent!Pag_06.Visible = False
                                                      End If
                                                      
                                                      
                                                  Case "SottoForm_07"
                                                      '//PAG_07
                                                      If .Fields("VISIBLE_b") = True Then
                                                         db_agg.Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE.Parent!Pag_07.Visible = True
                                                      Else
                                                         db_agg.Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE.Parent!Pag_07.Visible = False
                                                      End If
                                                      
                                                 Case "SottoForm_08"
                                                      '//PAG_08
                                                      If .Fields("VISIBLE_b") = True Then
                                                         db_agg.Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE.Parent!Pag_08.Visible = True
                                                      Else
                                                         db_agg.Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE.Parent!Pag_08.Visible = False
                                                      End If
    
                                                  Case "SottoForm_09"
                                                      '//PAG_09
                                                      If .Fields("VISIBLE_b") = True Then
                                                         db_agg.Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE.Parent!Pag_09.Visible = True
                                                      Else
                                                         db_agg.Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE.Parent!Pag_09.Visible = False
                                                      End If
    
                                                  Case "SottoForm_10"
                                                      '//PAG_10
                                                      If .Fields("VISIBLE_b") = True Then
                                                         db_agg.Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE.Parent!Pag_10.Visible = True
                                                      Else
                                                         db_agg.Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE.Parent!Pag_10.Visible = False
                                                      End If
                                                  
                                                  Case "SottoForm_11"
                                                      '//PAG_11
                                                      If .Fields("VISIBLE_b") = True Then
                                                         db_agg.Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE.Parent!PAG_11.Visible = True
                                                      Else
                                                         db_agg.Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE.Parent!PAG_11.Visible = False
                                                      End If
                                                      
                                                  Case "SottoForm_12"
                                                      '//PAG_12
                                                      If .Fields("VISIBLE_b") = True Then
                                                         db_agg.Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE.Parent!Pag_12.Visible = True
                                                      Else
                                                         db_agg.Form_MSys_TB0501Frm01_01_GE_TabContr_SEMPLICE.Parent!Pag_12.Visible = False
                                                      End If
    
    
                                                  
                                                  End Select
                                        
                                        
                                                .MoveNext
                                            Wend
                                        
                                            'FINE CICLO RS
                                            '..................................................................................................
                                        
                                        
                                        
                                          End With
                                        
                                        
                                        
                                        'Chiudo rs estrazione Sql
                                        DaoRs.Close
                                        Set DaoRs = Nothing
                                        
                                        '// *** FINE ***'
                                        '//       CONTROLLO PAGINE DELLA SOTTOFORM TAB CONTROL SEMPLICE ANNIDATA
                                        '//===========================================================================================//
    
                                      
                                      
                                 
                                Else
                                  DoEvents
                                  
                                  '//RESETTO SE TAB CONTROL SEMPLICE SI TROVA DA SOLO
                                  '//===========================================================================================//
    
                                      forms(frmName_s).Imposta_N21_RESETTA_LE_SOTTOFORM_Sub           '//02) ATTIVO IL RESET  = per evitare che rimangano visibili.
                                      forms(frmName_s).MSys_TUTTE_LE_PAGINE_Carica_pFunction          '//03) CARICO LE 12 PAGINE RESETTA LE FORM CARICO TUTTE LE 12 PAGINE
                                      forms(frmName_s).Recalc                                         '//04) RECALC  della form corrente _
                                                                                                    se non metti recalc da errore di procedura obbligatorio _
                                                                                                    dopo l 'aggiornamento delle 12 pagine.
                                                                                                    
                                      End If
                                      
                                  '//===========================================================================================//

                           
                            
                        End If '//If FormAperta_b = True Then
                        
                    
                '//--------------------------------------------------------------------------------//---------------//
            
                    
                '//CODICI PROGETTO IN TMP: SVUOTA LA TABELLA TMP E SALVA NELLA STESSA I CODICI DI PROGETTO
                '//--------------------------------------------------------------------------------//---------------//
                    '//todo: DA CANCELLARE FORSE NON SERVE ???
                    DoCmd.OpenQuery "MSys_TABELLE_TMP_Qry01_81_DELETE_TUTTE"
                    
                    '//sql apro il rs
                    sSql = ""
                    sSql = sSql & "SELECT MSys_TABELLE_Tmp.* "
                    sSql = sSql & "FROM MSys_TABELLE_Tmp;"
                    
                    'controllo  @aggiungo.nuovo.record.query
                    Debug.Print sSql
                    Set DaoRs = CurrentDb.OpenRecordset(sSql)
                    
                    '//controllo se il rs è popolato = FALSE esistono record True senza record
                    If DaoRs.EOF = True And DaoRs.BOF = True Then
                        
                        
                        DaoRs.AddNew
                            '//@SALVO@CODICE@PROGETTO@ @TABELLA@TMP
                            'inserisco un nuovo record con tutti i codici del progetto corrente
                            DaoRs.Fields("NRO_OGGETTO_i") = 1
                            DaoRs.Fields("COD_PROGETTO_s") = COD_PROGETTO_s
                            
                            '//DaoRs.Fields("NOME_OGGETTO_s") = DENOM_PROGETTO_s    PER ORA NON ESISTE
                            
                            
                        DaoRs.Update
                        DaoRs.Close
                        Set DaoRs = Nothing
                    
                    End If
                    
                '//--------------------------------------------------------------------------------//---------------//
             



'//ERRORI ED USCITA
'//--------------------------------------------------------------------------------//-----------------
Exit_COD_PROGETTO_s_Txt_DblClick:
    Exit Sub

Err_COD_PROGETTO_s_Txt_DblClick:
    MsgBox Err.Description
    Resume Exit_COD_PROGETTO_s_Txt_DblClick

End Sub

'//*** fine ***
'//COD_PROGETTO_s_Txt_DblClick MenuGestito_s, CodiceMenu_s
'//========================================================================================================//








'***********************************************************************************************************************
'                           HELP *** INIZIO ***
'// @01 = APRO HELP VISUALIZZA PAGINE MASTER
'***********************************************************************************************************************



'//EVENTO_DOPPIO_CLICK
  '//=================================================================================================================//
  '//MODELLO ROUTINE DOPPIO CLICK = ..... _
    CODICE----> Help_Txt_DblClick   @2024_10_21_ATTIVO_HELP_FORM_MASTER
  '//=================================================================================================================//
  '//Parametri che possono essere cambiati collettivamente:
  '//--------------------------------------------------------------@
  '//
  '//DENOMINAZIONE---------> Help_Txt_DblClick
  '//TIPO------------------> ROUTINE pubblica.
  '//ATTIVITA--------------> APRI FILE PDF HELP = Apro il file help
  '//ProceduraMessaggioErrore_s = "APRI FILE PDF HELP"
  '//ProceduraAttivaEseguita_s = "Help_Txt_DblClick"
  '//
  '//--------------------------------------------------------------@
  '//PARAMETRI-------------> Nessuno
  '//VALORE_DI_RITORNO-----> Nulla
  '//CODICE----------------> Help_Txt_DblClick.01.00
  '//
  '//=================================================================================================================//
  '//NOTE------------------> Apro il file help con il doppio click
  '// @01 = APRO HELP VISUALIZZA PAGINE MASTER
  '//


  Private Sub Help_Txt_DblClick(Cancel As Integer)

  On Error GoTo Help_Txt_DblClick_Err

    '//APRO_FORM_FALDONI = Chiamo procedura
    '//---------------------------------------------------------------------------------------//
    '//CODICE -----------> Help_Txt_DblClick.01.01

            '//Call Help_Txt_DblClick
    '//---------------------------------------------------------------------------------------//
    
   
    
    '//APRO_FORM_FALDONI = Apertura form + aggiornamento nome faldoni
    '//---------------------------------------------------------------------------------------//
    '//CODICE -----------> Help_Txt_DblClick.01.02
    
                '//RESET
                ProceduraMessaggioErrore_s = ""
                ProceduraAttivaEseguita_s = ""
   

                '//IMPOSTO LE VARIABILI
                ProceduraMessaggioErrore_s = "ERRORE NELLA PROCEDURA : APRI FILE PDF HELP"
                ProceduraAttivaEseguita_s = "Help_Txt_DblClick"


    
    '//---------------------------------------------------------------------------------------//


       '//APRO PATH O FILE GENERICI
           '//.....................................................................................................
              '//File da ricercare e la path Viene attivata la procedura solo se sono valorizzate le variabili
                                               
              '//IMPOSTO LE VARIABILI PATH E FILE
              '//
                Dim MyFile_s As String
              '//
                Dim MyPath_s  As String

           
               
                 MyFile_s = ""
                 MyPath_s = ""
                 
                 Str1 = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\OBJECT\HELP\HELP_FORM\"
                 '//PATH O FILE = Inserire o una path oppure una path+file con percoso esteso.
                 '
                 
                 
                 MyPath_s = Str1
                 MyFile_s = "HELP_MSys_TUTTI_TB01_FORM_MASTER_GE_PAGINE.pdf"
                 
                  
                 If Str1 > "" Then
                 '//IMPOSTO LA VARIABILE OGGETTO = per l'utilizzo della classe
                 '
                  Dim MyCls_LLPP_ATTI_Tb01_Gestione_Classe As LLPP_ATTI_Tb01_Gestione_Classe
                 
                 '//ISTANZIO LA CLASSE = MyCls_CLASSE_GESTIONE_OGGETTI_FORM assegnando _
                 l 'oggetto di classe alla variabile e creo creo l'oggetto di classe
                 '
                 Set MyCls_LLPP_ATTI_Tb01_Gestione_Classe = New LLPP_ATTI_Tb01_Gestione_Classe
                 
                 
                 '
                 MyCls_LLPP_ATTI_Tb01_Gestione_Classe.APRO_PATH_O_FILE_GENERICI_pSub (MyPath_s & MyFile_s)
                 
                 
                 '//RILASCIO LA MEMORIA = rilascio l'oggetto ed attivo il distruttore della classe
                 '
                 Set MyCls_LLPP_ATTI_Tb01_Gestione_Classe = Nothing
                 
                 Else
                         MsgBox "ATTENZIONE I PARAMETRI file e path non sono valorizzati per l'apertura del file - Uscita dalla Routine", vbExclamation
                 
                 End If 'If par_FileAtto_s > "" And Str1 > "" Then
                 
           '//.....................................................................................................
          
Help_Txt_DblClick_Err:
    Debug.Print "uscita dalla routine : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
   
    Exit Sub
    
    MsgBox Err.Description & " " & ProceduraMessaggioErrore_s & " - " & ProceduraAttivaEseguita_s & " --> errore Error$ -> : " & Error$
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Debug.Assert "BLOCCO PROCEDURA -> " & ProceduraAttivaEseguita_s

    Stop
    Resume Help_Txt_DblClick_Err
    
  End Sub
  '//'//DENOMINAZIONE---------> Help_Txt_DblClick *** FINE ***
  '//=================================================================================================================//




'***********************************************************************************************************************
'                           HELP *** FINE ***
'
'***********************************************************************************************************************


