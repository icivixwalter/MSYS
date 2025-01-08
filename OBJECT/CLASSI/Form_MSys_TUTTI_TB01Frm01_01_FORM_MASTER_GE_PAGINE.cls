VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_MSys_TUTTI_TB01Frm01_01_FORM_MASTER_GE_PAGINE"
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
    '//APRO LA TABELLA DEI PROGETTI DA CARICARE
        DoCmd.OpenTable "MSys_TUTTI_PROG"
        
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
   '@aggiornamento, @prima.di.aggiornare
    Me.DATAAGG_d_TXT = Date
    Me.ORA_AGG_d_TXT = Now()
 
    
End Sub

Private Sub Form_Open(Cancel As Integer)
    
Me.Recalc

End Sub


'//COD_PROGETTO_s_Txt_DblClick MenuGestito_s, CodiceMenu_s
'//========================================================================================================//
Public Sub COD_PROGETTO_s_Txt_DblClick(Cancel As Integer)
 
'//Modello di impostazione della proprieta Caption nella form Master passata come parametro
'MODELLO DI PROVA FUNZIONANTE ---> Application.Forms(LA_FORM_MASTER_s)!Pag_04.Caption = "dento la classe"
'//@CLASSE.MENU.GESTITI _
    @MENU.GESTITI.CLASSE _



On Error GoTo Err_COD_PROGETTO_s_Txt_DblClick
 
            
                Dim COD_PROGETTO_s As String
                Dim DENOM_PROGETTO_s As String
                Dim FormAperta_b As Boolean                 '//se la form è aperta = True
                            
                
                COD_PROGETTO_s = Me.COD_PROGETTO_s_TXT
                DENOM_PROGETTO_s = Me.DENOM_PROGETTO_s_Txt
                            
                  
                  

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
                            
                        '//01 svuoto la tabella tmp
                        DoCmd.OpenQuery "PROGETTI_Msys_TB01Qry01_01_PROJECT_TMP_Delete"
                        
                        '//controllo e stampa dell'area di validita della casella di testo
                        Vv1 = Me.COD_PROGETTO_s
                        Debug.Print "valore della casella di testo in VV1 = " & Vv1
                        
                        '//02 inserisco con una query sql il codice del progetto selezionato nella tabella tmp
                        sSql = ""
                        sSql = sSql & "INSERT INTO PROGETTI_Msys_TB01_PROJECT_TMP ( COD_PROGETTO_s ) "
                        sSql = sSql & "SELECT '" & Me.COD_PROGETTO_s_TXT.Value & "' AS COD_PROGETTO_s "
                        sSql = sSql & "WITH OWNERACCESS OPTION;"
                        Debug.Print sSql
                        CurrentDb.Execute sSql
                '//--------------------------------------------------------------------------------//---------------//
             
                  
            
                        '//PRIMA CONTROLLO SE ESITE UN PROGETTTO SELEZIONATO A TRUE
                        '//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++//
                        '//@CONTROLLO@PROGETTO@NULLO_(se viene deselezionato tutto e non viene scelto nessun progetto rimane impostato su msys)
                        '//@PROGETTO@NON@SELEZIONATO@_(se tutti i progetti sono False allora avviso di impostare un flag)
                            
                            
                            
                            '//controllo se la tabella TUTTI I PROGETTI HA un elemento selezionato
                              '//sql apro il rs
                                    Dim Valore_b As Boolean
                                    
                                    sSql = ""
                                    sSql = sSql & "SELECT MSys_TUTTI_PROG.COD_PROGETTO_s, MSys_TUTTI_PROG.Scel_b "
                                    sSql = sSql & "FROM MSys_TUTTI_PROG;"
                                    'sSql = sSql & "WHERE (((MSys_TUTTI_PROG.Scel_b)=True));"
                        
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
                
                 '//prima cancello i dati nella tmp
                 
                     sSql = ""
                     sSql = sSql & "DELETE MSys_TUTTI_PROG_tmp.* FROM MSys_TUTTI_PROG_tmp; "
                     CurrentDb.Execute (sSql)
                     
                    
                     '//dopo accodo i codici del progetto
                     sSql = ""
                     
                     sSql = sSql & "INSERT "
                     sSql = sSql & "INTO MSys_TUTTI_PROG_TMP "
                     sSql = sSql & "( NRO_OGGETTO_i, COD_PROGETTO_s, DENOM_PROGETTO_s ) "
                     sSql = sSql & "SELECT "
                     sSql = sSql & "MSys_TUTTI_PROG.NRO_OGGETTO_i, "
                     sSql = sSql & "MSys_TUTTI_PROG.COD_PROGETTO_s, "
                     sSql = sSql & "MSys_TUTTI_PROG.DENOM_PROGETTO_s "
                     sSql = sSql & "FROM MSys_TUTTI_PROG "
                     sSql = sSql & "WHERE (((MSys_TUTTI_PROG.COD_PROGETTO_s)='" & Me.COD_PROGETTO_s_TXT & "'));"
                     
                     
                     
                     '//CONTROLLO ED ESECUZIONE
                     Debug.Print sSql
                     
                     CurrentDb.Execute (sSql)
                                     
                '//---------------------------------------------------------------------------------//
                
                
                
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
                
                For Each frm In forms
                    ' 01) CICLO OGGETTI FORM = Stampa nome della maschera correntemente aperta, ma devi aggiungere il nome _
                      della pagina corrente
                      
                    '//@NOME@MASTER_(CONTROLLO IL NOME DELLA MASTER)_@FORM@MASTER
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
                            forms(frmName_s).Imposta_N21_RESETTA_LE_SOTTOFORM_Sub           '//02) ATTIVO IL RESET  = per evitare che rimangano visibili.
                            forms(frmName_s).MSys_TUTTE_LE_PAGINE_Carica_pFunction          '//03) CARICO LE 12 PAGINE RESETTA LE FORM CARICO TUTTE LE 12 PAGINE
                            forms(frmName_s).Recalc                                         '//04) RECALC  della form corrente _
                                                                                            se non metti recalc da errore di procedura obbligatorio _
                                                                                            dopo l'aggiornamento delle 12 pagine.
                        End If
                        
                    
                '//--------------------------------------------------------------------------------//---------------//
            
                    
                '//CODICI PROGETTO IN TMP: SVUOTA LA TABELLA TMP E SALVA NELLA STESSA I CODICI DI PROGETTO
                '//--------------------------------------------------------------------------------//---------------//
            
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


