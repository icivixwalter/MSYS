VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_MSys_TUTTI_PROG_FrmM01_01_PROGETTI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
                
                For Each frm In Forms
                    ' 01) CICLO OGGETTI FORM = Stampa nome della maschera correntemente aperta, ma devi aggiungere il nome _
                      della pagina corrente
                      
            
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
                        frmName_s = frm.Name
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
                            Forms(frmName_s).Imposta_N21_RESETTA_LE_SOTTOFORM_Sub           '//02) ATTIVO IL RESET  = per evitare che rimangano visibili.
                            Forms(frmName_s).MSys_TUTTE_LE_PAGINE_Carica_pFunction          '//03) CARICO LE 12 PAGINE RESETTA LE FORM CARICO TUTTE LE 12 PAGINE
                            Forms(frmName_s).Recalc                                         '//04) RECALC  della form corrente _
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








'....................................................................
