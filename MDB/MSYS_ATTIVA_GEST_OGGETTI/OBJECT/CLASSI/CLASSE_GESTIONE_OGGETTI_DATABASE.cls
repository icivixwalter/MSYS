VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLASSE_GESTIONE_OGGETTI_DATABASE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//CODICI ANSI: _
        .................................................................................... _
            TODO: inserire altri codici _
            & Chr$(13) & = INTERLINEA singola per avere una doppia interlinea ripere il comando. _
            & Chr$(34) & = CARATTERE ANSI della virgoletta singola (') o apostrofo _
            DONE: inseriti questi codici _
        .................................................................................... _




'//QUESTA E' LA CLASSE OGGETTI DI GESTIONE DEGLI OGGETTI DEL DATABASE MDB _
        del progetto mdb            --->    GE_OGGETTI_DB_MDB.mdb _
        si trova in questa path     --->    c:\CASA\LINGUAGGI\ACCESS\ACCESS_PROGETTI_MDB\ACCESS_GESTIONE_OGGETTI_DB_MDB\GESTIONE_OGGETTI\GESTIONE_OGGETTI_MDB\
'//DENOMINAZIONE                    --->    CLASSE_GESTIONE_OGGETTI_DATABASE


'//CLASSI DIPENDENTI: _
        01) UTIL_Nrx1001_n01_CLASSE_OGGETTI_DELLA_FORM _
        02) UTIL_Nrx1003_n01_RECUPERO_PARAMETRI_TABELLA_OGGETTI


'//PROCEDURE
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'//LE PROCEDURE CONTENUTE: _

        
        '//**************************************************************************************************************************//
        '//SEZIONE_01
        '//**************************************************************************************************************************//
        '//                                                    CONTROLLO DELLE FORM
        '//Controlla_TUTTE_LE_FORM() _
                    '//01.03
                    '//PULISCO I CAMPI TEXT E LABEL DELLA FORM CORRENTE = LoopFormProps_PULISCI_OGGETTI_TEXT_LABEL_PSub
                    '//CODICE----------->LoopFormProps_PULISCI_OGGETTI_TEXT_LABEL_PSub
        
                    '//01.04
                    '//FUNZIONE DI CONTROLLO SULLA FORM SE E' APERTA O CHIUSA = Controlla_FormGESTIONE_APERTA(..)
                    '//CODICE ----------->Controlla_FormGESTIONE_APERTA.ATTIVA
                
                    '//01.05
                    '//PROVA CONTROLLO INSIEME FORMS = tra l'insieme form trova la form e al suo interno trova la label -
                    '//CODICE ----------->PROVA_CONTROLLO_INSIEME_FORMS.ATTIVA

                   '

        '//**************************************************************************************************************************//
        
        
        
        '//**************************************************************************************************************************//
        '//SEZIONE_02
        '//**************************************************************************************************************************//
        '//                SEZIONE RELATIVA ALLA GESTIONE DELLA TABELLA OGGETTI TMP CON LE SUE FUNZIONI E ROUTINE
        '//
                '//02.06
                '//FUNZIONE DI SALVATAGGIO DEL PARAMETRO DEL CODOGGETTO NELLA TABELLA TMP
                '//CODICE ----------->SALVATAGGIO_PARAMETRO_CODOGGETTO_in_TMP_pFunct.AttivoLaFunzione

                '//02.07
                '//FUNZIONE DI RECUPERO DEL CODICE OGGETTO = la funzione Pubblica recupera il oggetto dalla tabella parametri oggetti _
                                                            denominata GEST_Tb01_N01_PARAM_OGGETTI
                '//CODICE----------->RECUPERO_CODOGGETTO_pbFunct.AttivoLaFunzione


                '//02.08
                '//ROUTINE DI SALVATAGGIO ATTIVITA CASELLA CMB_02 = La routine salva nella tabella parametri l'attivita _
                                                                GEST_Tb01_N01_PARAM_OGGETTI. I parametri sono : _
                                                                par_sxCODOGGETTO As String, _
                                                                par_sxTIPOGGETTO As String, _
                                                                par_sxATTIVITAEVENTO As String, _
                                                                par_sxCODATTIVITA As String)

                '//CODICE----------->SALVO_ATTIVITA_CMB_02_IN_TMP_Sub.AttivoLaRoutine
                
                
                

                
                '//02.09 TODO : GEST_Tb01_N01_PARAM_OGGETTI
                '//ROUTINE DI SALVATAGGIO ATTIVITA CASELLA CMB_02 = La routine salva nella tabella parametri l'attivita _
                                                                GEST_Tb01_N01_PARAM_OGGETTI. I parametri sono : _
                                                                par_sxCODOGGETTO As String, _
                                                                par_sxTIPOGGETTO As String, _
                                                                par_sxATTIVITAEVENTO As String, _
                                                                par_sxCODATTIVITA As String)

                '//CODICE----------->SALVO_ATTIVITA_CMB_02_IN_TMP_Sub.AttivoLaRoutine
                

        '//...
        '//...
        '//**************************************************************************************************************************//
        
        
        '//CLASSE.GESTIONE.OGGETTI.DATABASE
        '//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++//
        '//
        '//SEZIONE.05                     UTILIZZO DEGLI OGGETTI DAO - CODICE---------->UtilizzoDegliOggettiDao.Titolo
        '//
        '//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++//
        '//
        '//             CLASSE CHE CONTIENE  LE PROCEDURE DI CONTROLLO DEGLI OGGETTI TABLE NEL DATABASE CORRENTE O ESTERNO
        '//
            
            
            '//TODO: CERCA OGGETTI ... da implementare
            '//06
            '//CONTROLLO_ESISTENZA_OGGETTI_TABELLA
            '//****************************************************************************************************************************
            '//06)_FUNZIONE_CONTROLLO_OGGETTI_TABELLA = permette il controllo dell'esistenza della tabella oggetto all'interno del db corrente _
                                                                    oppure in un database esterno.
            '//CODICE----------> pfRICERCA_NEL_DB_OGGETTI_Table.Funzione
            
            
            '//TODO: CERCA OGGETTI ... da implementare
            '//SEZIONE.05
            '//06.30
            '//FUNZIONE CERCA OGGETTI = Con due parametri Tipo oggetto e Nome oggetto lo ricerca nella db.
            '//CODICE-------->CERCA_OGGETTI.01.Attiva

            '//SEZIONE.05
            '//05.40
            '//FUNZIONE CONTROLLA CAMPI TABELLA = attivo il controllo sulla tabella e sul campo
            '//CODICE---------->CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct.attiva
                    '//DONE: Implementazione ricerca campi tabella

        '//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++//
    
         

            '...
        '//**************************************************************************************************************************//
        
        
        '//SEZIONE_04
            '//**************************************************************************************************************************//
            '//                     PROCEDURE DI SALVATAGGIO DATI NELLA TABELLE TMP
            '//
                
                '//TODO:da implementare
                '//40
                '//SALVATAGGIO DATI NELLA TABELLA TMP
                '//****************************************************************************************************************************
                '//ROUTINE DI SALVATAGGIO DEI DATI NELLA TABELLA TMP = _
                        Nella tabella <<GEST_Tb01_N01_PARAM_OGGETTI_TMP>> _
                    viene effettuato il salvataggio dei parametri. I parametri in entrata sono i seguenti :
                    'par_sxCODOGGETTO       = CODICE OGGETTO OBBLIGATORIO
                    'par_sxTIPOGGETTO       =
                    'par_sxATTIVITAEVENTO   =
                    'par_sxCODATTIVITA      = CODICE ATTIVITA OBBLIGATORIO
                    '//CODICE----------> SALVO_ATTIVITA_CMB_02_IN_TMP_Sub.Attiva
                
                
                '//41
                '//RECUPERO IL CODOGGETTO DALLA TMP = dalla tabella tmp mediante query preimpostata e restituisce _
                            una stringa che rappresenta il CODICE OGGETTO.
                '//CODICE----------> RECUPERO_CODOGGETTO_da_TMP_pbFunct.attiva

    
                '...
            '//**************************************************************************************************************************//
        
        
'//PROCEDURE
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@






Option Compare Database

Dim Bool1 As Boolean
Dim NameForm_s As String
Dim Msg_s   As String



'EVENTI CLASSE
'______________________________________________________________________________

    
'INIZIALIZZO LA CLASSE

Private Sub Class_Initialize()

    On Error GoTo Err_Class_Initialize

            '//SE MANCA LA LIBRERIA CARICARE c:\CASA\LINGUAGGI\ACCESS\DLL_LIBRERIE\MSCOMCTL.OCX
            'Reset Variabili Oggetti Form
            m_sxTIPOGGETTO = ""
            m_sxPROPRIETA = ""
            m_sxMETODO = ""
            m_sxEVENTO = ""
            
'USCITA ED ERRORI
'..............................................................
Exit_Class_Initialize:
    Exit Sub

Err_Class_Initialize:
    MsgBox Err.Description
    Resume Exit_Class_Initialize

                                                      
End Sub



'TERMINO LA CLASSE
Private Sub Class_Terminate()

    On Error GoTo Err_Class_Terminate
    
    Dim TITOLO_Get As String


'USCITA ED ERRORI
'..............................................................
Exit_Class_Terminate:
    Exit Sub

Err_Class_Terminate:
    MsgBox Err.Description
    Resume Exit_Class_Terminate


End Sub




'//01.05
'//PROVA CONTROLLO INSIEME FORMS = tra l'insieme form trova la form e al suo interno trova la label -
'//CODICE ----------->PROVA_CONTROLLO_INSIEME_FORMS.ATTIVA
'//===========================================================================================================================//
'//NOTE: eseguo il ciclo nell'insieme forms e nell'insieme dei controlli della form filtrata con il nome e condizionata _
dalla if di filtro. Cicolo all'interno delle proprieta delle form e resetto LE CASELLE DI TESTO E LE LABEL.



Sub PROVA_CONTROLLO_INSIEME_FORMS()
Dim obj As Object
Dim frm As Form
    Dim Ctrl As Control
    
    '//Controllo insieme form  STAMPO LA FORM APERTA FILTRATA
    For Each obj In CurrentProject.AllForms
        If InStr(obj.Name, "MODELLO_Frm01_M01_GE_FILTRO_M15_H15_L35_NoButton") Then Debug.Print "Forms|" & obj.Name
        'Set frm = Forms(obj.Name) '- this had problems as well, so I commented out
        For Each Ctrl In forms(obj.Name).Controls
            If InStr(Ctrl.Name, "LBL_FILTRO_01") Then Debug.Print "Form|" & obj.Name & "|" & Ctrl.Name
        Next Ctrl
    Next obj
    
End Sub

'//===========================================================================================================================//


'//**************************************************************************************************************************//
'//SEZIONE_01
'//**************************************************************************************************************************//
'//                                                    CONTROLLO DELLE FORM
'//

'//01.03
'//PULISCO I CAMPI TEXT E LABEL DELLA FORM CORRENTE -
'//CODICE ----------->LoopFormProps_PULISCI_OGGETTI_TEXT_LABEL_PSub.ATTIVA
'//===========================================================================================================================//
'//NOTE: eseguo il ciclo nell'insieme forms e nell'insieme dei controlli della form filtrata con il nome e condizionata _
        dalla if di filtro. Cicolo all'interno delle proprieta delle form e resetto LE CASELLE DI TESTO E LE LABEL.
Public Sub LoopFormProps_PULISCI_OGGETTI_TEXT_LABEL_PSub(par_NameForm_s As String)
    On Error GoTo Err_Handler
    Dim frmLoop As Object
    Dim Ctrl As Control
    Dim propLoop As Property
    Dim strFormName As String
    
    '//La variabile BOOLEAN per indicare se esiste
    Dim Bool1 As Boolean
    Dim NameForm_s As String
    Dim Msg_s  As String
    
    
    '//CHIAMO LA ROUTINE DI PULITURA DEI CAMPI TEXT E LABEL
    '//--------------------------------------------------------------------------------//-//
    '//CODICE----------->LoopFormProps_PULISCI_OGGETTI_TEXT_LABEL_PSub.puliscoTutto
    
        '//ATTIVO LA ROUTINE Call LoopFormProps_PULISCI_OGGETTI_TEXT_LABEL_PSub
    
    '//--------------------------------------------------------------------------------//-//
    
    For Each frmLoop In CurrentProject.AllForms
        
        '//imposto la variabile bool
        Bool1 = False
        '//la form da ricercare nel db
        'NameForm_s = "MOD30_Frm01_S01_ELENCO_SCH01_65txt_campi_1key"
        
        NameForm_s = par_NameForm_s
        
        '//TROVA QUELLA CORRENTE
        'If frmLoop.NAME = "MOD30_Frm01_S01_ELENCO_SCH01_65txt_campi_1key" Then
        If frmLoop.Name = par_NameForm_s Then
      
        '//LA FORM CORRENTE
        '//=====================================================================================================//
                
                '//form trovata imposto il nome e  la variabile
                strFormName = frmLoop.Name
                '//imposto la variabile bool a true per individuare che è stata trovata
                Bool1 = True
                
                
                '//APRO LA FORM IN MODALITA DESIGN
                DoCmd.OpenForm strFormName, acDesign, , , , acHidden
                
                '//Qualifico l'oggetto form pe le proprieta
                With forms(strFormName)
                    
                    For Each Ctrl In .Controls
                        
                        DoEvents
                         
                        For Each propLoop In Ctrl.Properties
                        
                             If Ctrl.Name = "TXT_01" Then
                                'Stop
                             End If
                                
                            
                        '//PULISCO I CAMPI TEXT E LABEL DELLA FORM CORRENTE - CICOLO FOR EACH -
                        '//==================================================================================================//
                        '//NOTE : pulisco le caselle di testo impostando la proprieta recordsource a null mentre le label _
                                corrispondenti vengono impostate al nome oggetto label corrente.
                        '//CODICE----------->LoopFormProps_PULISCI_OGGETTI_TEXT_LABEL_PSub.puliscoTutto
                        
                        
                            
                            '//PULISCO LA PROPRIETA RECORD SOURCE DELLE CASELLE DI TESTO
                            '//--------------------------------------------------------------------------------//------------//
                            '//NELLA CASELLE TEXT riconosciute con il valore ControlSource _
                            viene azzerata la proprieta control source con valore ""
                                If propLoop.Name = "ControlSource" Then
                                    Debug.Print propLoop.Value
                                    propLoop.Value = ""
                                End If
                                
                            '//--------------------------------------------------------------------------------//------------//
                            
                                                    
                            
                            '//PULISCO LA PROPRIETA CAPTION DELLE LABEL
                            '//--------------------------------------------------------------------------------//------------//
                            '//se è una label riconosciuta con il nome della proprieta _
                            Caption, viene sostituita il valore del contenuto della _
                            etichetta con il nome della label
                                 If propLoop.Name = "Caption" Then
                                            Debug.Print propLoop.Value
                                            '//inserisco il nome label?
                                            propLoop.Value = Ctrl.Name
                                                    
                                                    
                                End If
                                
                            '//--------------------------------------------------------------------------------//------------//
                        
                        '//*** FINE ***
                        '//PULISCO I CAMPI TEXT E LABEL DELLA FORM CORRENTE - CICOLO FOR EACH -
                        '//==================================================================================================//
                             
                            '//CONTROLLO oggetto
                            Debug.Print "Form [" & .Name & "] - Control[" & Ctrl.Name & "](" & propLoop.Name & "):" & propLoop.Value
                            
                        Next '//For Each propLoop In ctrl.Properties
                    
                    
                    Next    '//For Each ctrl In .Controls
                
                            '//SALVATAGGIO FORM APERTA IN DESIGN
                            'Richiama la funzione che chiude la maschera aperta
                            DoCmd.SetWarnings False
                            DoCmd.Close acForm, strFormName, acSaveYes
                            DoCmd.SetWarnings True
                            DoCmd.Close acForm, strFormName

                End With '//With Forms(strFormName)
                
                               
        '//=====================================================================================================//
        
                
            
        End If
    Next
    
                '//messaggio di avviso se Bool1=false e cioè non esiste la form
                
                If Bool1 = False Then
                    '//imposto il testo del messaggio
                    Msg_s = "ATTENZIONE LA FORM " & NameForm_s & " NON E' STATA TROVATA IN QUESTO DATABASE!"
                    MsgBox Msg_s, vbExclamation
                End If
    
    
    
 
    Exit Sub
Err_Handler:
    Select Case Err.number
        Case 2186
            '//SE LA PROPRIETA NON ESISTE ATTIVO L'ERRORE E RIENTRO NEL CICLO FOR EACH
            'Property not available in design view  SE NON CI SONO LE  PROPRIETA RIPROVA.
            Resume Next
        Case Else
            MsgBox "Err: " & Err.number & " - " & Err.Description
    End Select
End Sub
 
'//*** fine ***
'//PULISCO I CAMPI TEXT E LABEL DELLA FORM CORRENTE -
'//===========================================================================================================================//





'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'//              MODULO ---> UTIL_Nro77_CONTROLLO_FORM_APERTA
'//
'//NOTE         :
'//             Questa funzione controlla se la form parent è aperta. Esegue un ciclo negli oggetti form ed individua se
'//             la form passata come parametro è aperta. Rilascia TRUE se la trova aperta e quindi permette la modifica
'//             DEL CAMPO che si trova nella form superiore altrimenti no.
'//             Il modulo in cui si trova il prototipo è il seguente: MODULO ------------> UTIL_Nro77_CONTROLLO_FORM_APERTA
'//
'//PARAMETRI    :
'//             par_sxNameForm As String = "nome form aperta"
'//
'//RESTITUISCE  :
'//             As Boolean = True se aperta oppure False se la form è chiusa
'//
'//CLASSE       :
'//             la classe ospitante è la seguente : CLASSE_GESTIONE_OGGETTI_DATABASE
'//             CODICE_CLASSE_Ospitante_}---------------@UTIL_Nro77_CONTROLLO_FORM_APERTA.CHIAMA_Controlla_FormGESTIONE_APERTA
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'//
'//01.04
'//FUNZIONE_DI_CONTROLLO_DELLA_FORM_GESTIONE_APERTA (FORM APERTA CONTROLLO)
'//CHIAMA_Controlla_FormGESTIONE_APERTA.attiva
'//=======================================================================================================//
Private Sub CHIAMA_Controlla_FormGESTIONE_APERTA()
    Dim Str1 As String
    Str1 = "CUP_Tb01_Frm_M01_GESTIONE_CUP_SCHEDE_H15_L35_NoButton"
    Bool1 = False
    Bool1 = Controlla_FormGESTIONE_APERTA(Str1)
    
    
    If Bool1 = True Then
    
        MsgBox "La form " & Str1 & " è  APERTA", vbExclamation, "CONTROLLO FORM APERTA"
    
    End If
    
    
    '// CONTROLLA SE LA FORM GESTIONE E' APERTA
    '//============================================================================================//
    '//NOTE : chiamo la funzione per il controllo se la form principale  è aperta _
    viene passato come parametro il nome della form MASTER e restituisce un Boolean _
    TRUE se la form è aperta. @CONTROLLO.FORM.APERTA.CLASSE.GestioneOggettiDb
    
        '//STR1 = Controlla_FormGESTIONE_APERTA()
         '//    Str1 = "CUP_Tb01_Frm_M01_GESTIONE_CUP_SCHEDE_H15_L35_NoButton"
        '//     Bool1 = False
        '//     Bool1 = Controlla_FormGESTIONE_APERTA(Str1)
   
    '//============================================================================================//
  
End Sub
Public Function Controlla_FormGESTIONE_APERTA(par_sxNameForm As String) As Boolean
    
     
    Dim formName_s As String
    Dim ctl As Control
        
    '//reset
    FormGestione_b = False
    
        
    'application.Form[GEST_FrmM00_GESTIONE].
    Dim obj As AccessObject, Dbs As Object
    Set Dbs = Application.CurrentProject
    ' Search for open AccessObject objects in AllForms collection.
    For Each obj In Dbs.AllForms
        If obj.IsLoaded = True Then
            ' Print name of obj.
            Debug.Print obj.Name
            formName_s = obj.Name
             
            '//CONTROLLO SUL NOME
            If (obj.Name = par_sxNameForm) And obj.IsLoaded = True Then
                
               '//FormGestione_b = true per indicare che la form è aperta e restituisco il nome della form _
               alla routine chiamante
               FormGestione_b = True
               Controlla_FormGESTIONE_APERTA = True
               
            End If
            
        End If
    Next obj
    
        Set obj = Nothing
        Set Dbs = Nothing
        Set ctl = Nothing
        formName_s = vbNull
        
    
    

End Function
'//*** FINE ***
'//FUNZIONE_DI_CONTROLLO_DELLA_FORM_GESTIONE_APERTA
'//=======================================================================================================//


'//*** fine ***
'//CONTROLLO FORM APERTA -
'//===========================================================================================================================//






'//**************************************************************************************************************************//
'//SEZIONE_02
'//**************************************************************************************************************************//
'//                SEZIONE RELATIVA ALLA GESTIONE DELLA TABELLA OGGETTI TMP CON LE SUE FUNZIONI E ROUTINE
'//
        '//TODO:da implementare
        '//02.06
        '//SALVATAGGIO_PARAMETRO_CODOGGETTO_in_TMP_pbFunctsxCODOGGETTO = FUNZIONE DI SALVATAGGIO DEL PARAMETRO DEL CODOGGETTO NELLA TABELLA TMP








'//**************************************************************************************************************************//
'//SEZIONE_02           *** FINE ***
'//**************************************************************************************************************************//








'//SEZIONE_03
'//**************************************************************************************************************************//
'//                     PROCEDURE DI CONTROLLO OGGETTI TABLE NEL DATABASE CORRENTE O ESTERNO
'//                     CONTROLLO OGGETTI DEL DABASE
'//06.30
'//FUNZIONE CERCA OGGETTI = Con due parametri Tipo oggetto e Nome oggetto lo ricerca nella db.
'//CODICE-------->CERCA_OGGETTI.01.Attiva
'//
'//**************************************************************************************************************************//


'//06.30
'//FUNZIONE CERCA OGGETTI = Con due parametri Tipo oggetto e Nome oggetto lo ricerca nella db.
'//CODICE-------->CERCA_OGGETTI.01.Attiva
'//=======================================================================================================//
'//MODULO   : UTIL_MDL40_N01_CONTROLLO_OGGETTI_DB _
            modulo dove si trova la funzione del database .mdb _
            PATH -> c:\CASA\LINGUAGGI\ACCESS\ACCESS_PROGETTI_MDB\ACCESS_GESTIONE_OGGETTI_DB_MDB\GESTIONE_OGGETTI\GESTIONE_OGGETTI_MDB\GE_OGGETTI_DB_MDB.mdb
'//NOTE     : la funzione cerca oggetti attiva il comando ricerca oggetto nel database corrente _
            con il primo parametro TIPO OGGETTO filtra gli oggetti query, tabella ecc... _
            e con il secondo parametro NOMEOGGETTO attiva la ricerca con il comando SelectObject _
            ma in caso di fallimento il messaggio di errore avvisa che non ci sono oggetti nel db corrente.
Public Function CERCA_OGGETTI(par_sxTIPOGGETTO As String, par_sxNOMEOGGETTO As String)

On Error GoTo CERCA_OGGETTI_Err


        'RICERCA OGGETTO @CLASSE.CERCA.OGGETTO
        '-------------------------------------------------------------------------------------
                Select Case par_sxTIPOGGETTO
                
                Case "TABLE"
                  ' RICERCA LE TABELLE
                  DoCmd.SelectObject acTable, par_sxNOMEOGGETTO, True
                  Exit Function
                  
                
                Case "QUERY"
                    ' RICERCA LE QUERY
                    DoCmd.SelectObject acQuery, par_sxNOMEOGGETTO, True
                    Exit Function
                
                
                Case "FORMS"
                  ' RICERCA LE FORM
                  DoCmd.SelectObject acForm, par_sxNOMEOGGETTO, True
                  Exit Function
            
            
                Case "REPORTS"
                  ' RICERCA REPORT
                  DoCmd.SelectObject acReport, par_sxNOMEOGGETTO, True
                  Exit Function
            
                
                Case "Pagine"
                  ' RICERCA LE PAGINE INTERNET
                  DoCmd.SelectObject acPage, par_sxNOMEOGGETTO, True
                  Exit Function
                
                
                Case "MACRO"
                  ' RICERCA LE MACRO
                  DoCmd.SelectObject acMacro, par_sxNOMEOGGETTO, True
                  Exit Function
                
                
                
                Case "MODULES"
                  ' RICERCA DEI MODULI
                  DoCmd.SelectObject acModule, par_sxNOMEOGGETTO, True
                  Exit Function
                
                
                Case Else
                    MsgBox "NON CI SONO OGGETTI DA RICERCARE", vbCritical
                    
                End Select

        '-------------------------------------------------------------------------------------


CERCA_OGGETTI_Exit:
    Exit Function

CERCA_OGGETTI_Err:
    MsgBox Error$
    Resume CERCA_OGGETTI_Exit

End Function
'//=======================================================================================================//




'//**************************************************************************************************************************//
'//SEZIONE_03                                CONTROLLO OGGETTI DEL DABASE *** FINE ***
'//
'//**************************************************************************************************************************//














   
'//**************************************************************************************************************************//
'//SEZIONE_02
'//**************************************************************************************************************************//
'//                SEZIONE RELATIVA ALLA GESTIONE DELLA TABELLA OGGETTI TMP CON LE SUE FUNZIONI E ROUTINE
'//
'//                                 TABELLA OGGETTI TMP FUNZIONI E ROUTINE
'//le funzioni e le routine applicate sulla tabella tmp
'//02.06
'//FUNZIONE DI SALVATAGGIO DEL PARAMETRO DEL CODOGGETTO NELLA TABELLA TMP
'//CODICE ----------->SALVATAGGIO_PARAMETRO_CODOGGETTO_in_TMP_pFunct.AttivoLaFunzione
'//
'//--------------------------------------------------------------------------------//------------------------------------------------//
'NOTE: funzione Pubblica di salvataggio del codice oggetto nella tabella tmp parametri oggetti TMP
'//MODULO DI PARTENZA = UTIL_Nro61_N01_SALVATAGGIO_PARAMETRO_CODICE_OGGETTO
Public Function SALVATAGGIO_PARAMETRO_CODOGGETTO_in_TMP_pFunct(par_sxCODOGGETTO) As String
                


On Error GoTo Err_SALVATAGGIO_PARAMETRO_CODOGGETTO_in_TMP_pFunct

        'Controllo parametro passato
        If par_sxCODOGGETTO = "" Then
            
            MsgBox ("Funzione : SALVATAGGIO_PARAMETRO_CODOGGETTO_in_TMP_pFunct - Parametro Codice Oggetto = null ? - Uscita dalla funzione")
            
            Exit Function
        
        End If
        
        

'//MODELLO_ATTIVAZIONE_CLASSE
'//CHIAMO LA ROUTINE SALVATAGGIO_PARAMETRO_CODOGGETTO_in_TMP_pFunctsxCODOGGETTO DELLA CLASSE OGGETTI FORM
'//--------------------------------------------------------------------------------//------------------------------------------------//


        
        
        'SALVO IL PARAMETRO NELLA TABELLA OGGETTI TMP
        'solo se esiste un codice oggetto.
        sSql = ""
        sSql = "UPDATE GEST_Tb01_N01_PARAM_OGGETTI_TMP SET GEST_Tb01_N01_PARAM_OGGETTI_TMP.CODOGGETTO_s = '" & par_sxCODOGGETTO & "';"
        
        'controllo
        Debug.Print sSql
        CurrentDb.Execute (sSql)
        
        
                SALVATAGGIO_PARAMETRO_CODOGGETTO_in_TMP_pFunct = "Salvataggio Parametro Eseguito"
                
                
                
        
'USCITA  E GESTIONE ERRORI
'..............................................................................................................

Exit_SALVATAGGIO_PARAMETRO_CODOGGETTO_in_TMP_pFunct:
        Exit Function
    
Err_SALVATAGGIO_PARAMETRO_CODOGGETTO_in_TMP_pFunct:
    MsgBox Err.Description
    Resume Exit_SALVATAGGIO_PARAMETRO_CODOGGETTO_in_TMP_pFunct
    
                
End Function






'//02.07
'//FUNZIONE DI RECUPERO DEL CODICE OGGETTO = la funzione Pubblica recupera il oggetto dalla tabella parametri oggetti _
                                            denominata GEST_Tb01_N01_PARAM_OGGETTI
'//CODICE----------->RECUPERO_CODOGGETTO_pbFunct.AttivoLaFunzione
Public Function RECUPERO_CODOGGETTO_pbFunct(par_sxCODATTIVITA As String) As String



On Error GoTo Err_RECUPERO_CODOGGETTO_pbFunct



     


        'RECUPERO PARAMETRO DA TABELLA OGGETTI
        sSql = ""
        sSql = sSql & "SELECT GEST_Tb01_N01_PARAM_OGGETTI.*, GEST_Tb01_N01_PARAM_OGGETTI.CODATTIVITA "
        sSql = sSql & "FROM GEST_Tb01_N01_PARAM_OGGETTI "
        sSql = sSql & "WHERE (((GEST_Tb01_N01_PARAM_OGGETTI.CODATTIVITA)='" & par_sxCODATTIVITA & "'));"
        
        'controllo
        Debug.Print sSql
        Set DaoRs = CurrentDb.OpenRecordset(sSql)
        
        If DaoRs.EOF = False And DaoRs.BOF = False Then
            
            DaoRs.MoveFirst
                'leggo il valore e lo assegno
                sxCODOGGETTO = DaoRs.Fields("CODOGGETTO")
                
            DaoRs.Close
            Set DaoRs = Nothing
        
        End If
        
                RECUPERO_CODOGGETTO_pbFunct = sxCODOGGETTO
                
                
                
        
'USCITA  E GESTIONE ERRORI
'..............................................................................................................

Exit_RECUPERO_CODOGGETTO_pbFunct:
    Exit Function

Err_RECUPERO_CODOGGETTO_pbFunct:
    MsgBox Err.Description
    Resume Exit_RECUPERO_CODOGGETTO_pbFunct

                
End Function



'//02.08
'//ROUTINE DI SALVATAGGIO ATTIVITA CASELLA CMB_02 = La routine salva nella tabella parametri l'attivita _
                                                GEST_Tb01_N01_PARAM_OGGETTI. I parametri sono : _
                                                par_sxCODOGGETTO As String, _
                                                par_sxTIPOGGETTO As String, _
                                                par_sxATTIVITAEVENTO As String, _
                                                par_sxCODATTIVITA As String)
'//CODICE----------->SALVO_ATTIVITA_CMB_02_IN_TMP_Sub.AttivoLaRoutine
'//
'SALVATAGGIO PARAMETRI NELLA TABELLA TMP
'Note: Nella tabella <<GEST_Tb01_N01_PARAM_OGGETTI_TMP>>
'viene effettuato il salvataggio dei parametri.
'
Public Sub SALVO_ATTIVITA_CMB_02_IN_TMP_Sub(par_sxCODOGGETTO As String, _
                                            par_sxTIPOGGETTO As String, _
                                            par_sxATTIVITAEVENTO As String, _
                                            par_sxCODATTIVITA As String)


                                        
                                 
On Error GoTo Err_SALVO_ATTIVITA_CMB_02_IN_TMP_Sub


            If par_sxCODOGGETTO > "" And par_sxCODATTIVITA > "" Then
                    
                    'SALVATAGGIO ATTIVITA NELLA TABELLA GEST_Tb01_N01_PARAM_OGGETTI_TMP
                    '___________________________________________________________________
                        'Stringa ssql
                        'di salvataggio nella tabella tmp della attivita da svolgere
                        sSql = ""
                        sSql = sSql & "UPDATE GEST_Tb01_N01_PARAM_OGGETTI_TMP "
                        sSql = sSql & "SET GEST_Tb01_N01_PARAM_OGGETTI_TMP.CODATTIVITA = '" & par_sxCODATTIVITA & "', "
                                                'sSql = sSql & "SET GEST_Tb01_N01_PARAM_OGGETTI_TMP.CODATTIVITA = '" & sxCODATTIVITA & "'; "
                        sSql = sSql & "GEST_Tb01_N01_PARAM_OGGETTI_TMP.CODOGGETTO_s = '" & par_sxCODOGGETTO & "'; "

                        
                        'controlo
                        Debug.Print sSql
                        
                        'comando eseguito
                        CurrentDb.Execute (sSql)
                    '___________________________________________________________________
                    
            Else
                    MsgBox "I PARAMETRI CODICE ATTIVIA E CODICE OGGETTO NON SONO VALORIZZATI : par_sxCODATTIVITA --> " & par_sxCODATTIVITA & _
                                                                    "  par_sxCODOGGETTO --> " & par_sxCODOGGETTO
            End If
        
'USCITA  E GESTIONE ERRORIGEST_Tb01_N01_PARAM_OGGETTI
'..............................................................................................................

Exit_SALVO_ATTIVITA_CMB_02_IN_TMP_Sub:
    Exit Sub

Err_SALVO_ATTIVITA_CMB_02_IN_TMP_Sub:
    MsgBox Err.Description
    Resume Exit_SALVO_ATTIVITA_CMB_02_IN_TMP_Sub

End Sub




'//02.09 TODO :
'//FUNZIONE TROVA LA CLASSE DI GESTIONE = Trova il codice oggetto relativo alla classe di gestione _
                                        mediante il parametro stringa.
'NOTE:
'Operazione : Trova il codice oggetto relativo alla classe di gestione.
'IMPUT      :Riceve come parametro una stringa che permette l'individuazione
'l'individuazione della classe da richiamare.
'//RESTITUISCE : il codice oggetto
'//MODULO ---> UTIL_Nro10_N03_TROVA_CLASSE_DI_GESTIONE_pFunct

Public Function TROVA_CLASSE_DI_GESTIONE_pFunct(par_sxItemText As String) As String

On Error GoTo TROVA_CLASSE_DI_GESTIONE_pFunct_Err


'IMPOSTAZIONE DELLE PROPRIETA
'-------------------------------------------------------------------------------------------
                
'-------------------------------------------------------------------------------------------
                
                
                  
                     
             
           
                
                
       'LETTURA OGGETTI DELLA CLASSE
       '==================================================================================================
       'Note: Leggo i dati dalla tabella OGGETTI


    
    
                'CLASSE DI GESTIONE DEL LOTTO
                '...............................................................
                        
                        
                        If par_sxItemText > "" Then
                            
                             'APRO IL RS
                                sSql = ""
                        
                
                                sSql = sSql & "SELECT GEST_Tb01_N01_PARAM_OGGETTI.* "
                                sSql = sSql & "FROM GEST_Tb01_N01_PARAM_OGGETTI "
                                sSql = sSql & "WHERE (((GEST_Tb01_N01_PARAM_OGGETTI.CODATTIVITA)='" & par_sxItemText & "'));"
                                
                            
                                   'Controllo
                                   Debug.Print sSql
                
                
                'Apro rs estrazione Sql
                Set DaoRs = CurrentDb.OpenRecordset(sSql)
                    
                      
                      'RECUPERA I DATI DALLA TABELLA OGGETTI
                      '_________________________________________________________________
                      'Le proprieta del controllo a schede :
                      '               - Visible
                      '               - Enabled
                      
                        If DaoRs.EOF = False And DaoRs.BOF = False Then
                            
                            DaoRs.MoveFirst
                      
                                  'Imposto la variabile della funzione
                                  '..........................................................
                                  
                                        'controllo oggetto
                                        Debug.Print DaoRs.Fields("CODOGGETTO_s").Value
                                        
                                        TROVA_CLASSE_DI_GESTIONE_pFunct = DaoRs.Fields("CODOGGETTO_s").Value
                            
                                  '..........................................................
                                
                                    
                        End If
                                         
                      '_________________________________________________________________
                    
                    
    
                
                End If
        
    
                            
                       
        '==================================================================================================
        
        
                
'EXIT E GESTIONE ERRORI
'-----------------------------------------------------------------------------------------------
        
        

TROVA_CLASSE_DI_GESTIONE_pFunct_Exit:
    Exit Function

TROVA_CLASSE_DI_GESTIONE_pFunct_Err:
    MsgBox Error$
    Resume TROVA_CLASSE_DI_GESTIONE_pFunct_Exit


End Function





'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'                                   LE FUNZIONI
'                                   DELLA FORM
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<




'RECUPERO SQL DI SALVATAGGIO DEI PARAMETRI TMP
'Nota   : Recupero una stringa ssql da attivare all'evento activate del campo 1 della form corrente
'per il salvataggio dei parametri.

Private Static Function RECUPERO_SQL_SALVATAGGIO_PARAMETRI_TMP_Func() As Variant

On Error GoTo Err_RECUPERO_SQL_SALVATAGGIO_PARAMETRI_TMP_Func



    
      
                      'RECUPERA LA STRINGA SSQL
                      '_________________________________________________________________
                      'Nota: Il valore del campo VALOREPROPRIETA, contiene la stringa
                      'sql da utilizzare per il salvataggio
                      'al flag di numerazione o meno dei campi
                            
                                         
                        'ISTANZIO LA CLASSE
                        'Assegno e creo l'oggetto di classe
                        Set Cls_RECUPERO_PARAMETRI_TABELLA_OGGETTI = New UTIL_Nrx1003_n01_RECUPERO_PARAMETRI_TABELLA_OGGETTI
                      
                        'recupero i parametri mediante i metodi della classe
                        Call Cls_RECUPERO_PARAMETRI_TABELLA_OGGETTI.RECUPERO_PARAMETRI_TABELLA_OGGETTI_pFc(sxCODOGGETTO, _
                                                                                                           sxTIPOGGETTO, _
                                                                                                           sxPROPRIETA, _
                                                                                                           vxVALOREPROPRIETA, _
                                                                                                           sxEVENTO, _
                                                                                                           sxATTIVITAEVENTO, _
                                                                                                           sxMETODO, _
                                                                                                           sxCODATTIVITA, _
                                                                                                           ixNROGGETTO, _
                                                                                                           lngxID)
                        
                       
                      
                    
                                  
                                  'RECUPERO LA STRINGA SSQL
                                  '..........................................................
                                  'Nota : stringa ssql passata al chiamante
                                  '
                                    
                                    'restituisco al chiamante
                                    sxQuerySalvataggio_Parametri = Cls_RECUPERO_PARAMETRI_TABELLA_OGGETTI.QuerySalvataggio_Parametri_Tmp_sxGET
                                  '..........................................................
                                                                    
                                    'RECUPERO IL NUMERO DEI PARAMETRI
                                    '..........................................................
                                    'Nota : recupero dal metodo della classe il numero dei parametri necessari
                                    'per utilizzare la query parametrica
                                    ixNRO_PARAMETRI = Cls_RECUPERO_PARAMETRI_TABELLA_OGGETTI.Nro_Parametri_ixGET
                                                                  
                                  
                                    'RECUPERO TIPO PARAMETRO_01
                                    '..........................................................
                                    'Nota : recupero dal metodo della classe il tipo di parametro_01
                                    'da utilizzare nella query di aggiornamento: es. Tipo = Integer, Tipo = String ecc..
                                    'restituisco al chiamante
                                    vxTIPO_PARAMETRO_01 = Cls_RECUPERO_PARAMETRI_TABELLA_OGGETTI.vxTIPO_PARAMETRO_01_Tmp_sxGET
                                  
                                    
                                    'RECUPERO TIPO PARAMETRO_02
                                    '..........................................................
                                    'Nota : recupero dal metodo della classe il tipo di parametro_02
                                    'da utilizzare nella query di aggiornamento: es. Tipo = Integer, Tipo = String ecc..
                                    'restituisco al chiamante
                                    vxTIPO_PARAMETRO_02 = Cls_RECUPERO_PARAMETRI_TABELLA_OGGETTI.vxTIPO_PARAMETRO_02_Tmp_sxGET
                                    
                                    'RECUPERO TIPO PARAMETRO_03
                                    '..........................................................
                                    'Nota : recupero dal metodo della classe il tipo di parametro_03
                                    'da utilizzare nella query di aggiornamento: es. Tipo = Integer, Tipo = String ecc..
                                    'restituisco al chiamante
                                    vxTIPO_PARAMETRO_03 = Cls_RECUPERO_PARAMETRI_TABELLA_OGGETTI.vxTIPO_PARAMETRO_03_Tmp_sxGET
                                    
                                    
                                    'RECUPERO TIPO PARAMETRO_04
                                    '..........................................................
                                    'Nota : recupero dal metodo della classe il tipo di parametro_04
                                    'da utilizzare nella query di aggiornamento: es. Tipo = Integer, Tipo = String ecc..
                                    'restituisco al chiamante
                                    vxTIPO_PARAMETRO_04 = Cls_RECUPERO_PARAMETRI_TABELLA_OGGETTI.vxTIPO_PARAMETRO_04_Tmp_sxGET
                                    
                                    
                                    'RECUPERO TIPO PARAMETRO_05
                                    '..........................................................
                                    'Nota : recupero dal metodo della classe il tipo di parametro_05
                                    'da utilizzare nella query di aggiornamento: es. Tipo = Integer, Tipo = String ecc..
                                    'restituisco al chiamante
                                    vxTIPO_PARAMETRO_05 = Cls_RECUPERO_PARAMETRI_TABELLA_OGGETTI.vxTIPO_PARAMETRO_05_Tmp_sxGET


                                  
                                                                    
                                  
                                  
                                  '..........................................................
                                                                    
                                    
                            'chiudo l'oggetto
                            Set Cls_RECUPERO_PARAMETRI_TABELLA_OGGETTI = Nothing
                                         
                      '_________________________________________________________________
                    
                
    


'USCITA  E GESTIONE ERRORI
'..............................................................................................................

Exit_RECUPERO_SQL_SALVATAGGIO_PARAMETRI_TMP_Func:
    Exit Function

Err_RECUPERO_SQL_SALVATAGGIO_PARAMETRI_TMP_Func:
    MsgBox Err.Description
    Resume Exit_RECUPERO_SQL_SALVATAGGIO_PARAMETRI_TMP_Func


End Function
'//=============================================================================================================//




'//RECUPERO IL CODOGGETTO DALLA TMP = dalla tabella tmp mediante query preimpostata e restituisce _
                            una stringa che rappresenta il CODICE OGGETTO.
'//CODICE----------> RECUPERO_CODOGGETTO_da_TMP_pbFunct.attiva
'funzione Pubblica di recupero oggetto dalla tabella parametri oggetti TMP
'//MODULO = si trova nel UTIL_Nro60_N07_RECUPERO_PARAMETRI_da_TMP_Function
'//=============================================================================================================//
Public Function RECUPERO_CODOGGETTO_da_TMP_pbFunct() As String



On Error GoTo Err_RECUPERO_CODOGGETTO_da_TMP_pbFunct


        'RECUPERO PARAMETRO DA TABELLA OGGETTI TMP
        'solo se esiste un codice oggetto.
        sSql = ""
        sSql = sSql & "SELECT GEST_Tb01_N01_PARAM_OGGETTI_tmp.* "

        sSql = sSql & "FROM GEST_Tb01_N01_PARAM_OGGETTI_tmp "
        sSql = sSql & "WHERE (((GEST_Tb01_N01_PARAM_OGGETTI_tmp.CODOGGETTO)Is Not Null));"

        
        
        
        
        'controllo
        Debug.Print sSql
        Set DaoRs = CurrentDb.OpenRecordset(sSql)
        
        If DaoRs.EOF = False And DaoRs.BOF = False Then
            
            DaoRs.MoveFirst
                'leggo il valore e lo assegno
                sxCODOGGETTO = DaoRs.Fields("CODOGGETTO_s")
                
            DaoRs.Close
            Set DaoRs = Nothing
        
        End If
        
                RECUPERO_CODOGGETTO_da_TMP_pbFunct = sxCODOGGETTO
                
                
                
        
'USCITA  E GESTIONE ERRORI
'..............................................................................................................

Exit_RECUPERO_CODOGGETTO_da_TMP_pbFunct:
    Exit Function

Err_RECUPERO_CODOGGETTO_da_TMP_pbFunct:
    MsgBox Err.Description
    Resume Exit_RECUPERO_CODOGGETTO_da_TMP_pbFunct

                
End Function
'//=============================================================================================================//





'RECUPERO TIPO PARAMETRO 01
'Per il salvataggio del dato nella tabella parametri tmp
'INPUT          :
'RESTITUISCE    : Un dato variant, il quale rappresente il tipo di parametro (es. Integer, String, Long. ecc.)
'che la query di aggiornamento deve utilizzare.

Private Function Funzione_da_utilizzare_Funct() As Variant

On Error GoTo Err_TROVA_CODOGGETTO_Funct


    



'USCITA  E GESTIONE ERRORI
'..............................................................................................................

Exit_TROVA_CODOGGETTO_Funct:
    Exit Function

Err_TROVA_CODOGGETTO_Funct:
    MsgBox Err.Description
    Resume Exit_TROVA_CODOGGETTO_Funct

End Function
'//==============================================================================================================================//


'//ZZFINE_CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct()
'
'//SEZIONE.05                     UTILIZZO DEGLI OGGETTI DAO - CODICE---------->UtilizzoDegliOggettiDao.Titolo
'//05.40
'//FUNZIONE CONTROLLA CAMPI TABELLA = attivo il controllo sulla tabella e sul campo
'//CODICE---------->CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct.attiva
'//CLASSE_CHE_LA_IMPLEMENTA:CODICE---------->CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct.attiva
'//==============================================================================================================================//
Public Function CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct(par_NameTable_s As String, par_NameField_s As String) As String
'//MODULO:                  La funzione si trova nel modulo UTIL_MDL40_N05_CONTROLLO_OGGETTI_DB_TABLES del database GE_OGGETTI_DB_MDO.mdb in _
                            questa path -> c:\CASA\LINGUAGGI\ACCESS\ACCESS_PROGETTI_MDB\ACCESS_GESTIONE_OGGETTI_DB_MDB\GESTIONE_OGGETTI\GESTIONE_OGGETTI_MDB\
'//CHIAVI DI RICERCA        : Trova campi tabella e tipo di campo, ITERA NELLA TABELLA, Campo Tabella, Field type _
                              itera campi tabella, trova campo, trova field, itera oggetti field table, _
                              trova type field, ciclo field table, ciclo campi tabella access, _
                              esempio itera campi tabella, esempio ciclo campi tabella. _
                                                ******* _
                                            IN QUESTA CLASSE ulteriori chiavi di ricerca : _
                                                    FUNZIONE DELLA CLASSE RICERCA TIPO DI CAMPI TABELLA, FUNZIONE TIPO FIELD


On Error GoTo CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct_Err





Dim Msg As String ' For error handling.
Dim ws As Workspace, db As Database
Dim tblDef As TableDef, fldDef As Field
Dim i As Integer, Handle As Integer
Dim fldName As String, fldDataInfo As String
Dim Size_s As String
Dim TotField_i As Integer
    
         ' -----------------------------------------------
            Set db = CurrentDb()
            Str1 = par_NameTable_s
         ' -----------------------------------------------
       
    
'//imposto l'oggetto tabella con il nome della tabella passata come parametro
Set tblDef = db.TableDefs(Str1)
         With tblDef
                '//reset variabili
                    TotField_i = .Fields.count - 1              '//calcolo il totale dei field della tabella
            For i = 0 To .Fields.count - 1
               Set fldDef = .Fields(i)
               
                '//qualifico l'oggetto
                With fldDef
                        fldName = .Name
                    
                        '//CONTROLLO FIELD
                        Debug.Print i & ") NOME CAMPO -> " & fldName
                
                        '//effettuo il controllo tra il campo della tabella e quello della ricerca _
                          se trovato il campo di ricerca trovo il tipo e restituisco il timpo di _
                          campo con l'enventuale grandezza
                        If fldName = par_NameField_s Then
                        
                                  Select Case .Type
                                     Case dbBoolean
                                        fldDataInfo = "Bit/Boolean"
                                     Case dbByte
                                        fldDataInfo = "Byte"
                                     Case dbInteger
                                        fldDataInfo = "Short"
                                     Case dbLong
                                        fldDataInfo = "Integer"
                                     Case dbCurrency
                                        fldDataInfo = "Currency"
                                     Case dbSingle
                                        fldDataInfo = "Single"
                                     Case dbDouble
                                        fldDataInfo = "Double"
                                     Case dbDate
                                        fldDataInfo = "Date"
                                     Case dbText
                                        fldDataInfo = "Char"
                                            '//FUTURI UTILIZZI _
                                            Al momento non è necessariore restituire la grandezza del campo perchè crea problemi _
                                            con il filtro impostato nella routine chiamante. Nel futuro puoi costruire una
                                            Debug.Print "per futuri utilizzi puoi visualizzare"
                                            Debug.Print "il char e la grandezza del campo"
                                            Debug.Print "Char Width " & Format(.DefinedSize, ">")
                                        
                                                            
                                     Case dbLongBinary
                                        fldDataInfo = "OLE"
                                     Case dbMemo
                                        fldDataInfo = "LongChar"
                                     Case dbGUID
                                        fldDataInfo = "Char Width 16"
                                  End Select
                                

                                    '//se trova il campo di ricerca lo restituisco al _
                                    chiamante e esco dal ciclo impostando il contatore dei campi _
                                    FIELD al massimo con l'operazione i = TotField_i in modo da uscire _
                                    in anticipo dal ciclo.
                                            CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct = fldDataInfo
                                                i = TotField_i

                        Else
                                '//se non trova il campo di ricerca restituisco null
                                CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct = ""

                        End If
                                    '//controllo tipo di campo iterato
                                    Debug.Print "         tipo campo ---> " & fldDataInfo
                                    Debug.Print
                 
               End With
            Next i
         End With
      
                    '//reset variabili ed oggetti
                    i = vbNull
                    Set tblDef = Nothing


'EXIT E GESTIONE ERRORI
'-----------------------------------------------------------------------------------------------

CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct_Exit:
    Exit Function

CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct_Err:
    MsgBox "ERRORE FUNCTION PUBLIC    " & Err.number & " - " & Err.Description, vbCritical, "pfRICERCA_NEL_DB_OGGETTI_Table"
       
    Resume CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct_Exit


End Function
'//*** fine **
'//FUNZIONE CONTROLLA CAMPI TABELLA = attivo il controllo sulla tabella e sul campo
'//==============================================================================================================================//




'//##########################################################################################################################################//
'//
'//                     PROCEDURE DI CONTROLLO OGGETTI TABLE NEL DATABASE CORRENTE O ESTERNO
'//
'//##########################################################################################################################################//
'//
'//NOTE                     : Procedura a funzione che controlla l'esistenza dell'oggetto tabella nel database esterno o interno.
'//MODULO DI CREAZIONE  --> UTIL_MDL40_N05_CONTROLLO_OGGETTI_DB_TABLES
'//DATABASE             --> GE_OGGETTI_DB_MDB.mdb
'//DOVE SI TROVA        --> c:\CASA\LINGUAGGI\ACCESS\ACCESS_PROGETTI_MDB\ACCESS_GESTIONE_OGGETTI_DB_MDB\GESTIONE_OGGETTI\GESTIONE_OGGETTI_MDB\
'//
'//LE FUNZIONI E ROUTINE :
'//                         01)_FUNZIONE_CONTROLLO_OGGETTI_TABELLA = permette il controllo dell'esistenza della tabella oggetto all'interno del db corrente _
                                                                    oppure in un database esterno. _
                                                                    CODICE----------> pfRICERCA_NEL_DB_OGGETTI_Table.Funzione
                                                            
 
'//                         02)...
'//
'//
'//##########################################################################################################################################//

'//CONTROLLO_ESISTENZA_OGGETTI_TABELLA
'//****************************************************************************************************************************
'//01)_FUNZIONE_CONTROLLO_OGGETTI_TABELLA = permette il controllo dell'esistenza della tabella oggetto all'interno del db corrente _
                                                        oppure in un database esterno.
'//CODICE----------> pfRICERCA_NEL_DB_OGGETTI_Table.Funzione
'//PARAMETRI: _
        01) sx_Scelta_db    = parametro di SCELTA DATABASE DATABASE da aprire e rappresenta il database esterno o interno; _
                       è esterno a condizione che ci sia una path estesa ed il mome di un db.mdb altrimenti in caso di parametro null _
                       viene inteso come parametro interno. Es.: _
                                sx_Scelta_db = null il database è quello corrente; _
                                sx_Scelta_db = C:\MyDb.mdb il database è esterno.
'//         02) par_sx_Name_Table =?
'//         03) par_sx_Proprieta_Table =?
'//         04) par_ix_Operazione =?
'//         05) par_ix_Indice1 =?
'//         06) par_bx_Flag_b =?
                       
            
Public Function pfRICERCA_NEL_DB_OGGETTI_Table( _
                                        par_sx_Scelta_db As String, _
                                        par_sx_Name_Table As String, _
                                        par_sx_Proprieta_Table As String, _
                                        par_ix_Operazione As Integer, _
                                        par_ix_Indice1 As Integer, _
                                        par_bx_Flag_b As Boolean) As Boolean


    
    On Error GoTo Err_pfRICERCA_NEL_DB_OGGETTI_Table
       
            
                
                '//CHIAMA IL CONTROLL TABELLE
                '//--------------------------------------------------------------------------------//----------//
                '//CODICE----> pfRICERCA_NEL_DB_OGGETTI_Table.CALL
                
                        '//LE VARIABILE PER IL CONTROLLO OGGETTI
                        'Dim sx_Scelta_db As String
                        'Dim sx_Name_Table As String
                        'Dim sx_Proprieta_Table As String
                        'Dim ix_Operazione As Integer
                        'Dim ix_Indice1 As Integer
                        'Dim bx_Flag_b As Boolean
                        
                        '//IMPOSTO LE VARIABILI
                        'sx_Scelta_db = ""
                        'sx_Name_Table = "GEST_MENU_TB01_VOCI"       '//RICERCO QUESTA TABELLA
                        'sx_Proprieta_Table = "Proprieta"
                        'ix_Operazione = 0
                        'ix_Indice1 = 0
                        'bx_Flag_b = False
    
    
                
                          '//ATTIVO LA FUNZIONE CON I PARAMETRI
                          'Bool1 = pfRICERCA_NEL_DB_OGGETTI_Table(sx_Scelta_db, _
                                                             sx_Name_Table, _
                                                             sx_Proprieta_Table, _
                                                             ix_Operazione, _
                                                             ix_Indice1, _
                                                             bx_Flag_b)
            
                '//CHIAMA IL CONTROLL TABELLE *** FINE ***
                '//--------------------------------------------------------------------------------//----------//
                        
                        
                
                
                '//--------------------------------------------------------------------------------//----------//
                '//1.1    SET VARIABILI, OPEN DB, CONTROLLO OGGETTO NEL DB
                 
                        
                        'SET VARIABILI:
                        sx_Name_Tab = par_sx_Name_Table
                        
                        ix_Operazione = par_ix_Operazione
                            'ix_Indice1= par_Index
                        bx_Flag_b = par_bx_Flag_b
                        sx_Name_Query = ""
                        
                        'OPEN DB: Apre il database scelto per la creazione o cancellazione a condizione _
                        che esista una path ed il db altrimenti il databse da controllare è quello corrente.
                        If par_sx_Scelta_db > "" Then
                            Set Dbs = OpenDatabase(par_sx_Scelta_db)
                        Else
                            Set Dbs = CurrentDb
                        End If
            
                        
                        
                        '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della tabella Nell'insieme Tabledefs
                        '//CONTROLLO PARAMETRI
                        '//-----------------------------------------------------------------------------//
                                
                            Debug.Print "........................................................"
                            Debug.Print "controllo parametri passati alla routine SALVATAGGIO_OGGETTI_DB_PFunct"
                            Debug.Print
                            
                                  Debug.Print "par_sx_Scelta_db ----> " & par_sx_Scelta_db
                                  Debug.Print "par_sx_Name_Table ---> " & par_sx_Name_Table
                                  Debug.Print "par_sx_Name_Query ---> " & par_sx_Proprieta_Table
                                  Debug.Print "par_sx_Name_FORM  ---> " & par_ix_Operazione
                                  Debug.Print "par_sx_Name_Report --> " & par_sx_Name_Report
                                  Debug.Print "par_sx_Name_MACRO ---> " & par_ix_Indice1
                                  Debug.Print "par_sx_Name_Modulo --> " & par_bx_Flag_b
                                  Debug.Print
                                        
                            Debug.Print "........................................................"
                                        
                        '//-----------------------------------------------------------------------------//
                        
                        
                        '//DBS CICLO
                        '//-----------------------------------------------------------------------------//
                        
                            With Dbs
                                
                                Debug.Print .TableDefs.count & _
                                    " TableDefs in database  " & sx_Scelta_db
                                For intCiclo = 0 To .TableDefs.count - 1
                                    
                                    '//BLOCCO ITERAZIONE _
                                    Ricerca di una tabella di sistema o di una tabella _
                                    inesistente con attivazione del blocco. Il nome della tabella _
                                    deve essere attivato manualmente.
                                    
                                        If .TableDefs(intCiclo).Name = "MSysIMEXColumns" Then
                                            MsgBox "BLOCCO ITERERAZIONE sulla tabella " & .TableDefs(intCiclo).Name, vbExclamation, "BLOCCO ITERAZIONE"
                                            Stop
                                            End If
                                    
                                    Debug.Print
                                    Debug.Print "  " & .TableDefs(intCiclo).Name
                                    
                                    
                                    '//TABELLA TROVATA
                                    If .TableDefs(intCiclo).Name = par_sx_Name_Table Then
                                        
                                         'COMANDO RICERCA TABELLA
                                        DoCmd.SelectObject acTable, par_sx_Name_Table, True
                          
                                        Bln1 = True    ' bx_Flag_b True = trovato tabelle si può cancellare
                                        pfRICERCA_NEL_DB_OGGETTI_Table = Bln1 'riassegno alla funzione il valore del bx_Flag_b
                                        GoTo Exit_pfRICERCA_NEL_DB_OGGETTI_Table  'esci
                                    Else
                                        '//TABELLA NON TROVATA
                                        Bln1 = False   ' bx_Flag_b False = non trovato TABELLA
                                        pfRICERCA_NEL_DB_OGGETTI_Table = Bln1 'riassegno alla funzione il valore del bx_Flag_b
                                    End If
                                    
                                Next intCiclo
            
                            End With
                        '//-----------------------------------------------------------------------------//
                           
            
            
            '------------------------------------------------------------------------------
            '                       FINE FUNCTIONE E GESTIONE ERRORI

'//ERRORI ED USCITA
'//--------------------------------------------------------------------------------//------------------------------------------//
Exit_pfRICERCA_NEL_DB_OGGETTI_Table:
            Dbs.Close
            Set Dbs = Nothing
            
            Exit Function
            
Err_pfRICERCA_NEL_DB_OGGETTI_Table:
            MsgBox "ERRORE FUNCTION PUBLIC    " & Err.number & " - " & Err.Description, vbCritical, "pfRICERCA_NEL_DB_OGGETTI_Table"
            Resume Exit_pfRICERCA_NEL_DB_OGGETTI_Table
            
            
End Function
            
'//*** FINE ***
'//CONTROLLO_ESISTENZA_OGGETTI_TABELLA
'//****************************************************************************************************************************
    

'//##########################################################################################################################################//
'//
'//                     PROCEDURE DI CONTROLLO OGGETTI TABLE NEL DATABASE CORRENTE O ESTERNO *** fine ***
'//
'//##########################################################################################################################################//
    



    
'//##########################################################################################################################################//
'//
'//                     MODELLI DI ISTANZIAMENTO CLASSE
'//
'//##########################################################################################################################################//



Private Sub ZZFINE_____________________________________________()

End Sub


Private Sub ZZFINE_RECUPERO_CODOGGETTO_da_TMP_pbFunct()
        
        '//CHIAMO LA ROUTINE CONTROLLO_OGGETTO_FORM_LaFormAPERTA DELLA CLASSE OGGETTI FORM
        '//CODICE----------->CONTROLLO_OGGETTO_FORM_LaFormAPERTA.Salvataggio_ATTIVITA.Attivazione
        '//--------------------------------------------------------------------------------//-//
        '//CLASSE_OSPITE        : CLASSE_GESTIONE_OGGETTI_DATABASE
        '//PARAMETRI            : sono 4 di seguito indicati
        'par_sxCODOGGETTO       = CODICE OGGETTO OBBLIGATORIO
        'par_sxTIPOGGETTO       =
        'par_sxATTIVITAEVENTO   =
        'par_sxCODATTIVITA      = CODICE ATTIVITA OBBLIGATORIO
        '//NOTA  :Salvo i parametri del codice oggetto e del codice attiva nella tabella tmp solo nel caso in cui _
        sia IL CODICE OGGETTO che il CODICE ATTIVITA sono VALORIZZATI altrimenti la classe ricevente emette _
        un msg di avviso.
        


                '1) ISTANZIO LA CLASSE
                '........................................................
                
                'Recupero la stringa da applicare alla
                'successiva classe
                '
                'sxCODOGGETTO = ""
                '//trovo il codice oggetto mediante il codice attivita
                'sxCODOGGETTO = TROVA_CLASSE_DI_GESTIONE_pFunct(par_sxCODATTIVITA)
                 
                'ISTANZIO LA CLASSE CLASSE_GESTIONE_OGGETTI_DATABASE
                'Assegno e creo l'oggetto di classe
                '--->                     Dim MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE As CLASSE_GESTIONE_OGGETTI_DATABASE
                '--->                     Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = New CLASSE_GESTIONE_OGGETTI_DATABASE
                    
                    
                
                    'PARAMETRI DI ESTRAZIONE
                    'Recupero la stringa da applicare alla
                    'successiva classe
                     '---> sxCODOGGETTO = ""
                     '---> sxCODOGGETTO = MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE.RECUPERO_CODOGGETTO_da_TMP_pbFunct
               
                
                '--->                     Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = Nothing

        '//--------------------------------------------------------------------------------//-//
       
End Sub



'//MODELLO ATTIVA LA CLASSE PER LA RICERCA DEI TIPI DI CAMPI DI UNA TABELLA ACCESS
Private Sub ZZFINE_CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct()
'//CODICE----------->ZZFINE_CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct.ATTIVA_LA_CLASSE
  

'...........................................................................................................................................
'                                       **** COPIARE DA QUI *****
'   Devi copiare siala Routine che la funzione ed inserirla in un modulo, classe, form o report per utilizzare le funzioni comuni.
'...........................................................................................................................................



'//ROUTINE ATTIVA LA CLASSE PER I CONTROLLO OGGETTI TABLE =  Questa Routine multiuso _
                    puo' essere attivata per diverse chiamate sul tipo di campo, da utilizzare nel modulo _
                    di classe, di una form, report o modulo del progetto ospite. Questa Routine e la Funzione _
                    associata per l'attivazione della classe si trova nella classe identificata da codice _
                    di seguito indicato:
'//CODICE----------->ZZFINE_CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct.ATTIVA_LA_CLASSE
'//^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^//
'---> Private Sub ATTIVA_CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct()
'---> Dim NameTable_s As String
'---> Dim NameField_s As String
'---> Dim AttivaMessaggio_b As Boolean
    
    '//PER ATTIVARE IL MESSAGGIO -> TRUE
    '---> AttivaMessaggio_b = True                        '//FALSE = MESSAGGIO NON ATTIVATO, TRUE MESSAGGIO DI CONTROLLO ATTIVATO
    
    '---> NameTable_s = "MSys_DF02_TIPO_OGGETTO"          '// nome della tabella da ricerca
    '---> NameField_s = "Scel_b"                          '// il campo da controllare. es. Scel_b, TIPOGGETTO_s
    '//Restituisco una stringa se trovato oppure null
    '---> vV1 = Cls_CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct(NameTable_s, NameField_s)
        
        '//CONTROLLO MSGBOX POSITIVO O NEGATIVO
        '//..........................................................................//
            '---> If AttivaMessaggio_b = True Then '//SE TRUE IL MESSAGGIO VIENE ATTIVATO
                '---> If vV1 > "" Then
                   '---> MsgBox NameTable_s & " QUESTA TABELLA CONTIENE questo tipo di campo : " & vV1, _
                                        vbExclamation, "MSG CONTROLLO TIPO DI CAMPO"
                '---> Else
                   '---> MsgBox "ATTENZIONE! QUESTA TABELLA NON ESISTE -> " & Chr$(13) & Chr$(13) & NameTable_s & _
                            Chr$(13) & Chr$(13) & _
                           "OPPURE NON ESISTE IL TIPO DI CAMPO RICERCATO ->" & _
                            Chr$(13) & Chr$(13) & _
                            NameField_s, vbCritical, "ERRORE RICERCA TIPO DI CAMPO"
                '---> End If
            '---> End If
        '//..........................................................................//

'--->End Sub




'//FUNZIONE CHE ATTIVA LA CLASSE OGGETTI = la quale permette la ricerca dei campi della tabella _
                    passata come parametro. Questa funzione è riutilizzabile dalla _
                    routine di seguito indicata. Restituisce una stringa che rappresenta _
                    il campo e le sue proprieta. Puo essere utilizzata per verificare ogni _
                    singolo campo della tabella.
'---> Private Function Cls_CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct(par_NameTable_s As String, par_NameField_s As String) As String
        
'//CODICE---------->CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct.attiva.classe
'//--------------------------------------------------------------------------------//-//
'//CLASSE_OSPITE        : CLASSE_GESTIONE_OGGETTI_DATABASE
'//PARAMETRI            : stringa che rappresenta il nome della della tabella + _
                          il nome del campo da ricerca.
'//RESTITUISCE          : Stringa ossia il tipo di campo esaminato
'//NOTE         : L'attivazione della classe viene inserita in una Funzione denominata _
                ATTIVA_LA_FUNZIONE_MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE _
'//MODULO_DOVE_SI_TROVA:CODICE---------->CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct.attiva
'//CALSSE_CHE_LA_IMPLEMENTA:CODICE---------->CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct.attiva

 
    '1) ISTANZIO LA CLASSE
    '........................................................

    'ISTANZIO LA CLASSE CLASSE_GESTIONE_OGGETTI_DATABASE
    'Assegno e creo l'oggetto di classe
    '---> Dim MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE As CLASSE_GESTIONE_OGGETTI_DATABASE
    '---> Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = New CLASSE_GESTIONE_OGGETTI_DATABASE



        'CONTROLLO SE IL TIPO DI CAMPO
        'Passo come parametri il nome della tabella in cui ricercare il campo
         '--->                            NameField_s = "nome tabella?"
         '--->                            NameField_s = "nome field?"
         'Restituisce il nome del campo ed il titolo altrimenti se non esiste la tabella _
            o non esiste il field viene emesso un msg di avviso dalla funzione.
        '---> Cls_CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct = MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE.CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct(par_NameTable_s, par_NameField_s)
        '---> Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = Nothing
'//--------------------------------------------------------------------------------//-//

'---> End Function


'...........................................................................................................................................
'                                       **** FINO QUI *****
'...........................................................................................................................................


End Sub




Private Sub ZZFINE_Controlla_FormGESTIONE_APERTA()





        '//CODICE----------->CONTROLLO_OGGETTO_FORM_LaFormAPERTA.Salvataggio_ATTIVITA.Attivazione
        '//--------------------------------------------------------------------------------//-//
        '//CLASSE_OSPITE        : CLASSE_GESTIONE_OGGETTI_DATABASE
        '//PARAMETRI            : stringa che rappresenta il nome della form.
        '//RESTITUISCE          : bool True se la form è aperta
        
                '//LE VARIABILI
                '--->                    Dim par_sxNameForm   As String                      '//La variabile della form passato come parametro

                '1) ISTANZIO LA CLASSE
                '........................................................
                        
                'ISTANZIO LA CLASSE CLASSE_GESTIONE_OGGETTI_DATABASE
                'Assegno e creo l'oggetto di classe
                '--->                    Dim MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE As CLASSE_GESTIONE_OGGETTI_DATABASE
                '--->                    Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = New CLASSE_GESTIONE_OGGETTI_DATABASE
                    
                    
                
                    'CONTROLLO SE LA FORM E' APERTA
                    'Passo come parametro la stringa nome form e ritora True se è aperta
                     '--->                            par_sxNameForm = Str1
                     '--->                            Routine_CONTROLLO_FORM_DI_GESTIONE_APERTA = MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE.Controlla_FormGESTIONE_APERTA(par_sxNameForm)
                

                    '--->                            Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = Nothing

        '//--------------------------------------------------------------------------------//-//



'//IN FONDO PUOI INSERIRE LA ROUTINE nel caso di parecchie operazioni di controllo della form se _
        aperta oppure no. La routine contiene un solo richiamo della classe.
'//ROUTINE PER IL CONTROLLO DELLA FORM APERTA
'//CODICE -----> Routine_CONTROLLO_FORM_DI_GESTIONE_APERTA
'//==========================================================================================================================================//
'Private Function Routine_CONTROLLO_FORM_DI_GESTIONE_APERTA(Str1) As Boolean


'On Error GoTo Err_Routine_CONTROLLO_FORM_DI_GESTIONE_APERTA

        '//CHIAMO LA FUNZIONE DI CONTROLLO SE LA FORM E' APERTA
        '//--------------------------------------------------------------------------------//-//
        '//NOTE: passare come stringa il nome della form da controllare. La funzione attiva _
                la classe di gestione oggetti con la quale viene effettuato il controllo tra _
                gli oggetti del db corrente, ossia le form e per quella specificata se la _
                form passata come parametro è aperta viene restituito TRUE.
                
                '----> Routine_CONTROLLO_FORM_DI_GESTIONE_APERTA (Str1)

        '//--------------------------------------------------------------------------------//-//


        '// CONTROLLA SE LA FORM GESTIONE E' APERTA
        '//============================================================================================//
        '//NOTE : chiamo la funzione per il controllo se la form principale  è aperta _
        viene passato come parametro il nome della form MASTER e restituisce un Boolean _
        TRUE se la form è aperta.
        
                '//1) ISTANZIO LA CLASSE
                '//--------------------------------------------------------------------------------//-//
                
                           
                    'ISTANZIO LA CLASSE CLASSE_GESTIONE_OGGETTI_DATABASE
                    'Assegno e creo l'oggetto di classe
                    '--->
                    'Dim MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE As CLASSE_GESTIONE_OGGETTI_DATABASE
                    '--->
                    'Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = New CLASSE_GESTIONE_OGGETTI_DATABASE
                        
                        
                        '//
                        'Bolean1 = False
                     
                        'Bolean1 = MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE.Controlla_FormGESTIONE_APERTA(Str1)
                    
                    
                    '---> Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = Nothing
    
            '//--------------------------------------------------------------------------------//-//

        '//============================================================================================//
          



        
'//USCITA E CONTROLLO ERRORI
'//==========================================================================================================================================//

'Exit_Routine_CONTROLLO_FORM_DI_GESTIONE_APERTA:
'    Exit Function

'Err_Routine_CONTROLLO_FORM_DI_GESTIONE_APERTA:
'    MsgBox Err.Description
'    Resume Exit_Routine_CONTROLLO_FORM_DI_GESTIONE_APERTA
'End Function

'End Function
'//ROUTINE PER IL CONTROLLO DELLA FORM APERTA *** FINE ***
'//==========================================================================================================================================//

End Sub

Private Sub ZZFINE_Istanzia_TROVA_CLASSE_DI_GESTIONE_pFunct()



      


                '//CHIAMO LA ROUTINE SALVO_ATTIVITA_CMB_02_IN_TMP_Sub DELLA CLASSE OGGETTI FORM
                '//CODICE----------->SALVO_ATTIVITA_CMB_02_IN_TMP_Sub.Salvataggio_ATTIVITA.Attivazione
                '//--------------------------------------------------------------------------------//-//
                '//CLASSE_OSPITE        : CLASSE_GESTIONE_OGGETTI_DATABASE
                '//PARAMETRI            : sono 4 di seguito indicati
                'par_sxCODOGGETTO       = CODICE OGGETTO OBBLIGATORIO
                'par_sxTIPOGGETTO       =
                'par_sxATTIVITAEVENTO   =
                'par_sxCODATTIVITA      = CODICE ATTIVITA OBBLIGATORIO
                '//NOTA  :Salvo i parametri del codice oggetto e del codice attiva nella tabella tmp solo nel caso in cui _
                sia IL CODICE OGGETTO che il CODICE ATTIVITA sono VALORIZZATI altrimenti la classe ricevente emette _
                un msg di avviso.



                            '1) ISTANZIO LA CLASSE
                            '........................................................
                            'Uscita dalla routine se non trova la classe tramite il codice _
                            oggetto allora viene emesso un msg di errore.
                            'Recupero la stringa da applicare alla
                            'successiva classe tramite il codice attivita ritrovo il _
                            codice oggetto.
                            '---> sxCODOGGETTO = ""
                            '//trovo il codice oggetto mediante il codice attivita
                            '---> sxCODOGGETTO = TROVA_CLASSE_DI_GESTIONE_pFunct(par_sxCODATTIVITA)
                             
                            'ISTANZIO LA CLASSE CLASSE_GESTIONE_OGGETTI_DATABASE
                            'Assegno e creo l'oggetto di classe
                            '---> Dim MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE As CLASSE_GESTIONE_OGGETTI_DATABASE
                            '---> Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = New CLASSE_GESTIONE_OGGETTI_DATABASE
                                
                                'PARAMETRI DI ESTRAZIONE
                                'Recupero la stringa da applicare alla
                                'successiva classe
                                '---> sxCODOGGETTO = MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE.TROVA_CLASSE_DI_GESTIONE_pFunct(par_sxCODATTIVITA)
                                '---> Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = Nothing


                                'SE NON ESISTE LA FUNZIONE USCITA DALLA ROUTINE
                                '........................................................
                                'Uscita dalla routine se non trova la classe tramite il codice _
                                oggetto allora viene emesso un msg di errore.
                                    
                                    '---> If sxCODOGGETTO = "" Then
                                        
                                            '---> MsgBox "Funzione non Abilitata!!", vbExclamation, "MSG GESTIONI"
                                        
                                            '---> GoTo Exit_RECUPERO_I_DATI_DALLA_CLASSE_Sub
                                        
                                    '---> End If
                                '........................................................


            
                    '//--------------------------------------------------------------------------------//-//
                




End Sub
Private Sub ZZFINE_Istanzia_SALVATAGGIO_PARAMETRO_CODOGGETTO_in_TMP()


'//CODICE----------->SALVATAGGIO_PARAMETRO_CODOGGETTO_in_TMP_pFunct.AttivoLaClasse
'//--------------------------------------------------------------------------------//-//
'//CLASSE_OSPITE        : CLASSE_GESTIONE_OGGETTI_DATABASE
'//PARAMETRI            : sono il par_sxCODOGGETTO       = CODICE OGGETTO OBBLIGATORIO
       
'//NOTA  : AGGIORNO IL CODICE OGGETTO nella tabella oggetti tmp GEST_Tb01_N01_PARAM_OGGETTI_TMP
       
   
   
               'ISTANZIO LA CLASSE CLASSE_GESTIONE_OGGETTI_DATABASE
               'Assegno e creo l'oggetto di classe
               'Creo la variabile oggetto                              ---> Dim MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE As CLASSE_GESTIONE_OGGETTI_DATABASE
               'Creo l'oggetto classe                                  ---> Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = New CLASSE_GESTIONE_OGGETTI_DATABASE
               
               'RECUPERO PARAMETRO DA TABELLA OGGETTI
                'Utilizzo la funzione pubblica di recupero del codice oggetto dal
                'codice attivita. Restituisco la chiamante il CODICE OGGETTO
                'Creo l'oggetto e attivo la funzione e ricevo la stringa ---> Str1 = MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE.SALVATAGGIO_PARAMETRO_CODOGGETTO_in_TMP_pFunct(par_sxCODOGGETTO)
               
               'Libero la memoria                                       ---> Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = Nothing
                                           
'//--------------------------------------------------------------------------------//------------------------------------------------//
End Sub


Private Sub ZZFINE_Istanzia_SALVO_ATTIVITA_CMB_02_IN_TMP_Sub()

        '//CHIAMO LA ROUTINE SALVO_ATTIVITA_CMB_02_IN_TMP_Sub DELLA CLASSE OGGETTI FORM
        '//CODICE----------->SALVO_ATTIVITA_CMB_02_IN_TMP_Sub.Salvataggio_ATTIVITA.Attivazione
        '//--------------------------------------------------------------------------------//-//
        '//CLASSE_OSPITE        : CLASSE_GESTIONE_OGGETTI_DATABASE
        '//PARAMETRI            : sono 4 di seguito indicati
        'par_sxCODOGGETTO       = CODICE OGGETTO OBBLIGATORIO
        'par_sxTIPOGGETTO       =
        'par_sxATTIVITAEVENTO   =
        'par_sxCODATTIVITA      = CODICE ATTIVITA OBBLIGATORIO
        '//NOTA  :Salvo i parametri del codice oggetto e del codice attiva nella tabella tmp solo nel caso in cui _
        sia IL CODICE OGGETTO che il CODICE ATTIVITA sono VALORIZZATI altrimenti la classe ricevente emette _
        un msg di avviso.
        
                     
            'ISTANZIO LA CLASSE CLASSE_GESTIONE_OGGETTI_DATABASE
            'Assegno e creo l'oggetto di classe
            '---> Dim MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE As CLASSE_GESTIONE_OGGETTI_DATABASE
            '---> Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = New CLASSE_GESTIONE_OGGETTI_DATABASE
                
                
            
            'PARAMETRI DI ESTRAZIONE
            'Attivo il metodo della classe rappresentata dalla routine passando
            'i 4 parametri necessari
            'sxCODOGGETTO = s
            'sxTIPOGGETTO = s
            'sxATTIVITAEVENTO = s
            'sxCODATTIVITA = s
            '---> MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE.SALVO_ATTIVITA_CMB_02_IN_TMP_Sub sxCODOGGETTO, sxTIPOGGETTO, sxATTIVITAEVENTO, sxCODATTIVITA
            
            
            '---> Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = Nothing

        
        '//--------------------------------------------------------------------------------//-//
        




End Sub

Private Sub ZZFINE_Istanzia_RECUPERO_CODOGGETTO()

'//CODICE----------->RECUPERO_CODOGGETTO_pbFunct.AttivoLaClasse
'//--------------------------------------------------------------------------------//-//
'//CLASSE_OSPITE        : CLASSE_GESTIONE_OGGETTI_DATABASE
'//PARAMETRI            : sono il par_sxCODOGGETTO       = CODICE OGGETTO OBBLIGATORIO
       
'//NOTA  : AGGIORNO IL CODICE OGGETTO nella tabella oggetti tmp GEST_Tb01_N01_PARAM_OGGETTI_TMP
       
   
   
               'ISTANZIO LA CLASSE CLASSE_GESTIONE_OGGETTI_DATABASE
               'Assegno e creo l'oggetto di classe
               'Creo la variabile oggetto                              ---> Dim MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE As CLASSE_GESTIONE_OGGETTI_DATABASE
               'Creo l'oggetto classe                                  ---> Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = New CLASSE_GESTIONE_OGGETTI_DATABASE
               
               'RECUPERO PARAMETRO DA TABELLA OGGETTI
                'Utilizzo la funzione pubblica di recupero del codice oggetto dal
                'codice attivita. Restituisco la chiamante il CODICE OGGETTO
                'Creo l'oggetto e attivo la funzione e ricevo la stringa ---> TROVA_CODOGGETTO_Funct = MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE.RECUPERO_CODOGGETTO_pbFunct(sxCODATTIVITA)
               
               'Libero la memoria                                       ---> Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = Nothing
                                           
    '//--------------------------------------------------------------------------------//------------------------------------------------//
       
 
End Sub

Private Sub ZZFINE_Istanzia_CERCA_OGGETTI()

    
    '//01) CLASSE_GESTIONE_OGGETTI_DATABASE
    '//ISTANZIA LA CLASSE
    '//06.30
    '//FUNZIONE CERCA OGGETTI = Con due parametri Tipo oggetto e Nome oggetto lo ricerca nella db.
    '//CODICE-------->CERCA_OGGETTI.01.Attiva
    '//.....................................................................................................
        '//IMPOSTO LA VARIABILE OGGETTO = per l'utilizzo della classe
            '//---> Dim MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE As CLASSE_GESTIONE_OGGETTI_DATABASE
    
        '//ISTANZIO LA CLASSE = MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE assegnando _
             l'oggetto di classe alla variabile e creo creo l'oggetto di classe
            '//---> Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = New CLASSE_GESTIONE_OGGETTI_DATABASE
        
        '//PARAMETRI= TIPO OGGETTO E NOME OGGETTO.
            '//---> Str1 = MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE.CERCA_OGGETTI(sxTIPOGGETTO, sxNOMEOGGETTO)
            
        '//RILASCIO LA MEMORIA = rilascio l'oggetto ed attivo il distruttore della classe
            '//---> Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = Nothing
    
    '//.....................................................................................................



End Sub


'//@ESEMPI@ISTANZIAMENTO@CLASSE
Private Sub ZZFINE_ESEMPI_ISTANZIAMENTO_CLASSE()


    
    '//01) CLASSE CON LA PATH ED IL FILE DA APRIRE
    '//
    '//MODELLO ISTANZIAMENTO CLASSE ATTI DI GESTIONE
    '//.....................................................................................................
        '//IMPOSTO LA VARIABILE OGGETTO = per l'utilizzo della classe
            'Dim MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE As CLASSE_GESTIONE_OGGETTI_DATABASE
    
        '//ISTANZIO LA CLASSE = MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE assegnando _
           l 'oggetto di classe alla variabile e creo creo l'oggetto di classe
            'Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = New CLASSE_GESTIONE_OGGETTI_DATABASE
        
        '//PATH O FILE = Inserire o una path oppure una path+file con percoso esteso.
            'Str1 = "C:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHVI_MDB\LLPP_GESTIONE\LPP_ANALISI\GESTIONE_SERVIZI\MDB"
            'MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE.APRO_PATH_O_FILE_GENERICI_pSub (Str1)
            
        '//RILASCIO LA MEMORIA = rilascio l'oggetto ed attivo il distruttore della classe
            'Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = Nothing
    
    '//.....................................................................................................
    

    '//02) CLASSE CON PER L'ATTIVAZIONE DI UNA FUNZIONE O ROUTINE E CON RETURN DI UN BOOLEAN
    '//
    '//MODELLO ISTANZIAMENTO CLASSE ATTI DI GESTIONE
    '//.....................................................................................................
        '//IMPOSTO LA VARIABILE OGGETTO = per l'utilizzo della classe
            'Dim MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE As CLASSE_GESTIONE_OGGETTI_DATABASE
    
        '//ISTANZIO LA CLASSE = MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE assegnando _
           l 'oggetto di classe alla variabile e creo creo l'oggetto di classe
            'Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = New CLASSE_GESTIONE_OGGETTI_DATABASE
        
        '//CHIAMO LA FUNZIONE CON I PARAMETRI E CON RESTITUZIONE DI UN BOOLEAN
            'Str1 = "C:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHVI_MDB\LLPP_GESTIONE\LPP_ANALISI\GESTIONE_SERVIZI\MDB"
            'Bool1 =  MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE.pfRICERCA_NEL_DB_OGGETTI_Table(sx_Scelta_db, _
                                                                                      sx_Name_Table, _
                                                                                      sx_Proprieta_Table, _
                                                                                      ix_Operazione, _
                                                                                      ix_Indice1, _
                                                                                      bx_Flag_b)
                                
            
            
            
        '//RILASCIO LA MEMORIA = rilascio l'oggetto ed attivo il distruttore della classe
            'Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = Nothing
    
    '//.....................................................................................................
    

End Sub


       



















