VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_UTILITA_Frm12_SERVIZI_TUTTI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//DENOMINAZIONE    -> Modello_Funzione_Pfunct
Option Compare Database
'//Variabili di database e GENERALI
'//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>:>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'DAO
Dim daoDB As DAO.Database
Dim DaoWks As DAO.Workspace
Dim daoRS As DAO.Recordset
Dim DaoRs_Lett As DAO.Recordset

'ADO
Dim ADODB As Database
Dim AdodaoRs As Recordset

'Contatori
Dim icount As Integer
Dim iTOTcount As Integer
Dim dbl_count As Double

'Le variabili generiche
Dim sSql As String                                          ' Stringa di estrazione
Dim sSql_Lett As String                                     ' Stringa di estrazione

'Variabili generali
Dim Str1 As String
Dim Int1 As Integer
Dim Int2 As Integer
Dim Int3 As Integer
Dim Lng1 As Long
Dim Dbl1 As Double
Dim Bln1 As Boolean
Dim Vv1 As Variant
Dim obj1 As Object

'Gestione parametri comandi
Dim sxParamCmd_3 As String


'Larghezza e numero di colonna
Dim sLarg_Col As String
Dim iNum_Col As Integer


'//ERRORI PROCEDURA_FUNCTION O ROUTINE
Dim ProceduraMessaggioErrore_s As String
Dim ProceduraAttivaEseguita_s As String



'//RECUPERO PARAMETRO DA TABELLA OGGETTI
'//File da ricercare e la path
Dim MyFile_s As String
Dim MyPath_s As String
        


'//Variabili di database e GENERALI     *** FINE ***
'//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>:>>>>>>>>>>>>>>>>>>>>>>>>>>>>



'//CMD          EVENTI COMANDI BUTTON
'//***********************************************************************//

'//@BUTTON_01@PAG_01_(Comando pagina 01)
Private Sub Cmd_01_Click()

    '//CANCELLA TMP
    DoCmd.OpenQuery "MsysDbEstTb01Qry01_Dlt01_OBJECT_TMP"
    Me.LABEL_TIT_01.Caption = "CANCELLAZIONE OGGETTI IN TMP AVVENUTA"
End Sub

'//@BUTTON_01@PAG_03_(Comando 01 della  pagina 03)
Private Sub Cmd_01_Pg_03_Click()
    
     '//CANCELLA TMP
    DoCmd.OpenQuery "MsysDbEstTb01Qry01_Dlt01_OBJECT_TMP"
    Me.LABEL_TIT_01.Caption = "CANCELLAZIONE OGGETTI IN TMP AVVENUTA"

End Sub

'//@BUTTON_02
'//@CMD@02@Cmd_02_Click
Private Sub Cmd_02_Click()
    '//ATTIVO LA ROUTINE per il controllo del db esterno
    Call ListTablesInExternalDB
    Me.LABEL_TIT_01.Caption = "IMPORTATO OGGETTI DB ESTERNO IN TMP"
    
    '//codice : @ROUTINE@ATTIVA_(@controllo@esterno@db)_(chiama la routine)
 
End Sub

'//@BUTTON_02@PAG_03_(Comando 02 della  pagina 03)
Private Sub Cmd_02_Pg_03_Click()
    'Todo:da correggere
    'Call ListFormsInExternalDB
    Me.LABEL_TIT_01.Caption = "IMPORTATO OGGETTI FORM DAL DB ESTERNO IN TMP"

End Sub

Private Sub Cmd_03_Pg_03_Click()
    
    '//CANCELLA ARCHIVI OGGETTI
    DoCmd.OpenQuery "MsysDbEstTb01Qry01_Dlt01_OBJECT"
    Me.LABEL_TIT_01.Caption = "CANCELLAZIONE OGGETTI NELLA TABELLA DI ARCHIVIO"

End Sub

'//@BUTTON_04
Private Sub Cmd_04_Click()
    '//@@button_(cancella tabelle in db esterno)
    '//ATTENZIONE = se la tabella è stata gia cancellata nel db esterno e quindi non esiste, _
       l'attivazione della rountine non ha nessun effetto anche se viene rilasciato il segnale _
       del messaggio sotto riportato.  TODO = fare integrazione alla routine che dia il messaggio se le tabelle non sono esistenti _
                                        nel db esterno quando si manda l'evento di cancellazione.
    DeleteSpecificTableInExternalDB_Psub
    
    Me.LABEL_TIT_01.Caption = "CANCELLATI GLI OGGETTI SELEZIONATI NEL DB ESTERNO"
    
End Sub


Private Sub Cmd_04_Pg_03_Click()

    '//@IMPORTI@DATI_{DAL @DB@ESTERNO, importo le form da tmp in archivio}
    DoCmd.OpenQuery "MsysDbEstTb01Qry01_Ins01_OBJECT_Tmp_To_Archivio"
    Me.LABEL_TIT_01.Caption = "IMPORTAZIONE OGGETTI FORM DAL DB ESTERNO NELLA TABELLA DI ARCHIVIO ESEGUITO"

End Sub
'//PAGINA 3 COMANDI LINEA 5
Private Sub Cmd_05_01_Pg_03_Click()
    Call DELETE_FormsInExternalDB
End Sub

Private Sub Cmd_05_Click()
    DELETE_TablesInExternalDB
End Sub

Private Sub Cmd_06_Click()


    '//@IMPORTI@DATI_{DAL @DB@ESTERNO}
    DoCmd.OpenQuery "MsysDbEstTb01Qry01_Ins01_OBJECT_Tmp_To_Archivio"
    Me.LABEL_TIT_01.Caption = "IMPORTAZIONE OGGETTI DAL DB ESTERNO NELLA TABELLA DI ARCHIVIO ESEGUITO"


End Sub

Private Sub Cmd_07_Click()

    '//CANCELLA ARCHIVI OGGETTI
    DoCmd.OpenQuery "MsysDbEstTb01Qry01_Dlt01_OBJECT"
    Me.LABEL_TIT_01.Caption = "CANCELLAZIONE OGGETTI NELLA TABELLA DI ARCHIVIO"

End Sub

'//lato query
Private Sub Cmd_08_Click()

    '//CANCELLA TMP
    DoCmd.OpenQuery "MsysDbEstTb01Qry01_Dlt01_OBJECT_TMP"
    Me.LABEL_TIT_01.Caption = "CANCELLAZIONE OGGETTI IN TMP AVVENUTA"

End Sub


Private Sub Cmd_09_Click()
    '//ATTIVO LA ROUTINE per il controllo del db esterno
    Call ListQUERYESInExternalDB
    Me.LABEL_TIT_01.Caption = "IMPORTATO OGGETTI QUERY DAL DB ESTERNO IN TMP"

End Sub


'//CMD          EVENTI COMANDI BUTTON *** FINE **
'//***********************************************************************//




'//COLLEGA_SERVIZI  @COLLEGA@TABELLE@LINK  TODO: DA FARE IL COLLEGAMENTO CIVILIA ??? @COLLEGA@TUTTE@LE@TABELLE@CIVILIA??
'//========================================================================================================================================//
'//Tipo           : Routine pubblica.
'//Attività       : COLLEGA TABELLE
'//Note           : Collego tutte le tabelle Civilia al Bd
'//Parametro      : par_iAnnoImp = anno di imposta e par_sCodiceTributo = Codice Tributo F24.
'//Codice         : COLLEGA_SUB_N01_TabelleCivilia_pSub.01
'//

Public Sub COLLEGA_SUB_N01_TabelleCivilia_pSub()

'//MessaggiDiErrore
Dim ProceduraMessaggioErrore_s As String
Dim ProceduraAttivaEseguita_s As String
 
 
'//Campo
Dim CampoCercato_s As String

'//Campi parametri
Dim par_AnnoImp_i As Integer
Dim par_CodiceTributo_s As String

            
    '//....
On Error GoTo Err_COLLEGA_SUB_N01_TabelleCivilia_pSub


        
        '//Imposto i parametri
        ProceduraAttivaEseguita_s = "COLLEGA_SUB_N01_TabelleCivilia_pSub"
        ProceduraMessaggioErrore_s = "Errore nella procedura"
        
    '//ITERO NELLA TABELLA
    '//.....................................................................................................
    '//Note           : Tramite una Select vengono individuati i valori da restiuire.

        '//RECUPERO PARAMETRO DA TABELLA OGGETTI
        '//-------------------------------------------------------------------------------
        'CIVILIA_{@================================================@}_Tab
        
        
            
    '//-------------------------------------------------------------------------------
    '//*** fine ***
    '//ITERO NELLA TABELLA
    '//.....................................................................................................

'//USCITA  E GESTIONE ERRORI
'//..............................................................................................................



Exit_COLLEGA_SUB_N01_TabelleCivilia_pSub:
    Exit Sub

Err_COLLEGA_SUB_N01_TabelleCivilia_pSub:
 '//-------------------------------------------------------------------------------
    MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & " Procedura -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Stop
    Resume Exit_COLLEGA_SUB_N01_TabelleCivilia_pSub
'//-------------------------------------------------------------------------------
        
End Sub

'//*** FINE ***
'//COLLEGA_SERVIZI
'//========================================================================================================================================//

'------------------------------------------------------------
' COLLEGA_CIVILIA_TB01bloccata
' parametro Blocco_b = True procedura  con False si esporta.
' Note: collega tutte le tabelle civilia al database corrente
'------------------------------------------------------------


Public Function COLLEGA_CIVILIA_TB01(Blocco_b As Boolean)
On Error GoTo COLLEGA_CIVILIA_TB01_Err

'//COLLEGA TABELLA LLPP_DF06_FASCICOLI
'//-----------------------------------------------------------------------------//
'//NOTE: parametro Blocco_b = True procedura bloccata con False si esporta.
        '// Call COLLEGA_CIVILIA_TB01(False)

'//-----------------------------------------------------------------------------//


    
    If BloccoProcedura_b = True Then
        MsgBox "BLOCCO PROCEDURA DI COLLEGAMENTO LLPP_DF06_FASCICOLI per evitare la cancellazione della tabella dati! " & _
            "Impostare a False la variabile BloccoProcedura_b", vbExclamation, "MSG BLOCCO COLLEGAMENTO"
        GoTo COLLEGA_CIVILIA_TB01_Exit
    End If


    ' In caso di errore non viene eseguita l'istruzione elimina oggetto ma quella successiva
    On Error Resume Next
    ' viene prima cancellata con
    ' EliminaOggetto
    ' e poi riscotruito il collegamento con
    ' TrasferisciDatabase
    ' e cioè RICOSTRUISCE IL COLLEGAMENTO trasferisce database. Se invece esiste la tabella  ATTI DI IMPEGNO La stessa
    ' prima elimina la tabella collegata gia esistente
    
    
    
    '//01)=DELETE LINK
    
    '//delete civilia atti
    DoCmd.DeleteObject acTable, "CIVILIA_Tb01_}-------------------------------------------------@"
    DoCmd.DeleteObject acTable, "CIVILIA_Tb01_ATTI"
    DoCmd.DeleteObject acTable, "CIVILIA_Tb01_ATTI_TMP"
    
    '//delete tabelle export
    
    DoCmd.DeleteObject acTable, "MSys_DF00_}----------------------------------------------------@"
    DoCmd.DeleteObject acTable, "MSys_DF02_TIPO_OGGETTO"
    DoCmd.DeleteObject acTable, "MSys_DF10_DATABASE_TMP"
    DoCmd.DeleteObject acTable, "Msys_DF11_PROGETTI"
    DoCmd.DeleteObject acTable, "Msys_DF12_TIPO_OGGETTI"
    
    DoCmd.DeleteObject acTable, "MSys_ELENCO_}-------------------------------------------------@"
    DoCmd.DeleteObject acTable, "MSys_ELENCO_ObjProgetti"
    
    '//DELETE LE DEFINIZIONI
    
    DoCmd.DeleteObject acTable, "CIVILIA_DF01_}-------------------------------------------------@"
    DoCmd.DeleteObject acTable, "CIVILIA_DF01_SETTORE"
    DoCmd.DeleteObject acTable, "CIVILIA_DF02_TIPO_ATTO"
    
    
    
    
    
    
    
    
    '//02)=ACCODA LINK
    
    ' collego le tabelle atti da CIVILIA_N02_Tb01_ATTI.mdb
    DoCmd.TransferDatabase acLink, "Microsoft Access", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\CIVILIA_GESTIONE\CIVILIA_MDB\CIVILIA_N02_Tb01_ATTI.mdb", acTable, "CIVILIA_Tb01_ATTI", "CIVILIA_Tb01_ATTI", False
    
    DoCmd.TransferDatabase acLink, "Microsoft Access", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\CIVILIA_GESTIONE\CIVILIA_MDB\CIVILIA_N02_Tb01_ATTI.mdb", acTable, "CIVILIA_Tb01_ATTI_TMP", "CIVILIA_Tb01_ATTI_TMP", False
    
    DoCmd.TransferDatabase acLink, "Microsoft Access", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\CIVILIA_GESTIONE\CIVILIA_MDB\CIVILIA_N02_Tb01_ATTI.mdb", acTable, "CIVILIA_Tb01_}-------------------------------------------------@", "CIVILIA_Tb01_}-------------------------------------------------@", False
    
     
     '//esport tabelle da CIVILIA_N04_Tb07_EXPORT.mdb - todo: CAMBIARE LA PATH CON GLI OGGETTI ACCESS
    '//OGETTI ACCESS per ora sono sulla path di civilianext
    DoCmd.TransferDatabase acLink, "Microsoft Access", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\CIVILIA_GESTIONE\CIVILIA_MDB\CIVILIA_N04_Tb07_EXPORT.mdb", acTable, "MSys_DF00_}----------------------------------------------------@", "MSys_DF00_}----------------------------------------------------@", False
    DoCmd.TransferDatabase acLink, "Microsoft Access", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\CIVILIA_GESTIONE\CIVILIA_MDB\CIVILIA_N04_Tb07_EXPORT.mdb", acTable, "MSys_DF02_TIPO_OGGETTO", "MSys_DF02_TIPO_OGGETTO", False
    DoCmd.TransferDatabase acLink, "Microsoft Access", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\CIVILIA_GESTIONE\CIVILIA_MDB\CIVILIA_N04_Tb07_EXPORT.mdb", acTable, "MSys_DF10_DATABASE_TMP", "MSys_DF10_DATABASE_TMP", False
    DoCmd.TransferDatabase acLink, "Microsoft Access", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\CIVILIA_GESTIONE\CIVILIA_MDB\CIVILIA_N04_Tb07_EXPORT.mdb", acTable, "Msys_DF11_PROGETTI", "Msys_DF11_PROGETTI", False
    
    DoCmd.TransferDatabase acLink, "Microsoft Access", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\CIVILIA_GESTIONE\CIVILIA_MDB\CIVILIA_N04_Tb07_EXPORT.mdb", acTable, "Msys_DF12_TIPO_OGGETTI", "Msys_DF12_TIPO_OGGETTI", False
    
    DoCmd.TransferDatabase acLink, "Microsoft Access", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\CIVILIA_GESTIONE\CIVILIA_MDB\CIVILIA_N04_Tb07_EXPORT.mdb", acTable, "MSys_ELENCO_}-------------------------------------------------@", "MSys_ELENCO_}-------------------------------------------------@", False
    DoCmd.TransferDatabase acLink, "Microsoft Access", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\CIVILIA_GESTIONE\CIVILIA_MDB\CIVILIA_N04_Tb07_EXPORT.mdb", acTable, "MSys_ELENCO_ObjProgetti", "MSys_ELENCO_ObjProgetti", False
    
    
      '//ACCODA LINK DELLE DEFINIZIONI  "CIVILIA_DF01_}-------------------------------------------------@" _
                                        "CIVILIA_DF01_SETTORE" _
                                        "CIVILIA_DF02_TIPO_ATTO"
    
      
    DoCmd.TransferDatabase acLink, "Microsoft Access", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\CIVILIA_GESTIONE\CIVILIA_MDB\CIVILIA_N02_Tb01_ATTI.mdb", acTable, "CIVILIA_DF01_}-------------------------------------------------@", "CIVILIA_DF01_}-------------------------------------------------@", False
    DoCmd.TransferDatabase acLink, "Microsoft Access", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\CIVILIA_GESTIONE\CIVILIA_MDB\CIVILIA_N02_Tb01_ATTI.mdb", acTable, "CIVILIA_DF01_SETTORE", "CIVILIA_DF01_SETTORE", False
    DoCmd.TransferDatabase acLink, "Microsoft Access", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\CIVILIA_GESTIONE\CIVILIA_MDB\CIVILIA_N02_Tb01_ATTI.mdb", acTable, "CIVILIA_DF02_TIPO_ATTO", "CIVILIA_DF02_TIPO_ATTO", False
    
  
    
     
     
     
     
    '//CERCA TABELLA  ---> ULTIMA INSERITA
    DoCmd.RunMacro "CIVILIA_DF02_TIPO_ATTO"
    
    
   
    
    
    

COLLEGA_CIVILIA_TB01_Exit:
    Exit Function

COLLEGA_CIVILIA_TB01_Err:
    MsgBox Error$
    Resume COLLEGA_CIVILIA_TB01_Exit

End Function




'//@ASSEGNAZIONI  @LINK@TABELLE@CIVILIA
Public Function COLLEGA_CIVILIA_TB05(Blocco_b As Boolean)
On Error GoTo COLLEGA_CIVILIA_TB05_Err

'//COLLEGA TABELLA LLPP_DF06_FASCICOLI
'//-----------------------------------------------------------------------------//
'//NOTE: parametro Blocco_b = True procedura bloccata con False si esporta.
        '// Call COLLEGA_CIVILIA_TB05(False)

'//-----------------------------------------------------------------------------//


    
    If BloccoProcedura_b = True Then
        MsgBox "BLOCCO PROCEDURA DI COLLEGAMENTO LLPP_DF06_FASCICOLI per evitare la cancellazione della tabella dati! " & _
            "Impostare a False la variabile BloccoProcedura_b", vbExclamation, "MSG BLOCCO COLLEGAMENTO"
        GoTo COLLEGA_CIVILIA_TB05_Exit
    End If


    ' In caso di errore non viene eseguita l'istruzione elimina oggetto ma quella successiva
    On Error Resume Next
    ' viene prima cancellata con
    ' EliminaOggetto
    ' e poi riscotruito il collegamento con
    ' TrasferisciDatabase
    ' e cioè RICOSTRUISCE IL COLLEGAMENTO trasferisce database. Se invece esiste la tabella  ATTI DI IMPEGNO La stessa
    ' prima elimina la tabella collegata gia esistente
    DoCmd.DeleteObject acTable, "CIVILIA_Tb05_}-------------------------------------------------@"
    DoCmd.DeleteObject acTable, "CIVILIA_Tb05_ATTI_ASSEGNAZIONI"
    DoCmd.DeleteObject acTable, "CIVILIA_Tb05_ATTI_ASSEGNAZIONI_TMP"
    
    
    
    
    ' collego le tabelle atti
    DoCmd.TransferDatabase acLink, "Microsoft Access", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\CIVILIA_GESTIONE\CIVILIA_MDB\CIVILIA_N02_Tb05_ASSEGNAZIONI.mdb", acTable, "CIVILIA_Tb05_}-------------------------------------------------@", "CIVILIA_Tb05_}-------------------------------------------------@", False
    
    DoCmd.TransferDatabase acLink, "Microsoft Access", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\CIVILIA_GESTIONE\CIVILIA_MDB\CIVILIA_N02_Tb05_ASSEGNAZIONI.mdb", acTable, "CIVILIA_Tb05_ATTI_ASSEGNAZIONI", "CIVILIA_Tb05_ATTI_ASSEGNAZIONI", False
    
    DoCmd.TransferDatabase acLink, "Microsoft Access", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\CIVILIA_GESTIONE\CIVILIA_MDB\CIVILIA_N02_Tb05_ASSEGNAZIONI.mdb", acTable, "CIVILIA_Tb05_ATTI_ASSEGNAZIONI_TMP", "CIVILIA_Tb05_ATTI_ASSEGNAZIONI_TMP", False
    
    '//CERCA TABELLA  ---> SOSPESO
    DoCmd.RunMacro "CIVILIA_Tb05_ATTI_ASSEGNAZIONI_TMP"
    
    

COLLEGA_CIVILIA_TB05_Exit:
    Exit Function

COLLEGA_CIVILIA_TB05_Err:
    MsgBox Error$
    Resume COLLEGA_CIVILIA_TB05_Exit

End Function




'------------------------------------------------------------
' ESPORTA_Mcr01_01_Esporta_GESTIONE_AGGIORNA_DATI_ACCODA_NW_CODICI_ESPORTA  @ESPORTA@CSV@XLS
'
'------------------------------------------------------------
Function ESPORTA_Mcr01_01_Esporta_GESTIONE_AGGIORNA_DATI_ACCODA_NW_CODICI_ESPORTA()
On Error GoTo ESPORTA_Mcr01_01_Esporta_GESTIONE_AGGIORNA_DATI_ACCODA_NW_CODICI_ESPORTA_Err

    'AggiornaDati_ed_EsportaDati_Pfunct   TODO:DA FARE.
    
    


ESPORTA_Mcr01_01_Esporta_GESTIONE_AGGIORNA_DATI_ACCODA_NW_CODICI_ESPORTA_Exit:
    Exit Function

ESPORTA_Mcr01_01_Esporta_GESTIONE_AGGIORNA_DATI_ACCODA_NW_CODICI_ESPORTA_Err:
    MsgBox Error$
    Resume ESPORTA_Mcr01_01_Esporta_GESTIONE_AGGIORNA_DATI_ACCODA_NW_CODICI_ESPORTA_Exit

End Function




'//********************************************************//
'//
'// @INIZIO@ANALISI@FILE
'//
'//********************************************************//






'//*** INIZIO ***
'//APRI FILE @APRO@ANALISI@SUBLIME.TEXT  @CMD@03@CMD.03@CMD.3
'//========================================================================================================================================//

Private Sub Cmd_03_Click()

On Error GoTo Cmd_03_Err
    
      
      
      
                     
        
        '//APRO SORGENTE FORM FILTRO 66 CAMPI
        '//.....................................................................................................
        '//Note           : Tramite una Select vengono individuati i valori da restiuire.
                   
            '//reset
            '//Imposto i parametri
            ProceduraAttivaEseguita_s = ""
            ProceduraMessaggioErrore_s = ""
            
            MyFile_s = ""
            MyPath_s = ""
            
            '//Imposto i parametri
            ProceduraAttivaEseguita_s = ""
            ProceduraMessaggioErrore_s = "Errore nella procedura - apri FILE GESTIONE """ & MyPath_s & MyFile_s
            
            
        
                    
                '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF
                '//--------------------------------------------------------------------------------//--------//
                '//NOTE                 -> Apro il file di tipo doc, zip o pdf
                '//CODICE               -> Function ApriFile_Pfunct.01.01
                '//PARAMETRI            -> par_Path_s         = PATH _
                                        -> par_NameFile_s               = NOME FILE _
                                        -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                        
                        '//imposto i parametri - LA PATH
                        MyPath_s = "c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\"
                        '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
                        MyFile_s = "Project_PROGETTI_MDB.sublime-project "
                
                         '//chiamo la sub con i parametri
                            Call ApriFile_Pfunct(MyPath_s, MyFile_s, 0)
                        
                '//--------------------------------------------------------------------------------//--------//
                        
                 
          
 

           
        '//*** fine ***
        '//APRO IL FILE PASSANDO COME PARAMETRO PATH E NOME
        '//.....................................................................................................
 
                    







        

'//*** fine ***
'//APRI FILE @APRO@ANALISI@SUBLIME.TEXT
'//###########################################################################################################################//


        
    '//........................................................//


'USCITA  E GESTIONE ERRORI
'..............................................................................................................

Cmd_03_Exit:
    Exit Sub

Cmd_03_Err:
    MsgBox Error$
    Resume Cmd_03_Exit

End Sub


'//*** FINE ***
'//APRI PATH
'//========================================================================================================================================//





'//*** INIZIO ***
'//APRI PROTOCOLLO
'//========================================================================================================================================//

Private Sub Cmd_APRI_PROTOCOLLO_Click()
On Error GoTo Cmd_APRI_PROTOCOLLO_Click_Err
    
    DoCmd.OpenForm "CIVILIA_FrmTb06_00_GE_PROTOCOLLO"
    
    
    

'USCITA  E GESTIONE ERRORI
'..............................................................................................................

Cmd_APRI_PROTOCOLLO_Click_Exit:
    Exit Sub

Cmd_APRI_PROTOCOLLO_Click_Err:
    MsgBox Error$ & " FORM PROTOCOLLO NON ESISTENTE - CARICARLO NEL DB CORRENTE!! "
    
    Resume Cmd_APRI_PROTOCOLLO_Click_Exit

End Sub


'//*** FINE ***
'//APRI PROTOCOLLO
'//========================================================================================================================================//



'//DELETE OGGETTI PROTOCOLLO
Private Sub Cmd_DELETE_PROTOCOLLO_Click()

On Error GoTo Cmd_DELETE_PROTOCOLLO_Click_Err

        '//CANCELLA LE FORM PROTOCOLLO
        '//@CANCELLA@FORM@GESTIONE_PROTOCOLLO  @FORM@PROTOCOLLO@MASTER
        DoCmd.DeleteObject acForm, "CIVILIA_FrmTb06_00_GE_PROTOCOLLO"
        DoCmd.DeleteObject acForm, "CIVILIA_FrmTb06_01_PROTOCOLLO"
        DoCmd.DeleteObject acForm, "CIVILIA_FrmTb06_01_PROTOCOLLO_TMP"
        DoCmd.DeleteObject acForm, "CIVILIA_FrmTb06_01_PROTOCOLLO_WORK_TMP"
        DoCmd.DeleteObject acForm, "CIVILIA_FrmTb06_11_GRUPPO_ANNO_PROTOCOLLO"
        DoCmd.DeleteObject acForm, "CIVILIA_FrmTb06_12_GRUPPO_DETTAGLIO_PROTOCOLLI"
        
        '//CANCELLO LE QUERY
        DoCmd.DeleteObject acQuery, "CIVILIA_QryTb06_01_SELECT_PROTOCOLLO"
        
        DoCmd.DeleteObject acQuery, "CIVILIA_QryTb06_01_SELECT_PROTOCOLLO_TMP"
        DoCmd.DeleteObject acQuery, "CIVILIA_QryTb06_01_SELECT_PROTOCOLLO_WORK_TMP"
        DoCmd.DeleteObject acQuery, "CIVILIA_QryTb06_11_GRUPPO_PROTOCOLLO_ANNO"
        DoCmd.DeleteObject acQuery, "CIVILIA_QryTb06_12_GRUPPO_DETTAGLIO_PROTOCOLLI"
        
        
        '//@CANCELLA@MODULI@PROTOCOLLI
        
        DoCmd.DeleteObject acModule, "CIVILIA_FrmTb06_00_GE_PROTOCOLLO_Mdl_00_}----------------------@"
        DoCmd.DeleteObject acModule, "CIVILIA_FrmTb06_01_GE_PROTOCOLLO_Mdl_01_MENU"
        
        MsgBox "DELETE OGGETTI PROTOCOLLI - TUTTI", vbExclamation, "CANCELLAZIONE OGGETTI PROTOCOLLO"
        
'USCITA  E GESTIONE ERRORI
'..............................................................................................................

Cmd_DELETE_PROTOCOLLO_Click_Exit:
    Exit Sub

Cmd_DELETE_PROTOCOLLO_Click_Err:
    MsgBox Error$ & " ATTENZIONE NON CI SONO QUERY O FORM DI GESTIONE DEL PROTOCOLLO - CARICARE GLI OGGETTI NEL DB CORRENTE!!"
    Resume Cmd_DELETE_PROTOCOLLO_Click_Exit
     
End Sub






'// HELP SUL COLLEGAMENTO DELLE TABELLE CIVILIA
'//===================================================================================================================//'
'//CODICI: @HELP@CMD@LINK.TABELLE@CIVILIA   @TABELLE@CIVILIA@LINK@AIUTO -
    Private Sub Cmd_HelpLinkCivilia_Click()

    On Error GoTo Cmd_HelpLinkCivilia_Click_Err



                '//CHIAMA LA PROCEDURA PRI FILE - @APRI.FILE
                '//-----------------------------------------------------------------------------------------------//
                '//apri file par1 = file; par2 = path
                    '//
                    ApriFile_Sub "HelpLinkCivilia.md", _
                    "C:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\CIVILIA_GESTIONE\CIVILIA_TEXT_ANALISI\CIVILIA_ANALISI_SUBLIME_TEXT\HELP\"
                
                
                '//-----------------------------------------------------------------------------------------------//
             

    'USCITA  E GESTIONE ERRORI
    '..............................................................................................................

Cmd_HelpLinkCivilia_Click_Exit:
        Exit Sub

Cmd_HelpLinkCivilia_Click_Err:
        MsgBox Error$
        Resume Cmd_HelpLinkCivilia_Click_Exit


    End Sub

'// HELP SUL COLLEGAMENTO DELLE TABELLE CIVILIA
'//===================================================================================================================//'


'// HELP SUL COLLEGAMENTO DELLA TABELLA 5
'//===================================================================================================================//'
'//CODICI: @HELP@CMD@LINK.TB5   @TB5@LIK@AIUTO -
    Private Sub Cmd_HelpLinkTb5_Click()

    On Error GoTo Cmd_HelpLinkTb5_Click_Err



                '//CHIAMA LA PROCEDURA PRI FILE - @APRI.FILE
                '//-----------------------------------------------------------------------------------------------//
                '//apri file par1 = file; par2 = path
                    '//
                    ApriFile_Sub "HelpLinkTb5.md", _
                    "C:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\CIVILIA_GESTIONE\CIVILIA_TEXT_ANALISI\CIVILIA_ANALISI_SUBLIME_TEXT\HELP\"
                
                
                '//-----------------------------------------------------------------------------------------------//
             

    'USCITA  E GESTIONE ERRORI
    '..............................................................................................................

Cmd_HelpLinkTb5_Click_Exit:
        Exit Sub

Cmd_HelpLinkTb5_Click_Err:
        MsgBox Error$
        Resume Cmd_HelpLinkTb5_Click_Exit


    End Sub



'// HELP SUL PROGETTO DI SUBLIME TEXT
'//===================================================================================================================//'
'//CODICI: @HELP@SUBLIMETEXT.PROGETTO

    Private Sub Cmd_HelpProgettoSublimeText_TUTTO_Click()

    On Error GoTo Cmd_HelpProgettoSublimeText_TUTTO_Click_Err



                '//CHIAMA LA PROCEDURA PRI FILE - @APRI.FILE
                '//-----------------------------------------------------------------------------------------------//
                '//apri file par1 = file; par2 = path
                    '//
                    ApriFile_Sub "HelpLinkCivilia.sublime-project", _
                    "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\CIVILIA_GESTIONE\CIVILIA_TEXT_ANALISI\CIVILIA_ANALISI_SUBLIME_TEXT\HELP\"
                
                
                '//-----------------------------------------------------------------------------------------------//
             

    'USCITA  E GESTIONE ERRORI
    '..............................................................................................................

Cmd_HelpProgettoSublimeText_TUTTO_Click_Exit:
        Exit Sub

Cmd_HelpProgettoSublimeText_TUTTO_Click_Err:
        MsgBox Error$
        Resume Cmd_HelpProgettoSublimeText_TUTTO_Click_Exit


    End Sub

'// HELP SUL PROGETTO DI SUBLIME TEXT
'//===================================================================================================================//'




'// HELP SUL COLLEGAMENTO DELLA TABELLA 5  *** FINE ***
'//===================================================================================================================//'


'//COLLEGA TABELLA PROTOCOLLO IMPORTANTO I LINK DAL DATABASE ESTERNO
'//========================================================================================================================================//

'//@PROTOCOLLO@CIVILIA @IMPORTA@TABELLE@PROTOCOLLO
Private Sub CmdProtocollo_Link_Click()

On Error GoTo CmdProtocollo_Link_Click_Err




'//COLLEGA TABELLE PROTOCOLLO CIVILIA
'//-----------------------------------------------------------------------------//
'//NOTE: parametro Blocco_b = True procedura bloccata con False si esporta.
        '// Call CmdProtocollo_Link_Click

'//-----------------------------------------------------------------------------//


    
    If BloccoProcedura_b = True Then
        MsgBox "BLOCCO PROCEDURA DI COLLEGAMENTO LLPP_DF06_FASCICOLI per evitare la cancellazione della tabella dati! " & _
            "Impostare a False la variabile BloccoProcedura_b", vbExclamation, "MSG BLOCCO COLLEGAMENTO"
        GoTo CmdProtocollo_Link_Click_Exit
    End If


    ' In caso di errore non viene eseguita l'istruzione elimina oggetto ma quella successiva
    On Error Resume Next
    ' viene prima cancellata con
    ' EliminaOggetto
    ' e poi riscotruito il collegamento con
    ' TrasferisciDatabase
    ' e cioè RICOSTRUISCE IL COLLEGAMENTO trasferisce database. Se invece esiste la tabella  ATTI DI IMPEGNO La stessa
    ' prima elimina la tabella collegata gia esistente
    DoCmd.DeleteObject acTable, "CIVILIA_Tb06_}-------------------------------------------------@"
    DoCmd.DeleteObject acTable, "CIVILIA_Tb06_PROTOCOLLO"
    DoCmd.DeleteObject acTable, "CIVILIA_Tb06_PROTOCOLLO_TMP"
    DoCmd.DeleteObject acTable, "CIVILIA_Tb06_PROTOCOLLO_WORK_TMP"
    
    
    
    
    ' collego le tabelle atti  @COLLEGA@TABELLE@PROTOCOLLI
    Dim PathDb_s As String
    Dim Tabella_s As String
    PathDb_s = "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\CIVILIA_GESTIONE\CIVILIA_MDB\CIVILIA_N03_Tb06_PROTOCOLLO.mdb"
    
    Tabella_s = "CIVILIA_Tb06_}-------------------------------------------------@"
    DoCmd.TransferDatabase acLink, "Microsoft Access", PathDb_s, acTable, Tabella_s, Tabella_s, False
    
    
    Tabella_s = "CIVILIA_Tb06_PROTOCOLLO"
    DoCmd.TransferDatabase acLink, "Microsoft Access", PathDb_s, acTable, Tabella_s, Tabella_s, False
    
    
    Tabella_s = "CIVILIA_Tb06_PROTOCOLLO_TMP"
    DoCmd.TransferDatabase acLink, "Microsoft Access", PathDb_s, acTable, Tabella_s, Tabella_s, False
    
    Tabella_s = "CIVILIA_Tb06_PROTOCOLLO_WORK_TMP"
    DoCmd.TransferDatabase acLink, "Microsoft Access", PathDb_s, acTable, Tabella_s, Tabella_s, False
    
    '//CERCA TABELLA  ---> SOSPESO
    DoCmd.RunMacro "CIVILIA_Tb05_ATTI_ASSEGNAZIONI_TMP"
    
    

'USCITA  E GESTIONE ERRORI
'..............................................................................................................

CmdProtocollo_Link_Click_Exit:
    Exit Sub

CmdProtocollo_Link_Click_Err:
    MsgBox Error$
    Resume CmdProtocollo_Link_Click_Exit

End Sub


'//*** FINE ***
'//COLLEGA TABELLA PROTOCOLLO IMPORTANTO I LINK DAL DATABASE ESTERNO
'//========================================================================================================================================//






'//IMPORTA FORM E QUERY DI GESTIONE DEL PROTOCOLLO DAL DATABASE ESTERNO
'//========================================================================================================================================//
'//Tipo           : Routine pubblica.
'//Attività       : IMPORTA GLI OGGETTI FORM
'//Note           : IL DATABASE E' ESTERNO
'//Parametro      : null
'//Codice         : Btn_IMPORTPROTOCOLLO_Click.01
'//
'//@IMPORTA@FORM@PROTOCOLLO,  @importazione@da.un.dabase.esterno, @IMPORTA@OGGETTI@DB.ESTERNO
Private Sub Btn_IMPORTPROTOCOLLO_Click()
    
        '//....
On Error GoTo Err_Btn_IMPORTPROTOCOLLO_Click
    
    
    '//IMPORTA OGGETTO DATABASE DALL'ESTERNO  ---> CODICE : Btn_IMPORTPROTOCOLLO_Click.01.CALL
    '//--------------------------------------------------------------------------------------------//
    '//Il metodo TransferDatabase esegue l'azione TransferDatabase in Visual Basic.
    '//Sintassi _
    espressione.TransferDatabase(TipoTrasferimento, TipoDatabase, NomeDatabase, TipoOggetto, Origine, Destinazione, SoloStruttura, AccessoArchivio) _
    espressione   Variabile che rappresenta un oggetto DoCmd.
        
        '//VARIABILI PER L'IMPORTAZIONE DELLA FORM
        Dim DatabaseOrigine_s As String
        Dim NomeOggettoOrigine_s As String
        Dim NomeOggettoFine_s As String
        
        '//IMPOSTO LA VARIABILE DEL DB
        '//database origine path+nome.mdb in cui prelevare gli oggetti
            DatabaseOrigine_s = "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\CIVILIA_GESTIONE\CIVILIA_MDB\CIVILIA_N03_Tb06_PROTOCOLLO.mdb"
            
        
        '//IMPORTA LE FORM DEL PROTOCOLLO  @IMPORTA@FORM@PROTOCOLLI
        '//-----------------------------------------------------------------------------------------//
            '//01_LA FORM MASTER DI GESTIONE DEL PROTOCOLLO
                
                '//nome dell'oggetto da importare nel db esterno
                NomeOggettoOrigine_s = "CIVILIA_FrmTb06_00_GE_PROTOCOLLO"
                '//nuovo nome dell'oggetto importato dall'esterno
                NomeOggettoFine_s = "CIVILIA_FrmTb06_00_GE_PROTOCOLLO"
                 
                '//@IMPORTA.OGGETTI.DATABASE
                DoCmd.TransferDatabase acImport, "Microsoft Access", _
                    DatabaseOrigine_s, acForm, NomeOggettoOrigine_s, _
                    NomeOggettoFine_s
                  
        '//02_LA_FORM_PROTOCOLLO
                
                '//nome dell'oggetto da importare nel db esterno
                NomeOggettoOrigine_s = "CIVILIA_FrmTb06_01_PROTOCOLLO"
                '//nuovo nome dell'oggetto importato dall'esterno
                NomeOggettoFine_s = "CIVILIA_FrmTb06_01_PROTOCOLLO"
                 
                '//@IMPORTA.OGGETTI.DATABASE
                DoCmd.TransferDatabase acImport, "Microsoft Access", _
                    DatabaseOrigine_s, acForm, NomeOggettoOrigine_s, _
                    NomeOggettoFine_s
      
      
                  
        '//03_LA_FORM_PROTOCOLLO_TMP
                
                '//nome dell'oggetto da importare nel db esterno
                NomeOggettoOrigine_s = "CIVILIA_FrmTb06_01_PROTOCOLLO_TMP"
                '//nuovo nome dell'oggetto importato dall'esterno
                NomeOggettoFine_s = "CIVILIA_FrmTb06_01_PROTOCOLLO_TMP"
                 
                '//@IMPORTA.OGGETTI.DATABASE
                DoCmd.TransferDatabase acImport, "Microsoft Access", _
                    DatabaseOrigine_s, acForm, NomeOggettoOrigine_s, _
                    NomeOggettoFine_s
      
            
        '//04_LA_FORM_PROTOCOLLO_WORK_TMP
                
                '//nome dell'oggetto da importare nel db esterno
                NomeOggettoOrigine_s = "CIVILIA_FrmTb06_01_PROTOCOLLO_WORK_TMP"
                '//nuovo nome dell'oggetto importato dall'esterno
                NomeOggettoFine_s = "CIVILIA_FrmTb06_01_PROTOCOLLO_WORK_TMP"
                 
                '//@IMPORTA.OGGETTI.DATABASE
                DoCmd.TransferDatabase acImport, "Microsoft Access", _
                    DatabaseOrigine_s, acForm, NomeOggettoOrigine_s, _
                    NomeOggettoFine_s
        
        '//04_LA_FORM_GRUPPO_ANNO_PROTOCOLLO
                
                '//nome dell'oggetto da importare nel db esterno
                NomeOggettoOrigine_s = "CIVILIA_FrmTb06_11_GRUPPO_ANNO_PROTOCOLLO"
                '//nuovo nome dell'oggetto importato dall'esterno
                NomeOggettoFine_s = "CIVILIA_FrmTb06_11_GRUPPO_ANNO_PROTOCOLLO"
                 
                '//@IMPORTA.OGGETTI.DATABASE
                DoCmd.TransferDatabase acImport, "Microsoft Access", _
                    DatabaseOrigine_s, acForm, NomeOggettoOrigine_s, _
                    NomeOggettoFine_s
      
        '//05_LA_FORM_GRUPPO_DETTAGLIO_PROTOCOLLI
                
                '//nome dell'oggetto da importare nel db esterno
                NomeOggettoOrigine_s = "CIVILIA_FrmTb06_12_GRUPPO_DETTAGLIO_PROTOCOLLI"
                '//nuovo nome dell'oggetto importato dall'esterno
                NomeOggettoFine_s = "CIVILIA_FrmTb06_12_GRUPPO_DETTAGLIO_PROTOCOLLI"
                 
                '//@IMPORTA.OGGETTI.DATABASE
                DoCmd.TransferDatabase acImport, "Microsoft Access", _
                    DatabaseOrigine_s, acForm, NomeOggettoOrigine_s, _
                    NomeOggettoFine_s
                    
          '//06_CIVILIA_FrmTb06_12_GRUPPO_DETTAGLIO_PROTOCOLLI
                
                '//nome dell'oggetto da importare nel db esterno
                NomeOggettoOrigine_s = "CIVILIA_FrmTb06_12_GRUPPO_DETTAGLIO_PROTOCOLLI"
                '//nuovo nome dell'oggetto importato dall'esterno
                NomeOggettoFine_s = "CIVILIA_FrmTb06_12_GRUPPO_DETTAGLIO_PROTOCOLLI"
                 
                '//@IMPORTA.OGGETTI.DATABASE
                DoCmd.TransferDatabase acImport, "Microsoft Access", _
                    DatabaseOrigine_s, acForm, NomeOggettoOrigine_s, _
                    NomeOggettoFine_s

      
        '//-----------------------------------------------------------------------------------------//
            
        '//IMPORTA LE QUERY DI GESTIONE DELLE FORM
        '//-----------------------------------------------------------------------------------------//
                '//01
                '//nome dell'oggetto da importare nel db esterno
                NomeOggettoOrigine_s = "CIVILIA_QryTb06_01_SELECT_PROTOCOLLO"
                '//nuovo nome dell'oggetto importato dall'esterno
                NomeOggettoFine_s = "CIVILIA_QryTb06_01_SELECT_PROTOCOLLO"
                 
                '//@IMPORTA.OGGETTI.DATABASE
                DoCmd.TransferDatabase acImport, "Microsoft Access", _
                    DatabaseOrigine_s, acQuery, NomeOggettoOrigine_s, _
                    NomeOggettoFine_s
             
             
                '//02
                '//nome dell'oggetto da importare nel db esterno
                NomeOggettoOrigine_s = "CIVILIA_QryTb06_01_SELECT_PROTOCOLLO_TMP"
                '//nuovo nome dell'oggetto importato dall'esterno
                NomeOggettoFine_s = "CIVILIA_QryTb06_01_SELECT_PROTOCOLLO_TMP"
                 
                '//@IMPORTA.OGGETTI.DATABASE
                DoCmd.TransferDatabase acImport, "Microsoft Access", _
                    DatabaseOrigine_s, acQuery, NomeOggettoOrigine_s, _
                    NomeOggettoFine_s
             

                '//03
                '//nome dell'oggetto da importare nel db esterno
                NomeOggettoOrigine_s = "CIVILIA_QryTb06_01_SELECT_PROTOCOLLO_WORK_TMP"
                '//nuovo nome dell'oggetto importato dall'esterno
                NomeOggettoFine_s = "CIVILIA_QryTb06_01_SELECT_PROTOCOLLO_WORK_TMP"
                 
                '//@IMPORTA.OGGETTI.DATABASE
                DoCmd.TransferDatabase acImport, "Microsoft Access", _
                    DatabaseOrigine_s, acQuery, NomeOggettoOrigine_s, _
                    NomeOggettoFine_s
             
             
                '//04
                '//nome dell'oggetto da importare nel db esterno
                NomeOggettoOrigine_s = "CIVILIA_QryTb06_11_GRUPPO_PROTOCOLLO_ANNO"
                '//nuovo nome dell'oggetto importato dall'esterno
                NomeOggettoFine_s = "CIVILIA_QryTb06_11_GRUPPO_PROTOCOLLO_ANNO"
                 
                '//@IMPORTA.OGGETTI.DATABASE
                DoCmd.TransferDatabase acImport, "Microsoft Access", _
                    DatabaseOrigine_s, acQuery, NomeOggettoOrigine_s, _
                    NomeOggettoFine_s
             
                
                '//05
                '//nome dell'oggetto da importare nel db esterno
                NomeOggettoOrigine_s = "CIVILIA_QryTb06_12_GRUPPO_DETTAGLIO_PROTOCOLLI"
                '//nuovo nome dell'oggetto importato dall'esterno
                NomeOggettoFine_s = "CIVILIA_QryTb06_12_GRUPPO_DETTAGLIO_PROTOCOLLI"
                 
                '//@IMPORTA.OGGETTI.DATABASE
                DoCmd.TransferDatabase acImport, "Microsoft Access", _
                    DatabaseOrigine_s, acQuery, NomeOggettoOrigine_s, _
                    NomeOggettoFine_s
           
        '//IMPORTA LE QUERY DI GESTIONE DELLE FORM   *** FINE ***
        '//-----------------------------------------------------------------------------------------//
        
        
        
        '//IMPORTA@MODULI@DB@ESTERNO
        '//-----------------------------------------------------------------------------------------//
       
                  
                '//01
                '//nome dell'oggetto da importare nel db esterno
                NomeOggettoOrigine_s = "CIVILIA_FrmTb06_00_GE_PROTOCOLLO_Mdl_00_}----------------------@"
                '//nuovo nome dell'oggetto importato dall'esterno
                NomeOggettoFine_s = "CIVILIA_FrmTb06_00_GE_PROTOCOLLO_Mdl_00_}----------------------@"
                 
                '//@IMPORTA.OGGETTI.DATABASE
                DoCmd.TransferDatabase acImport, "Microsoft Access", _
                    DatabaseOrigine_s, acModule, NomeOggettoOrigine_s, _
                    NomeOggettoFine_s
             
                  
                '//02
                '//nome dell'oggetto da importare nel db esterno
                NomeOggettoOrigine_s = "CIVILIA_FrmTb06_01_GE_PROTOCOLLO_Mdl_01_MENU"
                '//nuovo nome dell'oggetto importato dall'esterno
                NomeOggettoFine_s = "CIVILIA_FrmTb06_01_GE_PROTOCOLLO_Mdl_01_MENU"
                 
                '//@IMPORTA.OGGETTI.DATABASE
                DoCmd.TransferDatabase acImport, "Microsoft Access", _
                    DatabaseOrigine_s, acModule, NomeOggettoOrigine_s, _
                    NomeOggettoFine_s
           
            '//CERCA FORM IMPORTATA  @AZIONE@CERCA@FORM@NON.APERTA
                DoCmd.SelectObject acForm, "CIVILIA_FrmTb06_12_GRUPPO_DETTAGLIO_PROTOCOLLI", True
               
            
      '//--------------------------------------------------------------------------------------------//

Exit_Btn_IMPORTPROTOCOLLO_Click:
    Exit Sub

Err_Btn_IMPORTPROTOCOLLO_Click:
 '//-------------------------------------------------------------------------------
    MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & " Procedura -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Stop
    Resume Exit_Btn_IMPORTPROTOCOLLO_Click
'//-------------------------------------------------------------------------------
        

End Sub
'//IMPORTA FORM E QUERY DI GESTIONE DEL PROTOCOLLO DAL DATABASE ESTERNO  *** FINE ***
'//========================================================================================================================================//




'//********************************************************//
'//
'// @FINE@GESTIONE@PROTOCOLLO
'//
'//********************************************************//












'//*** INIZIO ***
'//APRI FILE   @APRI.FILE  @APRO.IL.FILE
'//========================================================================================================================================//

Private Sub ApriFile_Sub(par_MyFile_s As String, par_MyPath_s As String)

On Error GoTo Cmd_03_Err
    
    
    
            '//CHIAMA LA PROCEDURA PRI FILE - @APRI.FILE
            '//-----------------------------------------------------------------------------------------------//
            '//apri file par1 = file; par2 = path
                '//ApriFile_Sub "CIVILIA_SUBLIME_TEXT_ANALISI.sublime-project", _
                "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\CIVILIA_GESTIONE\CIVILIA_TEXT_ANALISI\CIVILIA_ANALISI_SUBLIME_TEXT\PROGETTO_SUBLIME_TEXT\"
            
            
            '//-----------------------------------------------------------------------------------------------//
            
            
            
             
            '//Imposto i parametri
            ProceduraAttivaEseguita_s = "ApriFile_Sub"
            ProceduraMessaggioErrore_s = "Errore nella procedura - apri FILE GESTIONE """ & MyPath_s & MyFile_s
            
    
    
        '//AGGIORNAMENTO_DEL_15_07_2020
        '//APRO LA PATH DEGLI ATTI DI GESTIONE -
        '//###########################################################################################################################//
        '//NOTE: apro il file .mdb _
        non ci sono parametri _
        AGGIORNAMENTO DEL = 13.07.2020
        '//CODICE----------->APRO_PATH_Tb01_ATTI_DI_GESTIONE_pSub
        
        
        
            
            
        
                    
        
         
        
        '//APRO SORGENTE FORM FILTRO 66 CAMPI
        '//.....................................................................................................
        '//Note           : Tramite una Select vengono individuati i valori da restiuire.
                   
            '//reset
            '//Imposto i parametri
            ProceduraAttivaEseguita_s = ""
            ProceduraMessaggioErrore_s = ""
            
            MyFile_s = ""
            MyPath_s = ""
            '//Imposto i parametri
            ProceduraAttivaEseguita_s = "ApriFile_Sub"
            ProceduraMessaggioErrore_s = "Errore nella procedura - apri FILE GESTIONE """ & MyPath_s & MyFile_s
            
            
        
                    
                '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF
                '//--------------------------------------------------------------------------------//--------//
                '//NOTE                 -> Apro il file di tipo doc, zip o pdf
                '//CODICE               -> Function ApriFile_Pfunct.01.01
                '//PARAMETRI            -> par_Path_s         = PATH _
                                        -> par_NameFile_s               = NOME FILE _
                                        -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                
                
                         '//chiamo la sub con i parametri
                            Call ApriFile_Pfunct(MyPath_s, MyFile_s, 0)
                        
                '//--------------------------------------------------------------------------------//--------//
                        
                 
          
             
            
     
        
        '//*** fine ***
        '//APRO SORGENTE FORM FILTRO 66 CAMPI
        '//.....................................................................................................
 
                    







        

'//*** fine ***
'//APRO LA PATH DEGLI ATTI DI GESTIONE -
'//###########################################################################################################################//


        
    '//........................................................//


'USCITA  E GESTIONE ERRORI
'..............................................................................................................

Cmd_03_Exit:
    Exit Sub

Cmd_03_Err:
    MsgBox Error$
    Resume Cmd_03_Exit

End Sub


'//*** FINE ***
'//APRI FILE
'//========================================================================================================================================//

