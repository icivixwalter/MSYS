VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_UTILITA_Msys_Frm21_IMPORTA_OGGETTI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'# FORM_(Form_UTILITA_Msys_Frm21_IMPORTA_OGGETTI)_Classe.md


'*****************************************************************************************************//
'*  CLASSE DELLA FORM: Form_UTILITA_Msys_Frm21_IMPORTA_OGGETTI
'*
'*  ATTIVITA        :questa classe della @form@IMPORTA@PROGETTO viene utilizzata per importare
'*                  tutti gli oggetti (tabelle, query, form, report ecc..) che in questo caso appartengono
'*                  al progetto MSYS_N01_OGGETTI.mdb ma puo essere adattata per importare
'*                  altri oggetti che appartengono ad al progetti access.
'*                  VANTAGGIO: è dato dal fatto che gli oggetti tabelle query ecc... NON SONO SALVATI
'*                  in un tabella anchessa da importare ma sono salvati all'interno della classe in
'*                  in diverse collezioni precaricate. In questo puo spostare tutta form master
'*                  in qualsiasi progetto mdb senza preoccuparti di portare dietro tabelle oggetti, query
'*                  ed altre form di gestione. L'importazione degli oggetti RISULTA SEMPLICATA.
'*
'*  FUNZIONAMENTO   : Il funzionamento si basa su queste collezione di oggetti precaricati che  possono
'*                  essere sostituite:
'*
'*                  @PARAMETRO_01 = LA PATH ED IL FILE.MDB DOVE SI TROVANO GLI OGGETTI
'*
'*                  @PARAMETRO_02= COLLECTION DELLE FORMS DA IMPORTARE
'*                              I° COLLEZIONE : OGGETTI FORM PRECARICATI PRONTI PER L'IMPORTAZIONE
'*                              VAI AL CODICE: '//CODICE ----> COLLECTION_FORM_PRECARICATE
'*
'*  APRI ANALISI    : puoi aprire l'analisi di questa classe con questo comando
'*                  START "APRI ANALISI GESTIONE DB ESTERNO " call "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\Project_MSYSDbEst_GE_OBJECT_ANALISI.sublime-project"
'*
'*  LIBRERIE
'*      Note
'*              Per il funzionamento occorrono le seguenti librerie:
'*              Dim db As DAO.Database  =
'*                   Microsoft DAO 3.6 Object Library                           =
'*                          C:\Program Files (x86)\Common Files\microsoft shared\DAO\dao360.dll
'*                   Microsoft Visual Basic for Application Extensibility 5.3   =
'*                          C:\Program Files (x86)\Common Files\microsoft shared\VBA\VBA6\VBE6EXT.olb
'*                   Microsoft Activex Data Object 2.1 Library                  =
'*                          C:\Program Files (x86)\Common Files\System\ado\msado21.tlb
    
'*  ANALISI ACCESS  : L'analisi generale di access si apre con questo comando
'*                      START "APRI ANALISI GENERALE DI ACCESS" call "c:\Casa\LINGUAGGI\ACCESS\Project_ACCESS.sublime-project"
'*  FAQ             :
'*                  COME IMPORTARE GLI OGGETTI TABLE, QUERY, FORM IN UN PROGETTO?
'*                  CLASSE IMPORTA PROGETTI IN UN DB ESTERNO
'*                  @FORM@IMPORTA@OGGETTI_(Questa form permette l'importazione degli oggetti table, query ecc. in modo autonomo)
'*
'*
'*****************************************************************************************************//
Option Compare Database
    
'//DIM LE VARIABILI
'//-----------------------------------------------------------------------------------//
    '//APRI FILE
   
' Dim objectTables As Collection
   
    '//DATABASE PATH SORGENTE, CONTATORI
    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim sourceDB As DAO.Database
    'DAO
    Dim daoDB As DAO.Database
    Dim dbCurrent As DAO.Database

    Dim DaoWks As DAO.Workspace
    Dim daoRS As DAO.Recordset
    Dim DaoRs_Lett As DAO.Recordset
    Dim tdf As DAO.TableDef

    
    'ADO
    Dim ADODB As Database
    Dim AdodaoRs As Recordset
    
    'Contatori
    Dim iCount As Integer
    Dim iTOTcount As Integer
    Dim dbl_count As Double
    

    
    '//DIM input box
    Dim Message As Variant
    Dim Title As Variant
    Dim Default As Variant
    Dim MyValue As Variant
    Dim MyMsg_V As Variant              '//risultato se 0 exit 1 continua procedura se annulla stringa = ""
    
    
    'Variabili generali
    Dim i As Integer
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
            
        
    ' // DIM variabili procedura importa e collega DLL  @DIM@DLL
    Dim PathDb_s As String
    
    ' // DIM MATRICI DINAMINCHE: senza valore - Array globali per i nomi degli oggetti, SOSTITUITO DA MATRICI STATICHE CON VALORE _
        A MATRICI DIMINAMICHE senza valore per essere ridimensionate : QueryArray() ecc..
    Dim QueryArray() As String
    Dim TableArray() As String
    Dim ModuleArray() As String
    Dim MacroArray() As String

    
    '//DIM LE VARIABILI DEL PROGETTO ORIGINALE CURRENT PROJECT DEL DB ESTERNO DA LEGGERE
    '//........................................................................//
         Dim myCurrentProject As String
         Dim myOriginaleProject As String
          '//PATH E FILE DOVE IMPORTARE
         Dim sourceDBPath_s As String                       '//path del db esterno
         
         Dim PROGETTO_s As String                           '//nome del progetto attivato
         
    '//........................................................................//
    
    
    
   '//@DIM@COLLECTION@TABLE COLLEZIONI TABELLE
   '//...................................
      
    '//DIM IMPORTA TABELLE
    Dim tableName As String
    Dim tblCiclo As TableDef
    Dim mytableName_s As String
    Dim objectTables As Collection              '//todo: doppia collezione di tabelle eliminare objectTables
    Dim tablesCollection As Collection
    
    Dim varItem As Variant
    '//i contatori delle tabelle
    Dim tableExistsInCollectionTables_Funct_b As Boolean
    Dim importedTablesCount_i As Integer
    Dim deletedTablesCount_i As Integer
    Dim deletedTables_s As String
       
   

    
   
    '//...................................

   '//@DIM@COLLECTION@FORM COLLEZIONI FORM
   '//...................................
     
        '//DIM IMPORTA FORM
        Dim myformName_s As String                      '//nome form da importare
        Dim formsToImport As Variant                    '//calcolo form da importare
        '//DIM della collezione form da importare
        Dim formsCollection As Collection                '//collezione delle form prestabilita
        
        '// i contatori delle form
        Dim deletedForms_s As String
        Dim deletetFormsCount_i As Integer
        Dim importedForms_s As String
        Dim importedFormsCount_i As Integer
    
        
   '//...................................


   '//@DIM@COLLECTION@QUERY COLLEZIONI QUERY
   '//...................................
     
        '//DIM IMPORTA QUERY
        Dim myNameQuery_s As String                         '//nome query da importare
        Dim queryesToImport As Variant                      '//calcolo query da importare
        '//DIM della collezione form da importare
        Dim queryesCollection As Collection                   '//collezione delle query prestabilita
        
        '// i contatori delle form
        Dim deleteNameQueryes_s As String
        Dim deleteQueryesCount_i As Integer
        Dim importedNameQueryes_s As String
        Dim importedQueryesCount_i As Integer
    
        
   '//...................................




   '//@DIM@COLLECTION@REPORT COLLEZIONI REPORTS
   '//...................................
     
        '//DIM IMPORTA QUERY
        Dim myreportName_s As String                         '//nome query da importare
        Dim reportsToImport As Variant                      '//calcolo REPORT da importare
        '//DIM della collezione form da importare
        Dim reportsCollection As Collection                   '//collezione delle REPORT prestabilita
        
        '// i contatori delle form
        Dim deletedReports_s As String
        Dim deletetReportesCount_i As Integer
        Dim importedReports_s As String
        Dim importedReportesCount_i As Integer
    
        
   '//...................................







    '//@DIM@COLLECTION@MACRO COLLEZIONI MACRO
    '//...................................
      
         '//DIM IMPORTA MACRO
         Dim myMacroName_s As String                        '//nome MACRO da importare
         Dim macrosToImport As Variant                      '//calcolo MACRO da importare
         '//DIM della collezione MACRO da importare
         Dim macrosCollection As Collection                 '//collezione delle MACRO prestabilita
         
         '// i contatori delle MACRO
         Dim deletedMacro_s As String
         Dim deletetMacroCount_i As Integer
         Dim importedMacro_s As String
         Dim importedMacroCount_i As Integer
     
         
    '//...................................

    


    '//@DIM@COLLECTION@MODULI COLLEZIONI MODULES
    '//...................................
      
         '//DIM IMPORTA MODULI
         Dim mymoduloName_s As String                         '//nome MODULI da importare
         Dim modulesToImport As Variant                      '//calcolo MODULI da importare
         '//DIM della collezione form da importare
         Dim modulesCollection As Collection                   '//collezione delle MODULI prestabilita
         
         '// i contatori delle form
         Dim deletedModules_s As String
         Dim deletetModulesCount_i As Integer
         Dim importedModules_s As String
         Dim importedModulesCount_i As Integer
     
         
    '//...................................

        
        
        
        
        'Le variabili generiche
        Dim sSql As String                                          ' Stringa di estrazione
        Dim sSql_Lett As String                                     ' Stringa di estrazione



'//DIM LE VARIABILI  *** FINE ***
'//-----------------------------------------------------------------------------------//







'//PARAMETRO PATH E FILE DOVE SI TROVANO GLI OGGETTI
'//-----------------------------------------------------------------------------------//

Public Function PathFile_s_pFunct() As String

On Error GoTo CollectionForm_PFunct_Err

        
        '//recupero la path solo se valorizzata altrimenti è null
        If Me.sourceDBPath_s_Txt > "" Or IsNull(Me.sourceDBPath_s_Txt) = False Then
                 
            PathFile_s_pFunct = Me.sourceDBPath_s_Txt
        
        Else
            '//se non esiste la path inserisco null
            PathFile_s_pFunct = ""
        End If
        
        


'//USCITA ED ERRORI
'//--------------------------------------------------------------------------------------//

CollectionForm_PFunct:
    Exit Function

CollectionForm_PFunct_Err:
    MsgBox Error$
    Resume CollectionForm_PFunct

End Function
'//-----------------------------------------------------------------------------------//



'//************************************************************************************************************************//
'//***                              I PARAMETRI DELLE COLLECTION,
'//***                              COLLECTION OBJECT   *** INIZIO ***
'//************************************************************************************************************************//



'//LA COLLEZIONE DELLE FORM DA IMPORTARE  *** INIZIO ***
'//-----------------------------------------------------------------------------------//
'//NOTE: in questa funzione precarico la collezione di oggetti form _
    @PARAMETRO_02= COLLECTION DELLE FORMS DA IMPORTARE
    
'//PRECARICO LA COLLECTION FORM DA IMPORTARE
Public Function CollectionForm_PFunct() As Collection       '//restituisco una collection
    

On Error GoTo CollectionForm_PFunct_Err
        


    Dim frm As Object             '//OGGETTO FORM DELLA COLLEZIONE PRECARICATA


    '//CODICE ----> COLLECTION_FORM_PRECARICATE
    '//CREO LA COLLEZIONE E LA POPOLO
    '//I° COLLEZIONE : OGGETTI FORM   @importa@collection@oggetti@form_(@form la collezione delle form)
    '//------------------------------------------------------------------------------//
    '//NOTE: questa collection, rappresenta gli OGGETTI FORM  CHE APPARTENGONO al progetto _
            MSYS_N01_OGGETTI.mdb e che vengono precaricate in un collezione. Quindi tali _
            oggetti devono esere eliminati o aggiunti a seconda che le form del progetto variano. _
            VANTAGGIO: il vantaggio di utilizzare la collezione facilita l'utilizzo della form di _
            importazione, in quanto le form non sono impostate una tabella esterna che dovrebbe essere _
            spostata insime alla form di importazione. MA ESSENDO PRECARICATE IN UNA COLLECTION si ha _
            sicurezza di spostare tutta la classe FORM con i dati relativi in questo caso agli oggetti FORM. _


    '//@OGGETTI@FORM@PRECARICATI
    '//popola la collection delle FORM di DA IMPORTARE
    ' Aggiungi i nomi delle FORM o ELIMINA O COMMENTA LE FORM DA ESCLUDERE. _
      SE VUOI UTILIZZARE LA COLLEZIONE anche per oggetti form DI PROGETTI DIVERSI inserisci i nome delle form _
      del progetto da importare. N.B. Rircordati di cambiare anche il parametro path+ file per rendere flessibile _
      questa collezione.
        Set objectFORMS = New Collection
        '//Aggiungo gli elementi della collezione
        
            
            '//@carico@collection@form_(carico la @collezione del @progetto scelto)_(@collezione@form)
             '//2024.08.03_Select_Case_Me.Cmb_02_txt = SOSTITUITO IL VALORE CON LA CMB_02 al posto della CMB_01 _
                    Le collection viene scelta in base combinata 2 che individua il progetto _
                    al posto della combinata 1 e la sostituzione riguarda tutte le collection.
     Select Case Me.Cmb_02_txt
        '//@collection@FORM@MSYS@OBJECT
        Case "MSYS_N01_OGGETTI.mdb"
            '//@popolo@collection@form@system_(delle form di sistema, inserisco nella @collezione)
            objectFORMS.Add "MsysDbEstTb01Frm01_Mts01_}-------------------------------------@"
            objectFORMS.Add "MsysDbEstTb01Frm01_Mts01_GE_OBJECT"
            objectFORMS.Add "MsysDbEstTb01Frm01_Stf01_OBJECT"
            objectFORMS.Add "MsysDbEstTb01Frm01_Stf01_OBJECT_TMP"
            objectFORMS.Add "MsysDbEstTb05Frm01_Mt00_}--------------------------------------@"
            objectFORMS.Add "MsysDbEstTb05Frm01_Stf01_DB_ESTERNO"
            objectFORMS.Add "UTILITA_MsysDbEst_Frm12_}--------------------------------------@"
            objectFORMS.Add "UTILITA_MsysDbEst_Frm12_SERVIZI_TUTTI"
        
        '//@collection@FORM@TREEVIEW
        Case "IMPORTA_OGGETTI_TREEVIEW"
            
            objectFORMS.Add "Tree_Frm_M00_00_{@============================================@}"
            objectFORMS.Add "TREE_VIEW_FRM01"
            objectFORMS.Add "TreeView_FIL_}-------------------------------------------@"
            '//2024.08.06_CORREZIONE_VOCE_(corretto errori della collection DA objectFORMS.Add "TreView_FILE" A objectFORMS.Add)
            objectFORMS.Add "TreeView_FILE"
            objectFORMS.Add "TreeView_FrmS01_GE_TREEVIEW"
            objectFORMS.Add "TreeView_FrmS01_TREE"
            objectFORMS.Add "TreeView_FrmS02_StampaDiControllo"
            objectFORMS.Add "TreeView_FrmS03_StampaRisultati"
            objectFORMS.Add "TreeView_Tb01Frm01_}-------------------------------------------@"
            objectFORMS.Add "TreeView_Tb01Frm01_M01_GE_ALBERO"
            objectFORMS.Add "TreeView_Tb01Frm01_M01_RADICE"
            objectFORMS.Add "TreeView_Tb01Frm02_S01_TRONCO"
            objectFORMS.Add "TreeView_Tb01Frm03_S01_RAMI"
            objectFORMS.Add "TreeView_Tb01Frm04_S01_FOGLIE"
            objectFORMS.Add "TreeView_UTILITA_Frm12_}------------------------------@"
            objectFORMS.Add "TreeView_UTILITA_Frm12_SERVIZI_TUTTI"
            
            

        '//@COLLECTION@FORM@MSys = form del progetto msys
        Case "MSYS_N01_OGGETTI.mdb"
            
            objectFORMS.Add "AA_MSys__}-----------------------------------------------@MASTER"
            objectFORMS.Add "AA_MSys_FORM_M01_GESTIONE_OGGETTI_MASTER"
            objectFORMS.Add "MSys_TABELLE_Frm01_S01_TUTTE"
            objectFORMS.Add "MSys_QUERY_Frm01_S01_TUTTE"
            objectFORMS.Add "MSys_FORMS_Frm01_S01_TUTTE"
            objectFORMS.Add "MSys_REPORT_Frm01_S01_TUTTI"
            objectFORMS.Add "MSys_MODULI_Frm01_S01_TUTTI"
            objectFORMS.Add "MSys_MACRO_Frm01_S01_TUTTE"
            objectFORMS.Add "MSys_DF10_DATABASE_Frm01_M0_ESPORTA"
            objectFORMS.Add "MSys_TABELLE_LINK_Frm01_S01_TUTTE"
            objectFORMS.Add "MSys_TUTTE_LE_PROPRIETA_Frm01_S01"
            objectFORMS.Add "Msys_DF11_PROGETTI_Frm00_M01_GESTIONE"
            objectFORMS.Add "Msys_DF11_PROGETTI_Frm01_M01_Obj"
            objectFORMS.Add "MSys_SERVIZI_Frm01_S01_TUTTI"
            objectFORMS.Add "MSys_DF10_DATABASE_Frm01_S01_ELENCO"
            objectFORMS.Add "MSys_TUTTI_PROG_FrmM01_01_PROGETTI"
            objectFORMS.Add "MSys_TUTTI_PROG_VIS_PAGINE_FrmS01_01_Carica"           '//AGGIUNTO
            objectFORMS.Add "Msys_DF11_PROGETTI_Frm02_S01_TIPO_OGGETTI"             '//AGGIUNTO
            objectFORMS.Add "MSys_ELENCO_ObjProgetti_Frm01_S01_Tutti"               '//AGGIUNTO

            
        '@IMPORTA@FORM@EPROPRI
        '//@COLLECTION@FORM@ESPROPRI = form del progetto ESPROPRI
        Case "IMPORTA_OGGETTI_ESPROPRI"
            '//@COLLECTION@@FORM@ESPROPRI  - 2024.08.06_COLLECTION@FORM@ESPROPRI_(inserito la collezione form per gli espropri)
            objectFORMS.Add "LLPP_TB09Frm01_}-----------------------------------------------@"
            objectFORMS.Add "LLPP_TB09Frm01_M01_GE_ESPROPRI"
            objectFORMS.Add "LLPP_TB09Frm01_S01_ESPROPRI"
            
            objectFORMS.Add "UTILITA_EPROPRI_Frm19_}----------------------------------------@"
            objectFORMS.Add "UTILITA_ESPROPRI_Frm20_SERVIZI_TUTTI"
            
        '//COLLECTION@IMPORTA@FORM@DLL@COLECTION   = @IMPORTA@FORM@DLL_(definisco gli oggetti collection FORM per le @liberie@DLL)
        Case "IMPORTA_OGGETTI_MSYS_DLL"
            '//2024.08.05_MODIFICHE_IMPORTA_DLL_(per l'importazione delle librerie DLL del progetto mysys inserito il blocco)
            '//@COLLECTION@DLL@FORM
            objectFORMS.Add ""
            objectFORMS.Add ""
            objectFORMS.Add ""
            objectFORMS.Add ""

        '//COLLECTION@IMPORTA@FORM@ESPROPRI   = @IMPORTA@FORM@DLL_(definisco gli oggetti collection FORM per GLI @ESPROPRI)
        Case "GE_ESPROPRI.mdb"
            '//@COLLECTION@ESPROPRI@FORM
            
                objectFORMS.Add "LLPP_TB09Frm01_}-----------------------------------------------@"
                objectFORMS.Add "LLPP_TB09Frm01_S01_ESPROPRI"
                objectFORMS.Add "UTILITA_EPROPRI_Frm19_}----------------------------------------@"
                objectFORMS.Add "UTILITA_ESPROPRI_Frm20_SERVIZI_TUTTI"
            

        Case Else
            MsgBox "ATTENZIONE! NON CI SONO FORM PER QUESTO PROGETTO: " & Me.Cmb_01_txt, vbCritical, "AVVISO)"
            
    End Select
        
    
        ' Inizializza le collezioni per FORM COLLEGATE
        Set connectedFORMS = New Collection
      
    '// *** FINE ****
    '//I° COLLEZIONE DI FORM DI SISTEMA E LA POPOLO
    '//------------------------------------------------------------------------------//
    
    
    
        '//CONTROLLO DELLE FORM PRECARICATE NELLA COLLEZIONE CARICO PATH + DB E COLLECTION NELLA CASELLA DI TESTO
        '//=====================================================================================================================//
        '//NOTE: queste sono le form da carica nel nuovo progetto.
                        
                        '//RESET CASELLA DI TESTO
                        Me.TXT_OBJECT_ESTERNI = ""
                        Me.TXT_OBJECT_ESTERNI.Requery
        
                                    
                  '//IMPOSTO LA PATH + FILE.MDB + TESTATA + OGGETTI VISUALIZZAZIONE DELLA CASELLA TXT
                  '//--------------------------------------------------------------------------------------------//
                  '//Note: @imposto@path_(imposto la path degli oggetti da recupeare, _
                           preparo la testata per la casela di testo, preparo gli oggetti e visualizzo tutto _
                           nella casella di testo.
                    
                    
                    '//RIEPILOGO COLLEZIONI E VISUALIZZAZIONE PATH ED OGGETTI PRECARICATI
                        Str1 = ""
                        Str1 = Str1 & "CONTROLLO PATH E DATABASE DA CUI RECUPERARE GLI OGGETTI" & vbCrLf
                        Str1 = Str1 & "-----------------------------------------------------------------" & vbCrLf
                        Str1 = Str1 & Me.sourceDBPath_s_Txt & vbCrLf
                        
                        '//INSERISCO LA PATH NELLA CASELLA DI TESTO ED AGGIORNO
                        Me.TXT_OBJECT_ESTERNI = Str1
                        Me.TXT_OBJECT_ESTERNI.Requery

                            '//PREPARO LA TESTATA DELLE COLLECTION PER IL II° INSERIMENTO
                            Str1 = Str1 & vbCrLf & "                          COLLECTION FORM DA IMPORTARE : " & vbCrLf
                            Str1 = Str1 & "//=========================================================================//" & vbCrLf
                            ' Stampa
                            Debug.Print
                            Debug.Print
                            Debug.Print Str1
                            
                            '//iterazione nella collezione per il controllo
                            For Each objfrm In objectFORMS
                                Debug.Print "LA FORM da importare nel progetto: " & objfrm
                                Str1 = Str1 & objfrm & vbCrLf
                            Next objfrm
                            
                            
                  '//--------------------------------------------------------------------------------------------//
                        
        '// *** FINE ***
        '//CONTROLLO DELLE FORM PRECARICATE NELLA COLLEZIONE CARICO PATH + DB E COLLECTION NELLA CASELLA DI TESTO
        '//=====================================================================================================================//
                        
                        
                        ' Restituisci la collezione
                        Set CollectionForm_PFunct = objectFORMS
                            
                    '//CHIAMO LA PROCEDURA UNICA DI STAMPA E CARICAMENTO NELLA CASELLA DI TESTO _
                            parametro_01= la collezione, parametro_02 = il tipo
                      Vv1 = Stampa_MSYS_N01_OGGETTI_funct(CollectionForm_PFunct, "COLLECTION_FORMS")



'//USCITA ED ERRORI
'//--------------------------------------------------------------------------------------//

CollectionForm_PFunct_Exit:
    Exit Function

CollectionForm_PFunct_Err:
    MsgBox Error$
    Resume CollectionForm_PFunct_Exit

End Function



'//LA COLLEZIONE DELLE FORM DA IMPORTARE  *** FINE ***
'//-----------------------------------------------------------------------------------//




'//LA COLLEZIONE DELLE QUERY DA IMPORTARE  *** INIZIO ***
'//-----------------------------------------------------------------------------------//
'//NOTE: in questa funzione precarico la collezione di oggetti QUERY _
    @PARAMETRO_02= COLLECTION DELLE QUERYS DA IMPORTARE
    
'//PRECARICO LA COLLECTION QUERY DA IMPORTARE
Public Function CollectionQuery_PFunct() As Collection       '//restituisco una collection
    

On Error GoTo CollectionQuery_PFunct_Err
        


    Dim frm As Variant             '//OGGETTO QUERY DELLA COLLEZIONE PRECARICATA


    '//CODICE ----> COLLECTION_QUERY_PRECARICATE
    '//CREO LA COLLEZIONE E LA POPOLO
    '//I° COLLEZIONE : OGGETTI QUERY   @importa@collection@oggetti@query_(la collezione delle @query)
    '//------------------------------------------------------------------------------//
    '//NOTE: questa collection, rappresenta gli OGGETTI QUERY  CHE APPARTENGONO al progetto _
            MSYS_N01_OGGETTI.mdb e che vengono precaricate in un collezione. Quindi tali _
            oggetti devono esere eliminati o aggiunti a seconda che le QUERY del progetto variano. _
            VANTAGGIO: il vantaggio di utilizzare la collezione facilita l'utilizzo della QUERY di _
            importazione, in quanto le QUERY non sono impostate una tabella esterna che dovrebbe essere _
            spostata insieme alla QUERY di importazione. MA ESSENDO PRECARICATE IN UNA COLLECTION si ha _
            sicurezza di spostare tutta la classe QUERY con i dati relativi in questo caso agli oggetti QUERY. _


    '//@OGGETTI@QUERY@PRECARICATI
    '//popola la collection delle QUERY di DA IMPORTARE
    ' Aggiungi i nomi delle QUERY o ELIMINA O COMMENTA LE QUERY DA ESCLUDERE. _
      SE VUOI UTILIZZARE LA COLLEZIONE anche per oggetti QUERY DI PROGETTI DIVERSI inserisci i nome delle QUERY _
      del progetto da importare. N.B. Rircordati di cambiare anche il parametro path+ file per rendere flessibile _
      questa collezione.
        Set objectQUERY = New Collection
        '//Aggiugno gli elementi della collezione
        
      
        '//2024.08.03_Select_Case_Me.Cmb_02_txt = SOSTITUITO IL VALORE CON LA CMB_02 al posto della CMB_01 _
                    Le collection viene scelta in base combinata 2 che individua il progetto _
                    al posto della combinata 1 e la sostituzione riguarda tutte le collection.
                    
        Select Case Me.Cmb_02_txt
        
            Case "MSYS_N01_OGGETTI.mdb"
                '//@popolo@collection@QUERY@system_(delle QUERY di sistema, inserisco nella @collezione)_@carico@collection@query
                objectQUERY.Add "MSSys_DF02_Qry01_01_SELECT_TIPO_OGGETTO"
                objectQUERY.Add "MsysDbEstTb01Qry01_Dlt00_}-------------------------------------@"
                objectQUERY.Add "MsysDbEstTb01Qry01_Dlt01_OBJECT"
                objectQUERY.Add "MsysDbEstTb01Qry01_Dlt01_OBJECT_TMP"
                objectQUERY.Add "MsysDbEstTb01Qry01_Ins00_}-------------------------------------@"
                objectQUERY.Add "MsysDbEstTb01Qry01_Ins01_OBJECT_Tmp_To_Archivio"
                objectQUERY.Add "MsysDbEstTb01Qry01_Slt00_}-------------------------------------@"
                objectQUERY.Add "MsysDbEstTb01Qry01_Slt01_OBJECT"
                objectQUERY.Add "MsysDbEstTb01Qry01_Slt01_OBJECT_TMP"
                objectQUERY.Add "MsysDbEstTb05Qry01_Slt00_}-------------------------------------@"
                objectQUERY.Add "MsysDbEstTb05Qry01_Slt01_DB_ESTERNO"
                
              
            
        '//@QUERY@TREEVIEW
        Case "IMPORTA_OGGETTI_TREEVIEW"
            
                objectQUERY.Add "Tree_Qry00_{@=============================================@}_Qry"
                objectQUERY.Add "Tree_Qry01_01_}------------------------------------------@SELECT"
                objectQUERY.Add "Tree_Qry01_CLASSI_STAMPA_DI_CONTROLLO_CANCELLA"
                objectQUERY.Add "Tree_Tb01_Qry01_N01_ELENCO"
                objectQUERY.Add "Tree_Tb01_Qry01_N01_ELENCO_TMP"
                objectQUERY.Add "Tree_Tb01_Qry01_n01_PADRI+figli"
                objectQUERY.Add "Tree_Tb01_Qry01_n01_PADRI+figli_Accoda_Tmp"
                objectQUERY.Add "TreeView_Tb01Qry01_01_RADICE"
                objectQUERY.Add "TreeView_Tb01Qry02_01_TRONCO"
                objectQUERY.Add "TreeView_Tb01Qry03_01_RAMI"
                objectQUERY.Add "TreeView_Tb01Qry04_01_FOGLIE"
        
        
            '//@QUERY@MSys
            Case "MSYS_N01_OGGETTI.mdb"
                
                objectQUERY.Add "MSSys_DF00_Qry00_}---------------------------------------------@"
                objectQUERY.Add "MSSys_DF02_Qry01_01_SELECT_TIPO_OGGETTO"
                objectQUERY.Add "MSsys_DF05_Qry01_ATTRIBUTI"
                objectQUERY.Add "MSys_{@-----------------------OGGETTI_DB----------------------@}"
                objectQUERY.Add "MSys_DF01_COMANDI_Qry_}----------------------------------------@"
                objectQUERY.Add "MSys_DF01_COMANDI_Qry01_01_SELECT_TUTTI"
                objectQUERY.Add "MSys_DF01_COMANDI_Qry01_02_GROUP_TUTTI"
                objectQUERY.Add "MSys_DF10_DATABASE_Qry01_01_SELECT_TUTTI"
                objectQUERY.Add "MSys_DF10_DATABASE_Qry01_01_SELECT_TUTTI_TMP"
                objectQUERY.Add "MSys_DF10_DATABASE_Qry01_81_DELETE_TUTTI_TMP"
                objectQUERY.Add "Msys_DF11_PROGETTI_Qry00_}-------------------------------------@"
                objectQUERY.Add "Msys_DF11_PROGETTI_Qry01_10_GROUP_TUTTI"
                objectQUERY.Add "Msys_DF11_Qry01_01_PROGETTI"
                objectQUERY.Add "Msys_DF12_Qry01_}----------------------------------------------@"
                objectQUERY.Add "Msys_DF12_Qry01_01_TIPO_OGGETTI"
                objectQUERY.Add "MSys_ELENCO_ObjProgetti_Qry00-}--------------------------------@"
                objectQUERY.Add "MSys_ELENCO_ObjProgetti_Qry01_01_SELECT_TUTTI"
                objectQUERY.Add "MSys_ELENCO_ObjProgetti_Qry01_81_DELETE_TUTTI"
                objectQUERY.Add "MSys_FORMS_Qry_}-----------------------------------------------@"
                objectQUERY.Add "MSys_FORMS_Qry01_01_SELECT_TUTTE"
                objectQUERY.Add "MSys_FORMS_Qry01_70_}------------------------------------@INSERT"
                objectQUERY.Add "MSys_FORMS_Qry01_71_INSERT_toMsys_DF12"
                objectQUERY.Add "MSys_FORMS_Qry01_81_DELETE_TUTTE"
                objectQUERY.Add "MSys_MACRO_Qry01_00_}------------------------------------@SELECT"
                objectQUERY.Add "MSys_MACRO_Qry01_01_SELECT_TUTTI"
                objectQUERY.Add "MSys_MACRO_Qry01_70_}-----------------------------------@INSERT"
                objectQUERY.Add "MSys_MACRO_Qry01_71_INSERT_toMsys_DF12"
                objectQUERY.Add "MSys_MACRO_Qry01_80_}------------------------------------@DELETE"
                objectQUERY.Add "MSys_MACRO_Qry01_81_DELETE_TUTTE"
                objectQUERY.Add "MSys_MODULI_Qry01_00_}-----------------------------------@SELECT"
                objectQUERY.Add "MSys_MODULI_Qry01_01_SELECT_TUTTI"
                objectQUERY.Add "MSys_MODULI_Qry01_70_}-----------------------------------@INSERT"
                objectQUERY.Add "MSys_MODULI_Qry01_71_INSERT_toMsys_DF12"
                objectQUERY.Add "MSys_MODULI_Qry01_80_}-----------------------------------@DELETE"
                objectQUERY.Add "MSys_MODULI_Qry01_81_DELETE_TUTTE"
                objectQUERY.Add "MSys_QUERY_Qry_}----------------------------------------------@"
                objectQUERY.Add "MSys_QUERY_Qry01_01_SELECT_TUTTE"
                objectQUERY.Add "MSys_QUERY_Qry01_70_}------------------------------------@INSERT"
                objectQUERY.Add "MSys_QUERY_Qry01_71_INSERT_toMsys_DF12"
                objectQUERY.Add "MSys_QUERY_Qry01_80_}------------------------------------@DELETE"
                objectQUERY.Add "MSys_QUERY_Qry01_81_DELETE_TUTTE"
                objectQUERY.Add "MSys_REPORT_Qry_}----------------------------------------------@"
                objectQUERY.Add "MSys_REPORT_Qry01_01_SELECT_TUTTE"
                objectQUERY.Add "MSys_REPORT_Qry01_70_}-----------------------------------@INSERT"
                objectQUERY.Add "MSys_REPORT_Qry01_71_INSERT_toMsys_DF12"
                objectQUERY.Add "MSys_REPORT_Qry01_80_}-----------------------------------@DELETE"
                objectQUERY.Add "MSys_REPORT_Qry01_81_DELETE_TUTTE"
                objectQUERY.Add "MSys_TABELLE_LINK_Qry_}----------------------------------------@"
                objectQUERY.Add "MSys_TABELLE_LINK_Qry01_01_SELECT_TUTTE"
                objectQUERY.Add "MSys_TABELLE_Qry_}---------------------------------------------@"
                objectQUERY.Add "MSys_TABELLE_Qry01_01_SELECT_Name1_s"
                objectQUERY.Add "MSys_TABELLE_Qry01_01_SELECT_TUTTE"
                objectQUERY.Add "MSys_TABELLE_Qry01_70_}----------------------------------@INSERT"
                objectQUERY.Add "MSys_TABELLE_Qry01_71_INSERT_toMsys_DF12"
                objectQUERY.Add "MSys_TABELLE_Qry01_80_}----------------------------------@DELETE"
                objectQUERY.Add "MSys_TABELLE_Qry01_81_DELETE_TUTTE"
                objectQUERY.Add "MSys_TABELLE_TMP_Qry00_}---------------------------------------@"
                objectQUERY.Add "MSys_TABELLE_TMP_Qry01_01_SELECT_TUTTE"
                objectQUERY.Add "MSys_TABELLE_TMP_Qry01_80_}------------------------------@DELETE"
                objectQUERY.Add "MSys_TABELLE_TMP_Qry01_81_DELETE_TUTTE"
                objectQUERY.Add "MSys_TUTTE_LE_PROPRIETA_Qry01_00_}-----------------------@SELECT"
                objectQUERY.Add "MSys_TUTTE_LE_PROPRIETA_Qry01_01_SELECT"
                objectQUERY.Add "MSys_TUTTI_PROG_}--------------------------------@Qry"
                objectQUERY.Add "MSys_TUTTI_PROG_INSERT_Qry01_01_PROGETTI_in_TMP"
                objectQUERY.Add "MSys_TUTTI_PROG_SELECT_Qry01_01_PROGETTI"
                objectQUERY.Add "MSys_TUTTI_PROG_VIS_PAGINE_SELECT_Qry01_01_Carica"
                objectQUERY.Add "MSys_TUTTI_Qry99_{@=======================================@}_Qry"
    

        
        
            
        '//2024.08.06_COLLECTION@QUERY@ESPROPRI_(inserito la collezione QUERY per gli espropri)
        '//COLLECTION@IMPORTA@QUERY@ESPROPRI   = @IMPORTA@QUERY@ESPROPRI_(definisco gli oggetti collection QUERY per GLI @ESPROPRI)
        Case "IMPORTA_OGGETTI_ESPROPRI"
        '//@COLLECTION@ESPROPRI@QUERY
            
            objectQUERY.Add "LLPP_TB09Qry00_{@===============ESPROPRI======================@}"
            objectQUERY.Add "LLPP_TB09Qry01_00_}--------------------------------------------@"
            objectQUERY.Add "LLPP_TB09Qry01_01_ESPROPRI"
            objectQUERY.Add "LLPP_TB09Qry01_01_ESPROPRI_UPDATE_TuttiCampi"
            objectQUERY.Add "LLPP_TB09Qry01_61_UPDATE_}-----------------------------------@"
            objectQUERY.Add "LLPP_TB09Qry01_62_UPDATE_SINGOLO_RECORD"
            objectQUERY.Add "LLPP_TB09Qry01_62_UPDATE_TUTTI_RECORD"
            
            
         '//@IMPORTA@QUERY@DLL@COLECTION   = @IMPORTA@QUERY@DLL_(definisco gli oggetti collection QUERY per le @liberie@DLL)
        Case "IMPORTA_OGGETTI_MSYS_DLL"
                '//@COLLECTION@DLL@QUERY
                objectQUERY.Add "Msys_DF13Qry01_00_}-----------------------------------------@DLL"
                objectQUERY.Add "Msys_DF13Qry01_01_DLL_PROGETTO"
                objectQUERY.Add "Msys_DF14Qry01_01_DLL_LIBRERIE"
                objectQUERY.Add "Msys_DLL_Qry01_81_}--------------------------------------------@"
                objectQUERY.Add "Msys_DLL_Qry01_81_DELETE_LIBRERIE"
                objectQUERY.Add "Msys_DLLQry01_00_}---------------------------------------------@"
                objectQUERY.Add "Msys_DLLQry01_01_LIBRERIE_ATTIVE"
                objectQUERY.Add "Msys_DLLQry01_10_GROUP_DLL_Correnti"
                objectQUERY.Add "Msys_GRUPPO_DLL_Correnti"


        Case Else
            MsgBox "ATTENZIONE! NON CI SONO QUERY PER QUESTO PROGETTO: " & Me.Cmb_01_txt, vbCritical, "AVVISO)"
            
        End Select
   
    
    
    '// *** FINE ****
    '//I° COLLEZIONE DI QUERY DI SISTEMA E LA POPOLO
    '//------------------------------------------------------------------------------//
    
    
    
    
    
        '//CONTROLLO DELLE QUERY PRECARICATE NELLA COLLEZIONE CARICO PATH + DB E COLLECTION NELLA CASELLA DI TESTO
        '//=====================================================================================================================//
        '//NOTE: queste sono le QUERY da carica nel nuovo progetto.
                        
                        '//RESET CASELLA DI TESTO
                        Me.TXT_OBJECT_ESTERNI = ""
                        Me.TXT_OBJECT_ESTERNI.Requery
        
                                    
                  '//IMPOSTO LA PATH + FILE.MDB + TESTATA + OGGETTI VISUALIZZAZIONE DELLA CASELLA TXT
                  '//--------------------------------------------------------------------------------------------//
                  '//Note: @imposto@path_(imposto la path degli oggetti da recupeare, _
                           preparo la testata per la casela di testo, preparo gli oggetti e visualizzo tutto _
                           nella casella di testo.
                    
                    
                    '//RIEPILOGO COLLEZIONI E VISUALIZZAZIONE PATH ED OGGETTI PRECARICATI
                        Str1 = ""
                        Str1 = Str1 & "CONTROLLO PATH E DATABASE DA CUI RECUPERARE GLI OGGETTI" & vbCrLf
                        Str1 = Str1 & "-----------------------------------------------------------------" & vbCrLf
                        Str1 = Str1 & Me.sourceDBPath_s_Txt & vbCrLf
                        
                        '//INSERISCO LA PATH NELLA CASELLA DI TESTO ED AGGIORNO
                        Me.TXT_OBJECT_ESTERNI = Str1
                        Me.TXT_OBJECT_ESTERNI.Requery

                            '//PREPARO LA TESTATA DELLE COLLECTION PER IL II° INSERIMENTO
                            Str1 = Str1 & vbCrLf & "                          COLLECTION QUERY DA IMPORTARE : " & vbCrLf
                            Str1 = Str1 & "//=========================================================================//" & vbCrLf
                            ' Stampa
                            Debug.Print
                            Debug.Print
                            Debug.Print Str1
                            
                            '//iterazione nella collezione per il controllo
                            For Each objqry In objectQUERY

                                Debug.Print "la MODULI da importare nel progetto: " & objqry
                                Str1 = Str1 & objqry & vbCrLf
                            Next objqry
                            
                            
                            '//II° INSERIMENTO = ACCODO TUTTI GLI OGGETTI DA CARICARE NELLA CASELLA DI STESTO
                            Me.TXT_OBJECT_ESTERNI = vbCrLf & vbCrLf & Str1
                            Me.Recalc
                  
                  '//--------------------------------------------------------------------------------------------//
                        
                        

        '// *** FINE ***
        '//CONTROLLO DELLE QUERY PRECARICATE NELLA COLLEZIONE CARICO PATH + DB E COLLECTION NELLA CASELLA DI TESTO
        '//=====================================================================================================================//
                                            
                        
                        ' Restituisci la collezione
                        Set CollectionQuery_PFunct = objectQUERY
               



        
        '//...............................................................................//
        

'//USCITA ED ERRORI
'//--------------------------------------------------------------------------------------//

CollectionQuery_PFunct_Exit:
    Exit Function

CollectionQuery_PFunct_Err:
    MsgBox Error$
    Resume CollectionQuery_PFunct_Exit

End Function



'//LA COLLEZIONE DELLE TABLE DA IMPORTARE  *** INIZIO ***
'//-----------------------------------------------------------------------------------//


'//PRECARICO LA COLLECTION TABELLE DA IMPORTARE
Public Function CollectionTable_PFunct() As Collection          '//restituisco una collection
    

On Error GoTo CollectionTable_PFunct_Err
        
    Dim tbl As Variant             '//OGGETTO TABELLE DELLA COLLEZIONE PRECARICATA


    '//CODICE ----> COLLECTION_TABELLE_PRECARICATE
    '//CREO LA COLLEZIONE E LA POPOLO
    '//I° COLLEZIONE : OGGETTI TABELLE    @importa@collection@oggetti@table_(la collezione dei @table)
    '//------------------------------------------------------------------------------//
    '//NOTE: questa collection, rappresenta gli OGGETTI TABELLE  CHE APPARTENGONO al progetto _
            MSYS_N01_OGGETTI.mdb e che vengono precaricate in un collezione. Quindi tali _
            oggetti devono esere eliminati o aggiunti a seconda che le TABELLE del progetto variano. _
            VANTAGGIO: il vantaggio di utilizzare la collezione facilita l'utilizzo della TABELLE di _
            importazione, in quanto le TABELLE non sono impostate una tabella esterna che dovrebbe essere _
            spostata insime alla TABELLE di importazione. MA ESSENDO PRECARICATE IN UNA COLLECTION si ha _
            sicurezza di spostare tutta la classe TABELLE con i dati relativi in questo caso agli oggetti TABELLE. _


    '//@OGGETTI@TABELLE@PRECARICATI
    '//popola la collection delle TABELLE di DA IMPORTARE
    ' Aggiungi i nomi delle TABELLE o ELIMINA O COMMENTA LE TABELLE DA ESCLUDERE. _
      SE VUOI UTILIZZARE LA COLLEZIONE anche per oggetti TABELLE DI PROGETTI DIVERSI inserisci i nome delle TABELLE _
      del progetto da importare. N.B. Rircordati di cambiare anche il parametro path+ file per rendere flessibile _
      questa collezione.
        
        Set objectTables = New Collection
        '//Aggiugno gli elementi della collezione
        
        '//SELECT DELLE COLLECTION IN BASE AL VALORE DELLA  COMBINATA_02
        '//------------------------------------------------------------------------------------//
        
            '//2024.08.03_Select_Case_Me.Cmb_02_txt = SOSTITUITO IL VALORE CON LA CMB_02 al posto della CMB_01 _
                    Le collection viene scelta in base combinata 2 che individua il progetto _
                    al posto della combinata 1 e la sostituzione riguarda tutte le collection.
        Select Case Me.Cmb_02_txt
            
            
            
            Case "MSYS_N01_OGGETTI.mdb"
            
                '//@popolo@collection@TABELLE@system_(delle TABELLE di sistema, inserisco nella @collezione)_@CARICO@COLLECTION@TABLE
                objectTables.Add "COLL_MSysTb06_IMPORTA_PROGETTO"
                objectTables.Add "MSys_{@===================================================@}_Tab"
                objectTables.Add "MSys_DF00_}----------------------------------------------------@"
                objectTables.Add "MSys_DF10_DATABASE"
                objectTables.Add "Msys_DF11_PROGETTI"
                objectTables.Add "Msys_DF12_GE_ObjProgetti"
                objectTables.Add "Msys_DF12_TIPO_OGGETTI"
                objectTables.Add "MSys_ELENCO_ObjProgetti"
                objectTables.Add "MSys_F_}-------------------------------------------------------@"
                objectTables.Add "MsysDbEstTb00_}------------------------------------------------@"
                objectTables.Add "MsysDbEstTb01_OBJECT"
                objectTables.Add "MsysDbEstTb01_OBJECT_TMP"
                objectTables.Add "MSysTb05_DB_EST"
                objectTables.Add "MSysTb06_IMPORTA_PROGETTO"
                                    
        
            '//@TABELLE@TREEVIEW
            Case "IMPORTA_OGGETTI_TREEVIEW"
                
                objectTables.Add "Tre99_{@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@}_Tab"
                objectTables.Add "Tree"

                objectTables.Add "Tree_CLASSI_STAMPA_DI_CONTROLLO"
                objectTables.Add "Tree_DF01_COMANDI"
                objectTables.Add "TreeTemporanea"
                objectTables.Add "TreTab_}-------------------------------------------------------@"
                
            '//@TABELLE@MSys
            Case "MSYS_N01_OGGETTI.mdb"
                
                objectTables.Add "MSys_{@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@}_Tab"
                objectTables.Add "MSys_{@===================================================@}_Tab"
                objectTables.Add "MSys_DF00_}----------------------------------------------------@"
                objectTables.Add "MSys_DF01_COMANDI"
                objectTables.Add "MSys_DF01_COMANDI_TMP"
                objectTables.Add "MSys_DF02_TIPO_OGGETTO"
                objectTables.Add "MSys_DF05_ATTRIBUTI_TABELLA"
                objectTables.Add "MSys_DF10_DATABASE"
                objectTables.Add "MSys_DF10_DATABASE_TMP"
                objectTables.Add "Msys_DF11_PROGETTI"
                objectTables.Add "Msys_DF12_GE_ObjProgetti"
                objectTables.Add "Msys_DF12_TIPO_OGGETTI"
                objectTables.Add "MSys_ELENCO_}-------------------------------------------------@"
                objectTables.Add "MSys_ELENCO_ObjProgetti"
                objectTables.Add "MSys_ELENCO_ObjProgetti_TMP"
                objectTables.Add "MSys_F_}-------------------------------------------------------@"
                objectTables.Add "MSys_FORMS"
                objectTables.Add "MSys_MACRO"
                objectTables.Add "MSys_MODULI"
                objectTables.Add "MSys_PAGINE"
                objectTables.Add "MSys_QUERY"
                objectTables.Add "MSys_REPORT"
                objectTables.Add "MSys_TABELLE"
                objectTables.Add "MSys_TABELLE_LINK"
                objectTables.Add "MSys_TABELLE_Tmp"
                objectTables.Add "MSys_TUTTE_}---------------------------------------------------@"
                objectTables.Add "MSys_TUTTE_BARRE_DEI_COMANDI"
                objectTables.Add "MSys_TUTTE_LE_PAGINE"
                objectTables.Add "MSys_TUTTE_LE_PROPRIETA"
                objectTables.Add "MSys_TUTTI_PRO_{@=========================================@}_Tab"
                objectTables.Add "MSys_TUTTI_PRO_}-----------------------------------------------@"
                objectTables.Add "MSys_TUTTI_PROG"
                objectTables.Add "MSys_TUTTI_PROG_TMP"
                objectTables.Add "MSys_TUTTI_PROG_VIS_PAGINE"

            
            
            
        '//@TABELLE@ESPROPRI = LE TABELLE del progetto ESPROPRI
        Case "GE_ESPROPRI.mdb"
     
            '//@COLLECTION@ESPROPRI@TABELLE
            objectTables.Add "LLPP_ATTI_Tb01_{@=============================================@}"
            objectTables.Add "LLPP_ATTI_Tb01_Gestione"
            objectTables.Add "LLPP_ATTI_Tb01_Gestione_TMP"
            objectTables.Add "LLPP_ATTI_Tb02_{@====================================@}_ALLEGATI"
            objectTables.Add "LLPP_ATTI_Tb02_Allegati"
            objectTables.Add "LLPP_DF01_{===============================================}_DF01"
            objectTables.Add "LLPP_DF01_CodiceOpera"
            objectTables.Add "LLPP_DF02_AttivitaOperaPubblica"
            objectTables.Add "LLPP_DF12_{@========================================@}_TIPO_ATTO"
            objectTables.Add "LLPP_DF12_TIPO_ATTO"
            objectTables.Add "LLPP_TB09_ESPR_}-----------------------------------------------@"
            objectTables.Add "LLPP_TB09_ESPROPRI"
            objectTables.Add "LLPP_TB09_ESPROPRI_TMP"

        
             '//@IMPORTA@TABELLE@DLL@COLECTION   = @IMPORTA@TABELLE@DLL_(definisco gli oggetti collection tabella per le @liberie@DLL)
            Case "IMPORTA_OGGETTI_MSYS_DLL"
                
                '//@COLLECTION@DLL@TABELLE
                objectTables.Add "MSys_DF13_}----------------------------------------------------@"
                objectTables.Add "Msys_DF13_DLL_PROGETTO"
                objectTables.Add "Msys_DF14_DLL_LIBRERIE"
                objectTables.Add "Msys_DLL_LIBRERIE"
        
        
            Case Else
            
                MsgBox "ATTENZIONE! NON CI SONO TABELLE PER QUESTO PROGETTO: " & Me.Cmb_01_txt, vbCritical, "AVVISO)"
            
        End Select
      
      '//SELECT DELLE COLLECTION IN BASE AL VALORE DELLA  COMBINATA_02  *** fine ***
      '//------------------------------------------------------------------------------------//
      
            
                    
    
        ' Inizializza le collezioni per TABELLE COLLEGATE
        Set connectedTABELLES = New Collection
        
              '//ATTIVO LA COLLEZIONE
              '//------------------------------------------------------------------------//
              '//Note: CHIAMO LA FUNZIONE COLLECTION TABLE per ottenere la collezione di TABELLE
                
                '//Set tablesCollection = CollectionTable_PFunct()
                
              '//------------------------------------------------------------------------//
        
      
    '// *** FINE ****
    '//I° COLLEZIONE DI TABELLE DI SISTEMA E LA POPOLO
    '//------------------------------------------------------------------------------//
    

        '//CONTROLLO DELLE_TABELLE_PRECARICATE  NELLA COLLEZIONE CARICO PATH + DB E COLLECTION NELLA CASELLA DI TESTO
        '//=====================================================================================================================//
        '//NOTE: queste sono DELLE_TABELLE_PRECARICATE  carica nel nuovo progetto.
                        
                        '//RESET CASELLA DI TESTO
                        Me.TXT_OBJECT_ESTERNI = ""
                        Me.TXT_OBJECT_ESTERNI.Requery
        
                                    
                  '//IMPOSTO LA PATH + FILE.MDB + TESTATA + OGGETTI VISUALIZZAZIONE DELLA CASELLA TXT
                  '//--------------------------------------------------------------------------------------------//
                  '//Note: @imposto@path_(imposto la path degli oggetti da recupeare, _
                           preparo la testata per la casela di testo, preparo gli oggetti e visualizzo tutto _
                           nella casella di testo.
                    
                    
                    '//RIEPILOGO COLLEZIONI E VISUALIZZAZIONE PATH ED OGGETTI PRECARICATI
                        Str1 = ""
                        Str1 = Str1 & "CONTROLLO PATH E DATABASE DA CUI RECUPERARE GLI OGGETTI" & vbCrLf
                        Str1 = Str1 & "-----------------------------------------------------------------" & vbCrLf
                        Str1 = Str1 & Me.sourceDBPath_s_Txt & vbCrLf
                        
                        '//INSERISCO LA PATH NELLA CASELLA DI TESTO ED AGGIORNO
                        Me.TXT_OBJECT_ESTERNI = Str1
                        Me.TXT_OBJECT_ESTERNI.Requery

                            '//PREPARO LA TESTATA DELLE COLLECTION PER IL II° INSERIMENTO
                            Str1 = Str1 & vbCrLf & "                COLLECTION DELLE_TABELLE_PRECARICATE DA IMPORTARE : " & vbCrLf
                            Str1 = Str1 & "//=========================================================================//" & vbCrLf
                            ' Stampa
                            Debug.Print
                            Debug.Print
                            Debug.Print Str1
                            
                            '//iterazione nella collezione per il controllo
                            For Each objtbl In objectTables

                                Debug.Print "DELLE_TABELLE_PRECARICATE  importare nel progetto: " & objtbl
                                Str1 = Str1 & objtbl & vbCrLf
                                
                            Next objtbl
                            
                            
                            '//II° INSERIMENTO = ACCODO TUTTI GLI OGGETTI DA CARICARE NELLA CASELLA DI TESTO
                            Me.TXT_OBJECT_ESTERNI = vbCrLf & vbCrLf & Str1
                            Me.Recalc
                  '//--------------------------------------------------------------------------------------------//
                        
        '// *** FINE ***
        '//CONTROLLO DELLE_TABELLE_PRECARICATE  NELLA COLLEZIONE CARICO PATH + DB E COLLECTION NELLA CASELLA DI TESTO
        '//=====================================================================================================================//
                                             
                                                     
                            
                            
                     ' Restituisci la collezione
                        Set CollectionTable_PFunct = objectTables
                        



        
        

'//USCITA ED ERRORI
'//--------------------------------------------------------------------------------------//

CollectionTable_PFunct_Exit:
    Exit Function

CollectionTable_PFunct_Err:
    MsgBox Error$
    Resume CollectionTable_PFunct_Exit

End Function


'//LA COLLEZIONE DELLE TABLE DA IMPORTARE  *** FINE ***
'//-----------------------------------------------------------------------------------//






'//LA COLLEZIONE DELLE MACRO DA IMPORTARE  *** INIZIO ***
'//-----------------------------------------------------------------------------------//

'//@carico@collection@macro_(la collezione di oggetti macro viene @precaricata)
'//PRECARICO LA COLLECTION MACRO DA IMPORTARE
Public Function CollectionMacro_PFunct() As Collection          '//restituisco una collection
    

On Error GoTo CollectionMacro_PFunct_Err
        
    Dim frm As Variant             '//OGGETTO MACRO DELLA COLLEZIONE PRECARICATA


    '//CODICE ----> COLLECTION_MACRO_PRECARICATE
    '//CREO LA COLLEZIONE E LA POPOLO
    '//I° COLLEZIONE : OGGETTI MACRO  @importa@collection@oggetti@macro_(la collezione delle @macro)
    '//------------------------------------------------------------------------------//
    '//NOTE: questa collection, rappresenta gli OGGETTI MACRO  CHE APPARTENGONO al progetto _
            MSYS_N01_OGGETTI.mdb e che vengono precaricate in un collezione. Quindi tali _
            oggetti devono esere eliminati o aggiunti a seconda che le MACRO del progetto variano. _
            VANTAGGIO: il vantaggio di utilizzare la collezione facilita l'utilizzo della MACRO di _
            importazione, in quanto le MACRO non sono impostate una tabella esterna che dovrebbe essere _
            spostata insime alla MACRO di importazione. MA ESSENDO PRECARICATE IN UNA COLLECTION si ha _
            sicurezza di spostare tutta la classe MACRO con i dati relativi in questo caso agli oggetti MACRO. _


    '//@OGGETTI@MACRO@PRECARICATI
    '//popola la collection dei MACRO di DA IMPORTARE
    ' Aggiungi i nomi dei MACRO o ELIMINA O COMMENTA I MACRO DA ESCLUDERE. _
      SE VUOI UTILIZZARE LA COLLEZIONE anche per oggetti MACRO DI PROGETTI DIVERSI inserisci i nome dei MACRO _
      del progetto da importare. N.B. Rircordati di cambiare anche il parametro path+ file per rendere flessibile _
      questa collezione.
        
        Set objectMACRO = New Collection
        '//Aggiugno gli elementi della collezione
        
    
        
        '//2024.08.03_Select_Case_Me.Cmb_02_txt = SOSTITUITO IL VALORE CON LA CMB_02 al posto della CMB_01 _
                    Le collection viene scelta in base combinata 2 che individua il progetto _
                    al posto della combinata 1 e la sostituzione riguarda tutte le collection.
        Select Case Me.Cmb_02_txt
            Case "MSYS_N01_OGGETTI.mdb"
                '//@popolo@collection@MACRO@system_(dei MACRO di sistema, inserisco nella @collezione)
                objectMACRO.Add "MsysDbEstTb00Mcr00_}--------------------------------------@Macro"
                objectMACRO.Add "MsysDbEstTb00Mcr01_N01_ESPORTA_XLS"
                objectMACRO.Add "MsysDbEstTb00Mcr01_N02_COLLEGA_XLS"

        
        '//@MACRO@TREEVIEW
        Case "IMPORTA_OGGETTI_TREEVIEW"
            
            objectMACRO.Add "CERCA_Mcr_{@==================================================@}}"
            objectMACRO.Add "CERCA_Mcr_FORM"
            objectMACRO.Add "CERCA_Mcr_MODULO"
            objectMACRO.Add "CERCA_Mcr_QUERY"
            objectMACRO.Add "CERCA_Mcr_TABELLA"
            objectMACRO.Add "MACRO_{@======================================================@}"
            objectMACRO.Add "MSys_Mcr00_}---------------------------------------------------@"
            objectMACRO.Add "MSys_Mcr01_ATTIVA_MACRO_ESPORTA_OGGETTI_MSSys"
            objectMACRO.Add "TREEVIEW_Mcr00_{@============================================@}"
            objectMACRO.Add "TREEVIEW_Mcr01_ATTIVA_LIBRERIE"
            objectMACRO.Add "TREEVIEW_Mcr01_STAMPA_LIBRERIE_ATTIVE"
            
        
        '//@MACRO@MSys
        Case "MSYS_N01_OGGETTI.mdb"
            '//@COLLECTION@ESPROPRI@MACRO
            objectMACRO.Add "CERCA_Mcr_{@==================================================@}"
            objectMACRO.Add "CERCA_Mcr_FORM"
            objectMACRO.Add "CERCA_Mcr_MODULO"
            objectMACRO.Add "CERCA_Mcr_QUERY"
            objectMACRO.Add "CERCA_Mcr_TABELLA"
            objectMACRO.Add "MSys_Mcr00_{@=================================================@}"
            objectMACRO.Add "MSys_Mcr00_}---------------------------------------------------@"
            objectMACRO.Add "MSys_Mcr01_ATTIVA_MACRO_ESPORTA_OGGETTI_MSSys"
            
        
          
        '//@MACRO@ESPROPRI = LE MACRO del progetto ESPROPRI
        Case "GE_ESPROPRI.mdb"
            

            objectMACRO.Add "COLLEGA_Mcr_}--------------------------------------------------@"
            objectMACRO.Add "COLLEGA_Mcr_GEST_MENU_TB01_VOCI_Espropri"
            objectMACRO.Add "COLLEGA_Mcr_LLPP_ATTI_Tb01_Gestione_Espropri"
            objectMACRO.Add "COLLEGA_Mcr_LLPP_ATTI_Tb01_Gestione_TMP_Espropri"
            objectMACRO.Add "COLLEGA_Mcr_LLPP_ATTI_Tb02_Allegati_Espropri"
            objectMACRO.Add "COLLEGA_Mcr_LLPP_DF01_CodiceOpera_Espropri"
            objectMACRO.Add "COLLEGA_Mcr_LLPP_DF02_AttivitaOperaPubblica_Espropri"
            objectMACRO.Add "COLLEGA_Mcr_LLPP_DF12_TIPO_ATTO_Espropri"
            objectMACRO.Add "COLLEGA_Mcr_LLPP_TUTTI_GLI_ESPROPRI"
            objectMACRO.Add "COLLEGA_TABELLA_Mcr_ATTI_Tb01_Gestione_TUTTO_Espropri"
            objectMACRO.Add "ESPR_Mcr00_}---------------------------------------------------@"
            objectMACRO.Add "ESPR_Mcr01_CollegaTUTTO_(Atti+Allegati+CodOpera+TipoAtto+OOPP)"
        
            
         '//@IMPORTA@MACRO@DLL@COLECTION   = @IMPORTA@MACRO@DLL_(definisco gli oggetti collection MACRO per le @liberie@DLL)
        Case "IMPORTA_OGGETTI_MSYS_DLL"
                '//@COLLECTION@DLL@MACRO
                objectMACRO.Add "UTILITA_Msys_Mdl03_DLL_REFERENZIA_LE_LIBRERIE"
                objectMACRO.Add "UTILITA_Msys_Mcr01_STAMPA_LIBRERIE_ATTIVE_DLL"
                '//2024.08.05_@macro@attiva@librerie@dll_(inserita la macro da importare per attivare le librerie FORM IMPORTA OGGETTI)
                objectMACRO.Add "UTILITA_Msys_Frm21_IMPORTA_OGGETTI_ATTIVA_LIBRERIE_DLL"
                
        
        '//2024.08.06_COLLECTION@FORM@ESPROPRI_(inserito la collezione MACRO per gli espropri)
        '//COLLECTION@IMPORTA@MACRO@ESPROPRI   = @IMPORTA@MACRO@DLL_(definisco gli oggetti collection MACRO per GLI @ESPROPRI)
        Case "IMPORTA_OGGETTI_ESPROPRI"
            '//@COLLECTION@ESPROPRI@MACRO
            
                                   
                    objectMACRO.Add "COLLEGA_Mcr_}--------------------------------------------------@"
                    objectMACRO.Add "COLLEGA_Mcr_GEST_MENU_TB01_VOCI_Espropri"
                    objectMACRO.Add "COLLEGA_Mcr_LLPP_ATTI_Tb01_Gestione_Espropri"
                    objectMACRO.Add "COLLEGA_Mcr_LLPP_ATTI_Tb01_Gestione_TMP_Espropri"
                    objectMACRO.Add "COLLEGA_Mcr_LLPP_DF01_CodiceOpera_Espropri"
                    objectMACRO.Add "COLLEGA_Mcr_LLPP_DF02_AttivitaOperaPubblica_Espropri"
                    objectMACRO.Add "COLLEGA_Mcr_LLPP_DF12_TIPO_ATTO_Espropri"
                    objectMACRO.Add "COLLEGA_Mcr_LLPP_TUTTI_GLI_ESPROPRI"
                    objectMACRO.Add "COLLEGA_TABELLA_Mcr_ATTI_Tb01_Gestione_TUTTO_Espropri"
                    objectMACRO.Add "ESPR_Mcr00_}---------------------------------------------------@"
                    objectMACRO.Add "ESPR_Mcr01_CollegaTUTTO_(Atti+Allegati+CodOpera+TipoAtto+OOPP)"
                    objectMACRO.Add "LLPP_TB09Mcr01_}-----------------------------------------------@"
                    objectMACRO.Add "LLPP_TB09Mcr01_01_ESPROPRI_COLLEGA"
        Case ""

        Case Else
            MsgBox "ATTENZIONE! NON CI SONO MACRO PER QUESTO PROGETTO: " & Me.Cmb_01_txt, vbCritical, "AVVISO)"
            
        End Select
   


    
        ' Inizializza le collezioni per MACRO COLLEGATE
        Set connectedMACRO = New Collection
        
              '//ATTIVO LA COLLEZIONE
              '//------------------------------------------------------------------------//
              '//Note: CHIAMO LA FUNZIONE COLLECTION MACRO per ottenere la collezione di MACRO
                
                '//Set macroCollection = CollectionMacro_PFunct()
                
              '//------------------------------------------------------------------------//
        
      
    '// *** FINE ****
    '//I° COLLEZIONE DI MACRO DI SISTEMA E LA POPOLO
    '//------------------------------------------------------------------------------//
    
        
        '//CONTROLLO DELLE MACRO PRECARICATE NELLA COLLEZIONE CARICO PATH + DB E COLLECTION NELLA CASELLA DI TESTO
        '//=====================================================================================================================//
        '//NOTE: queste sono le MACRO da carica nel nuovo progetto.
                        
                        '//RESET CASELLA DI TESTO
                        Me.TXT_OBJECT_ESTERNI = ""
                        Me.TXT_OBJECT_ESTERNI.Requery
        
                                    
                  '//IMPOSTO LA PATH + FILE.MDB + TESTATA + OGGETTI VISUALIZZAZIONE DELLA CASELLA TXT
                  '//--------------------------------------------------------------------------------------------//
                  '//Note: @imposto@path_(imposto la path degli oggetti da recupeare, _
                           preparo la testata per la casela di testo, preparo gli oggetti e visualizzo tutto _
                           nella casella ditesto.
                    
                    
                    '//RIEPILOGO COLLEZIONI E VISUALIZZAZIONE PATH ED OGGETTI PRECARICATI
                        Str1 = ""
                        Str1 = Str1 & "CONTROLLO PATH E DATABASE DA CUI RECUPERARE GLI OGGETTI" & vbCrLf
                        Str1 = Str1 & "-----------------------------------------------------------------" & vbCrLf
                        Str1 = Str1 & Me.sourceDBPath_s_Txt & vbCrLf
                                                                        
                        '//INSERISCO LA PATH NELLA CASELLA DI TESTO ED AGGIORNO
                        Me.TXT_OBJECT_ESTERNI = Str1
                        Me.TXT_OBJECT_ESTERNI.Requery

                            '//PREPARO LA TESTATA DELLE COLLECTION PER IL II° INSERIMENTO
                            Str1 = Str1 & vbCrLf & "                          COLLECTION MACRO DA IMPORTARE : " & vbCrLf
                            Str1 = Str1 & "//=========================================================================//" & vbCrLf
                                                       
                            ' Stampa le MACRO fisiche
                            Debug.Print
                            Debug.Print
                            Debug.Print "MACRO DA IMPORTARE:"
                            '//iterazione nella collezione per il controllo
                            For Each mcr In objectMACRO
                                Debug.Print "la MACRO da importare nel progetto: " & mcr
                                Str1 = Str1 & mcr & vbCrLf
                            Next mcr
         
                  '//--------------------------------------------------------------------------------------------//
                        
        '// *** FINE ***
        '//CONTROLLO DELLE MACRO PRECARICATE NELLA COLLEZIONE CARICO PATH + DB E COLLECTION NELLA CASELLA DI TESTO
        '//=====================================================================================================================//
                            
                    
                     ' Restituisci la collezione
                        Set CollectionMacro_PFunct = objectMACRO
               

                    '//CHIAMO LA PROCEDURA UNICA DI STAMPA E CARICAMENTO NELLA CASELLA DI TESTO _
                            parametro_01= la collezione, parametro_02 = il tipo
                      Vv1 = Stampa_MSYS_N01_OGGETTI_funct(CollectionMacro_PFunct, "COLLECTION_MACROS")



        
        '//...............................................................................//
        

'//USCITA ED ERRORI
'//--------------------------------------------------------------------------------------//

CollectionMacro_PFunct_Exit:
    Exit Function

CollectionMacro_PFunct_Err:
    MsgBox Error$
    Resume CollectionMacro_PFunct_Exit

End Function


'//LA COLLEZIONE DELLE MACRO DA IMPORTARE  *** FINE ***
'//-----------------------------------------------------------------------------------//










'//LA COLLEZIONE DELLE MODULI DA IMPORTARE  *** INIZIO ***
'//-----------------------------------------------------------------------------------//


'//PRECARICO LA COLLECTION MODULI DA IMPORTARE
Public Function CollectionModuli_PFunct() As Collection          '//restituisco una collection
    

On Error GoTo CollectionModuli_PFunct_Err
        
     Dim moduleName As Variant                       '//OGGETTO MODULI DELLA COLLEZIONE PRECARICATA


    '//CODICE ----> COLLECTION_MODULI_PRECARICATE
    '//CREO LA COLLEZIONE E LA POPOLO
    '//I° COLLEZIONE : OGGETTI MODULI   @importa@collection@oggetti@moduli_(la collezione dei @moduli)
    '//------------------------------------------------------------------------------//
    '//NOTE: questa collection, rappresenta gli OGGETTI MODULI  CHE APPARTENGONO al progetto _
            MSYS_N01_OGGETTI.mdb e che vengono precaricate in un collezione. Quindi tali _
            oggetti devono esere eliminati o aggiunti a seconda che le MODULI del progetto variano. _
            VANTAGGIO: il vantaggio di utilizzare la collezione facilita l'utilizzo della MODULI di _
            importazione, in quanto le MODULI non sono impostate una tabella esterna che dovrebbe essere _
            spostata insime alla MODULI di importazione. MA ESSENDO PRECARICATE IN UNA COLLECTION si ha _
            sicurezza di spostare tutta la classe MODULI con i dati relativi in questo caso agli oggetti MODULI. _


    '//@OGGETTI@MODULI@PRECARICATI
    '//popola la collection dei MODULI di DA IMPORTARE
    ' Aggiungi i nomi dei MODULI o ELIMINA O COMMENTA I MODULI DA ESCLUDERE. _
      SE VUOI UTILIZZARE LA COLLEZIONE anche per oggetti MODULI DI PROGETTI DIVERSI inserisci i nome dei MODULI _
      del progetto da importare. N.B. Rircordati di cambiare anche il parametro path+ file per rendere flessibile _
      questa collezione.
        
        Set objectModules = New Collection
        '//Aggiugno gli elementi della collezione
        
     
        
        '//2024.08.03_Select_Case_Me.Cmb_02_txt = SOSTITUITO IL VALORE CON LA CMB_02 al posto della CMB_01 _
                    Le collection viene scelta in base combinata 2 che individua il progetto _
                    al posto della combinata 1 e la sostituzione riguarda tutte le collection.
        Select Case Me.Cmb_02_txt
                
            '//@popolo@collection@MODULI@system_(dei MODULI di sistema, inserisco nella @collezione; )_@carico@collection@macro
            '//@MODULI@MSys
            Case "MSYS_N01_OGGETTI.mdb"
                '//@COLLECTION@ESPROPRI@MODULI
                objectModules.Add "UTIL_MDL40_{@=================================================@}}"
                objectModules.Add "UTIL_MDL40_N01_CONTROLLO_OGGETTI_DB_SALVATAGGIO_TUTTI"
                objectModules.Add "UTIL_MDL40_N08_RICERCA_TUTTI_GLI_OGGETTI_NEL_DB"
                objectModules.Add "UTIL_MDL40_N09_CONTROLLO_ESISTENZA_OGGETTO_DB"
                objectModules.Add "UTIL_MDL40_N10_CONTROLLO_OGGETTI_APRI_ACCESS"
                objectModules.Add "UTIL_MDL40_N11_MACRO_ESPORTA_OGGETTI_MSSys"
                objectModules.Add "UTIL_MDL90_01_(VARIABILI_GLOBAL_FORM)"
                objectModules.Add "UTIL_Nrx1001_{@=============CLASSI_OGGETTI====================@}"
                objectModules.Add "CLASSE_GESTIONE_OGGETTI_}--------------------------------------@"
                objectModules.Add "CLASSE_GESTIONE_OGGETTI_DATABASE"
                objectModules.Add "CLASSE_GESTIONE_OGGETTI_FORM"
                objectModules.Add "UTIL_Nrx1003_n01_RECUPERO_PARAMETRI_TABELLA_OGGETTI"
                objectModules.Add "COLLEGA_Mdl_PRES3000_TB25_DIPENDENTI"
                objectModules.Add "CLASSE_{@=====================================================@}"
                objectModules.Add "CLASSE_ATTI_Tb01_Gestione"
       
            
        '//@MODULI@TREEVIEW
        Case "IMPORTA_OGGETTI_TREEVIEW"
            
                objectModules.Add "ACCESS_{@=================================================@}_Mdl"
                objectModules.Add "ACCESS_Classe_LIBRERIE_RIFERIMENTI"
                objectModules.Add "ACCESS_Mdl_REFERENZIA_LE_LIBRERIE"
                objectModules.Add "API_WINDOWS_Main"
                objectModules.Add "API_WINDOWS_Test_Excel"
                objectModules.Add "APRI_WINDOWS_File"
                objectModules.Add "BARRA_COMANDI_{@==================MDL=========================@}"
                objectModules.Add "BARRA_COMANDI_MyContest"
                objectModules.Add "BARRA_COMANDI_MyTEST"
                objectModules.Add "BARRA_DEI_COMANDI_ClasseEventi"
                objectModules.Add "BARRA_DEI_COMANDI_CopiaDueBarre"
                objectModules.Add "BARRA_DEI_MENU+BARRA_DEI_COMANDI_MSys_OGGETTI"
                objectModules.Add "CLASSE_{@=====================================================@}"
                objectModules.Add "CLASSE_ALBERO"
                objectModules.Add "CLASSE_ALBERO_Test"
                objectModules.Add "CLASSE_ATTI_Tb01_Gestione"
                objectModules.Add "CLASSE_ClsStack_TEST"
                objectModules.Add "CLASSE_GESTIONE_OGGETTI_}--------------------------------------@"
                objectModules.Add "CLASSE_GESTIONE_OGGETTI_DATABASE"
                objectModules.Add "CLASSE_GESTIONE_OGGETTI_FORM"
                objectModules.Add "CLASSE_NOD_}---------------------------------------------------@"
                objectModules.Add "CLASSE_NODO"
                objectModules.Add "CLASSE_NODO_Test"
                objectModules.Add "CLASSE_STACK"
                objectModules.Add "CLASSE_TREE"
                objectModules.Add "clsFile"
                objectModules.Add "ComboBoxHandler"
                objectModules.Add "CreaAlberoDaTabel_}--------------------------------------------@"
                objectModules.Add "CreaAlberoDaTabella"
                objectModules.Add "CreaAlberoDaTabella_TEST"
                objectModules.Add "ESPORTA_Mdl_}--------------------------------------------------@"
                objectModules.Add "GESTIONE_Mdl10_FILE_APRI_PDF"
                objectModules.Add "LLPP_ATTI_Tb01_Gestione_Classe"
                objectModules.Add "LLPP_CLASSE_MENU_LAVORI_PUBBLICI"
                objectModules.Add "PROGETTI_GESTIONE_TUTTI_Classe"
                objectModules.Add "UTIL_{=====================================================}_Mdl"
                objectModules.Add "UTIL_MDL30_CANCELLA_I_MODULI_DbEsterno"
                objectModules.Add "UTIL_MDL40_{@=================================================@}"
                objectModules.Add "UTIL_MDL40_N01_CONTROLLO_OGGETTI_DB_SALVATAGGIO_TUTTI"
                objectModules.Add "UTIL_MDL40_N08_RICERCA_TUTTI_GLI_OGGETTI_NEL_DB"
                objectModules.Add "UTIL_MDL40_N09_CONTROLLO_ESISTENZA_OGGETTO_DB"
                objectModules.Add "UTIL_MDL40_N10_CONTROLLO_OGGETTI_APRI_ACCESS"
                objectModules.Add "UTIL_MDL40_N11_MACRO_ESPORTA_OGGETTI_MSSys"
                objectModules.Add "UTIL_MDL90_01_(VARIABILI_GLOBAL_FORM)"
                objectModules.Add "UTIL_Nrx1001_{@=============CLASSI_OGGETTI====================@}"
                objectModules.Add "UTIL_Nrx1003_n01_RECUPERO_PARAMETRI_TABELLA_OGGETTI"
                 
                
                
            
 
            '//2024.08.06_COLLECTION@MODULI@ESPROPRI_(inserito la collezione MODULI per gli espropri)
            '//COLLECTION@IMPORTA@MODULI@ESPROPRI   = @IMPORTA@MODULI@DLL_(definisco gli oggetti collection MODULI per GLI @ESPROPRI)
            Case "IMPORTA_OGGETTI_ESPROPRI"
            '//@COLLECTION@ESPROPRI@MODULI
            
            objectModules.Add "ESPORTA_FILE_Mdl00=============================================="
            objectModules.Add "FILE_Mdl01_TROVA_PATH_E_FILE"
            


        
        '//@IMPORTA@MODULI@DLL@COLECTION   = @IMPORTA@MODULI@DLL_(definisco gli oggetti collection MODULI per le @liberie@DLL)
        Case "IMPORTA_OGGETTI_MSYS_DLL"
                '//@COLLECTION@DLL@MODULI
                objectModules.Add "UTILITA_Msys_Mdl03_DLL_REFERENZIA_LE_LIBRERIE"
                
        
        Case Else
            MsgBox "ATTENZIONE! NON CI SONO MODULI PER QUESTO PROGETTO: " & Me.Cmb_01_txt, vbCritical, "AVVISO)"
            
        End Select
   
    
        ' Inizializza le collezioni per MODULI COLLEGATE
        Set connectedMODULIS = New Collection
        
              '//ATTIVO LA COLLEZIONE
              '//------------------------------------------------------------------------//
              '//Note: CHIAMO LA FUNZIONE COLLECTION MODULI per ottenere la collezione di MODULI
                
                '//Set tablesCollection = CollectionModuli_PFunct()
                
              '//------------------------------------------------------------------------//
        
      
    '// *** FINE ****
    '//I° COLLEZIONE DI MODULI DI SISTEMA E LA POPOLO
    '//------------------------------------------------------------------------------//
    
    
        '//CONTROLLO DELLE MODULI PRECARICATE NELLA COLLEZIONE CARICO PATH + DB E COLLECTION NELLA CASELLA DI TESTO
        '//=====================================================================================================================//
        '//NOTE: queste sono le MODULI da carica nel nuovo progetto.
                        
                        '//RESET CASELLA DI TESTO
                        Me.TXT_OBJECT_ESTERNI = ""
                        Me.TXT_OBJECT_ESTERNI.Requery
        
                                    
                  '//IMPOSTO LA PATH + FILE.MDB + TESTATA + OGGETTI VISUALIZZAZIONE DELLA CASELLA TXT
                  '//--------------------------------------------------------------------------------------------//
                  '//Note: @imposto@path_(imposto la path degli oggetti da recupeare, _
                           preparo la testata per la casela di testo, preparo gli oggetti e visualizzo tutto _
                           nella casella di testo.
                    
                    
                    '//RIEPILOGO COLLEZIONI E VISUALIZZAZIONE PATH ED OGGETTI PRECARICATI
                        Str1 = ""
                        Str1 = Str1 & "CONTROLLO PATH E DATABASE DA CUI RECUPERARE GLI OGGETTI" & vbCrLf
                        Str1 = Str1 & "-----------------------------------------------------------------" & vbCrLf
                        Str1 = Str1 & Me.sourceDBPath_s_Txt & vbCrLf
                        
                        '//INSERISCO LA PATH NELLA CASELLA DI TESTO ED AGGIORNO
                        Me.TXT_OBJECT_ESTERNI = Str1
                        Me.TXT_OBJECT_ESTERNI.Requery

                            '//PREPARO LA TESTATA DELLE COLLECTION PER IL II° INSERIMENTO
                            Str1 = Str1 & vbCrLf & "                          COLLECTION MODULI DA IMPORTARE : " & vbCrLf
                            Str1 = Str1 & "//=========================================================================//" & vbCrLf
                            ' Stampa
                            Debug.Print
                            Debug.Print
                            Debug.Print Str1
                            
                            '//iterazione nella collezione per il controllo
                            For Each objMdl In objectModules
                                Debug.Print "la MODULI da importare nel progetto: " & objMdl
                                Str1 = Str1 & objMdl & vbCrLf
                            Next objMdl
                            
                            
                  '//--------------------------------------------------------------------------------------------//
                        
        '// *** FINE ***
        '//CONTROLLO DELLE MODULI PRECARICATE NELLA COLLEZIONE CARICO PATH + DB E COLLECTION NELLA CASELLA DI TESTO
        '//=====================================================================================================================//
                        
                        
                        ' Restituisci la collezione
                        Set CollectionModuli_PFunct = objectModules

               


                    '//CHIAMO LA PROCEDURA UNICA DI STAMPA E CARICAMENTO NELLA CASELLA DI TESTO _
                            parametro_01= la collezione, parametro_02 = il tipo
                      Vv1 = Stampa_MSYS_N01_OGGETTI_funct(CollectionModuli_PFunct, "COLLECTION_MODULES")




        
        '//...............................................................................//
        

'//USCITA ED ERRORI
'//--------------------------------------------------------------------------------------//

CollectionModuli_PFunct_Exit:
    Exit Function

CollectionModuli_PFunct_Err:
    MsgBox Error$
    Resume CollectionModuli_PFunct_Exit

End Function


'//LA COLLEZIONE DELLE MODULI DA IMPORTARE  *** FINE ***
'//-----------------------------------------------------------------------------------//











'//LA COLLEZIONE DEI REPORT DA IMPORTARE  *** INIZIO ***
'//-----------------------------------------------------------------------------------//


'//PRECARICO LA COLLECTION REPORT DA IMPORTARE
Public Function CollectionReport_PFunct() As Collection          '//restituisco una collection
    

On Error GoTo CollectionReport_PFunct_Err
        
    Dim frm As Variant             '//OGGETTO REPORT DELLA COLLEZIONE PRECARICATA


    '//CODICE ----> COLLECTION_REPORT_PRECARICATE
    '//CREO LA COLLEZIONE E LA POPOLO
    '//I° COLLEZIONE : OGGETTI REPORT   @importa@collection@oggetti@Report_(la collezione dei @report)
    '//------------------------------------------------------------------------------//
    '//NOTE: questa collection, rappresenta gli OGGETTI REPORT  CHE APPARTENGONO al progetto _
            MSYS_N01_OGGETTI.mdb e che vengono precaricate in un collezione. Quindi tali _
            oggetti devono esere eliminati o aggiunti a seconda che le REPORT del progetto variano. _
            VANTAGGIO: il vantaggio di utilizzare la collezione facilita l'utilizzo delI REPORT di _
            importazione, in quanto le REPORT non sono impostate una tabella esterna che dovrebbe essere _
            spostata insime alI REPORT di importazione. MA ESSENDO PRECARICATE IN UNA COLLECTION si ha _
            sicurezza di spostare tutta la classe REPORT con i dati relativi in questo caso agli oggetti REPORT. _


    '//@OGGETTI@REPORT@PRECARICATI
    '//popola la collection delle REPORT di DA IMPORTARE
    ' Aggiungi i nomi delle REPORT o ELIMINA O COMMENTA LE REPORT DA ESCLUDERE. _
      SE VUOI UTILIZZARE LA COLLEZIONE anche per oggetti REPORT DI PROGETTI DIVERSI inserisci i nome delle REPORT _
      del progetto da importare. N.B. Rircordati di cambiare anche il parametro path+ file per rendere flessibile _
      questa collezione.
        
        Set objectreports = New Collection
        '//Aggiugno gli elementi della collezione
        
        

        '//2024.08.03_Select_Case_Me.Cmb_02_txt = SOSTITUITO IL VALORE CON LA CMB_02 al posto della CMB_01 _
                    Le collection viene scelta in base combinata 2 che individua il progetto _
                    al posto della combinata 1 e la sostituzione riguarda tutte le collection.
        Select Case Me.Cmb_02_txt
            Case "MSYS_N01_OGGETTI.mdb"
            
                '//@popolo@collection@REPORT@system_(delle REPORT di sistema, inserisco nella @collezione)_@carico@collection@REPORT
                objectreports.Add "MsysDbEstTb01Rpt00_}-------------------------------------------@"
                objectreports.Add "MsysDbEstTb01Rpt01_S01_OBJECT"
                objectreports.Add "MsysDbEstTb01Rpt01_S01_OBJECT_TMP"

                
            '//@REPORT@TREEVIEW
            Case "IMPORTA_OGGETTI_TREEVIEW"
                
                objectreports.Add "Tree_Rpt00_}---------------------------------------------------@"
                objectreports.Add "Tree_Rpt01_01_PROVA"
                objectreports.Add "Tree_Rpt90_}---------------------------------------------------@"
                
            '//@REPORTS@MSys
            Case "MSYS_N01_OGGETTI.mdb"
                
                objectreports.Add "MSys_Rpt00_{@------------------REPORT_DB----------------------@}"
                objectreports.Add "MSys_Rpt01_DF01_01_COMANDI_Qry01_01_SELECT_TUTTI"
                objectreports.Add "Msys_Rpt11_DF11_PROGETTI_Qry01_10_GROUP_TUTTI"
                objectreports.Add "Msys_Rpt12_DF12_Qry01_01_TIPO_OGGETTI}"
                
            
        
        '//@IMPORTA@REPORT@DLL@COLECTION   = @IMPORTA@REPORT@DLL_(definisco gli oggetti collection REPORT per le @liberie@DLL)
        Case "IMPORTA_OGGETTI_MSYS_DLL"
                '//2024.08.05_MODIFICHE_IMPORTA_DLL_(per l'importazione delle librerie DLL del progetto mysys inserito il blocco)
                '//@COLLECTION@DLL@FORM
                    objectreports.Add ""
        
        '//@REPORT@ESPROPRI = I REPORT del progetto ESPROPRI
        '//@IMPORTA@REPORT@ESPROPRI@COLECTION   = 2024.08.06_COLLECTION@REPORT@ESPROPRI_(inserito la collezione REPORT per gli espropri)
        Case "IMPORTA_OGGETTI_ESPROPRI"
     
            '//@COLLECTION@ESPROPRI@REPORT
            objectreports.Add "LLPP_TB09Rpt00_}-----------------------------------------------@"
            objectreports.Add "LLPP_TB09Rpt00_01_ESPROPRI"

      
            
        Case Else
            
            MsgBox "ATTENZIONE! NON CI SONO REPORT PER QUESTO PROGETTO: " & Me.Cmb_01_txt, vbCritical, "AVVISO)"
            
            
        End Select
   
        ' Inizializza le collezioni per REPORT COLLEGATE
        Set connectedREPORTS = New Collection
        
      
    '// *** FINE ****
    '//I° COLLEZIONE DI REPORT DI SISTEMA E LA POPOLO
    '//------------------------------------------------------------------------------//
    
        '//CONTROLLO DELLE REPORT PRECARICATE NELLA COLLEZIONE
        '//...............................................................................//
        '//NOTE: queste sono le REPORT da carica nel nuovo progetto.
                
                    
                    '//RIEPILOGO COLLEZIONI
                    '//---------------------------------------------------------//
                           
                            ' Stampa le REPORT fisiche
                            Debug.Print
                            Debug.Print
                            Debug.Print "REPORT DA IMPORTARE:"
                            '//iterazione nella collezione per il controllo
                            For Each frm In objectreports
                                Debug.Print "I REPORT da importare nel progetto: " & frm
                            Next frm
        
                    '//---------------------------------------------------------//
                    
                     
                     

    
        '//CONTROLLO I_REPORT PRECARICATE NELLA COLLEZIONE CARICO PATH + DB E COLLECTION NELLA CASELLA DI TESTO
        '//=====================================================================================================================//
        '//NOTE: queste sono I_REPORT da carica nel nuovo progetto.
                        
                        '//RESET CASELLA DI TESTO
                        Me.TXT_OBJECT_ESTERNI = ""
                        Me.TXT_OBJECT_ESTERNI.Requery
        
                                    
                  '//IMPOSTO LA PATH + FILE.MDB + TESTATA + OGGETTI VISUALIZZAZIONE DELLA CASELLA TXT
                  '//--------------------------------------------------------------------------------------------//
                  '//Note: @imposto@path_(imposto la path degli oggetti da recupeare, _
                           preparo la testata per la casela di testo, preparo gli oggetti e visualizzo tutto _
                           nella casella di testo.
                    
                    
                    '//RIEPILOGO COLLEZIONI E VISUALIZZAZIONE PATH ED OGGETTI PRECARICATI
                        Str1 = ""
                        Str1 = Str1 & "CONTROLLO PATH E DATABASE DA CUI RECUPERARE GLI OGGETTI" & vbCrLf
                        Str1 = Str1 & "-----------------------------------------------------------------" & vbCrLf
                        Str1 = Str1 & Me.sourceDBPath_s_Txt & vbCrLf
                        
                        '//INSERISCO LA PATH NELLA CASELLA DI TESTO ED AGGIORNO
                        Me.TXT_OBJECT_ESTERNI = Str1
                        Me.TXT_OBJECT_ESTERNI.Requery

                            '//PREPARO LA TESTATA DELLE COLLECTION PER IL II° INSERIMENTO
                            Str1 = Str1 & vbCrLf & "                          COLLECTION I_REPORT DA IMPORTARE : " & vbCrLf
                            Str1 = Str1 & "//=========================================================================//" & vbCrLf
                            ' Stampa
                            Debug.Print
                            Debug.Print
                            Debug.Print Str1
                            
                            '//iterazione nella collezione per il controllo
                            For Each objRpt In objectreports

                                Debug.Print "I_REPORT da importare nel progetto: " & objRpt
                                Str1 = Str1 & objRpt & vbCrLf
                                
                            Next objRpt
                            
                            
                            
                  '//--------------------------------------------------------------------------------------------//
                        
        '// *** FINE ***
        '//CONTROLLO I_REPORT PRECARICATE NELLA COLLEZIONE CARICO PATH + DB E COLLECTION NELLA CASELLA DI TESTO
        '//=====================================================================================================================//
                                             
                     
                     
                     
                     ' Restituisci la collezione
                        Set CollectionReport_PFunct = objectreports
        

                    '//CHIAMO LA PROCEDURA UNICA DI STAMPA E CARICAMENTO NELLA CASELLA DI TESTO _
                            parametro_01= la collezione, parametro_02 = il tipo
                      Vv1 = Stampa_MSYS_N01_OGGETTI_funct(CollectionModuli_PFunct, "COLLECTION_REPORTS")
       



        
'//USCITA ED ERRORI
'//--------------------------------------------------------------------------------------//

CollectionReport_PFunct_Exit:
    Exit Function

CollectionReport_PFunct_Err:
    MsgBox Error$
    Resume CollectionReport_PFunct_Exit

End Function



'//=====================================================================================================================//
'//                     CASELLE DI TESTO E COMBINATE DEL PROGETTO IMPORTA  *** INIZIO ***
'//                     + CurrentProject_txt + Cmb_01_txt + Cmb_02_txt + sourceDBPath_s_Txt  = @CMB_01, CMB_02 @CMB @PROGETTO
'//=====================================================================================================================//
'//NOTE: @caselle@combinate@progetto_(le caselle di testo e le caselle combinate che gestiscono il progetto + @cmb@help)





'//=====================================================================================================================//
'//                                             @CMB_01      *** INIZIO ***
'//
'//=====================================================================================================================//

'//CASELLA COMBINATA DI SCELTA DEL PROGETTO @CMB_01@CHANGE
'//-------------------------------------------------------------------------------------------//
'//@PATH@PROGETTI_(la path dei progetti da caricare)'
'//@CHANGE@CACELLA@COMBINATA_(se cambia la @casella@combinata abilito o disabilito i button)
'//NOTE
Private Sub Cmb_01_txt_AfterUpdate()

         
   
'//DIM della collezione form da importare
Dim formsCollection As Collection
Dim formName_v As Variant

'//DIM la collezione macro
Dim macroCollection As Collection
Dim macroName_v As Variant



'//DIM la collezione DEI MODULI
Dim modulesCollection As Collection
Dim moduloName_v As Variant


'//DIM la collezione DELLE QUERY
Dim queryCollection As Collection
Dim queryName_v As Variant




'//DIM la collezione DELLE REPORT
Dim reportCollection As Collection
Dim reportName_v As Variant



'//DIM la collezione DELLE TABELLE
Dim tableCollection As Collection
Dim tableName_v As Variant


            
    '//RESET I BUTTON
    '//........................................//
    '//note : RESET VARIABILI
        PROGETTO_s = ""
        
        '//RESET
        Me.TXT_OBJECT_ESTERNI = ""
       
       
    '//........................................//
    
    
    '//@CMB_01_@DOPO@AGGIORNAMENTO
    '//ABILITO O DISABILITO I COMAND BUTTON + IMPOSTO PATH E DB PER LA CASELLA DI TESTO - LA COMBINATA 01 - @CMB_01
    '//===============================================================================================//
    '//NOTE: @scelta@progetto_(casella combinata di scelta del progetto)
    '//      @imposta@path_(per tutti i progetti vengon impostati nella combinata 01)
    
    Select Case Me.Cmb_01_txt.Value
    
        
            '//@ELENCO@PROGETTI
            '//2024.08.03_Cmb_01_txt_AfterUpdate = inserito il case ELENCO PROGETTI
            '//@MENU@PROGETTI
            Case "ELENCO PROGETTI"
                        
                        
                        
                        'se non esiste nessun progetto disabilito i button ed imposto la path a null
                        sourceDBPath_s = ""
                        '//La variabile path aggiornata la salvo nella casella di testo:
                        Me.TXT_OBJECT_ESTERNI = ""
                        Me.sourceDBPath_s_Txt = ""
                        Me.sourceDBPath_s_Txt.Requery
                        
                        '//2024.08.03_Disabilito_sourceDBPath_Sub = creato la Disabilitazione della path
                        Call Disabilito_sourceDBPath_Sub      '//DISABILITO LA PATH
                        Call DisabilitaButton_Sub
                        
                        Me.Cmb_02_txt = ""
                        Me.Cmb_02_txt.Requery
                        
            '//@MENU@OGGETTI
            Case "MSYS_N01_OGGETTI.mdb"
                        
                        
                        '//TODO: nel caso di scelta DI MSYS occorre anticipare il controllo progetto, perche deve essere _
                                ed inibita l'importazione o la cancellazione degli oggetti.
                        
                        
                        '//IMPOSTO IL TITOTOLO DEL PROGETTO
                        PROGETTO_s = "PROGETTO MSYS_N01_OGGETTI"
                        
                        sourceDBPath_s = ""
                        Call DisabilitaButton_Sub
                          
                        '//MSYS_N01_OGGETTI.mdb = PER QUESTO DATABASE ABILITO LA PATH
                        ' Percorso del database di origine - todo: eliminato
                        sourceDBPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSYS_N01_OGGETTI.mdb"
                        Me.sourceDBPath_s_Txt = sourceDBPath_s
                        
                        '//Abilito la cmb_02 + la path
                        AbilitaCmd_02_Sub
                        
                        '//2024.08.03_Abilito_sourceDBPath_Sub = creato l'abilitazione della path
                        Call Abilito_sourceDBPath_Sub      '//ABILITO LA PATH
                        Call AbilitaButton_CMB_02_Sub      '//ABILITO LA CMB_02
    
        
                         '//2024.08.08_CMB_02_SELEZIONE_OGGETTI_IMPOSTA_PATH_MODIFICA
                         '//@CMB_02@IMPOSTO@_(I VALORI DELLA COMBINATA 02)
                         '//-------------------------------------------------------------------------//
                         '//@CMB_02@imposto_(reimposto agli oggetti TreeView della @casella@combinata)
                             Me.Cmb_02_txt.Value = ""
                             Str1 = ""
                             Str1 = Str1 & "OGGETTI MSYS;"
                             Str1 = Str1 & "IMPORTA_OGGETTI_MSYS_DLL;"
                             Str1 = Str1 & "IMPORTA_OGGETTI_MSYS_02??;"
                             Str1 = Str1 & ""
                         
                         
                             Me.Cmb_02_txt.RowSourceType = "elenco valori"
                             Me.Cmb_02_txt.RowSource = Str1
                             Me.Cmb_02_txt.Value = "OGGETTI MSYS"
                         '//-------------------------------------------------------------------------//
                            
                             '//RESET
                            Me.TXT_OBJECT_ESTERNI = ""
                            Me.Recalc
                            
                           '//02) SECONDO CONTROLLO
                           '//CHIAMO IL CONTROLLO DEL PROGETTO
                           '//..........................................................//
                           '//Note  : True = siamo nel progetto originale non possiamo _
                                      importare nessun oggetto perche distruggiamo il db corrente. _
                                      False = siamo in un progetto esterno possiamo caricare tutto.
                                          
                               '//SE VIENE RESTITUITO TRUE è FUORI PROGETTO PUO CARICARE, SE FALSE EXIT PERCHE PROGETTO ORIGINALE
                               Bool1 = CheckProjectName_b_PFunct
                               
                               If Bool1 = True Then Exit Sub
                           '//..........................................................//
                       
            '//@MENU@ESPROPRI
            Case "GE_ESPROPRI.mdb"
                                        
                        '//RESET
                        Me.TXT_OBJECT_ESTERNI = ""
                        Me.Recalc
                            
                    
                        '//IMPOSTO IL TITOTOLO DEL PROGETTO
                        PROGETTO_s = "PROGETTO ESPROPRI"
                        
                        '//Abilito la cmb_02 + la path + imposto la path
                        Call DisabilitaButton_Sub
                        Call Abilito_sourceDBPath_Sub      '//ABILITO LA PATH
                        Call AbilitaButton_CMB_02_Sub      '//ABILITO LA CMB_02
                        
                        sourceDBPath_s = "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\ESPROPRI\MDB\GE_ESPROPRI.mdb"       '//percorso ESPROPRI
                        Me.sourceDBPath_s_Txt = sourceDBPath_s
                    
        
                         '//2024.08.08_CMB_02_SELEZIONE_OGGETTI_IMPOSTA_PATH_MODIFICA
                         '//@CMB_02@IMPOSTO@_(I VALORI DELLA COMBINATA 02)
                         '//-------------------------------------------------------------------------//
                         '//@CMB_02@imposto_(reimposto agli oggetti TreeView della @casella@combinata)
                             Me.Cmb_02_txt.Value = ""
                             Str1 = ""
                             Str1 = Str1 & "OGGETTI ESPROPRI;"
                             Str1 = Str1 & "IMPORTA_OGGETTI_ESPORPRI_01??;"
                             Str1 = Str1 & "IMPORTA_OGGETTI_ESPORPRI_02??;"
                             Str1 = Str1 & ""
                         
                         
                             Me.Cmb_02_txt.RowSourceType = "elenco valori"
                             Me.Cmb_02_txt.RowSource = Str1
                             Me.Cmb_02_txt.Value = "OGGETTI ESPROPRI"
                         '//-------------------------------------------------------------------------//
                            
                             '//RESET
                            Me.TXT_OBJECT_ESTERNI = ""
                            Me.Recalc
                       
            '//@MENU@TREEVIEW
            Case "TREEVIEW.mdb"
                    
                    '//RESET
                    Me.TXT_OBJECT_ESTERNI = ""
                        
                    
                        '//IMPOSTO IL TITOTOLO DEL PROGETTO
                        PROGETTO_s = "PROGETTO TREEVIEW"
                        
                        '//Abilito la cmb_02 + la path + imposto la path
                        Call DisabilitaButton_Sub
                        sourceDBPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\TREVIEW\MDB\TREEVIEW.mdb"       '//percorso Treeview
                        
                        Call Abilito_sourceDBPath_Sub      '//ABILITO LA PATH
                        Call AbilitaButton_CMB_02_Sub      '//ABILITO LA CMB_02
                        
                        Me.sourceDBPath_s_Txt = sourceDBPath_s
                        Me.sourceDBPath_s_Txt.Requery
                        
        
                         '//2024.08.08_CMB_02_SELEZIONE_OGGETTI_IMPOSTA_PATH_MODIFICA
                         '//@CMB_02@IMPOSTO@_(I VALORI DELLA COMBINATA 02)
                         '//-------------------------------------------------------------------------//
                         '//@CMB_02@imposto_(reimposto agli oggetti TreeView della @casella@combinata)
                             Me.Cmb_02_txt.Value = ""
                             Str1 = ""
                             Str1 = Str1 & "OGGETTI TREVIEW;"
                             Str1 = Str1 & "IMPORTA_OGGETTI_TREEVIEW;"
                             Str1 = Str1 & "IMPORTA_OGGETTI_??;"
                             Str1 = Str1 & ""
                         
                         
                             Me.Cmb_02_txt.RowSourceType = "elenco valori"
                             Me.Cmb_02_txt.RowSource = Str1
                             Me.Cmb_02_txt.Value = "OGGETTI TREVIEW"
                         '//-------------------------------------------------------------------------//
                        
                            
                             '//RESET
                            Me.TXT_OBJECT_ESTERNI = ""
                            Me.Recalc
                       
                    
            '//@MENU@ESPROPRI
            Case "GE_ESPROPRI.mdb"
                    
                                  '//GE_ESPROPRI_MSYS_N01_OGGETTI.mdb = PER QUESTO DATABASE ABILITO LA PATH
                      ' Percorso del database di origine - todo: eliminato
                      sourceDBPath_s = "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\ESPROPRI\MDB\GE_ESPROPRI.mdb"
                      Call DisabilitaButton_Sub
                        
                        Call Abilito_sourceDBPath_Sub      '//ABILITO LA PATH
                        Call AbilitaButton_CMB_02_Sub      '//ABILITO LA CMB_02
                        
                        
                        '//IMPOSTO IL TITOTOLO DEL PROGETTO
                        PROGETTO_s = "PROGETTO ESPROPRI"
                        
                        '//VISUALIZZO NELLA CASELLA DI TESTO
                        Me.TXT_OBJECT_ESTERNI = Stampa_N02_OGGETTI_ESPROPRI_funct
                        
                        
                                                 '//2024.08.08_CMB_02_SELEZIONE_OGGETTI_IMPOSTA_PATH_MODIFICA
                         '//@CMB_02@IMPOSTO@_(I VALORI DELLA COMBINATA 02)
                         '//-------------------------------------------------------------------------//
                         '//@CMB_02@imposto_(reimposto agli oggetti TreeView della @casella@combinata)
                             Me.Cmb_02_txt.Value = ""
                             Str1 = ""
                             Str1 = Str1 & "OGGETTI ESPROPRI;"
                             Str1 = Str1 & "IMPORTA_OGGETTI_EPROPRI_01??;"
                             Str1 = Str1 & "IMPORTA_OGGETTI_EPROPRI_02??;"
                             Str1 = Str1 & ""
                         
                         
                             Me.Cmb_02_txt.RowSourceType = "elenco valori"
                             Me.Cmb_02_txt.RowSource = Str1
                             Me.Cmb_02_txt.Value = "OGGETTI ESPROPRI"
                         '//-------------------------------------------------------------------------//
                        

                             '//RESET
                            Me.TXT_OBJECT_ESTERNI = ""
                            Me.Recalc
                        
                '//@MENU@CONDOMINIO
                Case "GE_CDM.mdb"
                    
                    '//GE_ESPROPRI_MSYS_N01_OGGETTI.mdb = PER QUESTO DATABASE ABILITO LA PATH
                      ' Percorso del database di origine - todo: eliminato
                      sourceDBPath_s = "c:\CASA\CDM\GE_COND\"
                      Call DisabilitaButton_Sub
                        
                        Call Abilito_sourceDBPath_Sub      '//ABILITO LA PATH
                        Call AbilitaButton_CMB_02_Sub      '//ABILITO LA CMB_02
                          
                        '//IMPOSTO IL TITOTOLO DEL PROGETTO
                        PROGETTO_s = "PROGETTO CONDOMINIO"
                        
                        '//VISUALIZZO NELLA CASELLA DI TESTO
                        Me.TXT_OBJECT_ESTERNI = Stampa_N03_OGGETTI_CONDOMINIO_funct
                        
                        
                                                 '//2024.08.08_CMB_02_SELEZIONE_OGGETTI_IMPOSTA_PATH_MODIFICA
                         '//@CMB_02@IMPOSTO@_(I VALORI DELLA COMBINATA 02)
                         '//-------------------------------------------------------------------------//
                         '//@CMB_02@imposto_(reimposto agli oggetti TreeView della @casella@combinata)
                             Me.Cmb_02_txt.Value = ""
                             Str1 = ""
                             Str1 = Str1 & "OGGETTI ESPROPRI;"
                             Str1 = Str1 & "IMPORTA_OGGETTI_EPROPRI_01??;"
                             Str1 = Str1 & "IMPORTA_OGGETTI_EPROPRI_02??;"
                             Str1 = Str1 & ""
                         
                         
                             Me.Cmb_02_txt.RowSourceType = "elenco valori"
                             Me.Cmb_02_txt.RowSource = Str1
                             Me.Cmb_02_txt.Value = "OGGETTI ESPROPRI"
                         '//-------------------------------------------------------------------------//
                        

                             '//RESET
                            Me.TXT_OBJECT_ESTERNI = ""
                            Me.Recalc
                        
                        
        
            Case Else
                        'se non esiste nessun progetto disabilito i button ed imposto la path a null
                        sourceDBPath_s = ""
                        '//La variabile path aggiornata la salvo nella casella di testo:
                        Me.TXT_OBJECT_ESTERNI = ""
                        Me.sourceDBPath_s_Txt = ""
                        Me.sourceDBPath_s_Txt.Requery
                         
                        
                        Me.Cmb_02_txt = ""
                        Me.Cmb_02_txt.Requery
                        
                         
                        
                    
                    
                        
                        '//DISABILITO LA CMB_02
                        Call Disabilito_sourceDBPath_Sub      '//DISABILITO LA PATH
                        Call DisabilitaButton_CMB_02_Sub      '//DISABILITO LA CMB_02
                                       
                        
            End Select
            
                    
        
            '//ABILITO O DISABILITO I COMAND BUTTON + IMPOSTO PATH E DB PER LA CASELLA DI TESTO
            '//===============================================================================================//
                          

End Sub


'//HELP@CMB_01
Private Sub Cmb_01_txt_HELP_Click()
'//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
    '//--------------------------------------------------------------------------------//--------//
    '//NOTE                 -> Apro il file di tipo doc, zip o pdf
    '//CODICE               -> Function ApriFilePdf.01.01
    '//PARAMETRI            -> par_Path_s         = PATH _
                            -> par_NameFile_s               = NOME FILE _
                            -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                            
      '//IMPOSTO LE VARIABILI PATH E FILE
      '//
      Dim MyFile_s As String
      '//
      Dim MyPath_s  As String
      
            '//imposto i parametri - LA PATH
            '
            MyPath_s = ""
            MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\OBJECT\TUTORIAL\"
            '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
            '
            MyFile_s = "APRI_MSYS.bat"
                
                
                       
         '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
            Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto IMPORTA OGGETTI incluso in questa form _
            anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'importazione e cancellazione in caso di resettamento _
            totale.
            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
      
            
    '//--------------------------------------------------------------------------------//--------//

End Sub


'//=====================================================================================================================//
'//                                             @CMB_01      *** FINE ***
'//
'//=====================================================================================================================//





'//=====================================================================================================================//
'//                                             @CMB_02      *** INIZIO ***
'//
'//=====================================================================================================================//





'//todo: finre AL POSTO DI CREARE QUESTO EVENTO OCCORRE alla scelta IMPORTA OGGETT MSYS DLL, ossia importa _
         le tabelle, query, form, report, moduli e macro utili per la lavorazione delle librerie ad oggetti _
            1) abilitare il comando di importazione delle tabelle e di cancellazione.
            
'//CMB_02= CASELLA COMBINATA DI SCELTA DEL PROGETTO @CMB_02@CHANGE
'//-------------------------------------------------------------------------------------------//
'//Note: la combinata 02 deve solo filtrare gli oggetti che appartengono al progetto principale senza _
        commistioni es. SE LA CMB_01 imposta il progetto MSYS la CMB_02 deve filtrare solo OGGETTI MSYS e non _
        altro come ad esempio oggetti TREEVIEW che sarebbe errato in quanto la path impostata DALLA COMBINATA 01 è per gli _
        oggetti MSYS  e non per i Treeview.


Private Sub Cmb_02_txt_DblClick(Cancel As Integer)
    Cmb_02_txt_AfterUpdate
End Sub

Private Sub Cmb_02_txt_AfterUpdate()

    Select Case Me.Cmb_02_txt.Value
        
        '//2024.08.08_CMB_02_SELEZIONE_OGGETTI_IMPOSTA_PATH_MODIFICA
        '//IMPORTA OGGETTI LIBRERIE DLL  @CMB_02@IMPORTA@MSYS
        '//........................................................//
        Case "IMPORTA_OGGETTI_MSYS_DLL"
            
            '01_ABLITAZIONE_BUTTON= solo i comandi IMPORT e DELETE
            AbilitaButton_TABLE_Sub
            AbilitaButton_MODULI_Sub
            AbilitaButton_QUERY_Sub
            AbilitaButton_MACRO_Sub
            AbilitaButton_IMPORTA_TUTTO_Sub
            
            
            AbilitaButton_CMB_02_Sub
    
            
        '//........................................................//
                
        '//IMPORTA OGGETTI TREEVIEW
        '//........................................................//
        Case "IMPORTA_OGGETTI_TREEVIEW"
            
            '01_ABLITAZIONE_BUTTON= solo i comandi IMPORT e DELETE
            AbilitaButton_FORM_Sub
            AbilitaButton_TABLE_Sub
            AbilitaButton_MODULI_Sub
            AbilitaButton_QUERY_Sub
            AbilitaButton_MACRO_Sub
            AbilitaButton_IMPORTA_TUTTO_Sub
            
            AbilitaButton_CMB_02_Sub
    
        '//........................................................//
        
        
        '//IMPORTA OGGETTI CONDOMINIO
        '//........................................................//
        Case "IMPORTA_OGGETTI_CONDOMINIO"
            
            '01_ABLITAZIONE_BUTTON= solo i comandi IMPORT e DELETE
            AbilitaButton_FORM_Sub
            AbilitaButton_TABLE_Sub
            AbilitaButton_MODULI_Sub
            AbilitaButton_QUERY_Sub
            AbilitaButton_MACRO_Sub
            AbilitaButton_IMPORTA_TUTTO_Sub
            
            AbilitaButton_CMB_02_Sub
    
        '//........................................................//
        
    Case Else
        
        '//se non ci sono oggetti da importare disabilito tutti i button
                DisabilitaButton_FORM_Sub
                DisabilitaButton_TABLE_Sub
                DisabilitaButton_MODULI_Sub
                DisabilitaButton_QUERY_Sub
                DisabilitaButton_MACRO_Sub
                DisabilitaButton_IMPORTA_TUTTO_Sub
                DisabilitaButton_DELETE_TUTTO_Sub
                
                 '//RITORNO ALLA COMBINATA 01
                Me.Cmb_01_txt.SetFocus
                      
                DisabilitaButton_CMB_02_Sub
                
    
    End Select
    
    
End Sub


'//CMB_02 = CASELLA COMBINATA DI SCELTA DEL PROGETTO   *** FINE ***

'//-------------------------------------------------------------------------------------------//



'//=====================================================================================================================//
'//                                             @CMB_02      *** FINE ***
'//
'//=====================================================================================================================//







'//=====================================================================================================================//
'//                     CASELLE DI TESTO E COMBINATE DEL PROGETTO IMPORTA  *** FINE ***
'//                     + CurrentProject_txt + Cmb_01_txt + Cmb_02_txt + sourceDBPath_s_Txt  = @CMB_01, CMB_02 @CMB @PROGETTO
'//=====================================================================================================================//















'//@attiva@controllo@oggetti@esterni
'//CHIAMO LA FUNZIONE DI CONTROLLO OGGETTI DI SISTEMA ESTERNI
'//Note: controllo oggetti esterni.@CONTROLLO@OGGETTI_(attivo la procedura per il controllo degli oggetti esterni)
Private Sub Cmd_CONTROLLO_OGGETTI_ESTERNI_Click()
    ListExternalDBObjects
End Sub



'//LA COLLEZIONE DEI REPORT DA IMPORTARE  *** FINE ***
'//-----------------------------------------------------------------------------------//








'//=====================================================================================================================//
'//                     DELETE OGGETTI QUERY DEL DB ESTERNO IN TABELLA *** INIZIO ***
'//                     + DELETE QUERY HELP
'//=====================================================================================================================//


'//CANCELLA GLI OGGETTI QUERY *** INIZIO *** '//DELETE QUERIES
'//==========================================================================================================//
'//@GESTIONE@DELLE@QUERY                    @GESTONE@QUERY

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
    iCount = 0
    deleteQueryesCount_i = 0
    deletedTablesCount_i = 0
    deletedTables_s = ""
   
        
        
        
                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** INIZIO ***
                '//=====================================================================================================================//
                '//NOTE: chiamo la funzione che controlla se il progetto corrente è il progetto originale, infatti _
                         non si possono caricare gli oggetti nel PROGETTO MSYS_N01_OGGETTI.mdb perhè è originale _
                         e vi è una distruzione degli stessi. LA FUNZIONE CONTROLLO SE E' UN PROGETTO ESTERNO AMMISSIBILE _
                         per caricare gli oggetti table, query , form ecc. da importare. _
                         01) CONTROLLO = MSG DI AVVISO CON POSSIBILITA DI ANNULLARE _
                         02) CONTROLLO = CONTROLLO SE STIAMO NEL PROGETTO ORIGINALE  E QUINDI ESCLUSIONE.
                                 
                           '//01) PRIMO CONTROLLO _
                             MESSAGGIO DI IMPORTAZIONE 0 = ANNULLA  1 = CONTINUA
                           '//..........................................................//
                             '//Attenzione L'AVVISO produce il messaggio solo se il _
                                parametro è TRUE altrimenti il messaggio viene escluso.
                             MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)
                             
                             If MyValue = 0 Then Exit Sub
                           '//..........................................................//
                                 
                           '//02) SECONDO CONTROLLO
                           '//CHIAMO IL CONTROLLO DEL PROGETTO
                           '//..........................................................//
                           '//Note  : True = siamo nel progetto originale non possiamo _
                                      importare nessun oggetto perche distruggiamo il db corrente. _
                                      False = siamo in un progetto esterno possiamo caricare tutto.
                                          
                               '//SE VIENE RESTITUITO TRUE è FUORI PROGETTO PUO CARICARE, SE FALSE EXIT PERCHE PROGETTO ORIGINALE
                               Bool1 = CheckProjectName_b_PFunct
                               
                               If Bool1 = True Then Exit Sub
                           '//..........................................................//
                
                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** FINE ***
                '//=====================================================================================================================//

              '//ATTIVO LA COLLEZIONE
              '//------------------------------------------------------------------------//
              '//Note: CHIAMO LA FUNZIONE COLLECTION QUERY per ottenere la collezione di QUERY
                
                
                
                '//
                Set queryesCollection = CollectionQuery_PFunct()
                
              '//------------------------------------------------------------------------//
                    
                    ' Loop attraverso tutte I REPORT nel database corrente
                     Set dbCurrent = CurrentDb
                     
                     '//CALCOLO INDICE QUERY
                     iCount = dbCurrent.QueryDefs.count
                   
                    Set dbCurrent = CurrentDb

                        
                    '//se ci sono query da cancellare
                    If iCount > 0 Then
                    
                        '//reimposto la matrice sulle query esistenti nel db corrente
                        ReDim MATRICE_queryNames_s(iCount - 1)
                        
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
                                    
                                    If Err.number <> 0 Then
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


'//DELETE QUERY HELP *** INIZIO ***
'//-----------------------------------------------------------------------------------//

Private Sub Cmd_Delete_QUERY_HELP_Click()

    '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
    '//--------------------------------------------------------------------------------//--------//
    '//NOTE                 -> Apro il file di tipo doc, zip o pdf
    '//CODICE               -> Function ApriFilePdf.01.01
    '//PARAMETRI            -> par_Path_s         = PATH _
                            -> par_NameFile_s               = NOME FILE _
                            -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                            
      '//IMPOSTO LE VARIABILI PATH E FILE
      '//
      Dim MyFile_s As String
      '//
      Dim MyPath_s  As String
      
            '//imposto i parametri - LA PATH
            '
            MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\OBJECT\HELP\"
            '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
            '
            MyFile_s = "Help_delete_QUERY.pdf "
                
                
                       
         '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
            Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto IMPORTA OGGETTI incluso in questa form _
            anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'importazione e cancellazione in caso di resettamento _
            totale.
            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
      
            
    '//--------------------------------------------------------------------------------//--------//

End Sub


'//DELETE QUERY HELP *** FINE ***
'//-----------------------------------------------------------------------------------//


'//=====================================================================================================================//
'//                     DELETE OGGETTI QUERY DEL DB ESTERNO IN TABELLA *** FINE ***
'//                     + DELETE QUERY HELP
'//=====================================================================================================================//







'//************************************************************************************************************************//
'//***
'//***                              COLLECTION OBJECT   *** FINE ***
'//************************************************************************************************************************//



'//=====================================================================================================================//
'//                     CMD_IMPORTA OGGETTI FORM DEL DB ESTERNO *** INIZIO ***
'//                     @IMPORTA@OGGETTI@FORM@MSYS
'//=====================================================================================================================//



'//IMPORTA TUTTI GLI OGGETTI FORM DEL DB MSYSDBEST
'//----------------------------------------------------------------------------------------//
Private Sub Cmd_Importa_FORM_Click()
    '//DIM della collezione form da importare
    Dim formsCollection As Collection
    Dim formName_v As Variant
    Dim sourceDB As DAO.Database
    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim sourceDBPath As String
    Dim myformName_s As String
    Dim importedCount_i As Integer
    Dim deletedCount_i As Integer
    Dim deletedForms_s As String
    Dim importedForms_s As String
    Dim Bool1 As Boolean
    Dim MyValue As Integer

    On Error GoTo ErrorHandler
    
    ' Inizializza i contatori e le stringhe di riepilogo
    importedCount_i = 0
    deletedCount_i = 0
    deletedForms_s = ""
    importedForms_s = ""

    ' RECUPERO LA COLLEZIONE FORM DA IMPORTARE E FACCIO IL CONTROLLO
    Set formsCollection = CollectionForm_PFunct()

    ' SE LA COLLECTION E' VUOTA ESCI DALLA ROUTINE
    If formsCollection.count = 0 Then Exit Sub

    ' Controllo path ed il file.mdb per le importazioni
    sourceDBPath = PathFile_s_pFunct
    If sourceDBPath = "" Then
        MsgBox "ATTENZIONE ERRORE path e file sono null uscita dalla routine", vbCritical, "MSG CONTROLLO PATH E FILE VALORIZZATI"
        Exit Sub
    End If

    ' Primo controllo progetto
    MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)
    If MyValue = 0 Then Exit Sub

    ' Secondo controllo progetto
    Bool1 = CheckProjectName_b_PFunct
    If Bool1 = True Then Exit Sub

    ' Se siamo su un progetto esterno (Bool1 = False), procedo con l'importazione
    If Bool1 = False Then
        ' Apri il database di origine e il database di destinazione
        Set ws = DBEngine.Workspaces(0)
        Set sourceDB = ws.OpenDatabase(sourceDBPath)
        Set db = CurrentDb

        ' Ciclo nella collezione per il controllo delle form precaricate
        For Each formName_v In formsCollection
            myformName_s = formName_v
            
            ' Controlla se il modulo esiste nel database di origine
            If ObjectExists(sourceDB, acForm, myformName_s) Then
                ' Se esiste nel db esterno, controlla se esiste anche nel db corrente e cancellalo
                If ObjectExists(db, acForm, myformName_s) Then
                    DoCmd.DeleteObject acForm, myformName_s
                    deletedCount_i = deletedCount_i + 1
                    deletedForms_s = deletedForms_s & myformName_s & vbCrLf
                End If
                
                ' Importa il modulo dal database di origine
                DoCmd.TransferDatabase acImport, "Microsoft Access", sourceDBPath, acForm, myformName_s, myformName_s
                importedCount_i = importedCount_i + 1
                importedForms_s = importedForms_s & myformName_s & vbCrLf
            Else
                MsgBox "La form '" & myformName_s & "' non esiste nel database di origine.", vbExclamation
            End If
        Next formName_v

        ' Chiudi il database di origine
        sourceDB.Close
        Set sourceDB = Nothing
        Set ws = Nothing

        ' Messaggio finale di riepilogo
        If Me.ctr_AVVISO_ABILITATO_TXT = True Then
            MsgBox "Importazione completata." & vbCrLf & _
                "Forms importate: " & importedCount_i & vbCrLf & importedForms_s & _
                "Forms esistenti cancellate: " & deletedCount_i & vbCrLf & deletedForms_s, vbInformation
        End If

        ' Rilascia gli oggetti
        Set db = Nothing
        Set formsCollection = Nothing
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Errore durante l'importazione: " & Err.Description, vbCritical
    If Not sourceDB Is Nothing Then
        sourceDB.Close
    End If
    Set sourceDB = Nothing
    Set ws = Nothing
End Sub


'//2024.08.06_EFFICIENTAMENTO_(efficientata la procedura con l'interrogazione sulla tabella MSysObjects per il controllo delle form esterne)
'//FUNZIONE DI CONTROLLO SE ESISTE L'OGGETTO
Function ObjectExists(db As DAO.Database, objType As Integer, objName As String) As Boolean
    Dim rs As DAO.Recordset
    Dim moduleExists As Boolean
    
    moduleExists = False
    On Error Resume Next
    Set rs = db.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type=-32768 AND Name='" & objName & "'")
    If Not rs.EOF Then
        moduleExists = True
    End If
    rs.Close
    Set rs = Nothing
    ObjectExists = moduleExists
    On Error GoTo 0
End Function
'//IMPORTA TUTTI GLI OGGETTI FORM DEL DB MSYSDBEST *** FINE ***
'//----------------------------------------------------------------------------------------//



'//         IMPORTA PROGETTO OGGETTI FORM *** FINE ***
'//=========================================================================================//



'//=====================================================================================================================//
'//                     CMD_IMPORTA OGGETTI FORM DEL DB ESTERNO *** FINE ***
'//                     @IMPORTA@OGGETTI@FORM@MSYS
'//=====================================================================================================================//




'//=====================================================================================================================//
'//                     DELETE OGGETTI FORM DEL DB ESTERNO IN TABELLA *** INIZIO ***
'//                     + DELETE TABELLA HELP
'//=====================================================================================================================//



'//CANCELLA GLI OGGETTI FORM *** INIZIO *** DELETE_FORM
'//======================================================================================================//
'//@GESTIONE@DELLE@FORM    (@MODELLO@DELETE@OGGETTI)

'//NOTA: visto che la cancellazione riguarda il database corrente in cui è stata importata _
        la classe di IMPORTA PROGETTO, la routine cancella solo se le FORM esistono _
        nel database corrente mediante il cofronto tra : _
        LA COLLECTION FORM = GLI OGGETTI FORM  che sono attivi nel database ESTERNO. _
        Prima di cancellare viene effettuato il confronto di esistenza _
        @cancellazione@form@caricate_(cancellazione delle FORM gia caricate nel db esterno) _
        @cancella@TABLE@esterne_(attenzione puoi cancellare le FORM Nel @db@esterno)
            
'//FAQ : _
        DOVE SI CANCELLANO LE FORM PRECARICATE _
        COME CANCELLARE LE FORM ESTERNE _
        LE FORM ESTERNE DA CANCELLARE


'// PARAMETRI DA CAMBIARE : _
    -----> FORM _
    -----> DELLE@FORM _
    -----> DELETE_FORM _
    -----> CollectionForm_PFunct _
    -----> formsCollection _
    -----> objectFomr _
    -----> MATRICE_FORM_Names_s _
    -----> formsCount_i _
    -----> FormDef _
    -----> frm _


Private Sub Cmd_Delete_FORM_Click()

    Dim dbCurrent As DAO.Database
    Dim frm As Form
    Dim formsCollection As Collection
    Dim varItem As Variant
    Dim MATRICE_FORM_Names_s() As String             'matrice di FORM
    Dim tempCount As Integer
    Dim formsCount_i As Integer
    Dim fdf As AccessObject
    Dim iCount As Integer
    Dim i As Integer
    Dim Bool1 As Boolean
    
    On Error GoTo Err_Cmd_Delete_FORM_Click

    '//RESET
    iCount = 0
    deletetFormsCount_i = 0
    deletedForms_s = ""
    
    
        
                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** INIZIO ***
                '//=====================================================================================================================//
                '//NOTE: chiamo la funzione che controlla se il progetto corrente è il progetto originale, infatti _
                         non si possono caricare gli oggetti nel PROGETTO MSYS_N01_OGGETTI.mdb perhè è originale _
                         e vi è una distruzione degli stessi. LA FUNZIONE CONTROLLO SE E' UN PROGETTO ESTERNO AMMISSIBILE _
                         per caricare gli oggetti table, query , form ecc. da importare. _
                         01) CONTROLLO = MSG DI AVVISO CON POSSIBILITA DI ANNULLARE _
                         02) CONTROLLO = CONTROLLO SE STIAMO NEL PROGETTO ORIGINALE  E QUINDI ESCLUSIONE.
                                 
                           '//01) PRIMO CONTROLLO _
                             MESSAGGIO DI IMPORTAZIONE 0 = ANNULLA  1 = CONTINUA
                           '//..........................................................//
                             '//Attenzione L'AVVISO produce il messaggio solo se il _
                                parametro è TRUE altrimenti il messaggio viene escluso.
                             MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)
                             
                             If MyValue = 0 Then Exit Sub
                           '//..........................................................//
                                 
                           '//02) SECONDO CONTROLLO
                           '//CHIAMO IL CONTROLLO DEL PROGETTO
                           '//..........................................................//
                           '//Note  : True = siamo nel progetto originale non possiamo _
                                      importare nessun oggetto perche distruggiamo il db corrente. _
                                      False = siamo in un progetto esterno possiamo caricare tutto.
                                          
                               '//SE VIENE RESTITUITO TRUE è FUORI PROGETTO PUO CARICARE, SE FALSE EXIT PERCHE PROGETTO ORIGINALE
                               Bool1 = CheckProjectName_b_PFunct
                               
                               If Bool1 = True Then Exit Sub
                           '//..........................................................//
                
                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** FINE ***
                '//=====================================================================================================================//

    
    

    '//ATTIVO LA COLLEZIONE
    '//------------------------------------------------------------------------//
    '//Note: CHIAMO LA FUNZIONE COLLECTION FORM per ottenere la collezione di FORM

    Set formsCollection = CollectionForm_PFunct()
    '//------------------------------------------------------------------------//

    ' Loop attraverso tutte le FORM nel database corrente
    Set dbCurrent = CurrentDb

    '//CALCOLO INDICE FORM
    iCount = Application.CurrentProject.AllForms.count

    '//se ci sono FORM da cancellare
    If iCount > 0 Then

        '//reimposto la matrice sulle FORM esistenti nel db corrente
        ReDim MATRICE_FORM_Names_s(iCount - 1)

        ' Costruisci un array con i nomi dei FORM
        tempCount = 0

        '//CARICO LA MATRICE: ciclo for per caricare MATRICE_FORM_Names_s(tempCount)
        For i = 0 To iCount - 1
            Set fdf = Application.CurrentProject.AllForms(i)
            MATRICE_FORM_Names_s(tempCount) = fdf.Name
            tempCount = tempCount + 1
        Next i

        ' Cancella i FORM usando l'array dei nomi
        For i = 0 To tempCount - 1
            On Error Resume Next  ' Aggiungi gestione errori specifica per ogni iterazione
            Debug.Print "Cancellazione FORM: " & MATRICE_FORM_Names_s(i)

            '//CONTROLLO MATRICE = COLLECTION
            Bool1 = InCollectionFORMes_Funct_b(formsCollection, MATRICE_FORM_Names_s(i))

            '//cancello solo se la matrice (i) esiste nella collection
            If Bool1 = True Then
                Debug.Print
                Debug.Print "cancello la FORM trovata sia nella collezione che nella matrice : " & MATRICE_FORM_Names_s(i)
                DoCmd.DeleteObject acForm, MATRICE_FORM_Names_s(i)

                If Err.number <> 0 Then
                    Debug.Print "Errore durante la cancellazione del FORM: " & MATRICE_FORM_Names_s(i) & " - " & Err.Description
                    Err.Clear  ' Pulisci l'errore per la prossima iterazione
                Else
                    deletetFormsCount_i = deletetFormsCount_i + 1
                End If

            End If '//If Bool1 = True Then

            On Error GoTo 0  ' Reimposta gestione errori
        Next i
    End If

    ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "LE FORM CANCELLATA DA QUESTO DATABASE SONO :  " & deletetFormsCount_i, vbExclamation, "MSG CHIAMATO LA FUNZIONE PATHFILE"
    End If

'//USCITA  E GESTIONE ERRORI
'//..............................................................................................................

Exit_Cmd_Delete_FORM_Click:
    Exit Sub

Err_Cmd_Delete_FORM_Click:
    Set dbCurrent = Nothing
    MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & " Procedura -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Stop
    Resume Exit_Cmd_Delete_FORM_Click

End Sub

' Funzione per verificare se un elemento esiste nella collezione
Function InCollectionFORMes_Funct_b(col As Collection, key As String) As Boolean
    Dim varItem As Variant
    On Error Resume Next
    '//ITERNO NELLA COLLECTION PASSATA COME PARAMETRO
    For Each varItem In col
        Debug.Print
        Debug.Print "CONTROLLO COLLEZIONE = MATRICE (I)"
        Debug.Print "collezione: " & varItem & " = " & "matrice key: " & key
        Debug.Print
        If varItem = key Then
            InCollectionFORMes_Funct_b = True
            Debug.Print "trovato restituisco : " & InCollectionFORMes_Funct_b
            Exit Function
        End If
    Next varItem
    InCollectionFORMes_Funct_b = False
    On Error GoTo 0
End Function




'//CANCELLA GLI OGGETTI FORM *** FINE *** DELETE_FORM
'//=========================================================================================================








'//Cmd_DELETE_FORM_HELP_Click  *** INIZIO ***
'//-----------------------------------------------------------------------------------//

'//HELP DELETE FORM
Private Sub Cmd_Delete_FORM_HELP_Click()
                 
 '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
'//--------------------------------------------------------------------------------//--------//
'//NOTE                 -> Apro il file di tipo doc, zip o pdf
'//CODICE               -> Function ApriFilePdf.01.01
'//PARAMETRI            -> par_Path_s         = PATH _
                        -> par_NameFile_s               = NOME FILE _
                        -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                        
  '//IMPOSTO LE VARIABILI PATH E FILE
  '//
  Dim MyFile_s As String
  '//
  Dim MyPath_s  As String
  
        '//imposto i parametri - LA PATH
        '
        MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\OBJECT\HELP\"
        
        '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
        MyFile_s = "HELP_Cmd_Delete_FORM.pdf "
            
            
                     
         '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
            Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto IMPORTA OGGETTI incluso in questa form _
            anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'importazione e cancellazione in caso di resettamento _
            totale.
            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)

             
        
'//--------------------------------------------------------------------------------//--------//

End Sub


'//Cmd_DELETE_FORM_HELP_Click  *** FINE ***
'//-----------------------------------------------------------------------------------//






'//=====================================================================================================================//
'//                     DELETE OGGETTI FORM DEL DB ESTERNO IN TABELLA *** FINE ***
'//                     + DELETE DELETE HELP
'//=====================================================================================================================//








'//=========================================================================================//
'//
'//                 GESTIONE DELLE TABELLE ESTERNE O INTERNE   *** INIZIO ***
'//@GESTIONE@DELLE@TABELLE                    @GESTONE@TABELLE
'//=========================================================================================//
'//FAQ: _
        @HELP@TABELLE    @HELP@IMPORTA@TABELLE  @HELP@CANCELLA@TABELLE _
             @LA@GESTIONE@DELLE@TABELLE _
        @GESTIONE@TABELLE  @IMPORTA@TABELLE @CANCELLA@TABELLE


'//Cmd_Importa_FORM_HELP_Click()   ***INIZIO***
'//------------------------------------------------------------------------------------------------------//
'//Tipo           : Routine pubblica.
'//Attività       : Esegue codice help - ISTRUZIONI COMANDO IMPORTA FORM -
'//Note           : apre il file pdf HELP
'//Parametro      : NO PARAMETRI
'//Codice         : Cmd_Importa_FORM_HELP_Click.01
'//

Public Sub Cmd_Importa_FORM_HELP_Click()

'//MessaggiDiErrore
Dim ProceduraMessaggioErrore_s As String
Dim ProceduraAttivaEseguita_s As String
 
On Error GoTo Err_Cmd_Importa_FORM_HELP_Click


        
        '//Imposto i parametri
        ProceduraAttivaEseguita_s = "Cmd_Importa_FORM_HELP_Click "
        ProceduraMessaggioErrore_s = "Errore nella procedura HELP IMPORTA FORM ??"
        
        '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
        '//--------------------------------------------------------------------------------//--------//
        '//NOTE                 -> Apro il file di tipo doc, zip o pdf
        '//CODICE               -> Function ApriFilePdf.01.01
        '//PARAMETRI            -> par_Path_s         = PATH _
                               -> par_NameFile_s               = NOME FILE _
                               -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                               
         '//IMPOSTO LE VARIABILI PATH E FILE
         '//
         Dim MyFile_s As String
         '//
         Dim MyPath_s  As String
         
               
               '//TODO: help dei vari button importa o cancella si trova nella path : c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\OBJECT\HELP\ _
                        del progetto MSYS_OBJECT_DB_EST che dovra essere integrato completamente nel progetto MSYS che si trova nell'omonima cartella. _
                        Quando tutti gli oggetti saranno spostati nel progetto MSYS_N01_OGGETTI.mdb allora tutto l'Help dovra essere spostato nella _
                        cartella c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\OBJECT\HELP\HELP_FORM\. Per ora rimane qui.  ??? DA FINIRE ???
               '//imposto i parametri - LA PATH
               MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\OBJECT\HELP\"
               
               '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
               MyFile_s = "HELP_Cmd_Importa_FORM.pdf "
                   
                   
                '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
                   Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto IMPORTA OGGETTI incluso in questa form _
                   anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'importazione e cancellazione in caso di resettamento _
                   totale.
                Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
               
                    
               
        '//--------------------------------------------------------------------------------//--------//








'//USCITA  E GESTIONE ERRORI
'//..............................................................................................................


Exit_Cmd_Importa_FORM_HELP_Click:
    Exit Sub

Err_Cmd_Importa_FORM_HELP_Click:
 '//-------------------------------------------------------------------------------
    MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & " Procedura -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Stop
    Resume Exit_Cmd_Importa_FORM_HELP_Click
'//-------------------------------------------------------------------------------
        
End Sub

'//*** FINE ***
'//Cmd_Importa_FORM_HELP_Click()   ***INIZIO***
'//------------------------------------------------------------------------------------------------------//




'//**************************************************************************************************************************************//
'//*
'//*
'//*                                    GESTIONE DELLE MACRO   *** INIZIO ***
'//*
'//* FAQ: _
        @GESTIONE@DELLE@MACRO                    @GESTONE@MACRO _
        @LA@GESTIONE@DELLE@MACRO _
        @GESTIONE@MACRO  @IMPORTA@MACRO @CANCELLA@MACRO

'//**************************************************************************************************************************************//



'//=====================================================================================================================//
'//                            IMPORTO LE MACRO *** INIZIO ***
'//
'//=====================================================================================================================//



'//IMPORTA GLI OGGETTI MACRO DEL DB ESTERNO IN TABELLA *** INIZIO ***
'//----------------------------------------------------------------------------------------//
'//NOTE:   dal database di origine importa solo gli oggetti MACRO controllando prima se _
            nel db corrente di destinazione già esiste la MACRO per effettuare la cancellazione _
            prima dell'importazione.


'//PARAMETRI_DA_CAMBIARE _
        ----> Cmd_Importa_MACRO_Click _
        ----> Dim qdf As DAO.MacroDef _
        ----> Dim strSourceDbPath As String _
        ----> Dim objectMacros As Collection _
        ----> Dim objectExistsInCollectionMacros_Funct_b As Boolean _
        ----> MACRO _
        ----> MACRO _
        ----> COLLECTION MACRO _
        ----> importedMACROesCount_i _
        ----> deleteMACROCount_i _
        ----> deleteNameMACRO_s _


Private Sub Cmd_Importa_MACRO_Click()

    Dim dbCurrent As DAO.Database
    Dim dbSource As DAO.Database
    Dim dbs As Object
    Dim obj As AccessObject

    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim strSQL As String
    Dim strSourceDbPath As String
    Dim objectMacros As Collection
    Dim varItem As Variant
    Dim objectExistsInCollectionMacros_Funct_b As Boolean
    Dim importedMACROesCount_i As Integer
    Dim deleteMACROCount_i As Integer
    Dim deleteNameMACRO_s As String
    Dim importedNameMACRO_s As String
    Dim myNameMACRO_s As String
    Dim Bool1 As Boolean
    Dim MyValue As Integer

    On Error GoTo ErrorHandler

    ' Inizializza le variabili
    importedMACROesCount_i = 0
    deleteMACROCount_i = 0
    deleteNameMACRO_s = ""
    importedNameMACRO_s = ""
    
    
    
    
    
    
                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** INIZIO ***
                '//=====================================================================================================================//
                '//NOTE: chiamo la funzione che controlla se il progetto corrente è il progetto originale, infatti _
                         non si possono caricare gli oggetti nel PROGETTO MSYS_N01_OGGETTI.mdb perhè è originale _
                         e vi è una distruzione degli stessi. LA FUNZIONE CONTROLLO SE E' UN PROGETTO ESTERNO AMMISSIBILE _
                         per caricare gli oggetti table, query , form ecc. da importare. _
                         01) CONTROLLO = MSG DI AVVISO CON POSSIBILITA DI ANNULLARE _
                         02) CONTROLLO = CONTROLLO SE STIAMO NEL PROGETTO ORIGINALE  E QUINDI ESCLUSIONE.
                                 
                           '//01) PRIMO CONTROLLO _
                             MESSAGGIO DI IMPORTAZIONE 0 = ANNULLA  1 = CONTINUA
                           '//..........................................................//
                             '//Attenzione L'AVVISO produce il messaggio solo se il _
                                parametro è TRUE altrimenti il messaggio viene escluso.
                             MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)
                             
                             If MyValue = 0 Then Exit Sub
                           '//..........................................................//
                                 
                           '//02) SECONDO CONTROLLO
                           '//CHIAMO IL CONTROLLO DEL PROGETTO
                           '//..........................................................//
                           '//Note  : True = siamo nel progetto originale non possiamo _
                                      importare nessun oggetto perche distruggiamo il db corrente. _
                                      False = siamo in un progetto esterno possiamo caricare tutto.
                                          
                               '//SE VIENE RESTITUITO TRUE è FUORI PROGETTO PUO CARICARE, SE FALSE EXIT PERCHE PROGETTO ORIGINALE
                               Bool1 = CheckProjectName_b_PFunct
                               
                               If Bool1 = True Then Exit Sub
                           '//..........................................................//
                
                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** FINE ***
                '//=====================================================================================================================//

    

    '// ATTIVA LA COLLEZIONE DI MACRO
    '//------------------------------------------------------------------------//

        Set objectMacros = CollectionMacro_PFunct()
        
        
            '//2024.08.03_CollectionMacro_PFunct = SE LA COLLECTION E' VUOTA ESCI DALLA ROUTINE _
                        faccio un controllo sul numero degli oggetti della collection se = 0 allora _
                        non ci sono oggetti da caricare per cui si esce dalla routine.
                        
            If objectMacros.count = 0 Then Exit Sub
   
        
    '//------------------------------------------------------------------------//


    ' STEP 01: RECUPERO PATH E FILE DI ORIGINE
    strSourceDbPath = PathFile_s_pFunct()
    If strSourceDbPath = "" Then
        MsgBox "ATTENZIONE: Path e nome del file di origine non specificati.", vbCritical, "Errore di importazione"
        Exit Sub
    End If

    ' STEP 02: APERTURA DEI DATABASE
    Set dbCurrent = CurrentDb
    Set dbSource = DBEngine.Workspaces(0).OpenDatabase(strSourceDbPath)

    ' STEP 03: RECUPERO DELLE MACRO DAL DATABASE DI ORIGINE
    strSQL = "SELECT Name FROM MSysObjects WHERE Type=-32766;" ' Tipo -32764 corrisponde alle macro
    Set qdf = dbSource.CreateQueryDef("", strSQL)
    Set rs = qdf.OpenRecordset()

    ' Loop attraverso le macro trovate nel database di origine
    Do While Not rs.EOF
        myNameMACRO_s = rs.Fields("Name").Value
        rs.MoveNext

        ' Controlla se la macro è nella collezione di macro
        objectExistsInCollectionMacros_Funct_b = InCollectionMacros_Funct_b(objectMacros, myNameMACRO_s)

        ' Se la macro esiste nella collezione, procedi con l'importazione
        If objectExistsInCollectionMacros_Funct_b Then
            ' Prova a eliminare la macro esistente nel database corrente
            On Error Resume Next
            
            DoCmd.DeleteObject acMacro, myNameMACRO_s
            
            deleteMACROCount_i = deleteMACROCount_i + 1
            deleteNameMACRO_s = deleteNameMACRO_s & myNameMACRO_s & vbCrLf
            On Error GoTo 0

            ' Importa la macro dal database di origine al database corrente
            DoCmd.TransferDatabase acImport, "Microsoft Access", strSourceDbPath, acMacro, myNameMACRO_s, myNameMACRO_s
            importedMACROesCount_i = importedMACROesCount_i + 1
            importedNameMACRO_s = importedNameMACRO_s & myNameMACRO_s & vbCrLf
        End If
    Loop

    ' Messaggio di riepilogo
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "Importazione completata:" & vbCrLf & _
               "Macro importate: " & importedMACROesCount_i & vbCrLf & _
               "Macro eliminate: " & deleteMACROCount_i & vbCrLf & _
               "Elenco macro eliminate: " & vbCrLf & deleteNameMACRO_s, vbInformation
    End If
    
    
                '//ciclo nella collezione per il controllo DELLE MACRO per SELEZIONARE LA PRIMA
                '//--------------------------------------------------------------------------------------//
                     
                     '//itero nelle macro per la selezione dell'ultima
                    Set dbs = Application.CurrentProject
            
                     For Each obj In dbs.AllMacros
                     '//vado all'ultima query
                        DoCmd.SelectObject acMacro, obj.Name, True
                    
                     Next obj
                     
                     
            

                '//--------------------------------------------------------------------------------------//


    ' Libera le risorse
    rs.Close
    Set rs = Nothing
    Set qdf = Nothing
    Set dbSource = Nothing
    Set dbCurrent = Nothing

    Exit Sub

ErrorHandler:
    MsgBox "Errore durante l'importazione: " & Err.Description, vbCritical
    Debug.Print "Errore durante l'importazione: " & Err.Description
    If Not rs Is Nothing Then rs.Close
    If Not qdf Is Nothing Then Set qdf = Nothing
    If Not dbSource Is Nothing Then Set dbSource = Nothing
    If Not dbCurrent Is Nothing Then Set dbCurrent = Nothing
End Sub

'//funzione di controllo Macro in collection
Function InCollectionMacros_Funct_b(col As Collection, key As String) As Boolean
    Dim varItem As Variant
    On Error Resume Next
    For Each varItem In col
        If varItem = key Then
            InCollectionMacros_Funct_b = True
            Exit Function
        End If
    Next varItem
    InCollectionMacros_Funct_b = False
    On Error GoTo 0
End Function

'//IMPORTA GLI OGGETTI MACRO DEL DB ESTERNO IN TABELLA *** FINE ***
'//----------------------------------------------------------------------------------------//



'//=====================================================================================================================//
'//                            IMPORTO LE MACRO *** FINE ***
'//
'//=====================================================================================================================//





'//=====================================================================================================================//
'//                     DELETE OGGETTI MACRO DEL DB ESTERNO IN TABELLA *** INIZIO ***
'//                     + DELETE MACRO HELP
'//=====================================================================================================================//
'//----------------------------------------------------------------------------------------//
'//NOTE:   dal database di origine importa solo gli oggetti MACRO controllando prima se _
            nel db corrente di destinazione già esiste la MACRO per effettuare la cancellazione _
            prima dell'importazione.
'//NOTA: visto che la cancellazione riguarda il database corrente in cui è stata importata _
        la classe di IMPORTA PROGETTO, la routine cancella solo se le MACRO esistono _
        nel database corrente mediante il confronto tra : _
        LA COLLECTION MACRO = GLI OGGETTI MACRO che sono attivi nel database ESTERNO. _
        Prima di cancellare viene effettuato il confronto di esistenza _
        @cancellazione@macro@caricati_(cancellazione delle MACRO già caricate nel db esterno) _
        @cancella@MACRO@esterne_(attenzione puoi cancellare le MACRO nel @db@esterno)

'//FAQ : _
        DOVE SI CANCELLANO LE MACRO PRECARICATE _
        COME CANCELLARE LE MACRO ESTERNE _
        LE MACRO ESTERNE DA CANCELLARE
 
'// PARAMETRI DA CAMBIARE : _
    -----> MACRO _
    -----> DELLE@MACRO _
    -----> DELETE_MACRO _
    -----> CollectionMacro_PFunct _
    -----> macrosCollection _
    -----> objectMacro _
    -----> MATRICE_MACRO_Names_s _
    -----> macrosCount_i _
    -----> MacroDef _
    -----> mdf _


Private Sub Cmd_Delete_MACRO_Click()

    Dim dbCurrent As DAO.Database
    Dim dbs As Object

    Dim obj As AccessObject
    
    Dim macrosCollection As Collection
    Dim varItem As Variant
    Dim MATRICE_MACRO_Names_s() As String             'matrice di MACRO
    Dim tempCount As Integer
    Dim macrosCount_i As Integer
    Dim od As AccessObject
    Dim iCount As Integer
    Dim i As Integer
    Dim Bool1 As Boolean
    
    On Error GoTo Err_Cmd_Delete_MACRO_Click

    '//RESET
    iCount = 0
    deletetMacrosCount_i = 0
    deletedMacros_s = ""

    
    
    
    
                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** INIZIO ***
                '//=====================================================================================================================//
                '//NOTE: chiamo la funzione che controlla se il progetto corrente è il progetto originale, infatti _
                         non si possono caricare gli oggetti nel PROGETTO MSYS_N01_OGGETTI.mdb perhè è originale _
                         e vi è una distruzione degli stessi. LA FUNZIONE CONTROLLO SE E' UN PROGETTO ESTERNO AMMISSIBILE _
                         per caricare gli oggetti table, query , form ecc. da importare. _
                         01) CONTROLLO = MSG DI AVVISO CON POSSIBILITA DI ANNULLARE _
                         02) CONTROLLO = CONTROLLO SE STIAMO NEL PROGETTO ORIGINALE  E QUINDI ESCLUSIONE.
                                 
                           '//01) PRIMO CONTROLLO _
                             MESSAGGIO DI IMPORTAZIONE 0 = ANNULLA  1 = CONTINUA
                           '//..........................................................//
                             '//Attenzione L'AVVISO produce il messaggio solo se il _
                                parametro è TRUE altrimenti il messaggio viene escluso.
                             MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)
                             
                             If MyValue = 0 Then Exit Sub
                           '//..........................................................//
                                 
                           '//02) SECONDO CONTROLLO
                           '//CHIAMO IL CONTROLLO DEL PROGETTO
                           '//..........................................................//
                           '//Note  : True = siamo nel progetto originale non possiamo _
                                      importare nessun oggetto perche distruggiamo il db corrente. _
                                      False = siamo in un progetto esterno possiamo caricare tutto.
                                          
                               '//SE VIENE RESTITUITO TRUE è FUORI PROGETTO PUO CARICARE, SE FALSE EXIT PERCHE PROGETTO ORIGINALE
                               Bool1 = CheckProjectName_b_PFunct
                               
                               If Bool1 = True Then Exit Sub
                           '//..........................................................//
                
                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** FINE ***
                '//=====================================================================================================================//
    
    
    '//ATTIVO LA COLLEZIONE
    '//------------------------------------------------------------------------//
    '//Note: CHIAMO LA FUNZIONE COLLECTION MACRO per ottenere la collezione di MACRO

    Set macrosCollection = CollectionMacro_PFunct()
    '//------------------------------------------------------------------------//

    ' Loop attraverso tutte le MACRO nel database corrente
    Set dbCurrent = CurrentDb

    '//CALCOLO INDICE MACRO
    iCount = CurrentProject.AllMacros.count
    

    '//se ci sono MACRO da cancellare
    If iCount > 0 Then

        '//reimposto la matrice sulle MACRO esistenti nel db corrente
        ReDim MATRICE_MACRO_Names_s(iCount - 1)

        ' Costruisci un array con i nomi delle MACRO
        tempCount = 0

            '//CARICO LA MATRICE: ciclo for per caricare MATRICE_MACRO_Names_s(tempCount)
            Set dbs = Application.CurrentProject
            ' Ricerca oggetti AccessObject aperti in insieme AllMacros.
            For Each obj In dbs.AllMacros
                    ' Stampa nome di obj.
                    Debug.Print obj.Name
                    MATRICE_MACRO_Names_s(tempCount) = obj.Name
                    tempCount = tempCount + 1
            Next obj
    
    
        

        ' Cancella le MACRO usando l'array dei nomi
        For i = 0 To tempCount - 1
            On Error Resume Next  ' Aggiungi gestione errori specifica per ogni iterazione
            Debug.Print "Cancellazione MACRO: " & MATRICE_MACRO_Names_s(i)

            '//CONTROLLO MATRICE = COLLECTION
            Bool1 = InCollectionMacrosDELETE_Funct_b(macrosCollection, MATRICE_MACRO_Names_s(i))

            '//cancello solo se la matrice (i) esiste nella collection
            If Bool1 = True Then
                Debug.Print
                Debug.Print "cancello la MACRO trovata sia nella collezione che nella matrice : " & MATRICE_MACRO_Names_s(i)
                DoCmd.DeleteObject acMacro, MATRICE_MACRO_Names_s(i)

                If Err.number <> 0 Then
                    Debug.Print "Errore durante la cancellazione della MACRO: " & MATRICE_MACRO_Names_s(i) & " - " & Err.Description
                    Err.Clear  ' Pulisci l'errore per la prossima iterazione
                Else
                    deletetMacrosCount_i = deletetMacrosCount_i + 1
                End If

            End If '//If Bool1 = True Then

            On Error GoTo 0  ' Reimposta gestione errori
        Next i
    End If

    ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "Le MACRO CANCELLATE DA QUESTO DATABASE SONO :  " & deletetMacrosCount_i, vbExclamation, "MSG CHIAMATO LA FUNZIONE PATHFILE"
    End If

'//USCITA E GESTIONE ERRORI
'//..............................................................................................................

Exit_Cmd_Delete_MACRO_Click:
    Exit Sub

Err_Cmd_Delete_MACRO_Click:
    Set dbCurrent = Nothing
    MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & " Procedura -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Stop
    Resume Exit_Cmd_Delete_MACRO_Click

End Sub

' Funzione per verificare se un elemento esiste nella collezione di MACRO
Function InCollectionMacrosDELETE_Funct_b(col As Collection, key As String) As Boolean
    Dim varItem As Variant
    On Error Resume Next
    '//ITERNO NELLA COLLECTION PASSATA COME PARAMETRO
    For Each varItem In col
        Debug.Print
        Debug.Print "CONTROLLO COLLEZIONE = MATRICE (I)"
        Debug.Print "collezione: " & varItem & " = " & "matrice key: " & key
        Debug.Print
        If varItem = key Then
            InCollectionMacrosDELETE_Funct_b = True
            Debug.Print "trovato restituisco : " & InCollectionMacrosDELETE_Funct_b
            Exit Function
        End If
    Next varItem
    InCollectionMacrosDELETE_Funct_b = False
    On Error GoTo 0
End Function



Private Sub Cmd_Delete_MACRO_HELP_Click()

    '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
    '//--------------------------------------------------------------------------------//--------//
    '//NOTE                 -> Apro il file di tipo doc, zip o pdf
    '//CODICE               -> Function ApriFilePdf.01.01
    '//PARAMETRI            -> par_Path_s         = PATH _
                            -> par_NameFile_s               = NOME FILE _
                            -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                            
      '//IMPOSTO LE VARIABILI PATH E FILE
      '//
      Dim MyFile_s As String
      '//
      Dim MyPath_s  As String
      
            '//imposto i parametri - LA PATH
            '
            MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\OBJECT\HELP\"
            '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
            '
            MyFile_s = "HELP_Cmd_Delete_MACRO.pdf "
                
                
                       
         '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
            Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto DELETE OGGETTI incluso in questa MODULI _
            anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'DELETEzione e cancellazione in caso di resettamento _
            totale.
            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
      
            
    '//--------------------------------------------------------------------------------//--------//


End Sub


'//=====================================================================================================================//
'//                     DELETE OGGETTI MACRO DEL DB ESTERNO IN TABELLA *** FINE ***
'//                     + DELETE MACRO HELP
'//=====================================================================================================================//




'//**************************************************************************************************************************************//
'//*
'//*
'//*                                    GESTIONE DELLE MACRO   *** INIZIO ***
'//*
'//*
'//**************************************************************************************************************************************//







'//IMPORTA MODULI HELP *** INIZIO ***
'//-----------------------------------------------------------------------------------//

Private Sub Cmd_IMPORTA_MODULI_HELP_Click()

    '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
    '//--------------------------------------------------------------------------------//--------//
    '//NOTE                 -> Apro il file di tipo doc, zip o pdf
    '//CODICE               -> Function ApriFilePdf.01.01
    '//PARAMETRI            -> par_Path_s         = PATH _
                            -> par_NameFile_s               = NOME FILE _
                            -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                            
      '//IMPOSTO LE VARIABILI PATH E FILE
      '//
      Dim MyFile_s As String
      '//
      Dim MyPath_s  As String
      
            '//imposto i parametri - LA PATH
            '
            MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\OBJECT\HELP\"
            '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
            '
            MyFile_s = "HELP_Cmd_Importa_MODULI.pdf "
                
                
                       
         '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
            Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto DELETE OGGETTI incluso in questa MODULI _
            anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'DELETEzione e cancellazione in caso di resettamento _
            totale.
            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
      
            
    '//--------------------------------------------------------------------------------//--------//

End Sub

'//IMPORTA MODULI HELP *** FINE ***
'//-----------------------------------------------------------------------------------//


 
 






'//=========================================================================================//
'//
'//                 GESTIONE DELLE REPORT ESTERNE O INTERNE   *** INIZIO ***
'//@GESTIONE@DELLE@REPORT                   @GESTONE@REPORT
'//=========================================================================================//

 

 
 
'//==========================================================================================================//
'// IMPORTA GLI OGGETTI REPORT *** INIZIO *** '@IMPORT@REPORT
'//==========================================================================================================//

Private Sub Cmd_Importa_REPORT_Click()
    Dim obj As AccessObject
    Dim dbs As Object
    Dim ProceduraMessaggioErrore_s As String
    Dim reportsCollection As Collection
    Dim reportExists As Boolean
    Dim sourceDBPath As String
    Dim externalDB As DAO.Database
    Dim rs As DAO.Recordset
    
    On Error GoTo Err_Cmd_Import_REPORT_Click
   
    ' Reset counts and strings
    deletetReportesCount_i = 0
    importedReportesCount_i = 0
    importedReports_s = ""
    
    
    
    
    
            '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** INIZIO ***
                '//=====================================================================================================================//
                '//NOTE: chiamo la funzione che controlla se il progetto corrente è il progetto originale, infatti _
                         non si possono caricare gli oggetti nel PROGETTO MSYS_N01_OGGETTI.mdb perhè è originale _
                         e vi è una distruzione degli stessi. LA FUNZIONE CONTROLLO SE E' UN PROGETTO ESTERNO AMMISSIBILE _
                         per caricare gli oggetti table, query , form ecc. da importare. _
                         01) CONTROLLO = MSG DI AVVISO CON POSSIBILITA DI ANNULLARE _
                         02) CONTROLLO = CONTROLLO SE STIAMO NEL PROGETTO ORIGINALE  E QUINDI ESCLUSIONE.
                                 
                           '//01) PRIMO CONTROLLO _
                             MESSAGGIO DI IMPORTAZIONE 0 = ANNULLA  1 = CONTINUA
                           '//..........................................................//
                             '//Attenzione L'AVVISO produce il messaggio solo se il _
                                parametro è TRUE altrimenti il messaggio viene escluso.
                             MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)
                             
                             If MyValue = 0 Then Exit Sub
                           '//..........................................................//
                                 
                           '//02) SECONDO CONTROLLO
                           '//CHIAMO IL CONTROLLO DEL PROGETTO
                           '//..........................................................//
                           '//Note  : True = siamo nel progetto originale non possiamo _
                                      importare nessun oggetto perche distruggiamo il db corrente. _
                                      False = siamo in un progetto esterno possiamo caricare tutto.
                                          
                               '//SE VIENE RESTITUITO TRUE è FUORI PROGETTO PUO CARICARE, SE FALSE EXIT PERCHE PROGETTO ORIGINALE
                               Bool1 = CheckProjectName_b_PFunct
                               
                               If Bool1 = True Then Exit Sub
                           '//..........................................................//
                
                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** FINE ***
                '//=====================================================================================================================//

    
    
    
    
    
    
    
    
    '// ATTIVA LA COLLEZIONE DEI REPORT
    '//------------------------------------------------------------------------//

        ' Recupero la collezione di report da importare
        Set reportsCollection = CollectionReport_PFunct()
        
    
        
        
            '//2024.08.03_CollectionReport_PFunct  = SE LA COLLECTION E' VUOTA ESCI DALLA ROUTINE _
                        faccio un controllo sul numero degli oggetti della collection se = 0 allora _
                        non ci sono oggetti da caricare per cui si esce dalla routine.
                        
            If reportsCollection.count = 0 Then GoTo Exit_Cmd_Import_REPORT_Click
   
        
    '//------------------------------------------------------------------------//
    
    
    ' Recupero il percorso e il nome del file del database esterno
    sourceDBPath = PathFile_s_pFunct
    
    If sourceDBPath = "" Then
        MsgBox "ATTENZIONE ERRORE: il percorso del file del database esterno non è stato specificato. Uscita dalla routine.", vbCritical, "Errore: Percorso del File Non Specificato"
        Exit Sub
    End If
    
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "IMPORTO I REPORT DA QUESTO DATABASE; Percorso del file: " & sourceDBPath, vbExclamation, "Informazione: Chiamata Funzione PathFile"
    End If
    
    
    ' Apro il database esterno
    Set externalDB = DBEngine.Workspaces(0).OpenDatabase(sourceDBPath)
    
    ' Ciclo nella collezione per il controllo dei report precaricati
    For Each reportName In reportsCollection
        Debug.Print "Report nella collezione: " & myreportName_s
        
        ' Verifica se il report esiste nel database corrente
        reportExists = False
        Set dbs = Application.CurrentProject
        For Each obj In dbs.AllReports
            If obj.Name = reportName Then
                reportExists = True
                Exit For
            End If
        Next obj
        
        If reportExists Then
            ' Se il report esiste nel database corrente, cancellalo
            DoCmd.DeleteObject acReport, reportName
            deletetReportesCount_i = deletetReportesCount_i + 1
            Debug.Print "Report eliminato: " & reportName
        End If
        
        '//CONTROLLO SULLA TABELLA DI SISTEMA DEL DB ESTERNO MSysObjects
        '//----------------------------------------------------------------------------------------------------------------------//
        ' Verifica se il modulo esiste nel database esterno
        ' La riga SELECT Name FROM MSysObjects WHERE Type=-32761 è una query SQL utilizzata per selezionare i
        ' nomi degli oggetti di tipo modulo nel database di Access. Ecco una spiegazione dettagliata:

        ' MSysObjects: è una tabella di sistema in Microsoft Access che contiene informazioni su tutti gli oggetti nel database, _
          come tabelle, query, moduli, report, ecc. Type=-32761: questo filtro nella clausola WHERE specifica il tipo di oggetto _
          da selezionare. In Access, ogni tipo di oggetto ha un valore numerico associato nella tabella MSysObjects. _
          Il valore -32761 è il codice che rappresenta i moduli (VBA Modules).
         'Quindi, la query SELECT Name FROM MSysObjects WHERE Type=-32761 seleziona tutti i nomi degli oggetti nella tabella _
         di sistema MSysObjects che sono identificati come moduli VBA.
         ' Verifica se il report esiste nel database esterno
        ' Utilizzo una query SQL per verificare l'esistenza del report nel database esterno
        ' il codice del report è : -32764 controllato con la funzione codici esterni.
        

        reportExists = False
        Set rs = externalDB.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type= -32764 AND Name='" & reportName & "'")
            If Not rs.EOF Then
                reportExists = True
            End If
                rs.Close
            
            If reportExists Then
                ' Se il report esiste nel database esterno, importalo
                DoCmd.TransferDatabase acImport, "Microsoft Access", sourceDBPath, acReport, reportName, reportName
                importedReportesCount_i = importedReportesCount_i + 1
                importedReports_s = importedReports_s & reportName & vbCrLf
                Debug.Print "Report importato: " & reportName
            Else
                ' Report non trovato nel database esterno
                Debug.Print "Report non trovato nel database esterno: " & reportName
            End If
        Next reportName
        
            ' Chiudo il database esterno
            externalDB.Close
        '//----------------------------------------------------------------------------------------------------------------------//
        
    ' Messaggio di riepilogo finale
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "CONTROLLO REPORT PRECARICATI DA IMPORTARE:" & vbCrLf & _
               "Report importati: " & importedReportesCount_i & vbCrLf & importedReports_s & _
               "Report esistenti cancellati: " & deletetReportesCount_i, vbInformation
    End If
    
    
                '//ciclo nella collezione per il controllo delle QUERY precaricate per SELEZIONARE LA PRIMA
                '//--------------------------------------------------------------------------------------//
                     For Each obj In dbs.AllReports
                            '//vado all'ultima query
                        DoCmd.SelectObject acReport, obj.Name, True
                    
                     Next obj
                '//--------------------------------------------------------------------------------------//


Exit_Cmd_Import_REPORT_Click:
    Exit Sub

Err_Cmd_Import_REPORT_Click:
    MsgBox Err.Description & " - Errore: " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Stop
    Resume Exit_Cmd_Import_REPORT_Click

End Sub


'//FUNZIONE DI CONTROLLO DEGLI OGGETTI REPORT DEL DB MSYSDBEST *** INIZIO ***
'//----------------------------------------------------------------------------------------//

'//FUNZIONE DI CONTROLLO SE ESISTE L'OGGETTO = todo: modificata
Function ObjectReportExists(db As DAO.Database, objType As Integer, objName As String) As Boolean
    'Dim obj As DAO.Document
    'On Error Resume Next
    'Set obj = db.Containers("Reports").Documents(objName)
    'ObjectReportExists = (Err.Number = 0)
    'Set obj = Nothing
    'On Error GoTo 0
    
    Dim obj As AccessObject
    On Error Resume Next
    For Each obj In CurrentProject.AllReports
        Debug.Print "report corrente controllato :" & obj.Name
        If obj.Name = objName Then
            ReportExistsInCurrentDb = True
            Exit Function
        End If
    Next obj
    ReportExistsInCurrentDb = False
    On Error GoTo 0
    
    
    
End Function


'//FUNZIONE DI CONTROLLO DEGLI OGGETTI REPORT DEL DB MSYSDBEST *** FINE ***
'//----------------------------------------------------------------------------------------//

'//==========================================================================================================//
'// IMPORTA GLI OGGETTI REPORT *** FINE *** '@IMPORT@REPORT
'//==========================================================================================================//




'//=====================================================================================================================//
'//                     DELETE OGGETTI REPROT DEL DB ESTERNO IN TABELLA *** INIZIO ***
'//                     + DELETE REPORT HELP
'//=====================================================================================================================//



'//CANCELLA GLI OGGETTI REPORT *** INIZIO *** DELETE_REPORT
'//======================================================================================================//
'//@GESTIONE@DEI@REPORT    (@MODELLO@DELETE@OGGETTI)

'//NOTA: visto che la cancellazione riguarda il database corrente in cui è stata importata _
        la classe di IMPORTA PROGETTO, la routine cancella solo se i REPORT esistono _
        nel database corrente mediante il confronto tra : _
        LA COLLECTION REPORT = GLI OGGETTI REPORT che sono attivi nel database ESTERNO. _
        Prima di cancellare viene effettuato il confronto di esistenza _
        @cancellazione@report@caricati_(cancellazione dei REPORT già caricati nel db esterno) _
        @cancella@TABLE@esterne_(attenzione puoi cancellare i REPORT nel @db@esterno)

'//FAQ : _
        DOVE SI CANCELLANO I REPORT PRECARICATI _
        COME CANCELLARE I REPORT ESTERNI _
        I REPORT ESTERNI DA CANCELLARE

'// PARAMETRI DA CAMBIARE : _
    -----> REPORT _
    -----> DEI@REPORT _
    -----> DELETE_REPORT _
    -----> CollectionReport_PFunct _
    -----> reportsCollection _
    -----> objectReport _
    -----> MATRICE_REPORT_Names_s _
    -----> reportsCount_i _
    -----> ReportDef _
    -----> rpt _

Private Sub Cmd_Delete_REPORT_Click()

    Dim dbCurrent As DAO.Database
    Dim rpt As Report
    Dim reportsCollection As Collection
    Dim varItem As Variant
    Dim MATRICE_REPORT_Names_s() As String             'matrice di REPORT
    Dim tempCount As Integer
    Dim reportsCount_i As Integer
    Dim rdf As AccessObject
    Dim iCount As Integer
    Dim i As Integer
    Dim Bool1 As Boolean
    
    On Error GoTo Err_Cmd_Delete_REPORT_Click

    '//RESET
    iCount = 0
    deletetReportsCount_i = 0
    deletedReports_s = ""
    
    
    
    
    
                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** INIZIO ***
                '//=====================================================================================================================//
                '//NOTE: chiamo la funzione che controlla se il progetto corrente è il progetto originale, infatti _
                         non si possono caricare gli oggetti nel PROGETTO MSYS_N01_OGGETTI.mdb perhè è originale _
                         e vi è una distruzione degli stessi. LA FUNZIONE CONTROLLO SE E' UN PROGETTO ESTERNO AMMISSIBILE _
                         per caricare gli oggetti table, query , form ecc. da importare. _
                         01) CONTROLLO = MSG DI AVVISO CON POSSIBILITA DI ANNULLARE _
                         02) CONTROLLO = CONTROLLO SE STIAMO NEL PROGETTO ORIGINALE  E QUINDI ESCLUSIONE.
                                 
                           '//01) PRIMO CONTROLLO _
                             MESSAGGIO DI IMPORTAZIONE 0 = ANNULLA  1 = CONTINUA
                           '//..........................................................//
                             '//Attenzione L'AVVISO produce il messaggio solo se il _
                                parametro è TRUE altrimenti il messaggio viene escluso.
                             MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)
                             
                             If MyValue = 0 Then Exit Sub
                           '//..........................................................//
                                 
                           '//02) SECONDO CONTROLLO
                           '//CHIAMO IL CONTROLLO DEL PROGETTO
                           '//..........................................................//
                           '//Note  : True = siamo nel progetto originale non possiamo _
                                      importare nessun oggetto perche distruggiamo il db corrente. _
                                      False = siamo in un progetto esterno possiamo caricare tutto.
                                          
                               '//SE VIENE RESTITUITO TRUE è FUORI PROGETTO PUO CARICARE, SE FALSE EXIT PERCHE PROGETTO ORIGINALE
                               Bool1 = CheckProjectName_b_PFunct
                               
                               If Bool1 = True Then Exit Sub
                           '//..........................................................//
                
                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** FINE ***
                '//=====================================================================================================================//
     

    '//ATTIVO LA COLLEZIONE
    '//------------------------------------------------------------------------//
    '//Note: CHIAMO LA FUNZIONE COLLECTION REPORT per ottenere la collezione di REPORT

    Set reportsCollection = CollectionReport_PFunct()
    '//------------------------------------------------------------------------//

    ' Loop attraverso tutti i REPORT nel database corrente
    Set dbCurrent = CurrentDb

    '//CALCOLO INDICE REPORT
    iCount = Application.CurrentProject.AllReports.count

    '//se ci sono REPORT da cancellare
    If iCount > 0 Then

        '//reimposto la matrice sui REPORT esistenti nel db corrente
        ReDim MATRICE_REPORT_Names_s(iCount - 1)

        ' Costruisci un array con i nomi dei REPORT
        tempCount = 0

        '//CARICO LA MATRICE: ciclo for per caricare MATRICE_REPORT_Names_s(tempCount)
        For i = 0 To iCount - 1
            Set rdf = Application.CurrentProject.AllReports(i)
            MATRICE_REPORT_Names_s(tempCount) = rdf.Name
            tempCount = tempCount + 1
        Next i

        ' Cancella i REPORT usando l'array dei nomi
        For i = 0 To tempCount - 1
            On Error Resume Next  ' Aggiungi gestione errori specifica per ogni iterazione
            Debug.Print "Cancellazione REPORT: " & MATRICE_REPORT_Names_s(i)

            '//CONTROLLO MATRICE = COLLECTION
            Bool1 = InCollectionREPORTes_Funct_b(reportsCollection, MATRICE_REPORT_Names_s(i))

            '//cancello solo se la matrice (i) esiste nella collection
            If Bool1 = True Then
                Debug.Print
                Debug.Print "cancello il REPORT trovato sia nella collezione che nella matrice : " & MATRICE_REPORT_Names_s(i)
                DoCmd.DeleteObject acReport, MATRICE_REPORT_Names_s(i)

                If Err.number <> 0 Then
                    Debug.Print "Errore durante la cancellazione del REPORT: " & MATRICE_REPORT_Names_s(i) & " - " & Err.Description
                    Err.Clear  ' Pulisci l'errore per la prossima iterazione
                Else
                    deletetReportsCount_i = deletetReportsCount_i + 1
                End If

            End If '//If Bool1 = True Then

            On Error GoTo 0  ' Reimposta gestione errori
        Next i
    End If

    ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "I REPORT CANCELLATI DA QUESTO DATABASE SONO :  " & deletetReportsCount_i, vbExclamation, "MSG CHIAMATO LA FUNZIONE PATHFILE"
    End If

'//USCITA E GESTIONE ERRORI
'//..............................................................................................................

Exit_Cmd_Delete_REPORT_Click:
    Exit Sub

Err_Cmd_Delete_REPORT_Click:
    Set dbCurrent = Nothing
    MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & " Procedura -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Stop
    Resume Exit_Cmd_Delete_REPORT_Click

End Sub

' Funzione per verificare se un elemento esiste nella collezione
Function InCollectionREPORTes_Funct_b(col As Collection, key As String) As Boolean
    Dim varItem As Variant
    On Error Resume Next
    '//ITERNO NELLA COLLECTION PASSATA COME PARAMETRO
    For Each varItem In col
        Debug.Print
        Debug.Print "CONTROLLO COLLEZIONE = MATRICE (I)"
        Debug.Print "collezione: " & varItem & " = " & "matrice key: " & key
        Debug.Print
        If varItem = key Then
            InCollectionREPORTes_Funct_b = True
            Debug.Print "trovato restituisco : " & InCollectionREPORTes_Funct_b
            Exit Function
        End If
    Next varItem
    InCollectionREPORTes_Funct_b = False
    On Error GoTo 0
End Function


'//CANCELLA GLI OGGETTI REPORT *** FINE *** DELETE_REPORT
'//=========================================================================================================

'//DELETE REPORT HELP *** INIZIO ***
'//-----------------------------------------------------------------------------------//

Private Sub Cmd_Delete_REPORT_HELP_Click()

    '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
    '//--------------------------------------------------------------------------------//--------//
    '//NOTE                 -> Apro il file di tipo doc, zip o pdf
    '//CODICE               -> Function ApriFilePdf.01.01
    '//PARAMETRI            -> par_Path_s         = PATH _
                            -> par_NameFile_s               = NOME FILE _
                            -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                            
      '//IMPOSTO LE VARIABILI PATH E FILE
      '//
      Dim MyFile_s As String
      '//
      Dim MyPath_s  As String
      
            '//imposto i parametri - LA PATH
            '
            MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\OBJECT\HELP\"
            '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
            '
            MyFile_s = "Help_Cmd_DELETE_REPORT.pdf "
                
                
                       
         '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
            Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto DELETE OGGETTI incluso in questa REPORT _
            anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'DELETEzione e cancellazione in caso di resettamento _
            totale.
            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
      
            
    '//--------------------------------------------------------------------------------//--------//

End Sub


'//DELETE REPORT HELP *** FINE ***
'//-----------------------------------------------------------------------------------//




'//=====================================================================================================================//
'//                     DELETE OGGETTI REPROT DEL DB ESTERNO IN TABELLA *** FINE ***
'//                     + DELETE REPORT HELP
'//=====================================================================================================================//









'//=========================================================================================//
'//
'//                 GESTIONE DELLE REPORT ESTERNE O INTERNE   *** FINE ***
'//@GESTIONE@DELLE@REPORT                   @GESTONE@REPORT
'//=========================================================================================//









'//**************************************************************************************//
'//                     EVENTI FORM  @GLI@EVENTI@FORM@LOAD
'//**************************************************************************************//

'//FORM@SU@CARICAMENTO    @PARAMETRI@PRECARICATI_(i parametri del precaricati nella form corente)
Private Sub Form_Load()

    '//disabilito i butto
    Call DisabilitaButton_Sub
    Call Disabilito_sourceDBPath_Sub
    
    
    

    '//@avviso
    '//imposto il flag avviso ad abilitato per visualizzare il messaggio di proseguimento
    Me.ctr_AVVISO_ABILITATO_TXT = True
    
    '//@progetto@corrente @modificato = inserisco subito nella casella di testo ed evito _
        che il progetto corrente vada a finire nella casella combinata.
    Me.CurrentProject_txt = Me.Application.CurrentProject.Name
    
    
     
     '//@CMB_01
     '//-------------------------------------------------------------------------//
     '//@cmb@imposto_(i valore della @casella@combinata)
     '//IMPOSTO LA CASELLA COMBINATA IN PARTENZA @cmb@abilitazioni_(vari @codici dei @progetti)
        Str1 = Str1 & ";"
        Str1 = Str1 & "ELENCO PROGETTI;"
        Str1 = Str1 & "MSYS_N01_OGGETTI.mdb;"
        Str1 = Str1 & "GE_ESPROPRI.mdb;"
        Str1 = Str1 & "TREEVIEW.mdb;"
        Str1 = Str1 & ""
        
        
      
      
        Me.Cmb_01_txt.Value = "ELENCO PROGETTI"
        
        Me.Cmb_01_txt.RowSourceType = "elenco valori"
        Me.Cmb_01_txt.RowSource = Str1
    '//-------------------------------------------------------------------------//
    
    
    
    '//@CMB_02
    '//-------------------------------------------------------------------------//
    '//@cmb@imposto_(i valore della @casella@combinata)
    '//IMPOSTO LA CASELLA COMBINATA IN PARTENZA @cmb@abilitazioni_(vari @codici dei @progetti)
        Me.Cmb_02_txt.Value = ""
        Str1 = ""
        Str1 = Str1 & "ELENCO OGGETTI MSYS;"
        Str1 = Str1 & "IMPORTA_OGGETTI_MSYS_DLL;"                    '//@IMPORTA@OGGETTI@DLL
        Str1 = Str1 & "IMPORTA_OGGETTI_TREEVIEW;"                     '//@IMPORTA@OGGETTI@TREEVIEW
        Str1 = Str1 & ""
    
    
        Me.Cmb_02_txt.RowSourceType = "elenco valori"
        Me.Cmb_02_txt.RowSource = Str1
    '//-------------------------------------------------------------------------//
    
                    
                    
                    
    

    
End Sub






'//**************************************************************************************//
'//                     EVENTI FORM  *** FINE ***
'//**************************************************************************************//




'//=========================================================================================//
'//
'//                 GESTIONE DELLE MACRO ESTERNE O INTERNE   *** INIZIO ***
'//@GESTIONE@DELLE@MACRO                    @GESTONE@MACRO
'//=========================================================================================//


'//IMPORTA GLI OGGETTI MACRO DEL DB ESTERNO IN TABELLA *** INIZIO ***
'//----------------------------------------------------------------------------------------//
'//NOTE:   dal database di origine importa solo gli oggetti MACRO controllando prima se _
            nel db corrente di destinazione già esiste la MACRO per effettuare la cancellazione _
            prima dell'importazione.

'//IMPORTA GLI OGGETTI MACRO DEL DB ESTERNO IN TABELLA *** FINE ***
'//----------------------------------------------------------------------------------------//

'//=========================================================================================//
'//
'//                 GESTIONE DELLE MACRO ESTERNE O INTERNE   *** FINE ***
'//@GESTIONE@DELLE@MACRO                    @GESTONE@MACRO
'//=========================================================================================//







'//=========================================================================================//
'//
'//                 GESTIONE DELLE QUERY ESTERNE O INTERNE   *** INIZIO ***
'//@GESTIONE@DELLE@QUERY                    @GESTONE@QUERY
'//=========================================================================================//


'//IMPORTA GLI OGGETTI TABLE E QUERY DEL DB ESTERNO IN TABELLA *** INIZIO ***
'//----------------------------------------------------------------------------------------//
'//NOTE:   dal database di origine importa solo gli oggetti Table e Query controllando prima se _
            nel db corrente di destinazione già esiste la Table o Query per effettuare la cancellazione _
            prima dell'importazione.


'//PARAMETRI_DA_CAMBIARE _
        ----> Cmd_Importa_QUERY_Click _
        ----> Dim qdf As DAO.QueryDef _
        ----> Dim strSourceDbPath As String _
        ----> Dim objectQueries As Collection _
        ----> Dim objectExistsInCollectionQueries_Funct_b As Boolean _
        ----> TABELLE E QUERY _
        ----> QUERY _
        ----> Dim objectExistsInCollectionQueries_Funct_b As Boolean _
        ----> COLLECTION QUERY _
        ----> importedQueryesCount_i _
        ----> deleteQueryesCount_i _
        ----> deleteNameQueryes_s _

Private Sub Cmd_Importa_QUERY_Click()

Dim dbCurrent As DAO.Database
Dim dbSource As DAO.Database
Dim qdf As DAO.QueryDef
Dim strSourceDbPath As String
Dim strObjectName As String
Dim objectQueries As Collection
Dim varItem As Variant
Dim objectExistsInCollectionQueries_Funct_b As Boolean

On Error GoTo ErrorHandler

'//=====================================================================================================================//
'//                            STEP_01 = RECUPERO PATH E COLLECTION
'//RECUPERO LA COLLEZIONE TABELLE E QUERY DA IMPORTARE E FACCIO IL CONTROLLO
'//=====================================================================================================================//

    '//RESET
    'Inizializza i contatori e le stringhe di riepilogo
    iCount = 0
    importedQueryesCount_i = 0
    importedObjects_s = ""
    deleteQueryesCount_i = 0
    deleteNameQueryes_s = ""
              
    '//ATTIVO LA COLLEZIONE
    '//------------------------------------------------------------------------//
    '//Note: CHIAMO LA FUNZIONE COLLECTION QUERY per ottenere la collezione di TABELLE E QUERY
    Set objectQueries = CollectionQuery_PFunct()
    
            '//2024.08.03_CollectionQuery_PFunct = SE LA COLLECTION E' VUOTA ESCI DALLA ROUTINE _
                        faccio un controllo sul numero degli oggetti della collection se = 0 allora _
                        non ci sono oggetti da caricare per cui si esce dalla routine.
                        
            If objectQueries.count = 0 Then Exit Sub
                
    '//------------------------------------------------------------------------//
    
    
    

    '//I° CONTROLLO SULLA COLLECTION QUERY RECUPERATA
    '//ciclo nella collezione per il controllo delle TABELLE E QUERY precaricate
    For Each objCollection_objectName_v In objectQueries
        Debug.Print "Oggetto nella collezione: " & objCollection_objectName_v
        '//conteggio oggetti DA IMPORTARE
        iCount = iCount + 1
        importedNameQueryes_s = importedNameQueryes_s & objCollection_objectName_v & vbCrLf
    Next objCollection_objectName_v

    ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "CONTROLLO OGGETTI COLLECTION PRECARICATI DA IMPORTARE:" & vbCrLf & _
               "Oggetti COLLECTION PRECARICATA DA IMPORTARE: " & iCount & vbCrLf & importedNameQueryes_s & _
               "Oggetti esistenti cancellati: " & deleteQueryesCount_i & vbCrLf & deleteNameQueryes_s, vbInformation
    End If
    
            '//RESET
            iCount = 0
             
    

'//---------------------------------------------------------------------------//

'//RECUPERO LA PATH ED IL FILE.MDB PER LE IMPORTAZIONI
'//---------------------------------------------------------------------------//
'//NOTE: chiamo la funzione e recupero la path ed il file.mdb facendo un controllo _
        se la variabile è nulla msg di avviso ed uscita dalla routine.
    '//reset
    sourceDBPath_s = ""

    '//@RECUPERO@PATH_(ed il file
    sourceDBPath_s = PathFile_s_pFunct

    If sourceDBPath_s = "" Then
        MsgBox "ATTENZIONE ERRORE path e file sono null uscita dalla routine", vbCritical, "MSG CONTROLLO PATH E FILE VALORIZZATI"
        If MyValue = 0 Then Exit Sub
    End If

      ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "IMPORTO GLI OGGETTI DA QUESTO DATABASE; path e del file : " & sourceDBPath_s, vbExclamation, "MSG CHIAMATO LA FUNZIONE PATHFILE"
    End If

'//=====================================================================================================================//
'//                            STEP_01 = RECUPERO PATH E COLLECTION
'//RECUPERO LA COLLEZIONE FORM DA IMPORTARE E FACCIO IL CONTROLLO          **** FINE ***
'//=====================================================================================================================//

'//=====================================================================================================================//
'//                             STEP_02 = @CONTROLLO@PROGETTO
'//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** INIZIO ***
'//=====================================================================================================================//
'//NOTE: chiamo la funzione che controlla se il progetto corrente è il progetto originale, infatti _
         non si possono caricare gli oggetti nel PROGETTO MSYS_N01_OGGETTI.mdb perhè è originale _
         e vi è una distruzione degli stessi. LA FUNZIONE CONTROLLO SE E' UN PROGETTO ESTERNO AMMISSIBILE _
         per caricare gli oggetti table, query , form ecc. da importare. _
         01) CONTROLLO = MSG DI AVVISO CON POSSIBILITA DI ANNULLARE _
         02) CONTROLLO = CONTROLLO SE STIAMO NEL PROGETTO ORIGINALE  E QUINDI ESCLUSIONE.
                 
           '//01) PRIMO CONTROLLO _
             MESSAGGIO DI IMPORTAZIONE 0 = ANNULLA  1 = CONTINUA
           '//..........................................................//
             '//Attenzione L'AVVISO produce il messaggio solo se il _
                parametro è TRUE altrimenti il messaggio viene escluso.
             MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)
             
             If MyValue = 0 Then Exit Sub
           '//..........................................................//
                 
           '//02) SECONDO CONTROLLO
           '//CHIAMO IL CONTROLLO DEL PROGETTO
           '//..........................................................//
           '//Note  : True = siamo nel progetto originale non possiamo _
                      importare nessun oggetto perche distruggiamo il db corrente. _
                      False = siamo in un progetto esterno possiamo caricare tutto.
                          
               '//SE VIENE RESTITUITO TRUE è FUORI PROGETTO PUO CARICARE, SE FALSE EXIT PERCHE PROGETTO ORIGINALE
               Bool1 = CheckProjectName_b_PFunct
               
               If Bool1 = True Then Exit Sub
           '//..........................................................//

'//=====================================================================================================================//
'//                             STEP_02 = @CONTROLLO@PROGETTO
'//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** FINE ***
'//=====================================================================================================================//

'//=====================================================================================================================//
'//                            STEP_03 = IMPORTO LE TABELLE E LE QUERY
'//
'//=====================================================================================================================//
    
    '//RESET
    '//Inizializza i contatori e le stringhe di riepilogo
    
    
    importedQueriesCount_i = 0
    deleteQueryesCount_i = 0
    deleteNameQueryes_s = ""
    importedNameQueryes_s = ""
    

    ' Apri il database corrente
    Set dbCurrent = CurrentDb

    ' Apri il database di origine
    Set dbSource = DBEngine.Workspaces(0).OpenDatabase(sourceDBPath_s)
    
    'Loop attraverso tutte le QUERY nel database di origine ossia del DATABASE ESTERNO
    For Each qdf In dbSource.QueryDefs
        ' Salta le tabelle di sistema (i cui nomi iniziano con "MSys") aggiungo come parametro( la collezione table + nome tabella)
        If Left(qdf.Name, 3) <> "~sq" Or (Left(qdf.Name, 3) = "~sq" And InCollectionQueries_Funct_b(objectQueries, qdf.Name)) Then
            ' Nome della tabella nel database di origine
            myNameQuery_s = qdf.Name
            
            Debug.Print
            Debug.Print "CONTROLLO QUERY DA IMPORTARE --> : "; myNameQuery_s
            Debug.Print
            
            
            ' Controlla se la QUERY è nella collezione
            objectExistsInCollectionQueries_Funct_b = False
            
            For Each varItem In objectQueries
                If myNameQuery_s = varItem Then
                    objectExistsInCollectionQueries_Funct_b = True
                    Exit For
                End If
            Next varItem
            
            ' Se la QUERY è nella collezione, importala
            If objectExistsInCollectionQueries_Funct_b Then
                
                ' Elimina la QUERY se esiste già nel database corrente
                On Error Resume Next
                dbCurrent.QueryDefs.Delete myNameQuery_s
                
                deleteQueryesCount_i = deleteQueryesCount_i + 1
                deleteNameQueryes_s = deleteNameQueryes_s & myNameQuery_s & vbCrLf
                
                On Error GoTo 0
                
                ' Copia la tabella dal database di origine al database corrente
                DoCmd.TransferDatabase acImport, "Microsoft Access", sourceDBPath_s, acQuery, myNameQuery_s, myNameQuery_s
                
                '//conteggio TABELLE IMPORTATE
                
                importedQueryesCount_i = importedQueryesCount_i + 1
                importedNameQueryes_s = importedNameQueryes_s & myNameQuery_s & vbCrLf
                
            End If
        End If
    Next qdf
                
        '//TODO: MANCA IL CONTROLLO SULLA TABELLA DI SISTEMA DEL DB ESTERNO MSysObjects DA CREARE VEDI FUNZIONE Cmd_Importa_MODULI_Click
        '//----------------------------------------------------------------------------------------------------------------------//
        
            '//TODO: adattare il controllo sulla tabella di sistema esterna ... vai a Cmd_Importa_MODULI_Click

        '//----------------------------------------------------------------------------------------------------------------------//
                
    ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "CONTROLLO OGGETTI PRECARICATI DA IMPORTARE:" & vbCrLf & _
               "Oggetti importati: " & importedQueryesCount_i & vbCrLf & importedNameQueryes_s & vbCrLf & _
               "Oggetti esistenti cancellati: " & deleteQueryesCount_i & vbCrLf & deleteNameQueryes_s, vbInformation
    End If
    
                '//ciclo nella collezione per il controllo delle QUERY precaricate per SELEZIONARE LA PRIMA
                '//--------------------------------------------------------------------------------------//
                    For Each objCollection_objectName_v In objectQueries
                        
                        
                        '//vado all'ultima query
                        DoCmd.SelectObject acQuery, objCollection_objectName_v, True
                        
                    Next objCollection_objectName_v
                                
                    '//PRIMA DI USCIRE RILASCIO GLI OGGETTI
                    Set dbCurrent = Nothing
                    Set objectQueries = Nothing
                    
                '//--------------------------------------------------------------------------------------//

'//=====================================================================================================================//
'//                            STEP_03 = IMPORTO LE TABELLE E LE QUERY *** fine ***
'//
'//=====================================================================================================================//

'// CHIUSURA E CONTROLLO ERRORI
'//------------------------------------------------------------------------------------------------//
Exit Sub

ErrorHandler:
    MsgBox "Errore durante l'importazione: " & Err.Description, vbCritical
    Debug.Print "Errore durante l'importazione: " & Err.Description, vbCritical
    If Not dbSource Is Nothing Then
        dbSource.Close
    End If
    Set dbSource = Nothing
    Set dbCurrent = Nothing
End Sub   '//Private Sub Cmd_Importa_QUERY_Click()


' Funzione per verificare se un elemento esiste nella collezione
Function InCollectionQueries_Funct_b(col As Collection, key As String) As Boolean
    Dim varItem As Variant
    On Error Resume Next
    For Each varItem In col
        If varItem = key Then
            InCollectionQueries_Funct_b = True
            Exit Function
        End If
    Next varItem
    InCollectionQueries_Funct_b = False
    On Error GoTo 0
End Function

'//IMPORTA GLI OGGETTI TABLE E QUERY DEL DB ESTERNO IN TABELLA  *** FINE ***
'//----------------------------------------------------------------------------------------//



'//IMPORTA QUERY HELP *** INIZIO ***
'//-----------------------------------------------------------------------------------//

Private Sub Cmd_Importa_QUERY_HELP_Click()

    '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
    '//--------------------------------------------------------------------------------//--------//
    '//NOTE                 -> Apro il file di tipo doc, zip o pdf
    '//CODICE               -> Function ApriFilePdf.01.01
    '//PARAMETRI            -> par_Path_s         = PATH _
                            -> par_NameFile_s               = NOME FILE _
                            -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                            
      '//IMPOSTO LE VARIABILI PATH E FILE
      '//
      Dim MyFile_s As String
      '//
      Dim MyPath_s  As String
      
            '//imposto i parametri - LA PATH
            '
            MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\OBJECT\HELP\"
            '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
            '
            MyFile_s = "HELP_Cmd_IMPORTA_QUERY.pdf "
                
                
                       
         '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
            Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto IMPORTA OGGETTI incluso in questa form _
            anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'importazione e cancellazione in caso di resettamento _
            totale.
            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
      
            
    '//--------------------------------------------------------------------------------//--------//

End Sub

'//IMPORTA QUERY HELP *** FINE ***
'//-----------------------------------------------------------------------------------//










'//IMPORTA REPORT HELP *** INIZIO ***
'//-----------------------------------------------------------------------------------//

Private Sub Cmd_Importa_REPORT_HELP_Click()

    '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
    '//--------------------------------------------------------------------------------//--------//
    '//NOTE                 -> Apro il file di tipo doc, zip o pdf
    '//CODICE               -> Function ApriFilePdf.01.01
    '//PARAMETRI            -> par_Path_s         = PATH _
                            -> par_NameFile_s               = NOME FILE _
                            -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                            
      '//IMPOSTO LE VARIABILI PATH E FILE
      '//
      Dim MyFile_s As String
      '//
      Dim MyPath_s  As String
      
            '//imposto i parametri - LA PATH
            '
            MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\OBJECT\HELP\"
            '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
            '
            MyFile_s = "HELP_Cmd_IMPORTA_REPORT.pdf "
                
                
                       
         '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
            Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto IMPORTA OGGETTI incluso in questa REPORT _
            anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'importazione e cancellazione in caso di resettamento _
            totale.
            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
      
            
    '//--------------------------------------------------------------------------------//--------//

End Sub

'//IMPORTA REPORT HELP *** FINE ***
'//-----------------------------------------------------------------------------------//





'//=========================================================================================//
'//
'//                 GESTIONE DELLE QUERY ESTERNE O INTERNE   *** FINE ***
'//@GESTIONE@DELLE@QUERY
'//=========================================================================================//





    






'//=========================================================================================//
'//
'//                 GESTIONE DELLE TABELLE ESTERNE O INTERNE   *** INIZIO ***
'//@GESTIONE@DELLE@TABELLE                   @GESTONE@TABELLE
'//=========================================================================================//






'//IMPORTA GLI OGGETTI TABLE DEL DB ESTERNO IN TABELLA *** INIZIO ***
'//----------------------------------------------------------------------------------------//
'//NOTE:   dal database di origine importa solo gli oggetti Table controllando prima se _
            nel db corrente di destinazione già esiste la Table per effettuare la cancellazione _
            prima dell'importazione.


Private Sub Cmd_Importa_TABLE_Click()

 Dim dbCurrent As DAO.Database
 Dim dbSource As DAO.Database
 Dim tdf As DAO.TableDef
 Dim strSourceDbPath As String
 Dim strTableName As String
' Dim objectTables As Collection
 Dim connectedTABELLES As Collection
 Dim varItem As Variant
 Dim tableExistsInCollectionTables_Funct_b As Boolean
         


    On Error GoTo ErrorHandler



    '//=====================================================================================================================//
    '//                            STEP_01 = RECUPERO PATH E COLLECTION
    '//RECUPERO LA COLLEZIONE TABELLE DA IMPORTARE E FACCIO IL CONTROLLO
    '//=====================================================================================================================//
    '//NOte Chiamo la funzione in cui vengo precariate le form da importare in una _
            collection che viene restituiscta con oggetto dalla funzione e impostata _
            nella variabile oggetto formsCollection.
        
    
                '//RESET
                'Inizializza i contatori e le stringhe di riepilogo
                importedTablesCount_i = 0
                importedTables_s = ""
                deletedTablesCount_i = 0
                
                                  
              '//ATTIVO LA COLLEZIONE
              '//------------------------------------------------------------------------//
              '//Note: CHIAMO LA FUNZIONE COLLECTION TABLE per ottenere la collezione di TABELLE
                
                
                    Set tablesCollection = CollectionTable_PFunct()
                    
                    
                            '//2024.08.03_CollectionTable_PFunct  = SE LA COLLECTION E' VUOTA ESCI DALLA ROUTINE _
                        faccio un controllo sul numero degli oggetti della collection se = 0 allora _
                        non ci sono oggetti da caricare per cui si esce dalla routine.
                        
                    If tablesCollection.count = 0 Then Exit Sub
   
    
                
              '//------------------------------------------------------------------------//

                
                    
                
                
        '//ciclo nella collezione per il controllo delle TABELLE precaricate
        For Each objCollection_tableName_v In tablesCollection
            Debug.Print "Form nella collezione: " & objCollection_tableName_v
            ' Aggiungi qui il codice per lavorare con ciascuna TABELLA nella collezione
               '//conteggio form DA IMPORTARE
                importedTablesCount_i = importedTablesCount_i + 1
                importedTables_s = importedTables_s & objCollection_tableName_v & vbCrLf
    
        Next objCollection_tableName_v
        
            
            ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
            If Me.ctr_AVVISO_ABILITATO_TXT = True Then
                MsgBox "CONTROLLO TABELLE PRECARICATE DA IMPORTARE:" & vbCrLf & _
                       "Tables importate: " & importedTABLECount & vbCrLf & importedTables_s & _
                       "Tables esistenti cancellate: " & deletedTABLECount & vbCrLf & deletedTables, vbInformation
            End If

    '//---------------------------------------------------------------------------//
    
    
    '//RECUPERO LA PATH ED IL FILE.MDB PER LE IMPORTAZIONI
    '//---------------------------------------------------------------------------//
    '//NOTE: chiamo la funzione e recupero la path ed il file.mdb facendo un controllo _
            se la variabile è nulla msg di avviso ed uscita dalla routine.
            
        '//reset
        sourceDBPath_s = ""
        
        '//@RECUPERO@PATH_(ed il file
        sourceDBPath_s = PathFile_s_pFunct
        
        If sourceDBPath_s = "" Then
            MsgBox "ATTENZIONE ERRORE path e file sono null uscita dalla routine", vbCritical, "MSG CONTROLLO PATH E FILE VALORIZZATI"
              If MyValue = 0 Then Exit Sub
        End If
        
          ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
            If Me.ctr_AVVISO_ABILITATO_TXT = True Then
                MsgBox "IMPORTO LE TABELLE DA QUESTO DATABASE; path e del file : " & sourceDBPath_s, vbExclamation, "MSG CHIAMATO LA FUNZIONE PATHFILE"
            End If
    '//=====================================================================================================================//
    '//                            STEP_01 = RECUPERO PATH E COLLECTION
    '//RECUPERO LA COLLEZIONE FORM DA IMPORTARE E FACCIO IL CONTROLLO          **** FINE ***
    '//=====================================================================================================================//
    
            '//=====================================================================================================================//
            '//                             STEP_02 = @CONTROLLO@PROGETTO
            '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** INIZIO ***
            '//=====================================================================================================================//
            '//NOTE: chiamo la funzione che controlla se il progetto corrente è il progetto originale, infatti _
                     non si possono caricare gli oggetti nel PROGETTO MSYS_N01_OGGETTI.mdb perhè è originale _
                     e vi è una distruzione degli stessi. LA FUNZIONE CONTROLLO SE E' UN PROGETTO ESTERNO AMMISSIBILE _
                     per caricare gli oggetti table, query , form ecc. da importare. _
                     01) CONTROLLO = MSG DI AVVISO CON POSSIBILITA DI ANNULLARE _
                     02) CONTROLLO = CONTROLLO SE STIAMO NEL PROGETTO ORIGINALE  E QUINDI ESCLUSIONE.
                     
                        
                        
                   '//01) PRIMO CONTROLLO _
                     MESSAGGIO DI IMPORTAZIONE 0 = ANNULLA  1 = CONTINUA
                   '//..........................................................//
            
                         '//Attenzione L'AVVISO produce il messaggio solo se il _
                            parametro è TRUE altrimenti il messaggio viene escluso.
                         MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)
                         
                         If MyValue = 0 Then Exit Sub
                     '//..........................................................//
            
                         
                 
                   
                   '//02) SECONDO CONTROLLO
                   '//CHIAMO IL CONTROLLO DEL PROGETTO
                   '//..........................................................//
                   '//Note  : True = siamo nel progetto originale non possiamo _
                              importare nessun oggetto perche distruggiamo il db corrente. _
                              Fale = siamo in un progetto esterno possiamo caricare tutto.
                                  
                                  
                       '//SE VIENE RESTITUITO TRUE è FUORI PROGETTO PUO CARICARE, SE FALSE EXIT PERCHE PROGETTO ORIGINALE
                       Bool1 = CheckProjectName_b_PFunct
                       
                       If Bool1 = True Then Exit Sub
                   '//..........................................................//
                   
             
            '//=====================================================================================================================//
            '//                             STEP_02 = @CONTROLLO@PROGETTO
            '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** FINE ***
            '//=====================================================================================================================//
            
    
       
       
   
    
    '//=====================================================================================================================//
    '//                            STEP_03 = IMPORTO LE TABELLE
    '//
    '//=====================================================================================================================//
            
            
         
    
                '//RESET
                '//Inizializza i contatori e le stringhe di riepilogo
                importedTablesCount_i = 0
                deletedTablesCount_i = 0
                deletedTables_s = ""
                importedTables_s = ""
                
          
            ' Apri il database corrente
            Set dbCurrent = CurrentDb
            
            ' Apri il database di origine
            Set dbSource = DBEngine.Workspaces(0).OpenDatabase(sourceDBPath_s)
            
            'Loop attraverso tutte le tabelle nel database di origine ossia del DATABASE ESTERNO
            For Each tdf In dbSource.TableDefs
                
                ' Salta le tabelle di sistema (i cui nomi iniziano con "MSys") aggiungo come parametro( la collezione table + nome tabella)
                If Left(tdf.Name, 4) <> "MSys" Or (Left(tdf.Name, 4) = "MSys" And InCollectionTables_Funct_b(objectTables, tdf.Name)) Then
                    ' Nome della tabella nel database di origine
                    strTableName = tdf.Name
                    
                    ' Controlla se la tabella è nella collezione
                    tableExistsInCollectionTables_Funct_b = False
                    For Each varItem In objectTables
                        If strTableName = varItem Then
                            tableExistsInCollectionTables_Funct_b = True
                            Exit For
                        End If
                    Next varItem
                    
                    ' Se la tabella è nella collezione, importala
                    If tableExistsInCollectionTables_Funct_b Then
                        ' Elimina la tabella se esiste già nel database corrente SE DA ERRORE SALTA alla successiva istruzione
                        On Error Resume Next
                        dbCurrent.TableDefs.Delete strTableName
                        deletedTablesCount_i = deletedTablesCount_i + 1
                        
                        On Error GoTo 0
                        
                        ' Copia la tabella dal database di origine al database corrente
                        DoCmd.TransferDatabase acImport, "Microsoft Access", sourceDBPath_s, acTable, strTableName, strTableName
                        
                           '//conteggio TABELLE IMPORTATE
                            importedTablesCount_i = importedTablesCount_i + 1
                            importedTables_s = importedTables_s & strTableName & vbCrLf
                
                        
                    End If
                End If
            Next tdf
            
                    
        '//TODO: MANCA IL CONTROLLO SULLA TABELLA DI SISTEMA DEL DB ESTERNO MSysObjects DA CREARE VEDI FUNZIONE Cmd_Importa_MODULI_Click
        '//----------------------------------------------------------------------------------------------------------------------//
        
            '//TODO: adattare il controllo sulla tabella di sistema esterna ... vai a Cmd_Importa_MODULI_Click

        '//----------------------------------------------------------------------------------------------------------------------//
                    
                      ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
                        If Me.ctr_AVVISO_ABILITATO_TXT = True Then
                            MsgBox "CONTROLLO TABELLE PRECARICATE DA IMPORTARE:" & vbCrLf & _
                                   "Tables importate: " & importedTablesCount_i & vbCrLf & DELETETableCount_i & _
                                   "Tables esistenti cancellate: " & deletedTablesCount_i & vbCrLf & deletedTables_s, vbInformation
                        End If
            
                    
                    
                    Set dbCurrent = Nothing
                    
                '//ciclo nella collezione per il controllo delle TABELLE precaricate per SELEZIONARE LA PRIMA
                '//--------------------------------------------------------------------------------------//
                    For Each objCollection_objectName_v In tablesCollection
                    
                        
                        '//vado all'ultima query
                        DoCmd.SelectObject acTable, objCollection_tableName_v, True
                        
                    Next objCollection_objectName_v
                                
                    '//PRIMA DI USCIRE RILASCIO GLI OGGETTI
                    Set dbCurrent = Nothing
                    Set objectQueries = Nothing
                    
                '//--------------------------------------------------------------------------------------//
                
                
                    
  

                    
            
    '//=====================================================================================================================//
    '//                            STEP_03 = IMPORTO LE TABELLE *** fine ***
    '//
    '//=====================================================================================================================//


'// CHIUSURA E CONTROLLO ERRORI
'//------------------------------------------------------------------------------------------------//
Exit Sub

ErrorHandler:
    MsgBox "Errore durante l'importazione: " & Err.Description, vbCritical
    Debug.Print "Errore durante l'importazione: " & Err.Description, vbCritical
    If Not sourceDB Is Nothing Then
        sourceDB.Close
    End If
    Set sourceDB = Nothing
    Set ws = Nothing
End Sub   '//Private Sub Cmd_Importa_TABLE_Click()

'//@CONTROLLO@ESISTENZA@TABELLE@IN.COLLECTION
'Funzione per verificare se un elemento esiste nella collezione
Function InCollectionTables_Funct_b(col As Collection, key As String) As Boolean
    Dim varItem As Variant
    On Error Resume Next
    For Each varItem In col
        '//TODO: trovare il perche non cancella la tabella ???
        Debug.Print "CONTROLLO InCollectionTables_Funct_b TRA varItem = " & varItem & " - E KEY = "; key
        If varItem = key Then
            InCollectionTables_Funct_b = True
            Exit Function
        End If
    Next varItem
    InCollectionTables_Funct_b = False
    On Error GoTo 0
End Function



'//IMPORTA GLI OGGETTI TABLE DEL DB ESTERNO IN TABELLA  *** FINE ***
'//----------------------------------------------------------------------------------------//



'//=====================================================================================================================//
'//                     DELETE OGGETTI TABLE DEL DB ESTERNO IN TABELLA *** INIZIO ***
'//                     + DELETE TABLE HELP
'//=====================================================================================================================//


'//CANCELLA GLI OGGETTI TABLE *** INIZIO *** 'DELETE TABLES
'//==========================================================================================================//

Private Sub Cmd_Delete_TABLE_Click()

    Dim dbCurrent As DAO.Database
    Dim tdf As DAO.TableDef
    Dim objectTables As Collection
    Dim varItem As Variant
    
'//TODO: controllare se tale routine funziona???
On Error GoTo Err_Cmd_Delete_TABLE_Click
    
    
    '//RESET
    deletedTablesCount_i = 0
    deletedTables_s = ""
   
   
   
                   '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** INIZIO ***
                '//=====================================================================================================================//
                '//NOTE: chiamo la funzione che controlla se il progetto corrente è il progetto originale, infatti _
                         non si possono caricare gli oggetti nel PROGETTO MSYS_N01_OGGETTI.mdb perhè è originale _
                         e vi è una distruzione degli stessi. LA FUNZIONE CONTROLLO SE E' UN PROGETTO ESTERNO AMMISSIBILE _
                         per caricare gli oggetti table, query , form ecc. da importare. _
                         01) CONTROLLO = MSG DI AVVISO CON POSSIBILITA DI ANNULLARE _
                         02) CONTROLLO = CONTROLLO SE STIAMO NEL PROGETTO ORIGINALE  E QUINDI ESCLUSIONE.
                                 
                           '//01) PRIMO CONTROLLO _
                             MESSAGGIO DI IMPORTAZIONE 0 = ANNULLA  1 = CONTINUA
                           '//..........................................................//
                             '//Attenzione L'AVVISO produce il messaggio solo se il _
                                parametro è TRUE altrimenti il messaggio viene escluso.
                             MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)
                             
                             If MyValue = 0 Then Exit Sub
                           '//..........................................................//
                                 
                           '//02) SECONDO CONTROLLO
                           '//CHIAMO IL CONTROLLO DEL PROGETTO
                           '//..........................................................//
                           '//Note  : True = siamo nel progetto originale non possiamo _
                                      importare nessun oggetto perche distruggiamo il db corrente. _
                                      False = siamo in un progetto esterno possiamo caricare tutto.
                                          
                               '//SE VIENE RESTITUITO TRUE è FUORI PROGETTO PUO CARICARE, SE FALSE EXIT PERCHE PROGETTO ORIGINALE
                               Bool1 = CheckProjectName_b_PFunct
                               
                               If Bool1 = True Then Exit Sub
                           '//..........................................................//
                
                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** FINE ***
                '//=====================================================================================================================//

   
   
        
              '//ATTIVO LA COLLEZIONE
              '//------------------------------------------------------------------------//
              '//Note: CHIAMO LA FUNZIONE COLLECTION TABLE per ottenere la collezione di TABELLE
                
                '//
                Set tablesCollection = CollectionTable_PFunct()
                
              '//------------------------------------------------------------------------//
        
    
    ' Apri il database corrente
    Set dbCurrent = CurrentDb
    
    ' Loop attraverso tutte le tabelle nel database corrente
    For Each tdf In dbCurrent.TableDefs
        ' Controlla se la tabella è nella collezione e passo tutta la collezione objectTables + il nome della tabella da controllare
        If InCollectionTables_Funct_b(tablesCollection, tdf.Name) Then
            ' Elimina la tabella
            On Error Resume Next
                deletedTablesCount_i = deletedTablesCount_i + 1
                DoCmd.DeleteObject acTable, tdf.Name
            On Error GoTo 0
        End If
    Next tdf
    
    
            ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
            If Me.ctr_AVVISO_ABILITATO_TXT = True Then
                MsgBox "LE TABELLE CANCELLATA DA QUESTO DATABASE SONO :  " & deletedTablesCount_i, vbExclamation, "MSG CHIAMATO LA FUNZIONE PATHFILE"
            End If


                '




'//USCITA  E GESTIONE ERRORI
'//..............................................................................................................

Exit_Cmd_Delete_TABLE_Click:
    Exit Sub

Err_Cmd_Delete_TABLE_Click:
        
    Set dbCurrent = Nothing
    
    MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & " Procedura -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Stop
    Resume Exit_Cmd_Delete_TABLE_Click

End Sub




'//*** FINE ***
'//CANCELLA GLI OGGETTI TABLE *** INIZIO *** 'DELETE TABLES
'//==========================================================================================================//


'//DELETE TABLE HELP *** INIZIO ***
'//-----------------------------------------------------------------------------------//



Private Sub Cmd_Delete_TABLE_HELP_Click()
    
    '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
    '//--------------------------------------------------------------------------------//--------//
    '//NOTE                 -> Apro il file di tipo doc, zip o pdf
    '//CODICE               -> Function ApriFilePdf.01.01
    '//PARAMETRI            -> par_Path_s         = PATH _
                            -> par_NameFile_s               = NOME FILE _
                            -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                            
      '//IMPOSTO LE VARIABILI PATH E FILE
      '//
      Dim MyFile_s As String
      '//
      Dim MyPath_s  As String
      
            '//imposto i parametri - LA PATH
            '
            MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\OBJECT\HELP\"
            '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
            '
            MyFile_s = "HELP_Cmd_Delete_TABLE.pdf "
                
                
                       
         '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
            Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto DELETE OGGETTI incluso in questa MODULI _
            anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'DELETEzione e cancellazione in caso di resettamento _
            totale.
            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
      
            
    '//--------------------------------------------------------------------------------//--------//

End Sub




'//DELETE TABLE HELP *** FINE ***
'//-----------------------------------------------------------------------------------//


'//=====================================================================================================================//
'//                     DELETE OGGETTI TABLE DEL DB ESTERNO IN TABELLA *** FINE ***
'//                     + DELETE TABLE HELP
'//=====================================================================================================================//








'//@HELP@IMPORTA@TABELLE
Private Sub Importa_TABLE_HELP_Click()



    '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
    '//--------------------------------------------------------------------------------//--------//
    '//NOTE                 -> Apro il file di tipo doc, zip o pdf
    '//CODICE               -> Function ApriFilePdf.01.01
    '//PARAMETRI            -> par_Path_s         = PATH _
                            -> par_NameFile_s               = NOME FILE _
                            -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                            
      '//IMPOSTO LE VARIABILI PATH E FILE
      '//
      Dim MyFile_s As String
      '//
      Dim MyPath_s  As String
      
            '//imposto i parametri - LA PATH
            '
            MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\OBJECT\HELP\"
            '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
            '
            MyFile_s = "HELP_Cmd_IMPORTA_TABLE.pdf "
                
                
            
            
         '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
            Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto IMPORTA OGGETTI incluso in questa form _
            anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'importazione e cancellazione in caso di resettamento _
            totale.
            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
             
            
    '//--------------------------------------------------------------------------------//--------//

      
End Sub



'//=====================================================================================================================//
'//                     DELETE OGGETTI QUERY DEL DB ESTERNO IN TABELLA *** INIZIO ***
'//                     + DELETE QUERY HELP
'//=====================================================================================================================//





'//CANCELLA GLI OGGETTI TABELLE *** INIZIO *** DELETE TABLE
'//==========================================================================================================//
'//@GESTIONE@DELLE@TABELLE    (@MODELLO@DELETE@OGGETTI)

'// PARAMETRI DA CAMBIARE : _
    -----> TABELLE _
    -----> DELLE@TABELLE _
    -----> DELETE TABLE _
    -----> CollectionTable_PFunct _
    -----> objectTable _
    -----> MATRICE_TABLE_Names_s _
    -----> tableiCount _
    -----> TableDef _
    -----> tdf _


Private Sub Cmd_Delete_TABELLE_Click()

    Dim dbCurrent As DAO.Database
    Dim qry As DAO.TableDef
    Dim objectQueries As Collection
    Dim varItem As Variant
    Dim MATRICE_TABLE_Names_s() As String             'matrice di TABELLE
    Dim tempCount As Integer
    Dim tableiCount As Integer
    Dim tdf As TableDef



'//TODO: controllare se tale routine funziona???
On Error GoTo Err_Cmd_Delete_TABELLE_Click


    '//RESET
    iCount = 0
    deleteTABELLEesCount_i = 0
    deletedTablesCount_i = 0
    deletedTables_s = ""




                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** INIZIO ***
                '//=====================================================================================================================//
                '//NOTE: chiamo la funzione che controlla se il progetto corrente è il progetto originale, infatti _
                         non si possono caricare gli oggetti nel PROGETTO MSYS_N01_OGGETTI.mdb perhè è originale _
                         e vi è una distruzione degli stessi. LA FUNZIONE CONTROLLO SE E' UN PROGETTO ESTERNO AMMISSIBILE _
                         per caricare gli oggetti table, query , form ecc. da importare. _
                         01) CONTROLLO = MSG DI AVVISO CON POSSIBILITA DI ANNULLARE _
                         02) CONTROLLO = CONTROLLO SE STIAMO NEL PROGETTO ORIGINALE  E QUINDI ESCLUSIONE.
                                 
                           '//01) PRIMO CONTROLLO _
                             MESSAGGIO DI IMPORTAZIONE 0 = ANNULLA  1 = CONTINUA
                           '//..........................................................//
                             '//Attenzione L'AVVISO produce il messaggio solo se il _
                                parametro è TRUE altrimenti il messaggio viene escluso.
                             MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)
                             
                             If MyValue = 0 Then Exit Sub
                           '//..........................................................//
                                 
                           '//02) SECONDO CONTROLLO
                           '//CHIAMO IL CONTROLLO DEL PROGETTO
                           '//..........................................................//
                           '//Note  : True = siamo nel progetto originale non possiamo _
                                      importare nessun oggetto perche distruggiamo il db corrente. _
                                      False = siamo in un progetto esterno possiamo caricare tutto.
                                          
                               '//SE VIENE RESTITUITO TRUE è FUORI PROGETTO PUO CARICARE, SE FALSE EXIT PERCHE PROGETTO ORIGINALE
                               Bool1 = CheckProjectName_b_PFunct
                               
                               If Bool1 = True Then Exit Sub
                           '//..........................................................//
                
                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** FINE ***
                '//=====================================================================================================================//


          '//ATTIVO LA COLLEZIONE
         '//------------------------------------------------------------------------//
          '//Note: CHIAMO LA FUNZIONE COLLECTION TABELLE per ottenere la collezione di TABELLE



        '//
        Set tablesCollection = CollectionTable_PFunct()

          '//------------------------------------------------------------------------//

            ' Loop attraverso tutte I REPORT nel database corrente
             Set dbCurrent = CurrentDb

             '//CALCOLO INDICE TABELLE
             iCount = dbCurrent.TableDefs.count

            Set dbCurrent = CurrentDb


            '//se ci sono TABELLE da cancellare
            If iCount > 0 Then

            '//reimposto la matrice sulle TABELLE esistenti nel db corrente
            ReDim MATRICE_TABLE_Names_s(iCount - 1)

            ' Costruisci un array con i nomi dei TABELLE
            tempCount = 0

            '//CARICO LA MATRICE: ciclo for per caricare MATRICE_TABLE_Names_s(tempCount)
            For Each tdf In dbCurrent.TableDefs
                MATRICE_TABLE_Names_s(tempCount) = tdf.Name
                tempCount = tempCount + 1
            Next tdf

            ' Cancella i TABELLE usando l'array dei nomi
            For i = 0 To tempCount - 1
                On Error Resume Next  ' Aggiungi gestione errori specifica per ogni iterazione
                Debug.Print "Cancellazione TABELLE: " & MATRICE_TABLE_Names_s(i)

                '//CONTROLLO MATRICE = COLLECTION
                Bool1 = InCollectionTABELLEes_Funct_b(tablesCollection, MATRICE_TABLE_Names_s(i))

                '//cancello solo se la matrice (i) esiste nella collection
                If Bool1 = True Then
                    Debug.Print
                    Debug.Print "cancello la TABELLE trovata sia nella collezione che nella matrice : " & MATRICE_TABLE_Names_s(i)
                    DoCmd.DeleteObject acTABELLE, MATRICE_TABLE_Names_s(i)

                    If Err.number <> 0 Then
                    Debug.Print "Errore durante la cancellazione del TABELLE: " & MATRICE_TABLE_Names_s(i) & " - " & Err.Description
                    Err.Clear  ' Pulisci l'errore per la prossima iterazione
                    Else
                    deleteTABELLEesCount_i = deleteTABELLEesCount_i + 1
                    End If

                End If '//If Bool1 = True Then

                On Error GoTo 0  ' Reimposta gestione errori
            Next i
            End If

        ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
        If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "LE TABELLE CANCELLATA DA QUESTO DATABASE SONO :  " & deleteTABELLEesCount_i, vbExclamation, "MSG CHIAMATO LA FUNZIONE PATHFILE"
        End If





'//USCITA  E GESTIONE ERRORI
'//..............................................................................................................

Exit_Cmd_Delete_TABELLE_Click:
    Exit Sub

Err_Cmd_Delete_TABELLE_Click:

    Set dbCurrent = Nothing

    MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & " Procedura -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Stop
    Resume Exit_Cmd_Delete_TABELLE_Click

End Sub

'//*** FINE ***
'//SENZA_PARAMETRI
'//========================================================================================================================================//



' Funzione per verificare se un elemento esiste nella collezione
Function InCollectionTABELLEes_Funct_b(col As Collection, key As String) As Boolean
    Dim varItem As Variant
    On Error Resume Next
    '//ITERNO NELLA COLLECTION PASSATA COME PARAMETRO
    For Each varItem In col

    Debug.Print
    Debug.Print "CONTROLLO COLLEZIONE = MATRICE (I)"
    Debug.Print "collezione: " & varItem & " = " & "matrice key: " & key
    Debug.Print
    If varItem = key Then
        InCollectionTABELLEes_Funct_b = True
        Debug.Print "trovato restituisco : " & InCollectionTABELLEes_Funct_b
        Exit Function
    End If
    Next varItem
    InCollectionTABELLEes_Funct_b = False
    On Error GoTo 0
End Function



'//CANCELLA GLI OGGETTI TABELLE *** FINE *** DELETE TABLE
'//=========================================================================================================






'//=========================================================================================//
'//
'//                 GESTIONE DELLE TABELLE ESTERNE O INTERNE   *** FINE ***
'//
'//=========================================================================================//





'//=========================================================================================//
'//
'//                 GESTIONE DELI MODULI ESTERNE O INTERNE   *** inizio ***
'//@GESTIONE@DELLE@MODULI                    @GESTONE@MODULI
'//=========================================================================================//



'//==========================================================================================================//
'// IMPORTA GLI OGGETTI MODULI *** INIZIO *** '@IMPORT@MODULI
'//==========================================================================================================//


Private Sub Cmd_Importa_MODULI_Click()
    Dim obj As AccessObject
    Dim dbs As Object
    Dim ProceduraMessaggioErrore_s As String
    Dim modulesCollection As Collection
    Dim modulesCollection_v As Variant
    Dim moduleExists As Boolean
    Dim deletedModulesCount_i As Integer
    Dim deletetModulesCount_i As Integer
    Dim deletedModules_s As String
    Dim importedCount As Integer
    Dim importedModules_s As String
    Dim sourceDBPath As String
    Dim externalDB As DAO.Database
    Dim rs As DAO.Recordset
    
    On Error GoTo Err_Cmd_Import_MODULI_Click
   
    ' Reset count
    deletedModulesCount_i = 0
    deletetModulesCount_i = 0
    deletedModules_s = ""
    importedCount = 0
    importedModules_s = ""
    
    
    
    
    
                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** INIZIO ***
                '//=====================================================================================================================//
                '//NOTE: chiamo la funzione che controlla se il progetto corrente è il progetto originale, infatti _
                         non si possono caricare gli oggetti nel PROGETTO MSYS_N01_OGGETTI.mdb perhè è originale _
                         e vi è una distruzione degli stessi. LA FUNZIONE CONTROLLO SE E' UN PROGETTO ESTERNO AMMISSIBILE _
                         per caricare gli oggetti table, query , form ecc. da importare. _
                         01) CONTROLLO = MSG DI AVVISO CON POSSIBILITA DI ANNULLARE _
                         02) CONTROLLO = CONTROLLO SE STIAMO NEL PROGETTO ORIGINALE  E QUINDI ESCLUSIONE.
                                 
                           '//01) PRIMO CONTROLLO _
                             MESSAGGIO DI IMPORTAZIONE 0 = ANNULLA  1 = CONTINUA
                           '//..........................................................//
                             '//Attenzione L'AVVISO produce il messaggio solo se il _
                                parametro è TRUE altrimenti il messaggio viene escluso.
                             MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)
                             
                             If MyValue = 0 Then Exit Sub
                           '//..........................................................//
                                 
                           '//02) SECONDO CONTROLLO
                           '//CHIAMO IL CONTROLLO DEL PROGETTO
                           '//..........................................................//
                           '//Note  : True = siamo nel progetto originale non possiamo _
                                      importare nessun oggetto perche distruggiamo il db corrente. _
                                      False = siamo in un progetto esterno possiamo caricare tutto.
                                          
                               '//SE VIENE RESTITUITO TRUE è FUORI PROGETTO PUO CARICARE, SE FALSE EXIT PERCHE PROGETTO ORIGINALE
                               Bool1 = CheckProjectName_b_PFunct
                               
                               If Bool1 = True Then Exit Sub
                           '//..........................................................//
                
                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** FINE ***
                '//=====================================================================================================================//

    
    
    
    
    
    
    
    
    '// ATTIVA LA COLLEZIONE DEI MODULI
    '//------------------------------------------------------------------------//

        ' Recupero la collezione moduli da importare
        Set modulesCollection = CollectionModuli_PFunct()
    
        
        
            '//2024.08.03_CollectionModuli_PFunct  = SE LA COLLECTION E' VUOTA ESCI DALLA ROUTINE _
                        faccio un controllo sul numero degli oggetti della collection se = 0 allora _
                        non ci sono oggetti da caricare per cui si esce dalla routine.
                        
            If modulesCollection.count = 0 Then GoTo Exit_Cmd_Import_MODULI_Click
   
        
    '//------------------------------------------------------------------------//
    
    ' Recupero la path e il file dalla casella di testo
    sourceDBPath = Me.sourceDBPath_s_Txt
        '//controllo path recuperata
        Debug.Print sourceDBPath
        
    If sourceDBPath = "" Then
        MsgBox "ATTENZIONE ERRORE: path e file sono null, uscita dalla routine", vbCritical, "MSG CONTROLLO PATH E FILE VALORIZZATI"
        Exit Sub
    End If
    
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "IMPORTO LE FORM DA QUESTO DATABASE; path e del file: " & sourceDBPath, vbExclamation, "MSG CHIAMATO LA FUNZIONE PATHFILE"
    End If
    
    ' Apro il database esterno
    Set externalDB = DBEngine.Workspaces(0).OpenDatabase(sourceDBPath)
    
    ' Ciclo nella collezione per il controllo dei moduli precaricati
    For Each modulesCollection_v In modulesCollection
        Debug.Print "I MODULI nella collezione: " & modulesCollection_v
        
        ' Verifica se il modulo esiste nel database corrente
        moduleExists = False
        Set dbs = Application.CurrentProject
        For Each obj In dbs.AllModules
            If obj.Name = modulesCollection_v Then
                moduleExists = True
                Exit For
            End If
        Next obj
        
        If moduleExists Then
            ' Se il modulo esiste nel database corrente, cancellalo
            DoCmd.DeleteObject acModule, modulesCollection_v
            deletetModulesCount_i = deletetModulesCount_i + 1
            deletedModules_s = deletedModules_s & modulesCollection_v & vbCrLf
            Debug.Print "Modulo eliminato: " & modulesCollection_v
        End If
        
        
        
            '//CONTROLLO SULLA TABELLA DI SISTEMA DEL DB ESTERNO MSysObjects
            '//----------------------------------------------------------------------------------------------------------------------//
            ' Verifica se il modulo esiste nel database esterno
            ' La riga SELECT Name FROM MSysObjects WHERE Type=-32761 è una query SQL utilizzata per selezionare i
            ' nomi degli oggetti di tipo modulo nel database di Access. Ecco una spiegazione dettagliata:
    
            ' MSysObjects: è una tabella di sistema in Microsoft Access che contiene informazioni su tutti gli oggetti nel database, _
              come tabelle, query, moduli, report, ecc. Type=-32761: questo filtro nella clausola WHERE specifica il tipo di oggetto _
              da selezionare. In Access, ogni tipo di oggetto ha un valore numerico associato nella tabella MSysObjects. _
              Il valore -32761 è il codice che rappresenta i moduli (VBA Modules).
             'Quindi, la query SELECT Name FROM MSysObjects WHERE Type=-32761 seleziona tutti i nomi degli oggetti nella tabella _
             di sistema MSysObjects che sono identificati come moduli VBA.
            
                '//APRO IL DB ESTERNO su eseguot una query sql sulla tabella di sistema estraendo un recordset superando due filtri _
                    il PRIMO FILTRO     : -32761 il codice dei moduli (estraggo solo i moduli) _
                    il SECONDO FILETRO  : il nome del modulo recuperato dalla collection (se esisten nella tabella di sistema)
                moduleExists = False
                Set rs = externalDB.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type=-32761 AND Name='" & modulesCollection_v & "'")
                '//SE IL RS E' POPOLATO vuol dire che nella tabella di sistema, del progetto mdb esiste il modulo ricercato
                If Not rs.EOF Then
                    '//imposto a true = trovato modulo
                    moduleExists = True
                End If
                rs.Close
                
                '//TRUE = esiste il modulo e lo importo
                If moduleExists Then
                    ' Se il modulo esiste nel database esterno, importalo
                    DoCmd.TransferDatabase acImport, "Microsoft Access", sourceDBPath, acModule, modulesCollection_v, modulesCollection_v
                    importedCount = importedCount + 1
                    importedModules_s = importedModules_s & modulesCollection_v & vbCrLf
                    Debug.Print "Modulo importato: " & modulesCollection_v
                Else
                    ' Modulo non trovato nel database esterno messaggio print.
                    Debug.Print "Modulo non trovato nel database esterno: " & modulesCollection_v
                End If
                
                
            '//----------------------------------------------------------------------------------------------------------------------//
        
    Next modulesCollection_v
    
    ' Chiudo il database esterno
    externalDB.Close
    
    ' Messaggio finale di riepilogo
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "CONTROLLO FORM PRECARICATE DA IMPORTARE:" & vbCrLf & _
               "Modules importate: " & importedCount & vbCrLf & importedModules_s & _
               "Modules esistenti cancellate: " & deletetModulesCount_i & vbCrLf & deletedModules_s, vbInformation
    End If
    
                  '//ciclo nella collezione per il controllo delle MODULI precaricate per SELEZIONARE LA PRIMA
                '//--------------------------------------------------------------------------------------//
                     For Each obj In dbs.AllModules
                            '//vado all'ultima query
                        DoCmd.SelectObject acModule, obj.Name, True
                    
                     Next obj
                '//--------------------------------------------------------------------------------------//

    
    

Exit_Cmd_Import_MODULI_Click:
    Exit Sub

Err_Cmd_Import_MODULI_Click:
    MsgBox Err.Description & " - Error Message -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Stop
    Resume Exit_Cmd_Import_MODULI_Click

End Sub



'//==========================================================================================================//
'// IMPORTA GLI OGGETTI MODULI *** INIZIO *** '@IMPORT@MODULI
'//==========================================================================================================//




'//=====================================================================================================================//
'//                     DELETE OGGETTI MODULI DEL DB ESTERNO IN TABELLA *** INIZIO ***
'//                     + DELETE MODULI HELP
'//=====================================================================================================================//

Private Sub Cmd_Delete_MODULI_Click()
    Dim obj As AccessObject
    Dim dbs As Object
    Dim ProceduraMessaggioErrore_s As String
    Dim modulesCollection As Collection
    'Dim modulesCollection_v As Variant
    Dim moduleExists As Boolean
    'Dim deletedModulesCount_i As Integer
    'Dim deletetModulesCount_i As Integer
    'Dim deletedModules_s As String
    'Dim importedModules_s As String
    
    On Error GoTo Err_Cmd_Delete_MODULI_Click
   
    ' Reset count
    deletedModulesCount_i = 0
    deletetModulesCount_i = 0
    deletedModules_s = ""
    importedCount = 0
    importedModules_s = ""
     
     
     
     
     

                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** INIZIO ***
                '//=====================================================================================================================//
                '//NOTE: chiamo la funzione che controlla se il progetto corrente è il progetto originale, infatti _
                         non si possono caricare gli oggetti nel PROGETTO MSYS_N01_OGGETTI.mdb perhè è originale _
                         e vi è una distruzione degli stessi. LA FUNZIONE CONTROLLO SE E' UN PROGETTO ESTERNO AMMISSIBILE _
                         per caricare gli oggetti table, query , form ecc. da importare. _
                         01) CONTROLLO = MSG DI AVVISO CON POSSIBILITA DI ANNULLARE _
                         02) CONTROLLO = CONTROLLO SE STIAMO NEL PROGETTO ORIGINALE  E QUINDI ESCLUSIONE.
                                 
                           '//01) PRIMO CONTROLLO _
                             MESSAGGIO DI IMPORTAZIONE 0 = ANNULLA  1 = CONTINUA
                           '//..........................................................//
                             '//Attenzione L'AVVISO produce il messaggio solo se il _
                                parametro è TRUE altrimenti il messaggio viene escluso.
                             MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)
                             
                             If MyValue = 0 Then Exit Sub
                           '//..........................................................//
                                 
                           '//02) SECONDO CONTROLLO
                           '//CHIAMO IL CONTROLLO DEL PROGETTO
                           '//..........................................................//
                           '//Note  : True = siamo nel progetto originale non possiamo _
                                      importare nessun oggetto perche distruggiamo il db corrente. _
                                      False = siamo in un progetto esterno possiamo caricare tutto.
                                          
                               '//SE VIENE RESTITUITO TRUE è FUORI PROGETTO PUO CARICARE, SE FALSE EXIT PERCHE PROGETTO ORIGINALE
                               Bool1 = CheckProjectName_b_PFunct
                               
                               If Bool1 = True Then Exit Sub
                           '//..........................................................//
                
                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** FINE ***
                '//=====================================================================================================================//
         
     
    ' Recupero la collezione moduli da importare
    Set modulesCollection = CollectionModuli_PFunct()
                        
    ' Ciclo nella collezione per il controllo delle form precaricate
    For Each modulesCollection_v In modulesCollection
        Debug.Print "I MODULI nella collezione: " & modulesCollection_v
        
        ' Verifica se il modulo esiste nel database
        moduleExists = False
        Set dbs = Application.CurrentProject
        
        '//itero nell'insime ALLMODULES
        For Each obj In dbs.AllModules
        
            '//CONFRONTA SE ESISTE IMPOSTA TRUE
            If obj.Name = modulesCollection_v Then
                moduleExists = True
                Exit For
            End If
        Next obj
        
        '//TRUE= ESISTE E CANCELLO IL MODULO
        If moduleExists Then
            ' Se il modulo esiste, cancellalo
            DoCmd.DeleteObject acModule, modulesCollection_v
            deletetModulesCount_i = deletetModulesCount_i + 1
            deletedModules_s = deletedModules_s & modulesCollection_v & vbCrLf
            Debug.Print "Modulo eliminato: " & modulesCollection_v
        Else
            ' Se il modulo non esiste, considera per l'importazione
            importedModulesCount_i = importedModulesCount_i + 1
            importedModules_s = importedModules_s & modulesCollection_v & vbCrLf
        End If
    Next modulesCollection_v
                        
    ' Messaggio finale di riepilogo
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "CONTROLLO MODULI PRECARICATE DA IMPORTARE:" & vbCrLf & _
               "Modules importate: " & importedModulesCount_i & vbCrLf & importedModules_s & _
               "Modules esistenti cancellate: " & deletetModulesCount_i & vbCrLf & deletedModules_s, vbInformation
    End If
    
     ' Messaggio finale di riepilogo
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
   
        ' Optional: display a message indicating that the modules have been deleted
        MsgBox "Tutti i moduli standard sono stati cancellati. Totale moduli cancellati: " & deletetModulesCount_i, vbInformation, "Moduli Cancellati"
    
    End If

Exit_Cmd_Delete_MODULI_Click:
    Exit Sub

Err_Cmd_Delete_MODULI_Click:
    MsgBox Err.Description & " - Error Message -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Stop
    Resume Exit_Cmd_Delete_MODULI_Click

End Sub

'//DELETE MODULI        ***FINE ***
'//==========================================================================================================//



'//Cmd_DELETE_MODULI_HELP_Click  *** INIZIO ***
'//-----------------------------------------------------------------------------------//

Private Sub Cmd_DELETE_MODULI_HELP_Click()

    '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
    '//--------------------------------------------------------------------------------//--------//
    '//NOTE                 -> Apro il file di tipo doc, zip o pdf
    '//CODICE               -> Function ApriFilePdf.01.01
    '//PARAMETRI            -> par_Path_s         = PATH _
                            -> par_NameFile_s               = NOME FILE _
                            -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                           
      '//IMPOSTO LE VARIABILI PATH E FILE
      '//
      Dim MyFile_s As String
      '//
      Dim MyPath_s  As String
     
            '//imposto i parametri - LA PATH
            '
            MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\OBJECT\HELP\"
            '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
            '
            MyFile_s = "HELP_Cmd_Delete_MODULI.pdf "
               
               
             '//CHIAMO LA FUNZIONE INTERNA PER L'APERTURA DEI FILE
                Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
                 
           
    '//--------------------------------------------------------------------------------//--------//




End Sub

'//Cmd_DELETE_MODULI_HELP_Click  *** FINE ***
'//-----------------------------------------------------------------------------------//


'//=====================================================================================================================//
'//                     DELETE OGGETTI MODULI DEL DB ESTERNO IN TABELLA *** FINE ***
'//                     + DELETE MODULI HELP
'//=====================================================================================================================//






'//=========================================================================================//
'//
'//                 GESTIONE DEI MODULI ESTERNI O INTERNI   *** FINE ***
'//
'//=========================================================================================//
















'//*********************************************************************************************//
'//***
'//***                                            VARIE FUNZIONI
'//*********************************************************************************************//

'//CONTROLLO DEL PROGETTO CORRENTE *** INIZIO *** Function CheckProjectName_b_PFunct() As Boolean
'//----------------------------------------------------------------------------------------//
'//NOTE: QUESTA ROUTINE  controlla se il progetto in cui si trova la form carica oggetti, _
    è quello di origine ossia il database MSYS_N01_OGGETTI.mdb. Interroganto la _
    proprita name se ottengo il nome del DB DI ORIGINE non eseguo l'importazione perchè altrimenti _
    distrutto il db di origine. QUESTE FUNZIONI SI POSSONO UTILIZZARE SOLO IN DB DIVERSI DA _
    QUELLO DI ORIGINE.  @PROGETTO@SORGENTE_(il nome del progetto sorgente che impedesce le attivia MSYS_N01_OGGETTI.mdb)
'//CODICE: @CHECK@PROJECT@PROGETTO@SORGENTE

'// CHIAVI_SOSTITUIBILI .......................................................***INIZIO *** _
        PROGETTO ---------> MSYS_N01_OGGETTI.mdb

'// CHIAVI_SOSTITUIBILI .......................................................***FINE ***


Function CheckProjectName_b_PFunct() As Boolean
    
    '//CHIAMO IL CONTROLLO DEL PROGETTO
    '//..........................................................//
    '//Note  : True = siamo nel progetto originale non possiamo _
               importare nessun oggetto perche distruggiamo il db corrente. _
               Fale = siamo in un progetto esterno possiamo caricare tutto.
                   
        '//solo se FALSE attivo la procedura
        'Bool1 = CheckProjectName_b_PFunct
    '//..........................................................//
    
        
    ' Variabile per memorizzare il nome del progetto corrente
    Dim projectName As String
    
    ' Ottieni il nome del progetto corrente
    projectName = Application.CurrentProject.Name
    
    ' Controlla se il nome del DATABASE QUELLO ORIGINARIO OSSIA "MSYS_N01_OGGETTI.mdb"
    If projectName = "MSYS_N01_OGGETTI.mdb" Then
        MsgBox "USCITA DALLA FUNZIONE: ATTENZIONE CI TROVIAMO NEL PROGETTO ORIGINALE E NON POSSIAMO CARICARE NE CANCELLARE " & _
               "NESSUN OGGETTO SUL PROGETTO DI PARTENZA DENOMINATO : " & projectName & vbCrLf & _
               "La funzione puo essere usata solo in un progetto Esterno.", vbCritical
               
            Debug.Print "CONTROLLO PROGETTTO: ATTENZIONE IL PROGETTO E' QUELLO ORIGINALE  : " & projectName & " NON POSSIAMO CARICARE NESSUNO OGGETTO!!"
            '//TRUE = progetto originale  NON SI PUO CARICARE NESSU OGGETTO
            CheckProjectName_b_PFunct = True
    Else
            '//FALSE = progetto diverso dall'originale  possiamo importare gli oggetti."
            CheckProjectName_b_PFunct = False
            Debug.Print "Il progetto si chiama diversamente: " & projectName & " POSSIAMO CARICARE GLI OGGETTI !!"
    End If
    
    
End Function

'//CONTROLLO DEL PROGETTO CORRENTE *** FINE *** Function CheckProjectName_b_PFunct() As Boolean
'//----------------------------------------------------------------------------------------//







'//AVVISO DI CONTINUAZIONE *** INIZIO ***
'//----------------------------------------------------------------------------------------//

Function CheckAvviso_i_PFunct(par_Bool As Boolean) As Integer
    
    '//il messaggio viene attivato solo se il parametro è TRUE
    If par_Bool = True Then
      '//AVVISO  - @MSGBOX
      Message = "ATTENZIONE QUESTO COMANDO DI IMPORTAZIONE VALE SOLO PER II PROGETTI DIVERSI  " & Chr$(13) _
                & " PER ANNULLARE INSERIRE 0"               ' Imposta il messaggio.
                Title = "AVVISO DI CONTINUAZIONE"           ' Imposta il titolo.
                Default = "1"                               ' Imposta il valore predefinito.
            
                ' Visualizza la finestra di dialogo nella posizione 100,100
                
                      MyMsg_V = InputBox(Message, Title, Default, 100, 100)
            
                
                
                
                '//CONTROLLO SE LA SCELTA = 1 O SCELTA = 0
                If MyMsg_V <> "" Or MyMsg_V <> Null Then
                    '//se scelta =0 oppure = 1 restisce un valore
                    CheckAvviso_i_PFunct = MyMsg_V
                Else
                    '//se annulla restituisce stringa = null e quindi restituisce 0 ANNULLATA
                    CheckAvviso_i_PFunct = 0
                    MyMsg_V = 0             ' 0 = SCELTA ANNULLATA ANCHE SE HAI FATTO ANNULLA
                End If
                
                
                '//AVVISO DI USCITA PERCHE SCELTO 0
                If MyMsg_V = 0 Then
                
                        '//AVVISO  DI USCITA- @MSGBOX
                        
                        MsgBox "USCITA DALLA ROUTINE", vbInformation, "AVVISO DI USCITA"
                        
   
                End If
                       
                
                
    Else
            '//CON IL MESSAGGIO DISABILITATO SI INTENDE SEMPRE 1 = ATTIVARE LA PROCEURA
             MyMsg_i = 1
            CheckAvviso_i_PFunct = MyMsg_i
        
    End If
            
End Function

'//AVVISO DI CONTINUAZIONE *** FINE ***
'//----------------------------------------------------------------------------------------//



'//APRI_IL_FILE_PDF - FUNZIONE INTERNA DI APERTURA DEI FILE -
'//==================================================================================================================//
'//METODO DI APERTURA DI UN PROGRAMMA ESTERNO O DI UN COMANDO DOS MEDIANTE "WScript.Shell"
'//PARAMETRI            -> passo 2 stringhe per parametro, la path e il nome del file.pdf
'//VALORE_DI_RITORNO    -> Nulla
'//NOTE                 -> Apro il file di tipo doc, zip o pdf
'//CODICE               -> Function ApriFilePdf_Interna_s_PFunct.01.00

Public Function ApriFilePdf_Interna_s_PFunct(par_Path_s As String, par_NameFile_s As String, par_IDGestione_lng As Long) As String

Dim Path_programma_s As String
Dim NameFile_s As String
Dim Stringa1 As String

On Error GoTo ApriFilePdf_Interna_s_PFunct_Err
        
    
'//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
'//--------------------------------------------------------------------------------//--------//
'//NOTE                 -> Apro il file di tipo doc, zip o pdf
'//CODICE               -> Function ApriFilePdf_Interna_s_PFunct.01.01
'//PARAMETRI            -> par_Path_s         = PATH _
                        -> par_NameFile_s               = NOME FILE _
                        -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                        
  '//IMPOSTO LE VARIABILI PATH E FILE
  '//Dim MyFile_s As String
  '//Dim MyPath_s  As String
  
        '//imposto i parametri - LA PATH
        'MyPath_s = "c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\"
        '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
        'MyFile_s = "Project_PROGETTI_MDB.sublime-project "
            
            
         '//chiamo la sub con i parametri =   'CALL (ApriFilePdf_Interna_s_PFunct Path_programma_s, NameFile_s, IDGestione_lng)
            'Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
             
        
'//--------------------------------------------------------------------------------//--------//
                        
         
                        '//APRO FILE PDF SPECIFICO
                        '//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>//
                        '//Note       : Apro il file PDF speficico indicando come paramentro la path completa
                        '//           incluso il file PDF.
                          
                          '//La path del file = Unisco la path + il file+estensione sopra individuata
                          par_Path_s = par_Path_s & par_NameFile_s & MyFileEstensione_s
                                 
                          '//controllo
                          Debug.Print par_Path_s
                                 
                          '//CREO LA SHELL come oggetto, attivo il comando "%comspec% /c start " e gli assegno la path per estesa con il nome del file
                          '// perchè Shell lancia un EXE e non il PDF, quindi è necessario costruire la stringa in modo da far lanciare
                          '// prima la sessione dos e poi il pdf..
                          '// La stringa di comando ("%comspec% /c start ") deve essere unica : comando shell + path + file
                          Set WshPDF = CreateObject("wscript.shell")
                          
                          
                          '//CONTROLLO
                          Debug.Print " ---------------------CONTROLLO PATH -----------------------------------"
                          Debug.Print par_Path_s
                          Debug.Print
                          
                          '// Unisco il comando pdf start + path definitiva
                          Stringa1 = "%comspec% /c start " & par_Path_s                               '//Comando di apertura
                          '//Attivo il comando
                          WshPDF.Run Stringa1
                                  
                                  
                          '//libero la memoria dalle variabili e dagli oggetti creati
                          WshPDF = Null
                          Stringa1 = vbNull
                          MyFileEstensione_s = vbNull
                          MyPath_s = vbNull
                            
                            '//sistemare ??? invio tasti ???
                            Lng = 0
                            For Lng = 1 To 90000000 Step 1
                            
                            Next Lng
                            
                          'SendKeys "{ESC}"
                          'SendKeys "{ESC}"
                          'SendKeys "{ESC}"
                          
                            
                        '//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>//

       '//CREO LA SHELL come oggetto, attivo il comando "%comspec% /c start " e gli assegno la path per estesa con il nome del file
            '// perchè Shell lancia un EXE e non il PDF, quindi è necessario costruire la stringa in modo da far lanciare
            '// prima la sessione dos e poi il pdf..
            '// La stringa di comando ("%comspec% /c start ") deve essere unica : comando shell + path + file
            'Set WshPDF = CreateObject("wscript.shell")
            
            '// Unisco il comando pdf start + path definitiva
            'Stringa1 = "%comspec% /c start " & MyPath_s & MyFile_s                               '//Comando di apertura
                '//Attivo il comando ed aggiungo @exit per @chiudere@il@terminale occorre lo spazio per evitare l'errore & " ^exit"
            'WshPDF.Run Stringa1 & " ^exit"
      
      
ApriFilePdf_Interna_s_PFunct_Exit:
    Exit Function

ApriFilePdf_Interna_s_PFunct_Err:
    MsgBox Error$
    Resume ApriFilePdf_Interna_s_PFunct_Exit

End Function

'//APRI_IL_FILE_PDF - FUNZIONE INTERNA DI APERTURA DEI FILE -         *** FINE ***
'//==================================================================================================================//





'//*** INIZIO ***
'//APRI FILE @APRO@ANALISI@SUBLIME.TEXT  @CMD@03@CMD.03@CMD.3    @HELP@SUBLIME@TEXT@ANLISI
'//========================================================================================================================================//

Private Sub Cmd_HELP_Click()

On Error GoTo Cmd_HELP_Err
    
      
      
      
                     
        
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
                '//CODICE               -> Function ApriFilePdf.01.01
                '//PARAMETRI            -> par_Path_s         = PATH _
                                        -> par_NameFile_s               = NOME FILE _
                                        -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                        
                        '//imposto i parametri - LA PATH
                        MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\"
                        '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
                        MyFile_s = "Project_MSYSDbEst_GE_OBJECT_ANALISI.sublime-project "
                
                            '//CHIAMO LA FUNZIONE INTERNA PER L'APERTURA DEI FILE
                            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
             
                        
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

Cmd_HELP_Exit:
    Exit Sub

Cmd_HELP_Err:
    MsgBox Error$
    Resume Cmd_HELP_Exit

End Sub


'//*** FINE ***
'//APRI PATH
'//========================================================================================================================================//







'//ESEGUI QUERY SULLA TABELLA DI SISTEMA MSysObjects PER INDIVIDUARE GLI OGGETTI ATTIVI (TABELLE, QUERY FORM REPORT)
'//========================================================================================================================================//
'//Recupero del percorso del database esterno: La funzione PathFile_s_pFunct viene chiamata per ottenere il percorso del database esterno. _
    Se il percorso è vuoto, la funzione mostra un messaggio di errore e esce. Apertura del database esterno: Il database esterno viene aperto _
    utilizzando DBEngine.Workspaces(0).OpenDatabase. Esecuzione della query su MSysObjects: Viene eseguita una query su MSysObjects per _
    ottenere tutti gli oggetti che non sono nascosti (Flags=0) e che sono di tipo tabella, query, form, modulo, report o macro. _
    Ciclo attraverso i risultati: La funzione cicla attraverso i risultati della query e stampa il nome e il tipo di ciascun _
    oggetto nella finestra di debug. La variabile objectTypeDescription viene utilizzata per fornire una descrizione leggibile del tipo di oggetto. _
    Chiusura del recordset e del database: Il recordset e il database esterno vengono chiusi al termine della funzione. _
    Messaggio informativo: Viene visualizzato un messaggio informativo per indicare che la lista degli oggetti è stata stampata nella finestra di debug.


Private Sub ListExternalDBObjects()
    Dim sourceDBPath As String
    Dim externalDB As DAO.Database
    Dim rs As DAO.Recordset
    Dim objectType As Integer
    Dim objectName As String
    Dim objectTypeDescription As String
    
    ' Liste per ogni tipo di oggetto
    Dim tables As String
    Dim queries As String
    Dim forms As String
    Dim modules As String
    Dim reports As String
    Dim macros As String
    Dim others As String
    
    ' Recupero la path e il file dal database esterno
    sourceDBPath = PathFile_s_pFunct
    
    If sourceDBPath = "" Or IsNull(sourceDBPath) = True Then
        MsgBox "ATTENZIONE ERRORE: path e file sono null, uscita dalla routine", vbCritical, "MSG CONTROLLO PATH E FILE VALORIZZATI"
        Exit Sub
    End If
    
    ' RESET
    Me.TXT_OBJECT_ESTERNI.Value = ""
    Me.TXT_OBJECT_ESTERNI.Requery
    
    ' Apro il database esterno
    Set externalDB = DBEngine.Workspaces(0).OpenDatabase(sourceDBPath)
    
    ' Eseguo la query su MSysObjects per ottenere tutti gli oggetti
    Set rs = externalDB.OpenRecordset("SELECT Name, Type FROM MSysObjects WHERE Flags=0 AND Type IN (1, 4, 5, -32761, -32764, -32766)")
    
    ' Ciclo attraverso i risultati e raccolgo gli oggetti nelle liste appropriate
    Do While Not rs.EOF
    
        objectName = rs!Name
        objectType = rs!Type
        
        
        '// Chr$(13) =  è una funzione in VBA che restituisce il carattere corrispondente al codice ASCII specificato. _
            In questo caso, 13 è il codice ASCII per il carattere di ritorno a capo (Carriage Return, CR), che viene spesso _
            utilizzato in combinazione con Chr$(10) (line feed, LF) per creare una nuova linea su molte piattaforme. Tuttavia, per creare una _
            nuova linea in modo più robusto e compatibile, è meglio usare vbCrLf, che combina entrambi i caratteri di controllo (Chr$(13) & Chr$(10)). _
            vbCrLf è una costante predefinita in VBA che rappresenta una nuova linea e garantisce che il testo appaia correttamente formattato _
            su tutte le piattaforme. _
            Chr$(10) = Il Line Feed (LF) è un carattere di controllo utilizzato per indicare la fine di una linea di testo e l'inizio di una nuova. _
            Nel sistema ASCII, il codice per il Line Feed è 10. Questo carattere è comunemente utilizzato nei sistemi Unix e Linux _
            per terminare le linee di testo. Quando un Line Feed viene incontrato in un testo, il cursore o il puntatore si _
            sposta all 'inizio della riga successiva. Tuttavia, il Line Feed da solo non riporta il cursore all'inizio della riga corrente, _
            per questo motivo spesso viene combinato con il Carriage Return (CR, codice ASCII 13), che invece riporta il _
            cursore all'inizio della riga corrente.0
            
        ' Descrivo il tipo di oggetto
        Select Case objectType
            Case 1
                objectTypeDescription = "Tabella"
                tables = tables & "Nome: " & objectName & ", Tipo: " & objectTypeDescription & " - codice : " & objectType & vbCrLf
            Case 4
                objectTypeDescription = "Query"
                queries = queries & "Nome: " & objectName & ", Tipo: " & objectTypeDescription & " - codice : " & objectType & vbCrLf
            Case 5
                objectTypeDescription = "Form"
                forms = forms & "Nome: " & objectName & ", Tipo: " & objectTypeDescription & " - codice : " & objectType & vbCrLf
            Case -32761
                objectTypeDescription = "Modulo"
                modules = modules & "Nome: " & objectName & ", Tipo: " & objectTypeDescription & " - codice : " & objectType & vbCrLf
            Case -32764
                objectTypeDescription = "Report"
                reports = reports & "Nome: " & objectName & ", Tipo: " & objectTypeDescription & " - codice : " & objectType & vbCrLf
            Case -32766
                objectTypeDescription = "Macro"
                macros = macros & "Nome: " & objectName & ", Tipo: " & objectTypeDescription & " - codice : " & objectType & vbCrLf
            Case Else
                objectTypeDescription = "Altro"
                others = others & "Nome: " & objectName & ", Tipo: " & objectTypeDescription & " - codice : " & objectType & vbCrLf
        End Select
        
        rs.MoveNext
    Loop
    
    ' Stampo le liste in ordine
    Debug.Print "Oggetti nel database esterno (" & sourceDBPath & "):"
    Debug.Print "-----------------------------------------------"
    Debug.Print
    
    If tables <> "" Then Debug.Print tables
    Debug.Print
    If queries <> "" Then Debug.Print queries
    Debug.Print
    If forms <> "" Then Debug.Print forms
    Debug.Print
    If modules <> "" Then Debug.Print modules
    Debug.Print
    If reports <> "" Then Debug.Print reports
    Debug.Print
    If macros <> "" Then Debug.Print macros
    Debug.Print
    If others <> "" Then Debug.Print others
    Debug.Print
    
    ' Imposto il valore del controllo di testo
    Me.TXT_OBJECT_ESTERNI.Value = tables & vbCrLf & queries & vbCrLf & forms & vbCrLf & modules & vbCrLf & reports & vbCrLf & macros & vbCrLf & others
    
    ' Chiudo il recordset e il database
    rs.Close
    externalDB.Close
    
    MsgBox "Lista degli oggetti nel database esterno stampata nella finestra di debug.", vbInformation, "Completato"
End Sub


'//*** FINE ***
'//ESEGUI QUERY SULLA TABELLA DI SISTEMA MSysObjects PER INDIVIDUARE GLI OGGETTI ATTIVI (TABELLE, QUERY FORM REPORT)
'//========================================================================================================================================//








'//=====================================================================================================================//
'//                     DELETE TUTTI GLI OGGETTI DEL DB ESTERNO IN TABELLA *** INIZIO ***
'//                     + DELETE TUTTO HELP
'//=====================================================================================================================//




'//CANCELLA TUTTO
Private Sub Cmd_DELETE_TUTTO_Click()
        
        
    
                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** INIZIO ***
                '//=====================================================================================================================//
                '//NOTE: chiamo la funzione che controlla se il progetto corrente è il progetto originale, infatti _
                         non si possono caricare gli oggetti nel PROGETTO MSYS_N01_OGGETTI.mdb perhè è originale _
                         e vi è una distruzione degli stessi. LA FUNZIONE CONTROLLO SE E' UN PROGETTO ESTERNO AMMISSIBILE _
                         per caricare gli oggetti table, query , form ecc. da importare. _
                         01) CONTROLLO = MSG DI AVVISO CON POSSIBILITA DI ANNULLARE _
                         02) CONTROLLO = CONTROLLO SE STIAMO NEL PROGETTO ORIGINALE  E QUINDI ESCLUSIONE.
                                 
                           '//01) PRIMO CONTROLLO _
                             MESSAGGIO DI IMPORTAZIONE 0 = ANNULLA  1 = CONTINUA
                           '//..........................................................//
                             '//Attenzione L'AVVISO produce il messaggio solo se il _
                                parametro è TRUE altrimenti il messaggio viene escluso.
                             MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)
                             
                             If MyValue = 0 Then Exit Sub
                           '//..........................................................//
                                 
                           '//02) SECONDO CONTROLLO
                           '//CHIAMO IL CONTROLLO DEL PROGETTO
                           '//..........................................................//
                           '//Note  : True = siamo nel progetto originale non possiamo _
                                      importare nessun oggetto perche distruggiamo il db corrente. _
                                      False = siamo in un progetto esterno possiamo caricare tutto.
                                          
                               '//SE VIENE RESTITUITO TRUE è FUORI PROGETTO PUO CARICARE, SE FALSE EXIT PERCHE PROGETTO ORIGINALE
                               Bool1 = CheckProjectName_b_PFunct
                               
                               If Bool1 = True Then Exit Sub
                           '//..........................................................//
                
                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** FINE ***
                '//=====================================================================================================================//
                    
                    '//2024.08.03_MODIFICA_CMB_DELETE_TUTTO = la modifica riguarda la routine di attivazione
                    '//ATTIVO LA CANCELLAZIONE DI TUTTI GLI OGGETTI SE OGNI  BUTTON E' ABILITATO CON TRUE '
                                            
                        If Me.Cmd_Delete_FORM.Enabled = True Then
                            Cmd_Delete_FORM_Click
                        End If
                        If Me.Cmd_Delete_MODULI.Enabled = True Then
                            Cmd_Delete_MODULI_Click
                        End If
                        If Me.Cmd_Delete_QUERY.Enabled = True Then
                            Cmd_Delete_QUERY_Click
                        End If
                        If Me.Cmd_Delete_REPORT.Enabled = True Then
                            Cmd_Delete_REPORT_Click
                        End If
                        If Me.Cmd_Delete_TABLE.Enabled = True Then
                            Cmd_Delete_TABLE_Click
                        End If
                        
                        If Me.Cmd_Delete_MACRO.Enabled = True Then
                            Cmd_Delete_MACRO_Click
                        End If

                        '//FINE CON MESSAGGIO IN CASELLA DI TESTO'
                        Me.TXT_OBJECT_ESTERNI = vbCrLf & vbCrLf & "  //********************************************************************// " & vbCrLf & " ......  FINE CANCELLAZIONE DI TUTTI GLI OGGETTI "
                        Me.TXT_OBJECT_ESTERNI.Requery
                        

End Sub


Private Sub Cmd_Delete_Tutto_HELP_Click()

    '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
    '//--------------------------------------------------------------------------------//--------//
    '//NOTE                 -> Apro il file di tipo doc, zip o pdf
    '//CODICE               -> Function ApriFilePdf.01.01
    '//PARAMETRI            -> par_Path_s         = PATH _
                            -> par_NameFile_s               = NOME FILE _
                            -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                            
      '//IMPOSTO LE VARIABILI PATH E FILE
      '//
      Dim MyFile_s As String
      '//
      Dim MyPath_s  As String
      
            '//imposto i parametri - LA PATH
            '
            MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\OBJECT\HELP\"
            '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
            '
            MyFile_s = "HELP_Cmd_DELETE_TUTTO.pdf "
                
                
                       
         '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
            Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto IMPORTA OGGETTI incluso in questa REPORT _
            anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'importazione e cancellazione in caso di resettamento _
            totale.
            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
      
            
    '//--------------------------------------------------------------------------------//--------//

End Sub






'//=====================================================================================================================//
'//                     DELETE TUTTI GLI OGGETTI DEL DB ESTERNO IN TABELLA *** FINE ***
'//                     + DELETE TUTTO HELP
'//=====================================================================================================================//






Private Sub Cmd_IMPORTA_TUTTO_Click()



                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** INIZIO ***
                '//=====================================================================================================================//
                '//NOTE: chiamo la funzione che controlla se il progetto corrente è il progetto originale, infatti _
                         non si possono caricare gli oggetti nel PROGETTO MSYS_N01_OGGETTI.mdb perhè è originale _
                         e vi è una distruzione degli stessi. LA FUNZIONE CONTROLLO SE E' UN PROGETTO ESTERNO AMMISSIBILE _
                         per caricare gli oggetti table, query , form ecc. da importare. _
                         01) CONTROLLO = MSG DI AVVISO CON POSSIBILITA DI ANNULLARE _
                         02) CONTROLLO = CONTROLLO SE STIAMO NEL PROGETTO ORIGINALE  E QUINDI ESCLUSIONE.
                                 
                           '//01) PRIMO CONTROLLO _
                             MESSAGGIO DI IMPORTAZIONE 0 = ANNULLA  1 = CONTINUA
                           '//..........................................................//
                             '//Attenzione L'AVVISO produce il messaggio solo se il _
                                parametro è TRUE altrimenti il messaggio viene escluso.
                             MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)
                             
                             If MyValue = 0 Then Exit Sub
                           '//..........................................................//
                                 
                           '//02) SECONDO CONTROLLO
                           '//CHIAMO IL CONTROLLO DEL PROGETTO
                           '//..........................................................//
                           '//Note  : True = siamo nel progetto originale non possiamo _
                                      importare nessun oggetto perche distruggiamo il db corrente. _
                                      False = siamo in un progetto esterno possiamo caricare tutto.
                                          
                               '//SE VIENE RESTITUITO TRUE è FUORI PROGETTO PUO CARICARE, SE FALSE EXIT PERCHE PROGETTO ORIGINALE
                               Bool1 = CheckProjectName_b_PFunct
                               
                               If Bool1 = True Then Exit Sub
                           '//..........................................................//
                
                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** FINE ***
                '

                '//CHIAMO LE PROCEDURE MA LE TABELLE DEVONO ESSERE CHIAMATE PER PRIMA _
                   e le query per ultimo, altrimenti le query non si definiscono benei i campi.'
                
                '//2024.08.03_MODIFICA_CMB_IMPORTA_TUTTO = la modifica riguarda la routine di attivazione
                '//ATTIVO LA IMPORTAZIONE DI TUTTI GLI OGGETTI SE OGNI  BUTTON E' ABILITATO CON TRUE '
                    
                    If Me.Cmd_Importa_TABLE.Enabled = True Then
                        Cmd_Importa_TABLE_Click
                    End If
                    
                    If Me.Cmd_Importa_FORM.Enabled = True Then
                        Cmd_Importa_FORM_Click
                    End If
                    
                    If Me.Cmd_Importa_MODULI.Enabled = True Then
                        Cmd_Importa_MODULI_Click
                    End If
                    
                    If Me.Cmd_Importa_REPORT.Enabled = True Then
                        Cmd_Importa_REPORT_Click
                    End If
                    
                    If Me.Cmd_Importa_MACRO.Enabled = True Then
                        Cmd_Importa_MACRO_Click
                    End If
                    
                    If Me.Cmd_Importa_QUERY.Enabled = True Then
                        Cmd_Importa_QUERY_Click
                    End If


                     '//FINE CON MESSAGGIO IN CASELLA DI TESTO'
                        Me.TXT_OBJECT_ESTERNI = vbCrLf & vbCrLf & "  //********************************************************************// " & vbCrLf & " ......  FINE IMPORTAZIONE DI TUTTI GLI OGGETTI "
                        Me.TXT_OBJECT_ESTERNI.Requery
                       

End Sub




Private Sub Cmd_IMPORTA_Tutto_HELP_Click()

    '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
    '//--------------------------------------------------------------------------------//--------//
    '//NOTE                 -> Apro il file di tipo doc, zip o pdf
    '//CODICE               -> Function ApriFilePdf.01.01
    '//PARAMETRI            -> par_Path_s         = PATH _
                            -> par_NameFile_s               = NOME FILE _
                            -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                            
      '//IMPOSTO LE VARIABILI PATH E FILE
      '//
      Dim MyFile_s As String
      '//
      Dim MyPath_s  As String
      
            '//imposto i parametri - LA PATH
            '
            MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\OBJECT\HELP\"
            '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
            '
            MyFile_s = "HELP_Cmd_IMPORTA_TUTTO.pdf "
                
                
                       
         '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
            Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto IMPORTA OGGETTI incluso in questa REPORT _
            anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'importazione e cancellazione in caso di resettamento _
            totale.
            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
      
            
    '//--------------------------------------------------------------------------------//--------//

End Sub



'//*********************************************************************************************//
'//***
'//***                                            VARIE FUNZIONI *** FINE ***
'//*********************************************************************************************//












'//***********************************************************************************************************************//
'//                             @ABILITAZIONE E @DISABILLITAZIONI        *** INIZIO *****
'//
'//***********************************************************************************************************************//




'//DISABILITO LA PATH DEL DB
Private Sub Disabilito_sourceDBPath_Sub()
    
    '//DISABILITO LA PATH ESTERNA
    '//........................................//
       'Call DisabilitosourceDBPath_s_Txt_Sub
    '//........................................//
    
    '//disabilito la path
    Me.sourceDBPath_s_Txt.Enabled = False
    
    
End Sub

'//ABILITO LA PATH DEL DB
Private Sub Abilito_sourceDBPath_Sub()
    
    '//ABILITO LA PATH ESTERNA
    '//........................................//
       'Call Abilito_sourceDBPath_Sub
    '//........................................//
    
    '//Abilito la path
    Me.sourceDBPath_s_Txt.Enabled = True
    
    
End Sub




Private Sub DisabilitaButton_Sub()
    
    
    '//DISABILITO I BUTTON
    '//........................................//
       'Call DisAbilitaButton_Sub
    '//........................................//
    
    
    '//@DISABILITO@BUTTON_(disabilito tutti i comandi button se ci troviano nel progetto originale MSYS_N01_OGGETTI.mdb)
    Me.Cmd_Delete_FORM.Enabled = False
    Me.Cmd_Delete_MACRO.Enabled = False
    Me.Cmd_Delete_MODULI.Enabled = False
    Me.Cmd_Delete_QUERY.Enabled = False
    Me.Cmd_Delete_REPORT.Enabled = False
    Me.Cmd_Delete_TABLE.Enabled = False
    Me.Cmd_DELETE_TUTTO.Enabled = False
    Me.Cmd_Importa_FORM.Enabled = False
    Me.Cmd_Importa_MACRO.Enabled = False
    Me.Cmd_Importa_MODULI.Enabled = False
    Me.Cmd_Importa_QUERY.Enabled = False
    Me.Cmd_Importa_REPORT.Enabled = False
    Me.Cmd_Importa_TABLE.Enabled = False
    Me.Cmd_IMPORTA_TUTTO.Enabled = False
    
    
End Sub


'//ABILITAZIONE@BUTTON
'//NOTE : se corrisponde a diverse voci della combinata abilita il button
Private Sub AbilitaButton_Sub()
    
    '//ABILITO I BUTTON
    '//........................................//
    '//note : chiamo la funzione di abilitazione _
        dei button se corrispone ad una voce diversa _
        da ELENCO PROGETTI
       'Call AbilitaButton_Sub
    '//........................................//
 
            '//@ABILITO@BUTTON_(ABILITO tutti i comandi button se ci troviano nel progetto originale MSYS_N01_OGGETTI.mdb)
            Me.Cmd_Delete_FORM.Enabled = True
            Me.Cmd_Delete_MACRO.Enabled = True
            Me.Cmd_Delete_MODULI.Enabled = True
            Me.Cmd_Delete_QUERY.Enabled = True
            Me.Cmd_Delete_REPORT.Enabled = True
            Me.Cmd_Delete_TABLE.Enabled = True
            Me.Cmd_DELETE_TUTTO.Enabled = True
            Me.Cmd_Importa_FORM.Enabled = True
            Me.Cmd_Importa_MACRO.Enabled = True
            Me.Cmd_Importa_MODULI.Enabled = True
            Me.Cmd_Importa_QUERY.Enabled = True
            Me.Cmd_Importa_REPORT.Enabled = True
            Me.Cmd_Importa_TABLE.Enabled = True
            Me.Cmd_IMPORTA_TUTTO.Enabled = True
            
            
                
            '//ABILITO ANCHE LA COMBINATA 2 = CMB_02
            Me.Cmb_02_txt.Enabled = False
            Me.Cmb_02_txt.Value = "scelta"

    
End Sub





'//ABILITAZIONE E DIABLITAIONE DEI BUTTON , TABLE, FORM, QUERY MACRO REPORT *** INIZIO ***
'//-------------------------------------------------------------------------------------------------------------//
'2024.08.03_Abilito_Table+Report+Macro+Query+Moduli = creato routine per
'l 'abilitazione delle query, macro, report, table ecc.
'2024.08.03_DisabilitoAbilito_Table+Report+Macro+Query+Moduli = creato routine per
'la DISABILITAZIONE  delle query, macro, report, table ecc.

'//ABILITAZIONE@BUTTON
'//NOTE : ABILITA SOLO 1 COMANDO IMPORTA ED 1 CANCELLA
Private Sub AbilitaButton_FORM_Sub()
    
            '//ABILITO SOLO I COMANDI, IMPORTA E CANCELLA FORM
            '//@ABILITO@BUTTON_(abilito il @comando@form @cmd@form@delete.abilitato + cmd@form@importa.abilitato)
            Me.Cmd_Delete_FORM.Enabled = True
            Me.Cmd_Importa_FORM.Enabled = True
          
End Sub



'//DISABILITAZIONE@BUTTON@CMB_02
'//NOTE : DISABILITA LA COMBINATA 02
Private Sub DisabilitaButton_CMB_02_Sub()
    
            '//DISABILITO SOLO I COMANDI, IMPORTA TUTTO
            '//@DISABILITO@BUTTON_(DISABILITO il @comando@CMB_02)
            Me.Cmb_02_txt.Enabled = False
          
End Sub


'//BILITAZIONE@BUTTON@CMB_02
'//NOTE : ABILITA LA COMBINATA 02
Private Sub AbilitaButton_CMB_02_Sub()
    
            '//ABILITO SOLO I COMANDI, IMPORTA TUTTO
            '//@DISABILITO@BUTTON_(ABILITO il @comando@CMB_02)
            Me.Cmb_02_txt.Enabled = True
          
End Sub




'//DISABILITAZIONE@BUTTON
'//NOTE : DISABILITA SOLO 1 COMANDO IMPORTA ED 1 CANCELLA
Private Sub DisabilitaButton_FORM_Sub()
    
            '//DISABILITO SOLO I COMANDI, IMPORTA E CANCELLA FORM
            '//@DISABILITO@BUTTON_(DISABILITO il @comando@form @cmd@form@delete.disabilitato + cmd@form@importa.disabilitato)
            Me.Cmd_Delete_FORM.Enabled = False
            Me.Cmd_Importa_FORM.Enabled = False
          
End Sub


'//ABILITAZIONE@BUTTON
'//NOTE : ABILITA SOLO 1 COMANDO IMPORTA ED 1 CANCELLA
Private Sub AbilitaButton_TABLE_Sub()
    
            '//ABILITO SOLO I COMANDI, IMPORTA E CANCELLA TABLE
            '//@ABILITO@BUTTON_(abilito il @comando@TABLE @cmd@TABLE@delete.abilitato + cmd@TABLE@importa.abilitato)
            Me.Cmd_Delete_TABLE.Enabled = True
            Me.Cmd_Importa_TABLE.Enabled = True
          
End Sub



'//DISABILITAZIONE@BUTTON
'//NOTE : DISABILITA SOLO 1 COMANDO IMPORTA ED 1 CANCELLA
Private Sub DisabilitaButton_TABLE_Sub()
    
            '//DISABILITO SOLO I COMANDI, IMPORTA E CANCELLA TABLE
            '//@DISABILITO@BUTTON_(DISABILITO il @comando@TABLE @cmd@TABLE@delete.disabilitato + cmd@TABLE@importa.disabilitato)
            Me.Cmd_Delete_TABLE.Enabled = False
            Me.Cmd_Importa_TABLE.Enabled = False
          
End Sub




'//ABILITAZIONE@BUTTON
'//NOTE : ABILITA SOLO 1 COMANDO IMPORTA ED 1 CANCELLA
Private Sub AbilitaButton_QUERY_Sub()
    
            '//ABILITO SOLO I COMANDI, IMPORTA E CANCELLA QUERY
            '//@ABILITO@BUTTON_(abilito il @comando@QUERY @cmd@QUERY@delete.abilitato + cmd@QUERY@importa.abilitato)
            Me.Cmd_Delete_QUERY.Enabled = True
            Me.Cmd_Importa_QUERY.Enabled = True
          
End Sub



'//DISABILITAZIONE@BUTTON
'//NOTE : DISABILITA SOLO 1 COMANDO IMPORTA ED 1 CANCELLA
Private Sub DisabilitaButton_QUERY_Sub()
    
            '//DISABILITO SOLO I COMANDI, IMPORTA E CANCELLA QUERY
            '//@DISABILITO@BUTTON_(DISABILITO il @comando@QUERY @cmd@QUERY@delete.disabilitato + cmd@QUERY@importa.disabilitato)
            Me.Cmd_Delete_QUERY.Enabled = False
            Me.Cmd_Importa_QUERY.Enabled = False
          
End Sub


'//ABILITAZIONE@BUTTON
'//NOTE : ABILITA SOLO 1 COMANDO IMPORTA ED 1 CANCELLA
Private Sub AbilitaButton_MODULI_Sub()
    
            '//ABILITO SOLO I COMANDI, IMPORTA E CANCELLA MODULI
            '//@ABILITO@BUTTON_(abilito il @comando@MODULI @cmd@MODULI@delete.abilitato + cmd@MODULI@importa.abilitato)
            Me.Cmd_Delete_MODULI.Enabled = True
            Me.Cmd_Importa_MODULI.Enabled = True
          
End Sub



'//DISABILITAZIONE@BUTTON
'//NOTE : DISABILITA SOLO 1 COMANDO IMPORTA ED 1 CANCELLA
Private Sub DisabilitaButton_MODULI_Sub()
    
            '//DISABILITO SOLO I COMANDI, IMPORTA E CANCELLA MODULI
            '//@DISABILITO@BUTTON_(DISABILITO il @comando@MODULI @cmd@MODULI@delete.disabilitato + cmd@MODULI@importa.disabilitato)
            Me.Cmd_Delete_MODULI.Enabled = False
            Me.Cmd_Importa_MODULI.Enabled = False
          
End Sub



'//ABILITAZIONE@BUTTON
'//NOTE : ABILITA SOLO 1 COMANDO IMPORTA ED 1 CANCELLA
Private Sub AbilitaButton_REPORT_Sub()
    
            '//ABILITO SOLO I COMANDI, IMPORTA E CANCELLA REPORT
            '//@ABILITO@BUTTON_(abilito il @comando@REPORT @cmd@REPORT@delete.abilitato + cmd@REPORT@importa.abilitato)
            Me.Cmd_Delete_REPORT.Enabled = True
            Me.Cmd_Importa_REPORT.Enabled = True
          
End Sub


'//ABILITAZIONE@BUTTON
'//NOTE : ABILITA SOLO 1 COMANDO IMPORTA ED 1 CANCELLA
Private Sub AbilitaButton_MACRO_Sub()
    
            '//ABILITO SOLO I COMANDI, IMPORTA E CANCELLA MACRO
            '//@ABILITO@BUTTON_(abilito il @comando@MACRO @cmd@MACRO@delete.abilitato + cmd@MACRO@importa.abilitato)
            Me.Cmd_Delete_MACRO.Enabled = True
            Me.Cmd_Importa_MACRO.Enabled = True
          
End Sub







'//DISABILITAZIONE@BUTTON
'//NOTE : DISABILITA SOLO 1 COMANDO IMPORTA ED 1 CANCELLA
Private Sub DisabilitaButton_MACRO_Sub()
    
            '//DISABILITO SOLO I COMANDI, IMPORTA E CANCELLA MACRO
            '//@DISABILITO@BUTTON_(DISABILITO il @comando@MACRO @cmd@MACRO@delete.disabilitato + cmd@MACRO@importa.disabilitato)
            Me.Cmd_Delete_MACRO.Enabled = False
            Me.Cmd_Importa_MACRO.Enabled = False
          
End Sub




'//DISABILITAZIONE@BUTTON
'//NOTE : DISABILITA SOLO 1 COMANDO IMPORTA ED 1 CANCELLA
Private Sub DisabilitaButton_REPORT_Sub()
    
            '//DISABILITO SOLO I COMANDI, IMPORTA E CANCELLA REPORT
            '//@DISABILITO@BUTTON_(DISABILITO il @comando@REPORT @cmd@REPORT@delete.disabilitato + cmd@REPORT@importa.disabilitato)
            Me.Cmd_Delete_REPORT.Enabled = False
            Me.Cmd_Importa_REPORT.Enabled = False
          
End Sub




'//DISABILITAZIONE@BUTTON
'//NOTE : ABILITA SOLO 1 COMANDO IMPORTA ED 1 CANCELLA
Private Sub AbilitaButton_IMPORTA_TUTTO_Sub()
    
            '//ABILITO SOLO I COMANDI, IMPORTA TUTTO
            '//@DISABILITO@BUTTON_(ABILITO il @comando@IMPORTA@TUTTO)
            Me.Cmd_IMPORTA_TUTTO.Enabled = True
            Me.Cmd_DELETE_TUTTO.Enabled = True
            
End Sub



'//DISABILITAZIONE@BUTTON
'//NOTE : DISABILITA SOLO 1 COMANDO IMPORTA TUTTO
Private Sub DisabilitaButton_IMPORTA_TUTTO_Sub()
    
            '//DISABILITO SOLO I COMANDI, IMPORTA TUTTO
            '//@DISABILITO@BUTTON_(DISABILITO il @comando@IMPORTA@TUTTO)
            Me.Cmd_IMPORTA_TUTTO.Enabled = False
            
End Sub


'//DISABILITAZIONE@BUTTON@DELETE@TUTTO
'//NOTE : DISABILITA SOLO 1 COMANDO DELETE TUTTO
Private Sub DisabilitaButton_DELETE_TUTTO_Sub()
    
            '//DISABILITO SOLO I COMANDI, IMPORTA TUTTO
            '//@DISABILITO@BUTTON_(DISABILITO il @comando@IMPORTA@TUTTO)
            Me.Cmd_DELETE_TUTTO.Enabled = False
          
End Sub



'//ABILITAZIONE@BUTTON@DELETE@TUTTO
'//NOTE : ABILITA SOLO 1 COMANDO DELETE TUTTO
Private Sub AbilitaButton_DELETE_TUTTO_Sub()
    
            '//ABILITO SOLO I COMANDI, IMPORTA TUTTO
            '//@DISABILITO@BUTTON_(DISABILITO il @comando@IMPORTA@TUTTO)
            Me.Cmd_DELETE_TUTTO.Enabled = True
          
End Sub


'//ABILITAZIONE E DIABLITAIONE DEI BUTTON , TABLE, FORM, QUERY MACRO REPORT *** INIZIO ***
'//-------------------------------------------------------------------------------------------------------------//




'//@ABILITA@COMBINATA_01=@CMB_02
Private Sub AbilitaCmd_02_Sub()
    'Me.Cmb_01_txt.Selected
    Me.Cmb_02_txt.Enabled = True
    Me.Cmb_02_txt.Value = "ELENCO OGGETTI MSYS"

End Sub






'//***********************************************************************************************************************//
'//                             @ABILITAZIONE E @DISABILLITAZIONI        *** FINE *****
'//
'//***********************************************************************************************************************//




'//***********************************************************************************************************************//
'//                             @STAMPE@OGGETTI@CASELLA@DI@TESTO        *** INIZIO *****
'//
'//***********************************************************************************************************************//


'//@STAMPA@OGGETTI@MSYS
'//================================================================================================================//
'// STAMPA OGGETTI COLLETTIVA *** INIZIO ***
'//================================================================================================================//

Private Function Stampa_MSYS_N01_OGGETTI_funct(par_LaCollection As Collection, par_Tipo_s As String) As String


'//DIM della collezione form da importare
Dim formsCollection As Collection
Dim formName_v As Variant

'//DIM la collezione macro
Dim macroCollection As Collection
Dim macroName_v As Variant



'//DIM la collezione DEI MODULI
Dim modulesCollection As Collection
Dim moduloName_v As Variant


'//DIM la collezione DELLE QUERY
Dim queryCollection As Collection
Dim queryName_v As Variant




'//DIM la collezione DELLE REPORT
Dim reportCollection As Collection
Dim reportName_v As Variant



'//DIM la collezione DELLE TABELLE
Dim tableCollection As Collection
Dim tableName_v As Variant


'//Le variabili di controllo della sub
Dim ProceduraMessaggioErrore_s As String
Dim ProceduraAttivaEseguita_s As String
   



    '//....
On Error GoTo Err_Stampa_MSYS_N01_OGGETTI_funct


'//RESET VARIABILI
ProceduraMessaggioErrore_s = "PROCEDURA_ESEGUITA_Stampa_MSYS_N01_OGGETTI_funct"
ProceduraAttivaEseguita_s = "ATTIVITA_DI_CONTROLLO_controllo_record"


'//LA_SUB_Stampa_MSYS_N01_OGGETTI_funct_01.01
'//______________________________________________________________________________
            

            
                 
                '//CREO IL MESSAGGIO DA VISUALIZZARE NELLA CASELLA DI TESTO UNITAMENTE AGLI OGGETTI DELLA COLLEZIONE
                '//---------------------------------------------------------------------------------------------------------//
                '//NOTE: 01_TESTATA_MESSAGGIO = cre la testata del messaggio da riportare nella casella di testo
                ProceduraAttivaEseguita_s = ""
                ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " ************************* PROGETTO DA CARICARE : " & PROGETTO_s & " ********************" & vbCrLf & vbCrLf
                ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " PATH DEL PROGETTO: " & sourceDBPath_s & vbCrLf
                
                
                
                ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & vbCrLf
                ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " ..................................................................... " & vbCrLf
                ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & "         LE COLLEZIONI FORM DA CARICARE  " & vbCrLf
                ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " ..................................................................... " & vbCrLf
                
                
                Select Case par_Tipo_s
                
                    Case "COLLECTION_FORMS"
                
                    '//CHIAMO LA COLLEZIONE FORM ED ITERO NEGLI OGGETTI FORM DA IMPORTARE
                    '//-------------------------------------------------------------------------------------------//
                    '//Note: Chiamo la funzione in cui vengo precaricate le form da importare in una _
                    collection che viene restituita come oggetto dalla funzione e impostata _
                    nella variabile oggetto formsCollection.
                    
                        ' Chiama la funzione per ottenere la collezione di form
                        Set formsCollection = par_LaCollection
                        
                        '//ciclo nella collezione per il CARICAMENTO DELLE FORM NEL STRINGA MESSAGGIO
                        For Each formName_v In formsCollection
                            
                            ' Aggiungi l'oggetto della collezione form nella stringa di messaggio
                            ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & vbCrLf & formName_v
                            
                            Str1 = Str1 & objfrm & vbCrLf

                            '//CONTROLLO
                            Debug.Print "Form Della collezione: " & ProceduraAttivaEseguita_s
                    
                        Next formName_v
                        
                                                   
                            
                           
                        '//VISUALIZZO NELLA CASELLA DI TESTO IL TITOLO E LE FORM DA CARICARE
                        Me.TXT_OBJECT_ESTERNI = ProceduraAttivaEseguita_s & vbCrLf
                            
                        Me.TXT_OBJECT_ESTERNI.Requery
                      
                    '//-------------------------------------------------------------------------------------------//
                        
                        
                Case "COLLECTION_MACROS"
                        
                    ' ATTIVA LA COLLEZIONE DELLE MACRO
                    '//------------------------------------------------------------------------//
                
                        Set macroCollection = CollectionMacro_PFunct()
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & vbCrLf & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " ..................................................................... " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & "         LE COLLEZIONI MACRO DA CARICARE  " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " ..................................................................... " & vbCrLf
                        
                         '//ciclo nella collezione per il CARICAMENTO DELLE MACRO NEL STRINGA MESSAGGIO
                        For Each macroName_v In macroCollection
                            Debug.Print "MACRO nella collezione: " & macroName_v
                            ' Aggiungi l'oggetto della collezione MACRO nella stringa di messaggio
                            ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & vbCrLf & macroName_v
                            
                        Next macroName_v
                        
                        
                           
                        '//VISUALIZZO NELLA CASELLA DI TESTO IL TITOLO E LE MACRO DA CARICARE
                        Me.TXT_OBJECT_ESTERNI = ProceduraAttivaEseguita_s & vbCrLf

                        
                        
                    '//------------------------------------------------------------------------//

                Case "COLLECTION_MODULES"
                    ' ATTIVA LA COLLEZIONE DEI MODULI
                    '//------------------------------------------------------------------------//
                
                        ' Recupero la collezione moduli da importare
                        Set modulesCollection = CollectionModuli_PFunct()
                        
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & vbCrLf & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " ..................................................................... " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & "         LE COLLEZIONI DEI MODULI DA CARICARE  " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " ..................................................................... " & vbCrLf
                        
                        
                         '//ciclo nella collezione per il CARICAMENTO DELLE MACRO NEL STRINGA MESSAGGIO
                        For Each moduloName_v In modulesCollection
                            Debug.Print "MODULO nella collezione: " & macroName_v
                            ' Aggiungi l'oggetto della collezione MODULO nella stringa di messaggio
                            ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & vbCrLf & moduloName_v
                            
                            
                        '//VISUALIZZO NELLA CASELLA DI TESTO IL TITOLO I MODULI DA CARICARE
                        Me.TXT_OBJECT_ESTERNI = ProceduraAttivaEseguita_s & vbCrLf

                        Next moduloName_v
                        
                    '//------------------------------------------------------------------------//
                    
                    
                    
                Case "COLLECTION_QUERYES"
                    
                    '// ATTIVA LA COLLEZIONE DELLE QUERY
                    '//------------------------------------------------------------------------//
                
                        ' Recupero la collezione moduli da importare
                        Set queryCollection = CollectionQuery_PFunct()
                        
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & vbCrLf & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " ..................................................................... " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & "         LE COLLEZIONI DELLE QUERY DA CARICARE  " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " ..................................................................... " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & vbCrLf & vbCrLf
                        
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " Attenzione per le query viene prima esaminata tutti gli oggetti che sono   " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " salvati nel database esterno, sul quale viene asperta una interrogazione   " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " sql che filtra dal database esterno e dalla tabella di sistema denominata  " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " msysobject nel database MSYS_N01_OGGETTI.mdb. La sql di estrazione applica " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " due filtri;                                                                " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " 01 filtro = esclude l'estrazione dei nomi che iniziano                     " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & "  ~sq = solo le tabelle cancellate.                                         " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " 02 filtro = escludo dalla collection tali query e quindi vengono estratte  " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " solo le tabelle che sono in lavorazione.                                   " & vbCrLf
                        
                        
                       
                        
                         '//ciclo nella collezione per il CARICAMENTO DELLE QUERY NEL STRINGA MESSAGGIO
                        For Each queryName_v In queryCollection
                            
                            Debug.Print "QUERY nella collezione: " & queryName_v
                            
                            'Aggiungi l'oggetto della collezione  QUERY nella stringa di messaggio
                            ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & vbCrLf & queryName_v
                            
                        Next queryName_v
                                                            
                                    
                        '//VISUALIZZO NELLA CASELLA DI TESTO IL TITOLO LE QUERY DA CARICARE
                        Me.TXT_OBJECT_ESTERNI = ProceduraAttivaEseguita_s & vbCrLf
                        
                    '//------------------------------------------------------------------------//
                    
                    
                    
                Case "COLLECTION_REPORTS"
                
                    '// ATTIVA LA COLLEZIONE DELLE REPORT
                    '//------------------------------------------------------------------------//
                
                        ' Recupero la collezione moduli da importare
                        Set reportCollection = CollectionReport_PFunct()
                        
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & vbCrLf & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " ..................................................................... " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & "         LE COLLEZIONI DELLE REPORT DA CARICARE  " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " ..................................................................... " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & vbCrLf
                           
                         '//ciclo nella collezione per il CARICAMENTO DELLE REPORT NEL STRINGA MESSAGGIO
                        For Each reportName_v In reportCollection
                            
                            Debug.Print "REPORT nella collezione: " & reportName_v
                            
                            'Aggiungi l'oggetto della collezione  REPORT nella stringa di messaggio
                            ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & vbCrLf & reportName_v
                    
                        Next reportName_v
                        
                            
                        '//VISUALIZZO NELLA CASELLA DI TESTO IL TITOLO LE REPORT DA CARICARE
                        Me.TXT_OBJECT_ESTERNI = ProceduraAttivaEseguita_s & vbCrLf
                    
                    
                    '//------------------------------------------------------------------------//

            
            Case "COLLECTION_TABLES"
                    '// ATTIVA LA COLLEZIONE DELLE TABELLE
                    '//------------------------------------------------------------------------//
                
                        ' Recupero la collezione moduli da importare
                        Set tableCollection = CollectionTable_PFunct()
                        
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & vbCrLf & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " ..................................................................... " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & "         LE COLLEZIONI DELLE TABELLE DA CARICARE  " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " ..................................................................... " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & vbCrLf
                           
                         '//ciclo nella collezione per il CARICAMENTO DELLE TABELLE NEL STRINGA MESSAGGIO
                        For Each tableName_v In tableCollection
                            
                            Debug.Print "TABELLE nella collezione: " & tableName_v
                            
                            'Aggiungi l'oggetto della collezione  TABELLE nella stringa di messaggio
                            ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & vbCrLf & tableName_v
                    
                        Next tableName_v
                        
                        
                        '//VISUALIZZO NELLA CASELLA DI TESTO IL TITOLO LE TABELLE DA CARICARE
                        Me.TXT_OBJECT_ESTERNI = ProceduraAttivaEseguita_s & vbCrLf
                    '//------------------------------------------------------------------------//
                    
                Case Else
                        MsgBox "NESSUNA COLLECTION E' STATA STAMPATA PERCHE NON ESISTE", vbCritical, "CONTROLLO DI STAMPA"
                    
                End Select
                              
                    '//RESTITUISCO AL CHIAMATE L'INTERA STAMPA
                    
                    Stampa_MSYS_N01_OGGETTI_funct = ProceduraAttivaEseguita_s
            
          
'//______________________________________________________________________________




'//USCITA  E GESTIONE ERRORI
'//..............................................................................................................


Exit_Stampa_MSYS_N01_OGGETTI_funct:
    Exit Function

Err_Stampa_MSYS_N01_OGGETTI_funct:
    MsgBox Err.Description
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Stop
    Resume Exit_Stampa_MSYS_N01_OGGETTI_funct

End Function


'//================================================================================================================//
'// STAMPA OGGETTI COLLETTIVA *** FINE ***
'//================================================================================================================//


