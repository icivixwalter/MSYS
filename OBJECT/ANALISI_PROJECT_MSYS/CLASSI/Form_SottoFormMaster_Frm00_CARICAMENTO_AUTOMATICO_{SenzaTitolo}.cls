VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_SottoFormMaster_Frm00_CARICAMENTO_AUTOMATICO_{SenzaTitolo}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<7<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'                                   LE VARIABILI GENERALI
'//LA FORM MODELLO MASTER
'
'//CODICE-----------> \label{form:ModelloMasterTabcontroll-H14L35}
'//le librerie necessarie: _
        1) Microsoft Dao 3.6 Object Library _
        2) Microsoft Office 9.0 Object Library _
        3) Microsoft Activex Data Objects 2.1 Library _
        4) Microsoft Access 9.0 Objects Library _
        @librerie.necessarie _
        @le.librerie _
        @ddl.di.funzionamento
        
'//FORM MASTER _
    @SOTTOFORM@MASTER@MSYS_(form master a caricamento automatico)
'//
'//FORM MASTER PER LA GESTIONE DI TUTTI I PROGETTI = si trova in questa path ed in questo progetto : c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSYS_N01_OGGETTI.mdb
'//
'//@FORM@GE@MASTER_(@FORM @MASTER PER LA GESTIONE DEI @PROGETTI)
'//@FORM@GE@MASTER_(@SOTTOFORM@MASTER PER IL @CARICAMENTO@AUTOMATICO DELE SOTTOFORM)
'//
'//
'// FAQ
'//     QUALE E' LA FORM MASTER DI GESTIONE DELLE FORM
'//     FORM MASTER A CARICAMENTO AUTOMANTICO
'//     LA FORM MODELLO PER IL CARICAMENTO AUTOMATICO
'//     @MODELLO@SOTTOFORM@MASTER_(PER IL CARICAMENTO AUTOMATICO @DELLE SOTTOFORM)
'//
'// AGGIORNAMENTI
'//     2024_09_04_SOTTOFORM_MASTER_SENZA_TITOLO = CREATA LA SOTTOFORM MASTER A CARICAMENTO AUTOMATICO DI 12 PAGINE DEL TAB CONTROLL
'//
'//
'//
'//
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<


'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'Option

    Option Compare Text
    Option Explicit




'//DISCO ED IL FILE ATTO
'//....................................................................

    Dim Disco_s As String
    Dim PathFileAtto_s As String
    Dim MyFileEstensione_s As String
    Dim FileAtto_s As String
    Dim WshPDF As Variant

'//....................................................................



'Variabili di database
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'DAO
    Dim daoDB As DAO.Database          '//occorre il riferimento alla libreria  Microsoft DAO 3.6 Object Libray
    Dim DaoWks As DAO.Workspace
    Dim daoRS As DAO.Recordset
    Dim DaoRs_Filtro As DAO.Recordset
    
    'ADO
    Dim ADODB As Database
    Dim AdodaoRs As Recordset
    
    'Contatori
    Dim iCount As Integer
    Dim dbl_count As Double
    Dim i As Integer
    
    'Le variabili generiche
    Dim sSql As String                                          ' Stringa di estrazione
    
    'Variabili generali
    Dim Str1 As String, Str2 As String, Str3 As String
    Dim Int1 As Integer
    Dim Date1 As Date
    Dim Lng1 As Long
    Dim Dbl1 As Double
    Dim Bool1 As Boolean
    Dim Vv1 As Variant, vV2 As Variant, vV3 As Variant
    Dim obj1 As Object
    Dim Lng As Long

    '................................
    ' DIM LE VARIABILI LOCALI GENERALI
    Dim sS1 As String
    Dim Boolean1 As Boolean
    Dim iInt1    As Integer
    Dim dblDBL1   As Double
    Dim objObject1 As Object
    



'Gestione parametri comandi
    Dim sxParamCmd_03 As String


'Larghezza e numero di colonna
    Dim sLarg_Col As String
    Dim iNum_Col As Integer


'Variabili DELLA FORM
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::


    Dim sxFrm_CHIAMANTE_GENERALE As String                  '... Form CHIAMANTE generale
    Dim sxFrm_CHIAMANTE_GENERALE_CORRENTE As String         '... Form CHIAMANTE generale OLD
    Dim sxFrm_CHIAMANTE_GENERALE_PRECEDENTE As String       '... Form CHIAMANTE precedente


    'Variabili della Casella combinata
    Dim sxCmb_01  As String
    Dim sxCmb_02  As String
    Dim sxCmb_03  As String
    Dim sxCmb_04  As String

    '//@VARIABILI.CARICA.PAGINE, @pagine.da.caricare.variabili
    
        Dim SottoForm_s As String
        Dim NameSottoForm_s As String
        Dim VisibleSottoForm_b As Boolean
        Dim LabelPagina_s As String

'

'COMANDI FORM
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::


    'GLI OGGETTI COMANDI FORM
    '--------------------------------------------------------------------
    'Oggetti generali della form
    Dim sxCODOGGETTO As String                                          'Il codice form
    Dim sxCODOGGETTO_2 As String                                        'Il codice form
    
    'ATTIVITA CMD_02
    Dim sxOpenFORM_01_Cmd_02 As String                                   'Stringa per l'oggetto Docmd Apri form per il Cmd_02
    Dim sxOpenREPORT_01_Cmd_02 As String                                 'Stringa per l'oggetto Docmd Apri Report per il Cmd_02
    
    Dim sxOpenQUERY_01_Cmd_02 As String                                  'Stringa per l'oggetto Docmd Apri Query 1 per il Cmd_02
    Dim sxOpenQUERY_02_Cmd_02 As String                                  'Stringa per l'oggetto Docmd Apri Query 2 per il Cmd_02
    Dim sxOpenQUERY_03_Cmd_02 As String                                  'Stringa per l'oggetto Docmd Apri Query 3 per il Cmd_02
    
    Dim sxEseguiRoutine_01_Cmd_02 As String                              'Stringa per l'oggetto Docmd Esegui Macro 1 per il Cmd_02
    
    
    
    'ATTIVITA CMD_03
    Dim sxOpenFORM_01_Cmd_03 As String                                   'Stringa per l'oggetto Docmd Apri form per il Cmd_03
    Dim sxOpenREPORT_01_Cmd_03 As String                                 'Stringa per l'oggetto Docmd Apri Report per il Cmd_03
    
    Dim sxOpenQUERY_01_Cmd_03 As String                                  'Stringa per l'oggetto Docmd Apri Query 1 per il Cmd_03
    Dim sxOpenQUERY_02_Cmd_03 As String                                  'Stringa per l'oggetto Docmd Apri Query 2 per il Cmd_03
    Dim sxOpenQUERY_03_Cmd_03 As String                                  'Stringa per l'oggetto Docmd Apri Query 3 per il Cmd_03
    
    Dim sxEseguiRoutine_01_Cmd_03 As String                              'Stringa per l'oggetto Docmd Esegui Macro 1 per il Cmd_03
    
    'ATTIVITA CMD_04
    Dim sxOpenFORM_01_Cmd_04 As String                                   'Stringa per l'oggetto Docmd Apri form per il Cmd_4
    Dim sxOpenREPORT_01_Cmd_04 As String                                 'Stringa per l'oggetto Docmd Apri Report per il Cmd_4
    
    Dim sxOpenQUERY_01_Cmd_04 As String                                  'Stringa per l'oggetto Docmd Apri Query 1 per il Cmd_4
    Dim sxOpenQUERY_02_Cmd_04 As String                                  'Stringa per l'oggetto Docmd Apri Query 2 per il Cmd_4
    Dim sxOpenQUERY_03_Cmd_04 As String                                  'Stringa per l'oggetto Docmd Apri Query 3 per il Cmd_4
    
    Dim sxEseguiRoutine_01_Cmd_04 As String                              'Stringa per l'oggetto Docmd Esegui Macro 1 per il Cmd_4
    
    
    'ATTIVITA CMD_05
    Dim sxOpenFORM_01_Cmd_05 As String                                   'Stringa per l'oggetto Docmd Apri form per il Cmd_5
    Dim sxOpenREPORT_01_Cmd_05 As String                                 'Stringa per l'oggetto Docmd Apri Report per il Cmd_5
    
    Dim sxOpenQUERY_01_Cmd_05 As String                                  'Stringa per l'oggetto Docmd Apri Query 1 per il Cmd_5
    Dim sxOpenQUERY_02_Cmd_05 As String                                  'Stringa per l'oggetto Docmd Apri Query 2 per il Cmd_5
    Dim sxOpenQUERY_03_Cmd_05 As String                                  'Stringa per l'oggetto Docmd Apri Query 3 per il Cmd_5
    
    Dim sxEseguiRoutine_01_Cmd_05 As String                              'Stringa per l'oggetto Docmd Esegui Macro 1 per il Cmd_5
    
    'ATTIVITA Cmd_6
    Dim sxOpenFORM_01_Cmd_06 As String                                   'Stringa per l'oggetto Docmd Apri form per il Cmd_6
    Dim sxOpenREPORT_01_Cmd_06 As String                                 'Stringa per l'oggetto Docmd Apri Report per il Cmd_6
    
    Dim sxOpenQUERY_01_Cmd_06 As String                                  'Stringa per l'oggetto Docmd Apri Query 1 per il Cmd_6
    Dim sxOpenQUERY_02_Cmd_06 As String                                  'Stringa per l'oggetto Docmd Apri Query 2 per il Cmd_6
    Dim sxOpenQUERY_03_Cmd_06 As String                                  'Stringa per l'oggetto Docmd Apri Query 3 per il Cmd_6
    
    Dim sxEseguiRoutine_01_Cmd_06 As String                              'Stringa per l'oggetto Docmd Esegui Macro 1 per il Cmd_6
    
    
    'ATTIVITA Cmd_7
    Dim sxOpenFORM_01_Cmd_07 As String                                   'Stringa per l'oggetto Docmd Apri form per il Cmd_7
    Dim sxOpenREPORT_01_Cmd_07 As String                                 'Stringa per l'oggetto Docmd Apri Report per il Cmd_7
    
    Dim sxOpenQUERY_01_Cmd_07 As String                                  'Stringa per l'oggetto Docmd Apri Query 1 per il Cmd_7
    Dim sxOpenQUERY_02_Cmd_07 As String                                  'Stringa per l'oggetto Docmd Apri Query 2 per il Cmd_7
    Dim sxOpenQUERY_03_Cmd_07 As String                                  'Stringa per l'oggetto Docmd Apri Query 3 per il Cmd_7
    
    Dim sxEseguiRoutine_01_Cmd_07 As String                              'Stringa per l'oggetto Docmd Esegui Macro 1 per il Cmd_7
    
    'ATTIVITA Cmd_8
    Dim sxOpenFORM_01_Cmd_08 As String                                   'Stringa per l'oggetto Docmd Apri form per il Cmd_8
    Dim sxOpenREPORT_01_Cmd_08 As String                                 'Stringa per l'oggetto Docmd Apri Report per il Cmd_8
    
    Dim sxOpenQUERY_01_Cmd_08 As String                                  'Stringa per l'oggetto Docmd Apri Query 1 per il Cmd_8
    Dim sxOpenQUERY_02_Cmd_08 As String                                  'Stringa per l'oggetto Docmd Apri Query 2 per il Cmd_8
    Dim sxOpenQUERY_03_Cmd_08 As String                                  'Stringa per l'oggetto Docmd Apri Query 3 per il Cmd_8
    
    Dim sxEseguiRoutine_01_Cmd_08 As String                               'Stringa per l'oggetto Docmd Esegui Macro 1 per il Cmd_8
    
    
    'ATTIVITA Cmd_9
    Dim sxOpenFORM_01_Cmd_09 As String                                   'Stringa per l'oggetto Docmd Apri form per il Cmd_9
    Dim sxOpenREPORT_01_Cmd_09 As String                                 'Stringa per l'oggetto Docmd Apri Report per il Cmd_9
    
    
    Dim sxOpenQUERY_01_Cmd_09 As String                                  'Stringa per l'oggetto Docmd Apri Query 1 per il Cmd_9
    Dim sxOpenQUERY_02_Cmd_09 As String                                  'Stringa per l'oggetto Docmd Apri Query 2 per il Cmd_9
    Dim sxOpenQUERY_03_Cmd_09 As String                                  'Stringa per l'oggetto Docmd Apri Query 3 per il Cmd_9
    
    Dim sxEseguiRoutine_01_Cmd_09 As String                              'Stringa per l'oggetto Docmd Esegui Macro 1 per il Cmd_9
    
    'ATTIVITA Cmd_10
    Dim sxOpenFORM_01_Cmd_10 As String                                  'Stringa per l'oggetto Docmd Apri form per il Cmd_10
    Dim sxOpenREPORT_01_Cmd_10 As String                                'Stringa per l'oggetto Docmd Apri Report per il Cmd_10
    
    Dim sxOpenQUERY_01_Cmd_10 As String                                 'Stringa per l'oggetto Docmd Apri Query 1 per il Cmd_10
    Dim sxOpenQUERY_02_Cmd_10 As String                                 'Stringa per l'oggetto Docmd Apri Query 2 per il Cmd_10
    Dim sxOpenQUERY_03_Cmd_10 As String                                 'Stringa per l'oggetto Docmd Apri Query 3 per il Cmd_10
    
    Dim sxEseguiRoutine_01_Cmd_10 As String                             'Stringa per l'oggetto Docmd Esegui Macro 1 per il Cmd_10
    
    'ATTIVITA Cmd_11
    Dim sxOpenFORM_01_Cmd_11 As String                                  'Stringa per l'oggetto Docmd Apri form per il Cmd_11
    Dim sxOpenREPORT_01_Cmd_11 As String                                'Stringa per l'oggetto Docmd Apri Report per il Cmd_11
    
    Dim sxOpenQUERY_01_Cmd_11 As String                                 'Stringa per l'oggetto Docmd Apri Query 1 per il Cmd_11
    Dim sxOpenQUERY_02_Cmd_11 As String                                 'Stringa per l'oggetto Docmd Apri Query 2 per il Cmd_11
    Dim sxOpenQUERY_03_Cmd_11 As String                                 'Stringa per l'oggetto Docmd Apri Query 3 per il Cmd_11
    
    Dim sxEseguiRoutine_01_Cmd_11 As String                               'Stringa per l'oggetto Docmd Esegui Macro 1 per il Cmd_11
    
    'Gestione parametri comandi
    Dim sxParamCmd_3 As String
    
    '//I COMANDI FORM = PATH + LATEX + TUTTO
    '//File da ricercare e la path ANALISI DELLA FORM GESTIONE PRINCIPALE
    Dim MyFile_s As String
    Dim MyPath_s As String
 



'--------------------------------------------------------------------


'GESTIONE PROGETTI
'--------------------------------------------------------------------


    Dim sxTIPOGGETTO As String
    Dim sxTIPOGGETTO_2 As String
    
    Dim sxNOMEOGGETTO As String
    Dim sxNOMEOGGETTO_2 As String
    
    Dim sxPROPRIETA As String
    Dim sxPROPRIETA_2 As String
    
    Dim sxVALOREPROPRIETA As String
    Dim sxVALOREPROPRIETA_2 As String
    
    
    
    Dim sxMETODO    As String
    Dim sxMETODO_2    As String
    
    
    Dim sxCODATTIVITA As String
    Dim sxCODATTIVITA_2 As String
    
    Dim sxATTIVITAEVENTO As String
    
    
    Dim sxCRITERIO As String
    Dim sxCRITERIO_2 As String
    
    Dim stDocName As String                             'nome form da aprire con il comando
    Dim stLinkCriteria As String                        'criterio di ricerca record nella form da aprire con il comando
    
    
    'PAGINE TAB CONTROLL
    '-----------------------------------------------
    Dim ixPagCorrente As Integer
    Dim ixPage As Integer
    
    
    'Etichetta della pagina
    Dim sxPagCorrente_Label  As String
    'Indice pagina
    Dim ixPagCorrente_Index  As Integer
    'Nome elemento
    Dim sxPagCorrente_Name  As String
    
    'Nome elemento - Etichetta
    Dim sxPagCorrente_Caption  As String

    'I parametri del tab controll
    Dim ixCmb As Integer                            'Casella combinata : 0= Cmb_01, 1 = Cmb_02 ecc.
    Dim sxItemText  As String                       'Colonna zero casella combinata
    Dim ixCommand As Integer                        'Colonna 1 della casella combinata
    Dim vxArgument As Variant                       'Colonna 2 della casella combinata




'................................
' DIM OGGETTI E PARAMETRI
Dim obTipoDB As Object

'................................
' DIM SCELTA DATABASE DA APRIRE
Dim sOption_SEZ As Integer          'Numero Sezione di Progetto Scelta


Dim dbs As Database
Dim Apridbs As Database
Dim rs As Recordset

Dim intCiclo As Integer             'Intero per ciclo lettura oggetti

'................................
' DIM TABELLA DA ESPORTARE
Dim ExportTable1 As String
Dim ExportTable2 As String

'................................
' DIM CREAZIONE NUOVE TABELLE
Dim TableDefNuovo    As TableDef
Dim ApriTableDef    As TableDef
Dim FieldNuovo  As Field
Dim idxNuovo_Contatore_Univoco  As Index        ' Oggetto Index
Dim idxNuovo_Testo_Univoco  As Index
Dim prpPROPERTY1   As Property                  ' proprieta Description

Dim sNomeField  As String


'................................
' DIM QUERY E PARAMETRI DI CHIAMATA DELLA FUNZIONI



                                                    

Dim sName_Query As String
Dim sName_Form As String
Dim sName_Report As String
Dim sName_MODULI As String
Dim sName_MACRO As String
Dim sName_COD_PROGETTO_s As String
Dim iIndice1    As Integer                            ' per il passaggio dei parametri
Dim sTipo_Oggetto As String
Dim sTipo_OggettoMatrice(5) As String                             'LA MATRICE
Dim sName_Tab As String                                 'Nome Tabella da creare
Dim sScelta_db As String                                'scelta del database esterno  @dim.database.esterno

Dim CodCmd_s As String              '//il comando
Dim iOperazione As Integer




'................................
' DIM LE OPZIONI
Dim Opzione_1 As Integer
Dim Opzione_2 As Integer

Dim Flag As Boolean
        



'//ERRORI PROCEDURA_FUNCTION O ROUTINE
Dim ProceduraMessaggioErrore_s As String
Dim ProceduraAttivaEseguita_s As String









'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'                                   IMPOSTA I CONTROLLI
'                                   DELLA FORM
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<


'//FORM -> EVENTO @FORM@LOAD@CARICA@PAGINE@TUTTE
'//=========================================================================================// _
    All'apertura della form carico tutto le pagine del tab controll per un  max di 12
    




'//help@progetti
Private Sub Btn_HELP_ANALISI_MSYS_Click()

On Error GoTo Btn_HELP_ANALISI_MSYS_Click_Err

    
'//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
'//--------------------------------------------------------------------------------//--------//
'//NOTE                 -> Apro il file di tipo doc, zip o pdf
'//CODICE               -> Function ApriFilePdf_Interna_s_PFunct.01.01
'//PARAMETRI            -> par_Path_s         = PATH _
                        -> par_NameFile_s               = NOME FILE _
                        -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                        
  '//IMPOSTO LE VARIABILI PATH E FILE
  '//
  Dim MyFile_s As String
  '//
  Dim MyPath_s  As String
  
        '//imposto i parametri - LA PATH
        MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\OBJECT\HELP\"
        '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
        MyFile_s = "APRI_HELP_ANALISI_MSYS_{Word365}.bat "
            
            
         '//chiamo la sub con i parametri =   'CALL (ApriFilePdf_Interna_s_PFunct Path_programma_s, NameFile_s, IDGestione_lng)
            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
             
        
'//--------------------------------------------------------------------------------//--------//
                        
         

Btn_HELP_ANALISI_MSYS_Click_Err_Exit:
Exit Sub

Btn_HELP_ANALISI_MSYS_Click_Err:
    MsgBox Error$
    Resume Btn_HELP_ANALISI_MSYS_Click_Err_Exit
    MsgBox Err.Description & " " & ProceduraMessaggioErrore_s & " - " & ProceduraAttivaEseguita_s & " --> errore Error$ -> : " & Error$
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Debug.Assert "BLOCCO PROCEDURA -> " & ProceduraAttivaEseguita_s
  
    Stop
    Resume Btn_HELP_ANALISI_MSYS_Click_Err
    
End Sub















Private Sub Form_Load()

Dim SottoForm_s  As String
Dim NameSottoForm_s  As String

    '//ATTIVITA ESEGUITE ALL'EVENTO LOAD DELLA FORM
    '//----------------------------------------------------------------------------------------------
    '// @RESET@PAGINE_(sempre all'apertura vengono resettare le sottoform le pagine del tab controll)
  
        '//01) RESETTA_TUTTE_LE_SOTTOFORM+LE_SCHEDE
        '//......................................................
            Call Imposta_N21_RESETTA_LE_SOTTOFORM_Sub
        '//......................................................
        
        '//02) RESETTA TUTTE LE PAGINE
        '//......................................................
            Call Imposta_N20_RESETTA_LE_PAGINE_Sub
        
        '//......................................................
        
        '//03) IMPOSTO LE SOTTOFORM
        '//......................................................................................................................
        
            
            '//resetto la pag_01  - @carica.sottoform, @pagine.sottoform.carica, @carica.dati.sottoform
            '//Non è necessario la successione, ATTENZIONE SE SBAGLI IL NOME DELLE FORMA DA ERRORE.
            'Call Imposta_N22_IMPOSTA_LE_SOTTOFORM_Sub("SottoForm_01", "MODELLO_FrS01_01_MODELLO_ELENCO", True, "")
           ' Call Imposta_N22_IMPOSTA_LE_SOTTOFORM_Sub("SottoForm_01", "MSsys_Frm01_S01_TABELLE_ELENCO", True, "01 TABELLE")
                               '//PARAMETRI ----->    LA SOTTOFORM....LA FORM  SOTTOSTANTE ...........VISIBLE...LABEL PAG

            'Call Imposta_N22_IMPOSTA_LE_SOTTOFORM_Sub("SottoForm_02", "MSys_QUERY_Frm01_S02_TUTTE", True, "02 QUERY")
            'Call Imposta_N22_IMPOSTA_LE_SOTTOFORM_Sub("SottoForm_03", "MSsys_Frm01_S02_COLLEGA_TABELLE", True, "03 COLLEGA TABELLE")
            'Call Imposta_N22_IMPOSTA_LE_SOTTOFORM_Sub("SottoForm_04", "MSsys_Frm01_S01_DF02_TIPO_OGGETTO", True, "04 TIPO OGGETTO")
            'Call Imposta_N22_IMPOSTA_LE_SOTTOFORM_Sub("SottoForm_05", "MSsys_Frm01_S01_ELENCO_OGGETTI_MSsys", True, "05 ELENCO OGGETTI")
            'Call Imposta_N22_IMPOSTA_LE_SOTTOFORM_Sub("SottoForm_06", "MSsys_Frm01_S11_ELENCO_DISCO_ESPORTA", True, "06 ESPORTA")
            'Call Imposta_N22_IMPOSTA_LE_SOTTOFORM_Sub("SottoForm_07", "", True, "07")
            'Call Imposta_N22_IMPOSTA_LE_SOTTOFORM_Sub("SottoForm_08", "", True, "08")
            'Call Imposta_N22_IMPOSTA_LE_SOTTOFORM_Sub("SottoForm_09", "", True, "09")
            'Call Imposta_N22_IMPOSTA_LE_SOTTOFORM_Sub("SottoForm_10", "", True, "10")
            'Call Imposta_N22_IMPOSTA_LE_SOTTOFORM_Sub("SottoForm_11", "", True, "11")
            'Call Imposta_N22_IMPOSTA_LE_SOTTOFORM_Sub("SottoForm_12", "MSsys_Frm01_S12_SERVIZI", True, "12 SERVIZI")
                    
                '//TODO: SOSTITUIRE CON INTERAZIONE NEL DB con la tabella MSys_TUTTE_LE_PAGINE_CARICATE
                

        '//......................................................................................................................
        
                            
                        '//ATTIVA IL COMANDO DELLA COMBINATA
                        '==================================================================================================
                        'Note: Leggo i dati dalla tabella _
                        @ATTIVO@FUNZIONE@CARICA@PAGINE_(viene attivata la funzione che carica le pagine delle sottoform)

                                
                                '//CHIAMO QUESTA FUNZIONE
                                MSys_TUTTE_LE_PAGINE_Carica_pFunction
                        
                        '//ATTIVA IL COMANDO DELLA COMBINATA
                        '==================================================================================================
                                




        '//......................................................
    
    '//ATTIVITA ESEGUITE ALL'EVENTO LOAD DELLA FORM *** FINE ***
    '//----------------------------------------------------------------------------------------------

    '//04) REIMPOSTA .......................
    '//--------------------------------------------------
        '......
    '//--------------------------------------------------
    
     
    '//11) REIMPOSTA IL FORMATO INZIALE DEL CMB_01
    '//--------------------------------------------------
    '//Note     :Reimposto il formato iniziale del
    '//         comando ripristinando i colori iniziali

        Call Cmd_N02_ReimpostaFormati_Sub("Cmd_01")
    '//--------------------------------------------------
    
    
    '//SOSPES0****
    '//12)REIMPOSTA IL VALORE DELL'ETICHETTA CMB_01
    '//--------------------------------------------------
                
        '//Imposto il valore dell'etichetta
        'Me.Cmb_01_txt.Value = "Agg/Ins/Canc in Archivio"
        'Me.Recalc
        'Me.Cmb_01_txt.Requery
        
        'Me.TxtTit_01.Caption = "Gestione"
        
    '//--------------------------------------------------
              
   
    

End Sub
'//FORM -> EVENTO LOAD      *** FINE ***
'//=========================================================================================//


'//FORM -> EVENTO CLOSE      *** FINE ***
'//=========================================================================================//

Private Sub Form_Close()
    'Imposta_N20_RESETTA_LE_PAGINE_Sub
    'Imposta_N21_RESETTA_LE_SOTTOFORM_Sub
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '//2024.08.08_MODIFICA=resetto_solo_sottoform_NO_RESET_PAGINE
        'Imposta_N20_RESETTA_LE_PAGINE_Sub -- BLOCCATO PERCHE DA ERRORE
    Imposta_N21_RESETTA_LE_SOTTOFORM_Sub

End Sub


'//FORM -> EVENTO CLOSE      *** FINE ***
'//=========================================================================================//



'//REIMPOSTA IL FORMATO DEL COMANDO ALLE POSIZIONI INIZIALI
Private Sub Cmd_N02_ReimpostaFormati_Sub(par_sxComando)

On Error GoTo Err_Cmd_N02_ReimpostaFormati_Sub

    '//REIMPOSTA IL FORMATO INZIALE
    '//--------------------------------------------------
    '//Note     :Reimposto il formato iniziale del
    '//         comando ripristinando i colori iniziali

        'CALL Cmd_N02_ReimpostaFormati_Sub ("Cmd_01")
    '//--------------------------------------------------

'//IMPOSTA I FORMATI DEI COMANDI
'//..............................................................................................
  
        
        
           
        
    
'//IMPOSTA I FORMATI DEI COMANDI
'//..............................................................................................
   

Exit_Cmd_N02_ReimpostaFormati_Sub:
    Exit Sub

Err_Cmd_N02_ReimpostaFormati_Sub:
    MsgBox Err.Description
    Resume Exit_Cmd_N02_ReimpostaFormati_Sub

End Sub



'//PAGINE TAB -> RESETTA TUTTE LE PAGINE
'//=========================================================================================//

Private Sub Imposta_N20_RESETTA_LE_PAGINE_Sub()

On Error GoTo Err_Imposta_N20_RESETTA_LE_PAGINE_Sub



    
'//RESETTA TUTTE LE PAGINE
'//......................................................
    'call Imposta_N20_RESETTA_LE_PAGINE_Sub
     
'//......................................................


'//RESETTA TUTTE LE PAGINE, @visibili.pagine. @pagine.visibili
'//..............................................................................................
        
        Me.PAG_01.Caption = ""
        Me.PAG_01.Visible = False
        Me.PAG_02.Visible = False
        Me.PAG_03.Visible = False
        Me.Pag_04.Visible = False
        Me.PAG_05.Visible = False
        Me.Pag_06.Visible = False
        Me.Pag_07.Visible = False
        Me.Pag_08.Visible = False
        Me.Pag_09.Visible = False
        Me.Pag_10.Visible = False
        Me.PAG_11.Visible = False
        Me.Pag_12.Visible = True

'//..............................................................................................


Exit_Imposta_N20_RESETTA_LE_PAGINE_Sub:
    Exit Sub

Err_Imposta_N20_RESETTA_LE_PAGINE_Sub:
    MsgBox Err.Description
    Resume Exit_Imposta_N20_RESETTA_LE_PAGINE_Sub

End Sub

'//PAGINE TAB -> RESETTA TUTTE LE PAGINE    *** FINE ***
'//=========================================================================================//






'//PAGINE TAB SOTTOFORM -> @RESETTA@SOTTOFORM_(RESETTA TUTTE LE SOTTOFORM con la proprieta visible=false)
'//=========================================================================================//

Public Sub Imposta_N21_RESETTA_LE_SOTTOFORM_Sub()

On Error GoTo Err_Imposta_N21_RESETTA_LE_SOTTOFORM_Sub



    
'//RESETTA_TUTTE_LE_SOTTOFORM+LE_SCHEDE
'//......................................................
    'call Imposta_N21_RESETTA_LE_SOTTOFORM_Sub
     
'//......................................................


'//RESETTA TUTTE LE PAGINE+LE_SCHEDE  @resetta.tutte.le.pagine
'//..............................................................................................
'//controllo con lo stato attivo, lo imposto a pag 12

        '//@cotrollo@attivo_(imposto lo stato attivo sul controllo 12 altrimenti da errore in apertura)
        Me.SottoForm_12.SetFocus
        Me.Pag_12.SetFocus
        
        Me.SottoForm_01.SourceObject = ""
        Me.PAG_01.Visible = False
        Me.SottoForm_01.Visible = False
        Me.PAG_01.Caption = ""
        
        Me.SottoForm_02.SourceObject = ""
        Me.PAG_02.Visible = False
        Me.SottoForm_02.Visible = False
        Me.PAG_02.Caption = ""
        
        Me.SottoForm_03.SourceObject = ""
        Me.PAG_03.Visible = False
        Me.SottoForm_03.Visible = False
        Me.PAG_03.Caption = ""
        
        Me.SottoForm_04.SourceObject = ""
        Me.Pag_04.Visible = False
        Me.SottoForm_04.Visible = False
        Me.Pag_04.Caption = ""
        
        Me.SottoForm_05.SourceObject = ""
        Me.PAG_05.Visible = False
        Me.SottoForm_05.Visible = False
        Me.PAG_05.Caption = ""
        
        Me.SottoForm_06.SourceObject = ""
        Me.Pag_06.Visible = False
        Me.SottoForm_06.Visible = False
        Me.Pag_06.Caption = ""
        
        Me.SottoForm_07.SourceObject = ""
        Me.Pag_07.Visible = False
        Me.SottoForm_07.Visible = False
        Me.Pag_07.Caption = ""
        
        
        Me.SottoForm_08.SourceObject = ""
        Me.Pag_08.Visible = False
        Me.SottoForm_08.Visible = False
        Me.Pag_08.Caption = ""
        
        Me.SottoForm_09.SourceObject = ""
        Me.Pag_09.Visible = False
        Me.SottoForm_09.Visible = False
        Me.Pag_09.Caption = ""
        
        Me.SottoForm_10.SourceObject = ""
        Me.Pag_10.Visible = False
        Me.SottoForm_10.Visible = False
        Me.Pag_10.Caption = ""
        
        Me.SottoForm_11.SourceObject = ""
        Me.PAG_11.Visible = False
        Me.SottoForm_11.Visible = False
        Me.PAG_11.Caption = ""
        
        'Me.SottoForm_12.SourceObject = ""
        'Me.Pag_12.Visible = True
        'Me.SottoForm_12.Visible = True
        'Me.Pag_12.Caption = ""
        
        
'//..............................................................................................


Exit_Imposta_N21_RESETTA_LE_SOTTOFORM_Sub:
    Exit Sub

Err_Imposta_N21_RESETTA_LE_SOTTOFORM_Sub:
    MsgBox Err.Description
    Resume Exit_Imposta_N21_RESETTA_LE_SOTTOFORM_Sub

End Sub
'//PAGINE TAB SOTTOFORM -> RESETTA TUTTE LE SOTTOFORM *** FINE ***
'//=========================================================================================//



'//PAGINE TAB SOTTOFORM -> @ROUTINE@IMPOSTA@LE@SINGOLE@SOTTOFORM_(con la ROUTINE imposta le sottoform MA attivo la classe ricerca oggetto nel db)
'//=========================================================================================//

Private Sub Imposta_N22_IMPOSTA_LE_SOTTOFORM_Sub(par_SottoForm_s As String, _
                                                 par_SourceObject_s As String, _
                                                 par_Visible_b As Boolean, _
                                                 par_Caption_s As String)

On Error GoTo Err_Imposta_N22_IMPOSTA_LE_SOTTOFORM_Sub



    
'//RESETTA_TUTTE_LE_SOTTOFORM+LE_SCHEDE
'//......................................................
    '//CHIAMO la routine e resetto tutte le pagine a schede
    '//Call Imposta_N22_IMPOSTA_LE_SOTTOFORM_Sub("SottoForm_01", "", False, "")
     
'//......................................................

'//TODO: SOLO SE IL PARAMETRO VISIBLE= TRUE (IF DI CONTROLO) e cioè la sottoform deve essere caricata, _
         VIENE EFFETTUATO IL CONTROLLO PREVENTIVO se esiste la form (CICLO FOR EACH) e deve essere eseguito _
         prima della select altrimenti se non esiste la form da inserire nella pagina viene dato @errore@proprieta _
         nella @visible. Visto che la sottoform da caricare è stata definita nel parametro _
         par_SottoForm_s verificare con il ciclo for next sulle form esistente.
         
         
        If par_Visible_b = True Then



            '//FUNZIONE_DI_CONTROLLO_DELLA_ESISTENZA_DELL'OGGETTO_FORM
            '//CODICE-----> Controlla_SE_OGGETTO_FORM_ESISTE_pFunct.attiva.LaClasse
            '//---------------------------------------------------------------------------------//
            '//NOTA  :Controlla se esiste la form negli oggetti del database e restituisce TRUE _
                      se il nome della form passata con parametro par_sxNameForm esiste; altrimenti FALSE
                
            
            
                        'ISTANZIO LA CLASSE CLASSE_GESTIONE_OGGETTI_FORM
                        'Assegno e creo l'oggetto di classe
                        '
                        Dim MyCls_CLASSE_GESTIONE_OGGETTI_FORM As Classe_GESTIONE_OGGETTI_FORM
                        Dim sxNameForm As String
                        '
                        Set MyCls_CLASSE_GESTIONE_OGGETTI_FORM = New Classe_GESTIONE_OGGETTI_FORM
                        
                        
                        'PARAMETRI DI ESTRAZIONE
                        'Attivo il metodo della classe rappresentata dalla routine passando il nome _
                        dell'oggetto form da ricercare.
                        '
                        sxNameForm = par_SourceObject_s
                        '
                        Boolean1 = MyCls_CLASSE_GESTIONE_OGGETTI_FORM.Controlla_SE_OGGETTO_FORM_ESISTE_pFunct(sxNameForm)
                        
                        '
                        sxNameForm = vbNull
                        Set MyCls_CLASSE_GESTIONE_OGGETTI_FORM = Nothing
                                                    
                '//---------------------------------------------------------------------------------//
          
          
                  
                
                
                '//RESETTA TUTTE LE PAGINE+LE_SCHEDE
                '//..............................................................................................
                '//le sottoform sono caricate solo se esistono. Quindi il controllo con la variabile boolean
                    
                
                    '//solo se la form esiste con la variabile settata a TRUE vengono caricate le sottoform
                    If Boolean1 = True Then
                    
                            Select Case par_SottoForm_s
                                Case Is = "SottoForm_01"
                                    
                                    Me.SottoForm_01.SourceObject = par_SourceObject_s
                                    Me.SottoForm_01.Visible = par_Visible_b
                                    Me.PAG_01.Visible = par_Visible_b
                                    Me.PAG_01.Caption = par_Caption_s
                                
                                Case Is = "SottoForm_02"
                                    Me.SottoForm_02.SourceObject = par_SourceObject_s
                                    Me.SottoForm_02.Visible = par_Visible_b
                                    Me.PAG_02.Visible = par_Visible_b
                                    Me.PAG_02.Caption = par_Caption_s
                                        
                                Case Is = "SottoForm_03"
                                    Me.SottoForm_03.SourceObject = par_SourceObject_s
                                    Me.SottoForm_03.Visible = par_Visible_b
                                    Me.PAG_03.Visible = par_Visible_b
                                    Me.PAG_03.Caption = par_Caption_s
                                            
                                Case Is = "SottoForm_04"
                                    Me.SottoForm_04.SourceObject = par_SourceObject_s
                                    Me.SottoForm_04.Visible = par_Visible_b
                                    Me.Pag_04.Visible = par_Visible_b
                                    Me.Pag_04.Caption = par_Caption_s
                                
                                Case Is = "SottoForm_05"
                                    Me.SottoForm_05.SourceObject = par_SourceObject_s
                                    Me.SottoForm_05.Visible = par_Visible_b
                                    Me.PAG_05.Visible = par_Visible_b
                                    Me.PAG_05.Caption = par_Caption_s
                                    
                                Case Is = "SottoForm_06"
                                    Me.SottoForm_06.SourceObject = par_SourceObject_s
                                    Me.SottoForm_06.Visible = par_Visible_b
                                    Me.Pag_06.Visible = par_Visible_b
                                    Me.Pag_06.Caption = par_Caption_s
                                        
                                Case Is = "SottoForm_07"
                                    Me.SottoForm_07.SourceObject = par_SourceObject_s
                                    Me.SottoForm_07.Visible = par_Visible_b
                                    Me.Pag_07.Visible = par_Visible_b
                                    Me.Pag_07.Visible = par_Visible_b
                                    Me.Pag_07.Caption = par_Caption_s
                                    
                                    
                                    
                                Case Is = "SottoForm_08"
                                    Me.SottoForm_08.SourceObject = par_SourceObject_s
                                    Me.SottoForm_08.Visible = par_Visible_b
                                    Me.Pag_08.Visible = par_Visible_b
                                    Me.Pag_08.Visible = par_Visible_b
                                    Me.Pag_08.Caption = par_Caption_s
                                
                                Case Is = "SottoForm_09"
                                    Me.SottoForm_09.SourceObject = par_SourceObject_s
                                    Me.SottoForm_09.Visible = par_Visible_b
                                    Me.Pag_09.Visible = par_Visible_b
                                    Me.Pag_09.Caption = par_Caption_s
                                        
                                Case Is = "SottoForm_10"
                                    Me.SottoForm_10.SourceObject = par_SourceObject_s
                                    Me.SottoForm_10.Visible = par_Visible_b
                                    Me.Pag_10.Visible = par_Visible_b
                                    Me.Pag_10.Caption = par_Caption_s
                                
                                Case Is = "SottoForm_11"
                                    Me.SottoForm_11.SourceObject = par_SourceObject_s
                                    Me.SottoForm_11.Visible = par_Visible_b
                                    Me.PAG_11.Visible = par_Visible_b
                                    Me.PAG_11.Caption = par_Caption_s
                                
                                Case Is = "SottoForm_12"
                                    Me.SottoForm_12.SourceObject = par_SourceObject_s
                                    Me.SottoForm_12.Visible = par_Visible_b
                                    Me.Pag_12.Visible = par_Visible_b
                                    Me.Pag_12.Caption = par_Caption_s
                                
                                Case Else
                                    
                                End Select
                        '//..............................................................................................
                
                    End If 'If Boolean1 = True = se la form eiste viene caricata
                
                
      End If  '// If par_Visible_b = True = prma viene controllato se la form deve essere visibile = True
                
                

Exit_Imposta_N22_IMPOSTA_LE_SOTTOFORM_Sub:
    Exit Sub

Err_Imposta_N22_IMPOSTA_LE_SOTTOFORM_Sub:
    MsgBox Err.Description
    Resume Exit_Imposta_N22_IMPOSTA_LE_SOTTOFORM_Sub

End Sub

'//PAGINE TAB SOTTOFORM -> IMPOSTA LE SINGOLERE SOTTOFORM *** FINE ***
'//=========================================================================================//





'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'
'                                   I COMANDI E LE COMBINATE
'
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////



'//ESCI
Private Sub Cmd_01_Click()
On Error GoTo Err_Cmd_01_Click

    '//resetta le sottoform
    Imposta_N21_RESETTA_LE_SOTTOFORM_Sub
    '//chiudi
    DoCmd.Close

Exit_Cmd_01_Click:
    Exit Sub

Err_Cmd_01_Click:
    MsgBox Err.Description
    Resume Exit_Cmd_01_Click
    
End Sub

'//RICERCA
Private Sub Cmd_02_TXT_Click()
On Error GoTo Err_Cmd_02_TXT_Click


    Screen.PreviousControl.SetFocus
    DoCmd.DoMenuItem acFormBar, acEditMenu, 10, , acMenuVer70

Exit_Cmd_02_TXT_Click:
    Exit Sub

Err_Cmd_02_TXT_Click:
    MsgBox Err.Description
    Resume Exit_Cmd_02_TXT_Click
    
End Sub


'//---> button:cdm-tutto
Private Sub TUTTO_cmd_Click()
'APRO DB TABELLA LLPP\_ATTI\_Tb01\_GESTIONE.mdb + APRO la cartella di gestione + _
APRO la directory ....?

    '//---> button:cdm-tutto.01 _
    APRO tutto con la procedura .bat
    MyFile_s = "APRI_OGGETTI_TUTTI_FORM_PRINCIPALE.bat"
    MyPath_s = "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHVI_MDB\LLPP_GESTIONE\LPP_ANALISI\GESTIONE_FORM_PRINCIPALE\"
    'AnalisiProcedimento_N01_CORRENTE_PFucnt MyPath_s, MyFile_s
    
    '//LINK apro i collegamenti
    '//APRO la cartella di gestione
    MyFile_s = "AA_ANALISI_LINK.lnk"
    MyPath_s = "C:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHVI_MDB\LLPP_GESTIONE\LPP_ANALISI\AA_ANALISI_LINK\"
    'AnalisiProcedimento_N01_CORRENTE_PFucnt MyPath_s, MyFile_s
     
      
    '//DOVE SI TROVA IL FILE DI GESTIONE ORIGINALE _
    il file di gestione degli atti <<LLPP_ATTI_GESTIONE>> di lavoro si trova in questa directory _
    C:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHVI_MDB per il file LLPP_ATTI_GESTIONE.mdb
    'vV1 = MsgBox("Apro la directory dove si trova il di gestione originale, il file di lavoro (LLPP_ATTI_GESTIONE.mdb)  -> C:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP", vbInformation, "DOVE SI TROVA IL FILE DI GESTIONE")
    '//APRO la directory ....?
    MyFile_s = "APRI_OGGETTI_TUTTI_FORM_PRINCIPALE.bat"
    MyPath_s = "o:\Casa\Linguaggi\ACCESS\GESTIONE_MODELLI_OGGETTI\"
    'AnalisiProcedimento_N01_CORRENTE_PFucnt MyPath_s, MyFile_s
    
 
End Sub


'//CODICE--------> \label{button:cdm-LATEX}
'//APRO LA CARTELLA DEL DB GESTIONE FORM PRINCIPALE
Private Sub PATH_cmd_Click()
'//

    MyFile_s = "APRI_CARTELLA_FORM_PRINCIPALE.bat"
    MyPath_s = "o:\Casa\Linguaggi\ACCESS\GESTIONE_MODELLI_OGGETTI\"
    'AnalisiProcedimento_N01_CORRENTE_PFucnt MyPath_s, MyFile_s
    

End Sub


'//CODICE--------> \label{button:cdm-PAHT}
'//APRO LATEX DI ANALISI DEL DB DI GESTIONE DELLA FORM PRINCIPALE
Private Sub LATEX_Cmd_Click()
     
    MyPath_s = "o:\Casa\Linguaggi\ACCESS\GESTIONE_MODELLI_OGGETTI\"
    MyFile_s = "APRI_ANALISI_TEX_GE_OGGETTI_(MASTER_AUTOMATICO)_ANALISI.bat"
    'AnalisiProcedimento_N01_CORRENTE_PFucnt MyPath_s, MyFile_s
    

End Sub



'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'
'                                   I COMANDI E LE COMBINATE *** FINE ***
'   CASELLE COMBINATE _
        Cmb_01          = evento doppio click apri la tabella COMANDI               - @cmbd_01.Doppioclick.DblClick _
        Cmb_01          = Evento dopo aggiornamento                                 - @cmbd_01.DopoAggiornamento.AfterUpdate _

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////


'//@cmbd_01.casella.combinata.evento_(APRO GRUPPO COMANDI CASELLA COMBINATA 01)
Private Sub Cmb_01_TXT_DblClick(Cancel As Integer)
    DoCmd.OpenForm "GEST_MENU_Tb01Frm01_M01_GE_VOCI"
End Sub




    
'//ATTIVA IL COMANDO -
'//@cmbd_01.casella.combinata.evento
Private Sub Cmb_01_txt_AfterUpdate()


On Error GoTo Err_Cmb_01_Txt_AfterUpdate



    '//ATTIVA IL COMANDO DELLA COMBINATA @COMANDO.CMB
    '//==================================================================================================
        '//@ultimo.comando.cmb_01
        
        Dim UltimoCmd_s As String
        Dim par_sxItemText As String

        UltimoCmd_s = ""
        
         '//SOSPES0****
        'UltimoCmd_s = Me.Cmb_01_txt.Column(0)
        par_sxItemText = ""
        par_sxItemText = UltimoCmd_s
        
        
      
    '//==================================================================================================


       '//ATTIVA IL COMANDO DELLA COMBINATA - @MENU.CMD_01
       '//=========================================================================================================================//
       'Note: Leggo i dati dalla tabella MSys_DF01_COMANDI, _
                01) RECUPERO IL COMANDO DALLA TABELLA, _
                02) IMPOSTO LA VARIABILE ULTIMO COMANDO RECUPERANDO IL VALORE DAL CAMPO _
                03) CONTROLLO TIPO DI COMANDO > NULL _
                04) CONTROLLO TIPO DI OGGETTO > NULL _
                05) SALVO IL COMANDO DELLA TABELLA NELLA VARIABILE _
                06) SALVO IL TIPO OGGETTO NELLA VARIABILE _
                07) CHIAMO LA FUNZIONE DI ITERAZIONE OGGETTI DEL DB CORRENTE _
                08) AGGIORNO LA FORM CORENTE
                
       


    
                 '//@tabella.comandi = apro la tabella che contiene i comandi della cmb_01 e li confronto _
                    per l'attivazione dell'evento. Modificato con un filtro diretto alla scelta _
                    effettuata.
               ' sSql = "MSys_DF01_COMANDI_Qry01_01_SELECT_TUTTI"
                    sSql = ""
                    sSql = sSql & "SELECT MSys_DF01_COMANDI.* "
                    sSql = sSql & "FROM MSys_DF01_COMANDI "
                    sSql = sSql & "WHERE (((MSys_DF01_COMANDI.COMANDO_s)='" & UltimoCmd_s & "'));"

                
                'Apro rs estrazione Sql
                Set daoRS = CurrentDb.OpenRecordset(sSql)
                    
                      
                      '//ITERO NELLA TABELLA - CONTROLLA IL COMANDO SCELTO
                      '//--------------------------------------------------------------------------------
                      'Le proprieta del controllo a schede :
                      '               - Visible
                      '               - Enabled
                      
                        If daoRS.EOF = False And daoRS.BOF = False Then
                            
                            daoRS.MoveFirst
                            
                            While Not daoRS.EOF
                      
                                  'Imposto le variabili parametro
                                  '.....................................................................................
                                            DoEvents
                                            'RESET
                                            gsxOggetto = daoRS.Fields("COMANDO_s").Value
                                       
                                            sScelta_db = ""         'n.b. con la path nulla verrà aperto il db corrente
                                            sName_Tab = ""
                                            sName_Query = ""
                                            iOperazione = 0
                                            
                                            
                                        
                                  '.....................................................................................
                                  
                                    '//02) IMPOSTO LA VARIABILE ULTIMO COMANDO RECUPERANDO IL VALORE DAL CAMPO
                                    If UltimoCmd_s = daoRS.Fields("COMANDO_s").Value Then
                                            
                                         
                                            
                                            
                                              '//@ITERAZIONE.OGGETTI = se l'oggetto query   @finire??
                                 '______________________________________________________________________________
                                 
                                            
                                            '//CONTROLLI ERRORI BLOCCANTI PARAMETRI - PRELIMINARI -
                                            '//------------------------------------------------------------------------------------------
                                            '//Note: controllo preliminare dei parametri passati alla routine, errori bloccanti permettono _
                                                    l'uscita dalla routine. Venono controllati i seguenti parametri che devono essere valorizzati _
                                                    al fine di evitare il msg e l'uscita dalla routine: _
                                                    CODICE COMANDO              par_CodCmd_s _
                                                    TIPO DI OGGETTO             TIPOGGETTO_s

                                                '//03) CONTROLLO TIPO DI COMANDO > NULL
                                                If daoRS.Fields("COMANDO_s").Value = "" Or IsNull(daoRS.Fields("COMANDO_s").Value) = True Then
                                                    MsgBox "PARAMETRO - CodCmd_s  - Codice comando nullo -> Exit_Iterazione_oggetti - Uscita dalla routine", vbExclamation, "CONTROLLO PARAMETRI COMANDO"
                                                    GoTo Exit_Cmb_01_Txt_AfterUpdate
                                                End If
                                                
                                                
                                                                                            
                                            '//04) CONTROLLO TIPO DI OGGETTO > NULL
                                            If IsNull(daoRS.Fields("TIPOGGETTO_s").Value) = True Or daoRS.Fields("TIPOGGETTO_s").Value = "" Then
                                            
                                                MsgBox "Attenzione uscita dalla Routine OGGETTO NULLO -> " & daoRS.Fields("TIPOGGETTO_s").Value, vbExclamation, "MSGBOX CONTROLLO OGGETTO"
                                                '//USCITA
                                                GoTo Exit_Cmb_01_Txt_AfterUpdate
                                            End If

                                                
                                                    '//05) SALVO IL COMANDO DELLA TABELLA NELLA VARIABILE
                                                    'controllo oggetto
                                                    Debug.Print "COMANDO         -> " & daoRS.Fields("COMANDO_s").Value
                                                    Debug.Print "TIPO DI OGGETTO -> " & daoRS.Fields("TIPOGGETTO_s").Value
                                                    Debug.Print "codice comando  -> " & daoRS.Fields("CodCmd_s").Value
                                                    
                                                    CodCmd_s = daoRS.Fields("CodCmd_s").Value
                                                    
                                                    '//@imposto.tipo.oggetto
                                                    sTipo_Oggetto = daoRS.Fields("TIPOGGETTO_s").Value
                                                    

                                            '//------------------------------------------------------------------------------------------
                                    
                                         
                        
                                            
                                                                                
                                            
                                            
                                            '//06) SALVO IL TIPO OGGETTO NELLA VARIABILE + SALVO IL CODICE PROGETTO @todo:inserire un comando query di estrazione
                                            '//==========================================================================================================//
                                            
                                            '//@errore.corretto.22.05.2022 = durante il salvataggio delle tabelle, _
                                            macro, form query, report e moduli è stato corretto l'errore _
                                            in quanto occorre anche il salvataggio del codice progetto _
                                            che viene passato come parametro: par_sx_Name_COD_PROGETTO_s, IN QUESTO CASO IL _
                                            codice progetto viene recuperato con un comdando mediante la query particolare che estrae _
                                            il record con il campo progetto filtrato.
                                                                                        
                                         
                                            
                                                    sSql = ""
                                                    sSql = sSql & "SELECT MSys_TUTTI_PROG.*, MSys_TUTTI_PROG.Scel_b "
                                                    sSql = sSql & "FROM MSys_TUTTI_PROG "
                                                    sSql = sSql & "WHERE (((MSys_TUTTI_PROG.Scel_b) = True)) "
                                                    sSql = sSql & "WITH OWNERACCESS OPTION;"
                                                    
                                                    Debug.Print sSql

                                                     'Apro rs estrazione Sql
                                                     Set DaoRs_Filtro = CurrentDb.OpenRecordset(sSql)
                                                         
                                                           
                                                           '//estraggo il progetto attivo
                                                           '//--------------------------------------------------------------------------------
                                                             
                                                             If DaoRs_Filtro.EOF = False And DaoRs_Filtro.BOF = False Then
                                                                 
                                                                 '//estraggo il codice progetto
                                                                 daoRS.MoveFirst
                                                                    sName_COD_PROGETTO_s = DaoRs_Filtro.Fields("COD_PROGETTO_s").Value
                                                             
                                                             Else
                                                                '// non è stato scelto nessun progetto msgbox ed uscita
                                                                MsgBox "ATTENZIONE NON E' STATO SCELTO NESSUN PROGETTO - USCITA DALLA ROUTINE ", vbCritical, "MSGBOX"
                                                                DaoRs_Filtro.MoveNext
                                                                DaoRs_Filtro.Close
                                                                Set DaoRs_Filtro = Nothing
                                                                GoTo Exit_Cmb_01_Txt_AfterUpdate
                                                             End If
                                                             
                                                     DaoRs_Filtro.Close
                                                     Set DaoRs_Filtro = Nothing
                                            '//==========================================================================================================//
                                            
                                            
                                            '//ACCODA TUTTI GLI OGGETTI MSYS @attivo.menu.accoda.tutto.df12, @menu.accoda.tutto _
                                                        @menu.importa.tutto, @IMPORTA.TUTTO.
                                            '//=================================================================================================//
                                            '//SE ACCODA TUTTO DEVI ESEGUIRE LA ROUTINE ,
                                            '//Se vi è il comando importa tutto, TABELLE+FORM+QUERY+REPORT+MODULI ALLORA viene eseguita _
                                                la reiterazione del comando con il ciclo for ed alla fine il rs viene chiuso. Il comando _
                                                per attivare questa selezione completa degli oggetti è : _
                                                @03_IMPORTA_TUTTO_TABELLE+FORM+QUERY+REPORT+MODULI_IN_DB
                                            
                                                If CodCmd_s = "Cmd_Importa_TUTTO_TABELLE+FORM+QUERY+REPORT+MODULI+MACRO" Then
                                                         
                                                    sScelta_db = ""
                                                    sName_Tab = ""
                                                    sName_Query = ""
                                                    sName_Form = ""
                                                    sName_Report = ""
                                                    sName_MODULI = ""
                                                    sName_MACRO = ""
                                                    CodCmd_s = ""
                                                    iOperazione = 0
                                                    iIndice1 = 0
                                                    Flag = False
                                                     
                                                         
                                        
                                                             
                                                               
                                                              sTipo_OggettoMatrice(0) = "TABLE"
                                                              sTipo_OggettoMatrice(1) = "QUERY"
                                                              sTipo_OggettoMatrice(2) = "MODULI"
                                                              sTipo_OggettoMatrice(3) = "FORMS"
                                                              sTipo_OggettoMatrice(4) = "REPORTS"
                                                              sTipo_OggettoMatrice(5) = "MACRO"             '//procedura --->@IMPORTA_MACRO_IN_DB
                                                              
                                                 '//ACCODAMENTO COLLETTIVO DI TUTTI GLI OGGETTI _
                                                    QUERY+TABELLE+FORM+REPOR+MODULI+MACRO, attivando con il ciclo for per una matrice di 4 oggetti al fine di richimare 4 volte la funzione _
                                                    di gestione degli oggetti, alla fine per non richiamarla ulteriormente, chiuro il rs ed esco _
                                                    dalla funzione - @importa.tutto.tabelle+query+moduli+forms+reports+macros
                                                 For i = 0 To 5 Step 1
                                                        
                                                      
                                                        Vv1 = Iterazione_oggetti( _
                                                                        sScelta_db, _
                                                                        sName_Tab, _
                                                                        sName_Query, _
                                                                        sName_Form, _
                                                                        sName_Report, _
                                                                        sName_MODULI, _
                                                                        sName_MACRO, _
                                                                        sName_COD_PROGETTO_s, _
                                                                        sTipo_OggettoMatrice(i), _
                                                                        CodCmd_s, _
                                                                        iOperazione, _
                                                                        iIndice1, _
                                                                        Flag)
                                                 Next i
                                                 
                                                 MsgBox "ACCODA TUTTO, TABELLE, QUERY, FORM, REPORT E MODULI ESEGUITO", vbExclamation, "MSG ACCODA TUTTO"
                                                 '//con l'accodo collettivo esco dalla funzione chiudo vado all'ultimo rc e chiudo il rs.
                                                    daoRS.MoveLast
                                                    daoRS.Close
                                                    Set daoRS = Nothing
                                                    
                                                            '//08) AGGIORNO LA FORM CORENTE
                                                            'requery + refresh del record della form corrente
                                                            Me.Form.Requery
                                                            Me.Form.Refresh
                                                            
                                                 
                                                 GoTo Exit_Cmb_01_Txt_AfterUpdate
                                            End If
                                            '//*** FINE ***
                                            '//=================================================================================================//
                                            
                                            
                                            
                                            
                                            '//ACCODA TUTTI GLI OGGETTI IN DF12 @attivo.menu.accoda.tutto.df12, @menu.accoda.tutto, @accoda.tutto.df12
                                            '//=================================================================================================//
                                            '//SE ACCODA TUTTO DEVI ESEGUIRE LA ROUTINE ,
                                            '//Se vi è il comando accoda  tutto gli oggetti in DF12, TABELLE+FORM+QUERY+REPORT+MODULI ALLORA viene eseguita _
                                                la reiterazione del comando con il ciclo for ed alla fine il rs viene chiuso. @Cmd_Accoda_TUTTO_TO_Msys_DF12 _
                                                @04_ACCODA_TUTTO_TO_Msys_DF12_TABELLE+FORM+QUERY+REPORT+MODULI+MACRO

                                            
                                                If CodCmd_s = "Cmd_Accoda_TUTTO_TO_Msys_DF12" Then
                                                         
                                                    sScelta_db = ""
                                                    sName_Tab = ""
                                                    sName_Query = ""
                                                    sName_Form = ""
                                                    sName_Report = ""
                                                    sName_MODULI = ""
                                                    sName_MACRO = ""
                                                    sName_COD_PROGETTO_s = ""
                                                    CodCmd_s = ""
                                                    iOperazione = 0
                                                    iIndice1 = 0
                                                    Flag = False
                                                     
                                                         
                                                
                                                               
                                                              sTipo_OggettoMatrice(0) = "Cmd_Accoda_Tabelle_TO_Msys_DF12"
                                                              sTipo_OggettoMatrice(1) = "Cmd_Accoda_Report_TO_Msys_DF12"
                                                              sTipo_OggettoMatrice(2) = "Cmd_Accoda_Query_TO_Msys_DF12"
                                                              sTipo_OggettoMatrice(3) = "Cmd_Accoda_MODULI_TO_Msys_DF12"
                                                              sTipo_OggettoMatrice(4) = "Cmd_Accoda_Forms_TO_Msys_DF12"
                                                              sTipo_OggettoMatrice(5) = "Cmd_Accoda_Macro_TO_Msys_DF12"
                                                              
                                                 '//ACCODAMENTO COLLETTIVO DI TUTTI GLI OGGETTI _
                                                    QUERY+TABELLE+FORM+REPOR+MODULI, attivando con il ciclo for per una matrice di 4 oggetti al fine di richimare 4 volte la funzione _
                                                    di gestione degli oggetti, alla fine per non richiamarla ulteriormente, chiudo il rs ed esco _
                                                    dalla funzione
                                                    
                                                 For i = 0 To 5 Step 1
                                                        

                                                        Vv1 = Iterazione_oggetti( _
                                                                        sScelta_db, _
                                                                        sName_Tab, _
                                                                        sName_Query, _
                                                                        sName_Form, _
                                                                        sName_Report, _
                                                                        sName_MODULI, _
                                                                        sName_MACRO, _
                                                                        sName_COD_PROGETTO_s, _
                                                                        sTipo_Oggetto, _
                                                                        sTipo_OggettoMatrice(i), _
                                                                        iOperazione, _
                                                                        iIndice1, _
                                                                        Flag)
                                                 Next i
                                                 
                                                 MsgBox "IMPORTATO TUTTO SULLA TABELLA F12 TABELLE, QUERY, FORM, REPORT E MODULI ESEGUITO", vbExclamation, "MSG ACCODA TUTTO"
                                                 '//con l'accodo collettivo esco dalla funzione chiudo vado all'ultimo rc e chiudo il rs.
                                                    daoRS.MoveLast
                                                    daoRS.Close
                                                    Set daoRS = Nothing
                                                    
                                                            '//08) AGGIORNO LA FORM CORENTE
                                                            'requery + refresh del record della form corrente
                                                            Me.Form.Requery
                                                            Me.Form.Refresh
                                                            
                                                 
                                                 GoTo Exit_Cmb_01_Txt_AfterUpdate
                                            End If
                                            '//*** FINE ***
                                            '//=================================================================================================//
                                            
                                            
                                            
                                            
                                            '//07) CHIAMO LA FUNZIONE DI ITERAZIONE OGGETTI DEL DB CORRENTE - MODULO DA UTILIZZARE ----> UTIL_Nro40_N01_CONTROLLO_OGGETTI_DB
                                            '//faq: @controllo.oggetti, @chiama.controllo oggetti, @attivo.il.controllo.oggetti
                                            '//Funzione ITERAZIONE OGGETTI DEL DATABASE MDB CORRENTE
                                            '//PARAMETRI IN IMPUT:
                                            '//par_sx_Scelta_db =database ;                         = NON ATTIVATO          ----> _
                                               par_sx_Name_Tab                                      = ATTIVATO              ---> nome tabella da ricercare; _
                                               par_sx_Name_Query                                    = ATTIVATO              ---> nome query da ricercare _
                                               par_sTipo_Oggetto                                    = NON ATTIVATO          ---> tipo oggetto ; _
                                               par_ix_Operazione                                    = NON ATTIVATO          ----> Tipo di operazione =?  ; _
                                               par_ix_Indice1                                       = NON ATTIVATO          ----> ? ; _
                                               par_bx_Flag                                          = NON ATTIVATO          ----> True individuato ??
                                            '//OUTPUT : RESTITUISCE una stringa
                        
                                            Vv1 = Iterazione_oggetti(sScelta_db, _
                                                                    sName_Tab, _
                                                                    sName_Query, _
                                                                    sName_Form, _
                                                                    sName_Report, _
                                                                    sName_MODULI, _
                                                                    sName_MACRO, _
                                                                    sName_COD_PROGETTO_s, _
                                                                    sTipo_Oggetto, _
                                                                    CodCmd_s, _
                                                                    iOperazione, _
                                                                    iIndice1, _
                                                                    Flag)
                                            
                                                'Assegno alla Sottform la query - non funziona perchè restituisce
                                                ' un boolean. Devi modificarlo...
                                                If Vv1 > "" Then
                                                'Me.SottoForm_01.Form.RecordSource = Vv1
                                                'Application.Forms(0)!LBL_TITOLO_1.Caption = "CONTROLLO QUERY"
                        
                                                End If
                                            
                                            '//08) AGGIORNO LA FORM CORENTE
                                            'requery + refresh del record della form corrente
                                            Me.Form.Requery
                                            Me.Form.Refresh
                                            
                                        '______________________________________________________________________________
                        
                        

                                    Else
                                            '//se non si trova la comparazione tra comando scelto e record _
                                            vado all'ultimo record per uscire
                                            daoRS.MoveLast
                                            
                                    End If
                            
                            daoRS.MoveNext
                            
                            Wend
                            daoRS.Close
                            Set daoRS = Nothing
                            
                         
                      '//ITERO NELLA TABELLA - CONTROLLA IL COMANDO SCELTO *** FINE ***
                      '//--------------------------------------------------------------------------------
                            
                                  '..........................................................
                                  
                                  
                                  
                               
                                
                                    
                        End If
                                         
                      '_________________________________________________________________
                    
                    
    
                
                
        
    
                            
       '//ATTIVA IL COMANDO DELLA COMBINATA - @MENU.CMD_01  *** FINE ***
       '//=========================================================================================================================//
       
    
         
    


Exit_Cmb_01_Txt_AfterUpdate:
Exit Sub

Err_Cmb_01_Txt_AfterUpdate:
MsgBox Err.Description
Resume Exit_Cmb_01_Txt_AfterUpdate

End Sub


'//??? TODO:@da fare.  @ESPORTA.PROGETTO,  @TODO.PROGETTO.ESPORTA
Private Sub PROGETTO_DA_ESPORTARE()
Dim par_sxItemText As String
                '//@da.finire
                'ESPORTA LE TABELLE IN UN DATABASE ESTERNO - TODO : DA FARE  @da.finire, @da.fare
                '...........................................................................................
                        
                        '//PROGETTO DA ESPORTARE con oggetti salvati nella tabella MSsys_ELENCO_OGGETTI_DISCO_ESPORTA che sono da esportare, _
                           in un db esterno, ma si deve ancora implementare l'sportazione.
                          
                        If par_sxItemText > "" Then
                            
                             'APRO IL RS
                                sSql = ""
                        
                                sSql = sSql & "SELECT MSsys_ELENCO_OGGETTI_MSsys.* "
                                sSql = sSql & "FROM MSsys_ELENCO_OGGETTI_MSsys "
                                sSql = sSql & "WHERE (((MSsys_ELENCO_OGGETTI_MSsys.NOME_OGGETTO_s)='" & par_sxItemText & "'));"
                                
                        End If
                                   'Controllo
                                   Debug.Print sSql
                '...........................................................................................
                

End Sub




'//**************************************************************************************************************************//
'//         @LE.FUNZIONI.PUBBLICHE


'****************************************************************************************************************************
'           PROCEDURA CARICAMENTO OGGETTI MSys_TUTTE_LE_PAGINE
'//PARAMETRI                    = NULL
'//@funzione.carica.pagine
'
'
'****************************************************************************************************************************


'// FUNZIONE AGGIORNA PAGINE
'//--------------------------------------------------------------------------------
'//NOTE: funzione pubblica di iterazione nella tabella MSys_TUTTE_LE_PAGINE che _
        permette il caricamento delle pagine e delle sottoform attive in nel _
        tabcontroll della form master. @tab.controll.carica.pagine
'//
'//FUNZIONE             --> CALL MSys_TUTTE_LE_PAGINE_Carica_pFunction
'//PARAMETRI            --> Null
'//CODICE               --> @FUNZIONE@CARICA@PAGINE@TAB@CONTROLL, @AGGIORNA.IL.PROGETTO
'//--------------------------------------------------------------------------------
    
Public Function MSys_TUTTE_LE_PAGINE_Carica_pFunction()

Dim mysName_COD_PROGETTO_s As String              '//codice progetto da utilizzare per il filtro

On Error GoTo Err_MSys_TUTTE_LE_PAGINE_Carica_pFunction
                              


            '// 2024_09_04_SOTTOFORM_MASTER_SENZA_TITOLO = CREATA LA SOTTOFORM MASTER A CARICAMENTO AUTOMATICO DI 12 PAGINE DEL TAB CONTROLL _
                creato una sottoform Master abilitata al caricamento automatico di altre sottoform. La peculiarita che la _
                sottoform master senza titolo è quella di poter caricare 12 sottoform e quindi rendere il caricamento _
                automatico senza obbligo di codice statico. _



                        '//ATTIVA LA FUNZIONE DI RECUPERO DATI NELLA TABELLA
                        '==================================================================================================
                        'Note: Leggo i dati dalla tabella _
                        @FUNZIONE@CARICA@PAGINE@TAB@CONTROLL_(eseguo un ciclo sulla tabella MSys_TUTTI_PROG per trovare il progetto selezionato a True da caricare)

                        
                        
                                '//CARICO LE SOTTOFORM MASTER
                                'Apro rs estrazione Sql CON LA TABELLA DEI PROGETTI PER _
                                    individuare quale progetto è stato scelto da caricare.
                                Set daoRS = CurrentDb.OpenRecordset("SottoFormMaster_GE_Msys")
                                
                             If daoRS.EOF = False And daoRS.BOF = False Then
                                     daoRS.MoveFirst
                                    '//APRO LA TABELLA PROGETTO E SCELGO IL PROGETTO
                                    While Not daoRS.EOF
                                           
                                            If daoRS.Fields("Scel_b").Value = True Then
                                                    '//salvo il codice progetto
                                                    mysName_COD_PROGETTO_s = daoRS.Fields("COD_PROGETTO_s").Value
                                                    
                                                    '//se trovo il progetto vado alla finre record per evitare _
                                                     il msgbox di fallimento di individuazione del progetto
                                                     daoRS.MoveLast
                                                    '//stampo il progetto da caricare
                                                    Debug.Print " progetto da caricare ----> " & mysName_COD_PROGETTO_s
                                    
                                          
                                            End If
                                            
                                            
                                         daoRS.MoveNext
                                    Wend
                                    
                            End If
                                
                                    If mysName_COD_PROGETTO_s = "" Then
                                    DoEvents
                                                '//se non esiste il progetto da caricare, valorizzato con True avviene il msg di errore _
                                                    + uscita dalla routine
                                                MsgBox "ATTENZIONE IL PROGETTO DA CARICARE NON ESISTE : " & mysName_COD_PROGETTO_s, vbExclamation, " MSG PROGETTO, seleziona DALLA QUERY UN PROGETTO! "
                                                DoCmd.OpenQuery "MSys_TUTTI_PROG_SELECT_Qry01_01_PROGETTI"
                                                
                                                '//per uscire dal rs vado all'ultimo record + msg di avviso
                                                daoRS.MoveLast
                                                Debug.Print "ERRORE PROGGETTO VEDI QUESTA FUNZIONE :  Function MSys_TUTTE_LE_PAGINE_Carica_pFunction()"
                                                        
                                                        '//uscita dalla routine in quanto non esiste nessun progetto da caricare
                                                        GoTo Exit_MSys_TUTTE_LE_PAGINE_Carica_pFunction
                             
                                   End If
                                  
                             '//APRO LA TABELLA DELLE PAGINE
                             '//@tabella.comandi = apro la tabella che contiene le pagine da caricare, gia filtrata _
                                con la variabile del PROGETTO  scelto da caricare. _
                                @TABELLA.CARICA.PAGINE, @PAGINE.DA.CARICARE _
                                @PAGINE.FILTRATE.DA.CARICARE, @AGGIORNA.IL.PROGETTO.SQL
                           
                           

                                        sSql = ""
                                        sSql = sSql & "SELECT "
                                        sSql = sSql & "SottoFormMaster_Vis_Pagine_Msys.COD_PROGETTO_s, "
                                        sSql = sSql & "SottoFormMaster_Vis_Pagine_Msys.NRO_OGGETTO_i, "
                                        sSql = sSql & "SottoFormMaster_Vis_Pagine_Msys.TIPOGGETTO_s, "
                                        sSql = sSql & "SottoFormMaster_Vis_Pagine_Msys.NOMEOGGETTO_s, "
                                        sSql = sSql & "SottoFormMaster_Vis_Pagine_Msys.NOMEpagina_s, "
                                        sSql = sSql & "SottoFormMaster_Vis_Pagine_Msys.VISIBLE_b, "
                                        sSql = sSql & "SottoFormMaster_Vis_Pagine_Msys.NOTA_OGGETTO_s, "
                                        sSql = sSql & "SottoFormMaster_Vis_Pagine_Msys.NOTEex_s, "
                                        sSql = sSql & "SottoFormMaster_Vis_Pagine_Msys.ID "
                                        sSql = sSql & "FROM SottoFormMaster_Vis_Pagine_Msys "
                                        sSql = sSql & "WHERE (((SottoFormMaster_Vis_Pagine_Msys.COD_PROGETTO_s)='" & mysName_COD_PROGETTO_s & "'))"
                                        sSql = sSql & "ORDER BY SottoFormMaster_Vis_Pagine_Msys.NRO_OGGETTO_i "
                                        sSql = sSql & "WITH OWNERACCESS OPTION; "
                                    '//controllo stringa ssql
                                    Debug.Print sSql

                            
                            'Apro rs estrazione Sql
                            Set daoRS = CurrentDb.OpenRecordset(sSql)
                                
                                  
                                  '//ITERO NELLA TABELLA - CONTROLLA IL COMANDO SCELTO
                                  '//--------------------------------------------------------------------------------
                                  '//Note : resetto ed imposto le variabili
                                  
                                  
                                                        'RESET variabili pagine da caricare
                                                        SottoForm_s = ""
                                                        NameSottoForm_s = ""
                                                        VisibleSottoForm_b = False
                                                        LabelPagina_s = ""
                                                        
                                    '//se rs esiste
                                    If daoRS.EOF = False And daoRS.BOF = False Then
                                        
                                        daoRS.MoveFirst
                                        
                                        
                                        While Not daoRS.EOF
                                        
                                                        
                                        
                                        
                                              'Imposto le variabili parametro per caricare le pagine
                                              '.....................................................................................
                                              
                                              '//TODO:FINIRE occorre impostare true al progetto è lo stesso di quello salvato _
                                                        nella variabile attivo il caricamento   @attivo.il.caricamento.pagine
                                                        
                                                    '@FUNZIONE@CARICA@PAGINE@TAB@CONTROLL_(recupero i parametri CODICE PROGETTO TIPO OGGETTO NOME OGGETTO _
                                                        NOME PAGINA VIBILE e chiamo la funzione imposta le sottoform)
                                              
                                              '//controllo preventivo della form da caricare:
                                              Debug.Print "FORM DA CARICARE: " & daoRS.Fields("NOMEOGGETTO_s").Value

                                              If daoRS.Fields("COD_PROGETTO_s").Value = mysName_COD_PROGETTO_s Then

                                                        SottoForm_s = daoRS.Fields("TIPOGGETTO_s").Value
                                                        
                                                        If IsNull(daoRS.Fields("NOMEOGGETTO_s").Value) = True Then
                                                            NameSottoForm_s = ""
                                                        Else
                                                            NameSottoForm_s = daoRS.Fields("NOMEOGGETTO_s").Value
                                                        End If
                                                        
                                                        
                                                        If IsNull(daoRS.Fields("NOMEpagina_s").Value) = True Then
                                                            LabelPagina_s = ""
                                                        Else
                                                            
                                                            LabelPagina_s = daoRS.Fields("NOMEpagina_s").Value
                                                        End If

                                                        
                                                        VisibleSottoForm_b = daoRS.Fields("VISIBLE_b").Value
                                                        
                                                        
                                                        
                                              '.....................................................................................
                                              
                                                
                                                       'modelo Call Imposta_N22_IMPOSTA_LE_SOTTOFORM_Sub("SottoForm_03", "MSsys_Frm01_S02_COLLEGA_TABELLE", True, "03 COLLEGA TABELLE") _
                                                       @CHIAMO@LA@ROUTIN@IMPOSTA@LE@SOTTOFORM_(per impostare le sottoform e passo 4 parametri)

                                                       
                                                       Call Imposta_N22_IMPOSTA_LE_SOTTOFORM_Sub(SottoForm_s, NameSottoForm_s, VisibleSottoForm_b, LabelPagina_s)
                                               
                                               End If '//If DaoRs.Fields("sName_COD_PROGETTO_s").Value = mysName_COD_PROGETTO_s Then
                                             '______________________________________________________________________________
                                             
                                             
                                                                                                   
                                                                        
                                                                'controllo oggetto
                                                                'Debug.Print "COMANDO         -> " & DaoRs.Fields("COMANDO_s").Value
                                                                Debug.Print "TIPO DI OGGETTO -> " & daoRS.Fields("TIPOGGETTO_s").Value
                                                                'Debug.Print "codice comando  -> " & DaoRs.Fields("CodCmd_s").Value
                                                                
                                                                'CodCmd_s = DaoRs.Fields("CodCmd_s").Value
                                                                
                                                                
                                                        '//------------------------------------------------------------------------------------------
                                                
                                                     
                                    
                                                
                                                        
                                                        'requery + refresh del record della form corrente
                                                        Me.Form.Requery
                                                        Me.Form.Refresh
                                                        
                                                    '______________________________________________________________________________
                                    
                                    
                        
                                                    
                                                
                                        daoRS.MoveNext
                                        
                                        Wend
                                        daoRS.Close
                                        Set daoRS = Nothing
                                        
                                     
                                  '//ITERO NELLA TABELLA - CONTROLLA IL COMANDO SCELTO *** FINE ***
                                  '//--------------------------------------------------------------------------------
                                        
                                    Else
                                        '//@ERRORE@RECORD@VUOTO, se esiste il progetto da caricare, MA NON HA FORM DA IMPLEMENTARE allora il msg di errore _
                                            + uscita dalla routine,
                                        DoCmd.OpenQuery "MSys_TUTTI_PROG_SELECT_Qry01_01_PROGETTI", acViewNormal
                                        
                                        
                                        MsgBox "ATTENZIONE RECORD VUOTO NON CI SONO FORM DA CARICARE: " & mysName_COD_PROGETTO_s, vbExclamation, " MSG PROGETTO, seleziona DALLA QUERY UN PROGETTO CON LE FORM! "
                                        

                                    End If
                        
                        '//ATTIVA IL COMANDO DELLA COMBINATA
                        '==================================================================================================
                                



Exit_MSys_TUTTE_LE_PAGINE_Carica_pFunction:
Exit Function

Err_MSys_TUTTE_LE_PAGINE_Carica_pFunction:
MsgBox Err.Description
Resume Exit_MSys_TUTTE_LE_PAGINE_Carica_pFunction

End Function
'//=============================================================================================//


            





        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        '                                       TAB CONTROL
        ' @TabControllo@Evento@change_(@gestione @eventi del tab Controll, @cambio@pagina)
        '//@Trova@Progetti_(comando DOS apri trova @progetti@tutti)
        '//START "apri il progetto dei comandi di TUTTI@PROGETTI" "c:\\CASA\\LINGUAGGI\\TROVA_PROGETTI\\TROVA_PROGETTI.sublime-project"
        
        
        '//@APRI@PROGETTO@TAB@CONTROLL_(la path ed i comando DOS, per la routine di gestione del TAB CONTROL e le pagine che vengono aggiornate con @page@requery)
        
        '// START "apri il progetto TUTORIAL IN INGLESE" "c:/CASA/LINGUAGGI/INGLESE/ENGLISH_TUTORIAL/INGLESE_TUTORIAL.sublime-project"

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
            
            
                On Error GoTo Err_TabCtl_01_Change
        
        
                            'CONTROLLO SULLA FORM PRINCIPALE @controllo@pagine_(@routine per il controllo delle @pagine@a@schede)
                            '......................................................................................
                            ' La proprieta @IsLoaded Determina se la FORM  esiste e se è
                            ' attualmente aperta impostando la variabile booleana.
                                                    
                                    Dim ixPage As Integer                                           'variabili indice delle paginae
                                    Dim bFormAperta As Boolean                                      'variabile che indica se la form principale è aperta
                            
                            
                                                    
                                    
                                    'CONTROLLO OGGETTI DEL DABATABASE
                                    '---------------------------------------------------------------------
                                    'nota : @form@aperta_(Controllo se la form base di gestione è aperta.)
                                    
                                        
                                            Dim obj As AccessObject, dbs As Object
                                               Set dbs = Application.CurrentProject
                                               ' Search for open AccessObject objects in AllForms collection.
                                               For Each obj In dbs.AllForms
                                                   If obj.IsLoaded = True Then
                                                       ' Print name of obj la form è aperta.
                                                       Debug.Print "la form aperta IsLoaded= True " & obj.Name
                                                       
                                                       sxFrm_CHIAMANTE_GENERALE = obj.Name
                                                       
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
                                sxFrm_CHIAMANTE_GENERALE_CORRENTE = Me.Form.Name
                                
                                'ixPage = Me.TabCtl_01
                                
                                'chiamo la funzione
                                'sxPagCorrente_Caption = pfunc_Recupero_NOME_paginaA_CORRENTE(sxFrm_CHIAMANTE_GENERALE_CORRENTE , 1)
                            '............................................................................
                            
                            
                            '//@tab@controll@indice@pagina_(controllo dell'indice delle pagine del controllo a schede mediante @select@case@page)
                            If Me.TabCtl_01.Value = 0 Then
                           
                            
                                    
                                    'PG_01 INDICE_PAGINA = 0
                                    '............................................................................
                                    'Aggiorno la pagina 1 con il requery
                                    
                                    
                                         'REQUERY DELLA SOTTOFORM_1
                                         Me.SottoForm_01.Requery
                                         
                                         
                                         '//2024_08_23_CORREZIONE_ERRORE_PAG_01-12 = corretto errore delle sottoform non caricate con recordsource = null
                                         If Me.SottoForm_01.SourceObject > "" Then
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
                                         
                                         
                                         '//2024_08_23_CORREZIONE_ERRORE_PAG_01-12 = corretto errore delle sottoform non caricate con recordsource = null
                                         If Me.SottoForm_02.SourceObject > "" Then
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
                                         
                                         
                                         '//2024_08_23_CORREZIONE_ERRORE_PAG_01-12 = corretto errore delle sottoform non caricate con recordsource = null
                                         If Me.SottoForm_03.SourceObject > "" Then
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
                                         
                                         
                                         '//2024_08_23_CORREZIONE_ERRORE_PAG_01-12 = corretto errore delle sottoform non caricate con recordsource = null
                                         If Me.SottoForm_04.SourceObject > "" Then
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
                                         
                                         
                                         '//2024_08_23_CORREZIONE_ERRORE_PAG_01-12 = corretto errore delle sottoform non caricate con recordsource = null
                                         If Me.SottoForm_05.SourceObject > "" Then
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
                                         
                                         
                                         '//2024_08_23_CORREZIONE_ERRORE_PAG_01-12 = corretto errore delle sottoform non caricate con recordsource = null
                                         If Me.SottoForm_06.SourceObject > "" Then
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
                                         
                                         
                                         '//2024_08_23_CORREZIONE_ERRORE_PAG_01-12 = corretto errore delle sottoform non caricate con recordsource = null
                                         If Me.SottoForm_07.SourceObject > "" Then
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
                                         
                                         
                                         '//2024_08_23_CORREZIONE_ERRORE_PAG_01-12 = corretto errore delle sottoform non caricate con recordsource = null
                                         If Me.SottoForm_08.SourceObject > "" Then
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
                                         
                                         
                                         '//2024_08_23_CORREZIONE_ERRORE_PAG_01-12 = corretto errore delle sottoform non caricate con recordsource = null
                                         If Me.SottoForm_09.SourceObject > "" Then
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
                                                                                  
                                         '//2024_08_23_CORREZIONE_ERRORE_PAG_01-12 = corretto errore delle sottoform non caricate con recordsource = null
                                         If Me.SottoForm_10.SourceObject > "" Then
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
                                         
                                         
                                         '//2024_08_23_CORREZIONE_ERRORE_PAG_01-12 = corretto errore delle sottoform non caricate con recordsource = null
                                         If Me.SottoForm_11.SourceObject > "" Then
                                            Me.SottoForm_11.Form.Refresh
                                         End If
                                        
                                        
                                    '............................................................................
                    
                        '____________________________________________________________________________________________________________
                    
                        'Pag_12
                        '____________________________________________________________________________________________________________
                            ElseIf Me.TabCtl_01.Value = 11 Then
                        
                                    'PG_12 INDICE_PAGINA = 11
                                    '............................................................................
                                    'Aggiorno la pagina 11 con il requery
                                    
                                    
                                         'REQUERY DELLA SOTTOFORM_12
                                         Me.SottoForm_12.Requery
                                         
                                         '//2024_08_23_CORREZIONE_ERRORE_PAG_12 = corretto errore della sottoform non caricata
                                         If Me.SottoForm_12.SourceObject > "" Then
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







