VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_PROGETTI_Msys_TB01Frm01_01_PROJECT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<7<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'                                   LE VARIABILI GENERALI
'//LA FORM DI TIPO SOTTOFORM _
    CLASSE Form: PROGETTI_Msys_TB01Frm01_01_PROJECT

'//
'// FAQ
'//     QUALE E' LA FORM DI GESTIIONE DEI CODICI PROGETTO
'//
'//
'//
'//
'//
'//
'//
'//
'//
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<






'//@VARIABILI@GENERALI
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

Option Compare Database





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
    Dim sSql As String


    'ADO
    Dim ADODB As Database
    Dim AdodaoRs As Recordset
    
    
    'Contatori
    Dim icount As Integer
    Dim dbl_count As Double
    Dim i As Integer
    
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





Private Sub Form_Load()

'//SU CARICAMENTO IMPOSTO IL NOME DELLA FORM CORRENTE ED IL REQUERY
'//....................................................................//
'//NOTE:  assegno il nome della form corrente al campo NOME FORM _
        E LA VISUALIZZO.
        
      myNameForm = Me.Name
      Debug.Print Vv1
      '"01) LLPP_IMPEGNI_Frm_S01_12_GRUPPO_IMPEGNI_ANNO - MASTER
      NameForm_s_Txt = "01) " & myNameForm & " - FORM MASTER"
      Me.Requery
      Me.Recalc

'//....................................................................//

End Sub


Private Sub Form_Open(Cancel As Integer)
Me.RecordSource = "PROGETTI_Msys_TB01Qry01_01_PROJECT"
Me.Recalc

End Sub


'//HELP
Private Sub Help_Txt_DblClick(Cancel As Integer)
    MsgBox "help da definire???", vbCritical, "HELP PROGETTI"

End Sub









'//==================================================================================================//
'//             EVENTI CASELLA DI TESTO @COD_PROGETTO_s_(@CODICE@PROGETTO)   *** INIZIO ***
'//==================================================================================================//
'//EVENTI: _
    '//01_COD_PROGETTO_s_TXT_DblClick       = AGGIORNO LA TABELLA TMP E LA CASELLA DI TESTO FILTRO
    '//02_COD_PROGETTO_s_DblClick.?         = ?
    '//03_COD_PROGETTO_s_TXT.CHANGE         = AGGIORNO I CAMPI TXT PROGETTO E NOTE PROGETTO




'//01_COD_PROGETTO_s_TXT_DblClick       = AGGIORNO LA TABELLA TMP E LA CASELLA DI TESTO FILTRO
Private Sub COD_PROGETTO_s_TXT_DblClick(Cancel As Integer)

    
    '//01_COD_PROGETTO_s_TXT_DblClick.DELETE = CANCELLO I DATI NELLA TABELLA TMP
        DoCmd.OpenQuery "PROGETTI_Msys_TB01Qry01_02_PROJECT_TMP_Delete"
        
        
    ' Inserisce il valore di COD_PROGETTO_s nella tabella PROGETTI_Msys_TB01_PROJECT_TMP
'sSql = "INSERT INTO PROGETTI_Msys_TB01_PROJECT_TMP (COD_PROGETTO_s) VALUES ('" & Me.COD_PROGETTO_s_TXT & "')"

sSql = ""
sSql = sSql & "INSERT INTO PROGETTI_Msys_TB01_PROJECT_TMP ( COD_PROGETTO_s ) "
sSql = sSql & "SELECT '" & Me.COD_PROGETTO_s_TXT & "' AS COD_PROGETTO;"

'//CONTROLLO ED ESECUZIONE
Debug.Print sSql
CurrentDb.Execute sSql
    
        
        
    '//02_COD_PROGETTO_s_DblClick.UPDATE = AGGIORNO I CAMPI TMP CON L'ARCHIVIO _
        Aggiorno la tabella tmp e aggiorno la casella TXT_FILTRO
        DoCmd.OpenQuery "PROGETTI_Msys_TB01Qry01_03_PROJECT_TMP_UPDATE"
        
        Me.Parent!TXT_FILTRO = Me.COD_PROGETTO_s_TXT
        
        
        '//Imposto filtro sottoform_01 corrente dopo la scelta
        '//Esempio di impostazione della proprieta filter: Me.Filter = "Stato = 'USA'"
        Vv1 = Me.Parent!TXT_FILTRO
        
        Me.Form.filter = "COD_PROGETTO_s ='" & Vv1 & "'"
        
        Me.Form.FilterOn = True
        Me.Recalc
        Me.Requery
    
    
    
End Sub


'//03_COD_PROGETTO_s_TXT.CHANGE         = AGGIORNO I CAMPI TXT PROGETTO E NOTE PROGETTO
Private Sub COD_PROGETTO_s_TXT_Change()
        '//aggiorno i due campi testo.
        Me.PROGETTO_s_TXT = Me.COD_PROGETTO_s_TXT.Column(1)
        Me.NOTE_PROGETTO_s_Txt = Me.COD_PROGETTO_s_TXT.Column(2)
        Me.GROUP_COD_PROG_s_TXT = Me.COD_PROGETTO_s_TXT.Column(0)
End Sub


'//==================================================================================================//
'//             EVENTI CASELLA DI TESTO @COD_PROGETTO_s_(@CODICE@PROGETTO)   *** FINE ***
'//==================================================================================================//





'//==================================================================================================//
'//             EVENTI CASELLA DI TESTO @PROGETTO_s_(PROGETTO)   *** INIZIO ***
'//==================================================================================================//

Private Sub PROGETTO_s_TXT_DblClick(Cancel As Integer)
    '//APRO LA TABELLA DEI PROGETTI
    DoCmd.OpenTable "Msys_DF11_PROGETTI"
End Sub


'//==================================================================================================//
'//             EVENTI CASELLA DI TESTO @PROGETTO_s_(PROGETTO)   *** FINE ***
'//==================================================================================================//
