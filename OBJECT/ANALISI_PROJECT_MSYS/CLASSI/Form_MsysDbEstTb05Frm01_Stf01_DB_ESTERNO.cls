VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_MsysDbEstTb05Frm01_Stf01_DB_ESTERNO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'//Oggetti per il funzionamento : _
    MODULI: _
        UTIL_MDL41_01_CONTROLLO_OGGETTI_QUERY_DEL_DB _



'//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++//
'//       LE VARIABILI DI MODULO


'//LE VARIABILI DATABASE
'//....................................................................//
    Dim daoDB As DAO.Database          '//occorre il riferimento alla libreria  Microsoft DAO 3.6 Object Libray
    Dim DaoWks As DAO.Workspace
    Dim daoRS As DAO.Recordset

    Dim ADODB As Database
    Dim AdodaoRs As Recordset
    Dim sSql As String                          '//STRINGA SQL
    'Dim Path_s As String                        '//la path


    '//Contatori
    Dim iCount As Integer
    Dim dbl_count As Double
    
   
    'Le variabili generiche
    Dim Vv1 As Variant
    Dim Dbl1 As Double
    Dim Int1 As Integer
    Dim Long1 As Long
    Dim Str1 As String
    Dim Bool1  As Boolean
    
    '//Messaggi di errore
    Dim ProceduraMessaggioErrore_s As String    '//Errore procedura
    Dim ProceduraAttivaEseguita_s As String     '//Errore Attivita eseguita


Private Sub Attribute_Txt_DblClick(Cancel As Integer)
'//apro gli attributi
        DoCmd.OpenQuery "MSsys_DF05_Qry01_ATTRIBUTI"

End Sub

'....................................................................

'//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++//




Private Sub Scel_b_AfterUpdate()
    '//TODO: DA FINIRE --- ??NON E'AMMESSA LA SELEZIONE SE il campo TIPO OGGETTO = NULL _
       ATTENZIONE occorre il tipo oggetto valorizzato altrimenti non puoi AGGIORNARE IL CAMPO SCELTO _
       @blocco.Scel_b.record.nullo; SE ESISTE ALMENO UN RECORD viene attivata la selezione _
       vedi codice @SELEZIONA.TRUE.TUTTO.MSys_QUERY.Scel_b
    
    
End Sub




Private Sub Scel_b_DblClick(Cancel As Integer)
    
        
    '//SELEZIONA COLLETTIVA MSys_TABELLE
    '//--------------------------------------------------------------------------------//-------//
    '//note : seleziona tutto a True se gia impostata altrimente imposta tutto a False _
              a condizione che la tabella è valorizzata per cui non è ammessa la selezione _
              se il campo TIPO_OGGETTO  è a null; occorrono almeno 1 record presente _
              vedi QUESTA CHIAVE ----> @blocco.Scel_b.record.nullo. _
              @SELEZIONA.TRUE.TUTTO.MSys_TABELLE.Scel_b
        
        
        If Me.Scel_b = False Then
        DoCmd.RunSQL "UPDATE " & _
                     "MSys_TABELLE " & _
                     "SET " & _
                     "MSys_TABELLE.Scel_b = True " & _
                     "WITH OWNERACCESS OPTION;"
                     
                                 
                     Me.Refresh
                    '//refresh + recalc = perche altrimenti blocca il record corrente
                    Me.Recalc
                    
        End If
        
        If Me.Scel_b = True Then
        DoCmd.RunSQL "UPDATE " & _
                     "MSys_TABELLE " & _
                     "SET " & _
                     "MSys_TABELLE.Scel_b = false " & _
                     "WITH OWNERACCESS OPTION;"
                     '//refresh + recalc = perche altrimenti blocca il record corrente
                     Me.Refresh
                     Me.Recalc

        End If
    '//--------------------------------------------------------------------------------//-------//
    
    

End Sub
