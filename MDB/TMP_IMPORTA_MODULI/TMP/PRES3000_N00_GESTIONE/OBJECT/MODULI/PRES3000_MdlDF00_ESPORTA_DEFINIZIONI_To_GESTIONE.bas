Attribute VB_Name = "PRES3000_MdlDF00_ESPORTA_DEFINIZIONI_To_GESTIONE"
Option Compare Database

'########################################################################################
'
'LE FUNZIONI ESPORTA OGGETTI
'
'NOTA: Le funzioni pubbliche esportano nel db di destinazione i seguenti oggetti:
'Tabelle, Query Maschere, Report, Pagine di proprieta, Macro e Moduli
'@modulo@esporta@oggetti@calendario_(ROUTINE DI EPORTAZIONE DEGLI OGGETTI RELATIVI AL CALENDARIO sul progetto Gestione Calendario)
'@project@calendario_(@esporta@oggetti@calendario e le definizioni, @modulo per l'esportazione)
'@esempio@di@esportazione@oggetti
'@macro@esporta@oggetti
'########################################################################################



'------------------------------------------------------------
' ESPORTA_OGGETTI_GESTIONE
'
'Nota: Esporto le qry nei db.
'------------------------------------------------------------
Public Function ESPORTA_OGGETTI_GESTIONE()

On Error GoTo ESPORTA_OGGETTI_GESTIONE_Err

Dim sxPath As String
Dim sxNAME As String
Dim sxFILE As String
Dim iCount As Integer                   '//contatore DEGLI OGGETTI ESPORTATI'

'ESPORTA IN FILE.MDB
'____________________________________________________________________________________________________________________________________________

        'MODELLI
        'DoCmd.CopyObject sxPATH, "", acQuery, ""
        'DoCmd.CopyObject sxPATH, "", acTable, ""
        'DoCmd.CopyObject sxPATH, "", acForm, ""
        'DoCmd.CopyObject sxPATH, "", acModule, ""
   
      
        MsgBox "MACRO DI ESPORTAZIONE ATTIVATA", vbExclamation, "ESPORTO OGGETTI DEL CALENDARIO NEL PROGETTO DI GESTIONE DELLE PRESENZE ---> PRES3000_N00_GESTIONE.mdb"
   
    
    'reset
    sxPath = "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\PRES3000_GESTIONE\ARCHIVI_PRES3000\PRES3000_N00_GESTIONE\GESTIONE_MDB\PRES3000_N00_GESTIONE.mdb"
    
    
    'LE TABELLE IN GENERALE
     '   DoCmd.CopyObject sxPath, "", acTable, "GEST_____________________GESTIONI_______________________________"
        
    
    'LE QUERY IN GENERALE
        'DoCmd.CopyObject sxPath, "", acQuery, "CUP_Tb01_{@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@}_Qry"
        'DoCmd.CopyObject sxPath, "", acQuery, "CUP_Tb01_Qry_{@========================================@}_Tb01"
        'DoCmd.CopyObject sxPath, "", acQuery, "CUP_Tb01_Qry01_SELECT_ARCHIVIO"
        
        
        'DoCmd.CopyObject sxPath, "", acQuery, "CIG_Tb01_{@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@}_Qry"
        'DoCmd.CopyObject sxPath, "", acQuery, "CIG_Tb01_Qry_{@========================================@}_Tb01"
        'DoCmd.CopyObject sxPath, "", acQuery, "CIG_Tb01_Qry01_SELECT_ARCHIVIO"

        
    
    'LE FORM IN GENERALE
        DoCmd.CopyObject sxPath, "", acForm, "PRES3000_Frm_{@===============================================@}"
        DoCmd.CopyObject sxPath, "", acForm, "PRES3000_FrmTb01_S01_}-----------------------------------------@"
        iCount = iCount + 1
        '//LA FORM GRUPPO PROGETTO + QUERY SOTTOSTANTE
        DoCmd.CopyObject sxPath, "", acForm, "PRES3000_FrmTb01_S01_Calendario"
            iCount = iCount + 1
            DoCmd.CopyObject sxPath, "", acQuery, "PRES3000_QryTb01_01_Calendario"
            iCount = iCount + 1
            
            
        DoCmd.CopyObject sxPath, "", acForm, "PRES3000_FrmTb02_S01_}-----------------------------------------@"
            iCount = iCount + 1
        DoCmd.CopyObject sxPath, "", acForm, "PRES3000_FrmTb02_S01_CalendarioPianoFerie"
            iCount = iCount + 1
            DoCmd.CopyObject sxPath, "", acQuery, "PRES3000_QryTb02_01_CalendarioPianoFerie"
            iCount = iCount + 1
        
        '//LA FORM GRUPPO PROGETTO + QUERY SOTTOSTANTE
        'DoCmd.CopyObject sxPath, "", acForm, "CUP_Tb01_Frm_S03_GRUPPO_PROGETTI"
        '    DoCmd.CopyObject sxPath, "", acQuery, "CUP_DF01_Qry10_03_GROUP_PROGETTI"
        
            
    'I MODULI IN GENERALE
        
      'DoCmd.CopyObject sxPath, "", acModule, "GestLTTMdl____________________Moduli_Gestione_LOTTO_____________"
           
          
           
    MsgBox "OGGETTI ESPORTATI IN GESTIONE PRESENZE PER UN TOTALE DI --> " & iCount, vbExclamation, " ESPORTAZIONE ESEGUITA "
           
 
    
'FINE - ESPORTA IN BIL.MDB
'____________________________________________________________________________________________________________________________________________
   
    

ESPORTA_OGGETTI_GESTIONE_Exit:
    Exit Function

ESPORTA_OGGETTI_GESTIONE_Err:
    MsgBox Error$
    Resume ESPORTA_OGGETTI_GESTIONE_Exit

End Function
