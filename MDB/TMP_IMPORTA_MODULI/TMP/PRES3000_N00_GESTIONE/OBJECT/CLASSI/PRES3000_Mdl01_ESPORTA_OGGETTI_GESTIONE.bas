Attribute VB_Name = "PRES3000_Mdl01_ESPORTA_OGGETTI_GESTIONE"

Option Compare Database

'########################################################################################
'@MODULO@PRES3000_N00_GESTIONE_Mdl01_ESPORTA_OGGETTI
'
'LE FUNZIONI ESPORTA OGGETTI
'@progetto@PRES3000_N00_GESTIONE.mdb
'
'@database@DB_PRES3000_N00_GESTIONE
'NOTA: Le funzioni pubbliche esportano nel db di destinazione i seguenti oggetti:
'Tabelle, Query Maschere, Report, Pagine di proprieta, Macro e Moduli
'
' VIENE ATTIVATA DALLA MACRO ---> PRES3000_MCR01_GESTIONE_ESPORTA_OGGETTI_TO_LLPP_ATTI_GESTIONE
'
'########################################################################################



'------------------------------------------------------------
' PRES3000_N00_GESTIONE_Mdl01_ESPORTA_OGGETTI_Pfnc
'
'Nota: Esporto le qry nei db.
'------------------------------------------------------------
Public Function PRES3000_N00_GESTIONE_Mdl01_ESPORTA_OGGETTI_Pfnc()

'

On Error GoTo PRES3000_N00_GESTIONE_Mdl01_ESPORTA_OGGETTI_Pfnc_Err

Dim sxPath As String
Dim sxNAME As String
Dim sxFILE As String

'ESPORTA IN FILE.MDB
'____________________________________________________________________________________________________________________________________________

        'MODELLI
        'DoCmd.CopyObject sxPATH, "", acQuery, ""
        'DoCmd.CopyObject sxPATH, "", acTable, ""
        'DoCmd.CopyObject sxPATH, "", acForm, ""
        'DoCmd.CopyObject sxPATH, "", acModule, ""
   
    
    'reset
    sxPath = "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\LLPP_ATTI_GESTIONE.mdb"
    
    
    'LE TABELLE IN GENERALE
     '   DoCmd.CopyObject sxPath, "", acTable, "GEST_____________________GESTIONI_______________________________"
        
    
    'LE QUERY IN GENERALE
        'DoCmd.CopyObject sxPath, "", acQuery, "CUP_Tb01_{@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@}_Qry"
        'DoCmd.CopyObject sxPath, "", acQuery, "CUP_Tb01_Qry_{@========================================@}_Tb01"
        'DoCmd.CopyObject sxPath, "", acQuery, "CUP_Tb01_Qry01_SELECT_ARCHIVIO"
        
        
        'DoCmd.CopyObject sxPath, "", acQuery, "CIG_Tb01_{@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@}_Qry"
        'DoCmd.CopyObject sxPath, "", acQuery, "CIG_Tb01_Qry_{@========================================@}_Tb01"
        'DoCmd.CopyObject sxPath, "", acQuery, "CIG_Tb01_Qry01_SELECT_ARCHIVIO"

        
    
        '//LE FORM IN GENERALE - FORM A SCHEDE @LA@FORM@MASTER
        '//===============================================================================//
            DoCmd.CopyObject sxPath, "", acForm, "PRES3000_Frm01_}-----------------------------------------@MASTER"
            DoCmd.CopyObject sxPath, "", acForm, "PRES3000_Frm01_{@=============================================@}"
            DoCmd.CopyObject sxPath, "", acForm, "PRES3000_Frm01_M01_SOTTO_FORM_A_SCHEDE"
        
        
        '//===============================================================================//
        
        '//@FORM@RAGGRUPPAMENTI @FORM@ANNO @FORM@MESE
        '//===============================================================================//
            '//FORM
            DoCmd.CopyObject sxPath, "", acForm, "PRES3000_Frm04_S02_01_RAGGR_ANNO"
             '//QUERY
            DoCmd.CopyObject sxPath, "", acQuery, "PRES3000_Qry01_04_01_RAGGRUPPA_ANNO"
            '//FORM
            DoCmd.CopyObject sxPath, "", acForm, "PRES3000_Frm04_S02_RAGGR_MESE"
             '//QUERY
            DoCmd.CopyObject sxPath, "", acQuery, "PRES3000_Qry01_04_01_RAGGRUPPA_ANNO"
            
        
        '//===============================================================================//
        
        
        
        '//LE FORM GRUPPO PROGETTO + QUERY SOTTOSTANTI
         '//===============================================================================//
            '//FORM
            DoCmd.CopyObject sxPath, "", acForm, "PRES3000_Frm01_S04_03_RAGGRUPPA_ANNO_GIUSTIFICAZIONI"
            '//QUERY
            DoCmd.CopyObject sxPath, "", acQuery, "PRES3000_Qry01_04_03_RAGGRUPPA_ANNO_GIUSTIFICAZIONI"
            
            
            '//FORM
            DoCmd.CopyObject sxPath, "", acForm, "PRES3000_Frm01_S04_03_RAGGRUPPA_GIUSTIFICAZIONI_MESE"
            '//QUERY
            DoCmd.CopyObject sxPath, "", acQuery, "PRES3000_Qry01_04_04_RAGGRUPPA_GIUSTIFICAZIONI_MESE"
            
            
        
            '//FORM
            DoCmd.CopyObject sxPath, "", acForm, "PRES3000_Frm01_S04_05_RAGGRUPPA_GIUSTIFICAZIONI_X_MESE"
            '//QUERY
            DoCmd.CopyObject sxPath, "", acQuery, "PRES3000_Qry01_04_05_RAGGRUPPA_GIUSTIFICAZIONI_X_MESE"
            
            
            
            
        '//===============================================================================//
        
        
          '//PER IL CALENDARIO FERIE E CALENDARIO FORM+QUERY+MACRO
          '//===============================================================================//
                
            '//FORM CALENDARIO
            '//...................................
            
                DoCmd.CopyObject sxPath, "", acForm, "PRES3000_FrmDF01_S01_}-----------------------------------@CALEND"
                DoCmd.CopyObject sxPath, "", acForm, "PRES3000_FrmTb01_S01_Calendario"
                    '//QUERY
                    
                    DoCmd.CopyObject sxPath, "", acQuery, "PRES3000_QryTb01_01_Calendario"
                        
            '//...................................
         
                    
                    
          
          '//FORM CalendarioPianoFerie
          '//...................................
        DoCmd.CopyObject sxPath, "", acForm, "PRES3000_FrmTb02_S01_CalendarioPianoFerie"
            '//QUERY
            DoCmd.CopyObject sxPath, "", acQuery, "PRES3000_QryTb02_01_CalendarioPianoFerie"
            DoCmd.CopyObject sxPath, "", acQuery, "PRES3000_QryDF01_01_CalendarioFestivita"
            
            
            '//...................................
            
            '//FORM CALENDARIO FESTIVITA
            '//...................................
            DoCmd.CopyObject sxPath, "", acForm, "PRES3000_FrmDF01_S01_CalendarioFestivita"
                 '//QUERY
                DoCmd.CopyObject sxPath, "", acQuery, "PRES3000_QryDF01_01_CalendarioFestivita"
                DoCmd.CopyObject sxPath, "", acQuery, "PRES3000_QryDF01_02_GROUP_CalendarioFestivita"
                DoCmd.CopyObject sxPath, "", acQuery, "PRES3000_QryDF01_02_UPDATE_CalendarioFestivita"
               
                            
            '//...................................
                
    
             '//MACRO
            DoCmd.CopyObject sxPath, "", acMacro, "CERCA_Mcr_}----------------------------------------------------@"
            DoCmd.CopyObject sxPath, "", acMacro, "CERCA_Mcr_PRES3000_DF01_CalendarioFestivita"
            DoCmd.CopyObject sxPath, "", acMacro, "CERCA_Mcr_PRES3000_DF10_GIORNALIERE"
            DoCmd.CopyObject sxPath, "", acMacro, "CERCA_Mcr_PRES3000_DF11_TIMBRATURE"
            DoCmd.CopyObject sxPath, "", acMacro, "CERCA_Mcr_PRES3000_DF12_ORARI_TIMBRATURE"
            DoCmd.CopyObject sxPath, "", acMacro, "CERCA_Mcr_PRES3000_DF20_Help_Calendario"
            DoCmd.CopyObject sxPath, "", acMacro, "CERCA_Mcr_PRES3000_Tb01_Calendario"
            DoCmd.CopyObject sxPath, "", acMacro, "CERCA_TABELLA_Mcr_PRES3000_Tb02_CalendarioPianoFerie"
            DoCmd.CopyObject sxPath, "", acMacro, "CERCA_TABELLA_Mcr_PRES3000_Tb02_ElencoGiornaliere"
            DoCmd.CopyObject sxPath, "", acMacro, "COLLEGA_Mcr_}--------------------------------------------------@"
            DoCmd.CopyObject sxPath, "", acMacro, "COLLEGA_Mcr_PRES3000_DF00_}------------------------------------@"
            DoCmd.CopyObject sxPath, "", acMacro, "COLLEGA_Mcr_PRES3000_DF01_CalendarioFestivita"
            DoCmd.CopyObject sxPath, "", acMacro, "COLLEGA_Mcr_PRES3000_DF05_ORARI"
            DoCmd.CopyObject sxPath, "", acMacro, "COLLEGA_Mcr_PRES3000_DF10_GIORNALIERE"
            DoCmd.CopyObject sxPath, "", acMacro, "COLLEGA_Mcr_PRES3000_DF11_TIMBRATURE"
            DoCmd.CopyObject sxPath, "", acMacro, "COLLEGA_Mcr_PRES3000_DF12_ORARI_TIMBRATURE"
            DoCmd.CopyObject sxPath, "", acMacro, "COLLEGA_Mcr_PRES3000_DF20_Help_Calendario"
            DoCmd.CopyObject sxPath, "", acMacro, "COLLEGA_Mcr_PRES3000_TB01_Calendario"
            DoCmd.CopyObject sxPath, "", acMacro, "COLLEGA_Mcr_PRES3000_Tb02_CalendarioPianoFerie"
            DoCmd.CopyObject sxPath, "", acMacro, "COLLEGA_Mcr_PRES3000_Tb02_ElencoGiornaliere"
            DoCmd.CopyObject sxPath, "", acMacro, "PRES3000_N00_GESTIONE_Mcr_}------------------------------------@"
            DoCmd.CopyObject sxPath, "", acMacro, "PRES3000_N00_GESTIONE_Mcr_01_COLLEGA_TUTTO_PER_LA_GESTIONE"
            DoCmd.CopyObject sxPath, "", acMacro, "PRES3000_N00_GESTIONE_Mcr_02_COLLEGA_Parziali_Solo_CALENDARI"
            DoCmd.CopyObject sxPath, "", acMacro, "PRES3000_N00_GESTIONE_Mcr_03_COLLEGA_Parziali_Solo_TIMBRATURE"
 
            
            
            
          '//===============================================================================//
    
        
            
    'I MODULI IN GENERALE
        
      'DoCmd.CopyObject sxPath, "", acModule, "GestLTTMdl____________________Moduli_Gestione_LOTTO_____________"
           
          
           
    MsgBox "ESPORTAZIONE ESEGUITA", vbExclamation, "OGGETTI ESPORTATI GESTIONE"
           
 
    
'FINE - ESPORTA IN BIL.MDB
'____________________________________________________________________________________________________________________________________________
   
    

PRES3000_N00_GESTIONE_Mdl01_ESPORTA_OGGETTI_Pfnc_Exit:
    Exit Function

PRES3000_N00_GESTIONE_Mdl01_ESPORTA_OGGETTI_Pfnc_Err:
    MsgBox Error$
    Resume PRES3000_N00_GESTIONE_Mdl01_ESPORTA_OGGETTI_Pfnc_Exit

End Function
