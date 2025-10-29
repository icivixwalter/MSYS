'//MODULO = Msys_GEFILE_Mdl02_ESPORTA_OGGETTI

'//01_IMPOSTAZIONI MODULO
   Option Compare Database
   Option Explicit

'//02_ESPORTO TUTTI GLI OGGETTI DI GEFILE      ***INIZIO ***
'//============================================================================================================================//
       
        '//  ATTIVITA ESEGUITE'
        '//- gestisce più percorsi (sxPath1, sxPath2, ecc.)
        '//- usa Collection numeriche per i tipi oggetto (1, 2, 3, …)
        '//- verifica esistenza oggetto prima dell’esportazione
        '//- scrive tutto su un file di log (oggetti OK, errori, mancanti)
        '//- salta in sicurezza se una Collection è vuota
        '//- mostra un messaggio riepilogativo finale
        '//----------------------------------------------------------------------------'
        '//   FILTRO NUMERICO APPLICATO PER I DIVERSI OGGETTI '
        '//   Mappatura ufficiale dei tipi Access
        '//   Tipo numerico  Costante VBA   Oggetto Access
        '//   1  acQuery  Query
        '//   2  acForm   Form
        '//   3  acReport Report
        '//   4  acMacro  Macro
        '//   5  acModule Modulo
        '//   6  acTable  Tabella
        '// Qui il primo valore "2" è numerico (per acForm), il secondo è il nome dell’oggetto.
        '// Quando il codice gira, viene fatto:
        '// tipo = Split(itm, "|")(0)
        '// DoCmd.CopyObject sxPath, "", tipo, nome
        '// E quindi tipo è già un numero (1, 2, 3, …), quindi perfettamente compatibile con DoCmd.CopyObject.

    
   
   Public Function Msys_GEFILE_Mdl01_ESPORTA_OGGETTI_Pfnc()
   
       On Error GoTo GestErr
   
       Dim colOggetti As New Collection
       Dim colPercorsi As New Collection
       Dim itm As Variant, sxPath As Variant
       Dim fso As Object, logFile As Object
       Dim totaleOk As Long, totaleErr As Long
       Dim logPath As String
       Dim tipo As Integer, nome As String
       Dim startTime As Date
   
       startTime = Now()
       logPath = "C:\GESTIONI\LOG_ESPORTAZIONE.TXT"
       Set fso = CreateObject("Scripting.FileSystemObject")
       Set logFile = fso.CreateTextFile(logPath, True)
   
       logFile.WriteLine String(60, "=")
       logFile.WriteLine "LOG ESPORTAZIONE OGGETTI ACCESS"
       logFile.WriteLine "Database origine: " & CurrentDb.Name
       logFile.WriteLine "Data/Ora inizio: " & startTime
       logFile.WriteLine String(60, "=")
       logFile.WriteLine ""
   
       '==========================================================
       '=== PATH DI DESTINAZIONE ================================
       '==========================================================
       colPercorsi.Add "C:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\LLPP_ATTI_GESTIONE.mdb"  'sxPath1
       colPercorsi.Add "C:\GESTIONI\BACKUP\LLPP_ATTI_GESTIONE_BACKUP.mdb"                                   'sxPath2
   
       '==========================================================
       '=== OGGETTI DA ESPORTARE ================================
       '==========================================================
   
       'FORM MASTER
       colOggetti.Add "2|PRES3000_Frm01_}-----------------------------------------@MASTER"
       colOggetti.Add "2|PRES3000_Frm01_{@=============================================@}"
       colOggetti.Add "2|PRES3000_Frm01_M01_SOTTO_FORM_A_SCHEDE"
   
       'RAGGRUPPAMENTI
       colOggetti.Add "2|PRES3000_Frm04_S02_01_RAGGR_ANNO"
       colOggetti.Add "1|PRES3000_Qry01_04_01_RAGGRUPPA_ANNO"
       colOggetti.Add "2|PRES3000_Frm04_S02_RAGGR_MESE"
       colOggetti.Add "1|PRES3000_Qry01_04_01_RAGGRUPPA_ANNO"
   
       'GIUSTIFICAZIONI
       colOggetti.Add "2|PRES3000_Frm01_S04_03_RAGGRUPPA_ANNO_GIUSTIFICAZIONI"
       colOggetti.Add "1|PRES3000_Qry01_04_03_RAGGRUPPA_ANNO_GIUSTIFICAZIONI"
       colOggetti.Add "2|PRES3000_Frm01_S04_03_RAGGRUPPA_GIUSTIFICAZIONI_MESE"
       colOggetti.Add "1|PRES3000_Qry01_04_04_RAGGRUPPA_GIUSTIFICAZIONI_MESE"
       colOggetti.Add "2|PRES3000_Frm01_S04_05_RAGGRUPPA_GIUSTIFICAZIONI_X_MESE"
       colOggetti.Add "1|PRES3000_Qry01_04_05_RAGGRUPPA_GIUSTIFICAZIONI_X_MESE"
   
       'CALENDARI
       colOggetti.Add "2|PRES3000_FrmDF01_S01_}-----------------------------------@CALEND"
       colOggetti.Add "2|PRES3000_FrmTb01_S01_Calendario"
       colOggetti.Add "1|PRES3000_QryTb01_01_Calendario"
   
       colOggetti.Add "2|PRES3000_FrmTb02_S01_CalendarioPianoFerie"
       colOggetti.Add "1|PRES3000_QryTb02_01_CalendarioPianoFerie"
       colOggetti.Add "1|PRES3000_QryDF01_01_CalendarioFestivita"
   
       colOggetti.Add "2|PRES3000_FrmDF01_S01_CalendarioFestivita"
       colOggetti.Add "1|PRES3000_QryDF01_01_CalendarioFestivita"
       colOggetti.Add "1|PRES3000_QryDF01_02_GROUP_CalendarioFestivita"
       colOggetti.Add "1|PRES3000_QryDF01_02_UPDATE_CalendarioFestivita"
   
       'MACRO
       colOggetti.Add "4|CERCA_Mcr_}----------------------------------------------------@"
       colOggetti.Add "4|CERCA_Mcr_PRES3000_DF01_CalendarioFestivita"
       colOggetti.Add "4|COLLEGA_Mcr_PRES3000_DF01_CalendarioFestivita"
       colOggetti.Add "4|COLLEGA_Mcr_PRES3000_Tb02_CalendarioPianoFerie"
       colOggetti.Add "4|Msys_GEFILE_Mcr_01_COLLEGA_TUTTO_PER_LA_GESTIONE"
   
       '==========================================================
       '=== VERIFICHE PRELIMINARI ================================
       '==========================================================
   
       If colPercorsi.Count = 0 Then
           MsgBox "Nessun percorso di destinazione definito. Esportazione annullata.", vbExclamation
           logFile.WriteLine "?? Nessun percorso di destinazione definito. Esportazione annullata."
           GoTo Esci
       End If
   
       If colOggetti.Count = 0 Then
           MsgBox "Nessun oggetto da esportare. Esportazione annullata.", vbExclamation
           logFile.WriteLine "?? Nessun oggetto da esportare. Collection vuota."
           GoTo Esci
       End If
   
       '==========================================================
       '=== CICLO DI ESPORTAZIONE ================================
       '==========================================================
   
       For Each sxPath In colPercorsi
           logFile.WriteLine vbCrLf & String(50, "-")
           logFile.WriteLine "DESTINAZIONE: " & sxPath
           logFile.WriteLine String(50, "-")
   
           If Not fso.FileExists(sxPath) Then
               logFile.WriteLine "? ATTENZIONE: file non trovato -> " & sxPath
               Debug.Print "ATTENZIONE: file non trovato -> " & sxPath
               GoTo NextPath
           End If
   
           For Each itm In colOggetti
               tipo = CLng(Split(itm, "|")(0))
               nome = Split(itm, "|")(1)
   
               If OggettoEsiste(tipo, nome) Then
                   On Error Resume Next
                   DoCmd.CopyObject sxPath, "", tipo, nome
                   If Err.Number = 0 Then
                       totaleOk = totaleOk + 1
                       Debug.Print "OK -> " & nome
                       logFile.WriteLine "OK  -> " & nome
                   Else
                       totaleErr = totaleErr + 1
                       Debug.Print "ERRORE -> " & nome & " (" & Err.Description & ")"
                       logFile.WriteLine "ERRORE -> " & nome & " (" & Err.Description & ")"
                   End If
                   Err.Clear
                   On Error GoTo GestErr
               Else
                   totaleErr = totaleErr + 1
                   Debug.Print "NON TROVATO -> " & nome
                   logFile.WriteLine "NON TROVATO -> " & nome
               End If
           Next itm
   NextPath:
       Next sxPath
   
   Esci:
       logFile.WriteLine vbCrLf & "Totale OK: " & totaleOk & " - Errori: " & totaleErr
       logFile.WriteLine "Data/Ora fine: " & Now()
       logFile.WriteLine String(60, "=")
       logFile.Close
   
       MsgBox "? Esportazione completata!" & vbCrLf & _
              "Oggetti OK: " & totaleOk & vbCrLf & _
              "Errori o non trovati: " & totaleErr & vbCrLf & _
              "Log dettagliato: " & logPath, vbInformation, "RISULTATO ESPORTAZIONE"
   
       Exit Function
   
   GestErr:
       MsgBox "Errore durante l’esportazione: " & Err.Description, vbCritical, "ERRORE"
       If Not logFile Is Nothing Then logFile.Close
   
   End Function
   
   
   '------------------------------------------------------------
   ' FUNZIONE DI CONTROLLO ESISTENZA OGGETTO
   '------------------------------------------------------------
   Private Function OggettoEsiste(tipoOggetto As Integer, nomeOggetto As String) As Boolean
       On Error Resume Next
       Select Case tipoOggetto
           Case 1: OggettoEsiste = Not IsNull(CurrentData.AllQueries(nomeOggetto).Name)
           Case 2: OggettoEsiste = Not IsNull(CurrentProject.AllForms(nomeOggetto).Name)
           Case 3: OggettoEsiste = Not IsNull(CurrentProject.AllReports(nomeOggetto).Name)
           Case 4: OggettoEsiste = Not IsNull(CurrentProject.AllMacros(nomeOggetto).Name)
           Case 5: OggettoEsiste = Not IsNull(CurrentProject.AllModules(nomeOggetto).Name)
           Case 6: OggettoEsiste = Not IsNull(CurrentData.AllTables(nomeOggetto).Name)
           Case Else: OggettoEsiste = False
       End Select
   End Function
   
   
   '//ESPORTO TUTTI GLI OGGETTI DI GEFILE      ***FINE ***
   '//============================================================================================================================//
   
   
