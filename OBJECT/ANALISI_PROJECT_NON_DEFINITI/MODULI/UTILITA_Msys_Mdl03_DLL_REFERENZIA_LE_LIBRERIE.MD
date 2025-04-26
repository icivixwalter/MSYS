'//*****************************************************************************************//
'//*    Modulo di gestione delle librerie DLL di Msys
'//*    UTILITA_Msys_Mdl03_DLL_REFERENZIA_LE_LIBRERIE
'//*    Note: cona funzione STAMPARE_Refereces()  stampa nella tabella Msys_DLL_LIBRERIE
'//*            le librerie attive nel progetto corrente.
'//*
'//*
'//*
'//*****************************************************************************************//


Option Compare Database

'//DIM DLL
    Dim ref As access.Reference
    Dim blnBroken As Boolean

'//gli oggetti delle referenze delle librerie
  Dim objRefEvents As New References

'//dim per il test
    
    Dim oVBP As AccessObject
    Dim dbs As Object


'//REFERENZIARE LA LIBRERIA DAO DA CODICE
Public Function IsDAO_Referenced() As Boolean
'==========================================================================
'Name : IsDAO_Referenced (Function)
'Purpose : Verifica se Sono Settati i Riferimenti alle Lib. DAO
'Author : Alessandro Baraldi
'Web_Site : http://digilander.iol.it/ik2zok/
'E.Mail : ***@libero.it
'Date : 01 febbraio 2002
'Called by :
'Calls :
'Inputs : Nothing
'Output : True only if is DAO Reference SET
'@controllo.librerie.attive
'==========================================================================

On Error GoTo 0

IsDAO_Referenced = False
Application.SetOption "Error Trapping", 2

For Each ref In Application.References
    blnBroken = ref.IsBroken
    Debug.Print "Libreria attiva: " & ref.Name; " path   ------> " & ref.fullPath
    
    
    If ref.Name = "DAO" And ref.Kind = 0 And StrComp(ref.Guid, cDAOGUID, 1) = 0 And Not blnBroken Then
        IsDAO_Referenced = True
    Exit For

    End If
Next ref

Set ref = Nothing

End Function


'//ITERA NELLA LIBRERIA DI REFERECES
Function IteraReferenze() As Boolean

    On Error GoTo Error_IteraReferenze
    ' Create new reference on References object variable.
    'ERRATO --> Set ref = objRefEvents.evtReferences.AddFromFile(strFileName)
    'objRefEvents.Item (1)
    
    
Exit_IteraReferenze:
    Exit Function

Error_IteraReferenze:
    MsgBox Err & ": " & Err.Description
    IteraReferenze = False
    Resume Exit_IteraReferenze
End Function





'NON FUNZIONA MA E' UN ESEMPIO
Sub Test()
'esegue un ciclo sui riferimenti presenti
'nel progetto attivo scrivendo nella finestra
'immediata il Nome del riferimento, e gli
'argomenti necessari per aggiungere un riferimento
'tramite il metodo AddFromGuid
'Dim oVBP As VBProject     @non.funziona


Dim oRef As Reference
    Set oVBP = Application.CurrentProject         'Application.VBE.ActiveVBProject
        For Each oRef In oVBP
            Debug.Print oRef.Name
            Debug.Print oRef.Guid
            Debug.Print oRef.Major
            Debug.Print oRef.Minor
        Next
End Sub



'//controllo librerie del progetto corrente, _
@stampo.libreria,  @stampo.referenze @le.librerie.stampo
Public Function STAMPARE_Refereces()
Dim ref As access.Reference
Dim blnBroken As Boolean
Dim iCount As Integer
Dim Str1 As String
Dim myNameDLL As String
Dim obj As AccessObject, dbs As Object

 
 
'//RESET
    '//svuoto la tabella DLL
    DoCmd.OpenQuery "Msys_DLL_Qry01_81_DELETE_LIBRERIE"
 
    Set dbs = Application
    'Set ref = dbs.References
        
        Debug.Print "'LIBRERIE REFERENZIATE NEL PROGETTO DELL'UFFICIO"
        Debug.Print "'--------------------------------------------------------"
        Debug.Print
        For Each ref In Application.References
            iCount = iCount + 1
            Debug.Print "'nro               -- > " & iCount
            Str1 = Str1 & Chr$(13)
            Str1 = Str1 & "nro              -- > " & iCount & Chr$(13)
            
            Debug.Print "'ref.Name          -- > " & ref.Name
            Str1 = Str1 & "ref.Name         -- > " & ref.Name & Chr$(13)
            
            Debug.Print "'ref.Guid          -- > " & ref.Guid
            Str1 = Str1 & "ref.Guid         -- > " & ref.Guid & Chr$(13)
            
            Debug.Print "'ref.FullPath      -- > " & ref.fullPath
            Str1 = Str1 & "ref.FullPath     -- > " & ref.fullPath & Chr$(13)
            Str1 = Str1 & Chr$(13)
             
            '//RECUPERO IL NOME DELLA LIBRERIA DLL DALLA PATH
            myNameDLL = ""
            myNameDLL = DLL_GetFileName(ref.fullPath)
             
            '//ACCODO IN TABELLA Msys_DLL_LIBRERIE
            SQL = ""
            sSql = sSql & "INSERT INTO Msys_DLL_LIBRERIE (LIBRERIA_DLL_s, COD_PROGETTO_s, PROGETTO_s, NRO_i, NAME_s, GUID_s, PATH_s )"
            sSql = sSql & "SELECT " & "'" & myNameDLL & "'" & " AS MyLIBRERIA_DLL_s," & "'MSys'" & "AS MyCOD_PROGETTO_s," & "'PROGETTO MSys'" & " AS MyPROGETTO_s," & iCount & " AS NRO_is,'" & ref.Name & "' AS NAME_s, '" & ref.Guid & "' AS GUID_s,'" & ref.fullPath & "' AS PATH_s "
            sSql = sSql & "WITH OWNERACCESS OPTION;"
             
             '//controllo esecuzione e reset del ssql
             '//.............................................................//
                Debug.Print sSql
                CurrentDb.Execute sSql
                sSql = ""
                
             '//.............................................................//
        Next ref
        
                '//apro LA QUERY delle LIBRERIE CORRENTI
                DoCmd.OpenQuery "Msys_DF13Qry01_01_DLL_PROGETTO"

                '//apro la QUERY DEI PROGETTI LIBRERIE
                DoCmd.OpenQuery "Msys_DF14Qry01_01_DLL_LIBRERIE"

        

End Function

' Function to extract the file name from a full path _
    FUNZIONE DI ESTRAZIONE DEL SOLO FILE DALLA PATH DI SISTEMA
Function DLL_GetFileName(fullPath As String) As String
    Dim pos As Integer
    '//TROVO LA POSIZIONE DELLA BARRA FINALE prende l'ultima a cui aggiunge 1
    pos = InStrRev(fullPath, "\")
    If pos > 0 Then
        '//estrae solo il nome del file con l'estenzione. _
            mid restituisce i caratteri dalla posizione pos, ossia la \
        DLL_GetFileName = Mid(fullPath, pos + 1)
    Else
        DLL_GetFileName = fullPath
    End If
End Function





'//*******************************************************************************************//
'//                                 CARICA_LE_LIBRERIE_PER IL_TREEVIEW *** INIZIO ***
'//*******************************************************************************************//

'//ROUTINE DI PROVA DELLA FUNZIONE
Private Sub AttivaReferences()
    Call AddReferences_TREEVIEW
End Sub


'//FUNZIONE CHE CONTROLLA E CARICA_LE_LIBRERIE_PER IL_TREEVIEW MANCANTI
'//---------------------------------------------------------------------------------------//
Public Function AddReferences_TREEVIEW()
    On Error GoTo ErrHandler

    Dim ref As Reference
    Dim refs As References
    Dim libPath As String
    Dim fileName As String
    Dim alreadyAdded As Boolean
    Dim addedLibs As String
    Dim existingLibs As String
    Dim addedCount As Integer
    Dim existingCount As Integer
    Dim i As Integer
    Dim libAdded As Boolean

    ' Set the References object
    Set refs = Application.References
    
    '//ELENCO LIBRERIE NECESARIE PER IL FUNZIONAMENTO DEL TREEVIEW NELL'ORDINE DI FUNZIONAMENTO.
    ' Array of file paths for the libraries
    Dim libraries(1 To 7) As String
    libraries(1) = "c:\Casa\LINGUAGGI\ACCESS\DLL_LIBRERIE\VBE6EXT.OLB"
    libraries(2) = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\DLL_LIBRERIE\MSACC.OLB"
    libraries(3) = "c:\Casa\LINGUAGGI\ACCESS\DLL_LIBRERIE\MSADO21.TLB"
    libraries(4) = "c:\Casa\LINGUAGGI\ACCESS\DLL_LIBRERIE\MSO.DLL"
    libraries(5) = "c:\Casa\LINGUAGGI\ACCESS\DLL_LIBRERIE\scrrun.dll"
    libraries(6) = "c:\Casa\LINGUAGGI\ACCESS\DLL_LIBRERIE\MSCOMCTL.OCX"
    libraries(7) = "c:\Casa\LINGUAGGI\ACCESS\DLL_LIBRERIE\dao360.dll"
    
    '//SOLO I FILE DELLE LIBRERIE
    ' Array of file names for the libraries (for comparison)
    Dim libraryFileNames(1 To 7) As String
    libraryFileNames(1) = "VBE6EXT.OLB"
    libraryFileNames(2) = "MSACC.OLB"
    libraryFileNames(3) = "MSADO21.TLB"
    libraryFileNames(4) = "MSO.DLL"
    libraryFileNames(5) = "scrrun.dll"
    libraryFileNames(6) = "MSCOMCTL.OCX"
    libraryFileNames(7) = "dao360.dll"

    addedLibs = ""
    existingLibs = ""
    addedCount = 0
    existingCount = 0
    
    
    '//ESEGUO IL CICLO FOR TRA IL SISTEMA E CONTROLLO LE LIBRERIE GIA CARICATE
    ' Loop through the array and add each reference
    For i = 1 To 7
        libPath = libraries(i)
        fileName = libraryFileNames(i)
        alreadyAdded = False
        libAdded = False
        
        '//COFRONTA LA LIBRERIA DI SISTEMA E LA LIBRERIA DELLA MATRICE SENZA PATH
        ' Check if the reference is already added by comparing file names
        For Each ref In refs
            If GetFileName(ref.fullPath) = fileName Then
                alreadyAdded = True
                Exit For
            End If
        Next ref
        
        
        '//SE LA LIBRERIA NON E' PRECARICATA L'AGGIUNGE AL SISTEMA AGGIORNANDO IL CONTATTO DI QUELLE INSERITE
        ' If the reference is not already added, try to add it
        If Not alreadyAdded Then
            On Error Resume Next
            Application.References.AddFromFile libPath
            If Err.number = 0 Then
                libAdded = True
                '//CONTATORI DELLE LIBRERIE AGGIUNTE
                addedLibs = addedLibs & fileName & vbCrLf
                addedCount = addedCount + 1
            Else
                Debug.Print "Errore nell'aggiungere la libreria: " & libPath & " - " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        End If
        
        '//CONTATORI DELLE LIBRERIE ESISTENTI
        ' Update counts and lists based on the result
        If alreadyAdded Then
            existingLibs = existingLibs & fileName & vbCrLf
            existingCount = existingCount + 1
        End If
    Next i

    ' Show summary - CONTROLLO FINALE CON LA VISUALIZZAZIONE DI QUELLE AGGIUNTE E DI QUELLE ESISTENTI
    Dim summary As String
    summary = "Librerie aggiunte per il TreeView (" & addedCount & "):" & vbCrLf & addedLibs & vbCrLf & vbCrLf & _
              "Librerie già esistenti per il funzionamento del TreeView (" & existingCount & "):" & vbCrLf & existingLibs
    MsgBox summary, vbInformation, "Riepilogo delle librerie"
    
    ' Return True indicating success TODO: per ora bloccata
    ' AddReferences_TREEVIEW = True
    
    
'//uscita e controllo errori
Exit Function

ErrHandler:
    ' Handle errors
    MsgBox "Si è verificato un errore: " & Err.Description, vbCritical, "Errore"
    AddReferences_TREEVIEW = False
End Function


'//TODO: RECUPERARE IL NOME DELLA LIBRERIA  e salvarlo nel campo LIBRERIA_DLL_s???
' Function to extract the file name from a full path _
    FUNZIONE DI ESTRAZIONE DEL SOLO FILE DALLA PATH DI SISTEMA
Function GetFileName(fullPath As String) As String
    Dim pos As Integer
    '//TROVO LA POSIZIONE DELLA BARRA FINALE prende l'ultima a cui aggiunge 1
    pos = InStrRev(fullPath, "\")
    If pos > 0 Then
        '//estrae solo il nome del file con l'estenzione. _
            mid restituisce i caratteri dalla posizione pos, ossia la \
        GetFileName = Mid(fullPath, pos + 1)
    Else
        GetFileName = fullPath
    End If
End Function


'//FUNZIONE CHE CONTROLLA E CARICA_LE_LIBRERIE_PER IL_TREEVIEW MANCANTI
'//---------------------------------------------------------------------------------------//

'//*******************************************************************************************//
'//                                 CARICA_LE_LIBRERIE_PER IL_TREEVIEW *** FINE ***
'//*******************************************************************************************//


