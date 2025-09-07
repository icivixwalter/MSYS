Attribute VB_Name = "TREEVIEW_Mdl00_CARICA_LIBRERIE"
Option Compare Database
Option Explicit


'//TREEVIEW_Mdl00_CARICA_LIBRERIE = modulo per il caricamento delle librerie, viene attivato da una macro _
   quando il funzionamento del Treview non è possibile a casua della mancanza delle _
   librerie necessarie. La procedura predispone una MATRICE con 7 LIBRERIA  che devono _
   essere caricate nell'ordine in caso di loro mancanza. Viene fatto il confronto tra _
   quelle gia esistenti e quelle necessarie. Vengono scartate quelle esistente e caricate _
   solo le libreria mancante dopo un confronto dei file della matrice e del ciclo for each _
   avendo depurato la path.


'//LA FUNZIONE VIENE ATTIVATA CON LA MACRO denominata TREVIEW_Mcr01_ATTIVA_LIBRERIE CHE CHIAMA LA ROUTINE PUBBLICA

'//*******************************************************************************************//
'//                                 carica le librerie
'//*******************************************************************************************//

'//ROUTINE DI PROVA DELLA FUNZIONE
Private Sub AttivaReferences()
    Call AddReferences
End Sub


'//FUNZIONE CHE CONTROLLA E CARICA LE LIBRERIE MANCANTI
'//---------------------------------------------------------------------------------------//
Public Function AddReferences()
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
    ' AddReferences = True
    
    
'//uscita e controllo errori
Exit Function

ErrHandler:
    ' Handle errors
    MsgBox "Si è verificato un errore: " & Err.Description, vbCritical, "Errore"
    AddReferences = False
End Function

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


'//FUNZIONE CHE CONTROLLA E CARICA LE LIBRERIE MANCANTI
'//---------------------------------------------------------------------------------------//

