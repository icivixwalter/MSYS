'//************************************************************************************************************************************//
'//*'   CLASSE DELLA FORM : GEFILE_Frm01_GESTIONI.MD   @FORM@GESTIONE@FILE_(che utilizza le @API@WINDOWS)
'//*'                                                           *** INIZIO ***
'//*'
'//************************************************************************************************************************************//
'//*' NOTE: per il funzionamento di questa classe occorre il seguente modulo: _
        GEFILE_Mdl01_GESTIONI_FILE_API_VBA
'//*'
'//*' ATTIVITA ESEGUITE: _
      '//@01_@GESTIONI@COSTANTI@API                 = LE COSTANTI API PER IL FUNZIONAMENTO DELLA FORM DI GESTIONE DEI FILE
      '//@02_@GESTIONI@CMD_SELEZIONA FILE           = COMANDO PER LA SELEZIONE FILE DI WINDOWS
      '//@03_@GESTIONI@CMD_SELEZIONA FILE           = COMANDO PER LA SELEZIONE FILE DI WINDOWS
      '//@04_@GESTIONI@CONVERTI@FILTRO              = ROUTINE PER LA CONVERZIONE DEL FILTRO INTESTO
      '//@05_@GESTIONI@APRI@FILE                    = CHIAMO LA ROUTINE DEL MODULO ESTERNO PER L'APERTURA DEI FILE PDF
      '//@06_@GESTIONI@CLIK@FILTRO                  = IMPOSTO MANUALMENTE IL FILTRO DELLA FORM DI WINDOWOS
      '//@07_@GESTIONI@RECUPERO@FILE                = COMANDO PER IL RECUPERO DEL NOME DEL FILE
      '//@08_@GESTIONI@RECUPERO@PATH@FILE           = COMANDO PER IL RECUPERO DELLA PATH DEL FILE SELEZIONATO
      '//@09_@GESTIONI@RECUPERO@SHELL@EXE           = ROUTINE CHE APPLICA LA SCHELL EXE PER L'APERTURA DEL FILE **FUNZIONA?? CONTROLLARE ***




'// ARGOMENTO     = FileDialog della libreria ACCESS 2007      (@libreria@usata)
'// ATTENZIONE    = Microsoft Office 12.0 Object libary'
'//               Per il funzionamento occorre la @libreria 12
'//Fornisce alle finestre di dialogo dei file funzioni simili a quelle delle finestre di dialogo standard
'//Apri e Salva delle applicazioni di Microsoft Office.

'Metodi
'------------------------------------------------------------------------------------------------------------------------------------'
' Nome Descrizione
' Execute  Esegue un'azione utente non appena è stato richiamato il metodo Show.
'
' Show  Visualizza una finestra di dialogo dei file e restituisce un valore Long che indica se è stato premuto il pulsante Azione
'     (-1) o il pulsante  Annulla (0).
'     Quando si chiama il metodo Show non viene eseguito altro codice fino alla chiusura della finestra di dialogo.
'     Per le finestre di dialogo Apri e Salva con nome, utilizzare il metodo Execute subito dopo il metodo Show per eseguire l'azione dell'utente.
'------------------------------------------------------------------------------------------------------------------------------------'
'
'Proprietà
'Metodi
'------------------------------------------------------------------------------------------------------------------------------------'
' Nome Descrizione
' Execute  Esegue un'azione utente non appena è stato richiamato il metodo Show.
'
' Show  Visualizza una finestra di dialogo dei file e restituisce un valore Long che indica se è stato premuto il pulsante Azione
'     (-1) o il pulsante  Annulla (0).
'     Quando si chiama il metodo Show non viene eseguito altro codice fino alla chiusura della finestra di dialogo.
'     Per le finestre di dialogo Apri e Salva con nome, utilizzare il metodo Execute subito dopo il metodo Show per eseguire l'azione dell'utente.
'------------------------------------------------------------------------------------------------------------------------------------'
'
'Proprietà
'------------------------------------------------------------------------------------------------------------------------------------'
'
' Nome Descrizione
' AllowMultiSelect  Ha valore True se è consentito selezionare più file in una finestra di dialogo dei file. Proprietà di lettura/scrittura.
' Application       Ottiene un oggetto Application che rappresenta l'applicazione contenitore per l'oggetto FileDialog. È possibile utilizzare questa
'                   proprietà con un oggetto Automation per restituire l'applicazione contenitore dello stesso oggetto. Sola lettura.
' ButtonName        Imposta o ottiene un valore String che rappresenta il testo visualizzato sul pulsante di azione di una finestra di dialogo dei file.
'                   Proprietà di lettura/scrittura.
' Creator           Ottiene un numero intero a 32 bit che indica l'applicazione in cui è stato creato l'oggetto FileDialog. Sola lettura.
' DialogType        Ottiene una costante MsoFileDialogType che rappresenta il tipo di finestra di dialogo dei file visualizzata in base all'impostazione
'                   dell'oggetto FileDialog. Proprietà di sola lettura.
' FilterIndex       Ottiene o imposta un oggetto Long che indica il filtro di file predefinito per una finestra di dialogo dei file. Tale filtro determina
'                   quali tipi di file verranno visualizzati alla prima apertura della finestra di dialogo. Proprietà di lettura/scrittura.
' Filters           Ottiene un insieme FileDialogFilters. Proprietà di sola lettura.
' InitialFileName   Imposta o restituisce un oggetto String che rappresenta il percorso o il nome del file che viene visualizzato inizialmente in una
'                   finestra di dialogo dei file. Proprietà di lettura/scrittura.
' InitialView       Ottiene o imposta una costante MsoFileDialogView che rappresenta la presentazione iniziale dei file e delle cartelle in una finestra
'                   di dialogo dei file. Proprietà di lettura/scrittura.
' Item              Ottiene il testo associato all'oggetto specificato. Sola lettura.
' Parent            Ottiene l'oggetto Parent per l'oggetto FileDialog. Proprietà di sola lettura.
' SelectedItems     Ottiene un insieme FileDialogSelectedItems che contiene un elenco dei percorsi dei file selezionati da un utente in una finestra di
'                   dialogo dei file visualizzata utilizzando il metodo Show dell'oggetto FileDialog. Proprietà di sola lettura.
' Title             Imposta o ottiene il titolo di una finestra di dialogo di gestione dei file visualizzata utilizzando l'oggetto FileDialog. Lettura/scrittura.
'------------------------------------------------------------------------------------------------------------------------------------'







Option Compare Database
Option Explicit

'//@01_@GESTIONI@COSTANTI@API                = LE COSTANTI API PER IL FUNZIONAMENTO DELLA FORM DI GESTIONE DEI FILE
'//-------------------------------------------------------------------------------------------------------------------//

' Costanti per gestire il file
Private Const OFN_ALLOWMULTISELECT As Long = &H200
Private Const OFN_EXPLORER As Long = &H80000
Private Const OFN_PATHMUSTEXIST As Long = &H800
Private Const OFN_FILEMUSTEXIST As Long = &H1000

' Dichiarazione API per aprire il file con ShellExecute
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

' Tipo di dati per la finestra di dialogo
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type


' Dichiarazione API per aprire la finestra di selezione file
Private Declare Function GetOpenFileNameA Lib "comdlg32.dll" ( _
    ByRef ofn As OPENFILENAME) As Long

'//-------------------------------------------------------------------------------------------------------------------//


'//'//@02_@GESTIONI@CMD_SELEZIONA FILE          = COMANDO PER LA SELEZIONE FILE DI WINDOWS
'//-------------------------------------------------------------------------------------------------------------------//

' Funzione per aprire la finestra di selezione file
Public Function ApriFinestraSelezioneFile(Optional multiSelect As Boolean = False, _
                                          Optional filter As String = "Tutti i file (*.*)|*.*", _
                                          Optional titolo As String = "Seleziona un file") As Variant
    Dim ofn As OPENFILENAME
    Dim sFileName As String
    Dim result As Long

    ' Inizializza il buffer per il nome del file
    sFileName = String(260, 0)

    ' Configura la finestra di dialogo
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = Application.hWndAccessApp
    ofn.lpstrFilter = ConvertFilter(filter)  ' Assicurati che ConvertFilter sia definito correttamente
    ofn.lpstrFile = sFileName
    ofn.nMaxFile = Len(sFileName)
    ofn.lpstrTitle = titolo
    ofn.Flags = IIf(multiSelect, OFN_ALLOWMULTISELECT, OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST)

    ' Mostra la finestra di dialogo
    result = GetOpenFileNameA(ofn)  ' La funzione API dichiarata nel modulo esterno

    ' Restituisce il file selezionato o un array di file
    If result <> 0 Then
        Dim selectedFiles As String
        selectedFiles = Left(ofn.lpstrFile, InStr(ofn.lpstrFile, Chr(0)) - 1)
        If multiSelect Then
            ApriFinestraSelezioneFile = Split(selectedFiles, vbNullChar)
        Else
            ApriFinestraSelezioneFile = selectedFiles
        End If
    Else
        ApriFinestraSelezioneFile = Null
    End If
End Function
'//-------------------------------------------------------------------------------------------------------------------//

'//@04_@GESTIONI@CONVERTI@FILTRO              = ROUTINE PER LA CONVERZIONE DEL FILTRO INTESTO
' Funzione per convertire il filtro in un formato compatibile con l'API
Private Function ConvertFilter(filter As String) As String
    ConvertFilter = Replace(filter, "|", Chr(0)) & Chr(0)
End Function

'//@05_@GESTIONI@APRI@FILE                    = CHIAMO LA ROUTINE DEL MODULO ESTERNO PER L'APERTURA DEI FILE PDF
Private Sub Cmd_ApriFilePDF_Click()
  '//attivo la routine apri pdf del modulo @APRI_PDF_MODULO->GEFILE_Mdl01_GESTIONI_FILE_API_VBA
  Call APRI_FILE_PDF_pSub
End Sub

'//@06_@GESTIONI@CLIK@FILTRO                  = IMPOSTO MANUALMENTE IL FILTRO DELLA FORM DI WINDOWOS
' Implementazione di Cmd_FILTER_Click
Private Sub Cmd_FILTER_Click()
    On Error GoTo Err_Handler

    Dim result As Variant
    result = ApriFinestraSelezioneFile(False, "Immagini (*.gif;*.jpg;*.jpeg)|*.gif;*.jpg;*.jpeg", "Seleziona un'immagine")
    If Not IsNull(result) Then
        MsgBox "File selezionato: " & result
    Else
        MsgBox "Nessun file selezionato."
    End If

Exit_Handler:
    Exit Sub

Err_Handler:
    MsgBox Err.Description
    Resume Exit_Handler
End Sub

'//@07_@GESTIONI@RECUPERO@FILE                = COMANDO PER IL RECUPERO DEL NOME DEL FILE
' Implementazione di Cmd_RecuperaFILE_Click
Private Sub Cmd_RecuperaFILE_Click()
    On Error GoTo Err_Handler

    Dim result As Variant
    result = ApriFinestraSelezioneFile(True, "Tutti i file (*.*)|*.*", "Seleziona uno o più file")
    If Not IsNull(result) Then
        Dim i As Integer
        For i = LBound(result) To UBound(result)
            MsgBox "File selezionato: " & result(i)
        Next i
    Else
        MsgBox "Nessun file selezionato."
    End If

Exit_Handler:
    Exit Sub

Err_Handler:
    MsgBox Err.Description
    Resume Exit_Handler
End Sub

'//@08_@GESTIONI@RECUPERO@PATH@FILE           = COMANDO PER IL RECUPERO DELLA PATH DEL FILE SELEZIONATO
' Implementazione di Cmd_RecuperaPath_Click
Private Sub Cmd_RecuperaPath_Click()
    On Error GoTo Err_Handler

    Dim result As Variant
    result = ApriFinestraSelezioneFile(False, "Tutti i file (*.*)|*.*", "Seleziona un file")
    If Not IsNull(result) Then
        MsgBox "Percorso file selezionato: " & result
    Else
        MsgBox "Nessun file selezionato."
    End If

Exit_Handler:
    Exit Sub

Err_Handler:
    MsgBox Err.Description
    Resume Exit_Handler
End Sub

'//@09_@GESTIONI@RECUPERO@SHELL@EXE           = ROUTINE CHE APPLICA LA SCHELL EXE PER L'APERTURA DEL FILE **FUNZIONA?? CONTROLLARE ***
' Chiamata alla funzione ShellExecute per aprire un file selezionato
Private Sub ApriFileConShellExecute(filePath As String)
    If Dir(filePath) <> "" Then
        ShellExecute 0, "open", filePath, vbNullString, vbNullString, 1
        MsgBox "File aperto: " & filePath, vbInformation
    Else
        MsgBox "Errore: il file specificato non esiste.", vbCritical
    End If
End Sub

'//************************************************************************************************************************************//
'//*'   CLASSE DELLA FORM : GEFILE_Frm01_GESTIONI.MD                      *** FINE ***
'//*'
'//*'
'//************************************************************************************************************************************//

