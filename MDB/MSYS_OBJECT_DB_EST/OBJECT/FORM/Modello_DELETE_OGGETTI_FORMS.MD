## Modello_DELETE_OGGETTI_FORMS.MD







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
    Dim icount As Integer
    Dim i As Integer
    Dim Bool1 As Boolean
    
    On Error GoTo Err_Cmd_Delete_FORM_Click

    '//RESET
    icount = 0
    deletetFormsCount_i = 0
    deletedForms_s = ""

    '//ATTIVO LA COLLEZIONE
    '//------------------------------------------------------------------------//
    '//Note: CHIAMO LA FUNZIONE COLLECTION FORM per ottenere la collezione di FORM

    Set formsCollection = CollectionForm_PFunct()
    '//------------------------------------------------------------------------//

    ' Loop attraverso tutte le FORM nel database corrente
    Set dbCurrent = CurrentDb

    '//CALCOLO INDICE FORM
    icount = Application.CurrentProject.AllForms.Count

    '//se ci sono FORM da cancellare
    If icount > 0 Then

        '//reimposto la matrice sulle FORM esistenti nel db corrente
        ReDim MATRICE_FORM_Names_s(icount - 1)

        ' Costruisci un array con i nomi dei FORM
        tempCount = 0

        '//CARICO LA MATRICE: ciclo for per caricare MATRICE_FORM_Names_s(tempCount)
        For i = 0 To icount - 1
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

                If Err.Number <> 0 Then
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
