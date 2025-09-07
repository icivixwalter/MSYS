VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_TreeView_FrmS01_GE_TREEVIEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Risoluzione dei problemi di riferimento nel computer di sviluppo

'La creazione di un nuovo database vuoto e l'importazione di oggetti
'da un altro file di database possono creare problemi di riferimento se il codice o i controlli ActiveX si basano su riferimenti che non sono inclusi in un database per impostazione predefinita. I riferimenti predefiniti per un database di Access 2000 sono i seguenti:
'•Visual Basic, Applications Edition
'•Libreria oggetti di Microsoft Access 9,0
'•Automazione OLE
'•Raccolta Microsoft ActiveX Data Objects (ADO) 2,1

'La creazione di un nuovo database vuoto e l'importazione di oggetti da un altro file di database possono creare problemi di riferimento
'se il codice o i controlli ActiveX si basano su riferimenti che non sono inclusi in un database per impostazione predefinita.
'I riferimenti predefiniti per un database di Access 2000 sono i seguenti:

'•Visual Basic, Applications Edition
'•Libreria oggetti di Microsoft Access 9,0
'•Automazione OLE
'•Raccolta Microsoft ActiveX Data Objects (ADO) 2,1  = msado21.tlb


'// LE LIBRERIE DELL' UFFICIO SONO NELL'ORDINE: _
    Vba Visual basic for application                        = Vbe6ext.olb   ? FORSE ESTENSIONE
    'microsoft Access 12.0 Object Library                   = MSACC.OLB
    
    'Visual Basic For Applications                          = Vbe6intl.dll
    'microsoft Access 12.0 Object Library                   = MSACC.OLB
    'Microsoft Common Controls 6.0                          = Mscomctl.ocx

'LIBRERIE REFERENZIATE NEL PROGETTO DELL'UFFICIO
'--------------------------------------------------------


'// @LIBRERIA@VBA_(LIBRERIA VISUAL BASIC FOR APPLICATION))_@LIBRERIA@TREVIEW
'//------------------------------------------------------------------------------------------//
'NomeLibreria   -- > Visual Basic For Application _
                    la libreria puo essere _
                    VBE6.DLL oppure _
                    Vbe6ext.olb
'nro            -- > 1
'ref.Name       -- > VBA
'ref.Guid       -- > {000204EF-0000-0000-C000-000000000046}
'ref.FullPath   -- > C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6.DLL _
    in alternativa IN CASO DI ERRORI prelevarlo nella cartella librerie: _
                    c:\CASA\LINGUAGGI\ACCESS\DLL_LIBRERIE\VBE6.DLL

'//------------------------------------------------------------------------------------------//


'// @LIBRERIA@ACCESS@9.0_(LIBRERIA MICROSOFT ACCESS 9,0 OPPURE ACCESS 12.00)_@LIBRERIA@TREVIEW
'//------------------------------------------------------------------------------------------//
'NomeLibreria   -- > Microsoft Access 9,0 Object Library                   = MSACC9.OLB _
                     Microsoft Access 12.0 Object Library                  = MSACC9.OLB ma si trova in una cartella diversa _
                    la libreria è la MSACC.OLB

'nro            -- > 2
'ref.Name       -- > Access
'ref.Guid       -- > {4AFFC9A0-5F99-101B-AF4E-00AA003F0F07}

'                    se access  9.00 = Microsoft Access 12.0 Object Library  = MSACC.OLB
'                    c:\Program Files (x86)\Microsoft Office\Office\
'
'                    se access 12.00 = Microsoft Access 12.0 Object Library  = MSACC.OLB
'                    c:\CASA\LINGUAGGI\ACCESS\ACCESS_PROGETTI_MDB\ACCESS_TREVIEW\LIBRERIE\Microsoft Office\Office12\MSACC.OLB
'
'ref.FullPath   -- > C:\Program Files (x86)\Microsoft Office\Office12\MSACC.OLB _
 in alternativa IN CASO DI ERRORI prelevarlo nella cartella librerie: _
                     c:\CASA\LINGUAGGI\ACCESS\ACCESS_PROGETTI_MDB\ACCESS_TREVIEW\LIBRERIE\Microsoft Office\Office12\MSACC.OLB

'//------------------------------------------------------------------------------------------//



'// @LIBRERIA@DAO_(LIBRERIA PER RECORDSET DAO)__@LIBRERIA@TREVIEW@DA0360
'//------------------------------------------------------------------------------------------//
'nro            -- > 3
'ref.Name       -- > DAO
'ref.Guid       -- > {00025E01-0000-0000-C000-000000000046}
'ref.FullPath   -- > C:\Program Files (x86)\Common Files\Microsoft Shared\DAO\dao360.dll _
oppure se non funziona IN CASO DI ERRORI  prelevarlo nella cartella librerie: _
                -- > c:\CASA\LINGUAGGI\ACCESS\ACCESS_PROGETTI_MDB\ACCESS_TREVIEW\LIBRERIE\Common Files\Microsoft Shared\DAO\dao360.dll
'//------------------------------------------------------------------------------------------//


'// @LIBRERIA@VB6@CONTROL_(LIBRERIA PER I CONTROLLI DI VSISUAL BASIC 6.0 )_@MSCOM_@LIBRERIA@TREVIEW
'//------------------------------------------------------------------------------------------//
'//NOte: la libreria per i controlli; Microsofof Windows Common Controls 6.0 (SP6)
'nro            -- > 4
'ref.Name       -- > MSComctlLib
'ref.Guid       -- > {831FDD16-0C5C-11D2-A9FC-0000F8754DA1}
'ref.FullPath   -- > C:\CASA\LINGUAGGI\ACCESS\ACCESS_PROGETTI_MDB\ACCESS_TREVIEW\LIBRERIE\MSCOMCTL.OCX

'nro            -- > 5
'ref.Name       -- > ADODB
'ref.Guid       -- > {00000201-0000-0010-8000-00AA006D2EA4}
'ref.FullPath   -- > c:\Program Files (x86)\microsoft office\Office\MSO9.DLL _
    in alternativa IN CASO DI ERRORI prelevarlo nella cartella librerie: _
                    c:\CASA\LINGUAGGI\ACCESS\DLL_LIBRERIE\MSO9.DLL

'//------------------------------------------------------------------------------------------//

'@ORDINE@DELLE@LIBRERIE_@LIBRERIA@TREVIEW
'//------------------------------------------------------------------------------------------//
    'Visual Basic For Application
    'Microsoft Access 9,0 Object Library
    'Microsoft Activex Data Object Recordset 2.8 Object Library
    'Microsoft DAO 3.6 Object Library
    'Microsoft Activex Data Object 6.1 Library
    'Microsoft Windows Common Controls 6.0 (SP6)
    'Microsoft Office 9.0 Object Library
'//------------------------------------------------------------------------------------------//


Option Compare Database
Option Explicit

'//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++//
'//       LE VARIABILI DI MODULO

'//LIBRERIE : MSCOMCTL.OCX  SI TROVA IN -> c:\windows\SysWow64\
'//c:\windows\SysWow64\MSCOMCTL.OCX
'//La libreria di access è questa : Microsoft Windows Common Controls 6.0

'//LE VARIABILI DATABASE
'//....................................................................//
    Dim daoDB As DAO.Database
    Dim DaoWks As DAO.Workspace
    Dim daoRS As DAO.Recordset

    Dim ADODB As Database
    Dim AdodaoRs As Recordset
    Dim sSql As String                          '//STRINGA SQL
    Dim PATH_s As String                        '//la path
    Dim MyPath_s  As String


    '//Contatori
    Dim iCount As Integer
    Dim dbl_count As Double
    
   
    'Le variabili generiche
    Dim Vv1 As Variant
    Dim Dbl1 As Double
    Dim Int1 As Integer
    Dim Long1 As Long
    Dim Lng  As Long
    Dim Str1 As String
    
    '//Messaggi di errore
    Dim ProceduraMessaggioErrore_s As String    '//Errore procedura
    Dim ProceduraAttivaEseguita_s As String     '//Errore Attivita eseguita


'....................................................................

'//DISCO ED IL FILE ATTO
'//....................................................................

    Dim Disco_s As String
    Dim PathFileAtto_s As String
    Dim MyFileEstensione_s As String
    Dim FileAtto_s As String
    Dim WshPDF As Variant

'//....................................................................


    
    '//CLASSE CLASSE_ALBERO
    Dim my_CLASSE_ALBERO As CLASSE_ALBERO                             'VARIABILE DELLA CLASSE  che rappresenta l'intero CLASSE_ALBERO



'//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++//



'//********************************************************************************************************//
'//                             EVENTI FORM
'//01_FORM_LOAD = Carico la form aggiornando la tabella di stampa _
                  e chiamo il caricamento della TreeView
'
'//02_LIST VIEW EVENTO CLICK SCELTA SELEZIONE _
                = Evento click del Nodo selezionato dove viene individuato _
                quale Nodo e sottonodo o foglia che viene cliccato _
                CODICE------>TV_NodeClick.Attiva


'
'//********************************************************************************************************//


'//01_FORM_LOAD = Carico la form aggiornando la tabella di stampa
Private Sub Form_Load()

      
    
    '//pulisco la tabella di stampa
    'DoCmd.OpenQuery "Tree_Qry01_CLASSI_STAMPA_DI_CONTROLLO_CANCELLA"
    
    '//refresh della sottoform 02 la quale contiene TEST_TUTTE_LE_CLASSI_StampaRisultati.Requery
    'Me.SottoForm_02.Requery
    
    '//Carico la TreeView
    'Call CARICA_TREE_pSub
    
    
    '//CARICA CLASSE_ALBERO TREE = con la routine pubblica
    '//...........................................................................................//
    '//CODICE----> CARICA_TREE_pSub.Routine.ATTIVA
        '// CARICA CLASSE_ALBERO TREE --- >
        'Call CARICA_TREE_pSub
    '//...........................................................................................//
    
    '//IMPOSTO IL VALORE DEFINITO DELLA COMBINATA 01 -
    '//@01_modifica_cmb_01
        Me.Cmb_01.Value = ">>"
   
    
End Sub


'//********************************************************************************************************//



'//02_CARICO_TREEVIEW = Con il comando carico la TreView
Private Sub Cmd_01_Click()
On Error GoTo Err_Cmd_01_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    
    DoCmd.OpenQuery "Tree_Qry01_CLASSI_STAMPA_DI_CONTROLLO_CANCELLA"
    
    
    stDocName = "Report_STAMPA_TUTTO"
    'DoCmd.OpenReport stDocName, acViewPreview
       
       
       
    
    Call testCLASSE_ALBERO
    
    
    stDocName = "TEST_TUTTE_LE_CLASSI_StampaRisultati"
    'DoCmd.OpenForm stDocName, , , stLinkCriteria
    '//aggiorno la sottoform 02 che contine i risultati di stampa : ...TEST_TUTTE_LE_CLASSI_StampaRisultati.Requery
    ' sospeso ??? Me.SottoForm_02.Requery
       
       
Exit_Cmd_01_Click:
    Exit Sub

Err_Cmd_01_Click:
    MsgBox Err.Description
    Resume Exit_Cmd_01_Click
    
End Sub




'//                     SALVATAGGIO FIGLI IN TMP
'//**************************************************************************************************************************//
'//EX COMANDO CARICA_TREE_pSub TRASFORMATO IN ROUTINE PUBBICA CARICA TREE  OLD ----> Private Sub CARICA_TREE_pSub() _
                               Questa routine evento dava problemi, perchè è associata all'evento click. E' stata quindi _
                               trasformata in una ROUTINE PUBBLICA perchè in questo modo è attivabile in tutto il progetto _
                               senza dare gli errori di doppio caricamento nell'CLASSE_ALBERO Tree dell'oggetto TV (TreeView). _
                               Conclusione NON USARE EVENTI CLICK al di fuori della classe form perchè sono IMPREVEDIBILI.
'//nuova routine CARICA_TREE_pSub
'//CODICE----> CARICA_TREE_pSub.Routine
'//PROCEDIMENTO: _
    01) CARICA IL TREE, con la seguente routine     = CARICA_TREE_pSub() _
    02) IMPOSTA I NODI PADRIE E FIGLI = routine con passaggio collezione lista nordi ed il recordset nodi _
    03) CHIAMO IL TREE  passando la classe albero   = CaricaTreeView(ByVal Tree As CLASSE_ALBERO) _


'//01) CARICA IL TREE, con la seguente routine     = CARICA_TREE_pSub()
Public Sub CARICA_TREE_pSub()

On Error GoTo Err_CARICA_TREE_pSub
    
    Dim idCorrente As Integer
    Dim sottolista As Collection
    Dim lista_nodi_in_tabella As New Collection
    Dim idPadre As Integer
    Dim rsTutti As DAO.Recordset
    Dim Tree As CLASSE_ALBERO
    Dim CLASSE_NODO_in_lista As CLASSE_NODO
    
    '//CARICA CLASSE_ALBERO TREE = con la routine pubblica
    '//...........................................................................................//
    '//CODICE----> CARICA_TREE_pSub.Routine.ATTIVA
        '// CARICA CLASSE_ALBERO TREE --- > CALL CARICA_TREE_pSub()
    '//...........................................................................................//
    
    '//svuota la tabella tmp
    CurrentDb.Execute "DELETE TreeTemporanea.* FROM TreeTemporanea;"
    
    
    '//Pulisco il TreeView ma nel Pagina 01 del Tab control deve essere qualificato _
       con Me!TV.Nodes.Clear
       Me!TV.Nodes.Clear ' svuota il controllo treeview
    
    '//I° FASE CREO TUTTI I NODI CON IL CICLO WHILE SULLA TABELLA Tree
    '//01_CARICA_TREE.02 = Apri Recordset Tree I° FASE CREO TUTTI I NODI
    '//DATI DELLA TABELLA INSERITI NELLA CLASSE - (CARICAMENTO DATI)
    Set rsTutti = CurrentDb.OpenRecordset("SELECT * FROM Tree", , dbReadOnly)
    '//STEP 01 = Creo i nodi della tabella (ad es 8) e gli metto il nome
    Do While Not rsTutti.EOF:
        DoEvents
        
        '//01_CARICA_TREE.03 = imposta il nodo con il campo ("Id") nella CLASSE_NODO_in_lista
        Set CLASSE_NODO_in_lista = New CLASSE_NODO
        CLASSE_NODO_in_lista.ID = rsTutti.Fields("Id").Value
        '//01_CARICA_TREE.04 = imposta il nodo con il campo ("NomeCartella_s") nella CLASSE_NODO_in_lista
        CLASSE_NODO_in_lista.NOME = rsTutti.Fields("NomeCartella_s").Value
        '....... in questo modo abbiamo fatto lo step 1 e 2 del riempi CLASSE_ALBERO _
        senza usare la lista delle liste che potrebbero farti sbagliare perchè devi _
        ricordarti l'ordine dei dati. Invece con la classe chiami le proprieta. _

        '//01_CARICA_TREE.05 = Aggiungo alla lista_nodi_in_tabella CLASSE_NODO_in_lista.
        lista_nodi_in_tabella.Add CLASSE_NODO_in_lista
        
        '//CARICO I NODI
        '//AGGIORNAMENTO 06.09.2020 = ho aggiunto altri 3 valori nella classe (inserimenti WALTER) _
                Menu_s                  = CARICO IL MENU _
                ProgettoMENU_s          = CARICO IL PROGETTO _
                CodMENU_s               = CARICO IL CODICE MENU _

        '//.......................................................................
            '//I MENU NELLA CLASSE NODO IN LISTA :  vari campi
            If rsTutti.Fields("Menu_s").Value > "" Then
                CLASSE_NODO_in_lista.MenuGestito_s = rsTutti.Fields("Menu_s").Value
            Else
                CLASSE_NODO_in_lista.MenuGestito_s = ""
            End If
             
            If rsTutti.Fields("ProgettoMENU_s").Value > "" Then
                CLASSE_NODO_in_lista.ProgettoMENU_s = rsTutti.Fields("ProgettoMENU_s").Value
            Else
                CLASSE_NODO_in_lista.ProgettoMENU_s = ""
            End If
            
            
            If rsTutti.Fields("CodMENU_s").Value > "" Then
                CLASSE_NODO_in_lista.CodMENU_s = rsTutti.Fields("CodMENU_s").Value
            Else
                CLASSE_NODO_in_lista.CodMENU_s = ""
            End If

            '// CARICO IL DISCO
            If rsTutti.Fields("Disco_s").Value > "" Then
                CLASSE_NODO_in_lista.Disco_s = rsTutti.Fields("Disco_s").Value
            Else
                CLASSE_NODO_in_lista.Disco_s = ""
            End If

            
            '// CARICO LA PATH
            If rsTutti.Fields("PathFileAtto_s").Value > "" Then
                CLASSE_NODO_in_lista.PathFileAtto_s = rsTutti.Fields("PathFileAtto_s").Value
            Else
                CLASSE_NODO_in_lista.PathFileAtto_s = ""
            End If
            
            
            '// CARICO IL FILE ATTO
            If rsTutti.Fields("FileAtto_s").Value > "" Then
                CLASSE_NODO_in_lista.FileAtto_s = rsTutti.Fields("FileAtto_s").Value
            Else
                CLASSE_NODO_in_lista.FileAtto_s = ""
            End If
            
            
        '//.......................................................................
        
        rsTutti.MoveNext
    Loop
    '//II° FASE VENGONO CREATI I PADRI DEI NODI
    'dopo il while tutti i nodi sono stati creati e ora posso impostare i padri senza _
    rischiare che un certo CLASSE_NODO con id padre maggiore del proprio id non sia ancora creato
    rsTutti.MoveFirst
    
    '//STEP 02 e 03 imposto padri e figli
    imposta_padri_e_figli lista_nodi_in_tabella, rsTutti
    
    rsTutti.Close
    
    'Creo e stampo l'CLASSE_ALBERO a partire dalla tabella
    Set Tree = riempiCLASSE_ALBERO(lista_nodi_in_tabella)
    
    Tree.stampaCLASSE_ALBERO
    CaricaTreeView Tree
      
    '//AGGIUNTO = per riaggiornare la form
    Me.Form.Recalc
      
Exit_CARICA_TREE_pSub:
    Exit Sub

Err_CARICA_TREE_pSub:
    MsgBox Err.Description
    Resume Exit_CARICA_TREE_pSub

End Sub


'//02) IMPOSTA I NODI PADRI E FIGLI = routine con passaggio collezione lista nordi ed il recordset nodi
Public Sub imposta_padri_e_figli(ByVal lista_nodi As Collection, rs_nodi As DAO.Recordset)

    ' Imposta i padri e i figli dei nodi di una lista a partire da un recordset di una query.
    ' Non restituisce nulla, ma la lista_nodi data in input viene modificata (vedi STEP 02 e 03)
    Dim my_CLASSE_NODO As CLASSE_NODO
    Dim CLASSE_NODO_padre As CLASSE_NODO
    Dim id_padre As Variant
    
    For Each my_CLASSE_NODO In lista_nodi
        ' prendo l'id del padre dal recordset
        id_padre = rs_nodi.Fields("ParentID")
    
        If Not IsNull(id_padre) Then
            Stampa_ed_Inserisci "cerco il padre del CLASSE_NODO " & my_CLASSE_NODO.ID
            
            For Each CLASSE_NODO_padre In lista_nodi
                If CLASSE_NODO_padre.ID = Int(id_padre) Then
                    ' STEP 02 = Collego dal Figlio al Padre
                    my_CLASSE_NODO.father = CLASSE_NODO_padre ' Ho trovato il padre di my_CLASSE_NODO
                    ' STEP 03 = Collego dal Padre al figlio
                    CLASSE_NODO_padre.addSon my_CLASSE_NODO
                    Exit For
                End If
                    '//sospeso MA PUOI UITILIZZARLO PER IL CONTROLLO
                    'Stampa_ed_Inserisci "il padre di " & my_CLASSE_NODO.ID & " non è " & CLASSE_NODO_padre.ID
            Next
                '//AL POSTO DELL'ID della classe ho messo il nome _
                old --> Stampa_ed_Inserisci "Il padre di " & my_CLASSE_NODO.ID & " è " & CLASSE_NODO_padre.ID
            Stampa_ed_Inserisci "Il padre di " & my_CLASSE_NODO.NOME & " è " & CLASSE_NODO_padre.NOME
        End If
        rs_nodi.MoveNext
     Next
End Sub




'//03) CHIAMO IL TREE  passando la classe albero   = CaricaTreeView(ByVal Tree As CLASSE_ALBERO)
Private Sub CaricaTreeView(ByVal Tree As CLASSE_ALBERO)
    On Error GoTo Err_CaricaTreeView
    
    Dim tempNode As MSComctlLib.Node
    Dim lista_DFS As Collection
    Dim CLASSE_NODO_Corrente As CLASSE_NODO
    Dim idCLASSE_NODO As Variant
        
    Me!TV.Nodes.Clear ' svuota il controllo treeview
    Set lista_DFS = Tree.DFS(True)
    
    For Each idCLASSE_NODO In lista_DFS
        Set CLASSE_NODO_Corrente = Tree.find_node_by_id(idCLASSE_NODO, True)
        If CLASSE_NODO_Corrente.IsRadice Then
            Set tempNode = Me!TV.Nodes.Add(, , "C" & CLASSE_NODO_Corrente.ID, CLASSE_NODO_Corrente.NOME) 'Aggiungo la radice al TreevIEW
        Else
            Set tempNode = Me!TV.Nodes.Add("C" & CLASSE_NODO_Corrente.father.ID, tvwChild, "C" & CLASSE_NODO_Corrente.ID, CLASSE_NODO_Corrente.NOME)
        End If
    Next
    
    '//variabile pubblica CLASSE_ALBERO
    Set my_CLASSE_ALBERO = Tree
    
        
Exit_CaricaTreeView:
    Exit Sub

Err_CaricaTreeView:
    MsgBox Err.Description
    Resume Exit_CaricaTreeView
End Sub

'//**************************************************************************************************************************//




'//CLASSE NODO
'//***********************************************************************************************************//

Public Function CLASSE_NODOIdDaChiave(ByVal par_Node As MSComctlLib.Node) As Variant

     CLASSE_NODOIdDaChiave = Right(par_Node.Key, Len(par_Node.Key) - 1)
    
End Function

'//***********************************************************************************************************//








'//                             I COMANDI 1) ISERISCI NUOVA CARTELLA + FUNZIONE CONTROLLO DUPLICATO
'//***********************************************************************************************************//
'//NOTE : Le operazioni di inserimento sono : _
01_IMPOSTO_ID+NEW_CARTELLA = imposto il nuovo id ed il nome della nuova cartella da inserire _
02_FORM_INSERIMENTO_PARAMETRI = .... da creare .... _
03_INSERISCO_TMP_NELLA_TREEVIEW = faccio l'inserimento temporaneo nella Treeview _
04_CONTROLLO_DEL_DUPLICATO_NELLA_TABELLA = Prima dell'inserimento faccio il controllo se esiste il duplicato. _
05_INSERISCO_NELLA_TABELLA = inserisco l'ID e il nome della CARTELLA nella tabella Tree.


'//TODO : ancora da fare _
    a) come evitare l'errore del doppio inserimento nella collection del duplicato _
            ... TV.Nodes.Add  ....qui !!! _
    b) creare una form per l'inserimento dei 2 parametri in cui creare una nuova _
       cartella .... ----> ..... CREARE UNA FORM QUI = con i parametri ???? _
       02_FORM_INSERIMENTO_PARAMETRI = .... da creare .... _
    c)

Private Sub Cmd_NwCartella_Click()
    Dim NewIdParent_i  As Integer
    Dim NewNomeCartella_s As String
    Dim tempNode As MSComctlLib.Node
    
    
    
    
    '//ATTIVITA DI INSERIMENTO
    '//.......................................................................................................//
    '//imposto i valori ID + NOME CARTELLA del nuovo inserimento e prima _
       di inserire con la query INSERT controllo se esiste la nuova cartella _
       nella tabella restituendo il valore NON ESISTE oppure GIA ESISTE.
        
        '//01_IMPOSTO_ID+NEW_CARTELLA = imposto il nuovo id ed il nome della nuova cartella da inserire
            NewIdParent_i = 11
            NewNomeCartella_s = "nuova cartella"
                        
                        '// ..... CREARE UNA FORM QUI = con i parametri ????
        
        '//03_INSERISCO_TMP_NELLA_TREEVIEW = faccio l'inserimento temporaneo nella Treeview
        '//NOTA: l'inserimento nella Treeview è solo virtuale ma non è inserito nella tabella _
           per cui se non viene eseguito anche l'inserimento nella cartella questo inserimento _
           del CLASSE_NODO sarà perso. ??? ATTENZIONE IL DOPPIO INSERIMENTO CREA UN ERRORE di doppione _
           nella collezione inoltre bisogna recuperare il CLASSE_NODO selzionato????
            Set tempNode = Me!TV.Nodes.Add("C0", tvwChild, "CICI", "nuova cartella")
            
        '//04_CONTROLLO_DEL_DUPLICATO_NELLA_TABELLA = Prima dell'inserimento faccio il controllo se esiste il duplicato.
        '//NOTA : chiamo la funzione passando come parametro l'ID+NOME CARTELLA e controllo se esiste un duplicato _
                restituuendo una stringa filtrata dalla if che rilascia un messaggio
        Str1 = TrovaDuplicati(NewIdParent_i, NewNomeCartella_s)
        
            If Str1 = "IL VALORE GIA ESISTE" Then
                    MsgBox "IL VALORE DELLA NUOVA CARTELLA idParent -> " & NewIdParent_i & " NOME -> " & NewNomeCartella_s & _
                    " GIA ESISTE NELLA TABELLA, non inserire duplicati ?", vbExclamation
            End If
            
        '//05_INSERISCO_NELLA_TABELLA = inserisco l'ID e il nome della CARTELLA nella tabella Tree.
        '//NOTE: dopo l'esame del duplicato eseguo la query insert per l'inserimento.
            'Eseguo l'inserimento solo se non esiste il duplicato
            sSql = ""
            sSql = sSql & "INSERT INTO Tree ( ParentID, NomeCartella_s )"
            sSql = sSql & "SELECT " & NewIdParent_i & " AS NwParentID, '" & NewNomeCartella_s & "' AS NwNomeCartella;"
            Debug.Print sSql
            CurrentDb.Execute sSql

    '//*** FINE ***
    '//ATTIVITA DI INSERIMENTO
    '//.......................................................................................................//
    
End Sub



'//                     CONTROLLO TROVA DUPLICATI
'//**************************************************************************************************************************//
'//NOTE : controllo se nella tabella già esiste il duplicato restituendo la stringa _
         IL VALORE NON ESISTE oppure IL VALORE GIA ESISTE
Private Function TrovaDuplicati(NewIdParent_i As Integer, NewNomeCartella_s) As String

On Error GoTo Err_TrovaDuplicati
    
    Dim rsTutti As DAO.Recordset
    
    '//imposto le variabili e la imposta non esiste la cartella oggetto _
    di ricerca.
    TrovaDuplicati = "IL VALORE NON ESISTE"
   
    Set rsTutti = CurrentDb.OpenRecordset("SELECT * FROM Tree", , dbReadOnly)
    '//STEP 01 = controllo se esistono duplicati ed esco dal ciclo con il messaggio _
        IL VALORE ESISTE  restituendo la stringa al chiamante.

    Do While Not rsTutti.EOF:
        DoEvents
            If NewIdParent_i = rsTutti.Fields("ParentID") And _
               NewNomeCartella_s = rsTutti.Fields("NomeCartella_s") Then
                        TrovaDuplicati = "IL VALORE GIA' ESISTE"
                        '//vado alla fine del rs per chiudere ed _
                        uscire dal ciclo
                        rsTutti.MoveLast
                        TrovaDuplicati = "IL VALORE GIA ESISTE"
            End If
         
        rsTutti.MoveNext
    Loop
    
   
    rsTutti.Close
    
   
Exit_TrovaDuplicati:
    Exit Function

Err_TrovaDuplicati:
    MsgBox Err.Description
    Resume Exit_TrovaDuplicati

End Function


'//*** fine ***
'//                     CONTROLLO TROVA DUPLICATI
'//**************************************************************************************************************************//






'//DOPPIO CLICK SULL'ETICHETTA
'//CODICE---->LABEL_TREE_DblClick_CaricaTree
Private Sub LABEL_TREE_DblClick(Cancel As Integer)
    '//01_CARICA_TREE
    '//DONE = MODIFICA 24.09.2020 carica il tree al doppio click dell'etichetta
    Me.CARICA_TREE_pSub
End Sub




'//02_LIST VIEW EVENTO CLICK SCELTA SELEZIONE _
                = Evento click del Nodo selezionato dove viene individuato _
                quale Nodo e sottonodo o foglia che viene cliccato _
                CODICE------>TV_NodeClick.Attiva
'//***********************************************************************************************************//
'//
'//EVENTO CLICK DEL TREEVIEW
Private Sub TV_NodeClick(ByVal par_Node As Object)
    
    
    
    Dim CLASSE_NODOCliccato_cls As CLASSE_NODO
    Dim idCLASSE_NODO As Integer
    Dim CLASSE_NODOTreeView As MSComctlLib.Node
    
    '//VARIABILI DEL NODO
    Dim CLASSE_NODOCliccato_cls_s As String                     '//il PADRE DEL NODO
    Dim CLASSE_NODOTreeView_s As String                         '//Nodo del Treeview
    Dim CLASSE_NODOTreeViewKey_s As String                      '//LA CHIAVE
    
    
        Dim my_PROGETTI_GESTIONE_TUTTI_Classe As PROGETTI_GESTIONE_TUTTI_Classe
        '//@inizializzo la classe
        Set my_PROGETTI_GESTIONE_TUTTI_Classe = New PROGETTI_GESTIONE_TUTTI_Classe
        
    
    Set CLASSE_NODOTreeView = par_Node
            
            '//IMPOSTO LE VARIABILI del Nodo e della chiave del TreeView
            CLASSE_NODOTreeView_s = CLASSE_NODOTreeView.text
            CLASSE_NODOTreeViewKey_s = CLASSE_NODOTreeView.Key
    
                TV.Requery
  
      
    'DEVO RECUPERARE IL PADRE DEL CLASSE_NODO es. Padre = La quercia ---> CLASSE_NODO :2019_2020 _
    '...............................................................................
    
        idCLASSE_NODO = Right(par_Node.Key, Len(par_Node.Key) - 1)
        
        '//classe.albero.attivata,   @nodo.cliccato
        Set CLASSE_NODOCliccato_cls = my_CLASSE_ALBERO.find_node_by_id(idCLASSE_NODO, False)
        
        If Not CLASSE_NODOCliccato_cls.IsRadice Then
            '//salvo il nome della classe nodo se esiste
            CLASSE_NODOCliccato_cls_s = CLASSE_NODOCliccato_cls.father.NOME
            
            
       
        Else
             '//se non esiste il padre del nodo
            CLASSE_NODOCliccato_cls_s = "NON ESISTE IL PADRE DEL NODO"
        
        End If
       '...............................................................................
            
            
            '//CONTROLLO NODI CLICCATI - CLICK TREE - TREE SCELTA
            Debug.Print
            Debug.Print "=========================NODO CLICCATO =================================================="
    
            Debug.Print "PADRE DELLA CLASSE_NODO -------> " & CLASSE_NODOCliccato_cls_s
            Debug.Print "Classe ------------------------> " & CLASSE_NODOTreeView_s         'CLASSE_NODOTreeView.text
            Debug.Print "Chiave ------------------------> " & CLASSE_NODOTreeViewKey_s      'CLASSE_NODOTreeView.Key
            'old ---> Debug.Print "Hai cliccato il CLASSE_NODO: " & CLASSE_NODOTreeView.text & " con chiave: " & CLASSE_NODOTreeView.Key
            Debug.Print "Hai cliccato il CLASSE_NODO: " & CLASSE_NODOTreeView_s & " con chiave: " & CLASSE_NODOTreeViewKey_s
            
           
           
                       '//RECUPERO 3 CAMPI SCELTA MENU SOLO SE ESISTE I CODICE MENU
                       '//--------------------------------------------------------------------------------//-----------------------//
                       'a) colonna 0  = I MENU                                   -> MenuGestito_s _
                        b) colonna 1  = IL CODICE DEL MENU                       -> CodMENU_s _
                        c) colonna 3  = IL PROGETTO  a cui assegnato il menu     -> ProgettoMENU_s _
                        d) colonna    = LA PATH DEL FILE DA APIRE                -> PathFileAtto_s _
                        e) colonna    = IL TIPO DEL FILE                         -> TipoFileAtto_s _
                        f) colonna    = il disco                                 -> Disco_s
                        
                        
                        
                            '//NODO CLICCATO - scelta nodo - clik nodo
                            If CLASSE_NODOCliccato_cls.CodMENU_s > "" Then
                                
                                Debug.Print
                                Debug.Print
                                Debug.Print "Recupero i campi dei menu perche esiste il CODICE MENU"
                                Debug.Print "........................................................................."
                                    'CONTROLLO MENU GESTITO
                                    Debug.Print "IL CAMPO MENU MenuGestito_s -> " & CLASSE_NODOCliccato_cls.MenuGestito_s
                                    
                                    
                                    'CONTROLLO CODICE MENU
                                    Debug.Print "IL CAMPO CODICE MENU ProgettoMENU_s -> " & CLASSE_NODOCliccato_cls.CodMENU_s
                                    
                                    'CONTROLLO PROGETTO MENU
                                    Debug.Print "IL CAMPO PROGETTO MENU ProgettoMENU_s -> " & CLASSE_NODOCliccato_cls.ProgettoMENU_s
                                    
                                Debug.Print "........................................................................."
                                Debug.Print "                       I FILE DA APRIRE "
                                
                                
                                'CONTROLLO PROGETTO MENU
                                Debug.Print "IL CAMPO Disco_s                        -> " & CLASSE_NODOCliccato_cls.Disco_s
                                
                                'CONTROLLO PATH
                                Debug.Print "IL CAMPO PathFileAtto_s                 -> " & CLASSE_NODOCliccato_cls.PathFileAtto_s
                                
                                'CONTROLLO FileAtto_s
                                Debug.Print "IL CAMPO FileAtto_s                     -> " & CLASSE_NODOCliccato_cls.FileAtto_s
                                
                                            
                                
                                 Debug.Print "........................................................................."
                                
                                
                            
                                    
                                '//ATTIVO LA CLASSE DEI MENU PER LA GESTIONE DEI LAVORI PUBBLICI
                                '//===========================================================================================================//
                                '//Note : Creo le variabili oggetto della classe di gestione dei lavori pubblici e passo come parametri _
                                           il MENU GESTITO + IL CODICE MENU
                                        
                                        '//Imposto le variabili di classe + i parametri -->
                                        Dim my_LLPP_CLASSE_MENU_LAVORI_PUBBLICI  As LLPP_CLASSE_MENU_LAVORI_PUBBLICI
                                        Dim MenuGestito_s As String: Dim CodiceMenu_s As String: Dim Progetto_MENU_s As String
                                        
                                        '//Istanzio la classe -->
                                        Set my_PROGETTI_GESTIONE_TUTTI_Classe = New PROGETTI_GESTIONE_TUTTI_Classe
                                        
                                        
                                        '@menu.gestito.classe
                                        '//I MENU Reset + impostazioni delle variabili dei menu
                                        MenuGestito_s = "": CodiceMenu_s = "": Progetto_MENU_s = ""
                                        MenuGestito_s = CLASSE_NODOCliccato_cls.MenuGestito_s
                                        CodiceMenu_s = CLASSE_NODOCliccato_cls.CodMENU_s
                                        Progetto_MENU_s = CLASSE_NODOCliccato_cls.ProgettoMENU_s
                                        
                                        
                                         '//I FILE = reset + impostazioni delle variabili della path e del file
                                        Disco_s = "": PathFileAtto_s = "": FileAtto_s = ""
                                        Disco_s = CLASSE_NODOCliccato_cls.Disco_s
                                        FileAtto_s = CLASSE_NODOCliccato_cls.FileAtto_s
                                        PathFileAtto_s = CLASSE_NODOCliccato_cls.PathFileAtto_s
                                  
                                        '//ATTIVO LA CLASSE MENU LAVORI PUBBLICI
                                        '//___________________________________________________________________________________
                                        '//note: chiamo il menu dei lavori pubblici al quale ho eliminato LA PRIMA PRIMA PAGINA _
                                                 IL TREE  e la seconda pagina perche caricava 2 volte il Tre ed andava i crash.

                                                        '  = modificare la classe PROGETTI_GESTIONE_TUTTI_Classe nella _
                                                        IMPOSTA_GESTIONI MenuGestito_s, Progetto_CONDOMINIO_s dove deve recepire anche _
                                                        i tre campi : _
                                                        a) colonna 0 = il menu _
                                                        b) colonna 1 = IL CODICE DEL MENU _
                                                        c) colonna 3 = IL PROGETTO  a cui assegnato il menu _
                                                        d) colonna    = disco _
                                                        e) colonna    = path _
                                                        f) colonna    = file atto
                                                        
                                                          '//la Classe GESTIONE TUTTI, passo il MENU gestito + il codice Menu + _
                                                             il progetto MENU + i 3 campi dei file, il DISCO + LA PATH + IL FILE ATTO _
                                                             e la classe provvederà ad attivare un'altra classe _
                                                             specifica individuata con il valore PROGETTO CONDOMINIO al quale passera _
                                                             gli altri due parametri
                                                          '
                                                          '
                                                          my_PROGETTI_GESTIONE_TUTTI_Classe.IMPOSTA_GESTIONI MenuGestito_s, CodiceMenu_s, Progetto_MENU_s, _
                                                                                                             Disco_s, PathFileAtto_s, FileAtto_s

                                            
                                            
                                        '//*** FINE ***
                                        '//ATTIVO LA CLASSE MENU LAVORI PUBBLICI
                                        '//___________________________________________________________________________________
                                        
                                        '//LIBERO LA MEMORIA -->
                                        Set my_LLPP_CLASSE_MENU_LAVORI_PUBBLICI = Nothing
                                        MenuGestito_s = "":   Progetto_MENU_s = "": CodiceMenu_s = "": Disco_s = "": PathFileAtto_s = "": FileAtto_s = ""
                                        
                                  
                                        
                                      
                                        
                                '//*** FINE ***
                                '//ATTIVO LA CLASSE DEI MENU PER LA GESTIONE DEI LAVORI PUBBLICI
                                '//===========================================================================================================//
                                    
                            
                            
                            End If
                            
                        '//--------------------------------------------------------------------------------//-----------------------//
                    
                      
            
            
            Debug.Print "=========================NODO CLICCATO =================================================="
            
            
    
     '...............................................................................
        
        
        '//TODO: attenzione occorre passare 3 parametri
          '//TODO: da modificare per allinerlo al NODO. Occorre estrarre dalla combinata _
                                                                a) colonna 0 = il menu _
                                                                b) colonna 1 = IL CODICE DEL MENU _
                                                                c) colonna 3 = IL PROGETTO  a cui assegnato il menu _
                                                                Inoltre occorre modificare la select, in quanto bisogna attivare solo la classe _
                                                                e passare 3 parametri mentre i controlli li fa la classe GESTIONE TUTTI.
                                                                'COME FARE??? FORSE OCCORRE MODIFICARE LA TABELLA TREE???
                                                      
        
        '//SISTEMARE vedere aggiustarlo con i vari menu tipo , condominio, lavori pubblici ecc.. _
        CERCARE Di non ripertere il codice come quello previsto alla cmb01 della form.
        'COME CHIAMARE LA CLASSE MENU LAVORI PUBBLICI??
        'Diverso da Menu Gestioni altrimenti non apre i menu con le crocette.
        If CLASSE_NODOTreeView_s <> "Gestioni" Then
            '//Alla classe passo il MENU gestito NEL progetto CONDOMINIO
        '    my_PROGETTI_GESTIONE_TUTTI_Classe.IMPOSTA_GESTIONI CLASSE_NODOTreeView_s, "PROGETTO LAVORI PUBBLICI"
                
        
        End If
        
        '//libero la memoria
        Set my_PROGETTI_GESTIONE_TUTTI_Classe = Nothing

     
     
        

End Sub

'//*** FINE ***
'//LIST VIEW EVENTO CLICK SCELTA SELEZIONE
'//***********************************************************************************************************//


'//***********************************************************************************************************//
'//                     EVENTI DELLA COMBINATA 01 *** INIZIO ***
'// @cmb_01.03.@eventi
'//***********************************************************************************************************//


Private Sub Cmb_01_AfterUpdate()
        
    If Trim(Me.Cmb_01.Value) > "" Then
        Str1 = Me.Cmb_01.Value
        sSql = ""
        sSql = sSql & "SELECT Tree.*, Tree.NomeCartella_s AS RICERCA "
        sSql = sSql & "FROM Tree "
        sSql = sSql & "WHERE (((Tree.NomeCartella_s) Like '" & "*" & Str1 & "*" & "'));"
    Else
        sSql = ""
        sSql = sSql & "SELECT Tree.* FROM Tree; "
        
    
    End If
               



    Debug.Print sSql
    
    Me.SottoForm_01.Form.RecordSource = sSql '""
    
    'Me.SottoForm_01.SourceObject = sSql
    Me.SottoForm_01.Requery

End Sub



'//doppioclic della combinata 01 APRE TABELLA
Private Sub Cmb_01_DblClick(Cancel As Integer)
    '//LA TABELLA DEI COMANDI.
    DoCmd.OpenTable "Tree_DF01_COMANDI"
End Sub



'//***********************************************************************************************************//
'//                     EVENTI DELLA COMBINATA 01 *** FINE ***
'//
'//***********************************************************************************************************//



'//**************************************************************************************************************************//
'//                         HELP  ***INIZIO****
'//**************************************************************************************************************************//


Private Sub HELP_OGGETTO_TREE_Click()

'//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
'//--------------------------------------------------------------------------------//--------//
'//NOTE                 -> Apro il file di tipo doc, zip o pdf
'//CODICE               -> Function ApriFilePdf.01.01
'//PARAMETRI            -> par_Path_s         = PATH _
                        -> par_NameFile_s               = NOME FILE _
                        -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                        
  '//IMPOSTO LE VARIABILI PATH E FILE
  '//
  Dim MyFile_s As String
  '//
  Dim MyPath_s  As String
  
        '//imposto i parametri - LA PATH
        '
        MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\TREVIEW\OBIECT\HELP\"
                
        '//@01.02_HELP_OGGETTO_TREE
        '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
        MyFile_s = "HELP_OGGETTO_TREE.pdf "
            
            
         '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
            Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto IMPORTA OGGETTI incluso in questa form _
            anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'importazione e cancellazione in caso di resettamento _
            totale.
         Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
        
             
        
'//--------------------------------------------------------------------------------//--------//


End Sub




'//**************************************************************************************************************************//
'//                         HELP  ***FINE****
'//**************************************************************************************************************************//



'//**************************************************************************************************************************//
'//                         FILE  ***INIZIO****
'//**************************************************************************************************************************//



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




'//**************************************************************************************************************************//
'//                         FILE  ***FINE****
'//**************************************************************************************************************************//

