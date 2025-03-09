UTILITA_Msys_Frm21_IMPORTA_OGGETTI.md
   @form@UTILITA_(classe della form utilita)

@INIZIO

### CODICE_CLASSE








Option Compare Database



' Metodo 1: Usare un Controllo ProgressBar
'Se vuoi visualizzare la barra in una form dedicata, puoi usare un controllo ProgressBar su una form e aggiornarlo dinamicamente con VBA.

'Passaggi:
'1?? Crea una nuova Form (es. frmProgressBar).
'2?? Aggiungi una Casella di Testo (txtStatusMsg) per i messaggi di avanzamento.
'3?? Aggiungi una Etichetta (lblPercentuale) per la percentuale.
'4?? Aggiungi un Oggetto Rettangolo (boxProgressBar) per la barra grafica.

'?? Opzione 1: Usare un Rettangolo (Metodo più semplice)
'1?? Apri la Form frmProgressBar in modalità Struttura
'2?? Inserisci un Controllo Rettangolo:

'Vai su: Strumenti di progettazione > Controlli > Rettangolo
'Disegna il rettangolo e chiamalo, per esempio, boxProgressBar
'Imposta Colore Riempimento = Verde per simularne il riempimento
'3?? Aggiungi una Label chiamata lblPercentuale per visualizzare la percentuale.

'4?? Modifica il codice VBA per aggiornare la larghezza del rettangolo:

Dim IncrProgresso_dbl_i As Long
Dim frm As Form

'// LA FUNZIONE ATTIVA PROGRESS BAR RICEVE QUESTI PARAMETRI
'//...................................................................................................//
'//Note: passo i 5 parametri per aggiornare la label e il rettangolo alla funzione generale:
'//par_NroCollection_i    = Nro di oggetti collection
'//par_progressValue_dbl  = La progressione di incremento
'//par_Twips_dbl          = i twips per ogni oggetto da importare
'//par_ProgressPercentTwips_dbl        = la percetuale di incremento per ogni oggetto da importare
'//par_TipoImport_s       = il tipo di oggetto di importazione TOT IMPORTATI - TOT NON IMPORTATI

Public Function AttivaBarraDiProgressione(par_TotCollectioForms_i As Integer, _
                                          par_ProgressPercentTwips_dbl As Double, _
                                          par_Twips_dbl As Double, _
                                          par_TotPerc_dbl As Double, _
                                          par_TipoImport_s As String)
   
        '//chiamo le due routine private di aggiornamento LABEL e RETTANGOLO
        AggiornaBarralabel par_TotPerc_dbl, par_Twips_dbl, par_TipoImport_s
        
        AggiornaBarra par_ProgressPercentTwips_dbl, par_Twips_dbl, par_TipoImport_s
        
        DoEvents ' Aggiorna interfaccia
        
        ' Aggiorna la UI
        Me.Recalc
        

End Function
    
    
'//@AGGIORNA@PROGRESS@BAR
Private Sub AggiornaBarra(par_ProgressPercentTwips_dbl As Double, par_Twips_dbl As Double, par_TipoImport_s As String)
    'Set frm = forms("frmProgressBar")

    ' Imposta la lunghezza della barra (Assumiamo una larghezza max di 10)
    Me.boxProgressBar.Width = (par_ProgressPercentTwips_dbl)
    

    
End Sub
'//@AGGIORNO@LABEL_(aggiorno i dati della labelI
Private Sub AggiornaBarralabel(par_TotPerc_dbl As Double, par_Twips_dbl As Double, par_TipoImport_s As String)
        
    '//indico la percentuale di aggiornamento
    Me.lblPercentuale.Caption = par_TotPerc_dbl & "% completato"
    
    '//@messaggio@casella.di@testo
    Me.txtStatusMsg.Value = par_TotPerc_dbl & "% completato" & " --->  " & par_TipoImport_s & "   -   twips --->  " & par_Twips_dbl
    

End Sub





























