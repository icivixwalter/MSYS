README.md

# Note
		progetto Treeview

# LIBRERIE NECESSARIE
	ERRORI
		se manca la libreia ado:

		'// ERRORE *** = MANCA LA LIBRERIA ADO occorre fare riferimento ad ADO: _
			1. Verifica se ADO è installato sul tuo sistema _
			In genere, le librerie ADO sono già presenti su Windows, ma se non riesci a trovarle, potrebbe esserci un problema con la registrazione. Le versioni più comuni di ADO dovrebbero trovarsi nei seguenti percorsi: _
			ADO 6.1: C:\Program Files\Common Files\System\ado\msado15.dll _
			ADO 2.8: C:\Program Files (x86)\Common Files\System\ado\msado15.dll (su sistemi a 64 bit). _


		CARICARE_LA_LBRERIA_ADO
			Per caricare  la libreria ADO si trova qui:
	   	c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\DLL_LIBRERIE\MSCOMCTL.OCX


		TREEVIEW_REGISTRAZIONE_LIBRERIA_ADO
			In caso di errore perchè manca la libreria ado 

	
# GIORNALE_DELLE_REVISIONI
01 FORM TreeView_FrmS01_GE_TREEVIEW
	2024.07.21
		ho modicato la combinata cmb_01 impostando al momento del caricamento della form
		il valore ">>" e all'impostazione null della combinata, in modo da far capire che le doppie
		virgolette angolari hanno l'evento doppio click dove si puo aprire la tabella Tree_DF01_COMANDI
			
			vedi codice ----> @01_modifica_cmb_01

	01.02
		Inserito help nella form denominato HELP_OGGETTO_TREE

			vedi codice ----> @01.02_HELP_OGGETTO_TREE

	01.03
		Riordinato tutti gli eventi della casella combinata 01 cmb_01 in un unico posto.
			vedi codice ----> @cmb_01.03.@eventi

	01.04
		Creato ed riordinato gli eventi della casella combinata 02 cmb_02 in unico posto.
			vedi codice ----> @cmb_02.04.@eventi
			
02 FORM TreeView_Tb01Frm01_M01_GE_ALBERO
	02.01
		Reimposto la combinata all'evento load della form .
		vedi codice ----> @02.01.cmb_02.@eventi

		    '//IMPOSTO IL VALORE DEFINITO DELLA COMBINATA 02 -
		    '//@02.01.cmb_02.@eventi	 = quando la combinata è nulla imposta il valore di default >>
		    

	02.02
		IMPOSTO IL VALORE DEFINITO DELLA COMBINATA 02 SE NULL 
		vedi codice ----> @02.2_modifica_cmb_02	 = quando la combinata è nulla imposta il valore di default >>

	
	2024_07_24_TREVIEW = AGGIORNATO macro e creata la TREVIEW_Mcr02_STAMPA_LIBRERIE_ATTIVE
		che permette di stampare le librerie attive e funzionanti per il treeview
		con la funzione



