CALL "MODELLO_BAT_START_Aprofile_LATEX_(caricaDiscoEsterno).bat"
CALL "MODELLO_BAT_START_Aprofile_MDB_(caricaDiscoEsterno).bat"
CALL "MODELLO_BAT_START_Aprofile_QUERY_MODULI_CLASSI_(caricaDiscoEsterno).bat"
CALL "MODELLO_BAT_START_Aprofile_SALVATAGGI_(caricaDiscoEsterno).bat"
CALL "MODELLO_BAT_START_Aprofile_TEXT_ANALISI_(caricaDiscoEsterno).bat"




//@PARAMETRI DA CAMBIARE	= sono le 6 chiavi da sostituire in tutto il progetto DA OLD A NEW
//@CHIAVI GENERALI DI SOSTITUZIONE		@chiavi.di.sostituzione
//-------------------------------------------------------------------------------------------------------------------//
//					OLD					|	NEW	

//SEZIONI PER IL CAMBIO COLLETTIVO DELLE PAROLE CHIAVI


//NOTA per cambia tutti gli elementi basta cambiare le parole chiavi in comune.
@CHIAVI IN COMUNE ---> GESTIONI | GESTIONE_LLPP| 25_GESTIONE_LLPP | LLPP_ARCHIVI_MDB | AA_MODELLO | MODELLO_ | C: |



//------>	il DISCO + LA PATH DEL PROGRAMMA
//
//@disco_programma	= C:
//@PATH_PROGRAMMA	= \GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\AA_MODELLO\


@PATH_01_s=\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\
@PATH_02_s=\LLPP_ARCHIVI_MDB\AA_MODELLO\
@PATH_03_s=\MODELLO_TEX+TXT_ANALISI\

@FILE_s=MODELLO_ANALISI_TEX.tex
@FILE_s=AA_GE_MODELLI_DbBase.mdb
@FILE_s=MODELLO_AREA_DI_LAVORO.tws



//PATH RICOSTRUITA		--->	c:\GESTIONI\GESTIONE_LLPP\PRES3000_GESTIONE\ARCHIVI_PRES3000\PRES3000_N25_SALARIO_ACCESSORIO\SALARIO_ACCESSORIO_BAT\
//-------------------------------------------------------------------------------------------------------------------//
//NOTA: operazioni da eseguirC: 1) posizionarsi sulla nuova directory appena costruita o gia defiita, 
//				2) copiare tutti i file bat di tipo MODELLO,
//				3) aprire il file .txt e tutti gli altri file .bat e posizionarsi nel file o DISCO... oppure ... APRI TUTTO per 
//				   sostituzioni collettive.
//				4) cambiare il nome dei file .bat da old a nuovo con la parola chiave (es da MODELLO_BAT_ a PIPO_BAT_, ....
//					oppure da MODELLO_BAT_ A SALARIO_ACCESSORIO_)
//				5) cambiare le PATH A) , B, C) da OLD a NEW
//				6) cambiare il disco da OLD a NEW.
//-------------------------------------------------------------------------------------------------------------------//


