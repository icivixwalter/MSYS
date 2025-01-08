SottoFormMaster_Tb01_GE_Msys.md
	
	Note
		tabella di gestione delle SOTTOFORM MASTER A CARICAMENTO AUTOMATICO
		Per ogni sottoform occorre defire due chiavi

		@CODICE ----> @SOTTOFORM@MASTER@CARICAMENTO@AUTOMATICO_(la tabella TB01 di gestione dei codici)_@TABELLA@tb01




	CAMPI
		CHIAVI
				Note
					per le due chiavi CODICE PROGETTO E DENOMINAZIONE non sono ammessi duplicati per
					gestire in modo separato le SOTTOFORM MASTER a CARICAMENTO AUTOMATICO
			COD_PROGETTO_s			= pkey - chiave codice progetto 
			DENOM_PROGETTO_s		= pkey - chiave DENOMINAZIONE PROGETTO

		CAMPI

			COD_FORMS_s				= CAMPO CODICE FORM che deve essere univoco in modo da distinguere
									le diverse sottoform MASTER a caricamento automatico. In questo modo 
									si è costretti a creare SOTTOFORM MASTER con denominazioni diverse per
									evitare conflitti di nomi.

	PROPRIETA
				Note
					la tabella TB01 gestisce i codici identificativi delle sottoform MASTER ed collegata
					con la sottotabella denominata

						SottoFormMaster_Tb02_Vis_Pagine_Msys
