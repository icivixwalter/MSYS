MSys_TUTTE_BARRE_DEI_COMANDI_Qry01_01_SELECT_Archivio

   Note
      Query di estrazione delle barre dei comandi, tutte.
      Archivio delle barre degli strumenti
         @tutte@le@barre@degli@strumenti
         @barre@dei@comandi
         


   SQL

      SELECT 

      MSys_TUTTE_BARRE_DEI_COMANDI.CodMenu_s, 
      MSys_TUTTE_BARRE_DEI_COMANDI.Menu_s, 
      MSys_TUTTE_BARRE_DEI_COMANDI.NoteMenu_s, 
      MSys_TUTTE_BARRE_DEI_COMANDI.Visible_b, 
      MSys_TUTTE_BARRE_DEI_COMANDI.Enabled_b, 
      MSys_TUTTE_BARRE_DEI_COMANDI.Position_i, 
      MSys_TUTTE_BARRE_DEI_COMANDI.NameLocal_s, 
      MSys_TUTTE_BARRE_DEI_COMANDI.Nro_i, 

      MSys_TUTTE_BARRE_DEI_COMANDI.[02________________________________________UtilitaGenerali], 
      MSys_TUTTE_BARRE_DEI_COMANDI.bCanc, 
      MSys_TUTTE_BARRE_DEI_COMANDI.bAgg, 
      MSys_TUTTE_BARRE_DEI_COMANDI.bScel, 
      MSys_TUTTE_BARRE_DEI_COMANDI.IDKeyArch, 
      MSys_TUTTE_BARRE_DEI_COMANDI.IdOrd, 
      MSys_TUTTE_BARRE_DEI_COMANDI.Msg01, 

      MSys_TUTTE_BARRE_DEI_COMANDI.[AG___________________________________Agg+Id], 
      MSys_TUTTE_BARRE_DEI_COMANDI.DATAINS_d, 
      MSys_TUTTE_BARRE_DEI_COMANDI.DATAAGG_d, 
      MSys_TUTTE_BARRE_DEI_COMANDI.TIMEOPER_d, 
      MSys_TUTTE_BARRE_DEI_COMANDI.ORA_AGG_d, 
      MSys_TUTTE_BARRE_DEI_COMANDI.ID

      FROM 
      MSys_TUTTE_BARRE_DEI_COMANDI;
