# QUERY DI BACKUP

## TODO
è quella sql necessaria per l'esportazione dei dati di backup
Msys_TB99_LINK_Qry99_01_SELECT_BackUp


## SQL
SELECT 
      Msys_TB99_LINK.Name, 
      Msys_TB99_LINK.TIPO_OGGETTO_c
   
   FROM 
      Msys_TB99_LINK
   WHERE 
      (((Msys_TB99_LINK.Name) Not Like '*[_]{@*' And 
      (Msys_TB99_LINK.Name) Not Like '*[_]}*' And 
      (Msys_TB99_LINK.Name) Is Not Null And 
      (Msys_TB99_LINK.Name)<>''));
      
