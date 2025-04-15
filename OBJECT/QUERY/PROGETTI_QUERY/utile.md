Msys_DLL_LIBRERIEQry01_01_LIBRERIE_ATTIVE.md

	Note
		questa è la query di estrazione delle libreria salvate ed attive
		nel progetto corrente


	schema


		DELETE 
			Msys_DLL_LIBRERIE.*
			FROM 
			Msys_DLL_LIBRERIE;



		 sSql = ""
            sSql = sSql & "INSERT INTO Msys_DLL_LIBRERIE (LIBRERIA_DLL_s, COD_PROGETTO_s, PROGETTO_s, NRO_i, NAME_s, GUID_s, PATH_s )"
            sSql = sSql & "SELECT " & "'" & myNameDLL & "'" & " AS MyLIBRERIA_DLL_s," & "'MSys'" & "AS MyCOD_PROGETTO_s," & "'PROGETTO MSys'" & " AS MyPROGETTO_s," & iCount & " AS NRO_is,'" & ref.Name & "' AS NAME_s, '" & ref.Guid & "' AS GUID_s,'" & ref.fullPath & "' AS PATH_s "
            sSql = sSql & "WITH OWNERACCESS OPTION;"


