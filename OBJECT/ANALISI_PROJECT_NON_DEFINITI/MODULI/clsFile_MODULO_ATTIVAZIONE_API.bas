Attribute VB_Name = "clsFile_MODULO_ATTIVAZIONE_API"
Option Compare Database


'//SALVATAGGI DEL CODICE DEL MODULO : ----> c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\OBJECT\MODULI\clsFile_MODULO_ATTIVAZIONE_API.bas


'//@MODELLO@UTILIZZO@API@WINDOWS_(questa la la routine di prova per l'attivazione delle API per la gestione dei file in windows)

'// @UTILIZO@API@RICERCA@FILE_(Modello di utilizzo delle api windwos per la ricerca dei file)

'//la classe per l'utilizzo delle api windows si trova in questa classe clsFile, vedi codice: _
        ----> @CLASSE@API

'Ed ecco un esempio di come usarlo: _
  2024.09.18 = HO ATTIVATO L'ESEMPIO DI RICERCA DEL FILE e qui funziona se vuoi applicarlo per un'altro file _
                devi cambiare path + file.

Public Sub Main()
    Dim oFile As clsFile
    Set oFile = New clsFile             '//per la ricerca UTILIZZO LA CLASSE ClsFile

    oFile.OpenFile "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\\MSYS_README.MD", FILE_GENERIC_READ, NONE, OPEN_EXISTING

    Dim ChunkOfData() As Byte
    Const CHUNKSIZE As Long = 4096
    ReDim ChunkOfData(0 To CHUNKSIZE - 1)

    Dim lngCurrChunk As Long
    Dim lngBytesRead As Double


    'The SeekFile function works for seeks forward or backward in the file from [-2GB to +2GB).'
    'Past that you can use the SeekFile64bit function, but you'll have to be aware of the issues with using Currency to store the 64-bit number'
    Debug.Print oFile.SeekFile(&H40000000, so_Current) 'A 1GB seek

    lngBytesRead = oFile.ReadBytes(ChunkOfData, 0, CHUNKSIZE)
    While lngBytesRead > 0 'As soon as a call to ReadBytes returns 0, we've reached the end of the file.
        'Do something with the 4k chunk of data.  The buffer gets reused in this example.
        'Debug.Print ChunkOfData
        lngCurrChunk = lngCurrChunk + 1
        lngBytesRead = oFile.ReadBytes(ChunkOfData, 0, CHUNKSIZE)
    Wend
    MsgBox "Complete!"
End Sub

