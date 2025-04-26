VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//SALVATAGGI DEL CODICE DI CLASSE : ----> c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\OBJECT\MODULI\clsFile.cls


'//@CLASSE@API_(classe per l'utilizzo delle api windows per la gestione dei file)
'// la routine di attivazione della classe si trova qui _
    nel modulo clsFile_MODULO_ATTIVAZIONE_API, vedi questo codice : ----> @MODELLO@UTILIZZO@API@WINDOWS


'Based on the example on msdn:
'http://support.microsoft.com/kb/189981 _
    ALTRO: https://ftp.zx.net.nz/pub/archive/ftp.microsoft.com/MISC/KB/en-us/189/981.HTM _

    'VEDI L 'ESEMPIO QUI: https://stackoverflow.com/questions/20888220/is-there-a-way-to-navigate-backwards-in-a-textstream-file


'Some of the constants come from Winnt.h
'note: Mi sono imbattuto nella stessa frustrante limitazione. Di seguito è riportata una classe che esegue il wrapping dell'API Windows nativa _
       per eseguire l'I/O dei file. Come notato, si basa sull'esempio su msdn a http://support.microsoft.com/kb/189981. Non ho finito di testarlo a fondo, _
       quindi se riscontri problemi, fammelo sapere in modo che possa risolverli a vantaggio di entrambi. _
       Come nota a margine, le cose CanRead, CanWrite sono lì in modo da poter eventualmente implementare un'interfaccia di flusso, ma questo è un progetto futuro.
       





Option Compare Database
Option Explicit
       
       
Public Enum SeekOrigin
    so_Begin = 0
    so_Current = 1
    so_End = 2
End Enum

Public Enum FileAccess
'    FILE_READ_DATA = &H1                     ' winnt.h:1801
'    'FILE_LIST_DIRECTORY = &H1                ' winnt.h:1802
'    FILE_WRITE_DATA = &H2                    ' winnt.h:1804
'    'FILE_ADD_FILE = &H2                      ' winnt.h:1805
'    FILE_APPEND_DATA = &H4                   ' winnt.h:1807
'    'FILE_ADD_SUBDIRECTORY = &H4              ' winnt.h:1808
'    'FILE_CREATE_PIPE_INSTANCE = &H4          ' winnt.h:1809
'    FILE_READ_EA = &H8                       ' winnt.h:1811
'    FILE_READ_PROPERTIES = &H8               ' winnt.h:1812
'    FILE_WRITE_EA = &H10                     ' winnt.h:1814
'    FILE_WRITE_PROPERTIES = &H10             ' winnt.h:1815
'    FILE_EXECUTE = &H20                      ' winnt.h:1817
'    'FILE_TRAVERSE = &H20                     ' winnt.h:1818
'    'FILE_DELETE_CHILD = &H40                 ' winnt.h:1820
'    FILE_READ_ATTRIBUTES = &H80              ' winnt.h:1822
'    FILE_WRITE_ATTRIBUTES = &H100            ' winnt.h:1824
    FILE_ALL_ACCESS = &H1F01FF               ' winnt.h:1826
    FILE_GENERIC_READ = &H120089             ' winnt.h:1828
    FILE_GENERIC_WRITE = &H120116            ' winnt.h:1835
'    FILE_GENERIC_EXECUTE = &H1200A0          ' winnt.h:1843
'    FILE_SHARE_READ = &H1                    ' winnt.h:1848
'    FILE_SHARE_WRITE = &H2                   ' winnt.h:1849
'    FILE_NOTIFY_CHANGE_FILE_NAME = &H1       ' winnt.h:1860
'    FILE_NOTIFY_CHANGE_DIR_NAME = &H2        ' winnt.h:1861
'    FILE_NOTIFY_CHANGE_ATTRIBUTES = &H4      ' winnt.h:1862
'    FILE_NOTIFY_CHANGE_SIZE = &H8            ' winnt.h:1863
'    FILE_NOTIFY_CHANGE_LAST_WRITE = &H10     ' winnt.h:1864
'    FILE_NOTIFY_CHANGE_SECURITY = &H100      ' winnt.h:1865
'    'MAILSLOT_NO_MESSAGE = -1                ' winnt.h:1866
'    'MAILSLOT_WAIT_FOREVER = -1              ' winnt.h:1867
'    FILE_CASE_SENSITIVE_SEARCH = &H1         ' winnt.h:1868
'    FILE_CASE_PRESERVED_NAMES = &H2          ' winnt.h:1869
'    FILE_UNICODE_ON_DISK = &H4               ' winnt.h:1870
'    FILE_PERSISTENT_ACLS = &H8               ' winnt.h:1871
'    FILE_FILE_COMPRESSION = &H10             ' winnt.h:1872
'    FILE_VOLUME_IS_COMPRESSED = &H8000       ' winnt.h:1873
'    IO_COMPLETION_MODIFY_STATE = &H2         ' winnt.h:1874
'    IO_COMPLETION_ALL_ACCESS = &H1F0003      ' winnt.h:1875
'    DUPLICATE_CLOSE_SOURCE = &H1             ' winnt.h:1876
'    DUPLICATE_SAME_ACCESS = &H2              ' winnt.h:1877
'    DELETE = &H10000                         ' winnt.h:1935
'    READ_CONTROL = &H20000                   ' winnt.h:1936
'    WRITE_DAC = &H40000                      ' winnt.h:1937
'    WRITE_OWNER = &H80000                    ' winnt.h:1938
'    SYNCHRONIZE = &H100000                   ' winnt.h:1939
'    STANDARD_RIGHTS_REQUIRED = &HF0000       ' winnt.h:1941
'    STANDARD_RIGHTS_READ = &H20000           ' winnt.h:1943
'    STANDARD_RIGHTS_WRITE = &H20000          ' winnt.h:1944
'    STANDARD_RIGHTS_EXECUTE = &H20000        ' winnt.h:1945
'    STANDARD_RIGHTS_ALL = &H1F0000           ' winnt.h:1947
'    SPECIFIC_RIGHTS_ALL = &HFFFF             ' winnt.h:1949
'    ACCESS_SYSTEM_SECURITY = &H1000000
End Enum


Public Enum FileShare
    NONE = &H0
    FILE_SHARE_DELETE = &H4
    FILE_SHARE_READ = &H1
    FILE_SHARE_WRITE = &H2
End Enum


Public Enum FileCreationDisposition
    CREATE_ALWAYS = &H2
    CREATE_NEW = &H1
    OPEN_ALWAYS = &H4
    OPEN_EXISTING = &H3
    TRUNCATE_EXISTING = &H5
End Enum


'Public Enum FileFlagsAndAttributes
'    'Attributes
'    FILE_ATTRIBUTE_ENCRYPTED = &H4000
'    FILE_ATTRIBUTE_READONLY = &H1            ' winnt.h:1850
'    FILE_ATTRIBUTE_HIDDEN = &H2              ' winnt.h:1851
'    FILE_ATTRIBUTE_SYSTEM = &H4              ' winnt.h:1852
'    FILE_ATTRIBUTE_DIRECTORY = &H10          ' winnt.h:1853
'    FILE_ATTRIBUTE_ARCHIVE = &H20            ' winnt.h:1854
'    FILE_ATTRIBUTE_NORMAL = &H80             ' winnt.h:1855
'    FILE_ATTRIBUTE_TEMPORARY = &H100         ' winnt.h:1856
'    FILE_ATTRIBUTE_ATOMIC_WRITE = &H200      ' winnt.h:1857
'    FILE_ATTRIBUTE_XACTION_WRITE = &H400     ' winnt.h:1858
'    FILE_ATTRIBUTE_COMPRESSED = &H800        ' winnt.h:1859
'    'Flags
'    FILE_FLAG_BACKUP_SEMANTICS = &H2000000
'    FILE_FLAG_DELETE_ON_CLOSE = &H4000000
'    FILE_FLAG_NO_BUFFERING = &H20000000
'    FILE_FLAG_OPEN_NO_RECALL = &H100000
'    FILE_FLAG_OPEN_REPARSE_POINT = &H200000
'    FILE_FLAG_OVERLAPPED = &H40000000
'    FILE_FLAG_POSIX_SEMANTICS = &H100000
'End Enum


Private Const INVALID_FILE_HANDLE = -1 '&HFFFFFFFF
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const INVALID_FILE_SIZE As Long = -1 '&HFFFFFFFF
Private Const INVALID_SET_FILE_POINTER As Long = -1 '&HFFFFFFFF


Private Declare Function FormatMessage Lib "Kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, _
                                                                              lpSource As Long, _
                                                                              ByVal dwMessageId As Long, _
                                                                              ByVal dwLanguageId As Long, _
                                                                              ByVal lpBuffer As String, _
                                                                              ByVal nSize As Long, _
                                                                              Arguments As Any) As Long


Private Declare Function CreateFile Lib "Kernel32" Alias "CreateFileA" (ByVal lpFileName As String, _
                                                                        ByVal dwDesiredAccess As Long, _
                                                                        ByVal dwShareMode As Long, _
                                                                        lpSecurityAttributes As Long, _
                                                                        ByVal dwCreationDisposition As Long, _
                                                                        ByVal dwFlagsAndAttributes As Long, _
                                                                        hTemplateFile As Long) As Long


Private Declare Function SetFilePointer Lib "Kernel32" (ByVal hFile As Long, _
                                               ByVal lDistanceToMove As Long, _
                                               lpDistanceToMoveHigh As Long, _
                                               ByVal dwMoveMethod As Long) As Long


Private Declare Function ReadFile Lib "Kernel32" (ByVal hFile As Long, _
                                                  lpBuffer As Any, _
                                                  ByVal nNumberOfBytesToRead As Long, _
                                                  lpNumberOfBytesRead As Long, _
                                                  ByVal lpOverlapped As Long) As Long


Private Declare Function WriteFile Lib "Kernel32" (ByVal hFile As Long, _
                                                   lpBuffer As Any, _
                                                   ByVal nNumberOfBytesToWrite As Long, _
                                                   lpNumberOfBytesWritten As Long, _
                                                   ByVal lpOverlapped As Long) As Long


Private Declare Function FlushFileBuffers Lib "Kernel32" (ByVal hFile As Long) As Long


Private Declare Function GetFileSize Lib "Kernel32" (ByVal hFile As Long, _
                                                     lpFileSizeHigh As Long) As Long


Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (ByVal dest As Long, ByVal src As Long, ByVal size As Long)

Private m_Handle As Long

Private Sub Class_Terminate()
    If Not m_Handle = 0 Then
        Flush
        CloseFile
    End If
End Sub

Public Sub OpenFile(path As String, Optional access As FileAccess = FileAccess.FILE_GENERIC_READ, Optional share As FileShare = FileShare.NONE, Optional CreationDisposition As FileCreationDisposition = FileCreationDisposition.OPEN_ALWAYS)
    Dim Ret As Long
    Ret = CreateFile(path, access, share, ByVal 0&, CreationDisposition, 0&, ByVal 0&)
    If Ret = INVALID_FILE_HANDLE Then
        Err.Raise vbObjectError + Err.LastDllError, "clsFile.OpenFile", DecodeAPIErrors(Err.LastDllError)
    Else
        m_Handle = Ret
    End If
End Sub

'Properties

Public Property Get Length() As Double
    Dim Ret As Currency
    Dim FileSizeHigh As Long
    Ret = GetFileSize(m_Handle, FileSizeHigh)
    If Not Ret = INVALID_FILE_SIZE Then
        Length = Ret
    Else
        Err.Raise vbObjectError + Err.LastDllError, "clsFile.Length", DecodeAPIErrors(Err.LastDllError)
    End If
End Property

Public Property Get Position() As Long
    Dim Ret As Long
    Dim DistanceToMoveHigh As Long
    Ret = SetFilePointer(m_Handle, 0&, DistanceToMoveHigh, 1&) '1 is FILE_CURRENT
    If DistanceToMoveHigh = 0 Then
        If Ret = -1 Then
            Position = -1 'EOF'
        Else
            Position = Ret
        End If
    Else
        Class_Terminate
        Err.Raise vbObjectError + Err.LastDllError, "clsFile.Position", DecodeAPIErrors(Err.LastDllError)
    End If
End Property

Public Property Get Handle() As Long
    Handle = m_Handle
End Property

'Functions

Public Function ReadBytes(ByRef buffer() As Byte, ByVal buffer_offset As Long, ByVal count As Long) As Long
    Dim Ret As Long
    Dim BytesRead As Long
    Ret = ReadFile(m_Handle, buffer(buffer_offset), count, BytesRead, 0&)
    If Ret = 1 Then
        ReadBytes = BytesRead
    Else
        Class_Terminate
        Err.Raise vbObjectError + Err.LastDllError, "clsFile.ReadBytes", DecodeAPIErrors(Err.LastDllError)
    End If
End Function

Public Function ReadBytesPtr(ByVal ptrBuf As Long, ByVal buffer_offset As Long, ByVal count As Long) As Long
    Dim Ret As Long
    Dim BytesRead As Long
    Ret = ReadFile(m_Handle, ByVal ptrBuf + buffer_offset, count, BytesRead, 0&)
    If Ret = 1 Then
        ReadBytesPtr = BytesRead
    Else
        Class_Terminate
        Err.Raise vbObjectError + Err.LastDllError, "clsFile.ReadBytesPtr", DecodeAPIErrors(Err.LastDllError)
    End If
End Function

Public Function WriteBytes(ByRef buffer() As Byte, ByVal buffer_offset As Long, ByVal count As Long) As Long
    Dim Ret As Long
    Dim BytesWritten As Long
    Ret = WriteFile(m_Handle, buffer(buffer_offset), count, BytesWritten, 0&)
    If Ret = 1 Then
        WriteBytes = BytesWritten
    Else
        Class_Terminate
        Err.Raise vbObjectError + Err.LastDllError, "clsFile.WriteBytes", DecodeAPIErrors(Err.LastDllError)
    End If
End Function

Public Function WriteBytesPtr(ByVal ptrBuf As Long, ByVal buffer_offset As Long, ByVal count As Long) As Long
    Dim Ret As Long
    Dim BytesWritten As Long
    Ret = WriteFile(m_Handle, ByVal ptrBuf + buffer_offset, count, BytesWritten, 0&)
    If Ret = 1 Then
        WriteBytesPtr = BytesWritten
    Else
        Class_Terminate
        Err.Raise vbObjectError + Err.LastDllError, "clsFile.WriteBytes", DecodeAPIErrors(Err.LastDllError)
    End If
End Function

Public Function SeekFile(ByVal LoBytesOffset As Long, origin As SeekOrigin) As Long
    Dim Ret As Long
    Dim HiBytesOffset As Long
    Ret = SetFilePointer(m_Handle, LoBytesOffset, HiBytesOffset, origin)
    If Not Ret = INVALID_SET_FILE_POINTER Then
        SeekFile = Ret
    Else
        Err.Raise vbObjectError + Err.LastDllError, "clsFile.SeekFile", DecodeAPIErrors(Err.LastDllError)
    End If
End Function

Public Function SeekFile64bit(ByVal offset As Currency, origin As SeekOrigin) As Currency
'Take care with this function. A Currency variable is an 8-byte (64-bit) scaled (by 10,000) fixed-point number.'
'This means that setting a Currency variable to 0.0001 is the equivalent of a binary value of 1.'
'If you want to set an offset with an immediate value, write it like so:'
'1073741824 Bytes (1 GB) would be 107374.1824@, where @ is the symbol for an immediate Currency value.'
'Refer to http://support.microsoft.com/kb/189862 for hints on how to do 64-bit arithmetic'
    Dim Ret As Long
    Dim curFilePosition As Currency
    Dim LoBytesOffset As Long, HiBytesOffset As Long

    CopyMemory VarPtr(HiBytesOffset), VarPtr(offset) + 4, 4
    CopyMemory VarPtr(LoBytesOffset), VarPtr(offset), 4

    Ret = SetFilePointer(m_Handle, LoBytesOffset, HiBytesOffset, origin)

    CopyMemory VarPtr(curFilePosition) + 4, VarPtr(HiBytesOffset), 4
    CopyMemory VarPtr(curFilePosition), VarPtr(Ret), 4

    SeekFile64bit = curFilePosition
End Function

Public Sub CloseFile()
    Dim Ret As Long
    Ret = CloseHandle(m_Handle)
    m_Handle = 0
End Sub

Public Sub Flush()
    Dim Ret As Long
    Ret = FlushFileBuffers(m_Handle)
End Sub

 '***********************************************************************************
' Helper function, from Microsoft page as noted at top
Private Function DecodeAPIErrors(ByVal ErrorCode As Long) As String
    Dim sMessage As String, MessageLength As Long
    sMessage = Space$(256)
    MessageLength = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0&, _
                                  ErrorCode, 0&, sMessage, 256&, 0&)
    If MessageLength > 0 Then
        DecodeAPIErrors = Left(sMessage, MessageLength)
    Else
        DecodeAPIErrors = "Unknown Error."
    End If
End Function




