VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cls_zip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'This mod is the famoes Inflate routine used by several different
'Compression programs like ZIP,gZip,PNG,etc..
'This module is created by Marco v/d Berg but is heavely optimized by John Korejwa


'This class file can be used to show/extract the contents of an ZIP-archive
Event Persent(ByVal p100 As Integer)
Private Type Local_Header_Type
    VerExt As Integer                   'version needed to extract
    Flags As Integer                    'encrypt and compression flags
    Method As Integer                   'compression method
    FTime As Integer                    'time last modifies, dos format
    FDate As Integer                    'date last modifies, dos format
    CRC32 As Long                       'CRC32 for uncompressed file
    CSize As Long                       'compressed size
    USize As Long                       'uncompressed size
    LenFname As Integer                 'lenght filename
    LenExt As Integer                   'lenght for extra field
End Type
Private Type Central_Header_Type
    VerMade As Integer                  'version made by
    VerExt As Integer                   'version needed to extract
    Flags As Integer                    'encrypt and compression flags
    Method As Integer                   'compression method
    FTime As Integer                    'time last modifies, dos format
    FDate As Integer                    'date last modifies, dos format
    CRC32 As Long                       'CRC32 for uncompressed file
    CSize As Long                       'compressed size
    USize As Long                       'uncompressed size
    LenFname As Integer                 'lenght filename
    LenExt As Integer                   'lenght for extra field
    LenCom As Integer                   'lenght for comment field
    DiskStart As Integer                'start disk number
    AttribI As Integer                  'internal file attributes
    AttribX As Long                     'external file attributes
    Offset As Long                      'relative offset of local header
End Type
Private Type End_Header_Type
    signature As Long                   'Signature
    DiskNum As Integer                  'this disk number
    DiskStart As Integer                'start disk number
    Entries As Integer                  'Entries on this disk
    TotEntr As Integer                  'Number of total entries
    CenSize As Long                     'size of entire cetral directory
    CenOff As Long                      'offset of central on starting disk
    LenCom As Integer                   'lenght of comment field
End Type
Private Type Extended_Local_Header_Type
    CRC32 As Long                       'CRC32 for uncompressed file
    CSize As Long                       'compressed size
    USize As Long                       'uncompressed size
End Type

Private Type CentralData_Type
    VerMade As Integer                  'version made by
    VerExt As Integer                   'version needed to extract
    Flags As Integer                    'encrypt and compression flags
    Method As Integer                   'compression method
    FTime As Integer                    'time last modifies, dos format
    FDate As Integer                    'date last modifies, dos format
    CRC32 As Long                       'CRC32 for uncompressed file
    CSize As Long                       'compressed size
    USize As Long                       'uncompressed size
    DiskStart As Integer                'start disk number
    AttribI As Integer                  'internal file attributes
    AttribX As Long                     'external file attributes
    Offset As Long                      'relative offset of local header
    FileName As String                  'Name of the compressed file
    ExtField As String                  'Data from extended fields
    ComField As String                  'data from comments fields
End Type


'Signatures long version
Private Const ZipLocalSigLng As Long = &H4034B50
Private Const ZipCentralSigLng As Long = &H2014B50
Private Const ZipEndSigLng As Long = &H6054B50
Private Const ZipExtLocalSigLng As Long = &H8074B50

'Flags values for ZIP-files
Private Const ZipFlgEncrypted As Byte = 1   'bit 0 set = file is encrypted
Private Const ZipFlgUsedMed As Byte = 6     'bit 1+2 depending on compression type
                                            'type = 6 (imploding)
                                            'bit 1 set = use 8k dictionary else 4k dictionary
                                            'bit 2 set = use 3 trees else use 2 trees
                                            'type = 8 (deflating)
                                            'bit 2 : 1
                                            '    0   0 = Normal (-en) compression option was used.
                                            '    0   1 = Maximum (-exx/-ex) compression option was used
                                            '    1   0 = Fast (-ef) compression option was used
                                            '    1   1 = Super Fast (-es) compression option was used
                                            'bits are undified if other methods are used
Private Const ZipFlgExtLocHead As Byte = 8  'bit 3 set = Extended local header is used to store CRC and size
Private Const ZipFlgRes64 As Byte = 16      'bit 4 Reserved for ZIP64
Private Const ZipFlgPathed As Byte = 32     'bit 5 set = file is compressed pathed data
Private Const ZipFlgEncStrong As Byte = 64  'bit 6 set = file is encrypted using strong encryption

Private ZIPFileData() As CentralData_Type
Private CRC As New Cls_CRC32
Private Encrypt As New Cls_Encrypt
Private Const m_Unpack_Supported As Boolean = True

Public Function Get_Contents(ZipName As String) As Integer
    Dim NextByte As Byte        '1 byte
    Dim Byte2 As Integer        '2 bytes
    Dim FileNum As Long
    Dim FileLenght As Long
    Dim Header As Integer
    Dim LngHeader As Long
    Dim UnKnown As Boolean
    Dim TextBytes() As Byte
    Dim EndCentralSig As End_Header_Type
    Dim CentDat As Central_Header_Type
    Dim LocDat As Local_Header_Type
    Dim ExtDat As Extended_Local_Header_Type
    Dim LN As Long
    Dim X As Long
    PackFileName = ZipName
    PackComments = ""
    PackFileType = 0
    FileNum = FreeFile
    Open PackFileName For Binary Access Read As #FileNum
    FileLenght = LOF(FileNum)
    'get the end of central date
    Get #FileNum, FileLenght - Len(EndCentralSig) + 1, EndCentralSig
    If EndCentralSig.signature <> ZipEndSigLng Then
        'EndSignature not found (probably comments are added)
        'Search for central data from start
        Seek #FileNum, 1
        PackTotFiles = 0
        Do
            Get #FileNum, , LngHeader
            Select Case LngHeader
            Case ZipLocalSigLng
                Get #FileNum, , LocDat
                Seek #FileNum, Seek(FileNum) + LocDat.CSize + LocDat.LenFname + LocDat.LenExt
                PackTotFiles = PackTotFiles + 1
            Case ZipCentralSigLng
                Seek #FileNum, Seek(FileNum) - 4
                Exit Do
            Case ZipExtLocalSigLng
                Get #FileNum, , ExtDat
            Case Else
                MsgBox "Unknown header"
            End Select
        Loop
    Else
        'position pointer to the start of the central header
        Seek #FileNum, EndCentralSig.CenOff + 1
        PackTotFiles = EndCentralSig.Entries
    End If
    ReDim ZIPFileData(1 To PackTotFiles)
    'Read the central header and store the information
    PackFileType = ZipFileType     'Zip file type
    For X = 1 To PackTotFiles
        Get #FileNum, , LngHeader       'read the header
        Get #FileNum, , CentDat         'read the data
        With ZIPFileData(X)
            .VerMade = CentDat.VerMade
            .VerExt = CentDat.VerExt
            .Flags = CentDat.Flags
            .Method = CentDat.Method
            .FTime = CentDat.FTime
            .FDate = CentDat.FDate
            .CRC32 = CentDat.CRC32
            .CSize = CentDat.CSize
            .USize = CentDat.USize
            .DiskStart = CentDat.DiskStart
            .AttribI = CentDat.AttribI
            .AttribX = CentDat.AttribX
            .Offset = CentDat.Offset
            If CentDat.LenFname <> 0 Then
                ReDim TextBytes(0 To Int2Lng(CentDat.LenFname) - 1)
                Get #FileNum, , TextBytes
                .FileName = StrConv(TextBytes, vbUnicode)
            End If
            If CentDat.LenExt <> 0 Then
                ReDim TextBytes(0 To Int2Lng(CentDat.LenExt) - 1)
                Get #FileNum, , TextBytes
                .ExtField = StrConv(TextBytes, vbUnicode)
            End If
            If CentDat.LenCom Then
                ReDim TextBytes(0 To Int2Lng(CentDat.LenCom) - 1)
                Get #FileNum, , TextBytes
                .ComField = StrConv(TextBytes, vbUnicode)
            End If
        End With
    Next
    Get #FileNum, , EndCentralSig
    PackComments = String(EndCentralSig.LenCom, 0)
    Get #FileNum, , PackComments
    Close FileNum
End Function

'Unzip as file and return 0 for good decompression or others for error
Public Function UnPack(ZippedFile() As Boolean, ToPath As String) As Integer
    Dim ZipHead As Local_Header_Type        'Local Zip Header
    Dim Header As Long
    Dim X As Long
    Dim PassWord As String
    Dim FileNum As Long
    Dim Y As Long
    Dim TotDir As String                    'Used for new pathnames
    If PackTotFiles = 0 Then UnPack = -10: Exit Function 'nothing to UnPack
    If PackTotFiles <> UBound(ZippedFile) Then
        UnPack = -10                         'need same amount as files in zipfile
        Exit Function
    End If
    Erase PackData
    FileNum = FreeFile
    Open PackFileName For Binary Access Read As #FileNum
    For X = 1 To PackTotFiles
        If ZippedFile(X) = True Then
            If Not IsDir(X) Then                'decompress if it is not a pathname
                Seek #FileNum, ZIPFileData(X).Offset + 1
                Get #FileNum, , Header
                If Header = ZipLocalSigLng Then
                    'read the header
                    Get #FileNum, , ZipHead
                    'skip data whe already know
                    Seek #FileNum, Seek(FileNum) + ZipHead.LenFname + ZipHead.LenExt
                    If ZIPFileData(X).CSize = 0 Then
                        Erase PackData
                    Else
                        ReDim PackData(ZIPFileData(X).CSize - 1)
                        Get #FileNum, , PackData()          'Read the compressed file
                    End If
                    If Encrypted(X) Then
                        If PassWord = "" Then
                            PassWord = InputBox("Give Password", "Password requered")
                            If PassWord = "" Then
                                UnPack = -1
                                Close FileNum
                                MsgBox "Password is incorrect"
                                Exit Function
                            End If
                        End If
                        Encrypt.ZipPrepareKey PackData, PassWord
                        If PackData(11) <> (((ZIPFileData(X).CRC32 And &HFF000000) \ &H1000000) And 255&) Then
                            UnPack = -1
                            Close FileNum
                            MsgBox "Password is incorrect"
                            Exit Function
                        End If
'adjust the size of instream to delete the decryption data
                        For Y = 0 To UBound(PackData) - 12
                            PackData(Y) = PackData(Y + 12)
                        Next
                        ReDim Preserve PackData(UBound(PackData) - 12)
                        Encrypt.ZipDecryptArray PackData
                    End If
                    Select Case ZIPFileData(X).Method
                    Case 0
                        Call Write_Uncompressed_Data(X, ToPath)
                    Case 1
                        UnPack = -10                 'not supported
                    Case 2
                        UnPack = UnReduce(PackData, 2, ZIPFileData(X).USize - 1)
                    Case 3
                        UnPack = UnReduce(PackData, 3, ZIPFileData(X).USize - 1)
                    Case 4
                        UnPack = UnReduce(PackData, 4, ZIPFileData(X).USize - 1)
                    Case 5
                        UnPack = UnReduce(PackData, 5, ZIPFileData(X).USize - 1)
                    Case 6
                        UnPack = -10                 'not supported
                    Case 7
                        UnPack = -10                 'not supported
                    Case 8
                        Call Inflate(PackData, ZIPFileData(X).USize - 1, False)
                        If gbZipStop Then Exit Function
                        Call Write_Uncompressed_Data(X, ToPath)
                    Case 9
                        Call Inflate(PackData, ZIPFileData(X).USize - 1, True)
                        Call Write_Uncompressed_Data(X, ToPath)
                    Case 10
                        UnPack = -10                 'not supported
                    End Select
                Else
                    MsgBox "Error in zipfile"
                End If
            Else
                TotDir = ToPath
                If Right(TotDir, 1) <> "\" And Right(TotDir, 1) <> "/" Then TotDir = TotDir & "\"
                TotDir = TotDir & ZIPFileData(X).FileName
                If CreatePath(TotDir) = False Then
'room for error message
                End If
            End If
        End If
    Next
    Close FileNum
    Erase PackData
End Function

Public Function Pack(ZipName As String, Files() As String, CompType As Integer, CompLevel As Integer, Optional IncludeDir As String = "") As Integer
    
End Function

Public Property Get CanUnpack() As Boolean
    CanUnpack = m_Unpack_Supported
End Property

Public Property Get FileCount() As Long
    FileCount = PackTotFiles
End Property

Public Property Get FileName(FileNum As Long) As String
    If NotGood(FileNum) Then Exit Property
    FileName = ZIPFileData(FileNum).FileName
End Property

Public Property Get CommentsFile(FileNum As Long) As String
    If NotGood(FileNum) Then Exit Property
    CommentsFile = ZIPFileData(FileNum).ComField
End Property

Public Property Get CommentsPack() As String
    CommentsPack = PackComments
End Property

Public Property Get IsDir(FileNum As Long) As Boolean
    If NotGood(FileNum) Then Exit Property
    If ZIPFileData(FileNum).CSize = 0 Then
        If Right(ZIPFileData(FileNum).FileName, 1) = "/" Then IsDir = True
    End If
End Property

Public Property Get Method(FileNum As Long) As String
    If NotGood(FileNum) Then Exit Property
    Method = Methods(ZIPFileData(FileNum).Method)
End Property

Public Property Get CRC32(FileNum As Long) As Long
    If NotGood(FileNum) Then Exit Property
    CRC32 = ZIPFileData(FileNum).CRC32
End Property

Public Property Get Compressed_Size(FileNum As Long) As Long
    If NotGood(FileNum) Then Exit Property
    Compressed_Size = ZIPFileData(FileNum).CSize
End Property

Public Property Get UnCompressed_Size(FileNum As Long) As Long
    If NotGood(FileNum) Then Exit Property
    UnCompressed_Size = ZIPFileData(FileNum).USize
End Property

Public Property Get Encrypted(FileNum As Long) As Boolean
    If NotGood(FileNum) Then Exit Property
    Encrypted = (ZIPFileData(FileNum).Flags And 1) = 1
End Property

Public Property Get FileDateTime(FileNum As Long) As Date
    If NotGood(FileNum) Then Exit Property
    FileDateTime = GetZipDate(ZIPFileData(FileNum).FDate, ZIPFileData(FileNum).FTime)
End Property

Public Property Get SystemMadeBy(FileNum As Long) As String
    If NotGood(FileNum) Then Exit Property
    SystemMadeBy = SystemName(Fix(ZIPFileData(FileNum).VerMade / 256) And 255)
End Property

Public Property Get VersionMadeBy(FileNum As Long) As String
    If NotGood(FileNum) Then Exit Property
    VersionMadeBy = VersionTo(ZIPFileData(FileNum).VerMade And 255)
End Property

Public Property Get SystemNeeded(FileNum As Long) As String
    If NotGood(FileNum) Then Exit Property
    SystemNeeded = SystemName(Fix(ZIPFileData(FileNum).VerExt / 256) And 255)
End Property

Public Property Get VersionNeeded(FileNum As Long) As String
    If NotGood(FileNum) Then Exit Property
    VersionNeeded = VersionTo(ZIPFileData(FileNum).VerExt And 255)
End Property

Private Function NotGood(FileNum As Long) As Boolean
    If FileNum = 0 Then NotGood = True: Exit Function
    If FileNum > PackTotFiles Then NotGood = True: Exit Function
    If PackFileType = 0 Then NotGood = True: Exit Function
End Function

Private Function Int2Lng(Value As Integer) As Long
     If Value < 0 Then Int2Lng = &HFFFF& + Value + 1 Else Int2Lng = Value
End Function

Private Sub Write_Uncompressed_Data(FileNum As Long, ToPath As String)
    Dim DSize As Long
    DSize = DataSize
    If DSize > 0 Then
        If ZIPFileData(FileNum).CRC32 <> CRC.CalcCRC32File(PackData) Then 'calcCRC32(PackData) Then
            MsgBox "CRC calculation failed"
        End If
    Else
        If ZIPFileData(FileNum).CRC32 <> 0 Then
            MsgBox "CRC error"
        End If
    End If
    If ZIPFileData(FileNum).USize <> DSize Then
        MsgBox "Error in decompressed size"
    End If
    If Write_File(ZIPFileData(FileNum).FileName, ToPath, PackData(), ZIPFileData(FileNum).FDate, ZIPFileData(FileNum).FTime) <> 0 Then
        MsgBox "error writing file"
    End If
End Sub

Private Function DataSize() As Long
    On Error Resume Next
    DataSize = UBound(PackData) + 1
    If Err.Number <> 0 Then
        Err.Clear
        DataSize = 0
    End If
End Function

Private Function SystemName(System As Byte) As String
    Select Case System
    Case 0:     SystemName = "MS-DOS and OS/2 (FAT / VFAT / FAT32 file systems)"
    Case 1:     SystemName = "Amiga"
    Case 2:     SystemName = "2 - OpenVMS"
    Case 3:     SystemName = "UNIX"
    Case 4:     SystemName = "VM/CMS"
    Case 5:     SystemName = "Atari ST"
    Case 6:     SystemName = "OS/2 H.P.F.S."
    Case 7:     SystemName = "Macintosh"
    Case 8:     SystemName = "Z-System"
    Case 9:     SystemName = "CP/M"
    Case 10:    SystemName = "Windows NTFS"
    Case 11:    SystemName = "MVS (OS/390 - Z/OS)"
    Case 12:    SystemName = "VSE"
    Case 13:    SystemName = "Acorn Risc"
    Case 14:    SystemName = "VFAT"
    Case 15:    SystemName = "Alternate MVS"
    Case 16:    SystemName = "BeOS"
    Case 17:    SystemName = "Tandem"
    Case 18:    SystemName = "OS/400"
    Case Else:  SystemName = "UnKnown"
    End Select
End Function

Private Function VersionTo(Version As Byte) As String
    VersionTo = Fix(Version / 10) & "." & Version Mod 10
End Function

Private Function Methods(MethodType As Integer) As String
    Select Case MethodType
        Case 0: Methods = "No Compression"
        Case 1: Methods = "Shrunk"
        Case 2: Methods = "Reduce Factor 1"
        Case 3: Methods = "Reduce Factor 2"
        Case 4: Methods = "Reduce Factor 3"
        Case 5: Methods = "Reduce Factor 4"
        Case 6: Methods = "Imploded"
        Case 7: Methods = "Tokenized"
        Case 8: Methods = "Deflated"
        Case 9: Methods = "Enhanced Deflating"
        Case 10: Methods = "PKWARE Imploding"
        Case 11: Methods = "Reserved"
        Case 12: Methods = "BZip2"
        Case Else: Methods = "Unknown"
    End Select
End Function

Public Function Inflate(ByteArray() As Byte, UncompressedSize As Long, Optional ZIP64 As Boolean = False) As Long
'On Error GoTo errhandle

    Dim IsLastBlock  As Boolean
    Dim CompType     As Long
    Dim Char         As Long
    Dim Nubits       As Long
    Dim L1           As Long
    Dim L2           As Long
    Dim X            As Long


    InStream = ByteArray 'Copy local array to global array
    Call Init_Inflate(UncompressedSize) 'Init global variables

    Do
        IsLastBlock = GetBits(1) 'Read Last Block Flag
        CompType = GetBits(2)    'Read Block Type

        If CompType = 0 Then              'Block is Stored
            If Inpos + 4 > UBound(InStream) Then
                Inflate = -1 'InStream depleated
                Exit Do
            End If
'this is done couse if bitnum >= then next byte is already in ByteBuff
            Do While BitNum >= 8
                Inpos = Inpos - 1
                BitNum = BitNum - 8
            Loop
            CopyMemory L1, InStream(Inpos), 2&      'Read Count
            CopyMemory L2, InStream(Inpos + 2), 2&  'Read ones compliment of Count
            Inpos = Inpos + 4
            If L1 - (Not (L2) And &HFFFF&) Then Inflate = -2
            If Inpos + L1 - 1 > UBound(InStream) Then
                Inflate = -1 'InStream depleated
                Exit Do
            End If
            If OutPos + L1 - 1 > UBound(OutStream) Then
                Inflate = -1 'OutStream overflow
                Exit Do
            End If
            
           
            
            CopyMemory OutStream(OutPos), InStream(Inpos), L1 'Copy stored Block
            OutPos = OutPos + L1
            Inpos = Inpos + L1
            ByteBuff = 0
            BitNum = 0
        ElseIf CompType = 3 Then          'Error in compressed data
            Inflate = -1
            Exit Do
        Else
            If CompType = 1 Then          'Static Compression
                If Create_Static_Tree <> 0 Then
                    MsgBox "Error in tree creation (Static)"
                    Exit Function
                End If
            Else 'CompType = 2            'Dynamic Compression
                If Create_Dynamic_Tree <> 0 Then
                    MsgBox "Error in tree creation (Static)"
                    Exit Function
                End If
            End If
            Do
                NeedBits MaxLLenght
                Nubits = MinLLenght
                Do While LitLen.Lenght(ByteBuff And BitMask(Nubits)) <> Nubits
                    Nubits = Nubits + 1
                Loop
                Char = LitLen.code(ByteBuff And BitMask(Nubits))
                DropBits Nubits
                
                DoEvents
                If gbZipStop Then Exit Function
'                Debug.Print "Char", Char
                 If OutPos Mod 10000 = 0 Or OutPos - 1 = UncompressedSize Then
'                    Debug.Assert OutPos - 1 <> UncompressedSize
                    RaiseEvent Persent(CInt(OutPos / (UncompressedSize + 1) * 100))
                    
'                    Debug.Print CInt(OutPos / UncompressedSize * 100)
                 End If
                 
                If Char < 256 Then  'Character is Literal
                    OutStream(OutPos) = Char        'output the character
                    OutPos = OutPos + 1
                ElseIf Char > 256 Then 'Character is Length Symbol
                    'Decode Length L1
                    Char = Char - 257
                    L1 = LC.code(Char) + GetBits(LC.Lenght(Char))
                    If (L1 = 258) And ZIP64 Then L1 = GetBits(16) + 3
                    'Decode Distance L2 Symbol
                    NeedBits MaxDLenght
                    Nubits = MinDLenght
                    Do While Dist.Lenght(ByteBuff And BitMask(Nubits)) <> Nubits
                        Nubits = Nubits + 1
                    Loop
                    Char = Dist.code(ByteBuff And BitMask(Nubits))
                    DropBits Nubits
                    L2 = dc.code(Char) + GetBits(dc.Lenght(Char)) 'Decode Distance L2
                    'copy L2 positions back L1 characters
                    For X = 1 To L1
                        OutStream(OutPos) = OutStream(OutPos - L2)
                        OutPos = OutPos + 1
                    Next X
                End If
            Loop While Char <> 256 'EOB
        End If
    Loop While Not IsLastBlock
Stop_Decompression:
    If OutPos > 0 Then
        ReDim Preserve OutStream(OutPos - 1)
    Else
        Erase OutStream
    End If
'Clear memory
    Erase InStream
    Erase BitMask
    Erase Pow2
    Erase LC.code
    Erase LC.Lenght
    Erase dc.code
    Erase dc.Lenght
    Erase LitLen.code
    Erase LitLen.Lenght
    Erase Dist.code
    Erase Dist.Lenght
    Erase LenOrder
    ByteArray = OutStream
    
    Exit Function
errhandle:
If OutPos > UBound(OutStream) Then
    MsgBox "Incorrect Uncompressed Size"
    GoTo Stop_Decompression
ElseIf Inpos > UBound(InStream) Then
    MsgBox "Unexpected End of File"
    GoTo Stop_Decompression
Else
    Err.Raise Err.Number, , Err.Description
End If

End Function

