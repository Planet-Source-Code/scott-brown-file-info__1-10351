VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'Windows API/Global Declarations for: Obtaining Information Of A File
'***************************************************************
'Reviewed and used from PSC: 15-10-1999. If this is your code...apologies for no credit.
'Let me know your post and I'll amend !

Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function VerLanguageName Lib "kernel32" Alias "VerLanguageNameA" (ByVal wLang As Long, ByVal szLang As String, ByVal nSize As Long) As Long
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Long, puLen As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800

Private Type FILETIME
    LowDateTime As Long
    HighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 260 'MUST be set to 260
    cAlternate As String * 14
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
    
Private mvarCompanyName As String
Private mvarFileDescription As String
Private mvarFileVersion As String
Private mvarInternalName As String
Private mvarLegalCopyright As String
Private mvarOriginalFileName As String
Private mvarProductName As String
Private mvarProductVersion As String
Private mvarFileName As String
Private mvarByte As String
Private mvarCreationTime As String
Private mvarLastAccessTime As String
Private mvarLastWriteTime As String
Private mvarReadOnly As Boolean
Private mvarHidden As Boolean
Private mvarSystem As Boolean
Private mvarDirectory As Boolean
Private mvarArchive As Boolean
Private mvarNormal As Boolean
Private mvarTemporary As Boolean
Private mvarCompress As Boolean
                
Public Property Get Compress() As Boolean
    Compress = mvarCompress
End Property

Public Property Get Temporary() As Boolean
    Temporary = mvarTemporary
End Property

Public Property Get Normal() As Boolean
    Normal = mvarNormal
End Property

Public Property Get Archive() As Boolean
    Archive = mvarArchive
End Property

Public Property Get Directory() As Boolean
    Directory = mvarDirectory
End Property

Public Property Get System() As Boolean
    System = mvarSystem
End Property

Public Property Get Hidden() As Boolean
    Hidden = mvarHidden
End Property

Public Property Get ReadOnly() As Boolean
    ReadOnly = mvarReadOnly
End Property

Public Property Get LastWriteTime() As String
    LastWriteTime = mvarLastWriteTime
End Property

Public Property Get LastAccessTime() As String
    LastAccessTime = mvarLastAccessTime
End Property

Public Property Get CreationTime() As String
    CreationTime = mvarCreationTime
End Property

Public Property Get mByte() As String
    mByte = mvarByte
End Property

Public Property Get ProductVersion() As String
    ProductVersion = mvarProductVersion
End Property

Public Property Get ProductName() As String
    ProductName = mvarProductName
End Property

Public Property Get OriginalFileName() As String
    OriginalFileName = mvarOriginalFileName
End Property

Public Property Get LegalCopyright() As String
    LegalCopyright = mvarLegalCopyright
End Property

Public Property Get InternalName() As String
    InternalName = mvarInternalName
End Property

Public Property Get FileVersion() As String
    FileVersion = mvarFileVersion
End Property

Public Property Get FileDescription() As String
   FileDescription = mvarFileDescription
End Property

Public Property Get CompanyName() As String
    CompanyName = mvarCompanyName
End Property

Public Property Get FileName() As String
    FileName = mvarFileName
End Property

Public Function FindFileInfo(strFileName As String) As Long

    On Error GoTo GetFileVersionData_Error
    Dim sInfo As String, lSizeof As Long
    Dim lResult As Long, iDelim As Integer, n As Integer, lHandle As Long
    Dim ftime As SYSTEMTIME
    Dim filedata As WIN32_FIND_DATA

    If strFileName <> "" Then
        filedata = Findfile(strFileName)
        Call FileTimeToSystemTime(filedata.ftCreationTime, ftime)
        mvarCreationTime = ftime.wDay & "/" & ftime.wMonth & "/" & ftime.wYear & " " & ftime.wHour & ":" & ftime.wMinute & ":" & ftime.wSecond
        Call FileTimeToSystemTime(filedata.ftLastWriteTime, ftime) ' Determine Last Modified Date and time
        mvarLastWriteTime = ftime.wDay & "/" & ftime.wMonth & "/" & ftime.wYear & " " & ftime.wHour & ":" & ftime.wMinute & ":" & ftime.wSecond
        Call FileTimeToSystemTime(filedata.ftLastAccessTime, ftime) ' Determine Last accessed Date (note no time is recorded)
        mvarLastAccessTime = ftime.wDay & "/" & ftime.wMonth & "/" & ftime.wYear
        mvarHidden = ((filedata.dwFileAttributes And FILE_ATTRIBUTE_HIDDEN) = FILE_ATTRIBUTE_HIDDEN)
        mvarSystem = ((filedata.dwFileAttributes And FILE_ATTRIBUTE_SYSTEM) = FILE_ATTRIBUTE_SYSTEM)
        mvarReadOnly = ((filedata.dwFileAttributes And FILE_ATTRIBUTE_READONLY) = FILE_ATTRIBUTE_READONLY)
        mvarArchive = ((filedata.dwFileAttributes And FILE_ATTRIBUTE_ARCHIVE) = FILE_ATTRIBUTE_ARCHIVE)
        mvarTemporary = ((filedata.dwFileAttributes And FILE_ATTRIBUTE_TEMPORARY) = FILE_ATTRIBUTE_TEMPORARY)
        mvarNormal = ((filedata.dwFileAttributes And FILE_ATTRIBUTE_NORMAL) = FILE_ATTRIBUTE_NORMAL)
        mvarCompress = ((filedata.dwFileAttributes And FILE_ATTRIBUTE_COMPRESSED) = FILE_ATTRIBUTE_COMPRESSED)
        mvarFileName = UCase$(strFileName)


        If filedata.nFileSizeHigh = 0 Then
            mvarByte = Format$(filedata.nFileSizeLow, "###,###,###") & " bytes"
        Else
            mvarByte = Format$(filedata.nFileSizeHigh, "###,###,###") & " bytes"
        End If

        lHandle = 0
        lSizeof = GetFileVersionInfoSize(strFileName, lHandle)

        If lSizeof > 0 Then
            sInfo = String$(lSizeof, 0)
            lResult = GetFileVersionInfo(ByVal strFileName, 0&, ByVal lSizeof, ByVal sInfo)

            If lResult Then
                iDelim = InStr(sInfo, "CompanyName")

                If iDelim > 0 Then
                    iDelim = iDelim + 12
                    mvarCompanyName = Mid$(sInfo, iDelim)
                End If

                iDelim = InStr(sInfo, "FileDescription")

                If iDelim > 0 Then
                    iDelim = iDelim + 16
                    mvarFileDescription = Mid$(sInfo, iDelim)
                End If

                iDelim = InStr(sInfo, "FileVersion")

                If iDelim > 0 Then
                    iDelim = iDelim + 12
                    mvarFileVersion = Mid$(sInfo, iDelim)
                End If

                iDelim = InStr(sInfo, "InternalName")

                If iDelim > 0 Then
                    iDelim = iDelim + 16
                    mvarInternalName = Mid$(sInfo, iDelim)
                End If

                iDelim = InStr(sInfo, "LegalCopyright")

                If iDelim > 0 Then
                    iDelim = iDelim + 16
                    mvarLegalCopyright = Mid$(sInfo, iDelim)
                End If

                iDelim = InStr(sInfo, "OriginalFilename")

                If iDelim > 0 Then
                    iDelim = iDelim + 20
                    mvarOriginalFileName = Mid$(sInfo, iDelim)
                End If

                iDelim = InStr(sInfo, "ProductName")

                If iDelim > 0 Then
                    iDelim = iDelim + 12
                    mvarProductName = Mid$(sInfo, iDelim)
                End If

                iDelim = InStr(sInfo, "ProductVersion")

                If iDelim > 0 Then
                    iDelim = iDelim + 16
                    mvarProductVersion = Mid$(sInfo, iDelim)
                End If

                FindFileInfo = 1
            Else
                GoTo invalid_file_info_error
            End If

        Else
            GoTo invalid_file_info_error
        End If

    Else
        FindFileInfo = 0
    End If

GetFileVersionData_Exit:
    Exit Function
GetFileVersionData_Error:
    FindFileInfo = 0
    Resume GetFileVersionData_Exit
invalid_file_info_error:
    FindFileInfo = 1
    GoTo GetFileVersionData_Exit
End Function

Private Function Findfile(xStrfilename) As WIN32_FIND_DATA
    Dim Win32Data As WIN32_FIND_DATA
    Dim plngFirstFileHwnd As Long
    Dim plngRtn As Long
    
    plngFirstFileHwnd = FindFirstFile(xStrfilename, Win32Data) ' Get information of file using API call

    If plngFirstFileHwnd = 0 Then
        Findfile.cFileName = "Error" ' If file was not found, return error as name
    Else
        Findfile = Win32Data ' Else return results
    End If

    plngRtn = FindClose(plngFirstFileHwnd) ' It is important that you close the handle For FindFirstFile

End Function




