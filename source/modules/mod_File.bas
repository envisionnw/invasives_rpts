Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_File
' Level:        Framework module
' Version:      1.00
' Description:  File and directory related functions & subroutines
'
' Source/date:  Bonnie Campbell, April 2015
' Revisions:    BLC, 4/30/2015 - initial version
' =================================

' ---------------------------------
'  DIRECTORY RELATED
' ---------------------------------
' =================================
' FUNCTION:     fxnCreateFolder
' Description:  Creates a folder with the specified path
' Parameters:   strPath as a string
' Returns:      True or False
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 1/9/2009
' Revisions:    JRB, 1/9/2009 - initial version
'               BLC, 4/30/2015 - moved from mod_Utilities to mod_File
' =================================
Public Function fxnCreateFolder(ByVal strPath As String) As Boolean
    On Error GoTo Err_Handler

    fxnCreateFolder = False    ' Default in case of error

    Dim fs As Variant

    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FolderExists(strPath) = False Then
        fs.CreateFolder (strPath)
        fxnCreateFolder = True
    End If

Exit_Function:
    On Error Resume Next
    Set fs = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnCreateFolder[mod_File])"
    End Select
    Resume Exit_Function
End Function

' =================================
' FUNCTION:     fxnFolderExists
' Description:  Indicates whether or not the indicated folder exists
' Parameters:   strPath as a string
' Returns:      True or False
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 1/9/2009
' Revisions:    JRB, 1/9/2009 - initial version
'               BLC, 4/30/2015 - moved from mod_Utilities to mod_File
' =================================
Public Function fxnFolderExists(ByVal strPath As String) As Boolean
    On Error GoTo Err_Handler

    fxnFolderExists = False    ' Default in case of error

    Dim fs As Variant

    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FolderExists(strPath) Then fxnFolderExists = True

Exit_Function:
    On Error Resume Next
    Set fs = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnFolderExists[mod_File])"
    End Select
    Resume Exit_Function
End Function


' ---------------------------------
'  FILE RELATED
' ---------------------------------

' =================================
' FUNCTION:     fxnGetFile
' Description:  Opens the open/save file dialog and returns the file name selected by the user
' Parameters:   strInitialDir - the directory to start searching in (optional)
'               strFileType, varFileExt - file type and extension (optional)
'               strTitle - title of the dialog box (optional)
' Returns:      name of the file to open/import; or Null if user cancels
' Throws:       none
' References:   adhAddFilterItem, adhCommonFileOpenSave
' Source/date:  Susan Huse, fall 2004
' Revisions:    John R. Boetsch, May 17, 2006 - updated documentation and error trap
'               JRB, 6/22/2009 - revised from fxnGetLinkFile; added file type/ext variables
'               BLC, 4/30/2015 - moved from mod_Utilities to mod_File
' =================================
Public Function fxnGetFile(Optional ByVal strInitialDir As String, _
    Optional ByVal strFileType As String, _
    Optional ByVal varFileExt As Variant, _
    Optional ByVal strTitle As String = "Select File to Open") As Variant

    On Error GoTo Err_Handler

    Dim strFilter As String
    Dim lngFlags As Long

    ' Use the open file dialog to interactively browse to and select the desired file
    strFilter = adhAddFilterItem(strFilter, strFileType, varFileExt)

    lngFlags = adhOFN_HIDEREADONLY Or _
        adhOFN_HIDEREADONLY Or adhOFN_NOCHANGEDIR

    fxnGetFile = adhCommonFileOpenSave( _
        InitialDir:=strInitialDir, _
        OpenFile:=True, _
        Filter:=strFilter, _
        flags:=lngFlags, _
        DialogTitle:=strTitle)

Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnGetFile[mod_File])"
    End Select
    Resume Exit_Function
End Function

' =================================
' FUNCTION:     fxnSaveFile
' Description:  Opens the open/save file dialog and returns the file name selected by the user
' Parameters:   strFileName, strFileType, strFileExt - file name/path, type and extension
'               strTitle - title of the dialog box (optional)
' Returns:      name of the file to save; or Null if user cancels
' Throws:       none
' References:   adhAddFilterItem, adhCommonFileOpenSave
' Source/date:  Susan Huse, fall 2004
' Revisions:    John R. Boetsch, May 2005 - minor revisions and documentation
' Revisions:    JRB, 5/16/2006 - updated documentation, error traps
'               JRB, 6/22/2009 - added strTitle to parameters
'               BLC, 4/30/2015 - move from mod_Utilities to mod_File
' =================================
Public Function fxnSaveFile(ByVal strFilename As String, ByVal strFileType As String, _
    ByVal strFileExt As String, Optional ByVal strTitle As String = "Save As") As Variant

    On Error GoTo Err_Handler

    Dim strFilter As String
    Dim lngFlags As Long

    ' Use the save file dialog to interactively browse to and select the desired file
    strFilter = adhAddFilterItem(strFilter, strFileType, strFileExt)

    lngFlags = adhOFN_HIDEREADONLY Or adhOFN_OVERWRITEPROMPT Or _
        adhOFN_HIDEREADONLY Or adhOFN_NOCHANGEDIR

    fxnSaveFile = adhCommonFileOpenSave( _
        OpenFile:=False, _
        Filter:=strFilter, _
        flags:=lngFlags, _
        DialogTitle:=strTitle, _
        fileName:=strFilename)

Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnSaveFile[mod_File])"
    End Select
    Resume Exit_Function
End Function

' =================================
' FUNCTION:     fxnFileExists
' Description:  Indicates whether or not the indicated file exists
' Parameters:   strPath as a string
' Returns:      True or False
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 5/8/2006
' Revisions:    JRB, 5/8/2006 - initial version
'               BLC, 4/30/2015 - moved from mod_Utilities to mod_File
' =================================
Public Function fxnFileExists(ByVal strPath As String) As Boolean
    On Error GoTo Err_Handler

    fxnFileExists = False    ' Default in case of error

    Dim fs As Variant

    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(strPath) Then fxnFileExists = True

Exit_Function:
    On Error Resume Next
    Set fs = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnFileExists[mod_File])"
    End Select
    Resume Exit_Function
End Function

' =================================
' FUNCTION:     fxnDeleteFile
' Description:  Deletes the specified file; this is preferred over the Kill command
'               because it works for hidden files and read-only files
' Parameters:   strPath - the path and file name to be deleted
' Returns:      True if deleted, or False if error
' Throws:       none
' References:   fxnFileExists
' Source/date:  John R. Boetsch, 5/19/2006
' Revisions:    JRB, 5/19/2006 - initial version
'               BLC, 4/30/2015 - moved from mod_Utilities to mod_File
' =================================
Public Function fxnDeleteFile(ByVal strPath As String) As Boolean
    On Error GoTo Err_Handler

    fxnDeleteFile = False    ' Default in case of error

    Dim fs As Variant

    Set fs = CreateObject("Scripting.FileSystemObject")
    If fxnFileExists(strPath) Then
        fs.DeleteFile strPath, True
        fxnDeleteFile = True
    Else
        MsgBox "Unable to delete the specified file", vbCritical, _
            "File delete error (fxnDeleteFile)"
    End If

Exit_Function:
    On Error Resume Next
    Set fs = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnDeleteFile[mod_File])"
    End Select
    Resume Exit_Function
End Function

' =================================
' FUNCTION:     fxnParseFileName
' Description:  Parses an input path string to return only the file extension, if present
' Parameters:   strFullPath - string for the full file path
' Returns:      string including only the file name
' Throws:       none
' References:   none
' Source/date:  From Front-end Application Builder v1.1, Simon Kingston, date unknown
' Revisions:    John R. Boetsch, 6/17/2009 - error trapping, documentation
'               BLC, 4/30/2015 - moved from mod_Utilities to mod_File
' =================================
Public Function fxnParseFileName(ByVal strFullPath As String) As String
    On Error GoTo Err_Handler

    Dim strTemp As String

    Do While (InStr(strFullPath, "\") > 0)
        strTemp = strTemp & left(strFullPath, InStr(strFullPath, "\"))
        strFullPath = Mid(strFullPath, InStr(strFullPath, "\") + 1)
    Loop
    
    fxnParseFileName = strFullPath

Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnParseFileName[mod_File])"
    End Select
    Resume Exit_Function
End Function

' =================================
' FUNCTION:     fxnParseFileExt
' Description:  Parses an input path string to return only the file extension, if present
' Parameters:   strFullPath - string for the full file path
'               blnIncludeDot - flag to include the dot (".") in the return (default is True)
' Returns:      string including only the file extension, or an empty string ("") if missing
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 6/22/2009
' Revisions:    JRB, 6/22/2009 - initial version
'               BLC, 4/30/2015 - moved from mod_Utilities to mod_File
' =================================
Public Function fxnParseFileExt(ByVal strFullPath As String, _
    Optional blnIncludeDot As Boolean = True) As String

    On Error GoTo Exit_Procedure

    Dim arrPath() As String
    Dim strFile As String
    Dim strTemp As String
    Dim varPosition As Variant

    ' Split into an array based on the "\" delimiter; file name should be the uppermost segment
    arrPath = Split(strFullPath, "\")
    strFile = arrPath(UBound(arrPath))

    ' Get the position in the string of the dot
    varPosition = InStr(1, strFile, ".")
    If varPosition > 0 Then
        If blnIncludeDot = False Then varPosition = varPosition + 1
        strTemp = Mid(strFile, varPosition)
    Else
        strTemp = ""
    End If

    fxnParseFileExt = strTemp

Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnParseFileExt[mod_File])"
    End Select
    Resume Exit_Function
End Function

' =================================
' FUNCTION:     fxnOpenExcelFile
' Description:  Opens file in Excel - assumes that the file exists and can be opened by Excel
' Parameters:   strPath - full path of the file to be opened
' Returns:      none
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 6/22/2009
' Revisions:    JRB, 3/7/12 - fixed function header to indicate 'Public'
'               BLC, 4/30/2015 - moved from mod_Utilities to mod_File
' =================================
Public Function fxnOpenExcelFile(ByVal strPath As String)
    On Error GoTo Err_Handler

    Dim objExcel As Object

    ' Create a new instance of Excel
    Set objExcel = CreateObject("Excel.Application")
    objExcel.UserControl = True

    ' Open the file
    With objExcel
        .Visible = True
        .Workbooks.Open (strPath)
    End With
    
Exit_Function:
    On Error Resume Next
    Set objExcel = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnOpenExcelFile[mod_File])"
    End Select
    Resume Exit_Function
End Function

' =================================
' FUNCTION:     fxnParsePath
' Description:  Parses an input path string to return only the path without the file name
' Parameters:   strFullPath - string for the full file path
' Returns:      string including only the file path, or an empty string ("") if missing
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 6/22/2009
' Revisions:    JRB, 6/22/2009 - initial version
'               BLC, 4/30/2015 - moved from mod_Utilities to mod_File
' =================================
Public Function fxnParsePath(ByVal strFullPath As String) As String
    On Error GoTo Exit_Procedure

    Dim arrPath() As String
    Dim strFile As String

    ' Split into an array based on the "\" delimiter; file name should be the uppermost segment
    arrPath = Split(strFullPath, "\")
    strFile = arrPath(UBound(arrPath))

    ' Path is the full string minus length of the file name
    fxnParsePath = left(strFullPath, Len(strFullPath) - Len(arrPath(UBound(arrPath))))

Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnParsePath[mod_File])"
    End Select
    Resume Exit_Function
End Function