' =================================
' MODULE:       mod_Utilities
' Description:  Standard module containing useful or commonly-used generic functions
' Source/date:  John R. Boetsch, 5/17/2006
' Revisions:    JRB, 6/25/2009 - added fxnFiscalYear, fxnFormIsLoaded, fxnHideObject,
'                   fxnGetFile, fxnParseFileName, fxnParseFileExt, fxnParsePath; reorganized
'                   to include: fxnFolderExists, fxnCreateFolder, fxnSaveFile, fxnFileExists,
'                   fxnDeleteFile; removed fxnTrimSpaces and replaced with Trim()
'               JRB, 12/31/2009 - added fxnUserName

Option Compare Database
Option Explicit

' =================================
' FUNCTION:     fxnFiscalYear
' Description:  Returns the fiscal year corresponding to the input date
' Parameters:   datDate - date value to be converted to fiscal year
'               blnFourDigits - flag to use 4 digits to represent the year (default True)
'               blnAddPrefix - flag to add a prefix to the result (default True)
'               strPrefix - prefix to be added to the string
' Returns:      variant for the fiscal year string or integer (e.g., "FY2010")
' Throws:       none
' References:   none
' Source/date:  From Front-end Application Builder v1.1, Simon Kingston, date unknown
' Revisions:    John R. Boetsch, 6/17/2009 - error trapping, documentation, added prefix & digit flags
' =================================
Public Function fxnFiscalYear(ByVal datDate As Date, _
    Optional ByVal blnFourDigits As Boolean = True, _
    Optional ByVal blnAddPrefix As Boolean = True, _
    Optional ByVal strPrefix As String = "FY") As Variant

    On Error GoTo Err_Handler

    Dim intYear As Integer
    Dim strYear As String

    intYear = Year(datDate)
    If Month(datDate) >= 10 Then intYear = intYear + 1

    ' Year string depending on 2 or 4 characters
    If blnFourDigits Then
        strYear = CStr(intYear)
    Else
        strYear = Right(CStr(intYear), 2)
    End If

    If blnAddPrefix Then
        fxnFiscalYear = strPrefix & strYear
    Else
        fxnFiscalYear = strYear
    End If

Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnFiscalYear)"
    End Select
    Resume Exit_Procedure

End Function

' =================================
' FUNCTION:     fxnFormIsLoaded
' Description:  Returns whether the specified form is loaded in Form or Datasheet view
' Parameters:   strFormName - string for the name of the form to check
' Returns:      True if the specified form is open in Form view or Datasheet view
' Throws:       none
' References:   none
' Source/date:  From Northwind sample database, date unknown
' Revisions:    John R. Boetsch, 6/17/2009 - error trapping, documentation
' =================================
Public Function fxnFormIsLoaded(ByVal strFormName As String) As Integer
    On Error GoTo Err_Handler
 
    ' These variables are used to test the return values of the SysCmd function
    '  and the CurrentView property of the requested form.
    Const conObjStateClosed = 0
    Const conDesignView = 0

    ' Use the SysCmd function to check the current state of the requested form.
    '  Possible states: not open or nonexistent, open, new, or changed but not saved
    If SysCmd(acSysCmdGetObjectState, acForm, strFormName) <> conObjStateClosed Then
        ' Checks for the current view of the requested form, assuming the previous statement
        '   found it to be open ... return True if open and not in design view
        If Forms(strFormName).CurrentView <> conDesignView Then
            fxnFormIsLoaded = True
        End If
    End If
    
Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnFormIsLoaded)"
    End Select
    Resume Exit_Procedure

End Function

' =================================
' FUNCTION:     fxnHideObject
' Description:  Changes the hidden property of an object to hide / show in the database window
' Parameters:   strObjectName - name of the object (string)
'               blnHide - True to hide, False to show (default True)
'               varType - object type (default acTable)
' Returns:      none
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 6/25/2009
' Revisions:    <name, date, desc>
' =================================
Public Function fxnHideObject(strObjectName As String, _
    Optional blnHide As Boolean = True, Optional varType As Variant = acTable)

    On Error GoTo Err_Handler

    SetHiddenAttribute varType, strObjectName, blnHide

Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnHideObject)"
    End Select
    Resume Exit_Procedure

End Function

' =================================
' FUNCTION:     fxnTableExists
' Description:  Returns whether the specified table exists in the current database collection
' Parameters:   strTableName - string for the name of the table to check
' Returns:      True if the specified table exists in the master systems table, or False
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 6/29/2009
' Revisions:    <name, date, desc - add lines as you go>
' =================================
Public Function fxnTableExists(ByVal strTableName As String) As Boolean
    On Error GoTo Err_Handler

    fxnTableExists = DCount("*", "MSysObjects", "(([Type] In (1,4,6)) AND ([Name]=""" & _
        strTableName & """))")

Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnTableExists)"
    End Select
    Resume Exit_Procedure

End Function

' =================================
' FUNCTION:     fxnFolderExists
' Description:  Indicates whether or not the indicated folder exists
' Parameters:   strPath as a string
' Returns:      True or False
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 1/9/2009
' Revisions:    <name, date, desc - add lines as you go>
' =================================
Public Function fxnFolderExists(ByVal strPath As String) As Boolean
    On Error GoTo Err_Handler

    fxnFolderExists = False    ' Default in case of error

    Dim fs As Variant

    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FolderExists(strPath) Then fxnFolderExists = True

Exit_Procedure:
    On Error Resume Next
    Set fs = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnFolderExists)"
    End Select
    Resume Exit_Procedure

End Function

' =================================
' FUNCTION:     fxnCreateFolder
' Description:  Creates a folder with the specified path
' Parameters:   strPath as a string
' Returns:      True or False
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 1/9/2009
' Revisions:    <name, date, desc - add lines as you go>
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

Exit_Procedure:
    On Error Resume Next
    Set fs = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnCreateFolder)"
    End Select
    Resume Exit_Procedure

End Function

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

Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnGetFile)"
    End Select
    Resume Exit_Procedure
    
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

Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnSaveFile)"
    End Select
    Resume Exit_Procedure
    
End Function

' =================================
' FUNCTION:     fxnFileExists
' Description:  Indicates whether or not the indicated file exists
' Parameters:   strPath as a string
' Returns:      True or False
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 5/8/2006
' Revisions:    <name, date, desc - add lines as you go>
' =================================
Public Function fxnFileExists(ByVal strPath As String) As Boolean
    On Error GoTo Err_Handler

    fxnFileExists = False    ' Default in case of error

    Dim fs As Variant

    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(strPath) Then fxnFileExists = True

Exit_Procedure:
    On Error Resume Next
    Set fs = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnFileExists)"
    End Select
    Resume Exit_Procedure

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
' Revisions:    <name, date, desc - add lines as you go>
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

Exit_Procedure:
    On Error Resume Next
    Set fs = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnDeleteFile)"
    End Select
    Resume Exit_Procedure

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
' =================================
Public Function fxnParseFileName(ByVal strFullPath As String) As String
    On Error GoTo Err_Handler

    Dim strTemp As String

    Do While (InStr(strFullPath, "\") > 0)
        strTemp = strTemp & Left(strFullPath, InStr(strFullPath, "\"))
        strFullPath = Mid(strFullPath, InStr(strFullPath, "\") + 1)
    Loop
    
    fxnParseFileName = strFullPath

Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnParseFileName)"
    End Select
    Resume Exit_Procedure

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
' Revisions:    <name, date, desc - add lines as you go>
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

Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnParseFileExt)"
    End Select
    Resume Exit_Procedure

End Function

' =================================
' FUNCTION:     fxnParsePath
' Description:  Parses an input path string to return only the path without the file name
' Parameters:   strFullPath - string for the full file path
' Returns:      string including only the file path, or an empty string ("") if missing
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 6/22/2009
' Revisions:    <name, date, desc - add lines as you go>
' =================================
Public Function fxnParsePath(ByVal strFullPath As String) As String
    On Error GoTo Exit_Procedure

    Dim arrPath() As String
    Dim strFile As String

    ' Split into an array based on the "\" delimiter; file name should be the uppermost segment
    arrPath = Split(strFullPath, "\")
    strFile = arrPath(UBound(arrPath))

    ' Path is the full string minus length of the file name
    fxnParsePath = Left(strFullPath, Len(strFullPath) - Len(arrPath(UBound(arrPath))))

Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnParsePath)"
    End Select
    Resume Exit_Procedure

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
    
Exit_Procedure:
    On Error Resume Next
    Set objExcel = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnOpenExcelFile)"
    End Select
    Resume Exit_Procedure

End Function

' =================================
' FUNCTION:     fxnReplaceString
' Description:  Replaces a substring in a string with another
' Parameters:   strTextIn - string to work on
'               strFind - string to find
'               strReplace - string to replace with
'               fCaseSensitive - True for case sensitive search (default=False)
' Returns:      modified string
' Throws:       none
' References:   none
' Source/date:  Simon Kingston, date unknown
' Revisions:    John R. Boetsch, 5/17/2006 - error trapping, documentation
' =================================
Public Function fxnReplaceString(strTextIn As String, strFind As String, _
    strReplace As String, Optional fCaseSensitive As Boolean = False) As String

    On Error GoTo Err_Handler

    Dim strTemp As String
    Dim intPos As Integer
    Dim intCaseSensitive As Integer

    ' Convert the case-sensitive boolean to the comparison constant (1=binary, 2=textual)
    intCaseSensitive = fCaseSensitive + 1

    strTemp = strTextIn
    intPos = InStr(1, strTemp, strFind, intCaseSensitive)

    Do While intPos > 0
        strTemp = Left$(strTemp, intPos - 1) & strReplace & Mid$(strTemp, intPos + Len(strFind))
        intPos = InStr(intPos + Len(strReplace), strTemp, strFind, intCaseSensitive)
    Loop

    fxnReplaceString = strTemp

Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnReplaceString)"
    End Select
    Resume Exit_Procedure

End Function

' =================================
' FUNCTION:     fxnChangeDelimiter
' Description:  Replaces delimiters in an input string; default is to change double-quotes
'               to single quotes
' Parameters:   strInputText - string to work on
'               strCurrDelimiter - current delimiter in the string (default: double-quote)
'               strNewDelimiter - desired replacement delimiter (default: single-quote)
' Returns:      modified string
' Throws:       none
' References:   fxnReplaceString
' Source/date:  John R. Boetsch, 5/17/2006
' Revisions:    <name, date, desc - add lines as you go>
' =================================
Public Function fxnChangeDelimiter(strInputText As String, _
    Optional strCurrDelimiter As String = """", _
    Optional strNewDelimiter As String = "'") As String

    On Error GoTo Err_Handler

    Dim strTemp As String
    
    ' Call the replace string function, specifying the delimiter and no case-sensitive search
    strTemp = fxnReplaceString(strInputText, strCurrDelimiter, strNewDelimiter)
    fxnChangeDelimiter = strTemp

Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnChangeDelimiter)"
    End Select
    Resume Exit_Procedure

End Function

' =================================
' FUNCTION:     fxnUserName
' Description:  Returns the current user name
' Parameters:   none
' Returns:      string of the user login
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 12/31/2009
' Revisions:    <name, date - desc>
' =================================
Public Function fxnUserName() As String
    On Error GoTo Err_Handler

    fxnUserName = "Unknown"
    fxnUserName = Environ("Username")

Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnUserName)"
    End Select
    Resume Exit_Procedure

End Function

' ---------------------------------
' FUNCTION:     InsertSpace
' Description:  Inserts a space between capitalized letters
' Parameters:   str - string to inspect
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  theDBguy, May 20, 2010
'               http://www.utteraccess.com/forum/Split-string-capital-le-t1945127.html
' Adapted:      Bonnie Campbell, June 17, 2014
' Revisions:    6/17/2014 - BLC - XX
' ---------------------------------
Public Function InsertSpace(str As String) As String
     Dim strTemp As String
     Dim strChar As String
     Dim intLen As Integer
     
     If str > "" Then
          For intLen = 1 To Len(str)
               strChar = Mid(str, intLen, 1)
               If Asc(strChar) >= 65 And Asc(strChar) <= 90 Then
                    strTemp = strTemp & " " & strChar
               Else
                    strTemp = strTemp & strChar
               End If
          Next
     End If
        
     InsertSpace = strTemp
End Function

' ---------------------------------
' FUNCTION:     GetTempVarIndex
' Description:  Retrieves the index of a TempVar item
' Parameters:   strItem - item name(string)
' Returns:      index of item, if found (integer); not found returns -1
' Throws:       -
' References:   -
' Source/date:  Dal Jeanis, 7/11/2013
'               http://www.accessforums.net/modules/demo-module-vba-code-syntax-using-tempvars-36353.html
' Adapted:      Bonnie Campbell, Sep 1, 2014
' Revisions:    9/1/2014 - BLC - initial version
' ---------------------------------
Public Function GetTempVarIndex(strItem) As String
On Error GoTo Err_Handler

Dim i As Integer

    For i = 0 To [TempVars].count - 1
        If [TempVars].item(i).name = strItem Then
            'fetch the index and exit
            GetTempVarIndex = i
            Exit Function
        End If
    Next i
    
    'none found -> return -1
    GetTempVarIndex = -1
    
Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetTempVarIndex[mod_Utilities])"
    End Select
    Resume Exit_Function
End Function