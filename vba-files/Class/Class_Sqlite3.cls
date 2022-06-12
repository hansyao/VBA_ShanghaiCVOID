VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class_Sqlite3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'author by: Hans Yao <hansyow@gmail.com>
'blog: https://blog.oneplus-solution.com

'基于sqlite C语言接口
'see also: https://www.sqlite.org/capi3ref.html
'增删改查 语法
'增: insert into 表名(字段列表) values(值列表)
'删: delete from 表名 set 字段名 where 条件语句
'改: update from 表名 set 字段名=值 where 条件语句
'查: select * from 表名 where 条件语句 group by 分组字句 having...

'排序语句
'select * from person
'select * from person order by id desc
'select name from person group by name having count(*)>1

'分页(例:获取5条记录，跳过前三条记录)
'select * from Account limit 5 offset 3
'select * from Account limit 3,5

Private Const SQLITE_OK = 0          ' Successful result
Private Const SQLITE_ERROR = 1       ' Generic Error
Private Const SQLITE_BUSY = 5        ' The database file is locked
Private Const SQLITE_TOOBIG = 18     ' String or BLOB exceeds size limit
Private Const SQLITE_MISUSE = 21     ' Library used incorrectly
Private Const SQLITE_ROW = 100       ' sqlite3_step() has another row ready
Private Const SQLITE_DONE = 101      ' sqlite3_step() has finished executing
Private Const SQLITE_INTEGER = 1
Private Const SQLITE_FLOAT = 2
Private Const SQLITE_TEXT = 3
Private Const SQLITE_BLOB = 4
Private Const SQLITE_NULL = 5
Private Const CP_UTF8 = 65001
Private Const CP_GB2312 = 936
Private Const CP_GB18030 = 54936
Private Const CP_UTF7 = 65000

#If VBA7 Then
    Private Declare PtrSafe Function sqlite3_open Lib "winsqlite3.dll" (ByVal zFilename As String, ByRef ppDB As LongPtr) As LongPtr
    Private Declare PtrSafe Function sqlite3_close Lib "winsqlite3.dll" (ByVal db As LongPtr) As LongPtr
    Private Declare PtrSafe Function sqlite3_exec Lib "winsqlite3.dll" (ByVal db As LongPtr, ByVal sql As String, ByVal callback As LongPtr, ByVal argument_1 As LongPtr, ByRef errmsg As String) As LongPtr
    Private Declare PtrSafe Function sqlite3_prepare_v2 Lib "winsqlite3.dll" (ByVal db As LongPtr, ByVal zSql As String, ByVal nByte As LongPtr, ByRef ppStatement As LongPtr, ByRef pzTail As Long) As LongPtr
    Private Declare PtrSafe Function sqlite3_finalize Lib "winsqlite3.dll" (ByVal stmt As LongPtr) As LongPtr
    Private Declare PtrSafe Function sqlite3_bind_int Lib "winsqlite3.dll" (ByVal stmt As LongPtr, ByVal pos As Long, ByVal val As Long) As Long
    Private Declare PtrSafe Function sqlite3_bind_text_ Lib "winsqlite3.dll" Alias "sqlite3_bind_text" (ByVal stmt As LongPtr, ByVal pos As Long, ByVal val As LongPtr, ByVal len_ As Integer, ByVal whatIsThis As LongPtr) As Long
    Private Declare PtrSafe Function sqlite3_bind_null Lib "winsqlite3.dll" (ByVal stmt As LongPtr, ByVal pos As Long) As Long
    Private Declare PtrSafe Function sqlite3_step Lib "winsqlite3.dll" (ByVal stmt As LongPtr) As Long
    Private Declare PtrSafe Function sqlite3_reset Lib "winsqlite3.dll" (ByVal stmt As LongPtr) As Long
    Private Declare PtrSafe Function sqlite3_column_double Lib "winsqlite3.dll" (ByVal stmt As LongPtr, ByVal iCol As Integer) As Double
    Private Declare PtrSafe Function sqlite3_column_int Lib "winsqlite3.dll" (ByVal stmt As LongPtr, ByVal iCol As Integer) As Long
    Private Declare PtrSafe Function sqlite3_column_int64 Lib "winsqlite3.dll" (ByVal stmt As LongPtr, ByVal iCol As Integer) As LongPtr
    Private Declare PtrSafe Function sqlite3_column_text_ Lib "winsqlite3.dll" Alias "sqlite3_column_text" (ByVal stmt As LongPtr, ByVal iCol As Integer) As LongPtr
    Private Declare PtrSafe Function sqlite3_column_type Lib "winsqlite3.dll" (ByVal stmt As LongPtr, ByVal iCol As Integer) As Integer
    Private Declare PtrSafe Function sqlite3_column_count Lib "winsqlite3.dll" (ByVal stmt As LongPtr) As Long
    Private Declare PtrSafe Function sqlite3_column_name_ Lib "winsqlite3.dll" Alias "sqlite3_column_name" (ByVal stmt As LongPtr, ByVal iCol As Integer) As LongPtr
    Private Declare PtrSafe Function sqlite3_last_insert_rowid Lib "winsqlite3.dll" (ByVal stmt As LongPtr) As Long
    Private Declare PtrSafe Function sqlite3_get_table Lib "winsqlite3.dll" (ByVal db As LongPtr, ByVal sql As String, ByRef dbResult As LongPtr, ByRef nRow As LongPtr, ByRef nColumn As LongPtr, ByRef errmsg As LongPtr) As LongPtr
    Private Declare PtrSafe Function sqlite3_free_table Lib "winsqlite3.dll" (ByVal resultTbl As LongPtr)

    Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
    Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cchMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
#Else
    Private Declare Function sqlite3_open Lib "winsqlite3.dll" (ByVal zFilename As String, ByRef ppDB As Long) As Long
    Private Declare Function sqlite3_close Lib "winsqlite3.dll" (ByVal db As Long) As Long
    Private Declare Function sqlite3_exec Lib "winsqlite3.dll" (ByVal db As Long, ByVal sql As String, ByVal callback As Long, ByVal argument_1 As Long, ByRef errmsg As String) As Long
    Private Declare Function sqlite3_prepare_v2 Lib "winsqlite3.dll" (ByVal db As Long, ByVal zSql As String, ByVal nByte As Long, ByRef ppStatement As Long, ByRef pzTail As Long) As Long
    Private Declare Function sqlite3_finalize Lib "winsqlite3.dll" (ByVal stmt As Long) As Long
    Private Declare Function sqlite3_bind_int Lib "winsqlite3.dll" (ByVal stmt As Long, ByVal pos As Long, ByVal val As Long) As Long
    Private Declare Function sqlite3_bind_text_ Lib "winsqlite3.dll" Alias "sqlite3_bind_text" (ByVal stmt As Long, ByVal pos As Long, ByVal val As Long, ByVal len_ As Integer, ByVal whatIsThis As Long) As Long
    Private Declare Function sqlite3_bind_null Lib "winsqlite3.dll" (ByVal stmt As Long, ByVal pos As Long) As Long
    Private Declare Function sqlite3_step Lib "winsqlite3.dll" (ByVal stmt As Long) As Long
    Private Declare Function sqlite3_reset Lib "winsqlite3.dll" (ByVal stmt As Long) As Long
    Private Declare Function sqlite3_column_double Lib "winsqlite3.dll" (ByVal stmt As Long, ByVal iCol As Integer) As Double
    Private Declare Function sqlite3_column_int Lib "winsqlite3.dll" (ByVal stmt As Long, ByVal iCol As Integer) As Long
    Private Declare Function sqlite3_column_int64 Lib "winsqlite3.dll" (ByVal stmt As Long, ByVal iCol As Integer) As Long
    Private Declare Function sqlite3_column_text_ Lib "winsqlite3.dll" Alias "sqlite3_column_text" (ByVal stmt As Long, ByVal iCol As Integer) As Long
    Private Declare Function sqlite3_column_type Lib "winsqlite3.dll" (ByVal stmt As Long, ByVal iCol As Integer) As Integer
    Private Declare Function sqlite3_column_count Lib "winsqlite3.dll" (ByVal stmt As Long) As Long
    Private Declare Function sqlite3_column_name_ Lib "winsqlite3.dll" Alias "sqlite3_column_name" (ByVal stmt As Long, ByVal iCol As Integer) As Long
    Private Declare Function sqlite3_last_insert_rowid Lib "winsqlite3.dll" (ByVal stmt As Long) As Long
    Private Declare Function sqlite3_get_table Lib "winsqlite3.dll" (ByVal db As Long, ByVal sql As String, ByRef dbResult As Long, ByRef nRow As Long, ByRef nColumn As Long, ByRef errmsg As Long) As Long
    Private Declare Function sqlite3_free_table Lib "winsqlite3.dll" (ByVal resultTbl As Long)

    Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
    Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
#End If

''' Return length of byte array or zero if uninitialized
Private Function BytesLength(abBytes() As Byte) As Long
    ' Trap error if array is uninitialized
    On Error Resume Next
    BytesLength = UBound(abBytes) - LBound(abBytes) + 1
End Function

''' Return VBA "Unicode" string from byte array encoded in UTF-8
Public Function Utf8BytesToString(abUtf8Array() As Byte) As String
    Dim nBytes As Long
    Dim nChars As Long
    Dim strOut As String
'    Dim cp_UTF8 As Long
    Utf8BytesToString = ""
    ' Catch uninitialized input array
    nBytes = BytesLength(abUtf8Array)
    If nBytes <= 0 Then Exit Function
    ' Get number of characters in output string
    nChars = MultiByteToWideChar(CP_UTF8, 0&, VarPtr(abUtf8Array(0)), nBytes, 0&, 0&)
    ' Dimension output buffer to receive string
    strOut = String(nChars, 0)
    nChars = MultiByteToWideChar(CP_UTF8, 0&, VarPtr(abUtf8Array(0)), nBytes, StrPtr(strOut), nChars)
    Utf8BytesToString = Left$(strOut, nChars)
End Function

#If VBA7 Then
    Private Function utf8ptrToString(ByVal pUtf8String As LongPtr) As String
#Else
    Private Function utf8ptrToString(ByVal pUtf8String As Long) As String
#End If

    Dim buf     As String
    Dim cSize   As Long
    Dim retVal  As Long

    cSize = MultiByteToWideChar(CP_UTF8, 0, pUtf8String, -1, 0, 0)

    If cSize <= 1 Then
        utf8ptrToString = ""
        Exit Function
    End If

    utf8ptrToString = String(cSize - 1, "*")

    retVal = MultiByteToWideChar(CP_UTF8, 0, pUtf8String, -1, StrPtr(utf8ptrToString), cSize)
    If retVal = 0 Then
       Err.Raise 1000, "utf8ptrToString", "Utf8ptrToString error: " & Err.lastDllError
       Exit Function
    End If

End Function

#If VBA7 Then
    Private Function ptrToString(ByVal pString As LongPtr, ByVal CodePage As Long) As String
#Else
    Private Function ptrToString(ByVal pString As Long, ByVal CodePage As Long) As String
#End If
    Dim buf     As String
    Dim cSize   As Long
    Dim retVal  As Long

    cSize = MultiByteToWideChar(CodePage, 0, pString, -1, 0, 0)

    If cSize <= 1 Then
        ptrToString = ""
        Exit Function
    End If

    ptrToString = String(cSize - 1, "*")

    retVal = MultiByteToWideChar(CodePage, 0, pString, -1, StrPtr(ptrToString), cSize)
    If retVal = 0 Then
       Err.Raise 1000, "ptrToString", "ptrToString error: " & Err.lastDllError
       Exit Function
    End If

End Function

Public Function stringToUtf8bytes(ByVal txt As String) As Byte() ' {

    Dim bSize  As Long
    Dim retVal As Long
    Dim buf()  As Byte

    bSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(txt), -1, 0, 0, 0, 0)

    If bSize = 0 Then ' {
        Exit Function
    End If ' }

    ReDim buf(bSize)

    retVal = WideCharToMultiByte(CP_UTF8, 0, StrPtr(txt), -1, VarPtr(buf(0)), bSize, 0, 0)

    If retVal = 0 Then
        Err.Raise 1000, "stringToUtf8bytes", "stringToUtf8bytes error: " & Err.lastDllError
        Exit Function
    End If

    stringToUtf8bytes = buf

End Function

#If VBA7 And Win64 Then
    Private Function sqlite3_bind_text(ByVal stmt As LongPtr, ByVal pos As Long, ByVal val As String, ByVal len_ As Integer, ByVal whatIsThis As LongPtr) As Long
#Else
    Private Function sqlite3_bind_text(ByVal stmt As Long, ByVal pos As Long, ByVal val As String, ByVal len_ As Integer, ByVal whatIsThis As Long) As Long
#End If
  Dim arrayVariant As Variant
  arrayVariant = stringToUtf8bytes(val)

  sqlite3_bind_text = sqlite3_bind_text_(stmt, pos, VarPtr(arrayVariant), len_, whatIsThis)

End Function
#If VBA7 And Win64 Then
    Private Function sqlite3_column_name(ByVal stmt As LongPtr, ByVal iCol As Long, Optional ByVal encode As Long = CP_UTF8) As String
#Else
    Private Function sqlite3_column_name(ByVal stmt As Long, ByVal iCol As Long, Optional ByVal encode As Long = CP_UTF8) As String
#End If
    sqlite3_column_name = ptrToString(sqlite3_column_name_(stmt, iCol), encode)
End Function
#If VBA7 And Win64 Then
    Private Function sqlite3_column_text(ByVal stmt As LongPtr, ByVal iCol As Long, Optional ByVal encode As Long = CP_UTF8) As String
#Else
    Private Function sqlite3_column_text(ByVal stmt As Long, ByVal iCol As Long, Optional ByVal encode As Long = CP_UTF8) As String
#End If
    sqlite3_column_text = ptrToString(sqlite3_column_text_(stmt, iCol), encode)
End Function

#If VBA7 And Win64 Then
    Public Function recordCount(ByVal db As LongPtr, ByVal sql As String) As Long
    Dim stmt As LongPtr
#Else
    Public Function recordCount(ByVal db As Long, ByVal sql As String) As Long
    Dim stmt As Long
#End If
    Dim count As Long

    sql = "select count(0) from (" & sql & ")"

    stmt = prepareStmt(db, sql)
    While sqlite3_step(stmt) <> SQLITE_DONE
        count = sqlite3_column_int(stmt, 0)
    Wend

    '释放内存
    sqlite3_finalize stmt

    recordCount = count
End Function

#If VBA7 And Win64 Then
    Public Function columnCount(ByVal db As LongPtr, ByVal sql As String) As Long
    Dim stmt As LongPtr
#Else
    Public Function columnCount(ByVal db As Long, ByVal sql As String) As Long
    Dim stmt As Long
#End If
    Dim count As Long

    stmt = prepareStmt(db, sql)
    count = sqlite3_column_count(stmt)

    '释放内存
    sqlite3_finalize stmt

    columnCount = count
End Function

#If VBA7 And Win64 Then
    Public Function openDB(FileName As String) As LongPtr
    Dim res As LongPtr
#Else
    Public Function openDB(FileName As String) As Long
    Dim res As Long
#End If

    res = sqlite3_open(FileName, openDB)
    If res <> SQLITE_OK Then
        Err.Raise 1000, "openDB", "sqlite_open failed, res = " & res
    End If

'    Debug.Print ("SQLite db opened, db = " & openDB)

End Function

#If VBA7 And Win64 Then
    Public Function closeDB(db As LongPtr)
    Dim res As LongPtr
#Else
    Public Function closeDB(db As Long)
    Dim res As Long
#End If

    res = sqlite3_close(db)
    If res <> SQLITE_OK Then
        Err.Raise 1000, "closeDB", "sqlite_open failed, res = " & res
    End If

End Function

Private Function checkBindRetval(retVal As Long)

    If retVal = SQLITE_OK Then
        Exit Function
    End If

    If retVal = SQLITE_TOOBIG Then
        Err.Raise 1000, "checkBindRetval", "bind failed: String or BLOB exceeds size limit"
    End If

    If retVal = SQLITE_MISUSE Then
        Err.Raise 1000, "checkBindRetval", "bind failed: Library used incorrectly"
    End If

    Err.Raise 1000, "checkBindRetval", "bind failed, retVal = " & retVal

End Function

Private Function checkStepRetval(retVal As Long)

    If retVal = SQLITE_DONE Then
        Exit Function
    End If

    Err.Raise 1000, "checkStepRetval", "step failed, retVal = " & retVal

End Function
#If VBA7 And Win64 Then
    Public Function execSQL(db As LongPtr, sql As String)
    Dim res    As LongPtr
#Else
    Public Function execSQL(db As Long, sql As String)
    Dim res    As Long
#End If
    Dim errmsg As String

    res = sqlite3_exec(db, sql, 0, 0, errmsg)
    If res <> SQLITE_OK Then
        Err.Raise 1000, "execSQL", "sqlite3_exec failed, res = " & res
    End If

End Function

#If VBA7 And Win64 Then
    Private Function prepareStmt(db As LongPtr, sql As String) As LongPtr
    Dim res    As LongPtr
#Else
    Public Function prepareStmt(db As Long, sql As String) As Long
    Dim res As Long
#End If

    res = sqlite3_prepare_v2(db, sql, -1, prepareStmt, 0)
    If res <> SQLITE_OK Then
        Err.Raise 1000, "prepareStmt", "sqlite3_prepare failed, res = " & res
    End If

'    Debug.Print ("stmt = " & prepareStmt)

End Function

Function createTbl(ByVal tblName As String, ByVal colStr As String, Optional isExist As Boolean = True) As String
    Dim isExistStr As String

    colStr = Join(Split(colStr, " TEXT,"), "," & vbCrLf)
    If isExist Then isExistStr = "IF NOT EXISTS"
    createTbl = "CREATE TABLE " & isExistStr & " " & tblName & "(" & colStr & ");"
End Function

#If VBA7 And Win64 Then
    Function queryToArray(ByVal db As LongPtr, ByVal sql As String, Optional ByVal encode As Long = 936, Optional ByVal isTitle As Boolean = False)
    Dim stmt As LongPtr
#Else
    Function queryToArray(ByVal db As Long, ByVal sql As String, Optional ByVal encode As Long = 936, Optional ByVal isTitle As Boolean = False)
    Dim stmt As Long
#End If
    Dim RowNo As Long
    Dim colNo As Long
    Dim colCount As Long
    Dim RowCount As Long
    Dim myArray() As Variant

    RowCount = recordCount(db, sql)

    If RowCount = 0 Then Exit Function

    stmt = prepareStmt(db, sql)
    sqlite3_bind_int stmt, 1, 2

    colCount = sqlite3_column_count(stmt)

    If isTitle Then
        If colCount = 1 Then
            ReDim myArray(RowCount)
            myArray(0) = sqlite3_column_name(stmt, colNo, encode)
        Else
            ReDim myArray(RowCount, colCount)
            While colNo < colCount
                myArray(0, colNo) = sqlite3_column_name(stmt, colNo, encode)
                colNo = colNo + 1
            Wend
        End If
        RowNo = 1
    Else
        If colCount = 1 Then
            ReDim myArray(RowCount - 1)
        Else
            ReDim myArray(RowCount - 1, colCount - 1)
        End If
        RowNo = 0
    End If

    While sqlite3_step(stmt) <> SQLITE_DONE
        colNo = 0
        If colCount = 1 Then
            Select Case sqlite3_column_type(stmt, colNo)
                Case SQLITE_INTEGER
                    myArray(RowNo) = sqlite3_column_int64(stmt, colNo)
                Case SQLITE_FLOAT
                    myArray(RowNo) = sqlite3_column_double(stmt, colNo)
                Case SQLITE_TEXT
                    myArray(RowNo) = sqlite3_column_text(stmt, colNo, encode)
                Case SQLITE_NULL
                    myArray(RowNo) = "n/a"
                Case Else
                    myArray(RowNo) = "?"
            End Select
        Else
            While colNo < colCount
                Select Case sqlite3_column_type(stmt, colNo)
                    Case SQLITE_INTEGER
                        myArray(RowNo, colNo) = sqlite3_column_int64(stmt, colNo)
                    Case SQLITE_FLOAT
                        myArray(RowNo, colNo) = sqlite3_column_double(stmt, colNo)
                    Case SQLITE_TEXT
                        myArray(RowNo, colNo) = sqlite3_column_text(stmt, colNo, encode)
                    Case SQLITE_NULL
                        myArray(RowNo, colNo) = "n/a"
                    Case Else
                        myArray(RowNo, colNo) = "?"
                End Select
                colNo = colNo + 1
            Wend
        End If
        If (RowCount Mod 10000) = 0 Then DoEvents
        RowNo = RowNo + 1
    Wend

    '释放内存
    sqlite3_finalize stmt

    queryToArray = myArray

End Function
