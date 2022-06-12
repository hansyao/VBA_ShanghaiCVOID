VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class_HttpResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
'Attribute VB_Name = "HttpResponse"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] HttpResponseクラス
'* [詳  細] HttpResponse情報のデータクラス。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mStatusCd As String
Private mStatusTxt As String
Private mBody As String
Private mHeaders As String
Private mEncode As String
Private mOriginalBody As Boolean

Private Sub Class_Initialize()
    mEncode = RESPONSE_CHARSETS
    mOriginalBody = RESPONSE_ORIGINAL_BODY
End Sub

'******************************************************************************
'* プロパティ定義
'******************************************************************************
'Status Code
Public Property Get StatusCd() As String
    StatusCd = mStatusCd
End Property
Public Property Let StatusCd(arg As String)
    mStatusCd = arg
End Property

'Status Text
Public Property Get StatusText() As String
    StatusText = mStatusTxt
End Property
Public Property Let StatusText(arg As String)
    mStatusTxt = arg
End Property

'Response Body
Public Property Get body() As String
    body = mBody
End Property
Public Property Let body(arg As String)
    'check response header whether include gzip
    mBody = arg
End Property

'Response Headers
Public Property Get Headers() As String
    Headers = mHeaders
End Property
Public Property Let Headers(arg As String)
    mHeaders = arg
End Property

'charsets
Public Property Get encode() As String
    encode = mEncode
End Property
Public Property Let encode(arg As String)
    mEncode = arg
End Property

'charsets
Public Property Get OriginalBody() As String
    OriginalBody = mOriginalBody
End Property
Public Property Let OriginalBody(arg As String)
    mOriginalBody = arg
End Property

Public Function GetHeader(arg As String) As String
    Dim i As Integer
    Dim headArry As Variant
    Dim Key As String
    Dim Value As String
    
    GetHeader = ""
    On Error GoTo err_handler
    headArry = Split(mHeaders, vbNewLine)
    For i = 0 To UBound(headArry)
        On Error GoTo continue
        Key = Split(headArry(i), ":")(0)
        Value = Trim(Replace(headArry(i), Key & ":", ""))
        If LCase(Key) = LCase(arg) Then
            GetHeader = Value & ";" & GetHeader
            
        End If
continue:
        If Err.Number <> 0 Then
            Err.Clear
            Exit For
        End If
    Next

err_handler:
    If Err.Number <> 0 Then
        Debug.Print "getHeader: " & Err.Description
    End If
End Function
