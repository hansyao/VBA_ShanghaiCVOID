VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3345
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3075
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Unload Me
End Sub

Private Sub ListBox1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim i As Long
    Dim n As Integer
    Dim myJson As Object
    
    For i = 1 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) Then
                Set myJson = JSON.ParseJson("{}")
            For j = 0 To ListBox1.columnCount
                myJson(ListBox1.List(0, j)) = (ListBox1.List(i, j))
            Next j
            myJson("阳性历史日期") = sqlite3.queryToArray(db, "SELECT end_date from by_location where location = '" & myJson("居住地") & "' order by end_date desc")
            Select Case myJson("区域划分")
                Case "封控区"
                    n = vbCritical
                Case "管控区"
                    n = vbExclamation
                Case "防范区"
                    n = vbInformation
            End Select
            MsgBox JSON.ConvertToJson(myJson, Whitespace:=2), n, "新冠疫情状态: " & myJson("居住地")
            Exit Sub
        End If
    Next
    'Call addSeries(myChart, lng, lat)
End Sub


Private Sub UserForm_Activate()
    With Me
        .Left = Windows(1).Left
        .Top = Windows(1).Top
        .Width = 1000
        .Height = 190
    End With

    With Me.ListBox1
        .Left = 0
        .Top = 0
        .Height = 120
        .Width = Me.InsideWidth
        .MultiSelect = fmMultiSelectSingle
        .columnCount = UBound(userformVar, 2)
        .List = userformVar
        On Error Resume Next
        .Selected(0) = True
    End With

    With CommandButton1
        .Top = Me.InsideHeight - .Height
        .Width = Me.InsideWidth
        .Left = 0
    End With

    With Me.Label1
            .Width = Me.InsideWidth
            .Left = 0
            .Caption = "查询到" & UBound(userformVar) + 1 & "个疫情居住地"
    End With
    
    'Call showList
End Sub

Private Sub UserForm_Initialize()
    Dim i As Long
    
    Me.Caption = "选择查看"
End Sub

Private Sub UserForm_Terminate()
    'deleteMyShape "info"
    'removeLocationSeries "mapChart"
End Sub



