Attribute VB_Name = "ShCovid19"

'''+----                                                                   --+
'''|                             上海新冠疫情查询
'''|                tech stack: sqlite3, excel XY chart
'''|                @author Hans Yao <hansyow@gmail.com>
'''|                  Copyright (c) 2019-2022 Hans Yao
'''|          The Project Page: https://github.com/hansyao/VBA_ShanghaiCVOID
'''|          my blog: https://blog.oneplus-solution.com
'''+--                                                                   ----+

Option Explicit

Public Const RESPONSE_CHARSETS As String = "utf-8"
Public Const RESPONSE_ORIGINAL_BODY As Boolean = False
Public userformVar As Variant

Private Const CP_UTF8 = 65001
Private Const CP_GB2312 = 936
Private Const CP_GB18030 = 54936
Private Const CP_UTF7 = 65000
Private Const AMAPAPI = "8358809e179a92657376738b3508029c"

Public Const district = "浦东新,黄浦,静安,徐汇,长宁,普陀,虹口,杨浦,宝山,闵行,嘉定,金山,松江,青浦,奉贤,崇明"
Private Const adCode = "310000,310101,310104,310105,310106,310107,310109,310110,310112,310113,310114,310115,310116,310117,310118,310120,310151"
Private Const adCodeDescription = "上海市,黄浦区,徐汇区,长宁区,静安区,普陀区,虹口区,杨浦区,闵行区,宝山区,嘉定区,浦东新区,金山区,松江区,青浦区,奉贤区,崇明区"
Private Const exclude = "滑动,居住于,已对,资料,编辑,上海发布,各区信息,市卫健委,本市各区"
Private Const sHeader = "release_date,location,district,latest_date,dDiff,category,dUnlockdown,dDiffUnlockdown,lng,lat,formattedAddress,businessAreas,township"
Private Const sHeaderCN = "发布日期,居住地,辖区,最后一次阳性日期,未出现阳性天数,区域划分,估算解封日期,解封剩余天数,经度坐标,纬度坐标,详细地址,商圈,所属居委"

#If VBA7 Then
'// 64bit API Declarations
    Public Declare PtrSafe Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare PtrSafe Sub CopyMemorybyPtr Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByVal Source As LongPtr, ByVal Length As Long)
#Else
'// 32bit API Declarations
    Public Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare Sub CopyMemorybyPtr Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByVal Source As Long, ByVal Length As Long)
#End If

Private byCity() As Variant
Private byDistrict() As Variant
Private byLocation() As Variant
Private cookie As String
Private isHistoryExists As Boolean
Public JSON As New Class_JsonConverter

'sqlite3 database variant
#If VBA7 Then
    Public db As LongPtr
#Else
    Public db As Long
#End If
Private dbName As String
Public sqlite3 As New Class_Sqlite3
Private appData As String
Private dbFolder As String

Private db_struct As db_struct
Private Type db_struct
    tbl_url_list As String
    tbl_by_city As String
    tbl_by_district As String
    tbl_by_location As String
    tbl_mapdata As String
    tbl_poi As String
End Type

Public Function codeSet() As Long
    Dim wmi As Object
    Dim OS As Object
    Dim cs As Long

    Set wmi = GetObject("winmgmts:")

    For Each OS In wmi.ExecQuery("SELECT * FROM Win32_OperatingSystem")
      cs = VBA.CLng(OS.codeSet)
    Next
    codeSet = cs

    Set wmi = Nothing
End Function

Public Function ArrayDimension(ByVal Arrary1 As Variant) As Long
    '返回值: -1 非数组, 0 空数组, >0 当前数组维度
#If VBA7 Then
    Dim Ptr1 As LongPtr
#Else
    Dim Ptr1 As Long
#End If
    Dim i As Long
    Ptr1 = VarPtr(Arrary1)
    CopyMemorybyPtr i, Ptr1, Len(i)
    If i And 8192& Then
        If i And &H4000& Then
            CopyMemorybyPtr Ptr1, Ptr1 + 8, Len(Ptr1)
        Else
            Ptr1 = Ptr1 + 8
        End If
        CopyMemorybyPtr Ptr1, Ptr1, Len(Ptr1)
        If Ptr1 Then CopyMemorybyPtr ArrayDimension, Ptr1, 2
    Else
        ArrayDimension = -1
    End If
End Function

Function sheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet

    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    sheetExists = Not sht Is Nothing
    Err.Clear
End Function

''' @return As Object Is Scripting.FileSystemObject
Private Property Get fso() As Object
    Dim xxFSO As Object

    If xxFSO Is Nothing Then Set xxFSO = CreateObject("Scripting.FileSystemObject")
    Set fso = xxFSO
    Set xxFSO = Nothing
End Property

Private Property Get webreq() As Object
    Dim xxWebReq As New Class_HttpRequest

    xxWebReq.SetReqHeaderUserAgent "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36 Edg/98.0.1108.56"
    xxWebReq.SetReqHeaderAccept "application/json, text/javascript, text/html"
    xxWebReq.SetReqHeaderAcceptLanguage "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6"

    Set webreq = xxWebReq
    Set xxWebReq = Nothing
End Property

Private Function FolderExists(ByVal folderSpec As String) As Boolean
    FolderExists = fso.FolderExists(folderSpec)
End Function

Private Function FileExists(ByVal fileSpec As String) As Boolean
    FileExists = fso.FileExists(fileSpec)
End Function

'@getTimeStamp(t As Date) 获取时间戳return as Long
Private Function getTimeStamp(t As Date) As Long
    Dim t2 As Long, t1 As Date

    t1 = "1970-1-1 8:00"
    t2 = (t - t1) * 86400

    getTimeStamp = t2
End Function

Private Function isInArray(ByVal myStr As String, ByVal myArray) As Boolean
    Dim str As Variant

    For Each str In myArray
        If str = myStr Then
            isInArray = True
            Exit Function
        End If
    Next
    isInArray = False
End Function

Private Function findDistrict(ByVal myStr As String) As String
    Dim str As Variant

    For Each str In VBA.Split(district, ",")
        If VBA.InStr(1, myStr, str) > 0 Then
            findDistrict = str
            Exit Function
        End If
    Next
End Function

'@TransposeArray(ByVal arrA As Variant)  转置二维数组
Private Function TransposeArray(ByVal arrA As Variant) As Variant
    Dim aRes() As Variant
    Dim i As Long
    Dim j As Long

    If VBA.IsArray(arrA) Then
        ReDim aRes(LBound(arrA, 2) To UBound(arrA, 2), LBound(arrA, 1) To UBound(arrA, 1))
        For i = LBound(arrA, 1) To UBound(arrA, 1)
            For j = LBound(arrA, 2) To UBound(arrA, 2)
                aRes(j, i) = arrA(i, j)
            Next
        Next
        TransposeArray = aRes
    End If
End Function

Private Function isExclude(ByVal myStr As String, ByVal exclude As String) As Boolean
    Dim myArray As Variant
    Dim i As Long

    myArray = Split(exclude, ",")
    For i = 0 To UBound(Split(exclude, ","))
        If VBA.InStr(1, myStr, myArray(i)) > 0 Then
            isExclude = True
            Exit Function
        End If
    Next i
    isExclude = False
End Function

'@MatchesExp(ByRef strng): 正则 - 判断是否包含数字
Private Function isInclNum(ByRef strng) As Boolean
    Dim regEx As Object

    Set regEx = CreateObject("vbscript.regexp")
    With regEx
        .Global = True
        .Pattern = "\d+"
        If .test(strng) Then                   '复核提取情况
            isInclNum = True
        Else
            isInclNum = False
        End If
    End With

End Function

'@MatchesExp(ByRef strng,ByRef patrn): 正则 - 匹配字符串,返回匹配集合
Private Function MatchesExp(ByRef strng, ByRef patrn)
    Dim regEx As Object

    Set regEx = CreateObject("vbscript.regexp")
    regEx.Pattern = patrn
    regEx.IgnoreCase = True
    regEx.Global = True
    Set MatchesExp = regEx.Execute(strng)

    Set regEx = Nothing
End Function

#If VBA7 Then
Public Function mapDataToDb(ByVal db As LongPtr, ByVal tbl As String)
#Else
Public Function mapDataToDb(ByVal db As Long, ByVal tbl As String)
#End If
    Dim http As New Class_HttpUtils
    Dim Resp As New Class_HttpResponse
    Dim req As Object
    Dim url As String
    Dim adc As Variant
    Dim adcd As Variant
    Dim status As Integer
    Dim i As Long
    Dim mapData() As Variant

    Set req = webreq

    adc = VBA.Split(adCode, ",")
    adcd = VBA.Split(adCodeDescription, ",")
    For i = 0 To UBound(adc)
        If SafeArrayGetDim(mapData) > 0 Then
            ReDim Preserve mapData(1, UBound(mapData, 2) + 1)
        Else
            ReDim mapData(1, 0)
        End If
        req.url = "https://restapi.amap.com/v3/config/district?key=" & AMAPAPI & _
            "&keywords=" & adc(i) & "&subdistrict=3&extensions=all"
        Set Resp = http.GetReq(req)
        If Resp.StatusCd <> "200" Then GoTo continue
        If JSON.ParseJson(Resp.body)("status") <> 1 Then GoTo continue

        mapData(0, UBound(mapData, 2)) = adcd(i)
        mapData(1, UBound(mapData, 2)) = Resp.body
continue:
    Next

    '将查询结果写入sqlite数据库
    sqlite3.execSQL db, "BEGIN TRANSACTION"
    For i = 0 To UBound(mapData, 2)
        On Error GoTo err_handler
        sqlite3.execSQL db, "insert into " & tbl & " values(" & "'" & mapData(0, i) & "','" & mapData(1, i) & "'" & ");"
    Next i
    Erase mapData
    sqlite3.execSQL db, "COMMIT"

    Set req = Nothing
    Set Resp = Nothing

err_handler:
    If Err.Number <> 0 Then
        sqlite3.execSQL db, "ROLLBACK TRANSACTION"
        Err.Clear
    End If
End Function

Private Function mapDataToJson() As Object
    Dim sql As String
    Dim myArray As Variant
    Dim areaArray As Variant
    Dim area As String
    Dim m As Long
    Dim n As Long
    Dim i As Long
    Dim j As Long
    Dim recordCount As Long
    Dim myJson1 As Object
    Dim myJson2 As Object
    Dim coordinate As Variant
    Dim coordinate1 As Variant
    Dim coordinate2 As Variant
    Dim longitude As String
    Dim latitude As String
    Dim delimiter As String
    Dim center As String
    Dim centerList As String
    Dim longitudeCenter As String
    Dim latitudeCenter As String

    If db = 0 Then Call initDb

    Set myJson1 = JSON.ParseJson("{}")
    myArray = VBA.Split(adCodeDescription, ",")
    For i = 0 To UBound(myArray)
        area = myArray(i)
        sql = "select * from mapdata where area='" & area & "'"
        recordCount = sqlite3.recordCount(db, sql)
        If recordCount = 0 Then GoTo continue
        '将sqlite数据库导出到JSON对象
        areaArray = sqlite3.queryToArray(db, sql, CP_GB2312, False)
        myJson1(areaArray(0, 0)) = areaArray(0, 1)
        Erase areaArray
continue:
    Next i

    '获取区域边界坐标转换为excel数组myJSON2
    Set myJson2 = JSON.ParseJson("{}")
    coordinate = VBA.Split(adCodeDescription, ",")
    For m = 0 To UBound(coordinate)
        coordinate1 = VBA.Split(JSON.ParseJson(myJson1(coordinate(m)))("districts")(1)("polyline"), "|")
        center = VBA.Trim(JSON.ParseJson(myJson1(coordinate(m)))("districts")(1)("center"))
        If m = 0 Then delimiter = "" Else delimiter = ";"
        longitudeCenter = longitudeCenter & delimiter & Split(center, ",")(0)
        latitudeCenter = latitudeCenter & delimiter & Split(center, ",")(1)
        centerList = centerList & delimiter & coordinate(m)
        For i = 0 To UBound(coordinate1)
            Set myJson2(coordinate(m) & i + 1) = JSON.ParseJson("{}")
            Set myJson2(coordinate(m) & i + 1)("longitude") = JSON.ParseJson("{}")
            Set myJson2(coordinate(m) & i + 1)("latitude") = JSON.ParseJson("{}")
            coordinate2 = VBA.Split(coordinate1(i), ";")
            For j = 0 To UBound(coordinate2)
                If j = 0 Then delimiter = "" Else delimiter = ";"
                longitude = longitude & delimiter & VBA.Trim(VBA.Split(coordinate2(j), ",")(0))
                latitude = latitude & delimiter & VBA.Trim(VBA.Split(coordinate2(j), ",")(1))
            Next
            myJson2(coordinate(m) & i + 1)("longitude") = longitude
            myJson2(coordinate(m) & i + 1)("latitude") = latitude

            longitude = ""
            latitude = ""
        Next
    Next
    Set myJson2("中心标记") = JSON.ParseJson("{}")
    myJson2("中心标记")("centerList") = centerList
    myJson2("中心标记")("longitude") = longitudeCenter
    myJson2("中心标记")("latitude") = latitudeCenter

    Set mapDataToJson = myJson2

    Set myJson1 = Nothing
    Set myJson2 = Nothing
End Function

Private Function getEntrance(ByVal mainUri As String) As Variant
    Dim http As New Class_HttpUtils
    Dim Resp As New Class_HttpResponse
    Dim req As Object
    Dim htmlDoc  As Object 'As MSHTML.HTMLDocument
    Dim urlList As Object
    Dim patrn As String
    Dim totalPage As Long
    Dim i As Integer
    Dim j As Integer
    Dim urlListArray() As String
    Dim mDate As String
    Dim url As String
    Dim max As Variant

    Set req = webreq

    req.url = mainUri & "index.html"

    Set Resp = http.GetReq(req)
    Set htmlDoc = VBA.CreateObject("htmlfile")

    'Debug.Print Resp.body

    '获取页码
    patrn = """setPage"",[0-9],[0-9]*"
    totalPage = Split(MatchesExp(Resp.body, patrn)(0), ",")(2)

    For i = 1 To totalPage
        If i <> 1 Then
            req.url = mainUri & "index_" & i & ".html"
            cookie = Resp.GetHeader("set-cookie")
            If Resp.StatusCd <> "200" Then GoTo continue
        End If

        If Len(cookie) > 0 Then req.SetReqHeaderCookie cookie
        Set Resp = http.GetReq(req)

        If Resp.StatusCd <> "200" Then GoTo continue

        htmlDoc.body.innerhtml = Resp.body
        htmlDoc.body.innerhtml = htmlDoc.getElementsByTagName("ul")(0).innerhtml

        For Each urlList In htmlDoc.getElementsByTagName("li")
            If VBA.InStr(1, urlList.innerText, "居住地信息") > 0 Then
                mDate = VBA.Format(urlList.getElementsByTagName("SPAN")(0).innerText, "yyyy-mm-dd")
                max = getMax("url_list", "release_date")
                If mDate <= getMax("url_list", "release_date") And max <> "n/a" Then GoTo final
                url = VBA.Trim(VBA.Replace(urlList.getElementsByTagName("A")(0).href, "about:/", "https://wsjkw.sh.gov.cn/"))
                If SafeArrayGetDim(urlListArray) > 0 Then
                    ReDim Preserve urlListArray(1, UBound(urlListArray, 2) + 1)
                Else
                    ReDim urlListArray(1, 0)
                End If
                urlListArray(0, UBound(urlListArray, 2)) = mDate
                urlListArray(1, UBound(urlListArray, 2)) = url
            End If
        Next
continue:
    Next i

final:
    getEntrance = urlListArray

    Set htmlDoc = Nothing
    Set Resp = Nothing
    Set urlList = Nothing
    Set req = Nothing

End Function
Private Function getDetail(ByVal url As String, ByVal releaseDate As String) As Variant
    Dim http As New Class_HttpUtils
    Dim Resp As New Class_HttpResponse
    Dim req As Object
    Dim htmlDoc  As Object 'As MSHTML.HTMLDocument
    Dim location As Object
    Dim myLine As String
    Dim currentDistrict As String
    Dim mDate As String
    Dim confirmedCase As Long
    Dim asymptomaticCase As Long
    Dim sourceId As String
    Dim tag As String
    Dim mLocation As String
    Dim mUbound As Long
    Dim regMatch As Variant

    Set req = webreq
    req.url = url

    Set Resp = http.GetReq(req)
    If Resp.StatusCd <> "200" Then Exit Function

    Set htmlDoc = VBA.CreateObject("htmlfile")

    cookie = Resp.GetHeader("set-cookie")
    If Len(cookie) > 0 Then req.SetReqHeaderCookie cookie
    'Debug.Print Resp.body

    htmlDoc.body.innerhtml = Resp.body

    If VBA.InStr(1, req.url, "https://mp.weixin.qq.com") > 0 Then
        sourceId = "js_content"
    ElseIf VBA.InStr(1, req.url, "https://wsjkw.sh.gov.cn") > 0 Then
        sourceId = "ivs_content"
    Else
        Debug.Print "无效URL数据源"
        Exit Function
    End If
    tag = "p"

    htmlDoc.body.innerhtml = htmlDoc.getElementById(sourceId).innerhtml

    For Each location In htmlDoc.getElementsByTagName(tag)
        myLine = location.innerText

        '全市汇总
        If VBA.Left(myLine, 4) = "市卫健委" Then
            mDate = MatchesExp(myLine, "\d+年\d+月\d+日")(0)
            mDate = VBA.IIf(VBA.InStr(mDate, "年") = 0, VBA.Split(releaseDate, "-")(0) & "年" & mDate, mDate)
            mDate = VBA.Format(mDate, "yyyy-mm-dd")

            On Error Resume Next
            confirmedCase = VBA.Replace(MatchesExp(myLine, "确诊病例\d+")(0), "确诊病例", "")

            If Err.Number <> 0 Then confirmedCase = 0
            VBA.Err.Clear

            On Error Resume Next
            asymptomaticCase = VBA.Replace(MatchesExp(myLine, "无症状感染者\d+")(0), "无症状感染者", "")

            If Err.Number <> 0 Then asymptomaticCase = 0
            VBA.Err.Clear
            If SafeArrayGetDim(byCity) > 0 Then
                ReDim Preserve byCity(4, UBound(byCity, 2) + 1)
            Else
                ReDim byCity(4, 0)
            End If
            mUbound = UBound(byCity, 2)
            byCity(0, mUbound) = releaseDate
            byCity(1, mUbound) = mDate
            byCity(2, mUbound) = confirmedCase
            byCity(3, mUbound) = asymptomaticCase
            byCity(4, mUbound) = myLine
        End If


        'If VBA.InStr(myLine, "居住于") > 0 Then
        If RegExpTest("(?=.*(" & VBA.Replace(district, ",", "|") & "))(?=.*\d+年\d+月\d+日)", myLine, False) Then
            mDate = MatchesExp(myLine, "\d+月\d+日|\d+年\d+月\d+日")(0)
            mDate = VBA.IIf(VBA.Left(mDate, 1) = "年", VBA.Split(releaseDate, "-")(0) & mDate, mDate)
            mDate = VBA.IIf(VBA.InStr(mDate, "年") = 0, VBA.Split(releaseDate, "-")(0) & "年" & mDate, mDate)
            mDate = VBA.Format(mDate, "yyyy-mm-dd")

            'Debug.Assert Not RegExpTest("金山区", myLine, False)
            'Debug.Assert currentDistrict = "区"
            If findDistrict(myLine) & "区" = "区" Then GoTo continue
            currentDistrict = findDistrict(myLine) & "区"

        '区域汇总
            On Error Resume Next
            confirmedCase = MatchesExp(MatchesExp(myLine, "\d+例(.*)确诊")(0), "\d+")(0)
            If VBA.Err.Number <> 0 Then
                confirmedCase = 0
            End If
            If VBA.Err.Number <> 0 Then confirmedCase = 0
            VBA.Err.Clear

            Set regMatch = MatchesExp(MatchesExp(myLine, "\d+例(.*)无症状")(0), "\d+")
            If MatchesExp(myLine, "其中\d+例").count > 0 Then
                asymptomaticCase = regMatch(0)
            Else
                asymptomaticCase = regMatch(regMatch.count - 1)
            End If
            Set regMatch = Nothing

            If VBA.Err.Number <> 0 Then asymptomaticCase = 0
            VBA.Err.Clear

            If SafeArrayGetDim(byDistrict) > 0 Then
                ReDim Preserve byDistrict(5, UBound(byDistrict, 2) + 1)
            Else
                ReDim byDistrict(5, 0)
            End If
            mUbound = UBound(byDistrict, 2)
            byDistrict(0, mUbound) = releaseDate
            byDistrict(1, mUbound) = currentDistrict
            byDistrict(2, mUbound) = mDate
            byDistrict(3, mUbound) = confirmedCase
            byDistrict(4, mUbound) = asymptomaticCase
            byDistrict(5, mUbound) = ReReplace(myLine, "(.)居住于.*$|(.)分别居住于.*$", "")
        End If

        If currentDistrict = "" Or isExclude(myLine, exclude) Or myLine = "" Then GoTo continue

        '居住地汇总
        mLocation = VBA.Replace(VBA.Replace(location.innerText, "，", ""), "。", "")
        mLocation = VBA.Replace(mLocation, "、", "")
        mLocation = VBA.Replace(mLocation, ",", "")
        mLocation = VBA.Trim(mLocation)
        If isExclude(mLocation, adCodeDescription) Then GoTo continue
        If VBA.Trim(mLocation) = "" Or isExclude(mLocation, exclude) Then GoTo continue
        If SafeArrayGetDim(byLocation) > 0 Then
            ReDim Preserve byLocation(3, UBound(byLocation, 2) + 1)
        Else
            ReDim byLocation(3, 0)
        End If

        mUbound = UBound(byLocation, 2)
        byLocation(0, mUbound) = releaseDate
        byLocation(1, mUbound) = mDate
        byLocation(2, mUbound) = currentDistrict
        byLocation(3, mUbound) = mLocation

        'Debug.Print releaseDate, mDate, currentdistrict, mLocation

        If isHistoryExists Then Exit For

continue:
    Next

    Set htmlDoc = Nothing
    Set Resp = Nothing
End Function

Private Function clearAllSheets()
    sCity.Cells.ClearContents
    sDistrict.Cells.ClearContents
    sCity.Cells.ClearContents
    sLocation.Cells.ClearContents
    sUrl.Cells.ClearContents
End Function

#If VBA7 Then
    Private Function dbToExcel(ByVal db As LongPtr, ByVal tbl As String, ByVal ws As Worksheet)
#Else
    Private Function dbToExcel(ByVal db As Long, ByVal tbl As String, ByVal ws As Worksheet)
#End If
    Dim myArray As Variant

    '将sqlite数据库导出到excel
    Select Case tbl
        Case "by_city"
            On Error GoTo err_handler
            myArray = sqlite3.queryToArray(db, "select DISTINCT * from " & tbl & " order by release_date desc", CP_GB2312, True)
        Case "by_district"
            On Error GoTo err_handler
            myArray = sqlite3.queryToArray(db, "select DISTINCT * from " & tbl & " order by release_date desc", CP_GB2312, True)
        Case "by_location"
            On Error GoTo err_handler
            myArray = sqlite3.queryToArray(db, "select DISTINCT * from " & tbl & " order by release_date desc", CP_GB2312, True)
        Case "url_list"
            On Error GoTo err_handler
            myArray = sqlite3.queryToArray(db, "select DISTINCT * from " & tbl & " order by release_date desc", CP_GB2312, True)
        Case "summary"
            On Error GoTo err_handler
            myArray = sqlite3.queryToArray(db, "select DISTINCT * from " & tbl & " order by release_date desc", CP_GB2312, True)
        Case "poi"
            On Error GoTo err_handler
            myArray = sqlite3.queryToArray(db, "select DISTINCT * from " & tbl & " order by township desc", CP_GB2312, True)
    End Select
    ws.Range(ws.Cells(1, 1), ws.Cells(UBound(myArray, 1) + 1, UBound(myArray, 2) + 1)) = myArray
    Erase myArray
err_handler:
    If Err.Number <> 0 Then
        Debug.Print Err.Description
        Err.Clear
    End If
End Function

Private Function columnNoToLetter(ByVal ColumnNumber As Long) As String
    Dim ColumnLetter As String

    ColumnLetter = VBA.Split(Cells(1, ColumnNumber).Address, "$")(1)
    columnNoToLetter = ColumnLetter
End Function

Private Function rangeText2number(ByVal rng As Range)
    With rng
        .NumberFormat = "General"
        .Value = .Value
    End With
End Function

Private Function amapPoi(ByVal location As String) As Object
    Dim http As New Class_HttpUtils
    Dim Resp As New Class_HttpResponse
    Dim req As New Class_HttpRequest
    Dim url As String
    Dim myArray As Variant
    Dim count As Integer

    Call initDb

    'check whether isExist in database
    count = sqlite3.recordCount(db, "select DISTINCT location from poi where location='" & location & "'")
    If count > 0 Then GoTo final

    With req
        .url = "https://restapi.amap.com/v3/place/text?citylimit=true&output=json&offset=1&city=310000&extensions=all&keywords=" & location & "&key=" & AMAPAPI
        .SetReqHeaderUserAgent "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36 Edg/98.0.1108.56"
        .SetReqHeaderAccept "application/json"
        .SetReqHeaderAcceptLanguage "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6"
    End With

    Set Resp = http.GetReq(req)
    'Debug.Print Resp.body
    If Resp.StatusCd <> "200" Then Exit Function
    If JSON.ParseJson(Resp.body)("status") <> 1 Then Exit Function

    'write result to sqlite3 database
    sqlite3.execSQL db, "insert into poi values(" & "'" & location & "','" & Resp.body & "'" & ");"

final:
    myArray = sqlite3.queryToArray(db, "select * from poi where location=" & "'" & location & "'", CP_GB2312, True)

    'Debug.Print JSON.ConvertToJson(JSON.ParseJson(myArray(1, 1)), Whitespace:=2)
    Set amapPoi = JSON.ParseJson(myArray(1, 1))
End Function
Static Function amapFullPoi(ByVal location As String) As String
    Dim http As New Class_HttpUtils
    Dim Resp As New Class_HttpResponse
    Dim req As New Class_HttpRequest
    Dim url As String
    Dim myArray As Variant
    Dim count As Integer
    Dim lng As Double
    Dim lat As Double
    Dim formattedAddress As String
    Dim township As String
    Dim businessAreas As String
    Dim i As Long

    'check whether isExist in database

    '获取坐标
    Set req = Nothing
    With req
        .url = "http://127.0.0.1:1111/geocode"
        .SetRequestParam "city", "310000"
        .SetRequestParam "address", location
    End With

    Set Resp = http.GetReq(req)
    'Debug.Print Resp.body
    If Resp.StatusCd <> "200" Then GoTo continue
    On Error GoTo continue
    If JSON.ParseJson(Resp.body)("info") <> "OK" Then GoTo continue

    lng = JSON.ParseJson(Resp.body)("geocodes")(1)("location")("lng")
    lat = JSON.ParseJson(Resp.body)("geocodes")(1)("location")("lat")
    formattedAddress = JSON.ParseJson(Resp.body)("geocodes")(1)("formattedAddress")

    '获取商圈和居委会
    Set req = Nothing
    With req
        .url = "http://127.0.0.1:1111/geocode/reverse"
        .SetRequestParam "x", VBA.CStr(lng)
        .SetRequestParam "y", VBA.CStr(lat)
    End With

    Set Resp = http.GetReq(req)
    'Debug.Print Resp.body
    If Resp.StatusCd <> "200" Then GoTo continue
    On Error GoTo continue
    If VBA.StrComp(JSON.ParseJson(Resp.body)("info"), "OK", vbTextCompare) <> 0 Then GoTo continue

    township = JSON.ParseJson(Resp.body)("regeocode")("addressComponent")("township")
    On Error GoTo err_handler
    businessAreas = JSON.ParseJson(Resp.body)("regeocode")("addressComponent")("businessAreas")(1)("name")
err_handler:
    If Err.Number <> 0 Then
        businessAreas = "未知"
        Err.Clear
    End If

    With sLocationXY
        i = .Cells(.Rows.count, 1).End(xlUp).Row
        If i = 1 And VBA.LenB(.Cells(.Rows.count, 1).End(xlUp)) = 0 Then
            .Cells(1, 1) = "location"
            .Cells(1, 2) = "lng"
            .Cells(1, 3) = "lat"
            .Cells(1, 4) = "formattedAddress"
            .Cells(1, 5) = "businessAreas"
            .Cells(1, 6) = "township"
        Else
            .Cells(i + 1, 1) = location
            .Cells(i + 1, 2) = lng
            .Cells(i + 1, 3) = lat
            .Cells(i + 1, 4) = formattedAddress
            .Cells(i + 1, 5) = businessAreas
            .Cells(i + 1, 6) = township
        End If
     End With

    DoEvents
    '写入sqlite3数据库
    amapFullPoi = location & "||" & lng & "||" & lat & "||" & formattedAddress & "||" & businessAreas & "||" & township

continue:
    Set Resp = Nothing
    Set req = Nothing
End Function

Private Function allPoiInit()
    Dim myArray As Variant
    Dim myArray1() As Variant
    Dim myArray2() As Variant
    Dim RetStr() As Variant
    Dim mRecord As Variant
    Dim lineStr As Variant
    Dim results As Variant
    Dim myStr As String
    Dim i As Long

    If db = 0 Then Call initDb
    myArray = sqlite3.queryToArray(db, "select DISTINCT by_location.location from by_location LEFT JOIN poi on by_location.location=poi.location where poi.location IS NULL", CP_GB2312, False)

    If ArrayDimension(myArray) <= 0 Then Exit Function
    For Each results In myArray
        myStr = amapFullPoi(results)
        If VBA.Len(myStr) = 0 Then GoTo continue

        If SafeArrayGetDim(RetStr) > 0 Then
            ReDim Preserve RetStr(UBound(RetStr) + 1)
        Else
            ReDim RetStr(0)
        End If
        RetStr(UBound(RetStr)) = myStr

continue:
    Next results

    If ArrayDimension(RetStr) <= 0 Then Debug.Print "坐标查询失败": Exit Function

    '将查询结果写入sqlite数据库
    sqlite3.execSQL db, "BEGIN TRANSACTION"
    For Each lineStr In RetStr
        mRecord = Split(mRecord, "||")
        On Error GoTo err_handler
        sqlite3.execSQL db, "insert into poi values('" & _
            mRecord(0) & "','" & _
            mRecord(1) & "','" & _
            mRecord(2) & "','" & _
            mRecord(3) & "','" & _
            mRecord(4) & "','" & _
            mRecord(5) & _
            "')"
    Next lineStr
    Erase RetStr
    sqlite3.execSQL db, "COMMIT"

    Dim ws As Worksheet
    Set ws = sLocationXY
    myArray = sqlite3.queryToArray(db, "select DISTINCT * from poi order by township desc", CP_GB2312, True)
    ws.Range(ws.Cells(1, 1), ws.Cells(UBound(myArray, 1) + 1, UBound(myArray, 2) + 1)) = myArray

err_handler:
    If Err.Number <> 0 Then
        sqlite3.execSQL db, "ROLLBACK TRANSACTION"
        Err.Clear
    End If

    Set ws = Nothing
End Function

Private Function locationXY2Db()
    Dim ws As Worksheet
    Dim myArray As Variant
    Dim fields As Variant
    Dim totalRow As Long
    Dim totalCol As Long
    Dim i As Long
    Dim j As Long

    Set ws = sLocationXY

    With ws
        totalRow = .Cells(.Rows.count, 1).End(xlUp).Row
    End With

    If db = 0 Then Call initDb
    sqlite3.execSQL db, "DROP TABLE IF EXISTS poi"
    Call initDb

    '将查询结果写入sqlite数据库
    sqlite3.execSQL db, "BEGIN TRANSACTION"
    For i = 2 To totalRow
        On Error GoTo err_handler
        With ws
        sqlite3.execSQL db, "insert into poi values('" & _
            .Cells(i, 1) & "','" & _
            .Cells(i, 2) & "','" & _
            .Cells(i, 3) & "','" & _
            .Cells(i, 4) & "','" & _
            .Cells(i, 5) & "','" & _
            .Cells(i, 6) & _
            "')"
        End With
    Next i
    sqlite3.execSQL db, "COMMIT"

    'Set ws = Nothing
    'Set ws = sLocationXY
'    fields = Array("location", "lng", "lat", "formattedAddress", "businessAreas", "township")
'
'    myArray = sqlite3.queryToArray(db, "select DISTINCT * from poi order by township desc", CP_GB2312, True)
'    ws.Range(ws.Cells(1, 1), ws.Cells(UBound(myArray, 1) + 1, UBound(myArray, 2) + 1)) = myArray

err_handler:
    If Err.Number <> 0 Then
        sqlite3.execSQL db, "ROLLBACK TRANSACTION"
        sqlite3.closeDB db: db = 0
        Err.Clear
    End If

    Set ws = Nothing
End Function
Private Function initMapChart(ByVal chartName As String)
    Dim mapChart As New Class_Chart
    Dim myChart As ChartObject
    Dim ws As Worksheet
    Dim mapData As Object
    Dim mapArea As Variant
    Dim srs As Series
    Dim i As Long
    Dim j As Integer
    Dim longitude As Variant
    Dim latitude As Variant
    Dim rngLongitude As Range
    Dim rngLatitude As Range
    Dim colLetter As String
    Dim centerList As Variant
    Dim myDataLabel As DataLabel

    Set ws = home
    'add/define new chart
    On Error GoTo err_handler
    mapChart.name = chartName
    If Not mapChart.isChartExists(ws) Then
        Set myChart = mapChart.newChart(ws)
    Else
        'home.ChartObjects(chartName).Delete
        'Set myChart = mapChart.newChart(ws)
        Set myChart = home.ChartObjects(chartName)
    End If

    '清除所有序列
    For i = myChart.Chart.SeriesCollection.count To 1 Step -1
        myChart.Chart.SeriesCollection(i).Delete
    Next

    'add new series for district
    Set mapData = mapDataToJson()
    sMapdata.Cells.ClearContents
    i = 0
    For Each mapArea In mapData
        'put mapArea mapdata to worksheet
        longitude = VBA.Split(mapData(mapArea)("longitude"), ";")
        latitude = VBA.Split(mapData(mapArea)("latitude"), ";")

        With sMapdata
            .Range(.Cells(1, i * 2 + 1), .Cells(1, i * 2 + 1)) = mapArea & "-经度"
            .Range(.Cells(1, i * 2 + 2), .Cells(1, i * 2 + 2)) = mapArea & "-纬度"
            Set rngLongitude = .Range(.Cells(2, i * 2 + 1), .Cells(UBound(longitude) + 2, i * 2 + 1))
            Set rngLatitude = .Range(.Cells(2, i * 2 + 2), .Cells(UBound(latitude) + 2, i * 2 + 2))
        End With
        rngLongitude = Application.Transpose(longitude)
        rngLatitude = Application.Transpose(latitude)
        rangeText2number rngLongitude
        rangeText2number rngLatitude

        With mapChart
            .seriesValueX = "='" & sMapdata.name & "'!" & rngLongitude.Address
            .seriesValue = "='" & sMapdata.name & "'!" & rngLatitude.Address
            .SeriesName = mapArea
            .updateSeries myChart
        End With
        Erase longitude: Erase latitude
        Set rngLongitude = Nothing
        Set rngLatitude = Nothing
        i = i + 1
    Next mapArea

    'formatting chart
    With myChart
        .Chart.SetElement (msoElementChartTitleNone)
        .Chart.SetElement (msoElementLegendNone)
        .Chart.SetElement (msoElementPrimaryCategoryGridLinesNone)
        .Chart.SetElement (msoElementPrimaryValueGridLinesNone)
        .Chart.SetElement (msoElementPrimaryCategoryAxisNone)
        .Chart.SetElement (msoElementPrimaryValueAxisNone)
    End With

    'formatting series
    For j = myChart.Chart.SeriesCollection.count To 1 Step -1
        Set srs = myChart.Chart.SeriesCollection(j)
        With srs
            For Each mapArea In mapData
                If .name = mapArea Then GoTo continue
            Next
            srs.Delete: GoTo continue1

continue:
            If .name = "中心标记" Then
                .chartType = xlXYScatter
                .HasDataLabels = True
                centerList = VBA.Split(mapData(mapArea)("centerList"), ";")
                i = 0
                For Each myDataLabel In .DataLabels
                    myDataLabel.Caption = centerList(i)
                    myDataLabel.Position = xlLabelPositionRight
                    i = i + 1
                Next
            ElseIf VBA.Left(.name, 3) = "上海市" Then
                .IsFiltered = True
            Else
                .chartType = xlXYScatterLinesNoMarkers
            End If
        End With
continue1:
    Next

err_handler:
    If Err.Number <> 0 Then
        MsgBox Err.Description
        Err.Clear
    End If
    Set ws = Nothing
    Set myChart = Nothing
End Function

Private Function initMain()
    Dim urlList As Variant
    Dim mainUri As String
    Dim i As Long
    Dim startTime As Long
    Dim sql As String
    Dim count As Long

    Application.ScreenUpdating = False
    mainUri = "https://wsjkw.sh.gov.cn/yqtb/"

    Call initDb

    'get urllist
    progressShape "开始检查更新..."
    urlList = getEntrance(mainUri)
    If ArrayDimension(urlList) <= 0 Then
        Err.Number = 100
        Err.Description = "没有发现更新"
        GoTo err_handler
    End If

    '将urllist写入sqlite数据库
    progressShape "将urllist写入sqlite数据库..."
    sqlite3.execSQL db, "BEGIN TRANSACTION"
    For i = 0 To UBound(urlList, 2)
        sql = "select * from url_list where release_date = " & "'" & urlList(0, i) & "'"
        If sqlite3.recordCount(db, sql) > 0 Then GoTo continue
        sqlite3.execSQL db, "insert into url_list values(" & "'" & urlList(0, i) & "','" & urlList(1, i) & "');"
        getDetail urlList(1, i), urlList(0, i)
        progressShape "爬取" & urlList(0, i) & " 数据并写入内存数组"
        DoEvents
continue:
    Next i
    sqlite3.execSQL db, "COMMIT"

    progressShape "已完成数据清洗并写入sqlite数据库完成..."
    '判断是否有数据更新
    If SafeArrayGetDim(byCity) = 0 Then
        Err.Number = 100
        Err.Description = "没有发现更新"
        GoTo err_handler
    End If

    '获取地图数据并写入sqlite数据库
    progressShape "获取地图数据并写入sqlite数据库..."
    count = sqlite3.recordCount(db, "select DISTINCT * from mapdata")
    If count = 0 Then
        mapDataToDb db, "mapdata"
        progressShape "在线爬取地图坐标并写入sqlite数据库完成..."
    Else
        progressShape "地图数据已存在于sqlite数据库中，跳过在线爬取..."
    End If

    '清除数据表
    'Call clearAllSheets

    '将居住地POI写入sqlite
    progressShape "将居住地POI写入sqlite..."
    Call locationXY2Db

    '将爬取结果进行数据清洗后写入sqlite数据库
    progressShape "将数据清洗后的爬取结果从内存数组写入sqlite3数据库..."
    sqlite3.execSQL db, "BEGIN TRANSACTION"
    For i = 0 To UBound(byCity, 2)
        On Error GoTo err_handler
        sqlite3.execSQL db, "insert into by_city values(" & "'" & byCity(0, i) & "','" & byCity(1, i) & _
            "','" & byCity(2, i) & "','" & byCity(3, i) & "','" & byCity(4, i) & "'" & ");"
    Next i
    Erase byCity
    sqlite3.execSQL db, "COMMIT"

    sqlite3.execSQL db, "BEGIN TRANSACTION"
    For i = 0 To UBound(byDistrict, 2)
        On Error GoTo err_handler
        sqlite3.execSQL db, "insert into by_district values(" & "'" & byDistrict(0, i) & "','" & byDistrict(1, i) & _
            "','" & byDistrict(2, i) & "','" & byDistrict(3, i) & "','" & byDistrict(4, i) & "','" & byDistrict(5, i) & "'" & ");"
    Next i
    Erase byDistrict
    sqlite3.execSQL db, "COMMIT"

    sqlite3.execSQL db, "BEGIN TRANSACTION"
    For i = 0 To UBound(byLocation, 2)
        On Error GoTo err_handler
        sqlite3.execSQL db, "insert into by_location values(" & "'" & byLocation(0, i) & "','" & byLocation(1, i) & _
            "','" & byLocation(2, i) & "','" & byLocation(3, i) & "'" & ");"
    Next i
    Erase byLocation
    sqlite3.execSQL db, "COMMIT"

    '计算解封条件
    progressShape "计算解封条件..."
    fullPoiSummary (getMax("url_list", "release_date"))

    '创建地图
    progressShape "创建地图..."
    Call GenerateMap

    '显示分区坐标
    progressShape "显示分区坐标..."
    Call categoryXY

    '将sqlite数据库导出到excel
'    dbToExcel db, "by_city", sCity
'    dbToExcel db, "by_district", sDistrict
'    dbToExcel db, "by_location", sLocation
'    dbToExcel db, "url_list", sUrl
'    dbToExcel db, "summary", sSummary

err_handler:
    Select Case Err.Number
        Case 0
        Case 100
            progressShape "没有发现更新"
            Debug.Print Err.Description
        Case Else
            sqlite3.execSQL db, "ROLLBACK TRANSACTION"
    End Select
    Err.Clear

    sqlite3.closeDB db: db = 0
    Application.ScreenUpdating = True

    If shapeExists("statusMsg", home) Then home.Shapes("statusMsg").Delete

End Function
Function initDb()
    appData = Environ("AppData")
    dbFolder = appData & "\" & "shCovid19"
    If Not FolderExists(dbFolder) Then Exit Function
    dbName = dbFolder & "\" & "shcovid.db"

    With db_struct
        .tbl_url_list = "release_date,url"
        .tbl_by_city = "release_date,end_date,confirmed_case,asymptomatic_case,description"
        .tbl_by_district = "release_date,district,end_date,confirmed_case,asymptomatic_case,description"
        .tbl_by_location = "release_date,end_date,district,location"
        .tbl_mapdata = "area,mapdata"
        .tbl_poi = "location,lng,lat,formattedAddress,businessAreas,township"

        If db = 0 Then db = sqlite3.openDB(dbName)

        'sqlite3.execSQL db, "DROP TABLE IF EXISTS poi"

    '初始化sqlite数据库
        sqlite3.execSQL db, sqlite3.createTbl("url_list", .tbl_url_list)
        sqlite3.execSQL db, sqlite3.createTbl("by_city", .tbl_by_city)
        sqlite3.execSQL db, sqlite3.createTbl("by_district", .tbl_by_district)
        sqlite3.execSQL db, sqlite3.createTbl("by_location", .tbl_by_location)
        sqlite3.execSQL db, sqlite3.createTbl("mapdata", .tbl_mapdata)
        sqlite3.execSQL db, sqlite3.createTbl("poi", .tbl_poi)
    End With

End Function

Private Function getMax(ByVal tbl As String, ByVal field As String) As String
    Dim myArray As Variant

    If db = 0 Then Call initDb
    On Error GoTo err_handler
    myArray = sqlite3.queryToArray(db, "select DISTINCT MAX(" & field & ")  from " & tbl, CP_GB2312, False)
    getMax = myArray(0)

err_handler:
    If Err.Number <> 0 Then
        MsgBox Err.Description
    End If
End Function

Private Function getMin(ByVal tbl As String, ByVal field As String) As String
    Dim myArray As Variant

    If db = 0 Then Call initDb
    On Error GoTo err_handler
    myArray = sqlite3.queryToArray(db, "select DISTINCT MIN(" & field & ")  from " & tbl, CP_GB2312, False)
    getMin = myArray(0)

err_handler:
    If Err.Number <> 0 Then
        MsgBox Err.Description
    End If
End Function

Private Function estimatedArea(ByVal theDate As String)
    Dim dDiff As Long

    dDiff = DateDiff("d", theDate, Now)
    If dDiff <= 7 Then
        estimatedArea = "封控区"
    ElseIf dDiff > 7 And dDiff <= 14 Then
        estimatedArea = "管控区"
    Else
        estimatedArea = "防范区"
    End If
End Function

Private Function addMyShape(ByVal myShapeStr As Variant) As Shape
    Dim myShape As Shape

    For Each myShape In home.Shapes
        If myShape.name = myShapeStr Then GoTo final
    Next

    Set myShape = home.Shapes.AddShape(msoShapeRoundedRectangle, 200, 80, 180, 50)
    With myShape
        .name = myShapeStr
        .TextFrame2.TextRange = "等待查询结果"
        .TextFrame2.AutoSize = msoAutoSizeShapeToFitText
    End With
'    home.Shapes.Range(Array(myShapeStr)).ShapeRange.Adjustments.Item(1) = 0.07
    home.Range("c3").Select

final:
    Set addMyShape = myShape
End Function

Private Function deleteMyShape(ByVal myShapeStr As Variant)
    Dim myShape As Shape
    For Each myShape In home.Shapes
        If myShape.name = myShapeStr Then
            myShape.Delete
        End If
    Next
End Function

Private Function colorizedAreaType(ByVal shapeName As String) As Boolean
    Dim areaList As String
    Dim myShape As Shape
    Dim shapeText As String
    Dim matchIndex As Variant
    Dim matchStr As String
    Dim firstIndex As Long

    Set myShape = home.Shapes(shapeName)
    shapeText = myShape.TextFrame2.TextRange
    areaList = "封控区|管控区|防范区"

    For Each matchIndex In RegExpTest(areaList, shapeText)
        firstIndex = VBA.CLng(VBA.Split(matchIndex, "|")(0))
        matchStr = VBA.Split(matchIndex, "|")(1)
        With myShape.TextFrame2.TextRange.Characters(firstIndex + 1, 3).Font.Fill
            Select Case matchStr
                Case "封控区"
                    .ForeColor.RGB = RGB(255, 0, 0) '红色
                Case "管控区"
                    .ForeColor.RGB = RGB(255, 255, 0) '黄色
                Case "防范区"
                    .ForeColor.RGB = RGB(0, 255, 0) '绿色
            End Select
        End With
    Next

    Set myShape = Nothing
End Function

Private Function RegExpTest(patrn, strng, Optional ByVal isIndex As Boolean = True) As Variant
    Dim regEx, Match, Matches
    Dim RetStr() As Variant

    Set regEx = CreateObject("vbscript.regexp")
    regEx.Pattern = patrn
    regEx.IgnoreCase = True
    regEx.Global = True
    Set Matches = regEx.Execute(strng)
    For Each Match In Matches
        If SafeArrayGetDim(RetStr) > 0 Then
            ReDim Preserve RetStr(UBound(RetStr) + 1)
        Else
            ReDim RetStr(0)
        End If
        If isIndex Then
            RetStr(UBound(RetStr)) = Match.firstIndex & "|" & Match.Value
        Else
            RegExpTest = True
            Set regEx = Nothing
            Exit Function
        End If
    Next
    If SafeArrayGetDim(RetStr) = 0 Then RegExpTest = False: Exit Function
    RegExpTest = RetStr
    Set regEx = Nothing
End Function

Private Function removeLocationSeries(ByVal chartName As String)
    Dim myChart As ChartObject
    Dim mySeries As Series
    Dim i As Integer

    Set myChart = home.ChartObjects(chartName)
    For i = myChart.Chart.SeriesCollection.count To 1 Step -1
        Set mySeries = myChart.Chart.SeriesCollection(i)
        On Error Resume Next
        If RegExpTest("封控区|管控区|防范区", mySeries.name, False) Then
            mySeries.Delete
        End If
        'Debug.Print mySeries.name
    Next
    Set myChart = Nothing
End Function

Private Function addLocationSeries(jsonString As String, ByVal myChart As ChartObject)
    Dim myJson As Object
    Dim myJson1 As Object
    Dim longitude As String
    Dim latitude As String
    Dim area As String
    Dim delimiter As String
    Dim myJson2 As Object
    Dim SeriesName As Variant
    Dim chartName As String
    Dim mapChart As New Class_Chart
    Dim mySeries As Series
    Dim sPoint As Point

    Set myJson = JSON.ParseJson(jsonString)
    Set myJson2 = JSON.ParseJson("{}")
    For Each myJson1 In myJson
        Set myJson2(myJson1("估测区域")) = JSON.ParseJson("{}")
    Next myJson1

    For Each myJson1 In myJson
        If myJson2(myJson1("估测区域"))("area") <> "" Then delimiter = ";"
        myJson2(myJson1("估测区域"))("area") = myJson2(myJson1("估测区域"))("area") & delimiter & myJson1("小区名")
        myJson2(myJson1("估测区域"))("longitude") = myJson2(myJson1("估测区域"))("longitude") & delimiter & VBA.Split(myJson1("坐标"), ",")(0)
        myJson2(myJson1("估测区域"))("latitude") = myJson2(myJson1("估测区域"))("latitude") & delimiter & VBA.Split(myJson1("坐标"), ",")(1)
        delimiter = ""
    Next myJson1

    'add/define new chart
    'On Error GoTo err_handler
    For Each SeriesName In myJson2
        With mapChart
            .name = myChart.name
            .seriesValueX = "={" & myJson2(SeriesName)("longitude") & "}"
            .seriesValue = "={" & myJson2(SeriesName)("latitude") & "}"
            .SeriesName = SeriesName
            .updateSeries myChart, xlXYScatter
        End With
    Next

    'formatting series
    For Each mySeries In myChart.Chart.SeriesCollection
        With mySeries
            Select Case .name
                Case "封控区"
                    .MarkerStyle = -4168
                    .MarkerSize = 6
                    With .Format.Fill
                        .ForeColor.RGB = RGB(255, 0, 0)
                    End With
                Case "管控区"
                    .MarkerStyle = 9
                    .MarkerSize = 6
                    With .Format.Fill
                        .ForeColor.RGB = RGB(255, 255, 0)
                    End With
                Case "防范区"
                    .MarkerStyle = 9
                    .MarkerSize = 6
                    With .Format.Fill
                        .ForeColor.RGB = RGB(0, 255, 0)
                    End With
            End Select

        End With
    Next

err_handler:
    If Err.Number <> 0 Then
        Debug.Print Err.Description
    End If
End Function

Private Function inquiryByArray(ByVal myArray As Variant)
    Dim myChart As ChartObject
    Dim myLocation As Variant
    Dim myShape As Shape
    Dim myJson As Object
    Dim myJSONArray() As Object
    Dim sql As String
    Dim myArray1 As Variant
    Dim myArray2() As Variant
    Dim count As Long
    Dim i As Long
    Dim poiJSON As Object
    Dim ws As Worksheet
    Dim mapChart As New Class_Chart

    Call initDb
    Set ws = home
    'add/define new chart
    'On Error Resume Next
    mapChart.name = "mapChart"
    If Not mapChart.isChartExists(ws) Then
        initMapChart "mapChart"
    End If

    Set myShape = addMyShape("info")
    Set myChart = home.ChartObjects("mapChart")

    For Each myLocation In myArray
        If myLocation = "" Then GoTo continue
        If SafeArrayGetDim(myJSONArray) > 0 Then
            ReDim Preserve myJSONArray(UBound(myJSONArray) + 1)
        Else
            ReDim myJSONArray(0)
        End If

        sql = "select DISTINCT * from by_location where location='" & myLocation & "' order by end_date desc"
        count = sqlite3.recordCount(db, sql)
        If count = 0 Then
            Err.Number = 100
            Err.Description = "未查询到有效数据， 请重新输入"
            GoTo err_handler
        End If
        myArray1 = sqlite3.queryToArray(db, sql, CP_GB2312, False)
        Set myJson = JSON.ParseJson("{}")
        For i = 0 To UBound(myArray1, 1)
            If SafeArrayGetDim(myArray2) > 0 Then
                ReDim Preserve myArray2(UBound(myArray2) + 1)
            Else
                ReDim myArray2(0)
            End If
            myArray2(UBound(myArray2)) = myArray1(i, 0)
        Next i

        'get POI info
        On Error GoTo continue1
        Set poiJSON = amapPoi(myArray1(0, 3))
        myJson("小区名") = poiJSON("pois")(1)("name")
        On Error Resume Next
        myJson("所在社区") = poiJSON("pois")(1)("business_area") & "社区"
        myJson("坐标") = poiJSON("pois")(1)("location")

continue1:
        myJson("居住地") = myArray1(0, 3)
        myJson("估测区域") = estimatedArea(myArray1(0, 1))
        myJson("预计解封") = VBA.Format(VBA.DateAdd("d", 14, myArray1(0, 1)), "yyyy-mm-dd")
        myJson("所在区县") = myArray1(0, 2)
        myJson("阳性日期") = myArray2

        Set myJSONArray(UBound(myJSONArray)) = myJson
        Erase myArray1
        Erase myArray2
        Set myJson = Nothing
continue:
    Next myLocation

    If JSON.ConvertToJson(myJSONArray) = "[]" Then
        Err.Number = 100
        Err.Description = "未查询到有效数据， 请重新输入"
        GoTo err_handler
    End If

final:
    addLocationSeries JSON.ConvertToJson(myJSONArray), myChart
    myShape.TextFrame2.TextRange = JSON.ConvertToJson(myJSONArray, Whitespace:=2)
    myShape.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
    colorizedAreaType ("info")

err_handler:
    If Err.Number <> 0 Then
        myShape.TextFrame2.TextRange = Err.Description
        'home.Range("c3") = ""
        home.Range("c3").Select
    End If

    Set myShape = Nothing
    Set myChart = Nothing
End Function

Private Function fullPoiSummary(ByVal mDate As String, Optional ByVal nDate As String)
    Dim sql As String
    Dim criteria As String
    Dim myArray As Variant

    If db = 0 Then Call initDb

    If VBA.Len(nDate) <> 0 Then
        criteria = "where release_date='" & mDate & "'"
    Else
        criteria = "where release_date<='" & mDate & "'"
    End If

    '模糊查询
    sql = "select DISTINCT release_date,location,district,max(end_date) as latest_date," & _
        "JULIANDAY('" & mDate & "') - JULIANDAY(max(end_date)) as dDiff," & _
        "iif(JULIANDAY('" & mDate & "') - JULIANDAY(max(end_date))<7,'封控区',iif(JULIANDAY('" & mDate & "') - JULIANDAY(max(end_date)) <14, '管控区', '防范区')) as category," & _
        "date(max(end_date),'+14 day') as dUnlockdown," & _
        "iif(JULIANDAY(date(max(end_date),'+14 day')) - JULIANDAY('" & mDate & "')<0,'无',JULIANDAY(date(max(end_date),'+14 day')) - JULIANDAY('" & mDate & "')) as dDiffUnlockdown " & _
        "from by_location " & criteria & " group by location order by district asc"

    'create temp table
    sqlite3.execSQL db, "DROP TABLE IF EXISTS summary1"
    On Error GoTo err_handler
    sqlite3.execSQL db, "CREATE TEMPORARY TABLE IF NOT EXISTS summary1 AS with FT_CTE AS (" & sql & ") SELECT * FROM FT_CTE"

    sql = "SELECT DISTINCT summary1.*,poi.lng,poi.lat,poi.formattedAddress,poi.businessAreas,poi.township " & _
        "from summary1 LEFT OUTER JOIN poi on summary1.location = poi.location"
    sqlite3.execSQL db, "DROP TABLE IF EXISTS summary"
    On Error GoTo err_handler
    sqlite3.execSQL db, "CREATE TABLE IF NOT EXISTS summary AS with FT_CTE AS (" & sql & ") SELECT * FROM FT_CTE"

'    sql = "SELECT DISTINCT * from summary"
'    On Error GoTo err_handler
'    myArray = sqlite3.queryToArray(db, sql, CP_GB2312, True)
'
'    With sSummary
'        .Cells.Clear
'        .Range(.Cells(1, 1), .Cells(UBound(myArray, 1) + 1, UBound(myArray, 2) + 1)) = myArray
'    End With

    'sqlite3.closeDB db: db = 0

err_handler:
    If Err.Number <> 0 Then
        sqlite3.closeDB db: db = 0
        Debug.Print Err.Description
    End If
End Function

Private Sub inquiryByLocation(ByVal rng As Range)
    Dim sql As String
    Dim myArray As Variant
    Dim count As Long
    Dim myShape As Shape
    Dim keywords As String
    Dim i As Long
    Dim v As Variant

    'Debug.Print ActiveSheet.Buttons(Application.Caller).name
'    Debug.Print TypeName(Application.Caller)
'    Select Case TypeName(Application.Caller)
'        Case "Range"
'           v = Application.Caller.Address
'        Case "String"
'           v = Application.Caller
'        Case "Error"
'           v = "Error"
'        Case Else
'           v = "unknown"
'    End Select
'    Debug.Print "caller = " & v

    If Not FileExists(dbName) Then
        Call initMain
    End If

    Call initDb

    keywords = rng.Value

    'Set myShape = addMyShape("info")
    If VBA.Trim(keywords) = "" Then
        Err.Number = 100
        Err.Description = "未查询到有效数据， 请重新输入"
        GoTo err_handler
    End If

    '模糊查询
    sql = "select DISTINCT location from summary"
    On Error GoTo err_handler
    myArray = sqlite3.queryToArray(db, sql, CP_GB2312, False)

    If ArrayDimension(myArray) <= 0 Then
        Err.Number = 100
        Err.Description = "未查询到有效数据"
        GoTo err_handler
    End If

    '构建结果临时表
    sqlite3.execSQL db, "DROP TABLE IF EXISTS tempTbl"
    sqlite3.execSQL db, "CREATE TABLE IF NOT EXISTS tempTbl(location);"
    For i = 0 To UBound(myArray)
        If VBA.InStr(1, myArray(i), VBA.Trim(keywords)) > 0 Then
            sqlite3.execSQL db, "INSERT INTO tempTbl values('" & myArray(i) & "');"
        End If
    Next i
    Erase myArray

    '显示模糊查询结果，提示用户选择
    'sql = "SELECT DISTINCT summary1.*,poi.lng,poi.lat,poi.formattedAddress,poi.businessAreas,poi.township " & _
        "from summary1 INNER JOIN poi where summary1.location = poi.location"
    sql = "SELECT DISTINCT summary.* from summary INNER JOIN tempTbl on summary.location = tempTbl.location"

    myArray = sqlite3.queryToArray(db, sql, CP_GB2312, True)
    sqlite3.execSQL db, "DROP TABLE IF EXISTS tempTbl;"

    '将查询结果写入listView
    'Dim ws As Worksheet: Set ws = Sheet1
    'ws.Range(ws.Cells(1, 1), ws.Cells(UBound(myArray, 1) + 1, UBound(myArray, 2) + 1)) = myArray

    If ArrayDimension(myArray) <= 0 Then
        Err.Number = 100
        Err.Description = "未查询到有效数据"
        GoTo err_handler
    End If

    If UBound(myArray) >= 0 Then
        For i = 0 To UBound(myArray, 2)
            myArray(0, i) = headerChinese(myArray(0, i))
        Next i
        userformVar = myArray
        Load UserForm1
        UserForm1.Show
        Exit Sub
    End If

err_handler:
    If Err.Number <> 0 Then
        MsgBox Err.Description
        'myShape.TextFrame2.TextRange = Err.Description
        home.Range("c3").Select
    End If

    'Set myShape = Nothing
End Sub

Private Function zoomChart(Optional ByVal inOut As Integer = 0)
    Dim myChart As ChartObject
    Dim axesLen1 As Double
    Dim axesLen2 As Double

    On Error GoTo err_handler
    Set myChart = home.ChartObjects("mapChart")
    With myChart.Chart.Axes(xlValue)
        axesLen1 = (.MaximumScale - .MinimumScale) / 100
    End With

    With myChart.Chart.Axes(xlCategory)
        axesLen2 = (.MaximumScale - .MinimumScale) / 100
    End With

    Select Case inOut
        Case 0
            With myChart.Chart.Axes(xlValue)
                .MinimumScale = .MinimumScale - axesLen1
                .MaximumScale = .MaximumScale + axesLen1
            End With

            With myChart.Chart.Axes(xlCategory)
                .MinimumScale = .MinimumScale - axesLen2
                .MaximumScale = .MaximumScale + axesLen2
            End With
        Case 1
            With myChart.Chart.Axes(xlValue)
                    .MinimumScale = .MinimumScale + axesLen1
                    .MaximumScale = .MaximumScale - axesLen1
            End With

            With myChart.Chart.Axes(xlCategory)
                    .MinimumScale = .MinimumScale + axesLen2
                    .MaximumScale = .MaximumScale - axesLen2
            End With
        Case 2  'right
            With myChart.Chart.Axes(xlCategory)
                    .MinimumScale = .MinimumScale - axesLen1
                    .MaximumScale = .MaximumScale - axesLen1
            End With
        Case 3  'left
            With myChart.Chart.Axes(xlCategory)
                    .MinimumScale = .MinimumScale + axesLen1
                    .MaximumScale = .MaximumScale + axesLen1
            End With

        Case 4  'up
            With myChart.Chart.Axes(xlValue)
                    .MinimumScale = .MinimumScale + axesLen1
                    .MaximumScale = .MaximumScale + axesLen1
            End With
        Case 5  'down
            With myChart.Chart.Axes(xlValue)
                    .MinimumScale = .MinimumScale - axesLen1
                    .MaximumScale = .MaximumScale - axesLen1
            End With
        Case 6  'reset
            With myChart.Chart.Axes(xlValue)
                .MinimumScaleIsAuto = True
                .MaximumScaleIsAuto = True
            End With

            With myChart.Chart.Axes(xlCategory)
                .MinimumScaleIsAuto = True
                .MaximumScaleIsAuto = True
            End With
        End Select
    Set myChart = Nothing
err_handler:
    If Err.Number <> 0 Then
        MsgBox Err.Description
    End If
End Function

Private Function GenerateMap()
    Dim myChart As ChartObject

    Application.ScreenUpdating = False
    initMapChart "mapChart"

    Set myChart = home.ChartObjects("mapChart")

    With myChart.Chart
        .SetElement (msoElementChartTitleAboveChart)
        .ChartTitle.Text = "上海市新冠疫情分区防控图"
    End With

    Call categoryXY
    Call summaryChart
    Call addShapeInfo
    Call initTrendChart
    home.Shapes("mapChart").ZOrder msoSendToBack
    Application.ScreenUpdating = True
    DoEvents
End Function

Private Function headerChinese(ByVal headerEn As String) As String
    Dim myHeader As Variant
    Dim i As Integer

    myHeader = VBA.Split(sHeader, ",")
    For i = 0 To UBound(myHeader)
        If VBA.StrComp(myHeader(i), headerEn) = 0 Then
            headerChinese = VBA.Split(sHeaderCN, ",")(i)
            Exit Function
        End If
    Next i
End Function

Private Function categoryXY()
    Dim categoryArray As Variant
    Dim myArray1 As Variant
    Dim myArray2 As Variant
    Dim i As Integer
    Dim iChart As New Class_Chart
    Dim myChart As ChartObject
    Dim X As Range
    Dim Y As Range
    Dim srs As Series

    If db = 0 Then Call initDb

    On Error GoTo err_handler
    categoryArray = sqlite3.queryToArray(db, "SELECT DISTINCT category from summary", CP_GB2312, False)
    If ArrayDimension(categoryArray) <= 0 Then
        On Error GoTo err_handler
        Exit Function
    End If

    sCategoryXY.Cells.Clear
    For i = 0 To UBound(categoryArray)
        On Error GoTo err_handler
        myArray1 = sqlite3.queryToArray(db, "SELECT lng  from summary where category='" & categoryArray(i) & "' and lng<>'n/a'", CP_GB2312, True)
        myArray2 = sqlite3.queryToArray(db, "SELECT lat  from summary where category='" & categoryArray(i) & "' and lng<>'n/a'", CP_GB2312, True)
        myArray1(0) = categoryArray(i) & "-" & myArray1(0)
        myArray2(0) = categoryArray(i) & "-" & myArray2(0)
        With sCategoryXY
            '将坐标导入excel
            .Range(.Cells(1, 1 + i * 2), .Cells(UBound(myArray1) + 1, 1 + i * 2)) = Application.Transpose(myArray1)
            .Range(.Cells(1, 2 + i * 2), .Cells(UBound(myArray2) + 1, 2 + i * 2)) = Application.Transpose(myArray2)

            '添加到图表
            Set X = .Range(.Cells(2, 1 + i * 2), .Cells(UBound(myArray1) + 1, 1 + i * 2))
            Set Y = .Range(.Cells(2, 2 + i * 2), .Cells(UBound(myArray2) + 1, 2 + i * 2))
            Set myChart = home.ChartObjects("mapChart")
            With iChart
                .SeriesName = categoryArray(i)
                .seriesValueX = "='" & sCategoryXY.name & "'!" & X.Address
                .seriesValue = "='" & sCategoryXY.name & "'!" & Y.Address
                .updateSeries myChart, xlXYScatter
            End With
        End With

        Erase myArray1
        Erase myArray2
    Next i

    'formatting series
    On Error GoTo err_handler
    myChart.Chart.ChartTitle.Caption = "上海市新冠疫情分区防控图" & "(" & getMax("summary", "release_date") & ")"
     For Each srs In myChart.Chart.SeriesCollection
         With srs
            Select Case .name
                Case "封控区"
                    .HasDataLabels = False
                    .MarkerSize = 3
                    .MarkerStyle = 8
                    .Format.Line.Visible = msoFalse
                    With .Format.Fill
                        .ForeColor.RGB = RGB(244, 13, 100) '红色
                    End With
                Case "管控区"
                    .HasDataLabels = False
                    .MarkerSize = 3
                    .MarkerStyle = 8
                    .Format.Line.Visible = msoFalse
                    With .Format.Fill
                        .ForeColor.RGB = RGB(244, 208, 0) '黄色
                    End With
                Case "防范区"
                    .HasDataLabels = False
                    .MarkerSize = 3
                    .MarkerStyle = 8
                    .Format.Line.Visible = msoFalse
                    With .Format.Fill
                        .ForeColor.RGB = RGB(64, 116, 52)  '绿色
                    End With
            End Select
         End With
     Next
err_handler:
    If Err.Number <> 0 Then
        MsgBox Err.Description
    End If

    Set X = Nothing
    Set Y = Nothing
    Set iChart = Nothing

End Function

Private Function summaryChart()
    Dim sql As String
    Dim myArray As Variant
    Dim ws As Worksheet
    Dim myChart As ChartObject
    Dim iChart As New Class_Chart
    Dim X As String
    Dim Y As String
    Dim i As Integer
    Dim delimiter As String
    Dim fsc As SeriesCollection
    Dim srs As Series
    Dim totalCounts As Long

    If db = 0 Then Call initDb
    sql = "select DISTINCT category,count(location) from summary group by category"
    myArray = sqlite3.queryToArray(db, sql, CP_GB2312, False)

    If ArrayDimension(myArray) <= 0 Then
        Err.Number = 100
        Err.Description = "查询失败"
        GoTo err_handler
    End If

    For i = 0 To UBound(myArray, 1)
        If i = 0 Then delimiter = "" Else delimiter = ","
        X = X & delimiter & Chr(34) & myArray(i, 0) & Chr(34)
        Y = Y & delimiter & myArray(i, 1)
    Next i

    '添加扇形图
    Set ws = home
    Set myChart = createNewSummaryChart("summary", ws)
    myChart.Chart.chartType = xlColumnStacked

    '清空
    Set fsc = myChart.Chart.SeriesCollection
    For i = fsc.count To 1 Step -1
        fsc(i).Delete
    Next i

    '添加数据

'    With srs
'        .name = "分区防控占比"
'        .XValues = "={" & X & "}"
'        .Values = "={" & Y & "}"
'    End With
    totalCounts = sqlite3.recordCount(db, "select DISTINCT location from summary")
    For i = 0 To UBound(myArray, 1)
        Set srs = myChart.Chart.SeriesCollection.NewSeries
        With srs
            .name = myArray(i, 0)
            .XValues = "={" & Chr(34) & "合计共有" & totalCounts & "个居住地/小区受影响" & Chr(34) & "}"
            .Values = "={" & myArray(i, 1) & "}"
        End With
    Next i

    '格式化
    With ws.Shapes(myChart.name)
        .Line.Visible = msoFalse '隐藏边框
        .Fill.Visible = msoFalse
    End With

    With myChart.Chart

        '.SetElement (msoElementPrimaryCategoryAxisNone)
        .SetElement (msoElementPrimaryValueAxisNone)
        .SetElement (msoElementLegendNone) '不显示图例
        .SetElement (msoElementPrimaryValueGridLinesNone)
        .HasTitle = False
        '.ChartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255) 'chart title color
        With .PlotArea
            '.ClearFormats
            .Top = 0
            .InsideLeft = 0
            .InsideWidth = 100
            .InsideHeight = 280
'            .ClearFormats
            .Format.Fill.Visible = msoFalse
        End With
        .ChartArea.Format.Fill.Visible = msoFalse
        .ChartGroups(1).GapWidth = 0
        With .Axes(xlCategory)
            .TickLabels.Font.Bold = msoTrue
            .TickLabels.Font.Color = RGB(255, 255, 255)
            .Format.Fill.ForeColor.RGB = RGB(50, 50, 50)
        End With
        'On Error GoTo err_handler
        .Axes(xlValue).MaximumScale = sqlite3.recordCount(db, "select DISTINCT location from by_location")
    End With

    For i = fsc.count To 1 Step -1
        With fsc(i)
            .HasDataLabels = True
            'On Error Resume Next
            .DataLabels.ShowCategoryName = False
            '.DataLabels.ShowPercentage = True
            .DataLabels.Position = xlLabelPositionCenter
            .DataLabels.ShowSeriesName = True
            With .DataLabels.Format.TextFrame2.TextRange.Font.Fill
                .Visible = msoTrue
                .ForeColor.RGB = RGB(255, 255, 255)
            End With
            .DataLabels.Format.TextFrame2.WordWrap = msoFalse   '禁用标签文字自动换行

            With .Format.Fill
                Select Case fsc(i).name
                    Case "封控区"
                        .ForeColor.RGB = RGB(244, 13, 100) '红色
                    Case "管控区"
                        .ForeColor.RGB = RGB(244, 208, 0) '黄色
                    Case "防范区"
                        .ForeColor.RGB = RGB(64, 116, 52) '绿色
                End Select
            End With
        End With
    Next i

err_handler:
    If Err.Number <> 0 Then MsgBox Err.Description

    Set ws = Nothing
    Set srs = Nothing
    Set myChart = Nothing
    Err.Clear
End Function

Private Function addShapeInfo()
    Dim myShapeStr As String
    Dim myShapeName As String
    Dim sql As String
    Dim mDate As String
    Dim myShape As Shape

    myShapeName = "infoDescription"
    If db = 0 Then Call initDb
    mDate = getMax("summary", "release_date")
    sql = "select description from by_city where release_date = '" & mDate & "'"
    On Error GoTo err_handler
    myShapeStr = sqlite3.queryToArray(db, sql, CP_GB2312, False)(0)

    For Each myShape In home.Shapes
        If myShape.name = myShapeName Then
            Set myShape = home.Shapes(myShapeName)
            GoTo final
            Exit For
        End If
    Next
    Set myShape = home.Shapes.AddShape(msoShapeRoundedRectangle, 700, 50, 450, 70)
    myShape.name = myShapeName

final:
    With myShape
        .TextFrame2.TextRange = myShapeStr
        .Fill.Transparency = 0.85
        .TextFrame2.AutoSize = msoAutoSizeShapeToFitText
    End With

err_handler:
    If Err.Number <> 0 Then
        Debug.Print Err.Description
    End If
End Function

Private Function historicalStreamChart()
    Dim sql As String
    Dim mDate As Variant
    Dim myArray As Variant
    Dim myChart As ChartObject

    '计算解封条件
    If db = 0 Then Call initDb

    sql = "SELECT DISTINCT release_date from url_list order by release_date asc"
    myArray = sqlite3.queryToArray(db, sql, CP_GB2312, False)

    For Each mDate In myArray
        fullPoiSummary mDate
        DoEvents
        Call GenerateMap
        Set myChart = home.ChartObjects("mapChart")
'        With home.Shapes("shpDate").TextFrame2.TextRange
'            .Text = mDate
'            .Font.Size = 20
'        End With
        'myChart.Chart.Refresh
        DoEvents
        Sleep 1
    Next
End Function

Private Function Repeat$(ByVal n&, s$)
    Dim r&
    r = Len(s)
    If n < 1 Then Exit Function
    If r = 0 Then Exit Function
    If r = 1 Then Repeat = String$(n, s): Exit Function
    Repeat = Space$(n * r)
    Mid$(Repeat, 1) = s: If n > 1 Then Mid$(Repeat, r + 1) = Repeat
End Function

Private Function formatChart(ByVal myChart As ChartObject, Optional ByVal stage As Integer = 0)
    '格式化
    With myChart.Parent.Shapes(myChart.name)
        .Line.Visible = msoFalse '隐藏边框
        .Fill.Visible = msoFalse
    End With
    Select Case stage
        Case 0
            With myChart.Chart
                On Error Resume Next
                .ChartArea.ClearContents
            End With
            With myChart.Chart
                .chartType = xlBubble3DEffect
                .ClearToMatchStyle
                .ChartStyle = 233
                With .Axes(xlValue)
                        .TickLabels.Font.Size = 12
                End With
                With .Axes(xlCategory)
                    .MinorUnit = 1
                    .MajorUnit = 1
                    .TickLabels.NumberFormatLocal = "mm-dd"
                    .TickLabels.Font.Size = 12
                    .TickLabelPosition = xlLow
                End With

            End With
        Case 1
            With myChart.Chart
                    .SetElement (msoElementChartTitleNone)
                    .SetElement (msoElementPrimaryValueAxisNone)
                    .SetElement (msoElementPrimaryValueGridLinesNone)
                    '.SetElement (msoElementPrimaryCategoryAxisNone)
                    .SetElement (msoElementPrimaryCategoryGridLinesNone)
                    .SetElement (msoElementLegendRight)
                    On Error Resume Next
                    .SetElement (msoElementPlotAreaShow)
                    .SetElement (msoElementDataLabelCenter)
                    .Legend.Format.TextFrame2.TextRange.Font.Size = 14
                    With .PlotArea
                        .Top = 0
                        .InsideLeft = 0
                        .InsideWidth = 800
                        '.Interior.ColorIndex = 8
                    End With
                '格式化
                    With .Axes(xlCategory)
                        .TickLabels.Font.Size = 16
                        .Format.Line.Visible = msoFalse
                    End With
            End With
    End Select
    Err.Clear
End Function

Private Function createNewChart(ByVal chartName As String, ByVal ws As Worksheet) As ChartObject
    Dim myChart1 As ChartObject
    Dim myChart As ChartObject

    'create new trend chart
    For Each myChart In ws.ChartObjects
        If VBA.StrComp(myChart.name, chartName, vbBinaryCompare) = 0 Then
            Set myChart1 = myChart
            GoTo continue
        End If
continue:
    Next myChart
    If myChart1 Is Nothing Then
        Set myChart1 = ws.ChartObjects.add(Left:=1200, Width:=800, Top:=315, Height:=300)
        myChart1.name = chartName
    End If

    Set createNewChart = myChart1
    Set myChart1 = Nothing
End Function

Private Function createNewSummaryChart(ByVal chartName As String, ByVal ws As Worksheet) As ChartObject
    Dim myChart1 As ChartObject
    Dim myChart As ChartObject

    'create new trend chart
    For Each myChart In ws.ChartObjects
        If VBA.StrComp(myChart.name, chartName, vbBinaryCompare) = 0 Then
            Set myChart1 = myChart
            GoTo continue
        End If
continue:
    Next myChart
    If myChart1 Is Nothing Then
        Set myChart1 = ws.ChartObjects.add(Left:=400, Width:=100, Top:=15, Height:=300)
        myChart1.name = chartName
    End If

    Set createNewSummaryChart = myChart1
    Set myChart1 = Nothing
End Function

Private Function setBubbleSeries(ByVal mySeries As Series, _
    ByVal SeriesName As String, _
    ByVal X As String, _
    ByVal Y As String, _
    ByVal bubbleSize As String, _
    Optional ByVal chartType As Integer = xlBubble3DEffect)

    With mySeries
        .chartType = chartType
        .MarkerStyle = 8
        .MarkerSize = 24
        .name = SeriesName
        .XValues = X
        .Values = Y
        .Format.Line.Weight = 20
        .BubbleSizes = bubbleSize
        .DataLabels.Format.TextFrame2.TextRange.Font.Size = 12
        .DataLabels.Format.TextFrame2.TextRange.Font.Bold = True
    End With
End Function

Private Function setXYSeries(ByVal mySeries As Series, _
    ByVal SeriesName As String, _
    ByVal X As String, _
    ByVal Y As String, _
    Optional ByVal chartType As Integer = xlXYScatter)

    With mySeries
        .chartType = chartType
        .MarkerStyle = 8
        .MarkerSize = 24
        .name = SeriesName
        .XValues = X
        .Values = Y
        On Error Resume Next
        .HasLeaderLines = False
        .Format.Fill.Visible = msoFalse
        With .Format.Line
            .Weight = 20
            .Visible = msoFalse
        End With
        'On Error Resume Next
        With .DataLabels
            .ShowSeriesName = True
            .ShowValue = False
            .Position = xlLabelPositionRight
            .Font.Size = 14
        End With
        '.Parent.SetElement (msoElementDataLabelRight)
'        .DataLabels.Format.TextFrame2.TextRange.Font.Size = 12
'        .DataLabels.Format.TextFrame2.TextRange.Font.Bold = True
    End With
End Function

Private Function initTrendChart()
    Dim i As Integer
    Dim ws As Worksheet
    Dim chartName1 As String
    Dim chartName2 As String
    Dim myChart1 As ChartObject
    Dim myChart2 As ChartObject
    Dim seriesNames As String
    Dim mySeries As Series
    Dim myArray As Variant
    Dim sql As String
    Dim X As String
    Dim Y1 As String
    Dim Y2 As String
    Dim Y3 As String
    Dim delimiter As String
    Dim mDate As String
    Dim max As Long
    Dim totalCase As Long

    chartName1 = "trendChart1"
    chartName2 = "trendChart2"
    Set ws = home

    'get data
    If db = 0 Then Call initDb
    mDate = getMax("summary", "release_date")
    max = VBA.CLng(getMax("by_city", "cast(asymptomatic_case as int)"))

    'accumulated total_case
    sql = "select DISTINCT release_date,end_date,confirmed_case,asymptomatic_case,confirmed_case+asymptomatic_case as total_case, null as accumulated_case, description from by_city order by end_date asc"
    sqlite3.execSQL db, "CREATE TEMPORARY TABLE IF NOT EXISTS by_city_tmp AS with FT_CTE AS (" & sql & ") SELECT * FROM FT_CTE"

    sql = "update by_city_tmp " & _
        "set accumulated_case=" & _
        "ifnull(" & _
            "(SELECT ifnull(accumulated_case,0) FROM by_city_tmp ROWPRIOR " & _
                "WHERE ROWPRIOR.rowid = (by_city_tmp.rowid - 1 ) " & _
            "), " & _
        "0) +  " & _
        "total_case"
    sqlite3.execSQL db, sql

    sql = "select DISTINCT * from by_city_tmp where release_date<='" & mDate & "' order by end_date desc limit 10"
    myArray = sqlite3.queryToArray(db, sql, CP_GB2312, False)
'    Sheet1.Cells.Clear
'    Sheet1.Range(Sheet1.Cells(1, 1), Sheet1.Cells(UBound(myArray, 1) + 1, UBound(myArray, 2) + 1)) = myArray

    sqlite3.execSQL db, "DROP TABLE IF EXISTS by_city_tmp"

    If ArrayDimension(myArray) <= 0 Then
        Err.Number = 100
        Err.Description = "查询失败"
        GoTo err_handler
    End If

    For i = UBound(myArray, 1) To 0 Step -1
        If i = UBound(myArray, 1) Then delimiter = "" Else delimiter = ","
        X = X & delimiter & VBA.Format(myArray(i, 1), 0) 'end_date
        Y1 = Y1 & delimiter & myArray(i, 2) '确诊病例
        Y2 = Y2 & delimiter & myArray(i, 3) '无症状
        Y3 = Y3 & delimiter & myArray(i, 5) '累计值
    Next i

    'add new charts
    Set myChart1 = createNewChart(chartName1, ws)
    With myChart1
        .Left = 1200
        .Width = 800
        .Top = 315
        .Height = 300
    End With
    Set myChart2 = createNewChart(chartName2, ws)
    With myChart2
        .Left = 1200
        .Width = 800
        .Top = 15
        .Height = 300
    End With

    'add series for myChart1
    formatChart myChart1
    With myChart1.Chart
        With .Axes(xlValue)
                .MinimumScale = -max * 0.3
                .MaximumScale = max * 1.3
        End With
        With .Axes(xlCategory)
            .MinimumScale = VBA.Format(myArray(UBound(myArray, 1), 1), 0)
            .MaximumScale = .MinimumScale + 10
        End With
    End With

    'add series for myChart2
    formatChart myChart2
    With myChart2.Chart
        With .Axes(xlValue)
'                .MinimumScale = -max * 0.3
'                .MaximumScale = max * 1.3
                .TickLabels.Font.Size = 12
        End With
        With .Axes(xlCategory)
            .MinimumScale = VBA.Format(myArray(UBound(myArray, 1), 1), 0)
            .MaximumScale = .MinimumScale + 10
            .TickLabelPosition = xlLow
        End With
    End With

    'add data series
    '无症状病例
    Set mySeries = myChart1.Chart.SeriesCollection.NewSeries
    formatChart myChart1, 1
    With myChart1.Chart
        .SetElement (msoElementLegendNone)
    End With
    setBubbleSeries mySeries, "=" & Chr(34) & "无症状感染者" & Chr(34), "={" & X & "}", "={ " & Y2 & "}", "={ " & Y2 & "}", xlBubble3DEffect
    '添加图列 - 无症状病例
    Set mySeries = myChart1.Chart.SeriesCollection.NewSeries
    setXYSeries mySeries, "=" & Chr(34) & "无症状感染者" & Chr(34), "={" & VBA.Format(myArray(0, 1), "0") & "}", "={" & myArray(0, 3) & "}", xlXYScatterSmooth

    '确诊病例
    Set mySeries = myChart1.Chart.SeriesCollection.NewSeries
    'formatChart myChart1, 1
    With myChart1.Chart
        .SetElement (msoElementLegendNone)
    End With
    setBubbleSeries mySeries, "=" & Chr(34) & "确诊感染者" & Chr(34), "={" & X & "}", "={ " & Y1 & "}", "={ " & Y1 & "}", xlBubble3DEffect

    '添加图列 - 确诊病例
    Set mySeries = myChart1.Chart.SeriesCollection.NewSeries
    setXYSeries mySeries, "=" & Chr(34) & "确诊感染者" & Chr(34), "={" & VBA.Format(myArray(0, 1), "0") & "}", "={" & myArray(0, 2) & "}", xlXYScatterSmooth
    With myChart1.Chart
        .SetElement (msoElementLegendNone)
    End With

    '累计病例
    Set mySeries = myChart2.Chart.SeriesCollection.NewSeries
    formatChart myChart2, 1
    With myChart2.Chart
        .SetElement (msoElementPrimaryCategoryAxisNone)
        .SetElement (msoElementLegendNone)
        .ChartColor = 13
    End With
    setBubbleSeries mySeries, "=" & Chr(34) & "累计感染者" & Chr(34), "={" & X & "}", "={ " & Y3 & "}", "={ " & Y3 & "}", xlBubble
    '添加图列 - 累计病例
    Set mySeries = myChart2.Chart.SeriesCollection.NewSeries
    setXYSeries mySeries, "=" & Chr(34) & "累计感染者" & Chr(34), "={" & VBA.Format(myArray(0, 1), "0") & "}", "={" & myArray(0, 5) & "}", xlXYScatterSmooth

    Set mySeries = Nothing
    Set myChart1 = Nothing
    Set myChart2 = Nothing
    Err.Clear
err_handler:
    If Err.Number <> 0 Then
        Debug.Print Err.Description
    End If

End Function

Private Function exportToExcel()
    Dim wk As Workbook
    Dim ws As Worksheet
    Dim wsList As String
    Dim wsListName As String
    Dim tbl As Variant
    Dim i As Integer

    If db = 0 Then Call initDb

    wsList = "by_city,by_district,by_location,summary,url_list,poi"
    wsListName = "全市,各区,居住地明细,居住地汇总,数据源,兴趣点坐标库"
    Set wk = Application.Workbooks.add
    wk.Windows(1).Visible = False

    '将sqlite数据库导出到excel
    Application.ScreenUpdating = False
    i = 0
    For Each tbl In VBA.Split(wsList, ",")
        Set ws = wk.Worksheets.add
        ws.name = VBA.Split(wsListName, ",")(i)
        dbToExcel db, tbl, ws
        i = i + 1
    Next tbl

    Set ws = Nothing
    wk.Windows(1).Visible = True
    Application.ScreenUpdating = True
    wk.Activate
End Function

Private Function createBtn(ByVal ws As Worksheet, ByVal btn As String, ByVal btnCaption As String) As Button
        Dim myBtn As Button

        If btnExists(btn, ws) Then Set createBtn = ws.Buttons(btn): Exit Function

        Set myBtn = ws.Buttons.add(192.75, 108, 72, 72)
        With myBtn
            .name = btn
            .Caption = btnCaption
            .OnAction = "main"
        End With

        Set createBtn = myBtn

        Set myBtn = Nothing
End Function
Private Function createShape(ByVal ws As Worksheet, ByVal myShapeName As String, ByVal myShapeText As String) As Shape
        Dim myShape As Shape

        If shapeExists(myShapeName, ws) Then
            Set myShape = ws.Shapes(myShapeName)
        Else
            Set myShape = ws.Shapes.AddShape(msoShapeRoundedRectangle, 60, 70, 273, 25)
        End If

        With myShape
            .name = myShapeName
            .ShapeStyle = msoShapeStylePreset37
            .TextFrame2.TextRange = myShapeText
            .Line.ForeColor = .Fill.ForeColor
            .Reflection.Type = msoReflectionType1
            With .Fill
                .Visible = msoTrue
                .ForeColor.RGB = RGB(rndNumber(0, 255), rndNumber(0, 255), rndNumber(0, 255))
                .Transparency = 0
                .Solid
            End With
            .TextFrame2.WordWrap = msoTrue
            .TextFrame2.AutoSize = msoAutoSizeShapeToFitText
        End With
        Set createShape = myShape

        Set myShape = Nothing
End Function

Private Function btnExists(ByVal btn As String, ByVal ws As Worksheet) As Boolean
    Dim myBtn As Button

    btnExists = False
    For Each myBtn In ws.Buttons
        If myBtn.name = btn Then btnExists = True: Exit Function
    Next myBtn

End Function

Private Function shapeExists(ByVal myShapeName As String, ByVal ws As Worksheet) As Boolean
    Dim myShape As Shape

    shapeExists = False
    For Each myShape In ws.Shapes
        If myShape.name = myShapeName Then shapeExists = True: Exit Function
    Next myShape
End Function

Private Function createBtnList()
    Dim btnList As String
    Dim btnCaption As String
    Dim btn As Variant
    Dim i As Integer
    Dim ws As Worksheet
    Dim myBtn As Button

    Set ws = home
    btnList = "btnInquiry,btnUpdate,btnUpdateMap,btnHistryChart,btnExportToExcel,ZoomIn,ZoomOut,moveLeft,moveRight,moveUp,moveDown,sizeReset"
    btnCaption = "开始查询,更新数据,更新地图,动态显示,导出数据,放大,缩小,左移←,→右移,上移↑,↓下移,重置地图"
    i = 0
    For Each btn In VBA.Split(btnList, ",")
        Set myBtn = createBtn(ws, btn, VBA.Split(btnCaption, ",")(i))
        With myBtn
            .name = btn
            .Height = 30
            .Width = 54
            .Top = 35 + .Height * i
            .Left = 345
        End With
continue:
        i = i + 1
    Next btn

End Function

Private Function rndNumber(ByVal nMin As Long, ByVal nMax As Long)
    rndNumber = Application.RandBetween(nMin, nMax)
End Function

Function progressShape(ByVal msg As String)
    Dim scr_update_status As Boolean

    scr_update_status = Application.ScreenUpdating

    Application.ScreenUpdating = True
    createShape home, "statusMsg", msg
    DoEvents

    Application.ScreenUpdating = scr_update_status
End Function

Public Sub main()
    Dim myBtn As Button

    Set myBtn = home.Buttons(Application.Caller)
    Select Case myBtn.name
        Case "ZoomOut"
            zoomChart 0
        Case "ZoomIn"
            zoomChart 1
        Case "moveRight"
            zoomChart 2
        Case "moveLeft"
            zoomChart 3
        Case "moveDown"
            zoomChart 4
        Case "moveUp"
            zoomChart 5
        Case "sizeReset"
            zoomChart 6
        Case "btnInquiry"
            inquiryByLocation home.Range("c3")
        Case "btnUpdate"
            Call initMain
        Case "btnUpdateMap"
            Call GenerateMap
        Case "btnHistryChart"
            Call historicalStreamChart
        Case "btnExportToExcel"
            Call exportToExcel
    End Select

    Set myBtn = Nothing
End Sub

Private Function auto_open()
    Call createBtnList
End Function
