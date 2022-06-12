VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class_Chart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [类名] Class_Chart
'* [描述] Simple Chart control
'*
'* @author Hans Yao <hansyao@gmail.com>
'* Copyright (c) 2019-2022 Hans Yao
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* 内部函数定义
'******************************************************************************
Option Explicit

Private Type oSeries
    SeriesName  As String
    seriesValueX As String
    seriesValue As String
    seriesChartType As Integer
End Type
Private custSeries As oSeries
Private mName As String
Private WithEvents CEvents As Chart
Attribute CEvents.VB_VarHelpID = -1

Public Property Get SeriesName() As String
    SeriesName = custSeries.SeriesName
End Property
Public Property Let SeriesName(arg As String)
    custSeries.SeriesName = arg
End Property
Public Property Get seriesValueX() As String
    seriesValueX = custSeries.seriesValueX
End Property
Public Property Let seriesValueX(arg As String)
    custSeries.seriesValueX = arg
End Property
Public Property Get seriesValue() As String
    seriesValue = custSeries.seriesValue
End Property
Public Property Let seriesValue(arg As String)
    custSeries.seriesValue = arg
End Property

Public Property Get name() As String
    name = mName
End Property
Public Property Let name(arg As String)
    mName = arg
End Property

Public Function newChart(ByVal ws As Worksheet, Optional ByVal iChartType As Long = xlXYScatterSmoothNoMarkers) As ChartObject
    Dim myChart As ChartObject

    Set myChart = ws.ChartObjects.add(Left:=400, Width:=800, Top:=15, Height:=600)
    With myChart
        .name = mName
        .Chart.chartType = iChartType
        .Chart.ChartStyle = 248
    End With
    Set newChart = myChart
    Set myChart = Nothing
End Function

Public Function isSeriesExists(ByVal myChart As ChartObject, ByVal SeriesName As String) As Boolean
    Dim mySeries As Variant

    For Each mySeries In myChart.Chart.SeriesCollection
        If mySeries.name = SeriesName Then
            isSeriesExists = True
            Exit Function
        End If
    Next

    isSeriesExists = False
End Function

Public Function isChartExists(ByVal ws As Worksheet) As Boolean
    Dim myChart As Variant

    isChartExists = False
    For Each myChart In ws.ChartObjects
        If myChart.name = mName Then
            isChartExists = True
            Exit Function
        End If
    Next
End Function

Public Function appendSeries(ByVal myChart As ChartObject, Optional ByVal iChartType As Long = xlXYScatterSmoothNoMarkers)
    Dim mySeries As Series

    If isSeriesExists(myChart, custSeries.SeriesName) Then Exit Function

    Set mySeries = myChart.Chart.SeriesCollection.NewSeries
    With mySeries
        .name = custSeries.SeriesName
        .XValues = custSeries.seriesValueX
        .Values = custSeries.seriesValue
        .chartType = iChartType
    End With

    Set mySeries = Nothing
End Function

Public Function updateSeries(ByVal myChart As ChartObject, Optional ByVal iChartType As Long = xlXYScatterSmoothNoMarkers)
    Dim mySeries As Variant

    If Not isSeriesExists(myChart, custSeries.SeriesName) Then
        appendSeries myChart
        Exit Function
    End If
    Set mySeries = myChart.Chart.SeriesCollection(custSeries.SeriesName)
    If mySeries.name = custSeries.SeriesName Then
        With mySeries
            .name = custSeries.SeriesName
            .XValues = custSeries.seriesValueX
            .Values = custSeries.seriesValue
            .chartType = iChartType
        End With
    End If
End Function

Public Function getCharts(ByVal ws As Worksheet) As Boolean
    Dim myChart As Variant

    getCharts = False
    For Each myChart In ws.ChartObjects
        If myChart.name = mName Then
            getCharts = True
            Exit For
        End If
    Next
End Function
