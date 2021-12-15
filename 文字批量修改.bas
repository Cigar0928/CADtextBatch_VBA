Attribute VB_Name = "功能"
Option Private Module
Sub 文字批量修改1()
'直接换行

    On Error Resume Next
    
    Dim Sht As Worksheet
    Set Sht = ThisWorkbook.Worksheets("工作台")
    Sht.Range(Sht.Cells(2, 6), Sht.Cells(Rows.Count, 6)).ClearContents
    Dim EndRow
    EndRow = Sht.Cells(Rows.Count, 1).End(xlUp).Row
    If EndRow < 2 Then
        MsgBox "工作台中未检测到有效的内容", vbOKOnly + vbCritical, "错误"
        Exit Sub
    End If
    
    Dim TDZLDict As Object
    Set TDZLDict = CreateObject("scripting.dictionary")
    For i = 2 To EndRow
        k = Sht.Cells(i, 1).Text
        TDZLDict(k) = Array(Sht.Cells(i, 2).Text, _
            Sht.Cells(i, 4).Text, Sht.Cells(i, 5).Text, _
            Sht.Cells(i, 6).Value)
    Next i
    
    Dim acadApp As Object
    Set acadApp = GetObject(, "AutoCAD.Application")
    If Err Then
        MsgBox "请先启动AutoCAD，并打开需要修改土地坐落的dwg文件", vbOKOnly + vbCritical, "错误"
        Application.Visible = True
        Exit Sub
    End If
    ' 连接至 AutoCAD 图形
    MsgBox "点击确定之后，请切换到CAD窗口框选需要修改土地坐落的内容区域", vbOKOnly + vbInformation, "提示"
    t1 = Timer
    Set acadDoc = acadApp.activedocument
    
    '创建选择集
    'Err.Clear
    Set SSet1 = createSSet(acadDoc, "SS1")
    
    Dim FilterType1(3) As Integer
    Dim FilterData1(3) As Variant
    FilterType1(0) = -4
    FilterData1(0) = "<OR"
    FilterType1(1) = 0
    FilterData1(1) = "TEXT" 'IAcadText
    FilterType1(2) = 0
    FilterData1(2) = "MTEXT"
    FilterType1(3) = -4
    FilterData1(3) = "OR>"
    SSet1.SelectOnScreen FilterType1, FilterData1
    
    Dim geShu As Integer '获取选择集内对象个数
    geShu = SSet1.Count
    If geShu = 0 Then
        MsgBox "你没有选择有效的编码文字，无法进行修改替换", vbCritical + vbInformation, "提示"
        Exit Sub
    End If
    
    index0 = InputBox("请输入替换图型，" + Chr(13) + "例如：" + Chr(13) + "10，房产分层分户图" + Chr(13) + "08，宗地图", , "08")
    If index0 = "08" Then
        index1 = 4
    Else
        index1 = 2
    End If
    
    For n = 0 To SSet1.Count - 1
        Set objText = SSet1.Item(n)
        'objType = TypeName(objText) 'IAcadText
        num0 = objText.TextString
        If Left(num0, 8) = "43042610" Then
            TDZLItem = TDZLDict(Left(num0, 19))
            SSet1.Item(n + index1).TextString = TDZLItem(2)
            SSet1.Item(n + index1).Color = acMagenta
            TDZLItem(3) = TDZLItem(3) + 1
            TDZLDict(Left(num0, 19)) = TDZLItem
        End If
    Next
    Sht.Cells(1, 6) = "土地坐落处理结果"
    For i = 2 To EndRow
        k = Sht.Cells(i, 1).Text
        Sht.Cells(i, 6) = TDZLDict(k)(3)
    Next i
    ys = Format(Timer - t1, "0.0s")
    MsgBox "更新土地坐落处理完成，用时" & ys, vbOKOnly + vbInformation, "提示"
End Sub

Sub 文字批量修改2()
'双圆定位图框

    On Error Resume Next
    
    Dim Sht As Worksheet
    Set Sht = ThisWorkbook.Worksheets("工作台")
    Sht.Range(Sht.Cells(2, 6), Sht.Cells(Rows.Count, 6)).ClearContents
    Dim EndRow
    EndRow = Sht.Cells(Rows.Count, 1).End(xlUp).Row
    If EndRow < 2 Then
        MsgBox "工作台中未检测到有效的内容", vbOKOnly + vbCritical, "错误"
        Exit Sub
    End If
    
    Dim TDZLDict As Object
    Set TDZLDict = CreateObject("scripting.dictionary")
    For i = 2 To EndRow
        k = Sht.Cells(i, 1).Text
        TDZLDict(k) = Array(Sht.Cells(i, 2).Text, _
            Sht.Cells(i, 4).Text, Sht.Cells(i, 5).Text, _
            Sht.Cells(i, 6).Value)
    Next i
    
    Dim acadApp As Object
    Set acadApp = GetObject(, "AutoCAD.Application")
    If Err Then
        MsgBox "请先启动AutoCAD，并打开需要修改土地坐落的dwg文件", vbOKOnly + vbCritical, "错误"
        Application.Visible = True
        Exit Sub
    End If
    ' 连接至 AutoCAD 图形
    MsgBox "点击确定之后，请切换到CAD窗口框选需要修改土地坐落的内容区域", vbOKOnly + vbInformation, "提示"
    t1 = Timer
    Set acadDoc = acadApp.activedocument
    
    '创建选择集
    'Err.Clear
    Set SSet1 = createSSet(acadDoc, "SS1")
    
    Dim FilterType1(3) As Integer
    Dim FilterData1(3) As Variant
    FilterType1(0) = -4
    FilterData1(0) = "<OR"
    FilterType1(1) = 0
    FilterData1(1) = "TEXT" '"IAcadText"
    FilterType1(2) = 0
    FilterData1(2) = "CIRCLE" '"IAcadCircle"
    FilterType1(3) = -4
    FilterData1(3) = "OR>"
    SSet1.SelectOnScreen FilterType1, FilterData1
    
    Dim geShu As Integer '获取选择集内对象个数
    geShu = SSet1.Count
    If geShu = 0 Then
        MsgBox "你没有选择有效的编码文字，无法进行修改替换", vbCritical + vbInformation, "提示"
        Exit Sub
    End If

    'For n = 0 To SSet1.Count - 1
        'Set objSSet1 = SSet1.Item(n)
        'Debug.Print TypeName(objSSet1)
    'Next
    
    '获取编号位置
    Dim objCharts As Object
    Set objCharts = CreateObject("scripting.dictionary")
    pt = 0
    Dim coords(1)
    For n = 0 To SSet1.Count - 1
        Set objSSet1 = SSet1.Item(n)
        num0 = objSSet1.TextString
        If Left(num0, 8) = "43042610" Then
            key2 = n
        ElseIf TypeName(objSSet1) = "IAcadCircle" Then
            coords(pt) = objSSet1.Center
            If pt = 1 Then
                objCharts(key2) = coords
                coords(0) = Array(0, 0, 0)
                coords(1) = Array(0, 0, 0)
                pt = -1
            End If
            pt = pt + 1
        End If
    Next

    '开始批量修改
    Dim numItem As AcadEntity
    For Each key2 In objCharts.keys
        Set numItem = SSet1.Item(key2)
        num0 = numItem.TextString
        k = Left(num0, 19)
        pt1 = objCharts(key2)(0)
        pt2 = objCharts(key2)(1)
        Dim ptArr(12) As Double
        ptArr(0) = pt1(0) - 1: ptArr(1) = pt1(1): ptArr(2) = 0
        ptArr(3) = pt2(0): ptArr(4) = pt1(1): ptArr(5) = 0
        ptArr(6) = pt2(0): ptArr(7) = pt2(1): ptArr(8) = 0
        ptArr(9) = pt1(0) - 1: ptArr(10) = pt2(1): ptArr(11) = 0
        
        '创建选择集-土地坐落
        Set SSet2 = createSSet(acadDoc, "SS2")
        
        Dim FilterType2(0) As Integer
        Dim FilterData2(0) As Variant
        FilterType2(0) = 0
        FilterData2(0) = "TEXT"

        SSet2.SelectByPolygon acSelectionSetCrossingPolygon, ptArr, FilterType2, FilterData2
        
        Dim objSSet2 As AcadEntity
        For Each objSSet2 In SSet2
            TDZL0 = objSSet2.TextString
            TDZLItem = TDZLDict(k)
            If (InStr(TDZL0, "祁东县") > 0 Or InStr(TDZL0, "河洲镇") > 0) And TDZL0 <> "祁东县自然资源局" Then
                objSSet2.TextString = TDZLItem(2)
                objSSet2.Color = acMagenta
                TDZLItem(3) = TDZLItem(3) + 1
                TDZLDict(k) = TDZLItem
            End If
        Next
    Next
    
    Sht.Cells(1, 6) = "土地坐落处理结果"
    For i = 2 To EndRow
        k = Sht.Cells(i, 1).Text
        Sht.Cells(i, 6) = TDZLDict(k)(3)
    Next i
    ys = Format(Timer - t1, "0.0s")
    MsgBox "更新土地坐落处理完成，用时" & ys, vbOKOnly + vbInformation, "提示"
End Sub

Sub 文字批量修改3()
'自带图框

    On Error Resume Next
    
    Dim Sht As Worksheet
    Set Sht = ThisWorkbook.Worksheets("工作台")
    Sht.Range(Sht.Cells(2, 4), Sht.Cells(Rows.Count, 4)).ClearContents
    Dim EndRow
    EndRow = Sht.Cells(Rows.Count, 1).End(xlUp).Row
    If EndRow < 2 Then
        MsgBox "工作台中未检测到有效的内容", vbOKOnly + vbCritical, "错误"
        Exit Sub
    End If
    
    Dim TDZLDict1 As Object
    Set TDZLDict1 = CreateObject("scripting.dictionary")
    For i = 2 To EndRow
        k = Sht.Cells(i, 1).Text
        TDZLDict1(k) = Array(Sht.Cells(i, 2).Text, _
            Sht.Cells(i, 3).Text, Sht.Cells(i, 4).Value)
    Next i
    
    Dim acadApp As IAcadApplication
    Set acadApp = GetObject(, "AutoCAD.Application")
    If Err Then
        MsgBox "请先启动AutoCAD，并打开需要修改土地坐落的dwg文件", vbOKOnly + vbCritical, "错误"
        Application.Visible = True
        Exit Sub
    End If
    ' 连接至 AutoCAD 图形
    MsgBox "点击确定之后，请切换到CAD窗口框选需要修改土地坐落的内容区域", vbOKOnly + vbInformation, "提示"
    t1 = Timer
    Dim acadDoc As IAcadDocument
    Set acadDoc = acadApp.activedocument
    
    '创建选择集
    'Err.Clear
    Set SSet1 = createSSet(acadDoc, "SS1")
    
    Dim FilterType1(3) As Integer
    Dim FilterData1(3) As Variant
    FilterType1(0) = -4
    FilterData1(0) = "<OR"
    FilterType1(1) = 0
    FilterData1(1) = "TEXT" '"IAcadText"
    FilterType1(2) = 0
    FilterData1(2) = "LWPOLYLINE" '"IAcadLWPolyline"
    FilterType1(3) = -4
    FilterData1(3) = "OR>"
    SSet1.SelectOnScreen FilterType1, FilterData1
    
    Dim geShu As Integer '获取选择集内对象个数
    geShu = SSet1.Count
    If geShu = 0 Then
        MsgBox "你没有选择有效的编码文字，无法进行修改替换", vbCritical + vbInformation, "提示"
        Exit Sub
    End If
    
    index0 = InputBox("请输入替换图型，" + Chr(13) + "例如：" + Chr(13) + "10，房产分层分户图" + Chr(13) + "08，宗地图", , "10")
    If index0 = "08" Then
        ChartArea0 = 39000
    Else
        ChartArea0 = 1400
    End If
    

    '获取编号、图框位置
    pt = 0
    Dim coords() As Double
    Dim objSSet1 As AcadEntity
    Dim TDZLDict2 As Object
    Set TDZLDict2 = CreateObject("scripting.dictionary")
    Dim tempArr0(), tempArr1(), tempArr2()
    For n = 0 To SSet1.Count - 1
        Set objSSet1 = SSet1.Item(n)
        num0 = objSSet1.TextString
        If Left(num0, 8) = "43042610" Then
            key3 = Left(num0, 19)
            ReDim Preserve tempArr0(k0)
            tempArr0(k0) = key3
            ReDim Preserve tempArr1(k0)
            tempArr1(k0) = n
            k0 = k0 + 1
        ElseIf TypeName(objSSet1) = "IAcadLWPolyline" Then
            If objSSet1.Closed = True And objSSet1.Area > ChartArea0 Then
                tempCoord = objSSet1.Coordinates
                objSSet1.Color = acMagenta
                j = 0
                For i = 0 To UBound(tempCoord) Step 2
                    ReDim Preserve coords(j + 2)
                    coords(j) = tempCoord(i): coords(j + 1) = tempCoord(i + 1): coords(j + 2) = 0
                    j = j + 3
                Next
                ReDim Preserve tempArr2(k2)
                tempArr2(k2) = coords
                k2 = k2 + 1
            End If
        End If
    Next
    For k1 = 0 To k0
        Dim tempArr()
        tempArr = TDZLDict1(tempArr0(k1))
        ReDim Preserve tempArr(4)
        tempArr(3) = tempArr1(k1)
        tempArr(4) = tempArr2(k1)
        TDZLDict2(tempArr0(k1)) = tempArr
        Debug.Print tempArr0(k1), tempArr(3), tempArr(4)(0)
    Next

    '开始批量修改
    Dim numItem As AcadEntity
    For Each key2 In TDZLDict2.keys
        TDZLItem = TDZLDict2(key2)
        Set numItem = SSet1.Item(TDZLItem(3))
        num0 = numItem.TextString
        ptArr = TDZLItem(4)
        
        '创建选择集
        Dim SSet2 As AcadSelectionSet
        Set SSet2 = createSSet(acadDoc, "SS2")
        Dim FilterType2(0) As Integer
        Dim FilterData2(0) As Variant
        FilterType2(0) = 0
        FilterData2(0) = "TEXT"
        
        'acadApp.ZoomAll
        'ZoomCenter
        Dim zcenter(0 To 2) As Double
        Dim magnification As Double
        zcenter(0) = ptArr(6): zcenter(1) = ptArr(7): zcenter(2) = ptArr(8)
        magnification = 1.5
        'Call acadApp.ZoomCenter(zcenter, magnification)
        'ZoomScaled
        Dim scalefactor As Double
        Dim scaletype As Integer
        scalefactor = 0.01
        scaletype = acZoomScaledRelative
        'Call acadApp.ZoomScaled(scalefactor, scaletype)
        
        SSet2.SelectByPolygon acSelectionSetCrossingPolygon, ptArr, FilterType2, FilterData2
        
        Dim objSSet2 As AcadEntity
        For Each objSSet2 In SSet2
            TDZL0 = objSSet2.TextString
            If (InStr(TDZL0, "祁东县") > 0 Or InStr(TDZL0, "河洲镇") > 0 Or InStr(TDZL0, "归阳镇") > 0) And TDZL0 <> "祁东县自然资源局" Then
                objSSet2.TextString = TDZLItem(1)
                objSSet2.Color = acMagenta
                TDZLItem(2) = TDZLItem(2) + 1
                TDZLDict2(key2) = TDZLItem
                Exit For
            End If
        Next
    Next
    
    For i = 2 To EndRow
        k = Sht.Cells(i, 1).Text
        Sht.Cells(i, 4) = TDZLDict2(k)(2)
    Next i
    ys = Format(Timer - t1, "0.0s")
    MsgBox "更新土地坐落处理完成，用时" & ys, vbOKOnly + vbInformation, "提示"
End Sub

Public Function createSSet(ByVal acadDoc As AcadDocument, ByVal SSetName As String) As AcadSelectionSet
'安全创建选择集
    Dim SSet As AcadSelectionSet
    Dim i As Integer
    For i = 0 To acadDoc.SelectionSets.Count - 1
        Set SSet = acadDoc.SelectionSets.Item(i)
        If StrComp(SSet.Name, SSetName, vbTextCompare) = 0 Then
            SSet.Delete
            Exit For
        End If
    Next i
    Set createSSet = acadDoc.SelectionSets.Add(SSetName)
End Function

Public Function AddRec(ByVal acadDoc As AcadDocument, ByVal pt1 As Variant, ByVal pt2 As Variant) As AcadLWPolyline
    Dim ptArr(7) As Double
    ptArr(0) = pt1(0) - 1: ptArr(1) = pt1(1)
    ptArr(2) = pt2(0): ptArr(3) = pt1(1)
    ptArr(4) = pt2(0): ptArr(5) = pt2(1)
    ptArr(6) = pt1(0) - 1: ptArr(7) = pt2(1)
    
    Set AddRec = acadDoc.ModelSpace.AddLightWeightPolyline(ptArr)
    AddRec.Closed = True
End Function
Sub aaa()

    Dim acadApp As Object
    Set acadApp = GetObject(, "AutoCAD.Application")

    Set acadDoc = acadApp.activedocument
    LL = acadDoc.Utility.GetPoint(, "拾取第一个点:")
    UR = acadDoc.Utility.GetCorner(LL, "拾取对角点:")
    
    Dim objCircle As AcadCircle
    Set objCircle = acadDoc.ModelSpace.AddCircle(LL, 50)
    Set objCircle = acadDoc.ModelSpace.AddCircle(UR, 50)
    Set View = acadDoc.ModelSpace.ZoomWindow(LL, UR)
End Sub

