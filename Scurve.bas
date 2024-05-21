Attribute VB_Name = "Scurve"
'TODO:create a tmp sheet to keyin plot order

Sub plotS_Curve()

Dim shp As Shape

For Each shp In Sheets("天氣設定").Shapes

    If shp.OnAction = "" Then shp.Delete

Next

With Sheets("天氣設定")

lr = .Cells(.Rows.Count, 1).End(xlUp).Row

End With

    'Sheets("S-CURVE").Activate

    Sheets("天氣設定").Shapes.AddChart2(240, xlXYScatterSmoothNoMarkers).Select
    ActiveChart.SetSourceData Source:=Range("天氣設定!$A$2:$A$147,天氣設定!$D$2:$D$147,天氣設定!$E$2:$E$48")
    ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.FullSeriesCollection(1).XValues = "=天氣設定!$A$2:$A$" & lr
    ActiveChart.FullSeriesCollection(1).Values = "=天氣設定!$D$2:$D$" & lr
    ActiveChart.FullSeriesCollection(1).Name = "=""預定進度"""
    ActiveChart.FullSeriesCollection(2).XValues = "=天氣設定!$A$2:$A$" & lr
    ActiveChart.FullSeriesCollection(2).Values = "=天氣設定!$E$2:$E$" & lr
    ActiveChart.FullSeriesCollection(2).Name = "=""實際進度"""
    ActiveChart.FullSeriesCollection(3).Delete
    
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MinimumScale = 0
    ActiveChart.Axes(xlValue).MaximumScale = 1
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).MinimumScale = Sheets("天氣設定").Range("A2")
    ActiveChart.Axes(xlCategory).MaximumScale = Sheets("天氣設定").Range("A" & lr)
    
    For Each shp In Sheets("天氣設定").Shapes
    
        If shp.OnAction = "" Then ShpName = shp.Name
    
    Next
    
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.text = "[你的工程名稱]"
    Selection.Format.TextFrame2.TextRange.Characters.text = "[你的工程名稱]"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 2).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 2).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 14
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Spacing = 0
        .Strike = msoNoStrike
    End With
    ActiveChart.ChartArea.Select
    ActiveChart.SetElement (msoElementLegendRight)
    
    ActiveChart.ChartArea.Select
    ActiveSheet.Shapes(ShpName).IncrementLeft -410 + 300
    ActiveSheet.Shapes(ShpName).IncrementTop -140
    ActiveSheet.Shapes(ShpName).ScaleWidth 1.8, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes(ShpName).ScaleHeight 1.8, msoFalse, _
        msoScaleFromTopLeft
        
End Sub



