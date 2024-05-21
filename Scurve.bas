Attribute VB_Name = "Scurve"
'TODO:create a tmp sheet to keyin plot order

Sub plotS_Curve()

Dim shp As Shape

For Each shp In Sheets("�Ѯ�]�w").Shapes

    If shp.OnAction = "" Then shp.Delete

Next

With Sheets("�Ѯ�]�w")

lr = .Cells(.Rows.Count, 1).End(xlUp).Row

End With

    'Sheets("S-CURVE").Activate

    Sheets("�Ѯ�]�w").Shapes.AddChart2(240, xlXYScatterSmoothNoMarkers).Select
    ActiveChart.SetSourceData Source:=Range("�Ѯ�]�w!$A$2:$A$147,�Ѯ�]�w!$D$2:$D$147,�Ѯ�]�w!$E$2:$E$48")
    ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.FullSeriesCollection(1).XValues = "=�Ѯ�]�w!$A$2:$A$" & lr
    ActiveChart.FullSeriesCollection(1).Values = "=�Ѯ�]�w!$D$2:$D$" & lr
    ActiveChart.FullSeriesCollection(1).Name = "=""�w�w�i��"""
    ActiveChart.FullSeriesCollection(2).XValues = "=�Ѯ�]�w!$A$2:$A$" & lr
    ActiveChart.FullSeriesCollection(2).Values = "=�Ѯ�]�w!$E$2:$E$" & lr
    ActiveChart.FullSeriesCollection(2).Name = "=""��ڶi��"""
    ActiveChart.FullSeriesCollection(3).Delete
    
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MinimumScale = 0
    ActiveChart.Axes(xlValue).MaximumScale = 1
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).MinimumScale = Sheets("�Ѯ�]�w").Range("A2")
    ActiveChart.Axes(xlCategory).MaximumScale = Sheets("�Ѯ�]�w").Range("A" & lr)
    
    For Each shp In Sheets("�Ѯ�]�w").Shapes
    
        If shp.OnAction = "" Then ShpName = shp.Name
    
    Next
    
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.text = "[�A���u�{�W��]"
    Selection.Format.TextFrame2.TextRange.Characters.text = "[�A���u�{�W��]"
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



