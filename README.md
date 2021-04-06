# Stress-strain-curve
#Engineering Stress-Strain curve in Excel-VBA

Private Sub CommandButton4_Click()
 Dim xtitle, ytitle, Rowx, Columnx, Rowy, Columny, Startpoint, Endpoint, endp As String

Rowx = InputBox("Enter Row of strain")
Columnx = InputBox("Enter Column of strain")
Rowy = InputBox("Enter Row of stress")
Columny = InputBox("Enter Column of stress")
endp = Range("C10").Value + Rowx

If Rowx = Rowy Then
Startpoint = Columnx & Rowx
Endpoint = Columny & endp
End If

'MsgBox Startpoint & " " & Endpoint

xtitle = "Engineering strain"
ytitle = "Engineering stress"
    Range(Startpoint, Endpoint).Select
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.ChartType = xlXYScatterSmoothNoMarkers
   ' ActiveChart.SetSourceData Source:=Range("Sheet1!$E$14:$F$1000")
   ActiveChart.SetSourceData Source:=Range(Startpoint, Endpoint)
  
    ActiveChart.SeriesCollection(1).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .Weight = 0.75
    End With
    ActiveChart.ChartArea.Select
    ActiveChart.Legend.Select
    ActiveChart.SeriesCollection(1).Name = "=Sheet1!$A$11"
    With ActiveChart.Axes(xlValue)
        .Format.Line.Visible = msoTrue
        .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
        .Format.Line.Transparency = 0
        .HasTitle = True
        .AxisTitle.Text = ytitle
        End With
    With ActiveChart.Axes(xlCategory)
        .Format.Line.Visible = msoTrue
        .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
        .Format.Line.Transparency = 0
        .HasTitle = True
        .AxisTitle.Text = xtitle
    End With
    
End Sub
