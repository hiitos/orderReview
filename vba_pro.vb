' 取得する情報 ------
Type VisualFigureType
    '<<メイン>>  .top,.left,.width,.height
    Ctop As Double
    Cleft As Double
    Cwidth As Double
    Cheight As Double
    '<<プロットエリア>>  .PlotArea.Top,.PlotArea.Left,.PlotArea.Width,.PlotArea.Height,.PlotArea.InsideTop,.PlotArea.InsideLeft,.PlotArea.InsideWidth,.PlotArea.InsideHeight
    Ptop As Double
    Pleft As Double
    Pwidth As Double
    Pheight As Double
    PAInsideTop As Double
    PAInsideLeft As Double
    PAInsideWidth As Double
    PAInsideHeight As Double
    '<<フォント>>       .ChartArea.Font.Name,.ChartArea.Font.Size
    CAFontName As String
    '<<Title>>         .ChartTitle.Left,.ChartTitle.Top,.ChartTitle.Format.TextFrame2.TextRange.Font.Size
    CTleft As Double
    CTtop As Double
    CTfontSize As Single
    '<<Legend>>        .Legend.Top,.Legend.Left,.Legend.Width,.Legend.Height,.Legend.Font.Size,.SeriesCollection(1).DataLabels.Font.Size
    Ltop As Double
    Lleft As Double
    Lwidth As Double
    Lheight As Double
    LfontSize As Single
    DLFontSize As Single
    '<<ChartArea>>      .ChartArea.top,.ChartArea.left,.ChartArea.width,.ChartArea.height
    ChartTop As Double
    ChartLeft As Double
    ChartWidth As Double
    ChartHeight As Double
End Type
Sub InputNum(TrackSouceData As Range, AgeSouceData As Range, TrackNewSouceData_range As String, AgeNewSouceData_range As String)
    ' Range("AA7")  Range("AH7")  AN5:AN19,AP5:AP19  AQ5:AQ11,AS5:AS11 など
    
    Dim i As Long
    Dim c As Range
    Dim d As Range
    Dim tmp1 As Variant
    Dim tmp2 As Variant
    tmp1 = Split(TrackNewSouceData_range, ",")
    tmp2 = Split(AgeNewSouceData_range, ",")
    
    '<<コピペした項目に値を入れる>>
    ' トラック別の列の始まりから終わりまで
    For i = TrackSouceData.Row To TrackSouceData.End(xlDown).Row
        ' 検索する時じゃななので、スペース削除
        Cells(i, TrackSouceData.Column) = Replace(Cells(i, TrackSouceData.Column), " ", "")
        '検索し、一致する項目の値を入れる
        Set c = Range(tmp1(0)).Find(Cells(i, TrackSouceData.Column).Value, LookAt:=xlWhole, SearchOrder:=xlByColumns)
        Cells(c.Row, c.Column + 2) = Round(Cells(i, TrackSouceData.Column + 2).Value, 3)
        '　年齢別はトラック別より項目数が少ないので、
        If i <= AgeSouceData.End(xlDown).Row Then
            Cells(i, AgeSouceData.Column) = Replace(Cells(i, AgeSouceData.Column), " ", "")
            Set d = Range(tmp2(0)).Find(Cells(i, AgeSouceData.Column).Value, LookAt:=xlWhole, SearchOrder:=xlByColumns)
            Cells(d.Row, d.Column + 2) = Cells(i, AgeSouceData.Column + 1).Value / 100
        End If
    Next i
    PositionGet TrackNewSouceData_range, AgeNewSouceData_range
End Sub
Sub PositionGet(TrackNewSouceData_range As String, AgeNewSouceData_range As String)
    Dim figureArray() As Variant
    Dim targetTitle As String
    Dim PassingArguments As VisualFigureType  '上で定義したデータ型
    '検索元の配列を作成
    figureTitleArray = Array("再生数(トラフィック別)", "再生数(年齢別)")
    titleIndexArray = Array(TrackNewSouceData_range, AgeNewSouceData_range)

    Dim cho As ChartObject
    ' シートに表示されているグラフを一つずつ取り出す
    For Each cho In ActiveSheet.ChartObjects
        Debug.Print "<----------------" & cho.Chart.ChartTitle.Text & "を検索中---------------->"
        'グラフのタイトルを取得、このタイトルで最描画するグラフを判別する
        targetTitle = cho.Chart.ChartTitle.Text
        
        'ループ処理で配列を検索
        Dim i As Long
        Dim dataCD As String
        For i = 0 To UBound(figureTitleArray) 'figureTitleArrayの要素分forを回す
            If StrComp(figureTitleArray(i), targetTitle) = 0 Then ' グラフのタイトルがfigureTitleArrayに含まれていたら
                dataCD = titleIndexArray(i)    '円グラフの座標

                '<<メイン>>
                PassingArguments.Ctop = cho.Top
                PassingArguments.Cleft = cho.Left
                PassingArguments.Cwidth = cho.Width
                PassingArguments.Cheight = cho.Height
                '<<プロットエリア>>
                PassingArguments.Ptop = cho.Chart.PlotArea.Top
                PassingArguments.Pleft = cho.Chart.PlotArea.Left
                PassingArguments.Pwidth = cho.Chart.PlotArea.Width
                PassingArguments.Pheight = cho.Chart.PlotArea.Height
                PassingArguments.PAInsideTop = cho.Chart.PlotArea.InsideTop
                PassingArguments.PAInsideLeft = cho.Chart.PlotArea.InsideLeft
                PassingArguments.PAInsideWidth = cho.Chart.PlotArea.InsideWidth
                PassingArguments.PAInsideHeight = cho.Chart.PlotArea.InsideHeight
                '<<フォント>>
                PassingArguments.CAFontName = cho.Chart.ChartArea.Font.Name
                '<<Title>>  (グラフタイトルの配置左側)、(グラフタイトルの配置上側)
                PassingArguments.CTtop = cho.Chart.ChartTitle.Top
                PassingArguments.CTleft = cho.Chart.ChartTitle.Left
                PassingArguments.CTfontSize = cho.Chart.ChartTitle.Format.TextFrame2.TextRange.Font.Size
                '<<Legend>>
                PassingArguments.DLFontSize = cho.Chart.SeriesCollection(1).DataLabels.Font.Size
                PassingArguments.Ltop = cho.Chart.Legend.Top
                PassingArguments.Lleft = cho.Chart.Legend.Left
                PassingArguments.Lwidth = cho.Chart.Legend.Width
                PassingArguments.Lheight = cho.Chart.Legend.Height
                PassingArguments.LfontSize = cho.Chart.Legend.Font.Size
                
                '<<ChartArea>>
                PassingArguments.ChartTop = cho.Chart.ChartArea.Top
                PassingArguments.ChartLeft = cho.Chart.ChartArea.Left
                PassingArguments.ChartWidth = cho.Chart.ChartArea.Width
                PassingArguments.ChartHeight = cho.Chart.ChartArea.Height
                
                cho.Delete ' 元のグラフを削除
                VisualFigure targetTitle, dataCD, PassingArguments
            End If
        Next i
    Next cho
End Sub

Sub VisualFigure(title As String, dataCD As String, PassingArguments As VisualFigureType)
With ActiveSheet.Shapes.AddChart(xlPie, PassingArguments.Cleft, PassingArguments.Ctop, PassingArguments.Cwidth, PassingArguments.Cheight).Chart
    .SetSourceData Range(dataCD)
     '<<フォント>>
    .ChartArea.Font.Name = PassingArguments.CAFontName
    '<<Title>>
    .HasTitle = True
    .ChartTitle.Text = title
    .ChartTitle.Top = PassingArguments.CTtop
    .ChartTitle.Left = PassingArguments.CTleft
    .ChartTitle.Format.TextFrame2.TextRange.Font.Size = PassingArguments.CTfontSize
    .ChartTitle.Font.Bold = False
    '<<Legend>>
    .HasLegend = True
    .Legend.Top = PassingArguments.Ltop
    .Legend.Left = PassingArguments.Lleft
    .Legend.Width = PassingArguments.Lwidth
    .Legend.Height = PassingArguments.Lheight
    .Legend.Format.TextFrame2.TextRange.Font.Size = 11
    '<<プロットエリア>>
    .PlotArea.Top = PassingArguments.Ptop
    .PlotArea.Left = PassingArguments.Pleft
    .PlotArea.Width = PassingArguments.Pwidth
    .PlotArea.Height = PassingArguments.Pheight
    .PlotArea.InsideTop = PassingArguments.PAInsideTop
    .PlotArea.InsideLeft = PassingArguments.PAInsideLeft
    .PlotArea.InsideWidth = PassingArguments.PAInsideWidth
    .PlotArea.InsideHeight = PassingArguments.PAInsideHeight
    '<<ChartArea>>
    .ChartArea.Format.Fill.Visible = False
    .ChartArea.Format.Line.Visible = msoFalse
    .ChartArea.Top = PassingArguments.ChartTop
    .ChartArea.Left = PassingArguments.ChartLeft
    .ChartArea.Width = PassingArguments.ChartWidth
    .ChartArea.Height = PassingArguments.ChartHeight

    '  AN5:AN19,AP5:AP19
    Dim tmp_color1 As Variant
    tmp_color1 = Split(dataCD, ",")
    'Debug.Print tmp_color1(0)  ' AN5:AN19  AP5:AP19
    Dim tmp_color2 As Variant
    tmp_color2 = Split(tmp_color1(0), ":")
    'Debug.Print tmp_color2(0)    ' AN5
    ColorIndexRow = Range(tmp_color2(0)).Row - 1
    ColorIndexCol = Range(tmp_color2(0)).Column + 1
    
    Dim i As Long   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Dim LegendDleteList() As Variant
    LegendDleteList = Array()
    '指定されている色でグラフを描画
    For i = 1 To .SeriesCollection(1).Points.Count 'グラフの要素分
        .SeriesCollection(1).Points(i).Interior.color = Cells(i + ColorIndexRow, ColorIndexCol).Interior.color
        If Cells(i + ColorIndexRow, ColorIndexCol + 1).Value >= 0.05 Then ' 5%以下ならラベルの表示を行わない
            .SeriesCollection(1).Points.Item(i).HasDataLabel = True
            .SeriesCollection(1).Points.Item(i).DataLabel.NumberFormatLocal = "0%"
        End If

        If Cells(i + ColorIndexRow, ColorIndexCol + 1).Value = "" Then
            ReDim Preserve LegendDleteList(UBound(LegendDleteList) + 1)
            LegendDleteList(UBound(LegendDleteList)) = i
        End If
    Next i
    '円グラフのラベルのフォントサイズの設定
    .SeriesCollection(1).DataLabels.Format.TextFrame2.TextRange.Font.Size = PassingArguments.DLFontSize
    
    Dim temNum As Long
    Debug.Print "LegendDleteList"
    For k = UBound(LegendDleteList) To LBound(LegendDleteList) Step -1
        temNum = k
        Debug.Print LegendDleteList(temNum)
        .Legend.LegendEntries(LegendDleteList(temNum)).Delete
    Next k

End With
End Sub



