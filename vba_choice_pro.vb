Sub SubmitH1()
    'トラック別参照元データ
    Dim TrackSouceData As Range
    Set TrackSouceData = Range("AA7")
    '年齢別参照元データ
    Dim AgeSouceData As Range
    Set AgeSouceData = Range("AH7")
    'トラック別/年齢別の新たな座標
    Dim tmp1 As String
    tmp1 = Str(Range("AN5").End(xlDown).Row)
    Dim TrackNewSouceData_range As String
    TrackNewSouceData_range = Replace("AN5:AN" & tmp1 & ",AP5:AP" & tmp1, " ", "")
    Dim tmp2 As String
    tmp2 = Str(Range("AQ5").End(xlDown).Row)
    Dim AgeNewSouceData_range As String
    AgeNewSouceData_range = Replace("AQ5:AQ" & tmp2 & ",AS5:AS" & tmp2, " ", "")
    
    Debug.Print TrackSouceData
    Debug.Print AgeSouceData
    Debug.Print TrackNewSouceData_range
    Debug.Print AgeNewSouceData_range
    Debug.Print "+++++++++++++++++++++++"

    InputNum TrackSouceData, AgeSouceData, TrackNewSouceData_range, AgeNewSouceData_range
End Sub
Sub SubmitH2()
    'トラック別参照元データ
    Dim TrackSouceData As Range
    Set TrackSouceData = Range("AG7")
    '年齢別参照元データ
    Dim AgeSouceData As Range
    Set AgeSouceData = Range("AP7")
    'トラック別/年齢別の新たな座標
    Dim tmp1 As String
    tmp1 = Str(Range("AV5").End(xlDown).Row)
    Dim TrackNewSouceData_range As String
    TrackNewSouceData_range = Replace("AV5:AV" & tmp1 & ",AX5:AX" & tmp1, " ", "")
    Dim tmp2 As String
    tmp2 = Str(Range("AY5").End(xlDown).Row)
    Dim AgeNewSouceData_range As String
    AgeNewSouceData_range = Replace("AY5:AY" & tmp2 & ",BA5:BA" & tmp2, " ", "")
    
    Debug.Print TrackSouceData
    Debug.Print AgeSouceData
    Debug.Print TrackNewSouceData_range
    Debug.Print AgeNewSouceData_range
    Debug.Print "+++++++++++++++++++++++"

    InputNum TrackSouceData, AgeSouceData, TrackNewSouceData_range, AgeNewSouceData_range
End Sub
