Sub pick()
'
'
' *** for comments
    Dim E As Range, d As Object, i As Integer, score As Integer, record As Integer
    Dim winner As String
    Dim key As Variant
    Dim tmp_key As String
    Dim member_of_each_row() As String

    Set d = CreateObject("Scripting.Dictionary") '字典物件
    i = 2
    With Worksheets(1)
        Do While .Cells(i, "F") <> ""
            tmp_key = .Cells(i, "F").Value
            '讀取人員資料

            member_of_each_row = Split(tmp_key, "、")
            '處理職稱

            For Each man In member_of_each_row
                man = cut_title(man)
                'Debug.Print man
                If d.EXISTS(man) Then 'Dictionary 物件中指定的關鍵字存在，傳回 True，若不存在，傳回 False。
                    d(man).Item("counts") = d(man).Item("counts") + 1
                Else
                    Dim mydictionary As Object
                    Set mydictionary = CreateObject("Scripting.Dictionary")
                    mydictionary.Add "counts", man
                    mydictionary("counts") = 1
                    d.Add man, mydictionary
                End If
            Next

            i = i + 1
        Loop
    End With

    For Each key In d.Keys
        For Each key2 In d(key).Keys
            'Debug.Print "Key: " & key & " Value: " & d(key).Item(key2)
        Next
    Next

    'Debug.Print d.Item("陳怡君").Item("counts")

    Dim j As Integer
    j = 1
    For Each E In Worksheets(1).UsedRange.Columns(6).Offset(1).Cells
        If E = "" Then Exit For
        record = 0
        winner = ""
        member_of_each_row = Split(E, "、")
            '處理職稱
        For Each man In member_of_each_row
            man = cut_title(man)
            'Debug.Print man
            score = d.Item(man).Item("counts") '取出次數'
            If score > record Then
                record = score
                winner = man
            End If
        Next
        Debug.Print "受獎人: " & winner & " no. " & j
        Worksheets(1).Cells(j + 1, "H") = winner
        j = j + 1
    Next
End Sub

Function cut_title(ByVal man As String) As String
    man = Replace(man, "警員", "")
    man = Replace(man, "所長", "")
    man = Replace(man, "巡佐", "")
    man = Replace(man, "役男", "")
    cut_title = man
    Exit Function
End Function
