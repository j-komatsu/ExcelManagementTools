Sub CreateCalendar()
    Dim ws As Worksheet
    Dim startDate As Date, endDate As Date
    Dim currentDate As Date
    Dim rowOffset As Integer, colOffset As Integer
    Dim i As Integer

    ' シートを作成
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "進捗カレンダー"

    ' A1にタイトルを出力
    ws.Cells(1, 1).Value = "進捗カレンダー"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 16

    ' D1に「期間：開始日～終了日」を表示
    ws.Cells(1, 4).Value = "期間："
    ws.Cells(1, 4).Font.Bold = True

    ' 設定エリアを作成
    ws.Cells(2, 1).Value = "連番"
    ws.Cells(2, 2).Value = "タスク名"
    ws.Cells(2, 3).Value = "担当者"
    ws.Cells(2, 4).Value = "進捗状況"
    ws.Cells(2, 5).Value = "予定開始日"
    ws.Cells(2, 6).Value = "予定終了日"
    ws.Cells(2, 7).Value = "工数（日）" ' 工数列を追加

    ' ヘッダーの背景色と罫線を設定
    With ws.Rows(2)
        .Font.Bold = True
        .Interior.Color = RGB(173, 216, 230) ' 明るい青っぽい色
        .Borders.LineStyle = xlContinuous
    End With

    ' 列幅の調整
    ws.Columns(1).ColumnWidth = 5
    ws.Columns(2).ColumnWidth = 20
    ws.Columns(3).ColumnWidth = 15
    ws.Columns(4).ColumnWidth = 15
    ws.Columns(5).ColumnWidth = 15
    ws.Columns(6).ColumnWidth = 15
    ws.Columns(7).ColumnWidth = 10 ' 工数列の幅を調整

    ' デフォルトで100行分のフォーマットを設定
    Dim taskRow As Integer
    For taskRow = 3 To 102 ' 100行分
        ws.Cells(taskRow, 1).Value = taskRow - 2 ' 連番を自動入力
        ws.Cells(taskRow, 4).Value = "未設定" ' 初期状態で「未設定」を入力
    Next taskRow

    ' 仮データを1行のみ入力
    AddSampleData ws, 1 ' サンプルデータを1行だけ追加

    ' カレンダー設定エリア
    startDate = Application.InputBox("カレンダーの開始日を入力してください (YYYY/MM/DD):", Type:=2)
    endDate = Application.InputBox("カレンダーの終了日を入力してください (YYYY/MM/DD):", Type:=2)

    If Not IsDate(startDate) Or Not IsDate(endDate) Then
        MsgBox "日付が正しくありません。終了します。", vbExclamation
        Exit Sub
    End If

    If startDate > endDate Then
        MsgBox "開始日は終了日より前に設定してください。", vbExclamation
        Exit Sub
    End If

    ' D1セルに「期間：開始日～終了日」を表示
    ws.Cells(1, 4).Value = "期間：" & Format(startDate, "yyyy/mm/dd") & " ～ " & Format(endDate, "yyyy/mm/dd")

    ' カレンダーの背景色と日付を設定
    UpdateCalendarBackground ws, startDate, endDate

    ' データの入力規則を設定
    SetDateValidation ws, startDate, endDate

    ' 工数を計算して設定
    CalculateEffort ws

    MsgBox "進捗カレンダーを作成しました！", vbInformation
End Sub

Sub CalculateEffort(ws As Worksheet)
    Dim lastRow As Long
    Dim taskRow As Long
    Dim startDate As Date, endDate As Date
    Dim effort As Long

    ' データ範囲の最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, 5).End(xlUp).Row

    ' 各行の工数を計算
    For taskRow = 3 To lastRow
        If IsDate(ws.Cells(taskRow, 5).Value) And IsDate(ws.Cells(taskRow, 6).Value) Then
            startDate = ws.Cells(taskRow, 5).Value
            endDate = ws.Cells(taskRow, 6).Value

            ' 工数を平日かつ祝日を除外して計算
            effort = 0
            Dim currentDate As Date
            For currentDate = startDate To endDate
                If Weekday(currentDate, vbMonday) <= 5 And Not IsHoliday(currentDate) Then ' 平日（Monday=1, Friday=5）かつ祝日でない
                    effort = effort + 1
                End If
            Next currentDate

            ws.Cells(taskRow, 7).Value = effort
        Else
            ws.Cells(taskRow, 7).Value = "" ' 開始日または終了日が無効の場合は空白
        End If
    Next taskRow

    MsgBox "工数の計算が完了しました！", vbInformation
End Sub

Sub UpdateCalendarBackground(ws As Worksheet, startDate As Date, endDate As Date)
    Dim currentDate As Date
    Dim rowOffset As Integer, colOffset As Integer
    Dim i As Integer

    ' カレンダーの開始位置
    rowOffset = 2
    colOffset = 8 ' 工数列 (7列目) の次の列から開始するように調整
    currentDate = startDate
    i = 0

    Debug.Print "カレンダー背景色設定のテスト開始..."
    Do While currentDate <= endDate
        With ws.Cells(rowOffset, colOffset + i)
            .Value = currentDate
            .NumberFormat = "MM/DD"
            .Borders.LineStyle = xlContinuous

            ' 色設定の条件
            If IsHoliday(currentDate) Then
                .Interior.Color = RGB(255, 102, 102) ' 赤（祝日）
                Debug.Print "日付: " & currentDate & " は祝日（赤色設定）"
            ElseIf Weekday(currentDate) = vbSunday Then
                .Interior.Color = RGB(255, 182, 193) ' 薄い赤（日曜日）
                Debug.Print "日付: " & currentDate & " は日曜日（薄い赤設定）"
            ElseIf Weekday(currentDate) = vbSaturday Then
                .Interior.Color = RGB(173, 216, 230) ' 薄い青（土曜日）
                Debug.Print "日付: " & currentDate & " は土曜日（薄い青設定）"
            Else
                .Interior.Color = RGB(240, 240, 240) ' 薄い白（平日）
                Debug.Print "日付: " & currentDate & " は平日（薄い白設定）"
            End If
        End With
        currentDate = currentDate + 1
        i = i + 1
    Loop

    Debug.Print "カレンダー背景色設定のテスト終了"
End Sub



Sub ResetTargetRangeByDate()
    Dim ws As Worksheet
    Dim startDate As Date, endDate As Date

    ' シートの取得
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("進捗カレンダー")
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "シート「進捗カレンダー」が見つかりません。", vbCritical
        Exit Sub
    End If

    ' 開始日と終了日を取得
    startDate = Application.InputBox("新しい開始日を入力してください (YYYY/MM/DD):", Type:=2)
    endDate = Application.InputBox("新しい終了日を入力してください (YYYY/MM/DD):", Type:=2)

    If Not IsDate(startDate) Or Not IsDate(endDate) Then
        MsgBox "日付が正しくありません。終了します。", vbExclamation
        Exit Sub
    End If

    If startDate > endDate Then
        MsgBox "開始日は終了日より前に設定してください。", vbExclamation
        Exit Sub
    End If

    ' D1セルに「期間：開始日～終了日」を表示
    ws.Cells(1, 4).Value = "期間：" & Format(startDate, "yyyy/mm/dd") & " ～ " & Format(endDate, "yyyy/mm/dd")

    ' カレンダー範囲の更新
    UpdateCalendar ws, startDate, endDate
    MsgBox "カレンダーを新しい日付範囲で更新しました！", vbInformation
End Sub

Sub SetDateValidation(ws As Worksheet, startDate As Date, endDate As Date)
    Dim dateRange As Range

    ' 日付の入力範囲を指定（予定開始日と終了日）
    Set dateRange = ws.Range("E3:F" & ws.Cells(ws.Rows.Count, 5).End(xlUp).Row)

    ' 日付の範囲を指定
    With dateRange.Validation
        .Delete ' 既存の入力規則を削除
        .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=startDate, Formula2:=endDate
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "日付を入力してください"
        .ErrorTitle = "無効な入力"
        .InputMessage = "日付を選択または入力（範囲内）"
        .ErrorMessage = "有効な日付を入力してください（範囲内）"
    End With
End Sub



Sub AddSampleData(ws As Worksheet, sampleRows As Integer)
    ' サンプルデータを指定した行数だけ入力
    Dim i As Integer
    Dim startDate As Date
    
    startDate = DateSerial(2024, 12, 1) ' 開始日の基準

    For i = 3 To 2 + sampleRows ' サンプルデータを指定された行数分入力
        ws.Cells(i, 2).Value = "サンプルタスク" & (i - 2)
        ws.Cells(i, 3).Value = "担当者" & (i - 2)
        ws.Cells(i, 5).Value = startDate + (i - 3) * 5 ' 5日間隔で開始日を設定
        ws.Cells(i, 6).Value = startDate + (i - 3) * 5 + 3 ' 開始日+3日を終了日に設定
    Next i
End Sub



Sub UpdateProgressAndHighlightCalendar()
    Dim ws As Worksheet
    Dim lastRow As Integer
    Dim taskRow As Integer
    Dim startPlan As Date, endPlan As Date
    Dim progressStatus As String
    Dim calendarStartCol As Integer, calendarEndCol As Integer
    Dim calendarRow As Integer, col As Integer

    ' シートの取得
    Set ws = ThisWorkbook.Sheets("進捗カレンダー")

    ' カレンダーの開始列と終了列を取得
    calendarRow = 2 ' カレンダーの日付が入る行
    calendarStartCol = 7 ' カレンダー開始列
    calendarEndCol = ws.Cells(calendarRow, ws.Columns.Count).End(xlToLeft).Column

    ' データ範囲の最終行を取得（列EとFのうち最も下の行を自動検出）
    lastRow = Application.WorksheetFunction.Max(ws.Cells(ws.Rows.Count, 5).End(xlUp).Row, ws.Cells(ws.Rows.Count, 6).End(xlUp).Row)

    ' デバッグログ
    Debug.Print "[DEBUG] 最終行: " & lastRow

    ' カレンダー列全体の色をリセット
    For taskRow = 3 To lastRow
        For col = calendarStartCol To calendarEndCol
            ws.Cells(taskRow, col).Interior.ColorIndex = xlNone
        Next col
    Next taskRow

    ' 各タスクの進捗状況を更新し、カレンダーをハイライト
    For taskRow = 3 To lastRow
        ' 開始日または終了日が空白の場合、その行をスキップ
        If IsEmpty(ws.Cells(taskRow, 5).Value) Or IsEmpty(ws.Cells(taskRow, 6).Value) Then
            Debug.Print "[DEBUG] 行: " & taskRow & " - 開始日または終了日が空白のためスキップ"
            ws.Cells(taskRow, 4).Value = "未設定"
            ws.Cells(taskRow, 4).Interior.ColorIndex = xlNone ' 色なし
            GoTo NextTask
        End If

        ' 予定開始日と終了日の取得
        If IsDate(ws.Cells(taskRow, 5).Value) And IsDate(ws.Cells(taskRow, 6).Value) Then
            startPlan = ws.Cells(taskRow, 5).Value
            endPlan = ws.Cells(taskRow, 6).Value
            Debug.Print "[DEBUG] 行: " & taskRow & " - 開始日: " & startPlan & " | 終了日: " & endPlan

            ' 現在の日付と比較して進捗状況を判定
            If Date < startPlan Then
                progressStatus = "未着"
            ElseIf Date >= startPlan And Date <= endPlan Then
                progressStatus = "処理中"
            ElseIf Date > endPlan Then
                progressStatus = "終了済み"
            End If

            ' セルに進捗状況を記入
            ws.Cells(taskRow, 4).Value = progressStatus

            ' 色分け処理
            If progressStatus = "終了済み" Then
                ws.Cells(taskRow, 4).Interior.Color = RGB(255, 0, 0) ' 赤
            ElseIf progressStatus = "処理中" Then
                ws.Cells(taskRow, 4).Interior.Color = RGB(255, 255, 0) ' 黄色
            Else
                ws.Cells(taskRow, 4).Interior.ColorIndex = xlNone ' 色なし
            End If

            ' カレンダーをハイライト（各タスクの行に対応）
            For col = calendarStartCol To calendarEndCol
                If IsDate(ws.Cells(calendarRow, col).Value) Then
                    Debug.Print "[DEBUG] カレンダー日付: " & ws.Cells(calendarRow, col).Value
                    If ws.Cells(calendarRow, col).Value >= startPlan And ws.Cells(calendarRow, col).Value <= endPlan Then
                        Debug.Print "[DEBUG] ハイライト: 行 " & taskRow & " 列 " & col
                        ws.Cells(taskRow, col).Interior.Color = RGB(173, 216, 230) ' 青
                    End If
                Else
                    Debug.Print "[DEBUG] 列 " & col & " のカレンダー日付が無効"
                End If
            Next col
        Else
            Debug.Print "[DEBUG] 行: " & taskRow & " - 日付形式が無効"
            ws.Cells(taskRow, 4).Value = "未設定"
            ws.Cells(taskRow, 4).Interior.ColorIndex = xlNone ' 色なし
        End If

NextTask:
    Next taskRow

    ' カレンダー領域に罫線を付与（データがある行のみ）
    With ws.Range(ws.Cells(3, calendarStartCol), ws.Cells(lastRow, calendarEndCol))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    ' 工数を計算して設定
    CalculateEffort ws


    MsgBox "進捗状況とカレンダーの更新が完了しました！", vbInformation
End Sub




Sub UpdateProgress()
    Dim ws As Worksheet
    Dim lastRowTask As Integer
    Dim lastRowStartDate As Integer
    Dim lastRow As Integer
    Dim taskRow As Integer
    Dim startPlan As Date, endPlan As Date
    Dim progressStatus As String

    ' シートの取得
    Set ws = ThisWorkbook.Sheets("進捗カレンダー")

    ' 最終行の取得
    lastRowTask = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    lastRowStartDate = ws.Cells(ws.Rows.Count, 5).End(xlUp).Row
    lastRow = Application.WorksheetFunction.Max(lastRowTask, lastRowStartDate)

    Debug.Print "[DEBUG] Last row calculated: " & lastRow

    For taskRow = 3 To lastRow
        ' 開始日または終了日が空白の場合、その行をスキップ
        If IsEmpty(ws.Cells(taskRow, 5).Value) Or IsEmpty(ws.Cells(taskRow, 6).Value) Then
            Debug.Print "[DEBUG] Row: " & taskRow & " - Start or End Date is empty. Skipping."
            ws.Cells(taskRow, 4).Value = "未設定"
            ws.Cells(taskRow, 4).Interior.ColorIndex = xlNone ' 色なし
            GoTo NextRow
        End If

        ' 予定開始日と終了日の取得
        If IsDate(ws.Cells(taskRow, 5).Value) And IsDate(ws.Cells(taskRow, 6).Value) Then
            startPlan = ws.Cells(taskRow, 5).Value
            endPlan = ws.Cells(taskRow, 6).Value
            Debug.Print "[DEBUG] Row: " & taskRow & " - Start: " & startPlan & " | End: " & endPlan

            ' 現在の日付と比較して進捗状況を判定
            If Date < startPlan Then
                progressStatus = "未着"
            ElseIf Date >= startPlan And Date <= endPlan Then
                progressStatus = "処理中"
            ElseIf Date > endPlan Then
                progressStatus = "終了済み"
            End If

            ' セルに進捗状況を記入
            ws.Cells(taskRow, 4).Value = progressStatus

            ' 色分け処理
            If progressStatus = "終了済み" Then
                ws.Cells(taskRow, 4).Interior.Color = RGB(255, 0, 0) ' 赤
            ElseIf progressStatus = "処理中" Then
                ws.Cells(taskRow, 4).Interior.Color = RGB(255, 255, 0) ' 黄色
            Else
                ws.Cells(taskRow, 4).Interior.ColorIndex = xlNone ' 色なし
            End If
        Else
            Debug.Print "[DEBUG] Row: " & taskRow & " - Invalid date format in Start or End Date."
            ws.Cells(taskRow, 4).Value = "未設定"
            ws.Cells(taskRow, 4).Interior.ColorIndex = xlNone ' 色なし
        End If

NextRow:
    Next taskRow
    
    ' 工数を計算して設定
    CalculateEffort ws

    MsgBox "進捗状況を更新しました！", vbInformation
End Sub



Sub UpdateCalendar(ws As Worksheet, startDate As Date, endDate As Date)
    Dim rowOffset As Integer, colOffset As Integer
    Dim currentDate As Date
    Dim i As Integer
    Dim calendarStartCol As Integer
    Dim lastRow As Long

    ' カレンダーエリアのクリア
    rowOffset = 2 ' カレンダーの日付が表示される行
    colOffset = 7 ' カレンダーの開始列
    currentDate = startDate
    i = 0

    ' カレンダーの日付を再表示
    Do While currentDate <= endDate
        With ws.Cells(rowOffset, colOffset + i)
            .Value = currentDate
            .NumberFormat = "MM/DD" ' 日付をMM/DD形式で表示
            .Interior.ColorIndex = xlNone ' 色をリセット
            .Borders.LineStyle = xlContinuous ' 罫線を再設定
        End With

        ' 土日の色付け
        Select Case Weekday(currentDate)
            Case vbSaturday
                ws.Cells(rowOffset, colOffset + i).Interior.Color = RGB(173, 216, 230) ' 青
            Case vbSunday
                ws.Cells(rowOffset, colOffset + i).Interior.Color = RGB(255, 182, 193) ' 赤
        End Select

        currentDate = currentDate + 1
        i = i + 1
    Loop

    ' カレンダー列の余分な部分をクリア
    ws.Range(ws.Cells(rowOffset, colOffset + i), ws.Cells(rowOffset, ws.Columns.Count)).ClearContents

    ' 最終行を取得してカレンダーをハイライト
    lastRow = ws.Cells(ws.Rows.Count, 5).End(xlUp).Row
    UpdateProgressAndHighlightCalendar
End Sub


Sub FormatDateColumns(ws As Worksheet)
    Dim dateRange As Range

    ' 対象範囲の設定（予定開始日と終了日）
    Set dateRange = ws.Range("E3:F20") ' 必要に応じて範囲を変更

    ' 日付形式を設定
    dateRange.NumberFormat = "yyyy/mm/dd"

    MsgBox "日付形式が設定されました！", vbInformation
End Sub

Function IsHoliday(ByVal d As Date) As Boolean
    Dim holidays As Collection
    Dim holidayDate As Variant
    Set holidays = New Collection

    ' 2024年の祝日
    holidays.Add DateSerial(2024, 1, 1) ' 元日
    holidays.Add DateSerial(2024, 1, 8) ' 成人の日
    holidays.Add DateSerial(2024, 2, 11) ' 建国記念の日
    holidays.Add DateSerial(2024, 2, 23) ' 天皇誕生日
    holidays.Add DateSerial(2024, 3, 20) ' 春分の日
    holidays.Add DateSerial(2024, 4, 29) ' 昭和の日
    holidays.Add DateSerial(2024, 5, 3) ' 憲法記念日
    holidays.Add DateSerial(2024, 5, 4) ' みどりの日
    holidays.Add DateSerial(2024, 5, 5) ' こどもの日
    holidays.Add DateSerial(2024, 7, 15) ' 海の日
    holidays.Add DateSerial(2024, 8, 11) ' 山の日
    holidays.Add DateSerial(2024, 9, 16) ' 敬老の日
    holidays.Add DateSerial(2024, 9, 22) ' 秋分の日
    holidays.Add DateSerial(2024, 10, 14) ' スポーツの日
    holidays.Add DateSerial(2024, 11, 3) ' 文化の日
    holidays.Add DateSerial(2024, 11, 23) ' 勤労感謝の日

    ' 2025年の祝日
    holidays.Add DateSerial(2025, 1, 1) ' 元日
    holidays.Add DateSerial(2025, 1, 13) ' 成人の日
    holidays.Add DateSerial(2025, 2, 11) ' 建国記念の日
    holidays.Add DateSerial(2025, 2, 23) ' 天皇誕生日
    holidays.Add DateSerial(2025, 3, 20) ' 春分の日
    holidays.Add DateSerial(2025, 4, 29) ' 昭和の日
    holidays.Add DateSerial(2025, 5, 3) ' 憲法記念日
    holidays.Add DateSerial(2025, 5, 4) ' みどりの日
    holidays.Add DateSerial(2025, 5, 5) ' こどもの日
    holidays.Add DateSerial(2025, 7, 21) ' 海の日
    holidays.Add DateSerial(2025, 8, 11) ' 山の日
    holidays.Add DateSerial(2025, 9, 15) ' 敬老の日
    holidays.Add DateSerial(2025, 9, 23) ' 秋分の日
    holidays.Add DateSerial(2025, 10, 13) ' スポーツの日
    holidays.Add DateSerial(2025, 11, 3) ' 文化の日
    holidays.Add DateSerial(2025, 11, 23) ' 勤労感謝の日
    holidays.Add DateSerial(2025, 11, 24) ' 勤労感謝の日の振替休日

    ' 判定
    IsHoliday = False
    For Each holidayDate In holidays
        If holidayDate = d Then
            IsHoliday = True
            Exit Function
        End If
    Next holidayDate
End Function



Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("進捗カレンダー")

    Debug.Print "[DEBUG] Change detected in cell(s): " & Target.Address

    ' 列E（予定開始日）または列F（予定終了日）が変更された場合に進捗状況を更新
    Dim dataRange As Range
    Set dataRange = ws.Range("E3:F" & ws.Cells(ws.Rows.Count, "E").End(xlUp).Row)

    If Not Intersect(Target, dataRange) Is Nothing Then
        On Error GoTo ErrorHandler
        Application.EnableEvents = False ' イベントループを防止
        Debug.Print "[DEBUG] Relevant change detected in range: " & dataRange.Address
        UpdateProgressAndHighlightCalendar
    Else
        Debug.Print "[DEBUG] Change ignored. Out of range."
    End If

Cleanup:
    Application.EnableEvents = True ' イベントを再有効化
    Exit Sub

ErrorHandler:
    Debug.Print "[ERROR] An error occurred: " & Err.Description
    Resume Cleanup
End Sub

Sub TestIsHoliday()
    Dim testDates As Variant
    Dim i As Integer
    Dim result As Boolean

    ' テストする日付のリスト
    testDates = Array(DateSerial(2024, 1, 1), DateSerial(2024, 2, 11), _
                      DateSerial(2024, 4, 29), DateSerial(2024, 5, 3), _
                      DateSerial(2025, 1, 1), DateSerial(2025, 2, 11), _
                      DateSerial(2025, 4, 29), DateSerial(2025, 9, 23))

    Debug.Print "祝日判定テスト開始..."
    For i = LBound(testDates) To UBound(testDates)
        result = IsHoliday(testDates(i))
        Debug.Print "日付: " & testDates(i) & " | 祝日: " & result
    Next i
    Debug.Print "祝日判定テスト終了"
End Sub


