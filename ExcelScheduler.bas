Sub CreateCalendar()
    Dim ws As Worksheet
    Dim startDate As Date, endDate As Date
    Dim currentDate As Date
    Dim rowOffset As Integer, colOffset As Integer
    Dim i As Integer

    ' �V�[�g���쐬
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "�i���J�����_�["

    ' A1�Ƀ^�C�g�����o��
    ws.Cells(1, 1).Value = "�i���J�����_�["
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 16

    ' D1�Ɂu���ԁF�J�n���`�I�����v��\��
    ws.Cells(1, 4).Value = "���ԁF"
    ws.Cells(1, 4).Font.Bold = True

    ' �ݒ�G���A���쐬
    ws.Cells(2, 1).Value = "�A��"
    ws.Cells(2, 2).Value = "�^�X�N��"
    ws.Cells(2, 3).Value = "�S����"
    ws.Cells(2, 4).Value = "�i����"
    ws.Cells(2, 5).Value = "�\��J�n��"
    ws.Cells(2, 6).Value = "�\��I����"
    ws.Cells(2, 7).Value = "�H���i���j" ' �H�����ǉ�

    ' �w�b�_�[�̔w�i�F�ƌr����ݒ�
    With ws.Rows(2)
        .Font.Bold = True
        .Interior.Color = RGB(173, 216, 230) ' ���邢���ۂ��F
        .Borders.LineStyle = xlContinuous
    End With

    ' �񕝂̒���
    ws.Columns(1).ColumnWidth = 5
    ws.Columns(2).ColumnWidth = 20
    ws.Columns(3).ColumnWidth = 15
    ws.Columns(4).ColumnWidth = 15
    ws.Columns(5).ColumnWidth = 15
    ws.Columns(6).ColumnWidth = 15
    ws.Columns(7).ColumnWidth = 10 ' �H����̕��𒲐�

    ' �f�t�H���g��100�s���̃t�H�[�}�b�g��ݒ�
    Dim taskRow As Integer
    For taskRow = 3 To 102 ' 100�s��
        ws.Cells(taskRow, 1).Value = taskRow - 2 ' �A�Ԃ���������
        ws.Cells(taskRow, 4).Value = "���ݒ�" ' ������ԂŁu���ݒ�v�����
    Next taskRow

    ' ���f�[�^��1�s�̂ݓ���
    AddSampleData ws, 1 ' �T���v���f�[�^��1�s�����ǉ�

    ' �J�����_�[�ݒ�G���A
    startDate = Application.InputBox("�J�����_�[�̊J�n������͂��Ă������� (YYYY/MM/DD):", Type:=2)
    endDate = Application.InputBox("�J�����_�[�̏I��������͂��Ă������� (YYYY/MM/DD):", Type:=2)

    If Not IsDate(startDate) Or Not IsDate(endDate) Then
        MsgBox "���t������������܂���B�I�����܂��B", vbExclamation
        Exit Sub
    End If

    If startDate > endDate Then
        MsgBox "�J�n���͏I�������O�ɐݒ肵�Ă��������B", vbExclamation
        Exit Sub
    End If

    ' D1�Z���Ɂu���ԁF�J�n���`�I�����v��\��
    ws.Cells(1, 4).Value = "���ԁF" & Format(startDate, "yyyy/mm/dd") & " �` " & Format(endDate, "yyyy/mm/dd")

    ' �J�����_�[�̔w�i�F�Ɠ��t��ݒ�
    UpdateCalendarBackground ws, startDate, endDate

    ' �f�[�^�̓��͋K����ݒ�
    SetDateValidation ws, startDate, endDate

    ' �H�����v�Z���Đݒ�
    CalculateEffort ws

    MsgBox "�i���J�����_�[���쐬���܂����I", vbInformation
End Sub

Sub CalculateEffort(ws As Worksheet)
    Dim lastRow As Long
    Dim taskRow As Long
    Dim startDate As Date, endDate As Date
    Dim effort As Long

    ' �f�[�^�͈͂̍ŏI�s���擾
    lastRow = ws.Cells(ws.Rows.Count, 5).End(xlUp).Row

    ' �e�s�̍H�����v�Z
    For taskRow = 3 To lastRow
        If IsDate(ws.Cells(taskRow, 5).Value) And IsDate(ws.Cells(taskRow, 6).Value) Then
            startDate = ws.Cells(taskRow, 5).Value
            endDate = ws.Cells(taskRow, 6).Value

            ' �H���𕽓����j�������O���Čv�Z
            effort = 0
            Dim currentDate As Date
            For currentDate = startDate To endDate
                If Weekday(currentDate, vbMonday) <= 5 And Not IsHoliday(currentDate) Then ' �����iMonday=1, Friday=5�j���j���łȂ�
                    effort = effort + 1
                End If
            Next currentDate

            ws.Cells(taskRow, 7).Value = effort
        Else
            ws.Cells(taskRow, 7).Value = "" ' �J�n���܂��͏I�����������̏ꍇ�͋�
        End If
    Next taskRow

    MsgBox "�H���̌v�Z���������܂����I", vbInformation
End Sub

Sub UpdateCalendarBackground(ws As Worksheet, startDate As Date, endDate As Date)
    Dim currentDate As Date
    Dim rowOffset As Integer, colOffset As Integer
    Dim i As Integer

    ' �J�����_�[�̊J�n�ʒu
    rowOffset = 2
    colOffset = 8 ' �H���� (7���) �̎��̗񂩂�J�n����悤�ɒ���
    currentDate = startDate
    i = 0

    Debug.Print "�J�����_�[�w�i�F�ݒ�̃e�X�g�J�n..."
    Do While currentDate <= endDate
        With ws.Cells(rowOffset, colOffset + i)
            .Value = currentDate
            .NumberFormat = "MM/DD"
            .Borders.LineStyle = xlContinuous

            ' �F�ݒ�̏���
            If IsHoliday(currentDate) Then
                .Interior.Color = RGB(255, 102, 102) ' �ԁi�j���j
                Debug.Print "���t: " & currentDate & " �͏j���i�ԐF�ݒ�j"
            ElseIf Weekday(currentDate) = vbSunday Then
                .Interior.Color = RGB(255, 182, 193) ' �����ԁi���j���j
                Debug.Print "���t: " & currentDate & " �͓��j���i�����Ԑݒ�j"
            ElseIf Weekday(currentDate) = vbSaturday Then
                .Interior.Color = RGB(173, 216, 230) ' �����i�y�j���j
                Debug.Print "���t: " & currentDate & " �͓y�j���i�����ݒ�j"
            Else
                .Interior.Color = RGB(240, 240, 240) ' �������i�����j
                Debug.Print "���t: " & currentDate & " �͕����i�������ݒ�j"
            End If
        End With
        currentDate = currentDate + 1
        i = i + 1
    Loop

    Debug.Print "�J�����_�[�w�i�F�ݒ�̃e�X�g�I��"
End Sub



Sub ResetTargetRangeByDate()
    Dim ws As Worksheet
    Dim startDate As Date, endDate As Date

    ' �V�[�g�̎擾
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("�i���J�����_�[")
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "�V�[�g�u�i���J�����_�[�v��������܂���B", vbCritical
        Exit Sub
    End If

    ' �J�n���ƏI�������擾
    startDate = Application.InputBox("�V�����J�n������͂��Ă������� (YYYY/MM/DD):", Type:=2)
    endDate = Application.InputBox("�V�����I��������͂��Ă������� (YYYY/MM/DD):", Type:=2)

    If Not IsDate(startDate) Or Not IsDate(endDate) Then
        MsgBox "���t������������܂���B�I�����܂��B", vbExclamation
        Exit Sub
    End If

    If startDate > endDate Then
        MsgBox "�J�n���͏I�������O�ɐݒ肵�Ă��������B", vbExclamation
        Exit Sub
    End If

    ' D1�Z���Ɂu���ԁF�J�n���`�I�����v��\��
    ws.Cells(1, 4).Value = "���ԁF" & Format(startDate, "yyyy/mm/dd") & " �` " & Format(endDate, "yyyy/mm/dd")

    ' �J�����_�[�͈͂̍X�V
    UpdateCalendar ws, startDate, endDate
    MsgBox "�J�����_�[��V�������t�͈͂ōX�V���܂����I", vbInformation
End Sub

Sub SetDateValidation(ws As Worksheet, startDate As Date, endDate As Date)
    Dim dateRange As Range

    ' ���t�̓��͔͈͂��w��i�\��J�n���ƏI�����j
    Set dateRange = ws.Range("E3:F" & ws.Cells(ws.Rows.Count, 5).End(xlUp).Row)

    ' ���t�͈̔͂��w��
    With dateRange.Validation
        .Delete ' �����̓��͋K�����폜
        .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=startDate, Formula2:=endDate
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "���t����͂��Ă�������"
        .ErrorTitle = "�����ȓ���"
        .InputMessage = "���t��I���܂��͓��́i�͈͓��j"
        .ErrorMessage = "�L���ȓ��t����͂��Ă��������i�͈͓��j"
    End With
End Sub



Sub AddSampleData(ws As Worksheet, sampleRows As Integer)
    ' �T���v���f�[�^���w�肵���s����������
    Dim i As Integer
    Dim startDate As Date
    
    startDate = DateSerial(2024, 12, 1) ' �J�n���̊

    For i = 3 To 2 + sampleRows ' �T���v���f�[�^���w�肳�ꂽ�s��������
        ws.Cells(i, 2).Value = "�T���v���^�X�N" & (i - 2)
        ws.Cells(i, 3).Value = "�S����" & (i - 2)
        ws.Cells(i, 5).Value = startDate + (i - 3) * 5 ' 5���Ԋu�ŊJ�n����ݒ�
        ws.Cells(i, 6).Value = startDate + (i - 3) * 5 + 3 ' �J�n��+3�����I�����ɐݒ�
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

    ' �V�[�g�̎擾
    Set ws = ThisWorkbook.Sheets("�i���J�����_�[")

    ' �J�����_�[�̊J�n��ƏI������擾
    calendarRow = 2 ' �J�����_�[�̓��t������s
    calendarStartCol = 7 ' �J�����_�[�J�n��
    calendarEndCol = ws.Cells(calendarRow, ws.Columns.Count).End(xlToLeft).Column

    ' �f�[�^�͈͂̍ŏI�s���擾�i��E��F�̂����ł����̍s���������o�j
    lastRow = Application.WorksheetFunction.Max(ws.Cells(ws.Rows.Count, 5).End(xlUp).Row, ws.Cells(ws.Rows.Count, 6).End(xlUp).Row)

    ' �f�o�b�O���O
    Debug.Print "[DEBUG] �ŏI�s: " & lastRow

    ' �J�����_�[��S�̂̐F�����Z�b�g
    For taskRow = 3 To lastRow
        For col = calendarStartCol To calendarEndCol
            ws.Cells(taskRow, col).Interior.ColorIndex = xlNone
        Next col
    Next taskRow

    ' �e�^�X�N�̐i���󋵂��X�V���A�J�����_�[���n�C���C�g
    For taskRow = 3 To lastRow
        ' �J�n���܂��͏I�������󔒂̏ꍇ�A���̍s���X�L�b�v
        If IsEmpty(ws.Cells(taskRow, 5).Value) Or IsEmpty(ws.Cells(taskRow, 6).Value) Then
            Debug.Print "[DEBUG] �s: " & taskRow & " - �J�n���܂��͏I�������󔒂̂��߃X�L�b�v"
            ws.Cells(taskRow, 4).Value = "���ݒ�"
            ws.Cells(taskRow, 4).Interior.ColorIndex = xlNone ' �F�Ȃ�
            GoTo NextTask
        End If

        ' �\��J�n���ƏI�����̎擾
        If IsDate(ws.Cells(taskRow, 5).Value) And IsDate(ws.Cells(taskRow, 6).Value) Then
            startPlan = ws.Cells(taskRow, 5).Value
            endPlan = ws.Cells(taskRow, 6).Value
            Debug.Print "[DEBUG] �s: " & taskRow & " - �J�n��: " & startPlan & " | �I����: " & endPlan

            ' ���݂̓��t�Ɣ�r���Đi���󋵂𔻒�
            If Date < startPlan Then
                progressStatus = "����"
            ElseIf Date >= startPlan And Date <= endPlan Then
                progressStatus = "������"
            ElseIf Date > endPlan Then
                progressStatus = "�I���ς�"
            End If

            ' �Z���ɐi���󋵂��L��
            ws.Cells(taskRow, 4).Value = progressStatus

            ' �F��������
            If progressStatus = "�I���ς�" Then
                ws.Cells(taskRow, 4).Interior.Color = RGB(255, 0, 0) ' ��
            ElseIf progressStatus = "������" Then
                ws.Cells(taskRow, 4).Interior.Color = RGB(255, 255, 0) ' ���F
            Else
                ws.Cells(taskRow, 4).Interior.ColorIndex = xlNone ' �F�Ȃ�
            End If

            ' �J�����_�[���n�C���C�g�i�e�^�X�N�̍s�ɑΉ��j
            For col = calendarStartCol To calendarEndCol
                If IsDate(ws.Cells(calendarRow, col).Value) Then
                    Debug.Print "[DEBUG] �J�����_�[���t: " & ws.Cells(calendarRow, col).Value
                    If ws.Cells(calendarRow, col).Value >= startPlan And ws.Cells(calendarRow, col).Value <= endPlan Then
                        Debug.Print "[DEBUG] �n�C���C�g: �s " & taskRow & " �� " & col
                        ws.Cells(taskRow, col).Interior.Color = RGB(173, 216, 230) ' ��
                    End If
                Else
                    Debug.Print "[DEBUG] �� " & col & " �̃J�����_�[���t������"
                End If
            Next col
        Else
            Debug.Print "[DEBUG] �s: " & taskRow & " - ���t�`��������"
            ws.Cells(taskRow, 4).Value = "���ݒ�"
            ws.Cells(taskRow, 4).Interior.ColorIndex = xlNone ' �F�Ȃ�
        End If

NextTask:
    Next taskRow

    ' �J�����_�[�̈�Ɍr����t�^�i�f�[�^������s�̂݁j
    With ws.Range(ws.Cells(3, calendarStartCol), ws.Cells(lastRow, calendarEndCol))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    ' �H�����v�Z���Đݒ�
    CalculateEffort ws


    MsgBox "�i���󋵂ƃJ�����_�[�̍X�V���������܂����I", vbInformation
End Sub




Sub UpdateProgress()
    Dim ws As Worksheet
    Dim lastRowTask As Integer
    Dim lastRowStartDate As Integer
    Dim lastRow As Integer
    Dim taskRow As Integer
    Dim startPlan As Date, endPlan As Date
    Dim progressStatus As String

    ' �V�[�g�̎擾
    Set ws = ThisWorkbook.Sheets("�i���J�����_�[")

    ' �ŏI�s�̎擾
    lastRowTask = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    lastRowStartDate = ws.Cells(ws.Rows.Count, 5).End(xlUp).Row
    lastRow = Application.WorksheetFunction.Max(lastRowTask, lastRowStartDate)

    Debug.Print "[DEBUG] Last row calculated: " & lastRow

    For taskRow = 3 To lastRow
        ' �J�n���܂��͏I�������󔒂̏ꍇ�A���̍s���X�L�b�v
        If IsEmpty(ws.Cells(taskRow, 5).Value) Or IsEmpty(ws.Cells(taskRow, 6).Value) Then
            Debug.Print "[DEBUG] Row: " & taskRow & " - Start or End Date is empty. Skipping."
            ws.Cells(taskRow, 4).Value = "���ݒ�"
            ws.Cells(taskRow, 4).Interior.ColorIndex = xlNone ' �F�Ȃ�
            GoTo NextRow
        End If

        ' �\��J�n���ƏI�����̎擾
        If IsDate(ws.Cells(taskRow, 5).Value) And IsDate(ws.Cells(taskRow, 6).Value) Then
            startPlan = ws.Cells(taskRow, 5).Value
            endPlan = ws.Cells(taskRow, 6).Value
            Debug.Print "[DEBUG] Row: " & taskRow & " - Start: " & startPlan & " | End: " & endPlan

            ' ���݂̓��t�Ɣ�r���Đi���󋵂𔻒�
            If Date < startPlan Then
                progressStatus = "����"
            ElseIf Date >= startPlan And Date <= endPlan Then
                progressStatus = "������"
            ElseIf Date > endPlan Then
                progressStatus = "�I���ς�"
            End If

            ' �Z���ɐi���󋵂��L��
            ws.Cells(taskRow, 4).Value = progressStatus

            ' �F��������
            If progressStatus = "�I���ς�" Then
                ws.Cells(taskRow, 4).Interior.Color = RGB(255, 0, 0) ' ��
            ElseIf progressStatus = "������" Then
                ws.Cells(taskRow, 4).Interior.Color = RGB(255, 255, 0) ' ���F
            Else
                ws.Cells(taskRow, 4).Interior.ColorIndex = xlNone ' �F�Ȃ�
            End If
        Else
            Debug.Print "[DEBUG] Row: " & taskRow & " - Invalid date format in Start or End Date."
            ws.Cells(taskRow, 4).Value = "���ݒ�"
            ws.Cells(taskRow, 4).Interior.ColorIndex = xlNone ' �F�Ȃ�
        End If

NextRow:
    Next taskRow
    
    ' �H�����v�Z���Đݒ�
    CalculateEffort ws

    MsgBox "�i���󋵂��X�V���܂����I", vbInformation
End Sub



Sub UpdateCalendar(ws As Worksheet, startDate As Date, endDate As Date)
    Dim rowOffset As Integer, colOffset As Integer
    Dim currentDate As Date
    Dim i As Integer
    Dim calendarStartCol As Integer
    Dim lastRow As Long

    ' �J�����_�[�G���A�̃N���A
    rowOffset = 2 ' �J�����_�[�̓��t���\�������s
    colOffset = 7 ' �J�����_�[�̊J�n��
    currentDate = startDate
    i = 0

    ' �J�����_�[�̓��t���ĕ\��
    Do While currentDate <= endDate
        With ws.Cells(rowOffset, colOffset + i)
            .Value = currentDate
            .NumberFormat = "MM/DD" ' ���t��MM/DD�`���ŕ\��
            .Interior.ColorIndex = xlNone ' �F�����Z�b�g
            .Borders.LineStyle = xlContinuous ' �r�����Đݒ�
        End With

        ' �y���̐F�t��
        Select Case Weekday(currentDate)
            Case vbSaturday
                ws.Cells(rowOffset, colOffset + i).Interior.Color = RGB(173, 216, 230) ' ��
            Case vbSunday
                ws.Cells(rowOffset, colOffset + i).Interior.Color = RGB(255, 182, 193) ' ��
        End Select

        currentDate = currentDate + 1
        i = i + 1
    Loop

    ' �J�����_�[��̗]���ȕ������N���A
    ws.Range(ws.Cells(rowOffset, colOffset + i), ws.Cells(rowOffset, ws.Columns.Count)).ClearContents

    ' �ŏI�s���擾���ăJ�����_�[���n�C���C�g
    lastRow = ws.Cells(ws.Rows.Count, 5).End(xlUp).Row
    UpdateProgressAndHighlightCalendar
End Sub


Sub FormatDateColumns(ws As Worksheet)
    Dim dateRange As Range

    ' �Ώ۔͈͂̐ݒ�i�\��J�n���ƏI�����j
    Set dateRange = ws.Range("E3:F20") ' �K�v�ɉ����Ĕ͈͂�ύX

    ' ���t�`����ݒ�
    dateRange.NumberFormat = "yyyy/mm/dd"

    MsgBox "���t�`�����ݒ肳��܂����I", vbInformation
End Sub

Function IsHoliday(ByVal d As Date) As Boolean
    Dim holidays As Collection
    Dim holidayDate As Variant
    Set holidays = New Collection

    ' 2024�N�̏j��
    holidays.Add DateSerial(2024, 1, 1) ' ����
    holidays.Add DateSerial(2024, 1, 8) ' ���l�̓�
    holidays.Add DateSerial(2024, 2, 11) ' �����L�O�̓�
    holidays.Add DateSerial(2024, 2, 23) ' �V�c�a����
    holidays.Add DateSerial(2024, 3, 20) ' �t���̓�
    holidays.Add DateSerial(2024, 4, 29) ' ���a�̓�
    holidays.Add DateSerial(2024, 5, 3) ' ���@�L�O��
    holidays.Add DateSerial(2024, 5, 4) ' �݂ǂ�̓�
    holidays.Add DateSerial(2024, 5, 5) ' ���ǂ��̓�
    holidays.Add DateSerial(2024, 7, 15) ' �C�̓�
    holidays.Add DateSerial(2024, 8, 11) ' �R�̓�
    holidays.Add DateSerial(2024, 9, 16) ' �h�V�̓�
    holidays.Add DateSerial(2024, 9, 22) ' �H���̓�
    holidays.Add DateSerial(2024, 10, 14) ' �X�|�[�c�̓�
    holidays.Add DateSerial(2024, 11, 3) ' �����̓�
    holidays.Add DateSerial(2024, 11, 23) ' �ΘJ���ӂ̓�

    ' 2025�N�̏j��
    holidays.Add DateSerial(2025, 1, 1) ' ����
    holidays.Add DateSerial(2025, 1, 13) ' ���l�̓�
    holidays.Add DateSerial(2025, 2, 11) ' �����L�O�̓�
    holidays.Add DateSerial(2025, 2, 23) ' �V�c�a����
    holidays.Add DateSerial(2025, 3, 20) ' �t���̓�
    holidays.Add DateSerial(2025, 4, 29) ' ���a�̓�
    holidays.Add DateSerial(2025, 5, 3) ' ���@�L�O��
    holidays.Add DateSerial(2025, 5, 4) ' �݂ǂ�̓�
    holidays.Add DateSerial(2025, 5, 5) ' ���ǂ��̓�
    holidays.Add DateSerial(2025, 7, 21) ' �C�̓�
    holidays.Add DateSerial(2025, 8, 11) ' �R�̓�
    holidays.Add DateSerial(2025, 9, 15) ' �h�V�̓�
    holidays.Add DateSerial(2025, 9, 23) ' �H���̓�
    holidays.Add DateSerial(2025, 10, 13) ' �X�|�[�c�̓�
    holidays.Add DateSerial(2025, 11, 3) ' �����̓�
    holidays.Add DateSerial(2025, 11, 23) ' �ΘJ���ӂ̓�
    holidays.Add DateSerial(2025, 11, 24) ' �ΘJ���ӂ̓��̐U�֋x��

    ' ����
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
    Set ws = ThisWorkbook.Sheets("�i���J�����_�[")

    Debug.Print "[DEBUG] Change detected in cell(s): " & Target.Address

    ' ��E�i�\��J�n���j�܂��͗�F�i�\��I�����j���ύX���ꂽ�ꍇ�ɐi���󋵂��X�V
    Dim dataRange As Range
    Set dataRange = ws.Range("E3:F" & ws.Cells(ws.Rows.Count, "E").End(xlUp).Row)

    If Not Intersect(Target, dataRange) Is Nothing Then
        On Error GoTo ErrorHandler
        Application.EnableEvents = False ' �C�x���g���[�v��h�~
        Debug.Print "[DEBUG] Relevant change detected in range: " & dataRange.Address
        UpdateProgressAndHighlightCalendar
    Else
        Debug.Print "[DEBUG] Change ignored. Out of range."
    End If

Cleanup:
    Application.EnableEvents = True ' �C�x���g���ėL����
    Exit Sub

ErrorHandler:
    Debug.Print "[ERROR] An error occurred: " & Err.Description
    Resume Cleanup
End Sub

Sub TestIsHoliday()
    Dim testDates As Variant
    Dim i As Integer
    Dim result As Boolean

    ' �e�X�g������t�̃��X�g
    testDates = Array(DateSerial(2024, 1, 1), DateSerial(2024, 2, 11), _
                      DateSerial(2024, 4, 29), DateSerial(2024, 5, 3), _
                      DateSerial(2025, 1, 1), DateSerial(2025, 2, 11), _
                      DateSerial(2025, 4, 29), DateSerial(2025, 9, 23))

    Debug.Print "�j������e�X�g�J�n..."
    For i = LBound(testDates) To UBound(testDates)
        result = IsHoliday(testDates(i))
        Debug.Print "���t: " & testDates(i) & " | �j��: " & result
    Next i
    Debug.Print "�j������e�X�g�I��"
End Sub


