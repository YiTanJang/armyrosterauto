Attribute VB_Name = "Module1"
Option Explicit

' ==============================================================================
' [모듈 1] 근무표 생성 메인 로직
' ==============================================================================
Sub 근무표_이어쓰기_생성()
    Call Common_Scheduler(False)
End Sub

' 메인 스케줄러 (신규 생성 및 이어쓰기 통합)
Sub Common_Scheduler(isReset As Boolean)
    Dim wsData As Worksheet, wsRoster As Worksheet, wsSetting As Worksheet
    Dim lastRow As Long, settingLastRow As Long, r As Long
    Dim startDate As Date, targetDate As Date
    Dim leaderRow As Long, helperRow As Long
    Dim isHoliday As Boolean
    Dim eventName As String, eventType As String
    Dim addDays As Integer, dayIdx As Integer
    Dim shiftRow As Long
    
    ' 근무 설정 변수
    Dim shiftName As String, reqCount As Integer, isMandatory As Boolean
    Dim minRank As String, banConsecutive As Boolean
    Dim preFixedLeader As String, preFixedHelper As String
    
    ' 일별 근무자 추적용 배열 (하루 중복 근무 방지)
    Dim dutyToday() As Boolean
    
    Call Save_Snapshot
    
    Set wsData = ThisWorkbook.Sheets("인원관리")
    Set wsRoster = ThisWorkbook.Sheets("근무표")
    Set wsSetting = ThisWorkbook.Sheets("설정")
    
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    
    ' 데이터 유무 확인
    If lastRow < 2 Then
        MsgBox "인원 데이터가 없습니다. 인원관리 시트에 명단을 추가해주세요.", vbExclamation
        Exit Sub
    End If
    
    ' 날짜 설정 로직
    If isReset Then
        addDays = InputBox("생성할 기간(일수)을 입력하세요.", "기간 설정", "7")
        If addDays <= 0 Then Exit Sub
        Dim inputDate As String
        inputDate = InputBox("시작 날짜는?", "날짜 설정", Format(Date + 1, "yyyy-mm-dd"))
        If Not IsDate(inputDate) Then Exit Sub
        startDate = CDate(inputDate)
        
        wsRoster.Cells.Clear
        wsRoster.Range("A1:E1").Value = Array("날짜", "요일", "근무명", "사수", "부사수")
        wsRoster.Range("A1:E1").Font.Bold = True
        wsRoster.Columns("A:E").ColumnWidth = 14
    Else
        addDays = InputBox("며칠치를 추가하시겠습니까?", "기간 설정", "7")
        If addDays <= 0 Then Exit Sub
        Dim lastRosterRow As Long
        lastRosterRow = wsRoster.Cells(wsRoster.Rows.Count, "A").End(xlUp).Row
        If lastRosterRow < 2 Then startDate = Date Else startDate = CDate(wsRoster.Cells(lastRosterRow, 1).Value) + 1
    End If

    ' 마지막 근무일 저장용 배열 (연일 근무 방지용)
    Dim lastDutyDate() As Long
    ReDim lastDutyDate(2 To lastRow)
    For r = 2 To lastRow: lastDutyDate(r) = 0: Next r
    
    ' 기존 근무표 내역을 스캔하여 '직전 근무일' 메모리 로드
    Dim rosterLast As Long, checkRow As Long
    Dim checkDate As Date, checkName As String
    
    rosterLast = wsRoster.Cells(wsRoster.Rows.Count, "A").End(xlUp).Row
    ' 최근 200행 역추적
    For checkRow = rosterLast To Application.Max(2, rosterLast - 200) Step -1
        If IsDate(wsRoster.Cells(checkRow, 1).Value) Then
            checkDate = CDate(wsRoster.Cells(checkRow, 1).Value)
            ' 사수(4열), 부사수(5열) 확인
            Dim cIdx As Variant
            For Each cIdx In Array(4, 5)
                checkName = wsRoster.Cells(checkRow, cIdx).Value
                If checkName <> "" And checkName <> "-" And checkName <> "인원부족" Then
                    Dim foundR As Range
                    Set foundR = wsData.Columns(1).Find(checkName, LookAt:=xlWhole)
                    If Not foundR Is Nothing Then
                        If lastDutyDate(foundR.Row) < CLng(checkDate) Then
                            lastDutyDate(foundR.Row) = CLng(checkDate)
                        End If
                    End If
                End If
            Next cIdx
        End If
    Next checkRow
    
    settingLastRow = wsSetting.Cells(wsSetting.Rows.Count, "I").End(xlUp).Row
    
    ' === 날짜 루프 ===
    For dayIdx = 0 To addDays - 1
        targetDate = startDate + dayIdx
        isHoliday = Check_Is_Holiday(targetDate, wsSetting)
        
        Call Get_Event_Info(targetDate, wsSetting, eventName, eventType)
        
        ' 오늘 근무자 명단 초기화
        ReDim dutyToday(2 To lastRow)
        For r = 2 To lastRow: dutyToday(r) = False: Next r
        
        ' === 근무 목록 루프 ===
        For shiftRow = 2 To settingLastRow
            shiftName = wsSetting.Cells(shiftRow, 9).Value
            reqCount = Val(wsSetting.Cells(shiftRow, 10).Value)
            If UCase(wsSetting.Cells(shiftRow, 14).Value) = "O" Then isMandatory = True Else isMandatory = False
            
            minRank = wsSetting.Cells(shiftRow, 15).Value       ' 최소계급
            If UCase(wsSetting.Cells(shiftRow, 16).Value) = "O" Then banConsecutive = True Else banConsecutive = False ' 연일금지
            
            Dim writeRow As Long
            writeRow = wsRoster.Cells(wsRoster.Rows.Count, "A").End(xlUp).Row + 1
            
            wsRoster.Cells(writeRow, 1).Value = targetDate
            wsRoster.Cells(writeRow, 1).NumberFormat = "mm-dd"
            wsRoster.Cells(writeRow, 2).Value = Format(targetDate, "aaa")
            wsRoster.Cells(writeRow, 3).Value = shiftName
            
            preFixedLeader = wsRoster.Cells(writeRow, 4).Value
            preFixedHelper = wsRoster.Cells(writeRow, 5).Value
            
            ' 일정에 따른 스킵 여부 확인
            Dim skipShift As Boolean: skipShift = False
            If eventName <> "" Then
                Select Case eventType
                    Case "전체휴무": skipShift = True
                    Case "필수만": If Not isMandatory Then skipShift = True
                    Case "정상근무": skipShift = False
                    Case Else: skipShift = True
                End Select
            End If
            
            If skipShift Then
                wsRoster.Cells(writeRow, 4).Value = eventName
                wsRoster.Cells(writeRow, 5).Value = eventName
            Else
                ' -----------------------------------------------------
                ' (1) 사수 배정
                ' -----------------------------------------------------
                leaderRow = 0
                If preFixedLeader <> "" Then
                    For r = 2 To lastRow
                        If wsData.Cells(r, 1).Value = preFixedLeader Then
                            leaderRow = r: Exit For
                        End If
                    Next r
                Else
                    ' 함수명 변경됨 (V8 제거)
                    leaderRow = Find_Best_Soldier(wsData, targetDate, lastDutyDate, dutyToday, _
                                                    True, minRank, banConsecutive)
                    
                    If leaderRow > 0 Then
                        wsRoster.Cells(writeRow, 4).Value = wsData.Cells(leaderRow, 1).Value
                        
                        dutyToday(leaderRow) = True            ' 오늘 근무 체크
                        lastDutyDate(leaderRow) = targetDate ' 마지막 근무일 갱신
                        wsData.Cells(leaderRow, 4).Value = Val(wsData.Cells(leaderRow, 4).Value) + 1
                    Else
                        wsRoster.Cells(writeRow, 4).Value = "인원부족"
                    End If
                End If
                
                ' -----------------------------------------------------
                ' (2) 부사수 배정
                ' -----------------------------------------------------
                If reqCount >= 2 Then
                    helperRow = 0
                    If preFixedHelper <> "" Then
                        For r = 2 To lastRow
                            If wsData.Cells(r, 1).Value = preFixedHelper Then
                                helperRow = r: Exit For
                            End If
                        Next r
                    Else
                        ' 함수명 변경됨 (V8 제거)
                        helperRow = Find_Best_Soldier(wsData, targetDate, lastDutyDate, dutyToday, _
                                                        False, "이병", banConsecutive)
                        
                        If helperRow > 0 Then
                            wsRoster.Cells(writeRow, 5).Value = wsData.Cells(helperRow, 1).Value
                            dutyToday(helperRow) = True
                            lastDutyDate(helperRow) = targetDate
                            wsData.Cells(helperRow, 4).Value = Val(wsData.Cells(helperRow, 4).Value) + 1
                        Else
                            wsRoster.Cells(writeRow, 5).Value = "인원부족"
                        End If
                    End If
                Else
                    wsRoster.Cells(writeRow, 5).Value = "-"
                    wsRoster.Cells(writeRow, 5).HorizontalAlignment = xlCenter
                End If
            End If
        Next shiftRow
        
        ' 날짜 구분선
        With wsRoster.Range(wsRoster.Cells(wsRoster.Rows.Count, 1), wsRoster.Cells(wsRoster.Rows.Count, 5)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous: .Weight = xlThin
        End With
    Next dayIdx
    
    Call 통계_강제_갱신
End Sub

' ==============================================================================
' [모듈 2] 최적 근무자 선발 로직
' ==============================================================================
Function Find_Best_Soldier(ws As Worksheet, tDate As Date, lastDutyDate() As Long, dutyToday() As Boolean, _
                              isLeader As Boolean, reqRank As String, isBanConsecutive As Boolean) As Long
    Dim r As Long, lastRow As Long
    Dim minScore As Double, currentScore As Double
    Dim bestRow As Long
    Dim serviceDays As Long, startDate As Date
    Dim workerName As String
    Dim myScore As Double
    Dim myRankStr As String, myRankNum As Integer, reqRankNum As Integer
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    minScore = 999999
    bestRow = 0
    
    reqRankNum = Get_Rank_Num(reqRank)
    
    Randomize
    
    For r = 2 To lastRow
        workerName = ws.Cells(r, 1).Value
        
        ' 1. 오늘 이미 근무했으면 제외
        If dutyToday(r) = True Then GoTo ContinueLoop
        
        ' 2. 열외/휴가자 제외
        If Check_Is_Excluded(workerName, tDate) Then GoTo ContinueLoop
        
        ' 3. 연일 근무 금지 체크
        If isBanConsecutive Then
            If lastDutyDate(r) = (tDate - 1) Then GoTo ContinueLoop
        End If
        
        ' 4. 계급 제한 체크
        myRankStr = ws.Cells(r, 2).Value
        myRankNum = Get_Rank_Num(myRankStr)
        If myRankNum < reqRankNum Then GoTo ContinueLoop
        
        ' 복무일수 계산
        If IsDate(ws.Cells(r, 9).Value) Then
            startDate = CDate(ws.Cells(r, 9).Value)
        ElseIf IsDate(ws.Cells(r, 3).Value) Then
            startDate = CDate(ws.Cells(r, 3).Value)
        Else
            GoTo ContinueLoop
        End If
        serviceDays = tDate - startDate
        If serviceDays < 1 Then serviceDays = 1
        
        ' 점수 기반 선발 (점수/복무일수)
        myScore = Val(ws.Cells(r, 11).Value)
        currentScore = myScore / serviceDays
        
        ' 랜덤 노이즈 추가 (다양성 확보)
        currentScore = currentScore * (1 + ((Rnd() * 0.05) - 0.025))
        
        If currentScore < minScore Then
            minScore = currentScore
            bestRow = r
        End If
        
ContinueLoop:
    Next r
    
    Find_Best_Soldier = bestRow
End Function

Function Get_Rank_Num(rankStr As String) As Integer
    Select Case Trim(rankStr)
        Case "병장": Get_Rank_Num = 4
        Case "상병": Get_Rank_Num = 3
        Case "일병": Get_Rank_Num = 2
        Case "이병": Get_Rank_Num = 1
        Case Else: Get_Rank_Num = 0
    End Select
End Function

' ==============================================================================
' [모듈 3] 통계 및 데이터 갱신
' ==============================================================================
Sub 통계_강제_갱신()
    Dim wsData As Worksheet, wsRoster As Worksheet, wsSetting As Worksheet
    Dim lastRowRoster As Long, lastRowData As Long
    Dim i As Long, r As Long
    Dim workerName As String, targetDate As Date
    Dim isHol As Boolean
    Dim shiftName As String, score As Double, key As String
    
    Dim dictShiftScore As Object
    Set dictShiftScore = CreateObject("Scripting.Dictionary")
    
    Dim dictScore As Object, dictCount As Object
    Set dictScore = CreateObject("Scripting.Dictionary")
    Set dictCount = CreateObject("Scripting.Dictionary")
    
    Set wsData = ThisWorkbook.Sheets("인원관리")
    Set wsRoster = ThisWorkbook.Sheets("근무표")
    Set wsSetting = ThisWorkbook.Sheets("설정")
    
    Application.ScreenUpdating = False
    
    ' 설정값 로드
    Dim setLast As Long
    setLast = wsSetting.Cells(wsSetting.Rows.Count, "I").End(xlUp).Row
    If setLast >= 2 Then
        For i = 2 To setLast
            shiftName = wsSetting.Cells(i, 9).Value
            dictShiftScore(shiftName & "_평일") = Val(wsSetting.Cells(i, 11).Value)
            dictShiftScore(shiftName & "_휴일") = Val(wsSetting.Cells(i, 12).Value)
        Next i
    End If
    
    ' 근무표 스캔
    lastRowRoster = wsRoster.Cells(wsRoster.Rows.Count, "A").End(xlUp).Row
    If lastRowRoster >= 2 Then
        For i = 2 To lastRowRoster
            If IsDate(wsRoster.Cells(i, 1).Value) Then
                targetDate = CDate(wsRoster.Cells(i, 1).Value)
                isHol = Check_Is_Holiday(targetDate, wsSetting)
                shiftName = wsRoster.Cells(i, 3).Value
                
                If isHol Then key = shiftName & "_휴일" Else key = shiftName & "_평일"
                If dictShiftScore.Exists(key) Then score = dictShiftScore(key) Else score = 1
                
                Dim cols As Variant, c As Variant
                cols = Array(4, 5)
                For Each c In cols
                    workerName = wsRoster.Cells(i, c).Value
                    If workerName <> "" And _
                       workerName <> "전원부재" And workerName <> "인원부족" And _
                       workerName <> "-" And _
                       InStr(workerName, "휴무") = 0 And _
                       InStr(workerName, "근무없음") = 0 Then
                        
                        dictScore(workerName) = dictScore(workerName) + score
                        dictCount(workerName) = dictCount(workerName) + 1
                    End If
                Next c
            End If
        Next i
    End If
    
    ' 인원관리 시트에 반영
    lastRowData = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    For r = 2 To lastRowData
        workerName = wsData.Cells(r, 1).Value
        Dim baseTotal As Long: baseTotal = Val(wsData.Cells(r, 6).Value)
        Dim rosterTotal As Long: If dictCount.Exists(workerName) Then rosterTotal = dictCount(workerName) Else rosterTotal = 0
        
        wsData.Cells(r, 4).Value = baseTotal + rosterTotal
        
        Dim currentScore As Double: If dictScore.Exists(workerName) Then currentScore = dictScore(workerName) Else currentScore = 0
        wsData.Cells(r, 11).Value = (baseTotal * 1#) + currentScore
    Next r
    
    Application.ScreenUpdating = True
End Sub

' ==============================================================================
' [모듈 4] 유틸리티 및 편의 기능
' ==============================================================================
Sub 근무_교체_마법사()
    Dim rng As Range, c1 As Range, c2 As Range, temp As String
    Set rng = Selection
    If rng.Count <> 2 Then MsgBox "교체할 두 셀을 Ctrl+클릭으로 선택하세요.", vbExclamation: Exit Sub
    Set c1 = rng.Areas(1).Cells(1, 1)
    If rng.Areas.Count > 1 Then Set c2 = rng.Areas(2).Cells(1, 1) Else Set c2 = rng.Cells(1, 2)
    temp = c1.Value: c1.Value = c2.Value: c2.Value = temp
    Call 통계_강제_갱신
    MsgBox "교체 완료!", vbInformation
End Sub

Sub 인원_자동_정렬()
    Dim wsData As Worksheet, lastRow As Long, r As Long
    Dim rankVal As String, rankScore As Integer
    Set wsData = ThisWorkbook.Sheets("인원관리")
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    Application.ScreenUpdating = False
    For r = 2 To lastRow
        rankVal = Trim(wsData.Cells(r, 2).Value)
        Select Case rankVal
            Case "병장": rankScore = 1
            Case "상병": rankScore = 2
            Case "일병": rankScore = 3
            Case "이병": rankScore = 4
            Case Else: rankScore = 9
        End Select
        wsData.Cells(r, 12).Value = rankScore
    Next r
    
    With wsData.Sort
        .SortFields.Clear
        .SortFields.Add key:=wsData.Range("L2:L" & lastRow), Order:=xlAscending
        .SortFields.Add key:=wsData.Range("I2:I" & lastRow), Order:=xlAscending
        .SetRange wsData.Range("A1:L" & lastRow)
        .Header = xlYes
        .Apply
    End With
    wsData.Columns(12).ClearContents
    Application.ScreenUpdating = True
    MsgBox "계급순, 전입일순 정렬 완료!", vbInformation
End Sub

Sub 전역자_자동_보내기()
    Dim wsData As Worksheet, wsArchive As Worksheet
    Dim lastRow As Long, r As Long, archiveRow As Long, moveCount As Long
    Dim outDate As Date
    
    Set wsData = ThisWorkbook.Sheets("인원관리")
    On Error Resume Next
    Set wsArchive = ThisWorkbook.Sheets("전역자")
    If wsArchive Is Nothing Then
        Set wsArchive = ThisWorkbook.Sheets.Add(After:=wsData)
        wsArchive.Name = "전역자"
        wsData.Rows(1).Copy wsArchive.Rows(1)
        wsArchive.Range("J1").Value = "처리일자"
    End If
    On Error GoTo 0
    
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    If MsgBox("H열(전역일)이 지난 인원을 정리하시겠습니까?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    
    Application.ScreenUpdating = False
    For r = lastRow To 2 Step -1
        If IsDate(wsData.Cells(r, 8).Value) Then
            outDate = CDate(wsData.Cells(r, 8).Value)
            If Date >= outDate Then
                archiveRow = wsArchive.Cells(wsArchive.Rows.Count, "A").End(xlUp).Row + 1
                wsData.Rows(r).Copy wsArchive.Rows(archiveRow)
                wsArchive.Cells(archiveRow, 10).Value = Format(Date, "yyyy-mm-dd")
                wsData.Rows(r).Delete
                moveCount = moveCount + 1
            End If
        End If
    Next r
    Application.ScreenUpdating = True
    MsgBox moveCount & "명 처리 완료.", vbInformation
End Sub

Sub Save_Snapshot()
    Dim wsBack As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    Set wsBack = ThisWorkbook.Sheets("Backup_Hidden")
    If wsBack Is Nothing Then
        Set wsBack = ThisWorkbook.Sheets.Add: wsBack.Name = "Backup_Hidden": wsBack.Visible = xlSheetHidden
    End If
    wsBack.Cells.Clear
    ThisWorkbook.Sheets("근무표").Cells.Copy wsBack.Cells
    Application.DisplayAlerts = True
End Sub

Sub Undo_Last_Action()
    If MsgBox("실행 취소 하시겠습니까?", vbYesNo) = vbNo Then Exit Sub
    ThisWorkbook.Sheets("근무표").Cells.Clear
    ThisWorkbook.Sheets("Backup_Hidden").Cells.Copy ThisWorkbook.Sheets("근무표").Cells
    Call 통계_강제_갱신
    MsgBox "복구 완료", vbInformation
End Sub

Sub 공정성_히트맵_적용()
    Dim wsData As Worksheet, rng As Range
    Set wsData = ThisWorkbook.Sheets("인원관리")
    If wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row < 2 Then Exit Sub
    Set rng = wsData.Range("K2:K" & wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row)
    rng.FormatConditions.Delete
    With rng.FormatConditions.AddColorScale(ColorScaleType:=3)
        .ColorScaleCriteria(1).Type = xlConditionValueLowestValue: .ColorScaleCriteria(1).FormatColor.Color = RGB(135, 206, 250)
        .ColorScaleCriteria(2).Type = xlConditionValuePercentile: .ColorScaleCriteria(2).Value = 50: .ColorScaleCriteria(2).FormatColor.Color = RGB(255, 255, 255)
        .ColorScaleCriteria(3).Type = xlConditionValueHighestValue: .ColorScaleCriteria(3).FormatColor.Color = RGB(255, 99, 71)
    End With
    MsgBox "K열(근무점수)에 히트맵이 적용되었습니다.", vbInformation
End Sub

Sub 데이터_무결성_검사()
    Dim wsData As Worksheet, r As Long, msg As String, errCnt As Integer
    Set wsData = ThisWorkbook.Sheets("인원관리")
    For r = 2 To wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
        If Application.CountIf(wsData.Columns(1), wsData.Cells(r, 1).Value) > 1 Then
            errCnt = errCnt + 1: msg = msg & wsData.Cells(r, 1).Value & " (중복)" & vbCrLf
        End If
    Next r
    If errCnt > 0 Then MsgBox "데이터 오류 발견!" & vbCrLf & msg, vbCritical Else MsgBox "데이터 상태 양호.", vbInformation
End Sub

Sub 동명이인_자동_구분()
    Dim wsData As Worksheet, r As Long, dict As Object, nm As String, newNm As String
    Set wsData = ThisWorkbook.Sheets("인원관리")
    Set dict = CreateObject("Scripting.Dictionary")
    For r = 2 To wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
        nm = Trim(wsData.Cells(r, 1).Value)
        If nm <> "" Then
            If dict.Exists(nm) Then
                dict(nm) = dict(nm) + 1
                newNm = nm & "(" & dict(nm) & ")"
                wsData.Cells(r, 1).Value = newNm
                wsData.Cells(r, 1).Interior.Color = RGB(255, 255, 153)
            Else
                dict.Add nm, 1
            End If
        End If
    Next r
    MsgBox "동명이인 구분 완료.", vbInformation
End Sub

Sub 시스템_초기세팅_및_메뉴얼생성()
    Dim wsMain As Worksheet, wsManual As Worksheet
    Dim btn As Shape
    
    Application.ScreenUpdating = False
    
    Call Init_Sheet("메인")
    Set wsMain = ThisWorkbook.Sheets("메인"): wsMain.Move Before:=ThisWorkbook.Sheets(1)
    wsMain.Cells.Clear: wsMain.Cells.Interior.Color = RGB(245, 245, 245)
    
    With wsMain
        .Columns("A").ColumnWidth = 2: .Columns("B").ColumnWidth = 25
        .Columns("C").ColumnWidth = 2: .Columns("D").ColumnWidth = 2
        .Columns("E").ColumnWidth = 25: .Rows("5:35").RowHeight = 25
    End With
    
    With wsMain.Range("B2:E3")
        .Merge: .Value = "부대 근무표 자동 관리 시스템" ' 버전 표기 삭제
        .Font.Size = 20: .Font.Bold = True: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Interior.Color = RGB(50, 50, 50): .Font.Color = vbWhite
    End With
    
    For Each btn In wsMain.Shapes: btn.Delete: Next btn
    
    Call Make_Button(wsMain, 5, 2, "근무표 새로 만들기", RGB(0, 112, 192), "근무표_이어쓰기_생성")
    Call Make_Button(wsMain, 5, 5, "실행 취소 (복구)", RGB(255, 192, 0), "Undo_Last_Action")
    Call Make_Button(wsMain, 8, 2, "달력으로 변환", RGB(255, 102, 0), "근무표_달력으로_보기")
    Call Make_Button(wsMain, 8, 5, "공정성 히트맵 (점수)", RGB(0, 176, 80), "공정성_히트맵_적용")
    Call Make_Button(wsMain, 11, 2, "근무 교체 (Ctrl+선택)", RGB(75, 0, 130), "근무_교체_마법사")
    Call Make_Button(wsMain, 11, 5, "인원 자동 정렬", RGB(0, 128, 128), "인원_자동_정렬")
    Call Make_Button(wsMain, 14, 2, "초기 데이터 자동 추산", RGB(255, 0, 102), "초기_데이터_자동_계산")
    Call Make_Button(wsMain, 14, 5, "동명이인 자동 구분", RGB(237, 125, 49), "동명이인_자동_구분")
    Call Make_Button(wsMain, 17, 2, "지난 근무 정리 (오늘까지)", RGB(192, 0, 0), "지난근무_기록이관_및_초기화")
    Call Make_Button(wsMain, 17, 5, "데이터 오류 검사", RGB(112, 48, 160), "데이터_무결성_검사")
    Call Make_Button(wsMain, 20, 2, "점수 가중치 설정", RGB(100, 100, 100), "점수_가중치_설정_생성")
    Call Make_Button(wsMain, 20, 5, "전역자 명부 이동", RGB(127, 127, 127), "전역자_자동_보내기")
    
    With wsMain.Range("B23")
        .Value = "※ [지난 근무 정리]를 누르면 오늘 날짜까지의 근무만 기록으로 넘기고 시트를 비웁니다."
        .Font.Color = RGB(100, 100, 100): .Font.Size = 9
    End With
    
    Call Init_Sheet("사용설명서")
    Set wsManual = ThisWorkbook.Sheets("사용설명서"): wsManual.Cells.Clear
    wsManual.Cells(2, 2).Value = "시스템 사용 설명서"
    wsManual.Cells(2, 2).Font.Size = 16: wsManual.Cells(2, 2).Font.Bold = True
    
    wsMain.Activate
    Application.ScreenUpdating = True
    MsgBox "메인 화면 업데이트 완료!", vbInformation
End Sub

Sub 점수_가중치_설정_생성()
    Dim wsSet As Worksheet
    Call Init_Sheet("설정")
    Set wsSet = ThisWorkbook.Sheets("설정")
    If wsSet.Range("J2").Value = "" Then
        wsSet.Range("I1:J1").Value = Array("구분", "점수"): wsSet.Range("I1:J1").Font.Bold = True
        wsSet.Cells(2, 9).Value = "평일_주간": wsSet.Cells(2, 10).Value = 1#
        wsSet.Cells(3, 9).Value = "평일_야간": wsSet.Cells(3, 10).Value = 1.5
        wsSet.Cells(4, 9).Value = "주말_주간": wsSet.Cells(4, 10).Value = 1.5
        wsSet.Cells(5, 9).Value = "주말_야간": wsSet.Cells(5, 10).Value = 2#
        wsSet.Columns("I:J").AutoFit
        MsgBox "설정 시트 I~J열을 확인하세요.", vbInformation
    Else
        MsgBox "이미 설정값이 존재합니다.", vbInformation
    End If
End Sub

' ==============================================================================
' [모듈 5] 보조 함수 (날짜, 이벤트 체크)
' ==============================================================================
Function Check_Is_Holiday(chkDate As Date, wsSet As Worksheet) As Boolean
    If Weekday(chkDate, vbMonday) > 5 Then Check_Is_Holiday = True: Exit Function
    If Application.CountIf(wsSet.Columns(4), chkDate) > 0 Then Check_Is_Holiday = True
End Function

Function Get_Event_Name(chkDate As Date, wsSet As Worksheet) As String
    Dim fRange As Range: Set fRange = wsSet.Columns(6).Find(chkDate, LookIn:=xlValues, LookAt:=xlWhole)
    If Not fRange Is Nothing Then Get_Event_Name = fRange.Offset(0, 1).Value
End Function

Sub Get_Event_Info(tDate As Date, wsSet As Worksheet, ByRef eName As String, ByRef eType As String)
    Dim fRange As Range
    eName = ""
    eType = ""
    Set fRange = wsSet.Columns(6).Find(tDate, LookIn:=xlValues, LookAt:=xlWhole)
    If Not fRange Is Nothing Then
        eName = fRange.Offset(0, 1).Value
        eType = fRange.Offset(0, 2).Value
        If eType = "" Then eType = "전체휴무"
    End If
End Sub

Function Check_Is_Excluded(workerName As String, targetDate As Date) As Boolean
    Dim wsEx As Worksheet: On Error Resume Next: Set wsEx = ThisWorkbook.Sheets("열외"): On Error GoTo 0
    If wsEx Is Nothing Then Exit Function
    If Application.CountIfs(wsEx.Columns(1), workerName, wsEx.Columns(2), "<=" & CDbl(targetDate), wsEx.Columns(3), ">=" & CDbl(targetDate)) > 0 Then Check_Is_Excluded = True
End Function

Sub Init_Sheet(sName As String)
    On Error Resume Next
    If ThisWorkbook.Sheets(sName) Is Nothing Then ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).Name = sName
    On Error GoTo 0
End Sub

Sub Make_Button(ws As Worksheet, r As Integer, c As Integer, txt As String, clr As Long, mac As String)
    Dim btn As Shape
    Dim rng As Range
    Set rng = ws.Cells(r, c).Resize(2, 1)
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, rng.Left + 2, rng.Top + 2, rng.Width - 4, rng.Height - 4)
    With btn
        .TextFrame.Characters.Text = txt
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 12
        .TextFrame.Characters.Font.Color = vbWhite
        .Fill.ForeColor.RGB = clr
        .Line.Visible = msoFalse
        .OnAction = mac
        .Shadow.Type = msoShadow25
    End With
End Sub

Sub 근무표_달력으로_보기()
    Dim wsRoster As Worksheet, wsCal As Worksheet
    Dim lastRow As Long, i As Long
    Dim targetDate As Date, startDate As Date, globalStartMonday As Date
    Dim weekCount As Long, colIdx As Long, currentRow As Long
    Dim shiftName As String, sName As String, hName As String, content As String
    
    Set wsRoster = ThisWorkbook.Sheets("근무표")
    On Error Resume Next
    Set wsCal = ThisWorkbook.Sheets("달력보기")
    If wsCal Is Nothing Then
        Set wsCal = ThisWorkbook.Sheets.Add(After:=wsRoster)
        wsCal.Name = "달력보기"
    End If
    On Error GoTo 0
    
    Application.ScreenUpdating = False
    wsCal.Cells.Clear: wsCal.Cells.Interior.Color = xlNone
    
    lastRow = wsRoster.Cells(wsRoster.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then MsgBox "표시할 데이터가 없습니다.", vbExclamation: Exit Sub
    
    If IsDate(wsRoster.Cells(2, 1).Value) Then startDate = CDate(wsRoster.Cells(2, 1).Value) Else MsgBox "날짜 오류", vbCritical: Exit Sub
    globalStartMonday = startDate - Weekday(startDate, vbMonday) + 1
    
    With wsCal
        .Columns("A:G").ColumnWidth = 16
        .Rows.VerticalAlignment = xlVAlignTop
        .Range("A1:G1").Value = Array("월", "화", "수", "목", "금", "토", "일")
        .Range("A1:G1").Interior.Color = RGB(220, 230, 241)
        .Range("A1:G1").Font.Bold = True
        .Range("A1:G1").HorizontalAlignment = xlCenter
        .Range("A1:G1").Borders.LineStyle = xlContinuous
    End With
    
    For i = 2 To lastRow
        If IsDate(wsRoster.Cells(i, 1).Value) Then
            targetDate = CDate(wsRoster.Cells(i, 1).Value)
            shiftName = wsRoster.Cells(i, 3).Value
            sName = wsRoster.Cells(i, 4).Value
            hName = wsRoster.Cells(i, 5).Value
            
            weekCount = Int((targetDate - globalStartMonday) / 7)
            colIdx = Weekday(targetDate, vbMonday)
            currentRow = 2 + (weekCount * 2)
            
            If wsCal.Cells(currentRow, colIdx).Value = "" Then
                wsCal.Cells(currentRow, colIdx).Value = Format(targetDate, "mm/dd")
                wsCal.Cells(currentRow, colIdx).HorizontalAlignment = xlCenter
                wsCal.Cells(currentRow, colIdx).Interior.Color = RGB(245, 245, 245)
                wsCal.Cells(currentRow, colIdx).Font.Size = 9
                If colIdx = 6 Then wsCal.Cells(currentRow, colIdx).Font.Color = vbBlue
                If colIdx = 7 Then wsCal.Cells(currentRow, colIdx).Font.Color = vbRed
                wsCal.Range(wsCal.Cells(currentRow, 1), wsCal.Cells(currentRow + 1, 7)).Borders.LineStyle = xlContinuous
            End If
            
            content = "[" & shiftName & "] " & sName
            If hName <> "-" And hName <> "" And hName <> "인원부족" And hName <> "전원부재" Then content = content & ", " & hName
            
            With wsCal.Cells(currentRow + 1, colIdx)
                If .Value = "" Then .Value = content Else .Value = .Value & vbLf & content
                .WrapText = True: .Font.Size = 10
            End With
        End If
    Next i
    wsCal.Rows.AutoFit
    wsCal.Activate
    Application.ScreenUpdating = True
    MsgBox "달력 변환 완료!", vbInformation
End Sub

Sub 초기_데이터_자동_계산()
    Dim wsData As Worksheet
    Dim lastRow As Long, r As Long
    Dim serviceDays As Long, startDate As Date
    Dim interval As Double, estTotal As Long, estWeekend As Long
    Dim response As String
    
    Set wsData = ThisWorkbook.Sheets("인원관리")
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then MsgBox "인원 데이터가 없습니다.", vbExclamation: Exit Sub
    
    response = InputBox("이 부대는 보통 며칠에 한 번 근무를 섭니까? (예: 5)", "초기 데이터 추산", "5")
    If Not IsNumeric(response) Or response = "" Then Exit Sub
    interval = CDbl(response): If interval <= 0 Then Exit Sub
    
    If MsgBox("기존 F, G열 값이 덮어씌워집니다.", vbYesNo + vbExclamation) = vbNo Then Exit Sub
    
    Application.ScreenUpdating = False
    For r = 2 To lastRow
        If IsDate(wsData.Cells(r, 9).Value) Then
            startDate = CDate(wsData.Cells(r, 9).Value)
        ElseIf IsDate(wsData.Cells(r, 3).Value) Then
            startDate = CDate(wsData.Cells(r, 3).Value)
        Else
            startDate = 0
        End If
        If startDate > 0 And startDate <= Date Then
            serviceDays = Date - startDate
            estTotal = Int(serviceDays / interval)
            estWeekend = Int(estTotal * 0.285)
            wsData.Cells(r, 6).Value = estTotal
            wsData.Cells(r, 7).Value = estWeekend
        Else
            wsData.Cells(r, 6).Value = 0: wsData.Cells(r, 7).Value = 0
        End If
    Next r
    Call 통계_강제_갱신
    Application.ScreenUpdating = True
    MsgBox "계산 완료!", vbInformation
End Sub

Sub 지난근무_기록이관_및_초기화()
    Dim wsData As Worksheet, wsRoster As Worksheet, wsSetting As Worksheet
    Dim lastRowRoster As Long, lastRowData As Long
    Dim i As Long, r As Long, delCount As Long
    Dim targetDate As Date, workerName As String, shiftName As String
    Dim score As Double, isHol As Boolean
    
    Dim dictShiftScore As Object: Set dictShiftScore = CreateObject("Scripting.Dictionary")
    Dim dictPastScore As Object: Set dictPastScore = CreateObject("Scripting.Dictionary")
    
    Set wsData = ThisWorkbook.Sheets("인원관리")
    Set wsRoster = ThisWorkbook.Sheets("근무표")
    Set wsSetting = ThisWorkbook.Sheets("설정")
    
    If MsgBox("오늘(" & Format(Date, "mm-dd") & ")까지의 근무 기록을 정리하시겠습니까?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    
    Application.ScreenUpdating = False
    Dim setLast As Long: setLast = wsSetting.Cells(wsSetting.Rows.Count, "I").End(xlUp).Row
    If setLast >= 2 Then
        For i = 2 To setLast
            shiftName = wsSetting.Cells(i, 9).Value
            dictShiftScore(shiftName & "_평일") = Val(wsSetting.Cells(i, 11).Value)
            dictShiftScore(shiftName & "_휴일") = Val(wsSetting.Cells(i, 12).Value)
        Next i
    End If
    
    lastRowRoster = wsRoster.Cells(wsRoster.Rows.Count, "A").End(xlUp).Row
    delCount = 0
    For i = 2 To lastRowRoster
        If IsDate(wsRoster.Cells(i, 1).Value) Then
            targetDate = CDate(wsRoster.Cells(i, 1).Value)
            If targetDate <= Date Then
                isHol = Check_Is_Holiday(targetDate, wsSetting)
                shiftName = wsRoster.Cells(i, 3).Value
                Dim key As String
                If isHol Then key = shiftName & "_휴일" Else key = shiftName & "_평일"
                If dictShiftScore.Exists(key) Then score = dictShiftScore(key) Else score = 1
                
                Dim cols As Variant, c As Variant: cols = Array(4, 5)
                For Each c In cols
                    workerName = wsRoster.Cells(i, c).Value
                    If workerName <> "" And workerName <> "전원부재" And workerName <> "인원부족" And workerName <> "-" Then
                        dictPastScore(workerName) = dictPastScore(workerName) + score
                    End If
                Next c
            End If
        End If
    Next i
    
    lastRowData = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    For r = 2 To lastRowData
        workerName = wsData.Cells(r, 1).Value
        If dictPastScore.Exists(workerName) Then
            wsData.Cells(r, 6).Value = Val(wsData.Cells(r, 6).Value) + dictPastScore(workerName)
        End If
    Next r
    
    For i = lastRowRoster To 2 Step -1
        If IsDate(wsRoster.Cells(i, 1).Value) Then
            targetDate = CDate(wsRoster.Cells(i, 1).Value)
            If targetDate <= Date Then wsRoster.Rows(i).Delete: delCount = delCount + 1
        End If
    Next i
    Call 통계_강제_갱신
    Application.ScreenUpdating = True
    MsgBox delCount & "건 정리 완료.", vbInformation
End Sub

