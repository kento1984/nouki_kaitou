Attribute VB_Name = "Module1"
Option Explicit
' 実行開始時刻を保持
Public g_ExecutionTime As Date
Public g_CurrentSubFolder As String  '
' 営業所設定（起動時にセット）
Public g_BranchName As String
Public g_DefaultCutoff As Integer
Public g_BaseCenter As String
Public g_SharedEmail As String
Public g_Signature As String
Public g_StartDate As String
' ============================================
' パフォーマンス改善：マスターデータキャッシュ
' ============================================
Public g_MfgNameCache As Object       ' itemGroupCode -> メーカー名
Public g_MfgDaysCache As Object       ' itemGroupCode -> 配送日数
Public g_CustDaysCache As Object      ' customerName -> 配送曜日(文字列)
Public g_CustRetentionCache As Object ' customerName -> 保持日数
Public g_CustRouteCache As Object     ' customerName -> 路線便フラグ
Public g_ConfirmCache As Object       ' orderNum|detailNum -> Array(col8, col9, col10)
Public g_StorageCache As Object       ' orderNumber -> 保管場所
Public g_SourceData As Variant       ' sourceWsデータの配列キャッシュ
' ============================================
' 【v6.0】特別日カレンダーから祝日・特別締切時間を読み込む
' ============================================
Function Loadholidays(manufacturerMasterWb As Workbook) As Object
    Dim holidays As Object
    Set holidays = CreateObject("Scripting.Dictionary")
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim specialDate As Date
    Dim cutoffValue As Variant
    
    On Error Resume Next
    Set ws = manufacturerMasterWb.Sheets("特別日カレンダー")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set Loadholidays = holidays
        Exit Function
    End If
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        On Error Resume Next
        specialDate = CDate(ws.Cells(i, 1).Value)
        If Err.Number = 0 And specialDate > 0 Then
            cutoffValue = ws.Cells(i, 3).Value
            
            If IsEmpty(cutoffValue) Or Trim(CStr(cutoffValue)) = "" Then
                holidays(CLng(specialDate)) = ""
            ElseIf IsNumeric(cutoffValue) Then
                holidays(CLng(specialDate)) = CStr(CLng(cutoffValue))
            Else
                holidays(CLng(specialDate)) = ""
            End If
        End If
        Err.Clear
        On Error GoTo 0
    Next i
    
    Set Loadholidays = holidays
End Function

' ============================================
' 【v6.0】営業所設定を取得（特別締切時間対応）
' ============================================
Function GetBranchSettings(Optional holidays As Object = Nothing, _
                           Optional targetDate As Date = 0) As Variant
    Dim branchName As String
    Dim cutoffHour As Integer
    Dim signature As String
    Dim defaultCutoff As Integer
    
    If targetDate = 0 Then
        targetDate = Date
    End If
    
    branchName = g_BranchName
    defaultCutoff = g_DefaultCutoff
    
    cutoffHour = defaultCutoff
    
    If Not holidays Is Nothing Then
        If holidays.Exists(CLng(targetDate)) Then
            Dim specialValue As String
            specialValue = holidays(CLng(targetDate))
            
            If specialValue <> "" And IsNumeric(specialValue) Then
                cutoffHour = CInt(specialValue)
            End If
        End If
    End If
    
    signature = g_Signature
    
    GetBranchSettings = Array(branchName, cutoffHour, signature)
End Function
' ============================================
' 営業所設定を営業所設定シートから読み込み、モジュールレベル変数にセット
' ============================================
Sub LoadBranchSettings(manufacturerMasterWb As Workbook, sourceWs As Worksheet, cols As Object)
    Dim branchSettingsWs As Worksheet
    Dim branchCode As String
    Dim lastRow As Long
    Dim i As Long
    Dim firstOrderNum As String
    Dim dataLastRow As Long

    ' グローバル変数をリセット
    g_BranchName = ""
    g_DefaultCutoff = 0
    g_BaseCenter = ""
    g_SharedEmail = ""
    g_Signature = ""
    g_StartDate = ""

    ' 受注一覧の最初の注番から先頭2文字を取得
    dataLastRow = sourceWs.Cells(sourceWs.Rows.count, cols("受発注伝票")).End(xlUp).Row
    For i = 7 To dataLastRow
        firstOrderNum = Trim(g_SourceData(i, cols("受発注伝票")))
        If firstOrderNum <> "" And Len(firstOrderNum) >= 2 Then
            Dim firstTwo As String
            firstTwo = Left(firstOrderNum, 2)
            If firstTwo Like "[A-Z][A-Z]" Then
                branchCode = firstTwo
                Exit For
            End If
        End If
    Next i

    If branchCode = "" Then
        MsgBox "受注一覧に注番データが見つかりません。", vbExclamation
        Exit Sub
    End If

    ' 営業所設定シートから検索
    On Error Resume Next
    Set branchSettingsWs = manufacturerMasterWb.Sheets("営業所設定")
    On Error GoTo 0

    If branchSettingsWs Is Nothing Then
        MsgBox "メーカー一覧.xlsxに「営業所設定」シートが見つかりません。", vbExclamation
        Exit Sub
    End If

    lastRow = branchSettingsWs.Cells(branchSettingsWs.Rows.count, 1).End(xlUp).Row

    For i = 2 To lastRow
        If Trim(branchSettingsWs.Cells(i, 1).Value) = branchCode Then
            g_BranchName = Trim(branchSettingsWs.Cells(i, 2).Value)
            g_DefaultCutoff = CInt(branchSettingsWs.Cells(i, 3).Value)
            g_BaseCenter = Trim(branchSettingsWs.Cells(i, 4).Value)
            g_SharedEmail = Trim(branchSettingsWs.Cells(i, 5).Value)
            g_Signature = "マツモト産業" & vbLf & g_BranchName
            g_StartDate = Format(branchSettingsWs.Cells(i, 6).Value, "yyyy/mm/dd")
            Exit Sub
        End If
    Next i

    MsgBox "注番コード「" & branchCode & "」に対応する営業所設定が見つかりません。" & vbCrLf & _
           "メーカー一覧.xlsxの「営業所設定」シートを確認してください。", vbExclamation
End Sub

' ============================================
' 納期回答書作成マクロ v4.8（除外機能＋期間フィルタ修正版）
' 【v4.8新機能】確認中一覧の「除外」ステータスで納期回答書から除外
' 【v4.8修正】テーブル行追加時に入力規則を正しく設定
' 【v4.8修正】フォームで選択した期間が正しく反映されるように修正
' 【v4.7改善】フォント・罫線・色をモダンに刷新
' 【v4.6修正】当日送付分は除外しない（昼・夕方両方で表示）
' 【v4.6修正】送付履歴の重複チェック追加
' 【v4.5新機能】メーカー一覧.xlsxの「祝日」シートを参照して祝日を除外
' ============================================

Sub 納期回答書作成()
    Dim sourceFilePath As String
    Dim customerMasterPath As String
    Dim manufacturerMasterPath As String
    Dim deliveryHistoryPath As String
    Dim sourceWb As Workbook
    Dim customerMasterWb As Workbook
    Dim manufacturerMasterWb As Workbook
    Dim deliveryHistoryWb As Workbook
    Dim sourceWs As Worksheet
    Dim customerMasterWs As Worksheet
    Dim manufacturerMasterWs As Worksheet
    Dim deliveryHistoryWs As Worksheet
    Dim confirmingListWs As Worksheet
    Dim tempFilePath As String
    Dim isTemporaryFile As Boolean
    Dim mainTimer As Double
    mainTimer = Timer
    Application.ScreenUpdating = False
    Dim origCalcMode As Long
    origCalcMode = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    isTemporaryFile = False
    
    ' 実行開始時刻を記録（フォルダ名に使用）
    g_ExecutionTime = Now
    g_CurrentSubFolder = ""
    
    ' 1. 受注一覧ファイルを選択（.xlsと.xlsx両方OK）
    sourceFilePath = Application.GetOpenFilename("Excel Files (*.xls; *.xlsx), *.xls; *.xlsx", , "受注一覧ファイルを選択してください")
    If sourceFilePath = "False" Then
        MsgBox "キャンセルされました。", vbInformation
        GoTo Cleanup
    End If
    
    ' .xlsファイルの場合は.xlsxに変換
    If LCase(Right(sourceFilePath, 4)) = ".xls" Then
        tempFilePath = ConvertXlsToXlsx(sourceFilePath)
        If tempFilePath = "" Then
            MsgBox "ファイルの変換に失敗しました。", vbExclamation
            GoTo Cleanup
        End If
        sourceFilePath = tempFilePath
        isTemporaryFile = True
    End If
    
    ' 2. 各種マスターファイルを開く
    Dim toolPath As String
    toolPath = ThisWorkbook.Path
    
    ' 顧客マスターファイルを検索（xlsm優先）
    customerMasterPath = toolPath & "\顧客マスター_v2.xlsm"
    If Dir(customerMasterPath) = "" Then
        customerMasterPath = toolPath & "\顧客マスター_v2.xlsx"
    End If
    
    manufacturerMasterPath = toolPath & "\メーカー一覧.xlsx"
    deliveryHistoryPath = toolPath & "\送付履歴.xlsx"
    
    ' 顧客マスターが見つからない場合
    If Dir(customerMasterPath) = "" Then
        MsgBox "顧客マスター_v2が見つかりません。" & vbCrLf & _
               "（.xlsxまたは.xlsmファイル）" & vbCrLf & vbCrLf & _
               "ツールと同じフォルダに配置してください。" & vbCrLf & vbCrLf & _
               "場所：" & toolPath, vbExclamation
        If isTemporaryFile Then Kill tempFilePath
        GoTo Cleanup
    End If
    
    ' メーカー一覧が見つからない場合
    If Dir(manufacturerMasterPath) = "" Then
        MsgBox "メーカー一覧.xlsxが見つかりません。" & vbCrLf & _
               "ツールと同じフォルダに配置してください。" & vbCrLf & vbCrLf & _
               "場所：" & toolPath, vbExclamation
        If isTemporaryFile Then Kill tempFilePath
        GoTo Cleanup
    End If
    
    ' 送付履歴ファイルが存在しない場合は新規作成
    If Dir(deliveryHistoryPath) = "" Then
        Call InitializeDeliveryHistory(deliveryHistoryPath)
    End If
    
    ' 送付履歴が使用中かチェック
    If IsFileOpen(deliveryHistoryPath) Then
        MsgBox "━━━━━━━━━━━━━━━━━━" & vbCrLf & _
       "　▲ 送付履歴ファイル 使用中 ▲" & vbCrLf & _
       "━━━━━━━━━━━━━━━━━━" & vbCrLf & vbCrLf & _
       "他の人が作業中のため実行できません。" & vbCrLf & vbCrLf & _
       "【対処法】" & vbCrLf & _
       "　1. 使用中の人に声をかける" & vbCrLf & _
       "　2. 閉じてもらってから再実行", vbCritical, "実行できません"
        If isTemporaryFile Then Kill tempFilePath
        GoTo Cleanup
    End If
    
    ' ファイルを開く
    Set sourceWb = Workbooks.Open(sourceFilePath)
    Set sourceWs = sourceWb.Sheets(1)
    Set customerMasterWb = Workbooks.Open(customerMasterPath)
    Set customerMasterWs = customerMasterWb.Sheets("顧客マスター")
    
    ' 担当者マスターシートを取得（存在しなければNothing）
    Dim repMasterWs As Worksheet
    Set repMasterWs = Nothing
    On Error Resume Next
    Set repMasterWs = customerMasterWb.Sheets("担当者マスター")
    On Error GoTo 0
    
    Set manufacturerMasterWb = Workbooks.Open(manufacturerMasterPath)
    Set manufacturerMasterWs = manufacturerMasterWb.Sheets(1)
    Set deliveryHistoryWb = Workbooks.Open(deliveryHistoryPath)
    Set deliveryHistoryWs = deliveryHistoryWb.Sheets("送付履歴")
    Set confirmingListWs = deliveryHistoryWb.Sheets("確認中一覧")
    
    ' 【v4.5新機能】祝日を読み込み
    Dim holidays As Object
    Set holidays = Loadholidays(manufacturerMasterWb)
    
    ' 列位置を検索
    Dim cols As Object
    Set cols = GetColumnPositions(sourceWs)
    
    If cols Is Nothing Then
        MsgBox "必要な列が見つかりません。" & vbCrLf & _
               "ヘッダー行（5行目）を確認してください。", vbExclamation
        sourceWb.Close SaveChanges:=False
        customerMasterWb.Close SaveChanges:=False
        manufacturerMasterWb.Close SaveChanges:=False
        deliveryHistoryWb.Close SaveChanges:=False
        If isTemporaryFile Then Kill tempFilePath
        GoTo Cleanup
    End If
    
    ' 品目Group列のチェック
    If Not cols.Exists("品目Group") Then
        MsgBox "「品目 Group」列が見つかりません。" & vbCrLf & _
               "受注一覧ファイルを確認してください。", vbExclamation
        sourceWb.Close SaveChanges:=False
        customerMasterWb.Close SaveChanges:=False
        manufacturerMasterWb.Close SaveChanges:=False
        deliveryHistoryWb.Close SaveChanges:=False
        If isTemporaryFile Then Kill tempFilePath
        GoTo Cleanup
    End If
    
    ' 登録日列のチェック
    If Not cols.Exists("登録日") Then
        MsgBox "「登録日」列が見つかりません。" & vbCrLf & _
               "受注一覧ファイルを確認してください。", vbExclamation
        sourceWb.Close SaveChanges:=False
        customerMasterWb.Close SaveChanges:=False
        manufacturerMasterWb.Close SaveChanges:=False
        deliveryHistoryWb.Close SaveChanges:=False
        If isTemporaryFile Then Kill tempFilePath
        GoTo Cleanup
    End If
    
    ' 明細列のチェック
    If Not cols.Exists("明細") Then
        MsgBox "「明細」列が見つかりません。" & vbCrLf & _
               "受注一覧ファイルを確認してください。", vbExclamation
        sourceWb.Close SaveChanges:=False
        customerMasterWb.Close SaveChanges:=False
        manufacturerMasterWb.Close SaveChanges:=False
        deliveryHistoryWb.Close SaveChanges:=False
        If isTemporaryFile Then Kill tempFilePath
        GoTo Cleanup
    End If
    
    ' 営業所設定を読み込み
    ' ============================================
    ' sourceWsデータを配列に一括読み込み
    ' ============================================
    Dim srcLastRow As Long, srcLastCol As Long
    srcLastRow = sourceWs.Cells(sourceWs.Rows.Count, cols("受発注伝票")).End(xlUp).Row
    srcLastCol = 0
    Dim colKey As Variant
    For Each colKey In cols.Keys
        If cols(colKey) > srcLastCol Then srcLastCol = cols(colKey)
    Next colKey
    g_SourceData = sourceWs.Range(sourceWs.Cells(1, 1), sourceWs.Cells(srcLastRow, srcLastCol)).Value

    Call LoadBranchSettings(manufacturerMasterWb, sourceWs, cols)
    If g_BranchName = "" Then
        sourceWb.Close SaveChanges:=False
        customerMasterWb.Close SaveChanges:=False
        manufacturerMasterWb.Close SaveChanges:=False
        deliveryHistoryWb.Close SaveChanges:=False
        If isTemporaryFile Then Kill tempFilePath
        GoTo Cleanup
    End If
    
    ' 【v4.8修正】送付履歴を読み込み（確認中一覧の「除外」も含む）
    Dim sentOrders As Object
    ' ============================================
    ' マスターデータキャッシュ構築
    ' ============================================
    Call BuildManufacturerCache(manufacturerMasterWs)
    Call BuildCustomerCache(customerMasterWs)
    Call BuildConfirmingCache(confirmingListWs)
    Call BuildStorageCache(sourceWs, cols)
    
    Set sentOrders = LoadDeliveryHistory(deliveryHistoryWs, confirmingListWs, customerMasterWs, holidays)
    
    ' ============================================
    ' フォーム表示（注番/期間モード選択）
    ' ============================================
    Dim frmSelection As frmSelection
    Set frmSelection = New frmSelection
    
    frmSelection.SetDataSource sourceWs, cols, customerMasterWs
    
    frmSelection.Show
    
    ' キャンセルされた場合
    If frmSelection.IsCancelled Then
        sourceWb.Close SaveChanges:=False
        customerMasterWb.Close SaveChanges:=False
        manufacturerMasterWb.Close SaveChanges:=False
        deliveryHistoryWb.Close SaveChanges:=False
        If isTemporaryFile Then Kill tempFilePath
        Unload frmSelection
        GoTo Cleanup
    End If
    
    ' ============================================
    ' モード判定：注番指定 or 期間指定
    ' ============================================
    Dim IsOrderNumberMode As Boolean
    IsOrderNumberMode = frmSelection.IsOrderNumberMode
    
    ' 【v4.8追加】期間を取得
    Dim dateFrom As Date
    Dim dateTo As Date
    dateFrom = frmSelection.GetDateFrom
    dateTo = frmSelection.GetDateTo
    
    Dim createdFiles As Collection
    Set createdFiles = New Collection
    
    Dim newSentOrders As Collection
    Set newSentOrders = New Collection
    
    Dim newConfirmingOrders As Collection
    Set newConfirmingOrders = New Collection
    
    If IsOrderNumberMode Then
        ' ============================================
        ' 注番指定モード（顧客ごとにまとめる）
        ' ============================================
        Dim orderNumbers As Collection
        Set orderNumbers = frmSelection.GetOrderNumbers
        Unload frmSelection
        
        If orderNumbers.count = 0 Then
            MsgBox "注番が入力されませんでした。", vbInformation
            sourceWb.Close SaveChanges:=False
            customerMasterWb.Close SaveChanges:=False
            manufacturerMasterWb.Close SaveChanges:=False
            deliveryHistoryWb.Close SaveChanges:=False
            If isTemporaryFile Then Kill tempFilePath
            GoTo Cleanup
        End If
        
        ' 注番を顧客ごとにグループ化
        Dim customerGroups As Object
        Set customerGroups = GroupOrderNumbersByCustomer(sourceWs, cols, orderNumbers)
        
        ' 顧客ごとに納期回答書を作成
        Dim customerName As Variant
        For Each customerName In customerGroups.keys
            Dim orderNumList As Collection
            Set orderNumList = customerGroups(customerName)
            
            Dim resultData As Variant
            
            If Not repMasterWs Is Nothing And IsSplitByRep(CStr(customerName), repMasterWs) Then
                ' 担当者別分割
                Dim repListON As Collection
                Set repListON = GetRepList(CStr(customerName), repMasterWs)
                Dim repItemON As Variant
                For Each repItemON In repListON
                    resultData = CreateDeliveryReportByOrderNumbers(sourceWs, cols, CStr(customerName), _
                                       orderNumList, manufacturerMasterWs, holidays, confirmingListWs, _
                                       customerMasterWs, CStr(repItemON), repMasterWs)
                    If Not IsEmpty(resultData) Then
                        createdFiles.Add Array(resultData(0), resultData(1), resultData(2), resultData(3), resultData(4), CStr(repItemON))
                    End If
                Next repItemON
                ' その他バケット
                resultData = CreateDeliveryReportByOrderNumbers(sourceWs, cols, CStr(customerName), _
                                   orderNumList, manufacturerMasterWs, holidays, confirmingListWs, _
                                   customerMasterWs, "__OTHER__", repMasterWs)
                If Not IsEmpty(resultData) Then
                    createdFiles.Add Array(resultData(0), resultData(1), resultData(2), resultData(3), resultData(4), "")
                End If
            Else
                resultData = CreateDeliveryReportByOrderNumbers(sourceWs, cols, CStr(customerName), _
                                   orderNumList, manufacturerMasterWs, holidays, confirmingListWs, _
                                   customerMasterWs)
                If Not IsEmpty(resultData) Then
                    createdFiles.Add Array(resultData(0), resultData(1), resultData(2), resultData(3), resultData(4), "")
                End If
            End If
        Next customerName
        
        ' 元のファイルを閉じる
        sourceWb.Close SaveChanges:=False
        
        ' 一時ファイルを削除
        If isTemporaryFile Then
            On Error Resume Next
            Kill tempFilePath
            On Error GoTo 0
        End If
        
        Set g_MfgNameCache = Nothing
        Set g_MfgDaysCache = Nothing
        Set g_CustDaysCache = Nothing
        Set g_CustRetentionCache = Nothing
        Set g_CustRouteCache = Nothing
        Set g_ConfirmCache = Nothing
        Set g_StorageCache = Nothing
        g_SourceData = Empty
        Application.Calculation = origCalcMode
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        
        ' 完了メッセージ（送付履歴には記録しない）
        MsgBox createdFiles.count & "件の納期回答書を作成しました。" & vbCrLf & _
               "※注番指定モードのため、送付履歴には記録されていません。", vbInformation
        
        ' メール送信の確認（3択）
Dim mailChoice As VbMsgBoxResult
mailChoice = MsgBox("メールを送信しますか？" & vbCrLf & vbCrLf & _
                    "【はい】→ そのまま送信（" & createdFiles.count & "件）" & vbCrLf & _
                    "【いいえ】→ 確認してから送信" & vbCrLf & _
                    "【キャンセル】→ 送信しない", _
                    vbYesNoCancel + vbQuestion, "メール送信")

If mailChoice = vbYes Then
    Call CreateEmails(createdFiles, customerMasterWs, manufacturerMasterWs, holidays, confirmingListWs, True, repMasterWs)
ElseIf mailChoice = vbNo Then
    Call CreateEmails(createdFiles, customerMasterWs, manufacturerMasterWs, holidays, confirmingListWs, False, repMasterWs)
End If
        
        manufacturerMasterWb.Close SaveChanges:=False
        deliveryHistoryWb.Close SaveChanges:=False
        customerMasterWb.Close SaveChanges:=False
        
    Else
        ' ============================================
        ' 期間指定モード
        ' ============================================
        Dim customerList As Collection
        Set customerList = frmSelection.GetSelectedCustomers
        Unload frmSelection
        
        ' 選択された顧客が0件の場合
        If customerList Is Nothing Or customerList.count = 0 Then
            MsgBox "顧客が選択されませんでした。", vbInformation
            sourceWb.Close SaveChanges:=False
            customerMasterWb.Close SaveChanges:=False
            manufacturerMasterWb.Close SaveChanges:=False
            deliveryHistoryWb.Close SaveChanges:=False
            If isTemporaryFile Then Kill tempFilePath
            GoTo Cleanup
        End If
        
        ' 顧客マスターをチェック
        Dim missingEmails As String
        missingEmails = CheckCustomerMaster(customerList, customerMasterWs)
        
        If missingEmails <> "" Then
            Dim result As VbMsgBoxResult
            result = MsgBox("以下の顧客はメールアドレスが未登録です：" & vbCrLf & vbCrLf & _
                           missingEmails & vbCrLf & _
                           "このまま続けますか？" & vbCrLf & vbCrLf & _
                           "※未登録の顧客はメール作成をスキップします。", _
                           vbYesNo + vbQuestion, "メールアドレス未登録")
            If result = vbNo Then
                sourceWb.Close SaveChanges:=False
                customerMasterWb.Close SaveChanges:=False
                manufacturerMasterWb.Close SaveChanges:=False
                deliveryHistoryWb.Close SaveChanges:=False
                If isTemporaryFile Then Kill tempFilePath
                GoTo Cleanup
            End If
        End If
        
        ' 各顧客の納期回答書を作成（【v4.8修正】期間を渡す）
        Dim customerKey As Variant
        Dim customerCount As Long
        Dim currentCustomer As Long
        Dim customerName2 As String
        customerCount = customerList.count
        currentCustomer = 0
        
        MsgBox customerCount & "件の顧客を処理します。" & vbCrLf & _
               "完了までしばらくお待ちください。", vbInformation, "処理開始"
        
    ' ★Step7: 送付履歴を一時クローズ（メモリ負荷軽減）
    deliveryHistoryWb.Close SaveChanges:=False
    Set deliveryHistoryWs = Nothing
    Set confirmingListWs = Nothing
    

        For Each customerKey In customerList
            currentCustomer = currentCustomer + 1
            customerName2 = Split(CStr(customerKey), "|")(1)
            
            ' 進捗表示
            Application.ScreenUpdating = True
            Application.StatusBar = "処理中... (" & currentCustomer & "/" & customerCount & ") " & customerName2 & "様"
            DoEvents
            Application.ScreenUpdating = False
            
            Dim resultData2 As Variant
            
            If Not repMasterWs Is Nothing And IsSplitByRep(customerName2, repMasterWs) Then
                ' 担当者別分割
                Dim repListPM As Collection
                Set repListPM = GetRepList(customerName2, repMasterWs)
                Dim repItemPM As Variant
                For Each repItemPM In repListPM
                    resultData2 = CreateDeliveryReport(sourceWs, cols, customerName2, _
                          manufacturerMasterWs, sentOrders, holidays, _
                          dateFrom, dateTo, confirmingListWs, customerMasterWs, _
                          CStr(repItemPM), repMasterWs)
                    
                    If Not IsEmpty(resultData2) Then
                        createdFiles.Add Array(customerName2, resultData2(0), resultData2(3), resultData2(4), resultData2(5), CStr(repItemPM), resultData2(6))
                        Dim orderItemR As Variant
                        If Not resultData2(1) Is Nothing Then
                            For Each orderItemR In resultData2(1)
                                newSentOrders.Add orderItemR
                            Next orderItemR
                        End If
                        Dim confirmItemR As Variant
                        If Not resultData2(2) Is Nothing Then
                            For Each confirmItemR In resultData2(2)
                                newConfirmingOrders.Add confirmItemR
                            Next confirmItemR
                        End If
                    End If
                Next repItemPM
                ' その他バケット
                resultData2 = CreateDeliveryReport(sourceWs, cols, customerName2, _
                      manufacturerMasterWs, sentOrders, holidays, _
                      dateFrom, dateTo, confirmingListWs, customerMasterWs, _
                      "__OTHER__", repMasterWs)
                
                If Not IsEmpty(resultData2) Then
                    createdFiles.Add Array(customerName2, resultData2(0), resultData2(3), resultData2(4), resultData2(5), "", resultData2(6))
                    If Not resultData2(1) Is Nothing Then
                        Dim orderItemO As Variant
                        For Each orderItemO In resultData2(1)
                            newSentOrders.Add orderItemO
                        Next orderItemO
                    End If
                    If Not resultData2(2) Is Nothing Then
                        Dim confirmItemO As Variant
                        For Each confirmItemO In resultData2(2)
                            newConfirmingOrders.Add confirmItemO
                        Next confirmItemO
                    End If
                End If
            Else
                ' 従来処理
                resultData2 = CreateDeliveryReport(sourceWs, cols, customerName2, _
                      manufacturerMasterWs, sentOrders, holidays, _
                      dateFrom, dateTo, confirmingListWs, customerMasterWs)
                
                If Not IsEmpty(resultData2) Then
                    createdFiles.Add Array(customerName2, resultData2(0), resultData2(3), resultData2(4), resultData2(5), "", resultData2(6))
                    If Not resultData2(1) Is Nothing Then
                        Dim orderItem As Variant
                        For Each orderItem In resultData2(1)
                            newSentOrders.Add orderItem
                        Next orderItem
                    End If
                    If Not resultData2(2) Is Nothing Then
                        Dim confirmItem As Variant
                        For Each confirmItem In resultData2(2)
                            newConfirmingOrders.Add confirmItem
                        Next confirmItem
                    End If
                End If
            End If
        Next
    
    ' ★Step7: 送付履歴を再オープン
    Set deliveryHistoryWb = Workbooks.Open(deliveryHistoryPath)
    Set deliveryHistoryWs = deliveryHistoryWb.Sheets("送付履歴")
    Set confirmingListWs = deliveryHistoryWb.Sheets("確認中一覧")
        
        ' 元のファイルを閉じる
        sourceWb.Close SaveChanges:=False
        
        
        ' 一時ファイルを削除
        If isTemporaryFile Then
            On Error Resume Next
            Kill tempFilePath
            On Error GoTo 0
        End If
        
        ' 【v4.4新機能】確認中一覧のクリーンアップ（確定になった伝票を移動）
        On Error Resume Next
        Call CleanConfirmingList(deliveryHistoryWs, confirmingListWs, newSentOrders)
        If Err.Number <> 0 Then
            MsgBox "エラー番号: " & Err.Number & vbCrLf & "エラー内容: " & Err.Description & vbCrLf & "確認中一覧のクリーンアップ中にエラーが発生しましたが、処理を継続します。", vbExclamation
            Err.Clear
        End If
        On Error GoTo 0
        
        ' 送付履歴を更新（メール作成前に実行）
        If newSentOrders.count > 0 Then
            On Error Resume Next
            Call SaveDeliveryHistory(deliveryHistoryWs, newSentOrders)
            Call CleanOldHistory(deliveryHistoryWs, 180)
            Call CleanOldConfirmingList(confirmingListWs, 180)
            If Err.Number <> 0 Then
                MsgBox "エラー番号: " & Err.Number & vbCrLf & "エラー内容: " & Err.Description & vbCrLf & "送付履歴の保存中にエラーが発生しましたが、処理を継続します。", vbExclamation
                Err.Clear
            End If
            On Error GoTo 0
        End If
        
        ' 【v4.4新機能】確認中一覧を更新
        If newConfirmingOrders.count > 0 Then
            On Error Resume Next
            Call SaveConfirmingList(confirmingListWs, newConfirmingOrders)
            If Err.Number <> 0 Then
                MsgBox "エラー番号: " & Err.Number & vbCrLf & "エラー内容: " & Err.Description & vbCrLf & "確認中一覧の保存中にエラーが発生しましたが、処理を継続します。", vbExclamation
                Err.Clear
            End If
            On Error GoTo 0
        End If
        
        ' ステータスバーをリセット
        Application.StatusBar = False
        
        ' 保存
        deliveryHistoryWb.Save
        
        
        
        Set g_MfgNameCache = Nothing
        Set g_MfgDaysCache = Nothing
        Set g_CustDaysCache = Nothing
        Set g_CustRetentionCache = Nothing
        Set g_CustRouteCache = Nothing
        Set g_ConfirmCache = Nothing
        Set g_StorageCache = Nothing
        g_SourceData = Empty
        Application.Calculation = origCalcMode
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        
        ' 完了メッセージ
        Dim completionMsg As String
        completionMsg = createdFiles.count & "件の納期回答書を作成しました。"
        If newSentOrders.count > 0 Then
            completionMsg = completionMsg & vbCrLf & "送付履歴に" & newSentOrders.count & "件の確定伝票を記録しました。"
        End If
        If newConfirmingOrders.count > 0 Then
            completionMsg = completionMsg & vbCrLf & "確認中一覧に" & newConfirmingOrders.count & "件の確認中伝票を記録しました。"
        End If
        If newSentOrders.count = 0 And newConfirmingOrders.count = 0 Then
            completionMsg = completionMsg & vbCrLf & "※記録する伝票がありませんでした。"
        End If
        
        MsgBox completionMsg, vbInformation
        
        ' メール送信の確認（3択）
mailChoice = MsgBox("メールを送信しますか？" & vbCrLf & vbCrLf & _
                    "【はい】→ そのまま送信（" & createdFiles.count & "件）" & vbCrLf & _
                    "【いいえ】→ 確認してから送信" & vbCrLf & _
                    "【キャンセル】→ 送信しない", _
                    vbYesNoCancel + vbQuestion, "メール送信")

If mailChoice = vbYes Then
    Call CreateEmails(createdFiles, customerMasterWs, manufacturerMasterWs, holidays, confirmingListWs, True, repMasterWs)
ElseIf mailChoice = vbNo Then
    Call CreateEmails(createdFiles, customerMasterWs, manufacturerMasterWs, holidays, confirmingListWs, False, repMasterWs)
End If
        ' 送付履歴を閉じる
        deliveryHistoryWb.Close SaveChanges:=True
        manufacturerMasterWb.Close SaveChanges:=False
        ' 顧客マスターを閉じる（メール作成後）
        customerMasterWb.Close SaveChanges:=False
    End If
    
    Exit Sub
    
Cleanup:
    Set g_MfgNameCache = Nothing
    Set g_MfgDaysCache = Nothing
    Set g_CustDaysCache = Nothing
    Set g_CustRetentionCache = Nothing
    Set g_CustRouteCache = Nothing
    Set g_ConfirmCache = Nothing
    Set g_StorageCache = Nothing
    g_SourceData = Empty
    Application.Calculation = origCalcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub


' ============================================
' キャッシュ構築：メーカー一覧
' ============================================
Sub BuildManufacturerCache(manufacturerMasterWs As Worksheet)
    Set g_MfgNameCache = CreateObject("Scripting.Dictionary")
    Set g_MfgDaysCache = CreateObject("Scripting.Dictionary")
    
    If manufacturerMasterWs Is Nothing Then Exit Sub
    
    Dim lastRow As Long
    Dim i As Long
    Dim key As String
    Dim daysValue As Variant
    
    lastRow = manufacturerMasterWs.Cells(manufacturerMasterWs.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        key = Trim(CStr(manufacturerMasterWs.Cells(i, 1).Value))
        If key <> "" And Not g_MfgNameCache.Exists(key) Then
            g_MfgNameCache.Add key, Trim(CStr(manufacturerMasterWs.Cells(i, 2).Value))
            
            daysValue = manufacturerMasterWs.Cells(i, 3).Value
            If IsNumeric(daysValue) And CStr(daysValue) <> "" Then
                g_MfgDaysCache.Add key, CLng(daysValue)
            Else
                g_MfgDaysCache.Add key, 2  ' デフォルト値
            End If
        End If
    Next i
End Sub

' ============================================
' キャッシュ構築：顧客マスター
' ============================================
Sub BuildCustomerCache(customerMasterWs As Worksheet)
    Set g_CustDaysCache = CreateObject("Scripting.Dictionary")
    Set g_CustRetentionCache = CreateObject("Scripting.Dictionary")
    Set g_CustRouteCache = CreateObject("Scripting.Dictionary")
    
    If customerMasterWs Is Nothing Then Exit Sub
    
    Dim lastRow As Long
    Dim i As Long
    Dim key As String
    Dim retentionValue As Variant
    
    lastRow = customerMasterWs.Cells(customerMasterWs.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        key = Trim(CStr(customerMasterWs.Cells(i, 1).Value))
        If key <> "" And Not g_CustDaysCache.Exists(key) Then
            ' B列: 配送曜日
            g_CustDaysCache.Add key, Trim(CStr(customerMasterWs.Cells(i, 2).Value))
            
            ' C列: 保持日数
            retentionValue = customerMasterWs.Cells(i, 3).Value
            If IsNumeric(retentionValue) And CStr(retentionValue) <> "" And CLng(retentionValue) > 0 Then
                g_CustRetentionCache.Add key, CLng(retentionValue)
            Else
                g_CustRetentionCache.Add key, 0
            End If
            
            ' D列: 路線便フラグ
            g_CustRouteCache.Add key, (Trim(CStr(customerMasterWs.Cells(i, 4).Value)) <> "")
        End If
    Next i
End Sub

' ============================================
' キャッシュ構築：確認中テーブル
' ============================================
Sub BuildConfirmingCache(confirmingListWs As Worksheet)
    Set g_ConfirmCache = CreateObject("Scripting.Dictionary")
    
    If confirmingListWs Is Nothing Then Exit Sub
    
    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = confirmingListWs.ListObjects("確認中テーブル")
    On Error GoTo 0
    
    If tbl Is Nothing Then Exit Sub
    If tbl.ListRows.Count = 0 Then Exit Sub
    
    Dim i As Long
    Dim key As String
    Dim col8Val As String
    Dim col9Val As String
    Dim col10Val As Variant
    
    ' 配列一括読み取り
    Dim tblData As Variant
    tblData = tbl.DataBodyRange.Value
    
    Dim rowCount As Long
    rowCount = UBound(tblData, 1)
    
    For i = 1 To rowCount
        key = Trim(CStr(tblData(i, 4))) & "|" & Trim(CStr(tblData(i, 5)))
        
        If key <> "|" And Not g_ConfirmCache.Exists(key) Then
            col8Val = Trim(CStr(tblData(i, 8)))
            col9Val = Trim(CStr(tblData(i, 9)))
            col10Val = tblData(i, 10)
            
            g_ConfirmCache.Add key, Array(col8Val, col9Val, col10Val)
        End If
    Next i
End Sub

' ============================================
' キャッシュ構築：保管場所（sourceWsから注番→保管場所）
' ============================================
Sub BuildStorageCache(sourceWs As Worksheet, cols As Object)
    Set g_StorageCache = CreateObject("Scripting.Dictionary")
    
    If Not cols.Exists("保管場所") Then Exit Sub
    If Not cols.Exists("受発注伝票") Then Exit Sub
    
    Dim lastRow As Long
    Dim i As Long
    Dim orderNum As String
    Dim storagePlace As String
    
    lastRow = sourceWs.Cells(sourceWs.Rows.Count, cols("受発注伝票")).End(xlUp).Row
    
    For i = 7 To lastRow
        orderNum = Trim(CStr(g_SourceData(i, cols("受発注伝票"))))
        If orderNum <> "" And Not g_StorageCache.Exists(orderNum) Then
            storagePlace = Trim(CStr(g_SourceData(i, cols("保管場所"))))
            If storagePlace <> "" Then
                g_StorageCache.Add orderNum, storagePlace
            End If
        End If
    Next i
End Sub
' ============================================
' 【v4.3新機能】注番を顧客ごとにグループ化
' ============================================
Function GroupOrderNumbersByCustomer(sourceWs As Worksheet, cols As Object, _
                                     orderNumbers As Collection) As Object
    Dim customerGroups As Object
    Set customerGroups = CreateObject("Scripting.Dictionary")
    
    Dim orderNum As Variant
    Dim lastRow As Long
    Dim i As Long
    Dim customerName As String
    
    lastRow = sourceWs.Cells(sourceWs.Rows.count, cols("受発注伝票")).End(xlUp).Row
    
    ' 各注番の顧客名を検索してグループ化
    For Each orderNum In orderNumbers
        customerName = ""
        
        ' 注番から顧客名を検索
        For i = 7 To lastRow
            If Trim(g_SourceData(i, cols("受発注伝票"))) = Trim(CStr(orderNum)) Then
                customerName = Trim(g_SourceData(i, cols("受注先")))
                Exit For
            End If
        Next i
        
        ' 顧客名が見つかった場合、グループに追加
        If customerName <> "" Then
            If Not customerGroups.Exists(customerName) Then
                Dim newCollection As Collection
                Set newCollection = New Collection
                customerGroups.Add customerName, newCollection
            End If
            
            ' 重複チェック
            Dim isDuplicate As Boolean
            isDuplicate = False
            Dim existingNum As Variant
            For Each existingNum In customerGroups(customerName)
                If CStr(existingNum) = CStr(orderNum) Then
                    isDuplicate = True
                    Exit For
                End If
            Next existingNum
            
            If Not isDuplicate Then
                customerGroups(customerName).Add CStr(orderNum)
            End If
        End If
    Next orderNum
    
    Set GroupOrderNumbersByCustomer = customerGroups
End Function
' ============================================
' 【v6.1】複数注番で納期回答書を作成（showPrice削除版）
' ============================================
Function CreateDeliveryReportByOrderNumbers(sourceWs As Worksheet, cols As Object, _
                                           customerName As String, orderNumbers As Collection, _
                                           manufacturerMasterWs As Worksheet, _
                                           Optional holidays As Object = Nothing, _
                                           Optional confirmingListWs As Worksheet = Nothing, _
                                           Optional customerMasterWs As Worksheet = Nothing, _
                                           Optional repName As String = "", _
                                           Optional repMasterWs As Worksheet = Nothing) As Variant
    
    Dim newWb As Workbook
    Dim newWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' 【追加】欠品情報リスト
    Dim stockoutInfoList As Collection
    Set stockoutInfoList = New Collection
    ' 【追加】送り状情報リスト
    Dim trackingInfoList As Collection
    Set trackingInfoList = New Collection
    ' 【追加】分納情報リスト
    Dim bunnoInfoList As Collection
    Set bunnoInfoList = New Collection
    Dim commentDetail As String
    Dim productName As String
    Dim shortProductName As String
    Dim manufacturerName As String
    Dim qty As Variant
    Dim itemGroupCode As String
    Dim extComment As String
    Dim trackingInfo As Variant
    
    ' 新規ワークブックを作成
    Set newWb = Workbooks.Add
    Set newWs = newWb.Sheets(1)
    
    Dim sheetName As String
    If repName <> "" And repName <> "__OTHER__" Then
        sheetName = customerName & "_" & repName & "様"
    Else
        sheetName = customerName & "様"
    End If
    If Len(sheetName) > 31 Then
        sheetName = Left(sheetName, 31)
    End If
    newWs.Name = sheetName
    
    Call CreateHeader(newWs, customerName, repName)
    
    Dim currentRow As Long
    currentRow = 7
    
    lastRow = sourceWs.Cells(sourceWs.Rows.count, cols("受発注伝票")).End(xlUp).Row
    
    ' 担当者フィルタ用：マスター登録済みリストをキャッシュ
    Dim registeredRepListON As Collection
    Set registeredRepListON = Nothing
    If repName = "__OTHER__" And Not repMasterWs Is Nothing Then
        Set registeredRepListON = GetRepList(customerName, repMasterWs)
    End If
    
    ' 全ての注番のデータをコピー
    Dim orderNum As Variant
    For Each orderNum In orderNumbers
        For i = 7 To lastRow
            If Trim(g_SourceData(i, cols("受発注伝票"))) = Trim(CStr(orderNum)) Then
                Dim rejectReason As String
                rejectReason = ""
                If cols.Exists("拒否理由") Then
                    rejectReason = Trim(g_SourceData(i, cols("拒否理由")))
                End If
                
                Dim documentType As String
                documentType = ""
                If cols.Exists("伝票タイプ") Then
                    documentType = Trim(g_SourceData(i, cols("伝票タイプ")))
                End If
                
                ' コメント（社内）に##または＃＃が含まれている場合は除外
                Dim internalComment As String
                Dim isExcluded As Boolean
                isExcluded = False
                If cols.Exists("コメント（社内）") Then
                    internalComment = Trim(g_SourceData(i, cols("コメント（社内）")))
                    If InStr(internalComment, "##") > 0 Or InStr(internalComment, "＃＃") > 0 Then
                        isExcluded = True
                    End If
                End If
                
                ' 条件チェック（注番指定では送付済みチェックをスキップ）
                If rejectReason <> "明細削除" And _
                   (documentType = "【受注】直送販売" Or documentType = "【受注】在庫販売") And _
                   Not isExcluded Then
                    
                    ' 担当者フィルタ
                    If repName <> "" Then
                        Dim repCellValueON As String
                        repCellValueON = ""
                        If cols.Exists("得意先担当者") Then
                            repCellValueON = Trim(g_SourceData(i, cols("得意先担当者")))
                        End If
                        If Not ShouldIncludeForRep(repCellValueON, repName, registeredRepListON) Then GoTo NextRowON
                    End If
                    
                    Dim deliveryStatusForCopy As String
deliveryStatusForCopy = CopyDataRow(sourceWs, newWs, cols, i, currentRow, _
                 manufacturerMasterWs, holidays, confirmingListWs, False, _
                 customerMasterWs)

' 【追加】欠品情報を収集
commentDetail = ""
If cols.Exists("コメント（明細）") Then
    commentDetail = Trim(g_SourceData(i, cols("コメント（明細）")))
End If

productName = Trim(g_SourceData(i, cols("品名")))
Dim hasBunnoCommentON As Boolean
hasBunnoCommentON = (InStr(commentDetail, "分納:") > 0 Or InStr(commentDetail, "分納：") > 0)

' 【追加】出荷ステータス取得（欠品判定に使用）
Dim shipStatusForON As String
shipStatusForON = ""
If cols.Exists("出荷ステータス") Then
    shipStatusForON = Trim(g_SourceData(i, cols("出荷ステータス")))
End If

' 処理完了＝出荷処理済みなら欠品は解消済みなのでスキップ
If InStr(commentDetail, "欠品中") > 0 And productName <> "送料" And Not hasBunnoCommentON _
    And shipStatusForON <> "処理完了" Then
    itemGroupCode = Trim(g_SourceData(i, cols("品目Group")))
    manufacturerName = GetManufacturerName(itemGroupCode, manufacturerMasterWs)
    If manufacturerName = "" Then
        If cols.Exists("メーカー") Then
            manufacturerName = Trim(g_SourceData(i, cols("メーカー")))
        Else
            manufacturerName = itemGroupCode
        End If
    End If
    
    shortProductName = productName
    If Len(shortProductName) > 25 Then
        shortProductName = Left(shortProductName, 25) & "..."
    End If
    qty = g_SourceData(i, cols("受注数量"))
    
    ' アバウト納期を抽出
    Dim approxDeliveryForCopy As String
    approxDeliveryForCopy = ExtractApproxDelivery(commentDetail)
    
    stockoutInfoList.Add Array(manufacturerName, shortProductName, qty, deliveryStatusForCopy, approxDeliveryForCopy)
End If
                    
                    ' 【追加】送り状情報を収集（複数対応）
If cols.Exists("コメント（社外）") Then
    extComment = Trim(g_SourceData(i, cols("コメント（社外）")))
    Dim trackingResults As Collection
    Set trackingResults = ExtractTrackingInfo(extComment)
    
    If trackingResults.count > 0 Then
        productName = Trim(g_SourceData(i, cols("品名")))
        If productName <> "送料" Then
        itemGroupCode = Trim(g_SourceData(i, cols("品目Group")))
        manufacturerName = GetManufacturerName(itemGroupCode, manufacturerMasterWs)
        If manufacturerName = "" Then
            If cols.Exists("メーカー") Then
                manufacturerName = Trim(g_SourceData(i, cols("メーカー")))
            Else
                manufacturerName = itemGroupCode
            End If
        End If
        
        shortProductName = productName
        If Len(shortProductName) > 25 Then
            shortProductName = Left(shortProductName, 25) & "..."
        End If
        qty = g_SourceData(i, cols("受注数量"))
        
        Dim tInfo As Variant
        For Each tInfo In trackingResults
            trackingInfoList.Add Array(manufacturerName, shortProductName, qty, tInfo(0), tInfo(1))
        Next tInfo
        End If
    End If
End If

                    ' 【追加】分納情報を収集
                    Dim bunnoInfoForCopy As Collection
                    Set bunnoInfoForCopy = ExtractBunnoInfo(commentDetail)
                    
                    ' 処理完了＝出荷処理済みなら分納コメントは残骸なのでスキップ
                    If bunnoInfoForCopy.count > 0 And productName <> "送料" _
                        And shipStatusForON <> "処理完了" Then
                        ' メーカー名・品名を取得
                        productName = Trim(g_SourceData(i, cols("品名")))
                        itemGroupCode = Trim(g_SourceData(i, cols("品目Group")))
                        manufacturerName = GetManufacturerName(itemGroupCode, manufacturerMasterWs)
                        If manufacturerName = "" Then
                            If cols.Exists("メーカー") Then
                                manufacturerName = Trim(g_SourceData(i, cols("メーカー")))
                            Else
                                manufacturerName = itemGroupCode
                            End If
                        End If
                        
                        Dim isShipRuleForCopy As Boolean
                        isShipRuleForCopy = False
                        
                        Dim storagePlaceForBunno As String
                        storagePlaceForBunno = ""
                        If cols.Exists("保管場所") Then
                            storagePlaceForBunno = Trim(g_SourceData(i, cols("保管場所")))
                        End If

                        If storagePlaceForBunno = "転送中（直送用）" Then
                            isShipRuleForCopy = True
                        ElseIf documentType = "【受注】在庫販売" Then
                            If customerName <> Trim(g_SourceData(i, cols("出荷先名"))) Then
                                isShipRuleForCopy = True
                            End If
                        End If
                        
                        Dim isRosenbinForCopy As Boolean
                        isRosenbinForCopy = IsRouteDelivery(customerName, customerMasterWs)
                        
                        ' 在庫販売 + 路線便 → 出荷予定扱い
                        If Not isShipRuleForCopy And isRosenbinForCopy And documentType = "【受注】在庫販売" Then
                            isShipRuleForCopy = True
                        End If
                        
                        ' 【v12.1】分納の計算済み納期を事前に計算して保存
                        Dim daysToAddForCopy As Long
                        daysToAddForCopy = GetDeliveryDaysToAdd(itemGroupCode, manufacturerMasterWs)
                        
                        Dim orderNumForCopy As String
                        Dim detailNumForCopy As String
                        orderNumForCopy = Trim(g_SourceData(i, cols("受発注伝票")))
                        detailNumForCopy = Trim(g_SourceData(i, cols("明細")))
                        
                        Dim bunnoInfoWithCalcCopy As Collection
                        Set bunnoInfoWithCalcCopy = New Collection
                        
                        Dim bunnoLineForCopy As Variant
                        For Each bunnoLineForCopy In bunnoInfoForCopy
                            Dim calcDateForCopy As String
                            calcDateForCopy = CalculateBunnoDate(CStr(bunnoLineForCopy(1)), isShipRuleForCopy, _
                                                                  daysToAddForCopy, holidays, _
                                                                  confirmingListWs, orderNumForCopy, detailNumForCopy, _
                                                                  isRosenbinForCopy)
                            
                            Dim locationForCopy As String
                            locationForCopy = ""
                            If UBound(bunnoLineForCopy) >= 2 Then
                                locationForCopy = bunnoLineForCopy(2)
                            End If
                            
                            bunnoInfoWithCalcCopy.Add Array(bunnoLineForCopy(0), bunnoLineForCopy(1), locationForCopy, calcDateForCopy)
                        Next bunnoLineForCopy
                        
                        bunnoInfoList.Add Array(manufacturerName, productName, _
                            g_SourceData(i, cols("受注数量")), _
                            bunnoInfoWithCalcCopy, isShipRuleForCopy, itemGroupCode, _
                            orderNumForCopy, detailNumForCopy, _
                            isRosenbinForCopy)
                    End If
                    
                    currentRow = currentRow + 1
                End If
            End If
NextRowON:
        Next i
    Next orderNum
    
    If currentRow = 7 Then
        If repName = "" Then
            MsgBox "指定された注番は条件に該当しませんでした：" & customerName, vbExclamation
        End If
        newWb.Close SaveChanges:=False
        CreateDeliveryReportByOrderNumbers = Empty
        Exit Function
    End If
    
    ' 【変更】欠品リスト、分納リストを渡す
    Call FormatReport(newWs, currentRow - 1, trackingInfoList, stockoutInfoList, bunnoInfoList, manufacturerMasterWs, holidays, confirmingListWs)
    
    Dim fileName As String
    Dim filePath As String
    Dim savePath As String
    Dim subFolder As String
    Dim orderNumberSuffix As String
    
    ' ファイル名の注番部分を決定
    If orderNumbers.count = 1 Then
        orderNumberSuffix = CStr(orderNumbers(1))
    Else
        orderNumberSuffix = "複数注番"
    End If
    
    savePath = ThisWorkbook.Path
    Dim baseFolder As String
    baseFolder = savePath & "\納期回答書"
    
    ' 納期回答書フォルダがなければ作成
    If Dir(baseFolder, vbDirectory) = "" Then
        MkDir baseFolder
    End If
    
    ' サブフォルダがまだ決まっていない場合は決定
    If g_CurrentSubFolder = "" Then
        Dim todayStr As String
        Dim folderCount As Long
        Dim f As String
        
        todayStr = Format(g_ExecutionTime, "m月d日(aaa)")
        
        ' 同じ日のフォルダを数える
        folderCount = 0
        f = Dir(baseFolder & "\*", vbDirectory)
        Do While f <> ""
            If f <> "." And f <> ".." Then
                If Left(f, Len(todayStr)) = todayStr Then
                    folderCount = folderCount + 1
                End If
            End If
            f = Dir()
        Loop
        
        folderCount = folderCount + 1
        g_CurrentSubFolder = baseFolder & "\" & todayStr & "_" & ToCircledNumber(folderCount) & "回目"
        
        If Dir(g_CurrentSubFolder, vbDirectory) = "" Then
            MkDir g_CurrentSubFolder
        End If
    End If
    
    subFolder = g_CurrentSubFolder
    
    If repName <> "" And repName <> "__OTHER__" Then
        fileName = "納期回答書_" & Replace(customerName, "/", "_") & "様_" & repName & "様_" & orderNumberSuffix & "_" & Format(Date, "yyyymmdd") & ".xlsx"
    Else
        fileName = "納期回答書_" & Replace(customerName, "/", "_") & "様_" & orderNumberSuffix & "_" & Format(Date, "yyyymmdd") & ".xlsx"
    End If
    
    filePath = subFolder & "\" & fileName

    
    On Error Resume Next
    newWb.SaveAs filePath
    If Err.Number <> 0 Then
        MsgBox "ファイルの保存に失敗しました：" & customerName, vbExclamation
        newWb.Close SaveChanges:=False
        CreateDeliveryReportByOrderNumbers = Empty
        Exit Function
    End If
    On Error GoTo 0
    
    newWb.Close SaveChanges:=False
    
    ' ファイルパス、顧客名、欠品情報、送り状情報、分納情報を返す
    CreateDeliveryReportByOrderNumbers = Array(customerName, filePath, stockoutInfoList, trackingInfoList, bunnoInfoList)
End Function

' ============================================
' 顧客リストのチェック（Collection対応）
' ============================================
Function CheckCustomerMaster(customerList As Collection, customerMasterWs As Worksheet) As String
    Dim missingList As String
    Dim customerKey As Variant
    Dim customerName As String
    Dim found As Boolean
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim hasEmail As Boolean
    
    missingList = ""
    lastRow = customerMasterWs.Cells(customerMasterWs.Rows.count, 1).End(xlUp).Row
    
    For Each customerKey In customerList
        customerName = Split(CStr(customerKey), "|")(1)
        found = False
        
        For i = 2 To lastRow
            If Trim(customerMasterWs.Cells(i, 1).Value) = customerName Then
                hasEmail = False
                ' ★ E列（5列目）から開始に変更
                For j = 5 To customerMasterWs.Cells(i, customerMasterWs.Columns.count).End(xlToLeft).Column
                    If Trim(customerMasterWs.Cells(i, j).Value) <> "" Then
                        hasEmail = True
                        Exit For
                    End If
                Next j
                
                If hasEmail Then
                    found = True
                    Exit For
                End If
            End If
        Next i
        
        If Not found Then
            missingList = missingList & "・" & customerName & vbCrLf
        End If
    Next
    
    CheckCustomerMaster = missingList
End Function

' ============================================
' 【v4.8修正】送付履歴ファイルの初期化（確認中一覧に「除外」追加）
' ============================================
Sub InitializeDeliveryHistory(filePath As String)
    Dim newWb As Workbook
    Dim wsHistory As Worksheet
    Dim wsConfirming As Worksheet
    Dim tblRange As Range
    Dim tbl As ListObject
    
    Set newWb = Workbooks.Add
    
    ' ===== 送付履歴シート =====
    Set wsHistory = newWb.Sheets(1)
    wsHistory.Name = "送付履歴"
    
    With wsHistory
        ' ヘッダー行を作成（9列構成）
        .Cells(1, 1).Value = "送付日時"
        .Cells(1, 2).Value = "受注日"
        .Cells(1, 3).Value = "顧客名"
        .Cells(1, 4).Value = "受発注伝票"
        .Cells(1, 5).Value = "明細"
        .Cells(1, 6).Value = "メーカー名"
        .Cells(1, 7).Value = "品名"
        .Cells(1, 8).Value = "納期回答"
        .Cells(1, 9).Value = "送付者"
        
        ' ダミー行を追加（テーブル作成用）
        .Cells(2, 1).Value = ""
        
        ' テーブルとして設定
        Set tblRange = .Range("A1:I2")
        Set tbl = .ListObjects.Add(xlSrcRange, tblRange, , xlYes)
        tbl.Name = "送付履歴テーブル"
        tbl.TableStyle = "TableStyleMedium2"
        
        ' ダミー行を削除
        tbl.ListRows(1).Delete
        
        ' 列幅調整
        .Columns("A").ColumnWidth = 17
        .Columns("B").ColumnWidth = 12
        .Columns("C").ColumnWidth = 25
        .Columns("D").ColumnWidth = 15
        .Columns("E").ColumnWidth = 8
        .Columns("F").ColumnWidth = 20
        .Columns("G").ColumnWidth = 47
        .Columns("H").ColumnWidth = 22
        .Columns("I").ColumnWidth = 17.88
    End With
    
    ' ===== 確認中一覧シート =====
    Set wsConfirming = newWb.Sheets.Add(After:=wsHistory)
    wsConfirming.Name = "確認中一覧"
    
    With wsConfirming
        ' ヘッダー行を作成（11列構成）
        .Cells(1, 1).Value = "送付日時"
        .Cells(1, 2).Value = "受注日"
        .Cells(1, 3).Value = "顧客名"
        .Cells(1, 4).Value = "受発注伝票"
        .Cells(1, 5).Value = "明細"
        .Cells(1, 6).Value = "メーカー名"
        .Cells(1, 7).Value = "品名"
        .Cells(1, 8).Value = "問合せ状況"
        .Cells(1, 9).Value = "ステータス"
        .Cells(1, 10).Value = "受注納期"
        .Cells(1, 11).Value = "送付者"
        
        ' ダミー行を追加（テーブル作成用）
        .Cells(2, 1).Value = ""
        
        ' テーブルとして設定
        Set tblRange = .Range("A1:K2")
        Set tbl = .ListObjects.Add(xlSrcRange, tblRange, , xlYes)
        tbl.Name = "確認中テーブル"
        tbl.TableStyle = "TableStyleMedium1"
        
        ' ダミー行を削除
        tbl.ListRows(1).Delete
        
        ' 列幅調整
        .Columns("A").ColumnWidth = 17
        .Columns("B").ColumnWidth = 12
        .Columns("C").ColumnWidth = 25
        .Columns("D").ColumnWidth = 15
        .Columns("E").ColumnWidth = 8
        .Columns("F").ColumnWidth = 20
        .Columns("G").ColumnWidth = 47
        .Columns("H").ColumnWidth = 13
        .Columns("I").ColumnWidth = 12
        .Columns("J").ColumnWidth = 18
        .Columns("K").ColumnWidth = 17.88
    End With
    
    newWb.SaveAs filePath, FileFormat:=51
    newWb.Close SaveChanges:=False
End Sub
' ============================================
' 【v4.8修正】送付履歴の読み込み（確認中一覧の「除外」も含む）
' ============================================
Function LoadDeliveryHistory(wsHistory As Worksheet, wsConfirming As Worksheet, _
                             customerMasterWs As Worksheet, _
                             Optional holidays As Object = Nothing) As Object
    Dim sentOrders As Object
    Set sentOrders = CreateObject("Scripting.Dictionary")
    
    Dim tbl As ListObject
    Dim tblConfirming As ListObject
    Dim i As Long
    Dim orderNumber As String
    Dim detailNumber As String
    Dim deliveryStatus As String
    Dim historyKey As String
    Dim orderDateValue As Variant
    Dim orderDate As Date
    Dim today As Date
    Dim inquiryStatus As String
    Dim sentDateTime As Date
    Dim customerName As String
    Dim retentionDays As Long
    Dim businessDaysPassed As Long
    
    today = Date
    
    ' ===== 送付履歴テーブルを読み込み =====
    On Error Resume Next
    Set tbl = wsHistory.ListObjects("送付履歴テーブル")
    On Error GoTo 0
    
    If Not tbl Is Nothing And tbl.ListRows.Count > 0 Then
        Dim histData As Variant
        histData = tbl.DataBodyRange.Value
        Dim histRows As Long
        histRows = UBound(histData, 1)
        
        For i = 1 To histRows
            sentDateTime = 0
            On Error Resume Next
            sentDateTime = CDate(histData(i, 1))
            On Error GoTo 0
            
            orderDateValue = histData(i, 2)
            customerName = Trim(CStr(histData(i, 3)))
            orderNumber = Trim(CStr(histData(i, 4)))
            detailNumber = Trim(CStr(histData(i, 5)))
            deliveryStatus = Trim(CStr(histData(i, 8)))
            
            historyKey = orderNumber & "|" & detailNumber
            
            ' 確定伝票を登録（納期回答の値を保存）
            If historyKey <> "|" And deliveryStatus <> "" And deliveryStatus <> "確認中" Then
                orderDate = 0
                On Error Resume Next
                orderDate = Int(CDate(orderDateValue))
                On Error GoTo 0
                
                ' ★ 保持日数をチェック
                retentionDays = GetRetentionDays(customerName, customerMasterWs)
                businessDaysPassed = CountBusinessDaysBetween(Int(sentDateTime), today, holidays)
                
                ' 保持日数による判定
' 分納完了は日付・保持日数に関わらず常にスキップ対象
If deliveryStatus = "分納完了" Then
    If Not sentOrders.Exists(historyKey) Then
        sentOrders.Add historyKey, deliveryStatus
    End If
ElseIf retentionDays = 0 Then
    ' 従来通り：受注日が今日より前なら除外
    If orderDate > 0 And orderDate < today Then
        If Not sentOrders.Exists(historyKey) Then
            sentOrders.Add historyKey, deliveryStatus
        End If
    End If
Else
    ' 保持日数設定あり：営業日経過で判定
    If businessDaysPassed > retentionDays Then
        If Not sentOrders.Exists(historyKey) Then
            sentOrders.Add historyKey, deliveryStatus
        End If
    End If
End If
            End If
        Next i
    End If
    
    ' ===== 確認中テーブルから「除外」の伝票も読み込み =====
    On Error Resume Next
    Set tblConfirming = wsConfirming.ListObjects("確認中テーブル")
    On Error GoTo 0
    
    If Not tblConfirming Is Nothing And tblConfirming.ListRows.Count > 0 Then
        Dim confData As Variant
        confData = tblConfirming.DataBodyRange.Value
        Dim confRows As Long
        confRows = UBound(confData, 1)
        
        For i = 1 To confRows
            inquiryStatus = Trim(CStr(confData(i, 8)))
            
            If inquiryStatus = "除外" Then
                orderNumber = Trim(CStr(confData(i, 4)))
                detailNumber = Trim(CStr(confData(i, 5)))
                historyKey = orderNumber & "|" & detailNumber
                
                If historyKey <> "|" Then
                    If Not sentOrders.Exists(historyKey) Then
                        sentOrders.Add historyKey, "除外"
                    End If
                End If
            End If
        Next i
    End If
    
    Set LoadDeliveryHistory = sentOrders
End Function

' ============================================
' 【v4.6修正】送付履歴への書き込み（重複チェック追加）
' ============================================
Sub SaveDeliveryHistory(ws As Worksheet, newOrders As Collection)
    Dim tbl As ListObject
    Dim orderData As Variant
    Dim currentDateTime As Date
    Dim userName As String
    Dim i As Long, j As Long
    
    If newOrders.Count = 0 Then Exit Sub
    
    currentDateTime = Now
    userName = Environ("USERNAME")
    
    On Error Resume Next
    Set tbl = ws.ListObjects("送付履歴テーブル")
    On Error GoTo 0
    If tbl Is Nothing Then
        MsgBox "送付履歴テーブルが見つかりません。", vbExclamation
        Exit Sub
    End If
    
    ' テーブルのフィルター解除
    On Error Resume Next
    tbl.AutoFilter.ShowAllData
    On Error GoTo 0
    
    ' === 配列一括読み取り ===
    Dim tblData As Variant
    Dim existRows As Long
    existRows = 0
    If tbl.ListRows.Count > 0 Then
        tblData = tbl.DataBodyRange.Value
        existRows = UBound(tblData, 1)
    End If
    Dim colCount As Long
    colCount = tbl.ListColumns.Count
    
    ' === 既存データをDictionaryに格納（空行除外）===
    Dim existingKeys As Object
    Set existingKeys = CreateObject("Scripting.Dictionary")
    Dim validRows As Collection
    Set validRows = New Collection
    
    If existRows > 0 Then
        For i = 1 To existRows
            Dim keyVal As String
            keyVal = Trim(CStr(tblData(i, 4))) & "|" & Trim(CStr(tblData(i, 5)))
            If Trim(CStr(tblData(i, 4))) <> "" Then
                validRows.Add i
                If Not existingKeys.Exists(keyVal) Then
                    existingKeys.Add keyVal, i
                End If
            End If
        Next i
    End If
    
    ' === 新規データの処理（メモリ上で）===
    Dim newItems As Collection
    Set newItems = New Collection
    
    For i = 1 To newOrders.Count
        orderData = newOrders(i)
        Dim checkKey As String
        checkKey = orderData(2) & "|" & orderData(3)
        
        If existingKeys.Exists(checkKey) Then
            ' 重複：納品済みの場合は更新
            If orderData(6) = "納品済み" Then
                Dim updateRow As Long
                updateRow = existingKeys(checkKey)
                tblData(updateRow, 8) = "納品済み"
            End If
        Else
            ' 新規追加用の配列を作成
            Dim newRow As Variant
            ReDim newRow(1 To colCount)
            newRow(1) = currentDateTime   ' 送付日時
            newRow(2) = orderData(1)      ' 受注日
            newRow(3) = orderData(0)      ' 顧客名
            newRow(4) = orderData(2)      ' 受発注伝票
            newRow(5) = orderData(3)      ' 明細
            newRow(6) = orderData(4)      ' メーカー名
            newRow(7) = orderData(5)      ' 品名
            newRow(8) = orderData(6)      ' 納期回答
            newRow(9) = userName          ' 送付者
            newItems.Add newRow
            
            ' 後続の重複チェック用にDictionaryにも追加
            If Not existingKeys.Exists(checkKey) Then
                existingKeys.Add checkKey, -1
            End If
        End If
    Next i
    
    ' === 結果配列を構築 ===
    Dim totalRows As Long
    totalRows = validRows.Count + newItems.Count
    
    If totalRows = 0 Then Exit Sub
    Dim result() As Variant
    ReDim result(1 To totalRows, 1 To colCount)
    Dim r As Long
    r = 0
    
    ' 新規アイテムを先頭に
    Dim ni As Variant
    For Each ni In newItems
        r = r + 1
        For j = 1 To colCount
            result(r, j) = ni(j)
        Next j
    Next ni
    
    ' 既存の有効行（更新済みデータを含む）
    Dim vi As Variant
    For Each vi In validRows
        r = r + 1
        For j = 1 To colCount
            result(r, j) = tblData(CLng(vi), j)
        Next j
    Next vi
    
    ' === テーブルをクリアして一括書き戻し ===
    If tbl.ListRows.Count > 0 Then
        tbl.DataBodyRange.Delete
    End If
    
    On Error Resume Next
    If tbl.ListRows.Count = 0 Then
        tbl.ListRows.Add
    End If
    On Error GoTo 0
    
    tbl.Resize tbl.Range.Resize(totalRows + 1, colCount)
    tbl.DataBodyRange.Value = result
    
    ' === 表示形式を適用 ===
    On Error Resume Next
    If tbl.ListRows.Count > 0 And Not tbl.ListColumns("送付日時").DataBodyRange Is Nothing Then
        tbl.ListColumns("送付日時").DataBodyRange.NumberFormat = "mm/dd hh:nn"
    End If
    If tbl.ListRows.Count > 0 And Not tbl.ListColumns("受注日").DataBodyRange Is Nothing Then
        tbl.ListColumns("受注日").DataBodyRange.NumberFormat = "mm/dd"
    End If
    On Error GoTo 0
    
    ' === 送付日時の降順でソート ===
    If tbl.ListRows.Count > 0 Then
        On Error Resume Next
        With tbl.Sort
            .SortFields.Clear
            .SortFields.Add Key:=tbl.ListColumns("送付日時").DataBodyRange, Order:=xlDescending
            .Header = xlYes
            .Apply
        End With
        On Error GoTo 0
    End If
End Sub

' ============================================
' 【v4.8修正】確認中一覧への書き込み（行追加時に入力規則設定）
' ============================================
Sub SaveConfirmingList(ws As Worksheet, newOrders As Collection)
    Dim tbl As ListObject
    Dim orderData As Variant
    Dim currentDateTime As Date
    Dim userName As String
    Dim i As Long, j As Long
    
    If newOrders.Count = 0 Then Exit Sub
    
    currentDateTime = Now
    userName = Environ("USERNAME")
    
    On Error Resume Next
    Set tbl = ws.ListObjects("確認中テーブル")
    On Error GoTo 0
    If tbl Is Nothing Then
        MsgBox "確認中テーブルが見つかりません。", vbExclamation
        Exit Sub
    End If
    
    ' テーブルのフィルター解除
    On Error Resume Next
    tbl.AutoFilter.ShowAllData
    On Error GoTo 0
    
    ' === 配列一括読み取り ===
    Dim tblData As Variant
    Dim existRows As Long
    existRows = 0
    If tbl.ListRows.Count > 0 Then
        tblData = tbl.DataBodyRange.Value
        existRows = UBound(tblData, 1)
    End If
    Dim colCount As Long
    colCount = tbl.ListColumns.Count
    
    ' === 既存データをDictionaryに格納（空行除外）===
    Dim existingKeys As Object
    Set existingKeys = CreateObject("Scripting.Dictionary")
    Dim validRows As Collection
    Set validRows = New Collection
    
    If existRows > 0 Then
        For i = 1 To existRows
            Dim keyVal As String
            keyVal = Trim(CStr(tblData(i, 4))) & "|" & Trim(CStr(tblData(i, 5)))
            If Trim(CStr(tblData(i, 4))) <> "" Then
                validRows.Add i
                If Not existingKeys.Exists(keyVal) Then
                    existingKeys.Add keyVal, i
                End If
            End If
        Next i
    End If
    
    ' === 新規データの処理（メモリ上で）===
    Dim newItems As Collection
    Set newItems = New Collection
    
    For i = 1 To newOrders.Count
        orderData = newOrders(i)
        Dim checkKey As String
        checkKey = orderData(2) & "|" & orderData(3)
        
        If existingKeys.Exists(checkKey) Then
            ' 重複：ステータス（9列目）を更新
            Dim updateRow As Long
            updateRow = existingKeys(checkKey)
            If updateRow > 0 Then
                tblData(updateRow, 9) = orderData(6)
            End If
        Else
            ' 新規追加用の配列を作成
            Dim newRow As Variant
            ReDim newRow(1 To colCount)
            newRow(1) = currentDateTime   ' 送付日時
            newRow(2) = orderData(1)      ' 受注日
            newRow(3) = orderData(0)      ' 顧客名
            newRow(4) = orderData(2)      ' 受発注伝票
            newRow(5) = orderData(3)      ' 明細
            newRow(6) = orderData(4)      ' メーカー名
            newRow(7) = orderData(5)      ' 品名
            newRow(8) = "未"             ' 問合せ状況
            newRow(9) = orderData(6)      ' ステータス
            newRow(10) = ""              ' 受注納期
            newRow(11) = userName         ' 送付者
            newItems.Add newRow
            
            ' 後続の重複チェック用にDictionaryにも追加
            If Not existingKeys.Exists(checkKey) Then
                existingKeys.Add checkKey, -1
            End If
        End If
    Next i
    
    ' === 結果配列を構築 ===
    Dim totalRows As Long
    totalRows = validRows.Count + newItems.Count
    If totalRows = 0 Then Exit Sub
    
    Dim result() As Variant
    ReDim result(1 To totalRows, 1 To colCount)
    Dim r As Long
    r = 0
    
    ' 新規アイテムを先頭に
    Dim ni As Variant
    For Each ni In newItems
        r = r + 1
        For j = 1 To colCount
            result(r, j) = ni(j)
        Next j
    Next ni
    
    ' 既存の有効行（更新済みデータを含む）
    Dim vi As Variant
    For Each vi In validRows
        r = r + 1
        For j = 1 To colCount
            result(r, j) = tblData(CLng(vi), j)
        Next j
    Next vi
    
    ' === テーブルをクリアして一括書き戻し ===
    If tbl.ListRows.Count > 0 Then
        tbl.DataBodyRange.Delete
    End If
    
    On Error Resume Next
    If tbl.ListRows.Count = 0 Then
        tbl.ListRows.Add
    End If
    On Error GoTo 0
    
    tbl.Resize tbl.Range.Resize(totalRows + 1, colCount)
    tbl.DataBodyRange.Value = result
    
    ' === 問合せ状況列（8列目）の入力規則を一括設定 ===
    On Error Resume Next
    tbl.ListColumns(8).DataBodyRange.Validation.Delete
    tbl.ListColumns(8).DataBodyRange.Validation.Add Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, Formula1:="未,済,回答待ち,除外"
    On Error GoTo 0
    
    ' === 表示形式を適用 ===
    On Error Resume Next
    If tbl.ListRows.Count > 0 And Not tbl.ListColumns("送付日時").DataBodyRange Is Nothing Then
        tbl.ListColumns("送付日時").DataBodyRange.NumberFormat = "mm/dd hh:nn"
    End If
    If tbl.ListRows.Count > 0 And Not tbl.ListColumns("受注日").DataBodyRange Is Nothing Then
        tbl.ListColumns("受注日").DataBodyRange.NumberFormat = "mm/dd"
    End If
    On Error GoTo 0
    
    ' === 送付日時の降順でソート ===
    If tbl.ListRows.Count > 0 Then
        On Error Resume Next
        With tbl.Sort
            .SortFields.Clear
            .SortFields.Add Key:=tbl.ListColumns("送付日時").DataBodyRange, Order:=xlDescending
            .Header = xlYes
            .Apply
        End With
        On Error GoTo 0
    End If
    
    ' I列（9列目）の入力規則を削除（全行）
    On Error Resume Next
    If tbl.ListRows.Count > 0 Then
        tbl.ListColumns(9).DataBodyRange.Validation.Delete
    End If
    On Error GoTo 0
    
    ' 色分け表示（送付後3日以上）
    Call ColorConfirmingList(ws)
End Sub
' ============================================
' 【v4.4新機能】確認中一覧の色分け表示
' ============================================
Sub ColorConfirmingList(ws As Worksheet)
    Dim tbl As ListObject
    Dim i As Long
    Dim sentDate As Date
    Dim today As Date
    Dim daysDiff As Long
    Dim shipStatus As String
    
    today = Date
    
    ' テーブルを取得
    On Error Resume Next
    Set tbl = ws.ListObjects("確認中テーブル")
    On Error GoTo 0
    
    If tbl Is Nothing Or tbl.ListRows.Count = 0 Then
        Exit Sub
    End If
    
    ' テーブルデータを配列で一括読み取り
    Dim tblData As Variant
    tblData = tbl.DataBodyRange.Value
    
    ' カテゴリ別にRangeを収集
    Dim rngShipDone As Range     ' 出荷完了（赤系）
    Dim rngStockoutC As Range    ' 欠品中（薄紫）
    Dim rngBunnoC As Range       ' 分納（薄青）
    Dim rngWeekOld As Range      ' 1週間以上（オレンジ）
    Dim rngThreeDays As Range    ' 3日以上（黄色）
    Dim rngNormal As Range       ' それ以外（色なし）
    
    If IsArray(tblData) Then
        For i = 1 To UBound(tblData, 1)
            On Error Resume Next
            sentDate = CDate(tblData(i, 1))
            shipStatus = Trim(CStr(tblData(i, 9)))
            On Error GoTo 0
            
            If shipStatus = "出荷完了" Then
                If rngShipDone Is Nothing Then
                    Set rngShipDone = tbl.ListRows(i).Range
                Else
                    Set rngShipDone = Union(rngShipDone, tbl.ListRows(i).Range)
                End If
            ElseIf shipStatus = "欠品中" Then
                If rngStockoutC Is Nothing Then
                    Set rngStockoutC = tbl.ListRows(i).Range
                Else
                    Set rngStockoutC = Union(rngStockoutC, tbl.ListRows(i).Range)
                End If
            ElseIf shipStatus = "分納" Then
                If rngBunnoC Is Nothing Then
                    Set rngBunnoC = tbl.ListRows(i).Range
                Else
                    Set rngBunnoC = Union(rngBunnoC, tbl.ListRows(i).Range)
                End If
            ElseIf sentDate > 0 And (today - sentDate) >= 7 Then
                If rngWeekOld Is Nothing Then
                    Set rngWeekOld = tbl.ListRows(i).Range
                Else
                    Set rngWeekOld = Union(rngWeekOld, tbl.ListRows(i).Range)
                End If
            ElseIf sentDate > 0 And (today - sentDate) >= 3 Then
                If rngThreeDays Is Nothing Then
                    Set rngThreeDays = tbl.ListRows(i).Range
                Else
                    Set rngThreeDays = Union(rngThreeDays, tbl.ListRows(i).Range)
                End If
            Else
                If rngNormal Is Nothing Then
                    Set rngNormal = tbl.ListRows(i).Range
                Else
                    Set rngNormal = Union(rngNormal, tbl.ListRows(i).Range)
                End If
            End If
        Next i
    End If
    
    ' 一括書式設定
    If Not rngShipDone Is Nothing Then
        rngShipDone.Interior.Color = RGB(255, 180, 180)
    End If
    If Not rngStockoutC Is Nothing Then
        rngStockoutC.Interior.Color = RGB(220, 200, 255)
    End If
    If Not rngBunnoC Is Nothing Then
        rngBunnoC.Interior.Color = RGB(200, 220, 255)
    End If
    If Not rngWeekOld Is Nothing Then
        rngWeekOld.Interior.Color = RGB(255, 200, 150)
    End If
    If Not rngThreeDays Is Nothing Then
        rngThreeDays.Interior.Color = RGB(255, 255, 200)
    End If
    If Not rngNormal Is Nothing Then
        rngNormal.Interior.ColorIndex = xlNone
    End If
End Sub

' ============================================
' 【v4.4新機能】確認中一覧のクリーンアップ
' 確定になった伝票を確認中一覧から削除し、送付履歴に移動
' ============================================
Sub CleanConfirmingList(historyWs As Worksheet, confirmingWs As Worksheet, newConfirmedOrders As Collection)
    Dim tblConfirming As ListObject
    Dim tblHistory As ListObject
    Dim i As Long, j As Long
    Dim orderData As Variant
    
    ' 確定した伝票のキーを収集
    Dim confirmedKeys As Object
    Set confirmedKeys = CreateObject("Scripting.Dictionary")
    For Each orderData In newConfirmedOrders
        Dim hKey As String
        hKey = orderData(2) & "|" & orderData(3)
        confirmedKeys(hKey) = orderData
    Next orderData
    
    On Error Resume Next
    Set tblConfirming = confirmingWs.ListObjects("確認中テーブル")
    Set tblHistory = historyWs.ListObjects("送付履歴テーブル")
    On Error GoTo 0
    
    If tblConfirming Is Nothing Or tblHistory Is Nothing Then Exit Sub
    
    ' テーブルのフィルター解除
    On Error Resume Next
    tblConfirming.AutoFilter.ShowAllData
    tblHistory.AutoFilter.ShowAllData
    On Error GoTo 0
    
    If tblConfirming.ListRows.Count = 0 Then Exit Sub
    
    ' === 配列一括読み取り ===
    Dim confData As Variant
    confData = tblConfirming.DataBodyRange.Value
    Dim confRows As Long
    confRows = UBound(confData, 1)
    Dim confCols As Long
    confCols = tblConfirming.ListColumns.Count
    
    ' === 残す行と移動する行を分類 ===
    Dim keepRows As Collection
    Set keepRows = New Collection
    Dim movedOrders As Collection
    Set movedOrders = New Collection
    
    For i = 1 To confRows
        Dim confKey As String
        confKey = Trim(CStr(confData(i, 4))) & "|" & Trim(CStr(confData(i, 5)))
        
        If confirmedKeys.Exists(confKey) Then
            movedOrders.Add confirmedKeys(confKey)
        Else
            keepRows.Add i
        End If
    Next i
    
    ' === 確認中テーブルを残す行だけで再構築 ===
    If keepRows.Count = 0 Then
        ' 全行削除
        tblConfirming.DataBodyRange.Delete
    ElseIf keepRows.Count < confRows Then
        ' 一部削除 → 残す行だけの配列を作って書き戻し
        Dim keepResult() As Variant
        ReDim keepResult(1 To keepRows.Count, 1 To confCols)
        Dim kr As Long
        kr = 0
        Dim ki As Variant
        For Each ki In keepRows
            kr = kr + 1
            For j = 1 To confCols
                keepResult(kr, j) = confData(CLng(ki), j)
            Next j
        Next ki
        
        tblConfirming.DataBodyRange.Delete
        On Error Resume Next
        If tblConfirming.ListRows.Count = 0 Then
            tblConfirming.ListRows.Add
        End If
        On Error GoTo 0
        tblConfirming.Resize tblConfirming.Range.Resize(keepRows.Count + 1, confCols)
        tblConfirming.DataBodyRange.Value = keepResult
    End If
    ' keepRows.Count = confRows なら何も削除しない
    
    ' === 移動した伝票を送付履歴に追加 ===
    If movedOrders.Count > 0 Then
        Call SaveDeliveryHistory(historyWs, movedOrders)
    End If
End Sub

' ============================================
' 古いデータの削除
' ============================================
Sub CleanOldHistory(ws As Worksheet, daysToKeep As Long)
    Dim tbl As ListObject
    Dim cutoffDate As Date
    cutoffDate = Date - daysToKeep
    
    On Error Resume Next
    Set tbl = ws.ListObjects("送付履歴テーブル")
    On Error GoTo 0
    
    If tbl Is Nothing Or tbl.ListRows.Count = 0 Then Exit Sub
    
    On Error Resume Next
    tbl.AutoFilter.ShowAllData
    On Error GoTo 0
    
    ' === 配列一括読み取り ===
    Dim tblData As Variant
    tblData = tbl.DataBodyRange.Value
    Dim totalRows As Long
    totalRows = UBound(tblData, 1)
    Dim colCount As Long
    colCount = tbl.ListColumns.Count
    
    ' === 残す行を収集 ===
    Dim keepRows As Collection
    Set keepRows = New Collection
    Dim i As Long
    
    For i = 1 To totalRows
        Dim sentDate As Date
        sentDate = 0
        On Error Resume Next
        sentDate = CDate(Left(CStr(tblData(i, 1)), 10))
        On Error GoTo 0
        
        If sentDate = 0 Or sentDate >= cutoffDate Then
            keepRows.Add i
        End If
    Next i
    
    ' === 削除が必要な場合のみ書き戻し ===
    If keepRows.Count < totalRows Then
        If keepRows.Count = 0 Then
            tbl.DataBodyRange.Delete
        Else
            Dim result() As Variant
            ReDim result(1 To keepRows.Count, 1 To colCount)
            Dim r As Long, j As Long
            r = 0
            Dim ki As Variant
            For Each ki In keepRows
                r = r + 1
                For j = 1 To colCount
                    result(r, j) = tblData(CLng(ki), j)
                Next j
            Next ki
            
            tbl.DataBodyRange.Delete
            On Error Resume Next
            If tbl.ListRows.Count = 0 Then
                tbl.ListRows.Add
            End If
            On Error GoTo 0
            tbl.Resize tbl.Range.Resize(keepRows.Count + 1, colCount)
            tbl.DataBodyRange.Value = result
        End If
    End If
End Sub

' ============================================
' .xlsファイルを.xlsx形式に変換
' ============================================
Function ConvertXlsToXlsx(xlsFilePath As String) As String
    Dim tempWb As Workbook
    Dim tempFilePath As String
    
    On Error GoTo ErrorHandler
    
    tempFilePath = Environ("TEMP") & "\temp_" & Format(Now, "yyyymmddhhnnss") & ".xlsx"
    Set tempWb = Workbooks.Open(xlsFilePath, ReadOnly:=True)
    tempWb.SaveAs tempFilePath, FileFormat:=xlOpenXMLWorkbook
    tempWb.Close SaveChanges:=False
    
    ConvertXlsToXlsx = tempFilePath
    Exit Function
    
ErrorHandler:
    ConvertXlsToXlsx = ""
    If Not tempWb Is Nothing Then
        tempWb.Close SaveChanges:=False
    End If
End Function

' ============================================
' 列位置を検索する関数
' ============================================
Function GetColumnPositions(ws As Worksheet) As Object
    Dim cols As Object
    Set cols = CreateObject("Scripting.Dictionary")
    
    Dim lastCol As Long
    Dim i As Long
    Dim headerValue As String
    
    lastCol = ws.Cells(5, ws.Columns.count).End(xlToLeft).Column
    
    For i = 2 To lastCol
        headerValue = Trim(ws.Cells(5, i).Value)
        
        If InStr(headerValue, "受発注伝票") > 0 Then
            cols("受発注伝票") = i
        ElseIf InStr(headerValue, "明細") > 0 And Not cols.Exists("明細") Then
            cols("明細") = i
        ElseIf InStr(headerValue, "受注先") > 0 And Not cols.Exists("受注先") Then
            cols("受注先") = i
        ElseIf InStr(headerValue, "テキスト") > 0 And Not cols.Exists("品名") Then
            cols("品名") = i
        ElseIf InStr(headerValue, "受注数量") > 0 Then
            cols("受注数量") = i
        ElseIf InStr(headerValue, "受注単価") > 0 Then
            cols("受注単価") = i
        ElseIf InStr(headerValue, "正味額") > 0 Then
            cols("正味額") = i
        ElseIf InStr(headerValue, "名称") > 0 Then
            cols("メーカー") = i
        ElseIf InStr(headerValue, "保管場所") > 0 Then
            cols("保管場所") = i
        ElseIf InStr(headerValue, "出荷先名") > 0 Then
            cols("出荷先名") = i
        ElseIf InStr(headerValue, "出荷ステータス") > 0 Then
            cols("出荷ステータス") = i
        ElseIf InStr(headerValue, "受注納期") > 0 Then
            cols("受注納期") = i
        ElseIf InStr(headerValue, "品目 Group") > 0 Or InStr(headerValue, "品目Group") > 0 Then
            cols("品目Group") = i
        ElseIf InStr(headerValue, "得意先担当者") > 0 Then
            cols("得意先担当者") = i
        ElseIf InStr(headerValue, "得意先発注番号") > 0 Then
            cols("得意先発注番号") = i
        ElseIf InStr(headerValue, "コメント（明細）") > 0 Then
            cols("コメント（明細）") = i
        ElseIf InStr(headerValue, "コメント（社内）") > 0 Then
            cols("コメント（社内）") = i
        ElseIf InStr(headerValue, "コメント（社外）") > 0 Then
            cols("コメント（社外）") = i
        ElseIf InStr(headerValue, "伝票タイプ") > 0 Then
            cols("伝票タイプ") = i
        ElseIf InStr(headerValue, "時刻") > 0 Then
            cols("時刻") = i
        ElseIf InStr(headerValue, "登録日") > 0 Then
            cols("登録日") = i
        ElseIf InStr(headerValue, "拒否理由") > 0 Then
            cols("拒否理由") = i
        ElseIf InStr(headerValue, "指定納期") > 0 Then
            cols("指定納期") = i
        End If
    Next i
    
    ' 必須列チェック
    If cols.Exists("受発注伝票") And cols.Exists("明細") And _
       cols.Exists("受注先") And cols.Exists("品名") And _
       cols.Exists("受注数量") And cols.Exists("出荷先名") And _
       cols.Exists("受注納期") And cols.Exists("品目Group") And _
       cols.Exists("登録日") Then
        Set GetColumnPositions = cols
    Else
        Set GetColumnPositions = Nothing
    End If
End Function

' ============================================
' 【v6.1】納期回答書を作成する関数（showPrice削除版）
' ============================================
Function CreateDeliveryReport(sourceWs As Worksheet, cols As Object, _
                              customerName As String, _
                              manufacturerMasterWs As Worksheet, _
                              sentOrders As Object, _
                              Optional holidays As Object = Nothing, _
                              Optional dateFrom As Date = 0, _
                              Optional dateTo As Date = 0, _
                              Optional confirmingListWs As Worksheet = Nothing, _
                              Optional customerMasterWs As Worksheet = Nothing, _
                              Optional repName As String = "", _
                              Optional repMasterWs As Worksheet = Nothing) As Variant
    
    Dim newWb As Workbook
    Dim newWs As Worksheet
    Dim currentRow As Long
    Dim i As Long
    Dim lastRow As Long
    Dim confirmedOrders As Collection
    Dim confirmingOrders As Collection
    Dim today As Date
    Dim extComment As String
    Dim trackingInfo As Variant
    Dim shortProductName As String
    Dim qty As Variant
    Dim startTimer As Double
    startTimer = Timer


    today = Date
    
    Set confirmedOrders = New Collection
    Set confirmingOrders = New Collection
    Dim trackingInfoList As Collection
    Set trackingInfoList = New Collection
    Dim stockoutInfoList As Collection
    Set stockoutInfoList = New Collection
    Dim bunnoInfoList As Collection
    Set bunnoInfoList = New Collection
    Dim bunnoCompletedList As Collection
    Set bunnoCompletedList = New Collection
    Dim commentDetail As String
    
    Set newWb = Workbooks.Add
    Set newWs = newWb.Sheets(1)
    
    Dim sheetName As String
    If repName <> "" And repName <> "__OTHER__" Then
        sheetName = customerName & "_" & repName & "様"
    Else
        sheetName = customerName & "様"
    End If
    If Len(sheetName) > 31 Then
        sheetName = Left(sheetName, 31)
    End If
    newWs.Name = sheetName
    
    Call CreateHeader(newWs, customerName, repName)
    
    currentRow = 7
    lastRow = sourceWs.Cells(sourceWs.Rows.count, cols("受注先")).End(xlUp).Row
    
    ' 担当者フィルタ用：マスター登録済みリストをキャッシュ
    Dim registeredRepList As Collection
    Set registeredRepList = Nothing
    If repName = "__OTHER__" And Not repMasterWs Is Nothing Then
        Set registeredRepList = GetRepList(customerName, repMasterWs)
    End If

    
    For i = 7 To lastRow
        If Trim(g_SourceData(i, cols("受注先"))) = customerName Then
            
            ' 担当者フィルタ
            If repName <> "" Then
                Dim repCellValue As String
                repCellValue = ""
                If cols.Exists("得意先担当者") Then
                    repCellValue = Trim(g_SourceData(i, cols("得意先担当者")))
                End If
                If Not ShouldIncludeForRep(repCellValue, repName, registeredRepList) Then GoTo NextRow
            End If
            
            ' 【v4.8追加】期間フィルタ
            Dim registrationDateValue As Variant
            Dim registrationDate As Date
            registrationDateValue = g_SourceData(i, cols("登録日"))
            
            If IsDate(registrationDateValue) Then
                registrationDate = CDate(registrationDateValue)
                
                ' 期間外ならスキップ
                If dateFrom > 0 And dateTo > 0 Then
                    If registrationDate < dateFrom Or registrationDate > dateTo Then
                        GoTo NextRow
                    End If
                End If
            Else
                ' 日付が無効ならスキップ
                GoTo NextRow
            End If
            
            Dim rejectReason As String
            rejectReason = ""
            If cols.Exists("拒否理由") Then
                rejectReason = Trim(g_SourceData(i, cols("拒否理由")))
            End If
            
            Dim documentType As String
            documentType = ""
            If cols.Exists("伝票タイプ") Then
                documentType = Trim(g_SourceData(i, cols("伝票タイプ")))
            End If
            
            Dim orderNumber As String
            Dim detailNumber As String
            Dim historyKey As String
            
            orderNumber = Trim(g_SourceData(i, cols("受発注伝票")))
            detailNumber = Trim(g_SourceData(i, cols("明細")))
            
            ' 送付済みチェックを「注番|明細」で判定
            ' ※受注日（登録日）が今日の場合は何度でも送付OK
            historyKey = orderNumber & "|" & detailNumber
            Dim isAlreadySent As Boolean
            isAlreadySent = False
            
            
            ' 出荷ステータスを取得（処理完了判定用）
            Dim currentShipStatus As String
            currentShipStatus = ""
            If cols.Exists("出荷ステータス") Then
                currentShipStatus = Trim(g_SourceData(i, cols("出荷ステータス")))
            End If
            
            ' 処理完了かつ送付履歴に「納品済み」で記録済み → スキップ
            Dim previousDeliveryStatus As String
            previousDeliveryStatus = ""
            
            If sentOrders.Exists(historyKey) Then
                previousDeliveryStatus = CStr(sentOrders(historyKey))
    
            ' 除外判定：確認中一覧の「除外」かどうかチェック
            Dim isExcludedFromConfirming As Boolean
            isExcludedFromConfirming = False
            
            If previousDeliveryStatus = "除外" Then
                isExcludedFromConfirming = True
            End If
            
            If Not isExcludedFromConfirming Then
                Dim confirmCacheKey As String
                confirmCacheKey = orderNumber & "|" & detailNumber
                
                If Not g_ConfirmCache Is Nothing Then
                    If g_ConfirmCache.Exists(confirmCacheKey) Then
                        If CStr(g_ConfirmCache(confirmCacheKey)(0)) = "除外" Then
                            isExcludedFromConfirming = True
                        End If
                    End If
                Else
                    ' フォールバック：テーブルスキャン
                    If Not confirmingListWs Is Nothing Then
                        Dim tblConfirming As ListObject
                        On Error Resume Next
                        Set tblConfirming = confirmingListWs.ListObjects("確認中テーブル")
                        On Error GoTo 0
                        
                        If Not tblConfirming Is Nothing Then
                            Dim chkIdx As Long
                            For chkIdx = 1 To tblConfirming.ListRows.Count
                                If Trim(tblConfirming.ListRows(chkIdx).Range(1, 4).Value) = orderNumber And _
                                   Trim(tblConfirming.ListRows(chkIdx).Range(1, 5).Value) = detailNumber And _
                                   Trim(tblConfirming.ListRows(chkIdx).Range(1, 8).Value) = "除外" Then
                                    isExcludedFromConfirming = True
                                    Exit For
                                End If
                            Next chkIdx
                        End If
                    End If
                End If
            End If
            
            ' 除外の場合は無条件でスキップ
            If isExcludedFromConfirming Then
                isAlreadySent = True
            ' 分納完了は案内済みなので無条件でスキップ
            ElseIf previousDeliveryStatus = "分納完了" Then
                isAlreadySent = True
            ' 処理完了で前回「確認中」→ 今回「納品済み」で出す
            ElseIf currentShipStatus = "処理完了" And previousDeliveryStatus = "確認中" Then
                isAlreadySent = False
            ' 処理完了でそれ以外（すでに確定納期で送付済み）→ スキップ
            ElseIf currentShipStatus = "処理完了" Then
                isAlreadySent = True
            ' 通常の確定伝票は、当日登録以外はスキップ
            ElseIf Int(registrationDate) < today Then
                isAlreadySent = True
            End If
        End If
        
        ' 処理完了で送付履歴にない → 「納品済み」で出す
        ' ★紐付き（直送販売 + 保管場所≠転送中）は初回に配達予定を出すため除外
        Dim forceDelivered As Boolean
        forceDelivered = False

        Dim isHimozuki As Boolean
        isHimozuki = False
        If documentType = "【受注】直送販売" Then
            Dim storagePlaceForForce As String
            storagePlaceForForce = ""
            If cols.Exists("保管場所") Then
                storagePlaceForForce = Trim(g_SourceData(i, cols("保管場所")))
            End If
            If storagePlaceForForce = "" Then
                storagePlaceForForce = GetStoragePlaceFromSameOrder(sourceWs, cols, _
                    orderNumber, i)
            End If
            If storagePlaceForForce <> "転送中（直送用）" Then
                isHimozuki = True
            End If
        End If

        ' 【修正】確認中一覧のステータスを取得（分納チェック用）
        Dim confirmingStatus As String
        confirmingStatus = ""
        If Not confirmingListWs Is Nothing Then
            confirmingStatus = GetConfirmingStatus(confirmingListWs, orderNumber, detailNumber)
        End If
        
        ' 【修正】分納+処理完了 → 全出荷済みなので分納完了として扱う
        Dim isBunnoInConfirming As Boolean
        isBunnoInConfirming = (confirmingStatus = "分納")
        
        Dim isBunnoCompleted As Boolean
        isBunnoCompleted = False
        If isBunnoInConfirming And currentShipStatus = "処理完了" Then
            isBunnoCompleted = True
            isBunnoInConfirming = False  ' ブロック解除
        End If

        If currentShipStatus = "処理完了" And Not sentOrders.Exists(historyKey) Then
            If Not isHimozuki And Not isBunnoInConfirming Then
                forceDelivered = True
            End If
        End If
        If currentShipStatus = "処理完了" And sentOrders.Exists(historyKey) And previousDeliveryStatus = "確認中" Then
            If Not isHimozuki And Not isBunnoInConfirming Then
                forceDelivered = True
            End If
        End If
            
            ' コメント（社内）に##または＃＃が含まれている場合は除外
            Dim internalComment As String
            Dim isExcluded As Boolean
            isExcluded = False
            If cols.Exists("コメント（社内）") Then
                internalComment = Trim(g_SourceData(i, cols("コメント（社内）")))
                If InStr(internalComment, "##") > 0 Or InStr(internalComment, "＃＃") > 0 Then
                    isExcluded = True
                End If
            End If
            
            ' 条件チェック
            If rejectReason <> "明細削除" And _
               (documentType = "【受注】直送販売" Or documentType = "【受注】在庫販売") And _
               Not isAlreadySent And _
               Not isExcluded Then
                
                ' メーカー名と品名を取得
                Dim itemGroupCode As String
                Dim manufacturerName As String
                Dim productName As String
                Dim fullText As String
                Dim spacePos As Long
                Dim spacePos2 As Long
                
                itemGroupCode = Trim(g_SourceData(i, cols("品目Group")))
                
                ' メーカー名取得
                If itemGroupCode = "Z99" Or itemGroupCode = "Z97" Then
                    fullText = Trim(g_SourceData(i, cols("品名")))
                    
                    ' 半角スペースを検索
                    spacePos = InStr(fullText, " ")
                    ' 見つからない場合は全角スペースを検索
                    If spacePos = 0 Then
                        spacePos = InStr(fullText, "　")
                    End If
                    
                    If spacePos > 0 Then
                        manufacturerName = Left(fullText, spacePos - 1)
                    Else
                        manufacturerName = ""
                    End If
                Else
                    manufacturerName = GetManufacturerName(itemGroupCode, manufacturerMasterWs)
                End If
                
                If manufacturerName = "" And itemGroupCode <> "Z99" And itemGroupCode <> "Z97" Then
                    If cols.Exists("メーカー") Then
                        manufacturerName = Trim(g_SourceData(i, cols("メーカー")))
                    Else
                        manufacturerName = itemGroupCode
                    End If
                End If
                
                ' 品名取得
                productName = Trim(g_SourceData(i, cols("品名")))
                
                If (itemGroupCode = "Z99" Or itemGroupCode = "Z97") And manufacturerName <> "" Then
                    ' 半角スペースを検索
                    spacePos2 = InStr(productName, " ")
                    ' 見つからない場合は全角スペースを検索
                    If spacePos2 = 0 Then
                        spacePos2 = InStr(productName, "　")
                    End If
                    
                    If spacePos2 > 0 Then
                        productName = Trim(Mid(productName, spacePos2 + 1))
                    End If
                End If
                
                Dim deliveryStatus As String
                deliveryStatus = CopyDataRow(sourceWs, newWs, cols, i, currentRow, _
                                           manufacturerMasterWs, holidays, confirmingListWs, forceDelivered, _
                                           customerMasterWs)
                
                ' 処理完了かつ納期が確認中 → 納品済みに上書き
                If forceDelivered And (deliveryStatus = "確認中" Or deliveryStatus = "欠品中" _
    Or InStr(deliveryStatus, "（欠品）") > 0 Or deliveryStatus = "日程調整中") Then
                    deliveryStatus = "納品済み"
                    newWs.Cells(currentRow, 9).Value = "納品済み"
                    ' グレー色を適用
                    With newWs.Cells(currentRow, 9)
                        .Interior.Color = RGB(220, 220, 220)
                        .Font.Color = RGB(80, 80, 80)
                        .Font.Bold = True
                        .HorizontalAlignment = xlCenter
                    End With
                End If
                
                ' 分納完了 → 納品済みに上書き
                If isBunnoCompleted Then
                    deliveryStatus = "納品済み"
                    newWs.Cells(currentRow, 9).Value = "納品済み"
                    With newWs.Cells(currentRow, 9)
                        .Interior.Color = RGB(220, 220, 220)
                        .Font.Color = RGB(80, 80, 80)
                        .Font.Bold = True
                        .HorizontalAlignment = xlCenter
                    End With
                End If
                                           ' 送り状情報を収集（複数対応）
If cols.Exists("コメント（社外）") And productName <> "送料" Then
    extComment = Trim(g_SourceData(i, cols("コメント（社外）")))
    Dim trackingResultsCR As Collection
    Set trackingResultsCR = ExtractTrackingInfo(extComment)
    
    If trackingResultsCR.count > 0 Then
        shortProductName = productName
        If Len(shortProductName) > 25 Then
            shortProductName = Left(shortProductName, 25) & "..."
        End If
        qty = g_SourceData(i, cols("受注数量"))
        
        Dim tInfoCR As Variant
        For Each tInfoCR In trackingResultsCR
            trackingInfoList.Add Array(manufacturerName, shortProductName, qty, tInfoCR(0), tInfoCR(1))
        Next tInfoCR
    End If
End If
                
                
                ' 確定伝票と確認中伝票を分けて記録

                ' 欠品情報を収集（納期情報も含む）
' ※分納がある場合は分納セクションで表示するため除外
If cols.Exists("コメント（明細）") Then
    commentDetail = Trim(g_SourceData(i, cols("コメント（明細）")))
    Dim hasBunnoComment As Boolean
    hasBunnoComment = (InStr(commentDetail, "分納:") > 0 Or InStr(commentDetail, "分納：") > 0)
    
    ' 処理完了＝出荷処理済みなら欠品は解消済みなのでスキップ
    If InStr(commentDetail, "欠品中") > 0 And productName <> "送料" And Not hasBunnoComment _
        And currentShipStatus <> "処理完了" Then
    shortProductName = productName
        If Len(shortProductName) > 25 Then
            shortProductName = Left(shortProductName, 25) & "..."
        End If
        qty = g_SourceData(i, cols("受注数量"))
        
        ' アバウト納期を抽出
        Dim approxDelivery As String
        approxDelivery = ExtractApproxDelivery(commentDetail)
        
        stockoutInfoList.Add Array(manufacturerName, shortProductName, qty, deliveryStatus, approxDelivery)
    End If
End If
                
                ' 【追加】分納情報を収集
                Dim bunnoInfo As Collection
                Set bunnoInfo = ExtractBunnoInfo(commentDetail)
                
                ' 分納完了 → 品名・数量を通知用に収集
                If bunnoInfo.count > 0 And productName <> "送料" _
                    And isBunnoCompleted Then
                    bunnoCompletedList.Add Array(manufacturerName, productName, _
                        g_SourceData(i, cols("受注数量")))
                End If
                
                ' 処理完了＝出荷処理済みなら分納コメントは残骸なのでスキップ
                If bunnoInfo.count > 0 And productName <> "送料" _
                    And currentShipStatus <> "処理完了" Then
                    Dim isShipRuleForBunno As Boolean
                    isShipRuleForBunno = False
                    
                    Dim storagePlaceForBunno As String
                    storagePlaceForBunno = ""
                    If cols.Exists("保管場所") Then
                        storagePlaceForBunno = Trim(g_SourceData(i, cols("保管場所")))
                    End If

                    If storagePlaceForBunno = "転送中（直送用）" Then
                        isShipRuleForBunno = True
                    ElseIf documentType = "【受注】在庫販売" Then
                        If customerName <> Trim(g_SourceData(i, cols("出荷先名"))) Then
                            isShipRuleForBunno = True
                        End If
                            End If
                    
                    Dim isRosenbinForBunno As Boolean
                    isRosenbinForBunno = IsRouteDelivery(customerName, customerMasterWs)
                    
                    ' 在庫販売 + 路線便 → 出荷予定扱い
                    If Not isShipRuleForBunno And isRosenbinForBunno And documentType = "【受注】在庫販売" Then
                        isShipRuleForBunno = True
                    End If
                    
                    ' 【v12.1】分納の計算済み納期を事前に計算して保存
                    Dim daysToAddForBunnoCalc As Long
                    daysToAddForBunnoCalc = GetDeliveryDaysToAdd(itemGroupCode, manufacturerMasterWs)
                    
                    Dim bunnoInfoWithCalc As Collection
                    Set bunnoInfoWithCalc = New Collection
                    
                    Dim bunnoLineForCalc As Variant
                    For Each bunnoLineForCalc In bunnoInfo
                        Dim calcDateForBunno As String
                        calcDateForBunno = CalculateBunnoDate(CStr(bunnoLineForCalc(1)), isShipRuleForBunno, _
                                                              daysToAddForBunnoCalc, holidays, _
                                                              confirmingListWs, orderNumber, detailNumber, _
                                                              isRosenbinForBunno)
                        
                        Dim locationForBunno As String
                        locationForBunno = ""
                        If UBound(bunnoLineForCalc) >= 2 Then
                            locationForBunno = bunnoLineForCalc(2)
                        End If
                        
                        bunnoInfoWithCalc.Add Array(bunnoLineForCalc(0), bunnoLineForCalc(1), locationForBunno, calcDateForBunno)
                    Next bunnoLineForCalc
                    
                    bunnoInfoList.Add Array(manufacturerName, productName, _
                        g_SourceData(i, cols("受注数量")), _
                        bunnoInfoWithCalc, isShipRuleForBunno, itemGroupCode, _
                        orderNumber, detailNumber, _
                        isRosenbinForBunno)
                End If
                
                ' 確定伝票と確認中伝票を分けて記録
                If deliveryStatus <> "" Then
                    ' 【修正】確認中一覧のステータスを取得
                    Dim prevConfirmingStatus As String
                    prevConfirmingStatus = ""
                    If Not confirmingListWs Is Nothing Then
                        prevConfirmingStatus = GetConfirmingStatus(confirmingListWs, orderNumber, detailNumber)
                    End If
                    
                    ' 【修正】確認中一覧に「分納」で登録されている伝票は
                    ' 処理完了でなければ確認中一覧に残す（処理完了なら分納完了として送付履歴へ）
                    Dim keepInConfirming As Boolean
                    keepInConfirming = (prevConfirmingStatus = "分納" And currentShipStatus <> "処理完了")

                                    If deliveryStatus = "確認中" Or deliveryStatus = "欠品中" Or InStr(deliveryStatus, "（欠品）") > 0 Or deliveryStatus = "日程調整中" Or InStr(deliveryStatus, "分納") > 0 Or keepInConfirming Then
                        ' 確認中一覧に追加
                        Dim shipStatusForConfirm As String
                        ' 分納の場合（先にチェック）
                        If InStr(deliveryStatus, "分納") > 0 Or prevConfirmingStatus = "分納" Then
                            ' 未定ありなら確認中一覧へ
                            If HasBunnoMitei(bunnoInfo, confirmingListWs, orderNumber, detailNumber) Then
                                shipStatusForConfirm = "分納"
                            ElseIf prevConfirmingStatus = "分納" Then
                                ' 分納で確認中一覧に残っている → まだ未定があるか再判定
                                If HasBunnoMitei(bunnoInfo, confirmingListWs, orderNumber, detailNumber) Then
                                    shipStatusForConfirm = "分納"
                                Else
                                    ' 全分確定 → 送付履歴へ
                                    confirmedOrders.Add Array(customerName, registrationDate, orderNumber, detailNumber, manufacturerName, productName, "分納完了")
                                    GoTo SkipConfirming
                                End If
                            Else
                                ' 新規で未定なし→送付履歴へ
                                confirmedOrders.Add Array(customerName, registrationDate, orderNumber, detailNumber, manufacturerName, productName, "分納完了")
                                GoTo SkipConfirming
                            End If
                        ' 欠品の場合
                        ElseIf deliveryStatus = "欠品中" Or InStr(deliveryStatus, "（欠品）") > 0 Then
                            shipStatusForConfirm = "欠品中"
                        Else
                            shipStatusForConfirm = ""
                            If cols.Exists("出荷ステータス") Then
                                shipStatusForConfirm = Trim(g_SourceData(i, cols("出荷ステータス")))
                            End If
                        End If
                        confirmingOrders.Add Array(customerName, registrationDate, orderNumber, detailNumber, manufacturerName, productName, shipStatusForConfirm)
SkipConfirming:
                    Else
                        ' 送付履歴に追加（受注日を含む7要素）
                        Dim statusForHistory As String
                        If isBunnoCompleted Then
                            statusForHistory = "分納完了"
                        Else
                            statusForHistory = deliveryStatus
                        End If
                        confirmedOrders.Add Array(customerName, registrationDate, orderNumber, detailNumber, manufacturerName, productName, statusForHistory)
                    End If
                End If
                
                currentRow = currentRow + 1
            End If
        End If
NextRow:
    Next i
    
    If currentRow = 7 Then
        newWb.Close SaveChanges:=False
        CreateDeliveryReport = Empty
        Exit Function
    End If
    
    Call FormatReport(newWs, currentRow - 1, trackingInfoList, stockoutInfoList, bunnoInfoList, manufacturerMasterWs, holidays, confirmingListWs, bunnoCompletedList)
    
    Dim fileName As String
    Dim filePath As String
    Dim savePath As String
    Dim subFolder As String
    
    savePath = ThisWorkbook.Path
    Dim baseFolder As String
    baseFolder = savePath & "\納期回答書"
    
    If Dir(baseFolder, vbDirectory) = "" Then
        MkDir baseFolder
    End If
    
    If g_CurrentSubFolder = "" Then
        Dim todayStr As String
        Dim folderCount As Long
        Dim f As String
        
        todayStr = Format(g_ExecutionTime, "m月d日(aaa)")
        
        ' 同じ日のフォルダを数える
        folderCount = 0
        f = Dir(baseFolder & "\*", vbDirectory)
        Do While f <> ""
            If f <> "." And f <> ".." Then
                If Left(f, Len(todayStr)) = todayStr Then
                    folderCount = folderCount + 1
                End If
            End If
            f = Dir()
        Loop
        
        folderCount = folderCount + 1
        g_CurrentSubFolder = baseFolder & "\" & todayStr & "_" & ToCircledNumber(folderCount) & "回目"
        
        If Dir(g_CurrentSubFolder, vbDirectory) = "" Then
            MkDir g_CurrentSubFolder
        End If
    End If
    
    subFolder = g_CurrentSubFolder
    
    If repName <> "" And repName <> "__OTHER__" Then
        fileName = "納期回答書_" & Replace(customerName, "/", "_") & "様_" & repName & "様_" & Format(Date, "yyyymmdd") & ".xlsx"
    Else
        fileName = "納期回答書_" & Replace(customerName, "/", "_") & "様_" & Format(Date, "yyyymmdd") & ".xlsx"
    End If
    
    filePath = subFolder & "\" & fileName
    
    On Error Resume Next
    newWb.SaveAs filePath
    If Err.Number <> 0 Then
        newWb.Close SaveChanges:=False
        CreateDeliveryReport = Empty
        Exit Function
    End If
    On Error GoTo 0
    
    newWb.Close SaveChanges:=False
    
    ' ファイルパス、確定伝票リスト、確認中伝票リスト、欠品情報、送り状情報、分納完了情報を返す
    CreateDeliveryReport = Array(filePath, confirmedOrders, confirmingOrders, stockoutInfoList, trackingInfoList, bunnoInfoList, bunnoCompletedList)
End Function

' ============================================
' 【v6.1】ヘッダー作成（showPrice削除版）
' ============================================
Sub CreateHeader(ws As Worksheet, customerName As String, Optional repName As String = "")
    ' フォントを游ゴシックに統一
    ws.Cells.Font.Name = "游ゴシック"
    
    ' タイトル行（文字間スペース）
    ws.Range("A1").Value = "納　期　回　答　書"
    ws.Range("A1").Font.Size = 26
    ws.Range("A1").Font.Bold = True
    ws.Range("A1:L1").Merge
    ws.Range("A1").HorizontalAlignment = xlCenter
    ws.Range("A1").VerticalAlignment = xlCenter
    ws.Range("A1").Interior.Color = RGB(20, 40, 70)
    ws.Range("A1").Font.Color = RGB(255, 255, 255)
    ws.Rows("1").RowHeight = 55
    
    ' アクセントライン（2行目）
    ws.Range("A2:L2").Interior.Color = RGB(180, 150, 70)
    ws.Rows("2").RowHeight = 4
    
    ' 3行目：空白
    ws.Rows("3").RowHeight = 10
    
    ' 顧客名（4行目）- ラベル付き
    ws.Range("A4").Value = "お客様："
    ws.Range("A4").Font.Size = 12
    If repName <> "" And repName <> "__OTHER__" Then
        ws.Range("B4").Value = customerName & " 御中（ご担当：" & repName & " 様）"
    Else
        ws.Range("B4").Value = customerName & " 御中"
    End If
    ws.Range("B4").Font.Size = 16
    ws.Range("B4").Font.Bold = True
    ws.Range("B4").Font.Color = RGB(20, 40, 70)
    ws.Range("B4:E4").Merge
    
    ws.Range("L4").Value = "発行日： " & Format(Date, "yyyy年m月d日(aaa)")
    ws.Range("L4").Font.Size = 12
    ws.Range("L4").Font.Color = RGB(50, 50, 50)
    ws.Range("L4").HorizontalAlignment = xlRight
    ws.Rows("4").RowHeight = 35
    
    ' 5行目は空行（余白）
    ws.Rows("5").RowHeight = 8
    
    ' ヘッダー行（6行目）
    ws.Cells(6, 1).Value = "受注日"
    ws.Cells(6, 2).Value = "担当者様"
    ws.Cells(6, 3).Value = "貴社注番"
    ws.Cells(6, 4).Value = "メーカー名"
    ws.Cells(6, 5).Value = "品名"
    ws.Cells(6, 6).Value = "数量"
    ws.Cells(6, 7).Value = "単価"
    ws.Cells(6, 8).Value = "金額"
    ws.Cells(6, 9).Value = "納期回答"
    ws.Cells(6, 10).Value = "納入先名"
    ws.Cells(6, 11).Value = "備考"
    ws.Cells(6, 12).Value = "弊社注番"
    
    With ws.Range("A6:L6")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(35, 55, 90)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Rows("6").RowHeight = 28
End Sub

' ============================================
' 【v6.2】CalculateDeliveryDate関数（修正版）
' - 他支店在庫チェックを曜日制限より先に実行
' - 曜日出荷の締切時間ルールを適用
' 【納期表示ルール】
'
' ■ 特殊パターン（最優先、該当すれば即確定）
'   && → 作業予定/済み
'   @@ → 出荷→着予定
'   引取 → 引取予定/済み
'
' ■ 本流
'   [直送] メーカー→販売店/ユーザー直接
'     → 出荷予定/済み
'
'   [紐付き] メーカー→センター→販売店/ユーザー
'     → 受注納期＋加算日数＋曜日考慮 → 出荷予定/済み
'
'   [在庫販売] センター在庫→配達
'     受注先≠出荷先 → 路線便扱い → 出荷予定/済み
'     他拠点在庫     → 他拠点より出荷予定/済み
'     曜日制限あり   → 次の配送曜日 → 出荷予定/済み
'     曜日制限なし   → 15時前:翌日 / 15時後:翌々日 → 配達予定/済み
'
'   [路線便フラグON] 上記すべて「出荷予定/済み」に統一
'
'   [未確定] 受注納期12/31
'     在庫販売 → 日程調整中
'     直送販売 → 確認中
'
' ★「配達予定」は自社便ルート配達の場合のみ使用
' ★ 路線便は配達日をマツモト側で管理できないため「出荷予定」
' ============================================
Function CalculateDeliveryDate(sourceWs As Worksheet, cols As Object, rowNum As Long, _
                              itemGroupCode As String, manufacturerMasterWs As Worksheet, _
                              Optional holidays As Object = Nothing, _
                              Optional confirmingListWs As Worksheet = Nothing, _
                              Optional customerName As String = "", _
                              Optional customerMasterWs As Worksheet = Nothing) As String
    
    Dim deliveryDate As Date
    Dim shipStatus As String
    Dim storagePlace As String
    Dim adjustedDate As Date
    Dim today As Date
    Dim documentType As String
    Dim timeValue As String
    Dim registrationDate As Date
    Dim timeHour As Integer
    Dim timeMinute As Integer
    Dim daysToAdd As Long
    Dim internalComment As String
    Dim pickupDate As Date
    Dim timeParts As Variant
    Dim rawTime As Variant
    Dim branchSettings As Variant
    Dim cutoffHour As Integer
    Dim customerPlace As String
    Dim deliveryPlace As String
    Dim useShipRule As Boolean
    Dim originalUseShipRule As Boolean
    Dim isRosenbin As Boolean
    
    today = Date
    ' ============================================
    ' 【作業系チェック】コメント（社内）に「&&」があれば作業予定
    ' ============================================
    Dim internalCommentForWork As String
    internalCommentForWork = ""
    If cols.Exists("コメント（社内）") Then
        internalCommentForWork = Trim(g_SourceData(rowNum, cols("コメント（社内）")))
    End If
    
    If InStr(internalCommentForWork, "&&") > 0 Or InStr(internalCommentForWork, "＆＆") > 0 Then
        ' 指定納期または受注納期から日付を取得
        Dim workDate As Date
        workDate = 0
        
        If cols.Exists("指定納期") Then
            On Error Resume Next
            workDate = CDate(g_SourceData(rowNum, cols("指定納期")))
            On Error GoTo 0
            If Month(workDate) = 12 And Day(workDate) = 31 Then workDate = 0
        End If
        
        If workDate = 0 Then
            On Error Resume Next
            workDate = CDate(g_SourceData(rowNum, cols("受注納期")))
            On Error GoTo 0
            If Month(workDate) = 12 And Day(workDate) = 31 Then workDate = 0
        End If
        
        If workDate = 0 Then
            CalculateDeliveryDate = "日程調整中"
        ElseIf workDate <= Date Then
            CalculateDeliveryDate = Format(workDate, "m月d日") & "作業済み"
        Else
            CalculateDeliveryDate = Format(workDate, "m月d日") & "作業予定"
        End If
        Exit Function
    End If
    ' ============================================
    ' 【着日指定チェック】コメント（社内）に「@@○/○」があれば優先
    ' ============================================
    Dim arrivalDateFromComment As Date
    Dim shipDateForArrival As Date
    Dim internalCommentForArrival As String
    
    internalCommentForArrival = ""
    If cols.Exists("コメント（社内）") Then
        internalCommentForArrival = Trim(g_SourceData(rowNum, cols("コメント（社内）")))
    End If
    
    arrivalDateFromComment = ExtractArrivalDateFromInternal(internalCommentForArrival)
    
    If arrivalDateFromComment > 0 Then
        ' 出荷日を取得（指定納期 > 受注納期）
        shipDateForArrival = 0
        
        ' 指定納期をチェック
        If cols.Exists("指定納期") Then
            On Error Resume Next
            Dim specDateForArrival As Date
            specDateForArrival = CDate(g_SourceData(rowNum, cols("指定納期")))
            On Error GoTo 0
            
            ' 12月31日はデフォルト値なので無視
            If specDateForArrival > 0 And Not (Month(specDateForArrival) = 12 And Day(specDateForArrival) = 31) Then
                shipDateForArrival = specDateForArrival
            End If
        End If
        
        ' 指定納期がなければ受注納期
        If shipDateForArrival = 0 Then
            On Error Resume Next
            shipDateForArrival = CDate(g_SourceData(rowNum, cols("受注納期")))
            On Error GoTo 0
            
            ' 12月31日はデフォルト値なので無視
            If Month(shipDateForArrival) = 12 And Day(shipDateForArrival) = 31 Then
                shipDateForArrival = 0
            End If
        End If
        
        ' 出荷日があれば着日表示形式で返す
        If shipDateForArrival > 0 Then
            If shipDateForArrival <= today Then
                CalculateDeliveryDate = Format(shipDateForArrival, "m/d") & "出荷済→" & Format(arrivalDateFromComment, "m/d") & "着"
            Else
                CalculateDeliveryDate = Format(shipDateForArrival, "m/d") & "出荷→" & Format(arrivalDateFromComment, "m/d") & "着予定"
            End If
            Exit Function
        End If
    End If
    
    ' コメント（社外）とコメント（社内）から引き取り判定
    Dim externalComment As String
    Dim internalCommentForPickup As String
    Dim commentForPickup As String
    
    externalComment = ""
    internalCommentForPickup = ""
    
    If cols.Exists("コメント（社外）") Then
        externalComment = Trim(g_SourceData(rowNum, cols("コメント（社外）")))
    End If
    If cols.Exists("コメント（社内）") Then
        internalCommentForPickup = Trim(g_SourceData(rowNum, cols("コメント（社内）")))
    End If
    
    ' 両方のコメントを結合してチェック
    commentForPickup = externalComment & " " & internalCommentForPickup
    pickupDate = ExtractPickupDate(commentForPickup)
    
    If pickupDate > 0 Then
        If pickupDate < today Then
            CalculateDeliveryDate = Format(pickupDate, "m月d日") & "引取済み"
        Else
            CalculateDeliveryDate = Format(pickupDate, "m月d日") & "引取予定"
        End If
        Exit Function
    End If
    
    ' 【追加】在庫販売で受注先 ≠ 出荷先 → 出荷ルール適用
    useShipRule = False
    If cols.Exists("伝票タイプ") Then
        Dim docType As String
        docType = Trim(g_SourceData(rowNum, cols("伝票タイプ")))
        If docType = "【受注】在庫販売" Then
            customerPlace = Trim(g_SourceData(rowNum, cols("受注先")))
            deliveryPlace = Trim(g_SourceData(rowNum, cols("出荷先名")))
            If customerPlace <> deliveryPlace Then
                useShipRule = True
            End If
        End If
    End If
    
    ' 路線便フラグ判定
    isRosenbin = IsRouteDelivery(customerName, customerMasterWs)

    ' 在庫販売 + 路線便 → 出荷予定扱い
    originalUseShipRule = useShipRule
    If Not useShipRule And isRosenbin Then
        useShipRule = True
    End If
    
    ' 【v6.2新機能】指定納期を最優先でチェック（処理完了より前）
    If cols.Exists("指定納期") Then
        Dim specifiedDateEarly As Date
        On Error Resume Next
        specifiedDateEarly = CDate(g_SourceData(rowNum, cols("指定納期")))
        On Error GoTo 0
        
        ' 12月31日はデフォルト値なので無視
        If specifiedDateEarly > 0 And Not (Month(specifiedDateEarly) = 12 And Day(specifiedDateEarly) = 31) Then
            Dim docTypeForSpec As String
            docTypeForSpec = ""
            If cols.Exists("伝票タイプ") Then
                docTypeForSpec = Trim(g_SourceData(rowNum, cols("伝票タイプ")))
            End If
            
            If cols.Exists("保管場所") Then
    storagePlace = Trim(g_SourceData(rowNum, cols("保管場所")))
    
    If storagePlace = "" Then
        storagePlace = GetStoragePlaceFromSameOrder(sourceWs, cols, _
            Trim(g_SourceData(rowNum, cols("受発注伝票"))), rowNum)
    End If
End If
            
            ' 在庫販売：指定納期＝客の希望到着日
            If docTypeForSpec = "【受注】在庫販売" Then
                If storagePlace = "転送中（直送用）" Then
                    ' 転送中 → そのまま出荷予定
                    If specifiedDateEarly <= today Then
                        CalculateDeliveryDate = Format(specifiedDateEarly, "m月d日") & "出荷済み"
                    Else
                        CalculateDeliveryDate = Format(specifiedDateEarly, "m月d日") & "出荷予定"
                    End If
                ElseIf useShipRule Then
                    ' 路線便 or 受注先≠出荷先 → 到着希望日から1営業日逆算して出荷日
                    Dim shipDateFromArrival As Date
                    shipDateFromArrival = GetPreviousBusinessDay(specifiedDateEarly, holidays)
                    If shipDateFromArrival <= today Then
                        CalculateDeliveryDate = Format(shipDateFromArrival, "m月d日") & "出荷済み"
                    Else
                        CalculateDeliveryDate = Format(shipDateFromArrival, "m月d日") & "出荷予定"
                    End If
                Else
                    ' 自社便 → 到着希望日をそのまま配達予定
                    If specifiedDateEarly <= today Then
                        CalculateDeliveryDate = Format(specifiedDateEarly, "m月d日") & "配達済み"
                    Else
                        CalculateDeliveryDate = Format(specifiedDateEarly, "m月d日") & "配達予定"
                    End If
                End If
                Exit Function
            End If
            
            ' 直送販売：+営業日で計算
            daysToAdd = GetDeliveryDaysToAdd(itemGroupCode, manufacturerMasterWs)
            
            If storagePlace = "転送中（直送用）" Then
                ' 本当の直送 → 出荷曜日制限なし
                If specifiedDateEarly <= today Then
                    CalculateDeliveryDate = Format(specifiedDateEarly, "m月d日") & "出荷済み"
                Else
                    CalculateDeliveryDate = Format(specifiedDateEarly, "m月d日") & "出荷予定"
                End If
            Else
                ' 出荷曜日制限をチェック
                Dim deliveryDaysForSpec As Collection
                Set deliveryDaysForSpec = Nothing
                
                If customerName <> "" And Not customerMasterWs Is Nothing Then
                    Set deliveryDaysForSpec = GetCustomerDeliveryDays(customerName, customerMasterWs)
                End If
                
                ' まず+営業日で計算
                adjustedDate = AddBusinessDays(specifiedDateEarly, daysToAdd, holidays)
                
                If Not deliveryDaysForSpec Is Nothing Then
                If deliveryDaysForSpec.count > 0 Then
                    ' 出荷曜日制限あり → 次の出荷曜日を探す
                    adjustedDate = GetNextDeliveryDay(adjustedDate, deliveryDaysForSpec, holidays)
                    
                    If adjustedDate <= today Then
                        CalculateDeliveryDate = Format(adjustedDate, "m月d日") & "出荷済み"
                    Else
                        CalculateDeliveryDate = Format(adjustedDate, "m月d日") & "出荷予定"
                    End If
                Else
                    ' 出荷曜日制限なし
                    If isRosenbin Then
                        Dim rosenbinEarly As Date
                        rosenbinEarly = AddBusinessDays(specifiedDateEarly, WorksheetFunction.Max(daysToAdd - 1, 0), holidays)
                        If rosenbinEarly <= today Then
                            CalculateDeliveryDate = Format(rosenbinEarly, "m月d日") & "出荷済み"
                        Else
                            CalculateDeliveryDate = Format(rosenbinEarly, "m月d日") & "出荷予定"
                        End If
                    Else
                        If adjustedDate <= today Then
                            CalculateDeliveryDate = Format(adjustedDate, "m月d日") & "配達済み"
                        Else
                            CalculateDeliveryDate = Format(adjustedDate, "m月d日") & "配達予定"
                        End If
                    End If
                End If
                Else
                    ' 出荷曜日制限なし
                    If isRosenbin Then
                        Dim rosenbinEarly2 As Date
                        rosenbinEarly2 = AddBusinessDays(specifiedDateEarly, WorksheetFunction.Max(daysToAdd - 1, 0), holidays)
                        If rosenbinEarly2 <= today Then
                            CalculateDeliveryDate = Format(rosenbinEarly2, "m月d日") & "出荷済み"
                        Else
                            CalculateDeliveryDate = Format(rosenbinEarly2, "m月d日") & "出荷予定"
                        End If
                    Else
                        If adjustedDate <= today Then
                            CalculateDeliveryDate = Format(adjustedDate, "m月d日") & "配達済み"
                        Else
                            CalculateDeliveryDate = Format(adjustedDate, "m月d日") & "配達予定"
                        End If
                    End If
                End If
            End If
            Exit Function
        End If
    End If
    
    If cols.Exists("伝票タイプ") And cols.Exists("時刻") And cols.Exists("登録日") Then
        documentType = Trim(g_SourceData(rowNum, cols("伝票タイプ")))
        
        If cols.Exists("出荷ステータス") Then
            shipStatus = Trim(g_SourceData(rowNum, cols("出荷ステータス")))
        Else
            shipStatus = ""
        End If
        
        If documentType = "【受注】在庫販売" And shipStatus = "処理完了" Then
    
            ' 保管場所を先に取得
            If cols.Exists("保管場所") Then
    storagePlace = Trim(g_SourceData(rowNum, cols("保管場所")))
Else
    storagePlace = ""
End If

If storagePlace = "" Then
    storagePlace = GetStoragePlaceFromSameOrder(sourceWs, cols, _
        Trim(g_SourceData(rowNum, cols("受発注伝票"))), rowNum)
End If
            
            rawTime = g_SourceData(rowNum, cols("時刻"))
            If IsNumeric(rawTime) And rawTime > 0 And rawTime < 1 Then
                timeValue = Format(rawTime, "hh:mm:ss")
            Else
                timeValue = Trim(CStr(rawTime))
            End If
            
            If timeValue <> "" Then
                On Error Resume Next
                timeParts = Split(timeValue, ":")
                If UBound(timeParts) >= 1 Then
                    timeHour = CInt(timeParts(0))
                    timeMinute = CInt(timeParts(1))
                End If
                On Error GoTo 0
                
                On Error Resume Next
                registrationDate = CDate(g_SourceData(rowNum, cols("登録日")))
                On Error GoTo 0
                
                If registrationDate > 0 Then
                    branchSettings = GetBranchSettings(holidays, registrationDate)
                    cutoffHour = branchSettings(1)
                    
                    ' 【追加】受注先 ≠ 出荷先 → 出荷ルール
                    If useShipRule Then
                        Dim shipDateForDiff As Date
                        If timeHour < cutoffHour Or (timeHour = cutoffHour And timeMinute = 0) Then
                            shipDateForDiff = registrationDate
                        Else
                            shipDateForDiff = AddBusinessDays(registrationDate, 1, holidays)
                        End If
                        
                        If shipDateForDiff <= today Then
                            CalculateDeliveryDate = Format(shipDateForDiff, "m月d日") & "出荷済み"
                        Else
                            CalculateDeliveryDate = Format(shipDateForDiff, "m月d日") & "出荷予定"
                        End If
                        Exit Function
                    End If
                    
                    ' 受注先 = 出荷先 → 配達ルール（従来通り）
                    If timeHour < cutoffHour Or (timeHour = cutoffHour And timeMinute = 0) Then
                        adjustedDate = AddBusinessDays(registrationDate, 1, holidays)
                    Else
                        adjustedDate = AddBusinessDays(registrationDate, 2, holidays)
                    End If
                    
                    ' ベースの商品センターを判定
                    Dim baseCenter As String
                    baseCenter = g_BaseCenter
                    
                    ' 【先に】他支店在庫かチェック（曜日制限なし）
                    If storagePlace <> baseCenter Then
                        Dim shipDate As Date
                        If timeHour < cutoffHour Or (timeHour = cutoffHour And timeMinute = 0) Then
                            shipDate = registrationDate
                        Else
                            shipDate = AddBusinessDays(registrationDate, 1, holidays)
                        End If
                        
                        If shipDate <= today Then
                            CalculateDeliveryDate = Format(shipDate, "m月d日") & "他拠点より出荷済み"
                        Else
                            CalculateDeliveryDate = Format(shipDate, "m月d日") & "他拠点より出荷予定"
                        End If
                        Exit Function
                    End If
                    
                    ' 自拠点在庫の場合のみ曜日制限をチェック
                    Dim deliveryDaysForComplete As Collection
                    Set deliveryDaysForComplete = Nothing
                    
                    If customerName <> "" And Not customerMasterWs Is Nothing Then
                        Set deliveryDaysForComplete = GetCustomerDeliveryDays(customerName, customerMasterWs)
                    End If
                    
                    If Not deliveryDaysForComplete Is Nothing Then
                    If deliveryDaysForComplete.count > 0 Then
                        ' 出荷曜日制限あり → 締切時間を考慮して起点日を決定
                        Dim baseDate As Date
                        If timeHour < cutoffHour Or (timeHour = cutoffHour And timeMinute = 0) Then
                            baseDate = registrationDate
                        Else
                            baseDate = AddBusinessDays(registrationDate, 1, holidays)
                        End If
                        
                        adjustedDate = GetNextDeliveryDay(baseDate, deliveryDaysForComplete, holidays)
                        
                        If adjustedDate <= today Then
                            CalculateDeliveryDate = Format(adjustedDate, "m月d日") & "出荷済み"
                        Else
                            CalculateDeliveryDate = Format(adjustedDate, "m月d日") & "出荷予定"
                        End If
                        Exit Function
                    End If
                    End If
                    
                    ' 曜日制限なし＋自拠点 → 配達予定
                    If adjustedDate <= today Then
                        CalculateDeliveryDate = Format(adjustedDate, "m月d日") & "配達済み"
                    Else
                        CalculateDeliveryDate = Format(adjustedDate, "m月d日") & "配達予定"
                    End If
                    Exit Function
                    
                End If
            End If
        End If
    End If
    
    ' ===== ここから下は直送販売や処理完了でない場合の処理 =====
    
    If cols.Exists("出荷ステータス") Then
        shipStatus = Trim(g_SourceData(rowNum, cols("出荷ステータス")))
    Else
        shipStatus = ""
    End If
    
    If cols.Exists("保管場所") Then
    storagePlace = Trim(g_SourceData(rowNum, cols("保管場所")))
Else
    storagePlace = ""
End If

If storagePlace = "" Then
    storagePlace = GetStoragePlaceFromSameOrder(sourceWs, cols, _
        Trim(g_SourceData(rowNum, cols("受発注伝票"))), rowNum)
End If

    ' ============================================
    ' 【紐付き対応】直送販売 + 処理完了 + 紐付き → 在庫販売の処理完了と同じ扱い
    ' 紐付きの処理完了 = センターに入荷済み＋出荷処理済み。あとは配達するだけ
    ' 処理完了時刻がSAPから取れないため、マクロ実行時刻(g_ExecutionTime)で代用
    '   12時台実行 → 確実に締切前 → 翌営業日
    '   17時台実行 → 処理完了時刻不明 → 安全側で翌々営業日
    ' ============================================
    If documentType = "【受注】直送販売" And shipStatus = "処理完了" _
       And storagePlace <> "転送中（直送用）" Then

        branchSettings = GetBranchSettings(holidays, Date)
        cutoffHour = branchSettings(1)

        Dim execHour As Integer
        execHour = Hour(g_ExecutionTime)

        If execHour < cutoffHour Then
            adjustedDate = AddBusinessDays(Date, 1, holidays)
        Else
            adjustedDate = AddBusinessDays(Date, 2, holidays)
        End If

        ' 受注先≠出荷先チェック
        Dim useShipRuleHimozuki As Boolean
        useShipRuleHimozuki = False
        customerPlace = Trim(g_SourceData(rowNum, cols("受注先")))
        deliveryPlace = Trim(g_SourceData(rowNum, cols("出荷先名")))
        If customerPlace <> deliveryPlace Then
            useShipRuleHimozuki = True
        End If

        ' 受注先≠出荷先 → 配達日から1営業日逆算して出荷日を回答
        If useShipRuleHimozuki Then
            Dim shipDateHimozuki As Date
            shipDateHimozuki = GetPreviousBusinessDay(adjustedDate, holidays)
            If shipDateHimozuki <= today Then
                CalculateDeliveryDate = Format(shipDateHimozuki, "m月d日") & "出荷済み"
            Else
                CalculateDeliveryDate = Format(shipDateHimozuki, "m月d日") & "出荷予定"
            End If
            Exit Function
        End If

        ' 曜日制限チェック
        Dim deliveryDaysForHimozuki As Collection
        Set deliveryDaysForHimozuki = Nothing
        If customerName <> "" And Not customerMasterWs Is Nothing Then
            Set deliveryDaysForHimozuki = GetCustomerDeliveryDays(customerName, customerMasterWs)
        End If

        Dim hasDeliveryDaysHimozuki As Boolean
        hasDeliveryDaysHimozuki = False
        If Not deliveryDaysForHimozuki Is Nothing Then
            If deliveryDaysForHimozuki.count > 0 Then hasDeliveryDaysHimozuki = True
        End If

        If hasDeliveryDaysHimozuki Then
            adjustedDate = GetNextDeliveryDay(adjustedDate, deliveryDaysForHimozuki, holidays)
            If adjustedDate <= today Then
                CalculateDeliveryDate = Format(adjustedDate, "m月d日") & "出荷済み"
            Else
                CalculateDeliveryDate = Format(adjustedDate, "m月d日") & "出荷予定"
            End If
            Exit Function
        End If

        ' 路線便 → 配達日から1営業日逆算して出荷日を回答
        If isRosenbin Then
            Dim shipDateRosenbinHimozuki As Date
            shipDateRosenbinHimozuki = GetPreviousBusinessDay(adjustedDate, holidays)
            If shipDateRosenbinHimozuki <= today Then
                CalculateDeliveryDate = Format(shipDateRosenbinHimozuki, "m月d日") & "出荷済み"
            Else
                CalculateDeliveryDate = Format(shipDateRosenbinHimozuki, "m月d日") & "出荷予定"
            End If
            Exit Function
        End If

        ' 自社便配達 → 配達予定
        If adjustedDate <= today Then
            CalculateDeliveryDate = Format(adjustedDate, "m月d日") & "配達済み"
        Else
            CalculateDeliveryDate = Format(adjustedDate, "m月d日") & "配達予定"
        End If
        Exit Function
    End If

    On Error Resume Next
    deliveryDate = CDate(g_SourceData(rowNum, cols("受注納期")))
    On Error GoTo 0
    
    If deliveryDate = 0 Then
        CalculateDeliveryDate = "日程調整中"
        Exit Function
    End If
    
    If Month(deliveryDate) = 12 And Day(deliveryDate) = 31 Then
        ' 確認中一覧から確定納期をチェック
        If Not confirmingListWs Is Nothing Then
            Dim confirmedOrderNumber As String
            Dim confirmedDetailNumber As String
            Dim confirmedDate As Date
            
            confirmedOrderNumber = Trim(g_SourceData(rowNum, cols("受発注伝票")))
            confirmedDetailNumber = Trim(g_SourceData(rowNum, cols("明細")))
            
            confirmedDate = GetConfirmedDeliveryDate(confirmingListWs, confirmedOrderNumber, confirmedDetailNumber)
                        
            If confirmedDate > 0 Then
                daysToAdd = GetDeliveryDaysToAdd(itemGroupCode, manufacturerMasterWs)
                
                ' 【修正】紐付き+受注先≠出荷先チェック
                ' originalUseShipRuleは在庫販売専用のため、
                ' 紐付き（直送販売）の受注先≠出荷先は別途判定が必要
                Dim isHimozukiDiffPlace As Boolean
                isHimozukiDiffPlace = False
                If documentType = "【受注】直送販売" And storagePlace <> "転送中（直送用）" Then
                    customerPlace = Trim(g_SourceData(rowNum, cols("受注先")))
                    deliveryPlace = Trim(g_SourceData(rowNum, cols("出荷先名")))
                    If customerPlace <> deliveryPlace Then
                        isHimozukiDiffPlace = True
                    End If
                End If
                
                If storagePlace = "転送中（直送用）" Then
    If confirmedDate <= today Then
        CalculateDeliveryDate = Format(confirmedDate, "m月d日") & "出荷済み"
    Else
        CalculateDeliveryDate = Format(confirmedDate, "m月d日") & "出荷予定"
    End If
ElseIf originalUseShipRule Then
    If confirmedDate <= today Then
        CalculateDeliveryDate = Format(confirmedDate, "m月d日") & "出荷済み"
    Else
        CalculateDeliveryDate = Format(confirmedDate, "m月d日") & "出荷予定"
    End If
                ElseIf isRosenbin Then
                    adjustedDate = AddBusinessDays(confirmedDate, WorksheetFunction.Max(daysToAdd - 1, 0), holidays)
                    If adjustedDate <= today Then
                        CalculateDeliveryDate = Format(adjustedDate, "m月d日") & "出荷済み"
                    Else
                        CalculateDeliveryDate = Format(adjustedDate, "m月d日") & "出荷予定"
                    End If
                Else
                    adjustedDate = AddBusinessDays(confirmedDate, daysToAdd, holidays)
                    ' 曜日制限チェック
                    Dim deliveryDaysForConfirmed As Collection
                    Set deliveryDaysForConfirmed = Nothing
                    If customerName <> "" And Not customerMasterWs Is Nothing Then
                        Set deliveryDaysForConfirmed = GetCustomerDeliveryDays(customerName, customerMasterWs)
                    End If
                    Dim hasDeliveryDaysConfirmed As Boolean
                    hasDeliveryDaysConfirmed = False
                    If Not deliveryDaysForConfirmed Is Nothing Then
                        If deliveryDaysForConfirmed.count > 0 Then hasDeliveryDaysConfirmed = True
                    End If
                    If hasDeliveryDaysConfirmed Then
                        adjustedDate = GetNextDeliveryDay(adjustedDate, deliveryDaysForConfirmed, holidays)
                        If adjustedDate <= today Then
                            CalculateDeliveryDate = Format(adjustedDate, "m月d日") & "出荷済み"
                        Else
                            CalculateDeliveryDate = Format(adjustedDate, "m月d日") & "出荷予定"
                        End If
                    Else
                        ' 紐付き+受注先≠出荷先 → 配達日を計算するが表示は「出荷予定」
                        If isHimozukiDiffPlace Then
                            If adjustedDate <= today Then
                                CalculateDeliveryDate = Format(adjustedDate, "m月d日") & "出荷済み"
                            Else
                                CalculateDeliveryDate = Format(adjustedDate, "m月d日") & "出荷予定"
                            End If
                        Else
                            If adjustedDate <= today Then
                                CalculateDeliveryDate = Format(adjustedDate, "m月d日") & "配達済み"
                            Else
                                CalculateDeliveryDate = Format(adjustedDate, "m月d日") & "配達予定"
                            End If
                        End If
                    End If
                End If
                Exit Function
            End If
        End If
        
        ' 【v6.2変更】在庫販売は「日程調整中」、直送販売は「確認中」
        Dim docTypeForCheck As String
        docTypeForCheck = ""
        If cols.Exists("伝票タイプ") Then
            docTypeForCheck = Trim(g_SourceData(rowNum, cols("伝票タイプ")))
        End If
        
        If docTypeForCheck = "【受注】在庫販売" Then
            CalculateDeliveryDate = "日程調整中"
        Else
            CalculateDeliveryDate = "確認中"
        End If
        Exit Function
    End If
    
    daysToAdd = GetDeliveryDaysToAdd(itemGroupCode, manufacturerMasterWs)
    
    If storagePlace = "転送中（直送用）" Then
        ' 本当の直送 → 出荷曜日制限なし
        If deliveryDate <= today Then
            CalculateDeliveryDate = Format(deliveryDate, "m月d日") & "出荷済み"
        Else
            CalculateDeliveryDate = Format(deliveryDate, "m月d日") & "出荷予定"
        End If
    ElseIf useShipRule Or isRosenbin Then
        ' 紐付き/路線便 → 出荷曜日制限をチェック
        Dim deliveryDaysForShipRule As Collection
        Set deliveryDaysForShipRule = Nothing
        
        If customerName <> "" And Not customerMasterWs Is Nothing Then
            Set deliveryDaysForShipRule = GetCustomerDeliveryDays(customerName, customerMasterWs)
        End If
        
        If Not deliveryDaysForShipRule Is Nothing Then
            If deliveryDaysForShipRule.count > 0 Then
                ' 出荷曜日制限あり
                Dim rosenbinBase As Date
                If isRosenbin And Not originalUseShipRule Then
                    rosenbinBase = AddBusinessDays(deliveryDate, daysToAdd, holidays)
                Else
                    rosenbinBase = deliveryDate
                End If
                adjustedDate = GetNextDeliveryDay(rosenbinBase, deliveryDaysForShipRule, holidays)
                
                If adjustedDate <= today Then
                    CalculateDeliveryDate = Format(adjustedDate, "m月d日") & "出荷済み"
                Else
                    CalculateDeliveryDate = Format(adjustedDate, "m月d日") & "出荷予定"
                End If
            Else
                ' 出荷曜日制限なし
                Dim rosenbinDate As Date
                If isRosenbin And Not originalUseShipRule Then
                    rosenbinDate = AddBusinessDays(deliveryDate, WorksheetFunction.Max(daysToAdd - 1, 0), holidays)
                Else
                    rosenbinDate = deliveryDate
                End If
                If rosenbinDate <= today Then
                    CalculateDeliveryDate = Format(rosenbinDate, "m月d日") & "出荷済み"
                Else
                    CalculateDeliveryDate = Format(rosenbinDate, "m月d日") & "出荷予定"
                End If
            End If
        Else
            ' 出荷曜日制限なし
            Dim rosenbinDate2 As Date
            If isRosenbin And Not originalUseShipRule Then
                rosenbinDate2 = AddBusinessDays(deliveryDate, WorksheetFunction.Max(daysToAdd - 1, 0), holidays)
            Else
                rosenbinDate2 = deliveryDate
            End If
            If rosenbinDate2 <= today Then
                CalculateDeliveryDate = Format(rosenbinDate2, "m月d日") & "出荷済み"
            Else
                CalculateDeliveryDate = Format(rosenbinDate2, "m月d日") & "出荷予定"
            End If
        End If
    Else
        ' 出荷曜日制限があるかチェック
        Dim deliveryDays As Collection
        Set deliveryDays = Nothing
        
        If customerName <> "" And Not customerMasterWs Is Nothing Then
            Set deliveryDays = GetCustomerDeliveryDays(customerName, customerMasterWs)
        End If
        
        Dim hasDeliveryDays As Boolean
        hasDeliveryDays = False
        If Not deliveryDays Is Nothing Then
            If deliveryDays.count > 0 Then hasDeliveryDays = True
        End If
        
        If hasDeliveryDays Then
            ' 出荷曜日制限あり → +営業日で次の出荷曜日を探す → 「出荷予定」
            adjustedDate = AddBusinessDays(deliveryDate, daysToAdd, holidays)
            adjustedDate = GetNextDeliveryDay(adjustedDate, deliveryDays, holidays)
            
            If adjustedDate <= today Then
                CalculateDeliveryDate = Format(adjustedDate, "m月d日") & "出荷済み"
            Else
                CalculateDeliveryDate = Format(adjustedDate, "m月d日") & "出荷予定"
            End If
        Else
            ' 出荷曜日制限なし → 従来通り +営業日 → 「配達予定」
            adjustedDate = AddBusinessDays(deliveryDate, daysToAdd, holidays)
            
            If adjustedDate <= today Then
                CalculateDeliveryDate = Format(adjustedDate, "m月d日") & "配達済み"
            Else
                CalculateDeliveryDate = Format(adjustedDate, "m月d日") & "配達予定"
            End If
        End If
    End If
End Function
Function CopyDataRow(sourceWs As Worksheet, newWs As Worksheet, cols As Object, _
                     sourceRow As Long, targetRow As Long, _
                     manufacturerMasterWs As Worksheet, _
                     Optional holidays As Object = Nothing, _
                     Optional confirmingListWs As Worksheet = Nothing, _
                     Optional forceDelivered As Boolean = False, _
                     Optional customerMasterWs As Worksheet = Nothing) As String
    
    Dim orderNumber As String
    Dim deliveryPlace As String
    Dim deliveryDate As String
    Dim manufacturerName As String
    Dim itemGroupCode As String
    Dim customerContact As String
    Dim customerOrderNumber As String
    Dim commentDetail As String
    Dim registrationDate As Variant
    Dim internalComment As String
    Dim pickupDate As Date
    
    orderNumber = Trim(g_SourceData(sourceRow, cols("受発注伝票")))
    registrationDate = g_SourceData(sourceRow, cols("登録日"))
    
    If cols.Exists("得意先担当者") Then
        customerContact = Trim(g_SourceData(sourceRow, cols("得意先担当者")))
    Else
        customerContact = ""
    End If
    
    If cols.Exists("得意先発注番号") Then
        customerOrderNumber = Trim(g_SourceData(sourceRow, cols("得意先発注番号")))
    Else
        customerOrderNumber = ""
    End If
    
    If cols.Exists("コメント（明細）") Then
        commentDetail = Trim(g_SourceData(sourceRow, cols("コメント（明細）")))
    Else
        commentDetail = ""
    End If
    
    deliveryPlace = Trim(g_SourceData(sourceRow, cols("出荷先名")))
    
    ' お引き取り日付がある場合は「お引き取り」に変更（社外・社内両方チェック）
pickupDate = 0
Dim externalComment As String
Dim internalCommentForPickup As String
Dim commentForPickup As String

externalComment = ""
internalCommentForPickup = ""

If cols.Exists("コメント（社外）") Then
    externalComment = Trim(g_SourceData(sourceRow, cols("コメント（社外）")))
End If
If cols.Exists("コメント（社内）") Then
    internalCommentForPickup = Trim(g_SourceData(sourceRow, cols("コメント（社内）")))
End If

commentForPickup = externalComment & " " & internalCommentForPickup
pickupDate = ExtractPickupDate(commentForPickup)

If pickupDate > 0 Then
    deliveryPlace = "お引き取り"
End If
    ' コメント（社内）を取得（$$判定用）
    internalComment = ""
    If cols.Exists("コメント（社内）") Then
        internalComment = Trim(g_SourceData(sourceRow, cols("コメント（社内）")))
    End If
    
    ' 元のロジック（お引き取りでない場合のみ適用）
    If pickupDate = 0 Then
        If deliveryPlace = Trim(g_SourceData(sourceRow, cols("受注先"))) Then
            deliveryPlace = "貴社"
        ElseIf deliveryPlace <> "" And Right(deliveryPlace, 1) <> "様" Then
            deliveryPlace = deliveryPlace & "様"
        End If
    End If
    
    itemGroupCode = Trim(g_SourceData(sourceRow, cols("品目Group")))
    
    ' 納期計算を先に実行（祝日対応）
    ' 顧客名を取得
    Dim customerNameForCalc As String
    customerNameForCalc = Trim(g_SourceData(sourceRow, cols("受注先")))
    
    deliveryDate = CalculateDeliveryDate(sourceWs, cols, sourceRow, itemGroupCode, _
                                         manufacturerMasterWs, holidays, confirmingListWs, _
                                         customerNameForCalc, customerMasterWs)
    
    ' 【追加】出荷ステータス取得（分納・欠品判定に使用）
    Dim shipStatusInCopy As String
    shipStatusInCopy = ""
    If cols.Exists("出荷ステータス") Then
        shipStatusInCopy = Trim(g_SourceData(sourceRow, cols("出荷ステータス")))
    End If

    ' 【追加】分納判定
    ' 処理完了＝出荷処理済みなら分納コメントは残骸なのでスキップ
    Dim bunnoInfoForCopy As Collection
    Set bunnoInfoForCopy = ExtractBunnoInfo(commentDetail)

    If bunnoInfoForCopy.count > 0 And shipStatusInCopy <> "処理完了" Then
    deliveryDate = "分納"
End If

    ' 【修正】欠品中でも納期があれば「○月○日配達予定（欠品）」と表示
    ' ただし処理完了＝出荷処理済みなら欠品は解消済みなのでスキップ
    ' 確認中一覧で出荷日が手入力済みなら欠品overrideをスキップ（送付履歴に移動させるため）
    Dim hasConfirmedDateForKeppin As Boolean
    hasConfirmedDateForKeppin = False
    If Not confirmingListWs Is Nothing And cols.Exists("明細") Then
        Dim keppinDetailNum As String
        keppinDetailNum = Trim(g_SourceData(sourceRow, cols("明細")))
        If GetConfirmedDeliveryDate(confirmingListWs, orderNumber, keppinDetailNum) > 0 Then
            hasConfirmedDateForKeppin = True
        End If
    End If

    If InStr(commentDetail, "欠品中") > 0 And shipStatusInCopy <> "処理完了" _
       And Not hasConfirmedDateForKeppin Then
        If deliveryDate = "確認中" Or deliveryDate = "日程調整中" Then
            deliveryDate = "欠品中"
        Else
            ' 納期がある場合は（欠品）を追加
            deliveryDate = deliveryDate & "（欠品）"
        End If
    End If
    
    ' Z99/Z97品目のメーカー名抽出（全角・半角スペース両対応）
    If itemGroupCode = "Z99" Or itemGroupCode = "Z97" Then
        Dim fullText As String
        Dim spacePos As Long
        fullText = Trim(g_SourceData(sourceRow, cols("品名")))
        
        spacePos = InStr(fullText, " ")
        If spacePos = 0 Then
            spacePos = InStr(fullText, "　")
        End If
        
        If spacePos > 0 Then
            manufacturerName = Left(fullText, spacePos - 1)
        Else
            manufacturerName = ""
        End If
    Else
        manufacturerName = GetManufacturerName(itemGroupCode, manufacturerMasterWs)
    End If
    
    If manufacturerName = "" And itemGroupCode <> "Z99" And itemGroupCode <> "Z97" Then
        If cols.Exists("メーカー") Then
            manufacturerName = Trim(g_SourceData(sourceRow, cols("メーカー")))
        Else
            manufacturerName = itemGroupCode
        End If
    End If
    
    Dim productName As String
    productName = Trim(g_SourceData(sourceRow, cols("品名")))
    
    ' Z99/Z97品目の品名抽出（全角・半角スペース両対応）
    If (itemGroupCode = "Z99" Or itemGroupCode = "Z97") And manufacturerName <> "" Then
        Dim spacePos2 As Long
        spacePos2 = InStr(productName, " ")
        If spacePos2 = 0 Then
            spacePos2 = InStr(productName, "　")
        End If
        
        If spacePos2 > 0 Then
            productName = Trim(Mid(productName, spacePos2 + 1))
        End If
    End If
    
    Dim unitPrice As Variant
    Dim totalAmount As Variant
    
    unitPrice = g_SourceData(sourceRow, cols("受注単価"))
    totalAmount = g_SourceData(sourceRow, cols("正味額"))
    
    ' 【$$フラグ】コメント（社内）に「$$」があれば価格は確定として表示
    Dim priceConfirmed As Boolean
    priceConfirmed = False
    If InStr(internalComment, "$$") > 0 Or InStr(internalComment, "＄＄") > 0 Then
        priceConfirmed = True
    End If
    
    ' 【修正】納品済み（forceDelivered）の場合は納期理由での確認中表示をスキップ
    If unitPrice = 1 Or (deliveryDate = "確認中" And Not priceConfirmed And Not forceDelivered) Then
        unitPrice = "確認中"
        totalAmount = "確認中"
    End If
    
    newWs.Cells(targetRow, 1).Value = registrationDate
    newWs.Cells(targetRow, 2).Value = customerContact
    newWs.Cells(targetRow, 3).Value = customerOrderNumber
    newWs.Cells(targetRow, 4).Value = manufacturerName
    newWs.Cells(targetRow, 5).Value = productName
    newWs.Cells(targetRow, 6).Value = g_SourceData(sourceRow, cols("受注数量"))
    newWs.Cells(targetRow, 7).Value = unitPrice
    newWs.Cells(targetRow, 8).Value = totalAmount
    newWs.Cells(targetRow, 9).Value = deliveryDate
    newWs.Cells(targetRow, 10).Value = deliveryPlace
    ' 備考欄にセット（「欠品中」と「分納:?」を除外）
Dim remarkText As String
remarkText = RemoveStockoutText(commentDetail)
remarkText = RemoveBunnoText(remarkText)
remarkText = Trim(remarkText)
newWs.Cells(targetRow, 11).Value = remarkText
    newWs.Cells(targetRow, 12).Value = orderNumber
    
    ' 偶数行は薄いブルー
    If targetRow Mod 2 = 0 Then
        newWs.Range("A" & targetRow & ":L" & targetRow).Interior.Color = RGB(225, 235, 248)
    End If
    
    CopyDataRow = deliveryDate
End Function
' ============================================
' 分納テキストを除外
' ============================================
Function RemoveBunnoText(text As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim bunnoText As String
    Dim afterBunno As String
    
    RemoveBunnoText = text
    
    startPos = InStr(text, "分納:")
    If startPos = 0 Then startPos = InStr(text, "分納：")
    
    If startPos = 0 Then Exit Function
    
    ' 分納部分の終了位置を探す
    afterBunno = Mid(text, startPos)
    
    endPos = InStr(afterBunno, "  ")
    If endPos = 0 Then endPos = InStr(afterBunno, vbLf)
    
    If endPos > 0 Then
        ' 分納部分だけ除外
        bunnoText = Left(afterBunno, endPos - 1)
    Else
        ' 全部除外
        bunnoText = afterBunno
    End If
    
    RemoveBunnoText = Replace(text, bunnoText, "")
End Function
' ============================================
' 【v6.1】書式設定
' ============================================
Sub FormatReport(ws As Worksheet, lastDataRow As Long, Optional trackingInfoList As Collection = Nothing, Optional stockoutInfoList As Collection = Nothing, Optional bunnoInfoList As Collection = Nothing, Optional manufacturerMasterWs As Worksheet = Nothing, Optional holidays As Object = Nothing, Optional confirmingListWs As Worksheet = Nothing, Optional bunnoCompletedList As Collection = Nothing)
    Dim dataRange As Range
    Dim deliveryCol As Integer
    
    Dim bunnoItem As Variant
    Dim bunnoDetail As Collection
    Dim bunnoLine As Variant
    Dim bunnoCounter As Long
    Dim bunnoDateCalc As String
    Dim daysToAddForBunno As Long
    
    ws.Columns("A").ColumnWidth = 9
    ws.Columns("B").ColumnWidth = 16
    ws.Columns("C").ColumnWidth = 14
    ws.Columns("D").ColumnWidth = 20
    ws.Columns("E").ColumnWidth = 50
    ws.Columns("F").ColumnWidth = 7
    ws.Columns("G").ColumnWidth = 9
    ws.Columns("H").ColumnWidth = 11
    ws.Columns("I").ColumnWidth = 18
    ws.Columns("J").ColumnWidth = 20
    ws.Columns("K").ColumnWidth = 26
    ws.Columns("L").ColumnWidth = 14
    
    Set dataRange = ws.Range("A6:L" & lastDataRow)
    deliveryCol = 9
    
    ws.Range("I7:I" & lastDataRow).ShrinkToFit = True
    ws.Range("J7:J" & lastDataRow).ShrinkToFit = True
    ws.Range("K7:K" & lastDataRow).ShrinkToFit = True
    
    ' 納期回答列の色分け（配列読み取り + Union一括書式設定）
    Dim j As Long
    Dim cellValue As String
    Dim deliveryDate As Date
    Dim today As Date
    Dim tomorrow As Date
    Dim dayAfterTomorrow As Date
    
    today = Date
    tomorrow = today + 1
    dayAfterTomorrow = today + 2
    
    ' 配列で一括読み取り
    Dim deliveryValues As Variant
    If lastDataRow >= 7 Then
        deliveryValues = ws.Range(ws.Cells(7, deliveryCol), ws.Cells(lastDataRow, deliveryCol)).Value
    End If
    
    ' 共通書式を一括設定
    With ws.Range(ws.Cells(7, deliveryCol), ws.Cells(lastDataRow, deliveryCol))
        .Font.Bold = True
        .Font.Size = 10
        .HorizontalAlignment = xlCenter
    End With
    
    ' カテゴリ別にRangeを収集
    Dim rngDelivered As Range
    Dim rngStockout As Range
    Dim rngStockoutPartial As Range
    Dim rngBunno As Range
    Dim rngConfirming As Range
    Dim rngScheduling As Range
    Dim rngWork As Range
    Dim rngOtherBranch As Range
    Dim rngPickedUp As Range
    Dim rngDone As Range
    Dim rngShipToday As Range
    Dim rngShipLater As Range
    Dim rngDeliverSoon As Range
    Dim rngDeliverLater As Range
    Dim rngPickupPlan As Range
    Dim rngOtherPlan As Range
    Dim rngOtherPlanNoDate As Range
    Dim rngDefault As Range
    Dim targetRow As Long
    
    If IsArray(deliveryValues) Then
        For j = 1 To UBound(deliveryValues, 1)
            cellValue = Trim(CStr(deliveryValues(j, 1)))
            targetRow = j + 6
            
            If cellValue = "納品済み" Then
                If rngDelivered Is Nothing Then
                    Set rngDelivered = ws.Cells(targetRow, deliveryCol)
                Else
                    Set rngDelivered = Union(rngDelivered, ws.Cells(targetRow, deliveryCol))
                End If
            ElseIf cellValue = "欠品中" Then
                If rngStockout Is Nothing Then
                    Set rngStockout = ws.Cells(targetRow, deliveryCol)
                Else
                    Set rngStockout = Union(rngStockout, ws.Cells(targetRow, deliveryCol))
                End If
            ElseIf InStr(cellValue, "（欠品）") > 0 Then
                If rngStockoutPartial Is Nothing Then
                    Set rngStockoutPartial = ws.Cells(targetRow, deliveryCol)
                Else
                    Set rngStockoutPartial = Union(rngStockoutPartial, ws.Cells(targetRow, deliveryCol))
                End If
            ElseIf InStr(cellValue, "分納") > 0 Then
                If rngBunno Is Nothing Then
                    Set rngBunno = ws.Cells(targetRow, deliveryCol)
                Else
                    Set rngBunno = Union(rngBunno, ws.Cells(targetRow, deliveryCol))
                End If
            ElseIf cellValue = "確認中" Then
                If rngConfirming Is Nothing Then
                    Set rngConfirming = ws.Cells(targetRow, deliveryCol)
                Else
                    Set rngConfirming = Union(rngConfirming, ws.Cells(targetRow, deliveryCol))
                End If
            ElseIf cellValue = "日程調整中" Then
                If rngScheduling Is Nothing Then
                    Set rngScheduling = ws.Cells(targetRow, deliveryCol)
                Else
                    Set rngScheduling = Union(rngScheduling, ws.Cells(targetRow, deliveryCol))
                End If
            ElseIf InStr(cellValue, "作業") > 0 Then
                If rngWork Is Nothing Then
                    Set rngWork = ws.Cells(targetRow, deliveryCol)
                Else
                    Set rngWork = Union(rngWork, ws.Cells(targetRow, deliveryCol))
                End If
            ElseIf InStr(cellValue, "他拠点より") > 0 Then
                If rngOtherBranch Is Nothing Then
                    Set rngOtherBranch = ws.Cells(targetRow, deliveryCol)
                Else
                    Set rngOtherBranch = Union(rngOtherBranch, ws.Cells(targetRow, deliveryCol))
                End If
            ElseIf InStr(cellValue, "引取済み") > 0 Then
                If rngPickedUp Is Nothing Then
                    Set rngPickedUp = ws.Cells(targetRow, deliveryCol)
                Else
                    Set rngPickedUp = Union(rngPickedUp, ws.Cells(targetRow, deliveryCol))
                End If
            ElseIf InStr(cellValue, "済み") > 0 Or InStr(cellValue, "済") > 0 Then
                If rngDone Is Nothing Then
                    Set rngDone = ws.Cells(targetRow, deliveryCol)
                Else
                    Set rngDone = Union(rngDone, ws.Cells(targetRow, deliveryCol))
                End If
            ElseIf InStr(cellValue, "予定") > 0 Then
                deliveryDate = ExtractDateFromString(cellValue)
                If deliveryDate > 0 Then
                    If InStr(cellValue, "出荷予定") > 0 Then
                        If deliveryDate = today Or deliveryDate = tomorrow Then
                            If rngShipToday Is Nothing Then
                                Set rngShipToday = ws.Cells(targetRow, deliveryCol)
                            Else
                                Set rngShipToday = Union(rngShipToday, ws.Cells(targetRow, deliveryCol))
                            End If
                        Else
                            If rngShipLater Is Nothing Then
                                Set rngShipLater = ws.Cells(targetRow, deliveryCol)
                            Else
                                Set rngShipLater = Union(rngShipLater, ws.Cells(targetRow, deliveryCol))
                            End If
                        End If
                    ElseIf InStr(cellValue, "配達予定") > 0 Then
                        If deliveryDate = tomorrow Or deliveryDate = dayAfterTomorrow Then
                            If rngDeliverSoon Is Nothing Then
                                Set rngDeliverSoon = ws.Cells(targetRow, deliveryCol)
                            Else
                                Set rngDeliverSoon = Union(rngDeliverSoon, ws.Cells(targetRow, deliveryCol))
                            End If
                        Else
                            If rngDeliverLater Is Nothing Then
                                Set rngDeliverLater = ws.Cells(targetRow, deliveryCol)
                            Else
                                Set rngDeliverLater = Union(rngDeliverLater, ws.Cells(targetRow, deliveryCol))
                            End If
                        End If
                    ElseIf InStr(cellValue, "引取予定") > 0 Then
                        If rngPickupPlan Is Nothing Then
                            Set rngPickupPlan = ws.Cells(targetRow, deliveryCol)
                        Else
                            Set rngPickupPlan = Union(rngPickupPlan, ws.Cells(targetRow, deliveryCol))
                        End If
                    Else
                        If rngOtherPlan Is Nothing Then
                            Set rngOtherPlan = ws.Cells(targetRow, deliveryCol)
                        Else
                            Set rngOtherPlan = Union(rngOtherPlan, ws.Cells(targetRow, deliveryCol))
                        End If
                    End If
                Else
                    If rngOtherPlanNoDate Is Nothing Then
                        Set rngOtherPlanNoDate = ws.Cells(targetRow, deliveryCol)
                    Else
                        Set rngOtherPlanNoDate = Union(rngOtherPlanNoDate, ws.Cells(targetRow, deliveryCol))
                    End If
                End If
            Else
                If rngDefault Is Nothing Then
                    Set rngDefault = ws.Cells(targetRow, deliveryCol)
                Else
                    Set rngDefault = Union(rngDefault, ws.Cells(targetRow, deliveryCol))
                End If
            End If
        Next j
    End If
    
    ' 一括書式設定
    If Not rngDelivered Is Nothing Then
        rngDelivered.Interior.Color = RGB(220, 220, 220)
        rngDelivered.Font.Color = RGB(80, 80, 80)
    End If
    If Not rngStockout Is Nothing Then
        rngStockout.Interior.Color = RGB(220, 80, 20)
        rngStockout.Font.Color = RGB(255, 255, 255)
    End If
    If Not rngStockoutPartial Is Nothing Then
        rngStockoutPartial.Interior.Color = RGB(255, 150, 120)
        rngStockoutPartial.Font.Color = RGB(140, 40, 20)
    End If
    If Not rngBunno Is Nothing Then
        rngBunno.Interior.Color = RGB(200, 220, 255)
        rngBunno.Font.Color = RGB(0, 70, 140)
    End If
    If Not rngConfirming Is Nothing Then
        rngConfirming.Interior.Color = RGB(255, 200, 200)
        rngConfirming.Font.Color = RGB(180, 30, 30)
    End If
    If Not rngScheduling Is Nothing Then
        rngScheduling.Interior.Color = RGB(250, 245, 220)
        rngScheduling.Font.Color = RGB(140, 100, 40)
    End If
    If Not rngWork Is Nothing Then
        rngWork.Interior.Color = RGB(225, 210, 245)
        rngWork.Font.Color = RGB(80, 50, 120)
    End If
    If Not rngOtherBranch Is Nothing Then
        rngOtherBranch.Interior.Color = RGB(255, 220, 180)
        rngOtherBranch.Font.Color = RGB(180, 80, 0)
    End If
    If Not rngPickedUp Is Nothing Then
        rngPickedUp.Interior.Color = RGB(220, 205, 240)
        rngPickedUp.Font.Color = RGB(70, 40, 110)
    End If
    If Not rngDone Is Nothing Then
        rngDone.Interior.Color = RGB(200, 240, 210)
        rngDone.Font.Color = RGB(20, 100, 50)
    End If
    If Not rngShipToday Is Nothing Then
        rngShipToday.Interior.Color = RGB(200, 225, 255)
        rngShipToday.Font.Color = RGB(20, 70, 140)
    End If
    If Not rngShipLater Is Nothing Then
        rngShipLater.Interior.Color = RGB(255, 235, 180)
        rngShipLater.Font.Color = RGB(140, 90, 10)
    End If
    If Not rngDeliverSoon Is Nothing Then
        rngDeliverSoon.Interior.Color = RGB(200, 225, 255)
        rngDeliverSoon.Font.Color = RGB(20, 70, 140)
    End If
    If Not rngDeliverLater Is Nothing Then
        rngDeliverLater.Interior.Color = RGB(255, 235, 180)
        rngDeliverLater.Font.Color = RGB(140, 90, 10)
    End If
    If Not rngPickupPlan Is Nothing Then
        rngPickupPlan.Interior.Color = RGB(230, 215, 250)
        rngPickupPlan.Font.Color = RGB(90, 50, 130)
    End If
    If Not rngOtherPlan Is Nothing Then
        rngOtherPlan.Interior.Color = RGB(255, 235, 180)
        rngOtherPlan.Font.Color = RGB(140, 90, 10)
    End If
    If Not rngOtherPlanNoDate Is Nothing Then
        rngOtherPlanNoDate.Interior.Color = RGB(255, 235, 180)
        rngOtherPlanNoDate.Font.Color = RGB(140, 90, 10)
    End If
    If Not rngDefault Is Nothing Then
        rngDefault.Interior.Color = RGB(250, 245, 220)
        rngDefault.Font.Color = RGB(80, 80, 80)
    End If
    
    ' 単価・金額の「確認中」を赤字・太字にする（配列読み取り + Union一括設定）
    Dim ghValues As Variant
    Dim rngPriceConfirming As Range
    
    If lastDataRow >= 7 Then
        ghValues = ws.Range("G7:H" & lastDataRow).Value
        
        If IsArray(ghValues) Then
            For j = 1 To UBound(ghValues, 1)
                targetRow = j + 6
                If Trim(CStr(ghValues(j, 1))) = "確認中" Then
                    If rngPriceConfirming Is Nothing Then
                        Set rngPriceConfirming = ws.Cells(targetRow, 7)
                    Else
                        Set rngPriceConfirming = Union(rngPriceConfirming, ws.Cells(targetRow, 7))
                    End If
                End If
                If Trim(CStr(ghValues(j, 2))) = "確認中" Then
                    If rngPriceConfirming Is Nothing Then
                        Set rngPriceConfirming = ws.Cells(targetRow, 8)
                    Else
                        Set rngPriceConfirming = Union(rngPriceConfirming, ws.Cells(targetRow, 8))
                    End If
                End If
            Next j
        End If
    End If
    
    If Not rngPriceConfirming Is Nothing Then
        With rngPriceConfirming
            .Font.Color = RGB(180, 30, 30)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .NumberFormat = "@"
        End With
    End If
    
    ' 数値書式
    ws.Range("F:F").NumberFormat = "#,##0"
    ws.Range("A:A").NumberFormat = "m/d(aaa)"
    
    ' G列・H列の数値書式（配列読み取り + Union一括設定）
    Dim numValues As Variant
    Dim rngInteger As Range
    Dim rngDecimal As Range
    
    If lastDataRow >= 7 Then
        numValues = ws.Range("G7:H" & lastDataRow).Value
        
        If IsArray(numValues) Then
            Dim colIdx As Long
            For j = 1 To UBound(numValues, 1)
                targetRow = j + 6
                For colIdx = 1 To 2
                    If IsNumeric(numValues(j, colIdx)) And numValues(j, colIdx) <> "" Then
                        If CDbl(numValues(j, colIdx)) = Int(CDbl(numValues(j, colIdx))) Then
                            If rngInteger Is Nothing Then
                                Set rngInteger = ws.Cells(targetRow, 6 + colIdx)
                            Else
                                Set rngInteger = Union(rngInteger, ws.Cells(targetRow, 6 + colIdx))
                            End If
                        Else
                            If rngDecimal Is Nothing Then
                                Set rngDecimal = ws.Cells(targetRow, 6 + colIdx)
                            Else
                                Set rngDecimal = Union(rngDecimal, ws.Cells(targetRow, 6 + colIdx))
                            End If
                        End If
                    End If
                Next colIdx
            Next j
        End If
    End If
    
    If Not rngInteger Is Nothing Then
        rngInteger.NumberFormat = "#,##0"
    End If
    If Not rngDecimal Is Nothing Then
        rngDecimal.NumberFormat = "#,##0.##"
    End If
    
    ' 罫線
    With dataRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(160, 160, 160)
    End With
    
    With ws.Range("A6:L6").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(180, 150, 70)
    End With
    
    ws.Range("A7:L" & lastDataRow).Font.Size = 10
    ws.Rows("7:" & lastDataRow).RowHeight = 22
    
    ' 税抜き注記
    ws.Range("G" & lastDataRow + 1 & ":H" & lastDataRow + 1).Merge
    With ws.Range("G" & lastDataRow + 1)
        .Value = "※表示金額は税抜きです"
        .Font.Name = "游ゴシック"
        .Font.Size = 9
        .Font.Color = RGB(120, 120, 120)
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(lastDataRow + 1).RowHeight = 16
    
    ' ===== ご連絡事項と署名セクション =====
    Dim infoStartRow As Long
    Dim infoRow As Long
    Dim hasTrackingInfo As Boolean
    Dim hasStockoutInfo As Boolean
    Dim hasBunnoInfo As Boolean
    
    hasTrackingInfo = False
    hasStockoutInfo = False
    hasBunnoInfo = False
    
    If Not trackingInfoList Is Nothing Then
        If trackingInfoList.count > 0 Then hasTrackingInfo = True
    End If
    If Not stockoutInfoList Is Nothing Then
        If stockoutInfoList.count > 0 Then hasStockoutInfo = True
    End If
    If Not bunnoInfoList Is Nothing Then
        If bunnoInfoList.count > 0 Then hasBunnoInfo = True
    End If
    
    infoStartRow = lastDataRow + 2
    infoRow = infoStartRow
    
    ' 【ご連絡事項】ヘッダー
    ws.Range("A" & infoRow & ":G" & infoRow).Merge
    If hasTrackingInfo Or hasStockoutInfo Or hasBunnoInfo Then
        With ws.Range("A" & infoRow)
            .Value = "【ご連絡事項】"
            .Font.Name = "游ゴシック"
            .Font.Bold = True
            .Font.Size = 14
            .Font.Color = RGB(35, 55, 90)
        End With
    End If
    
    With ws.Range("A" & infoRow & ":H" & infoRow).Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .Color = RGB(180, 150, 70)
    End With
    
    ' 署名
    ws.Range("I" & infoRow & ":J" & infoRow).Merge
    With ws.Range("I" & infoRow)
        .Value = "◆ マツモト産業株式会社 ◆"
        .Font.Name = "游ゴシック"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(35, 55, 90)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .Interior.ColorIndex = xlNone
        .Characters(1, 1).Font.Color = RGB(180, 150, 70)
        .Characters(Len(.Value), 1).Font.Color = RGB(180, 150, 70)
    End With
    
    With ws.Range("K" & infoRow & ":L" & infoRow).Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .Color = RGB(180, 150, 70)
    End With
    
    ws.Rows(infoRow).RowHeight = 28
    infoRow = infoRow + 1
    
    ' 営業所名
    Dim branchSettingsForSign As Variant
    branchSettingsForSign = GetBranchSettings()
    
    ws.Range("I" & infoRow & ":J" & infoRow).Merge
    With ws.Range("I" & infoRow)
        .Value = branchSettingsForSign(0)
        .Font.Name = "游ゴシック"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(35, 55, 90)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .Interior.ColorIndex = xlNone
    End With
    
    ' ===== 連絡事項の詳細 =====
    If hasTrackingInfo Or hasStockoutInfo Then
        ' 送り状情報を表示
        If hasTrackingInfo Then
            ws.Range("A" & infoRow & ":G" & infoRow).Merge
            With ws.Range("A" & infoRow)
                .Value = "    下記商品の送り状番号をご連絡いたします。"
                .Font.Name = "游ゴシック"
                .Font.Size = 11
                .Font.Color = RGB(40, 40, 40)
            End With
            ws.Rows(infoRow).RowHeight = 24
            infoRow = infoRow + 1
            
            ' 商品ごとの送り状セットを作成
            Dim productToTrackingSet As Object
            Set productToTrackingSet = CreateObject("Scripting.Dictionary")
            Dim tItemLoop As Variant
            Dim productKeyLoop As String
            
            For Each tItemLoop In trackingInfoList
                productKeyLoop = tItemLoop(0) & "|" & tItemLoop(1) & "|" & tItemLoop(2)
                
                If Not productToTrackingSet.Exists(productKeyLoop) Then
                    productToTrackingSet.Add productKeyLoop, CreateObject("Scripting.Dictionary")
                End If
                
                Dim tKeyLoop As String
                tKeyLoop = tItemLoop(3) & "|" & tItemLoop(4)
                If Not productToTrackingSet(productKeyLoop).Exists(tKeyLoop) Then
                    productToTrackingSet(productKeyLoop).Add tKeyLoop, Array(tItemLoop(3), tItemLoop(4))
                End If
            Next tItemLoop
            
            ' 送り状セット文字列でグループ化
            Dim trackingSetToProducts As Object
            Set trackingSetToProducts = CreateObject("Scripting.Dictionary")
            Dim pKeyLoop As Variant
            Dim trackingSetKey As String
            Dim tKeyVar As Variant
            Dim trackingKeys() As String
            Dim keyCount As Long
            Dim tempKey As String
            Dim sortI As Long, sortJ As Long
            
            For Each pKeyLoop In productToTrackingSet.keys
                ' 送り状キーをソートして一意の文字列を作成
                keyCount = productToTrackingSet(pKeyLoop).count
                ReDim trackingKeys(0 To keyCount - 1)
                sortI = 0
                For Each tKeyVar In productToTrackingSet(pKeyLoop).keys
                    trackingKeys(sortI) = CStr(tKeyVar)
                    sortI = sortI + 1
                Next tKeyVar
                
                ' バブルソート
                For sortI = 0 To keyCount - 2
                    For sortJ = sortI + 1 To keyCount - 1
                        If trackingKeys(sortI) > trackingKeys(sortJ) Then
                            tempKey = trackingKeys(sortI)
                            trackingKeys(sortI) = trackingKeys(sortJ)
                            trackingKeys(sortJ) = tempKey
                        End If
                    Next sortJ
                Next sortI
                
                trackingSetKey = Join(trackingKeys, "||")
                
                If Not trackingSetToProducts.Exists(trackingSetKey) Then
                    trackingSetToProducts.Add trackingSetKey, New Collection
                End If
                trackingSetToProducts(trackingSetKey).Add Array(pKeyLoop, productToTrackingSet(pKeyLoop))
            Next pKeyLoop
            
            ' 送り状セットごとに表示
            Dim setKeyVar As Variant
            Dim productList As Collection
            Dim productData As Variant
            Dim trackingDict As Object
            Dim trackingItemLoop As Variant
            Dim trackingUrlLoop As String
            Dim canDirectLoop As Boolean
            Dim productPartsLoop As Variant
            Dim isMultiTracking As Boolean
            
            For Each setKeyVar In trackingSetToProducts.keys
                Set productList = trackingSetToProducts(setKeyVar)
                
                ' 最初の商品から送り状情報を取得
                Set trackingDict = productList(1)(1)
                isMultiTracking = (trackingDict.count >= 2)
                
                ' 送り状を表示
                For Each tKeyVar In trackingDict.keys
                    trackingItemLoop = trackingDict(tKeyVar)
                    trackingUrlLoop = GetTrackingUrl(CStr(trackingItemLoop(0)), CStr(trackingItemLoop(1)))
                    canDirectLoop = CanDirectTrack(CStr(trackingItemLoop(0)))
                    
                    ws.Range("A" & infoRow & ":G" & infoRow).Merge
                    
                    If canDirectLoop And trackingUrlLoop <> "" Then
                        With ws.Range("A" & infoRow)
                            .Value = "    ■ " & trackingItemLoop(0) & "  " & trackingItemLoop(1)
                            .Font.Name = "游ゴシック"
                            .Font.Size = 11
                            .Font.Bold = True
                            .Font.Color = RGB(0, 70, 180)
                        End With
                        ws.Hyperlinks.Add Anchor:=ws.Range("A" & infoRow), Address:=trackingUrlLoop, _
                            TextToDisplay:=ws.Range("A" & infoRow).Value
                    Else
                        With ws.Range("A" & infoRow)
                            .Value = "    ■ " & trackingItemLoop(0) & "  " & trackingItemLoop(1)
                            .Font.Name = "游ゴシック"
                            .Font.Size = 11
                            .Font.Bold = True
                            .Font.Color = RGB(35, 55, 90)
                        End With
                    End If
                    
                    With ws.Range("A" & infoRow).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .Weight = xlThick
                        .Color = RGB(180, 150, 70)
                    End With
                    ws.Rows(infoRow).RowHeight = 24
                    infoRow = infoRow + 1
                    
                    If Not canDirectLoop And trackingUrlLoop <> "" Then
                        ws.Range("A" & infoRow & ":G" & infoRow).Merge
                        With ws.Range("A" & infoRow)
                            .Value = "        → 追跡ページ（番号を入力してください）"
                            .Font.Name = "游ゴシック"
                            .Font.Size = 10
                            .Font.Color = RGB(0, 70, 180)
                        End With
                        ws.Hyperlinks.Add Anchor:=ws.Range("A" & infoRow), Address:=trackingUrlLoop, _
                            TextToDisplay:=ws.Range("A" & infoRow).Value
                        ws.Rows(infoRow).RowHeight = 20
                        infoRow = infoRow + 1
                    End If
                Next tKeyVar
                
                ' 商品を表示
                For Each productData In productList
                    productPartsLoop = Split(productData(0), "|")
                    ws.Range("A" & infoRow & ":G" & infoRow).Merge
                    With ws.Range("A" & infoRow)
                        .Value = "        - " & productPartsLoop(0) & "  " & productPartsLoop(1) & "  x" & productPartsLoop(2)
                        .Font.Name = "游ゴシック"
                        .Font.Size = 10
                        .Font.Bold = True
                        .Font.Color = RGB(40, 40, 40)
                    End With
                    ws.Rows(infoRow).RowHeight = 20
                    infoRow = infoRow + 1
                Next productData
                
                ' 複数送り状の場合は注釈
                If isMultiTracking Then
                    ws.Range("A" & infoRow & ":G" & infoRow).Merge
                    With ws.Range("A" & infoRow)
                        .Value = "        ※別々の場所からの出荷になります"
                        .Font.Name = "游ゴシック"
                        .Font.Size = 9
                        .Font.Color = RGB(100, 100, 100)
                        .Font.Italic = True
                    End With
                    ws.Rows(infoRow).RowHeight = 18
                    infoRow = infoRow + 1
                End If
                
                infoRow = infoRow + 1
            Next setKeyVar
        End If
        
        ' 欠品情報を表示
        If hasStockoutInfo Then
            ws.Range("A" & infoRow & ":G" & infoRow).Merge
            With ws.Range("A" & infoRow)
                .Value = "    下記商品は現在欠品中です。ご迷惑をおかけし申し訳ございません。"
                .Font.Name = "游ゴシック"
                .Font.Size = 11
                .Font.Bold = True
                .Font.Color = RGB(180, 0, 0)
            End With
            ws.Rows(infoRow).RowHeight = 24
            infoRow = infoRow + 1
            
            Dim stockoutItem As Variant
            Dim stockoutDeliveryEx As String
            Dim stockoutApprox As String
            Dim stockoutDisplayText As String
            
            For Each stockoutItem In stockoutInfoList
                stockoutDeliveryEx = ""
                stockoutApprox = ""
                On Error Resume Next
                stockoutDeliveryEx = stockoutItem(3)
                stockoutApprox = stockoutItem(4)
                On Error GoTo 0
                
                stockoutDisplayText = "        - " & stockoutItem(0) & "  " & stockoutItem(1) & "  x" & stockoutItem(2)
                
                ' 表示優先順位：アバウト納期 > 確定納期 > 入荷次第ご連絡
                If stockoutApprox <> "" Then
                    stockoutDisplayText = stockoutDisplayText & " → " & stockoutApprox
                ElseIf stockoutDeliveryEx = "" Or stockoutDeliveryEx = "欠品中" Or stockoutDeliveryEx = "確認中" Then
                    stockoutDisplayText = stockoutDisplayText & " → 入荷次第ご連絡"
                Else
                    stockoutDeliveryEx = Replace(stockoutDeliveryEx, "（欠品）", "")
                    stockoutDisplayText = stockoutDisplayText & " → " & stockoutDeliveryEx
                End If
                
                ws.Range("A" & infoRow & ":G" & infoRow).Merge
                With ws.Range("A" & infoRow)
                    .Value = stockoutDisplayText
                    .Font.Name = "游ゴシック"
                    .Font.Size = 10
                    .Font.Bold = True
                    .Font.Color = RGB(139, 0, 0)
                End With
                ws.Rows(infoRow).RowHeight = 20
                infoRow = infoRow + 1
            Next stockoutItem
        End If
    End If
    
    ' 分納情報を表示
    If hasBunnoInfo Then
        ' 未定/確認中があるかチェック
        Dim hasBunnoMiteiForNote As Boolean
        hasBunnoMiteiForNote = False
        Dim bunnoItemCheck As Variant
        Dim bunnoDetailCheck As Collection
        
        For Each bunnoItemCheck In bunnoInfoList
            Set bunnoDetailCheck = bunnoItemCheck(3)
            
            If HasBunnoKakuninchu(bunnoDetailCheck) Then
                hasBunnoMiteiForNote = True
                Exit For
            End If
        Next bunnoItemCheck
        
        ws.Range("A" & infoRow & ":G" & infoRow).Merge
        With ws.Range("A" & infoRow)
            .Value = "    下記商品は分納にてお届けいたします。"
            .Font.Name = "游ゴシック"
            .Font.Size = 11
            .Font.Color = RGB(0, 70, 140)
        End With
        ws.Rows(infoRow).RowHeight = 24
        infoRow = infoRow + 1
        
        ' 未定がある場合は注釈を追加
        If hasBunnoMiteiForNote Then
            ws.Range("A" & infoRow & ":G" & infoRow).Merge
            With ws.Range("A" & infoRow)
                .Value = "    ※一部納期未定のためご迷惑をおかけいたします。確定次第ご連絡いたします。"
                .Font.Name = "游ゴシック"
                .Font.Size = 10
                .Font.Color = RGB(180, 0, 0)
            End With
            ws.Rows(infoRow).RowHeight = 20
            infoRow = infoRow + 1
        End If
        
        For Each bunnoItem In bunnoInfoList
            ws.Range("A" & infoRow & ":G" & infoRow).Merge
            With ws.Range("A" & infoRow)
                .Value = "    ■ " & bunnoItem(0) & "  " & bunnoItem(1) & "  x" & bunnoItem(2)
                .Font.Name = "游ゴシック"
                .Font.Size = 11
                .Font.Bold = True
                .Font.Color = RGB(0, 70, 140)
            End With
            
            With ws.Range("A" & infoRow).Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlThick
                .Color = RGB(100, 150, 200)
            End With
            
            ws.Rows(infoRow).RowHeight = 24
            infoRow = infoRow + 1
            
            ' 分納詳細
            Set bunnoDetail = bunnoItem(3)
            daysToAddForBunno = GetDeliveryDaysToAdd(CStr(bunnoItem(5)), manufacturerMasterWs)
            bunnoCounter = 0
            
            ' 同じ日付があるかチェック
            Dim hasSameDate As Boolean
            hasSameDate = CheckSameDateInBunno(bunnoDetail)
            
            ' 【v12.0】注番・明細を取得
            Dim orderNumForBunno As String
            Dim detailNumForBunno As String
            orderNumForBunno = ""
            detailNumForBunno = ""
            On Error Resume Next
            orderNumForBunno = bunnoItem(6)
            detailNumForBunno = bunnoItem(7)
            On Error GoTo 0
            
            For Each bunnoLine In bunnoDetail
                bunnoCounter = bunnoCounter + 1
                
                ' 【v12.1】保存済みの計算済み納期を使用（インデックス3）
                bunnoDateCalc = ""
                If UBound(bunnoLine) >= 3 Then
                    bunnoDateCalc = CStr(bunnoLine(3))
                End If
                
                If bunnoDateCalc = "" Then
                    Dim isRosenbinForBunnoFmt As Boolean
                    isRosenbinForBunnoFmt = False
                    On Error Resume Next
                    isRosenbinForBunnoFmt = CBool(bunnoItem(8))
                    On Error GoTo 0
                    bunnoDateCalc = CalculateBunnoDate(CStr(bunnoLine(1)), CBool(bunnoItem(4)), _
                                                       daysToAddForBunno, holidays, _
                                                       confirmingListWs, orderNumForBunno, detailNumForBunno, _
                                                       isRosenbinForBunnoFmt)
                End If
                                
                ' 場所があれば追加
                Dim locationText As String
                locationText = ""
                If UBound(bunnoLine) >= 2 Then
                    If bunnoLine(2) <> "" Then
                        locationText = "（" & bunnoLine(2) & "）"
                    End If
                End If
                
                ws.Range("A" & infoRow & ":G" & infoRow).Merge
                With ws.Range("A" & infoRow)
                    .Value = "        " & ToCircledNumber(bunnoCounter) & bunnoLine(0) & " → " & bunnoDateCalc & locationText
                    .Font.Name = "游ゴシック"
                    .Font.Size = 10
                    .Font.Color = RGB(40, 40, 40)
                    
                    ' 元データが未定/欠品/確認中だった場合は赤字（確定後も）
                Dim originalDate As String
                originalDate = CStr(bunnoLine(1))
                
                ' 元データが未定/欠品/確認中だったかチェック
                Dim isFromMitei As Boolean
                isFromMitei = (originalDate = "未定" Or InStr(originalDate, "欠品") > 0 Or InStr(originalDate, "確認中") > 0)
                
                If bunnoDateCalc = "確認中" Or (InStr(bunnoDateCalc, "予定") > 0 And InStr(bunnoDateCalc, "出荷") = 0 And InStr(bunnoDateCalc, "配達") = 0) Then
                    ' まだ未確定（確認中、○旬予定）→ 赤
                    .Font.Color = RGB(180, 30, 30)
                    .Font.Bold = True
                ElseIf isFromMitei Then
                    ' 元は未定だったが今は確定 → オレンジ色
                    .Font.Color = RGB(200, 100, 0)
                    .Font.Bold = True
                End If
                End With
                ws.Rows(infoRow).RowHeight = 20
                infoRow = infoRow + 1
            Next bunnoLine
            
            ' 同じ日付がある場合は注釈を追加
            If hasSameDate Then
                ws.Range("A" & infoRow & ":G" & infoRow).Merge
                With ws.Range("A" & infoRow)
                    .Value = "        ※別々の場所からの出荷になります"
                    .Font.Name = "游ゴシック"
                    .Font.Size = 9
                    .Font.Color = RGB(100, 100, 100)
                    .Font.Italic = True
                End With
                ws.Rows(infoRow).RowHeight = 18
                infoRow = infoRow + 1
            End If
            
            infoRow = infoRow + 1
        Next bunnoItem
    End If
    
    ' 分納完了の通知
    Dim hasBunnoCompleted As Boolean
    hasBunnoCompleted = False
    If Not bunnoCompletedList Is Nothing Then
        If bunnoCompletedList.count > 0 Then hasBunnoCompleted = True
    End If
    
    If hasBunnoCompleted Then
        ws.Range("A" & infoRow & ":G" & infoRow).Merge
        With ws.Range("A" & infoRow)
            .Value = "    分納でご注文いただいた商品は全て出荷が完了しました。"
            .Font.Name = "游ゴシック"
            .Font.Size = 11
            .Font.Color = RGB(0, 120, 60)
            .Font.Bold = True
        End With
        ws.Rows(infoRow).RowHeight = 24
        infoRow = infoRow + 1
        
        Dim bcItem As Variant
        For Each bcItem In bunnoCompletedList
            ws.Range("A" & infoRow & ":G" & infoRow).Merge
            With ws.Range("A" & infoRow)
                .Value = "        ■ " & bcItem(0) & "  " & bcItem(1) & "  x" & bcItem(2)
                .Font.Name = "游ゴシック"
                .Font.Size = 10
                .Font.Color = RGB(0, 100, 50)
            End With
            ws.Rows(infoRow).RowHeight = 20
            infoRow = infoRow + 1
        Next bcItem
        
        infoRow = infoRow + 1
    End If
    
    ' ウィンドウ枠の固定
    ws.Activate
    ws.Range("A7").Select
    ActiveWindow.FreezePanes = True
    
    ' 印刷設定
    With ws.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .TopMargin = Application.CentimetersToPoints(1.5)
        .BottomMargin = Application.CentimetersToPoints(1.5)
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0)
        .PrintTitleRows = "$1:$6"
    End With
End Sub
' ============================================
' 分納に同じ日付があるかチェック
' ============================================
Function CheckSameDateInBunno(bunnoDetail As Collection) As Boolean
    Dim dates As Object
    Set dates = CreateObject("Scripting.Dictionary")
    
    Dim item As Variant
    Dim dateStr As String
    
    CheckSameDateInBunno = False
    
    For Each item In bunnoDetail
        dateStr = CStr(item(1))
        If dateStr <> "未定" And InStr(dateStr, "予定") = 0 Then
            If dates.Exists(dateStr) Then
                CheckSameDateInBunno = True
                Exit Function
            Else
                dates.Add dateStr, 1
            End If
        End If
    Next item
End Function

Function GetManufacturerName(itemGroupCode As String, manufacturerMasterWs As Worksheet) As String
    Dim lastRow As Long
    Dim i As Long
    Dim trimmedCode As String
    
    GetManufacturerName = ""
    
    trimmedCode = Trim(itemGroupCode)
    If trimmedCode = "" Then Exit Function
    
    ' キャッシュから検索
    If Not g_MfgNameCache Is Nothing Then
        If g_MfgNameCache.Exists(trimmedCode) Then
            GetManufacturerName = g_MfgNameCache(trimmedCode)
            Exit Function
        End If
    End If
    
    ' フォールバック：シートスキャン
    If manufacturerMasterWs Is Nothing Then Exit Function
    
    lastRow = manufacturerMasterWs.Cells(manufacturerMasterWs.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        If Trim(manufacturerMasterWs.Cells(i, 1).Value) = trimmedCode Then
            GetManufacturerName = Trim(manufacturerMasterWs.Cells(i, 2).Value)
            Exit Function
        End If
    Next i
End Function

Function GetDeliveryDaysToAdd(itemGroupCode As String, manufacturerMasterWs As Worksheet) As Long
    Dim lastRow As Long
    Dim i As Long
    Dim daysValue As Variant
    Dim trimmedCode As String
    
    GetDeliveryDaysToAdd = 2
    
    trimmedCode = Trim(itemGroupCode)
    If trimmedCode = "" Then Exit Function
    
    ' キャッシュから検索
    If Not g_MfgDaysCache Is Nothing Then
        If g_MfgDaysCache.Exists(trimmedCode) Then
            GetDeliveryDaysToAdd = g_MfgDaysCache(trimmedCode)
            Exit Function
        End If
    End If
    
    ' フォールバック：シートスキャン
    If manufacturerMasterWs Is Nothing Then Exit Function
    
    lastRow = manufacturerMasterWs.Cells(manufacturerMasterWs.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        If Trim(manufacturerMasterWs.Cells(i, 1).Value) = trimmedCode Then
            daysValue = manufacturerMasterWs.Cells(i, 3).Value
            If IsNumeric(daysValue) And daysValue <> "" Then
                GetDeliveryDaysToAdd = CLng(daysValue)
            End If
            Exit Function
        End If
    Next i
End Function


' ============================================
' 確認中一覧から確定納期を取得
' ============================================
Function GetConfirmedDeliveryDate(confirmingListWs As Worksheet, _
                                   orderNumber As String, _
                                   detailNumber As String) As Date
    Dim tbl As ListObject
    Dim i As Long
    Dim tblOrderNum As String
    Dim tblDetailNum As String
    Dim confirmedValue As Variant
    Dim cacheKey As String
    
    GetConfirmedDeliveryDate = 0
    
    ' キャッシュから検索
    cacheKey = orderNumber & "|" & detailNumber
    If Not g_ConfirmCache Is Nothing Then
        If g_ConfirmCache.Exists(cacheKey) Then
            confirmedValue = g_ConfirmCache(cacheKey)(2)  ' Array index 2 = col10
            If IsDate(confirmedValue) Then
                GetConfirmedDeliveryDate = CDate(confirmedValue)
            End If
            Exit Function
        End If
    End If
    
    ' フォールバック：テーブルスキャン
    On Error Resume Next
    Set tbl = confirmingListWs.ListObjects("確認中テーブル")
    On Error GoTo 0
    
    If tbl Is Nothing Or tbl.ListRows.Count = 0 Then
        Exit Function
    End If
    
    For i = 1 To tbl.ListRows.Count
        tblOrderNum = Trim(tbl.ListRows(i).Range(1, 4).Value)
        tblDetailNum = Trim(tbl.ListRows(i).Range(1, 5).Value)
        
        If tblOrderNum = orderNumber And tblDetailNum = detailNumber Then
            confirmedValue = tbl.ListRows(i).Range(1, 10).Value
            If IsDate(confirmedValue) Then
                GetConfirmedDeliveryDate = CDate(confirmedValue)
            End If
            Exit Function
        End If
    Next i
End Function
' ============================================
' 【新規】確認中一覧から伝票のステータスを取得
' 戻り値: "分納", "欠品中", "" など（9列目の値）
' ============================================
Function GetConfirmingStatus(confirmingListWs As Worksheet, _
                              orderNumber As String, _
                              detailNumber As String) As String
    Dim tbl As ListObject
    Dim i As Long
    Dim tblOrderNum As String
    Dim tblDetailNum As String
    Dim cacheKey As String
    
    GetConfirmingStatus = ""
    
    If orderNumber = "" Or detailNumber = "" Then Exit Function
    
    ' キャッシュから検索
    cacheKey = orderNumber & "|" & detailNumber
    If Not g_ConfirmCache Is Nothing Then
        If g_ConfirmCache.Exists(cacheKey) Then
            GetConfirmingStatus = CStr(g_ConfirmCache(cacheKey)(1))  ' Array index 1 = col9
            Exit Function
        End If
    End If
    
    ' フォールバック：テーブルスキャン
    If confirmingListWs Is Nothing Then Exit Function
    
    On Error Resume Next
    Set tbl = confirmingListWs.ListObjects("確認中テーブル")
    On Error GoTo 0
    
    If tbl Is Nothing Or tbl.ListRows.Count = 0 Then
        Exit Function
    End If
    
    For i = 1 To tbl.ListRows.Count
        tblOrderNum = Trim(tbl.ListRows(i).Range(1, 4).Value)
        tblDetailNum = Trim(tbl.ListRows(i).Range(1, 5).Value)
        
        If tblOrderNum = orderNumber And tblDetailNum = detailNumber Then
            GetConfirmingStatus = Trim(tbl.ListRows(i).Range(1, 9).Value)
            Exit Function
        End If
    Next i
End Function
' ============================================
Function ExtractPickupDate(comment As String) As Date
    Dim monthNum As Integer
    Dim dayNum As Integer
    Dim i As Long
    Dim currentChar As String
    Dim dateStr As String
    Dim slashPos As Long
    Dim resultDate As Date  '
    
    If InStr(comment, "引取") = 0 And InStr(comment, "引き取り") = 0 Then
        ExtractPickupDate = 0
        Exit Function
    End If
    
    dateStr = ""
    For i = 1 To Len(comment)
        currentChar = Mid(comment, i, 1)
        
        If IsNumeric(currentChar) Or currentChar = "/" Or currentChar = "／" Then
            dateStr = dateStr & currentChar
        Else
            If Len(dateStr) > 0 Then
                slashPos = InStr(dateStr, "/")
                If slashPos = 0 Then slashPos = InStr(dateStr, "／")
                
                If slashPos > 0 Then
                    On Error Resume Next
                    dateStr = Replace(dateStr, "／", "/")
                    monthNum = CInt(Left(dateStr, slashPos - 1))
                    dayNum = CInt(Mid(dateStr, slashPos + 1))
                    
                    If monthNum >= 1 And monthNum <= 12 And dayNum >= 1 And dayNum <= 31 Then
                        resultDate = DateSerial(Year(Date), monthNum, dayNum)
                        If resultDate < Date And (Date - resultDate) > 180 Then
    resultDate = DateSerial(Year(Date) + 1, monthNum, dayNum)
End If
                        ExtractPickupDate = resultDate
                        Exit Function
                    End If
                    On Error GoTo 0
                End If
                dateStr = ""
            End If
        End If
    Next i
    
    If Len(dateStr) > 0 Then
        slashPos = InStr(dateStr, "/")
        If slashPos = 0 Then slashPos = InStr(dateStr, "／")
        
        If slashPos > 0 Then
            On Error Resume Next
            dateStr = Replace(dateStr, "／", "/")
            monthNum = CInt(Left(dateStr, slashPos - 1))
            dayNum = CInt(Mid(dateStr, slashPos + 1))
            
            If monthNum >= 1 And monthNum <= 12 And dayNum >= 1 And dayNum <= 31 Then
                resultDate = DateSerial(Year(Date), monthNum, dayNum)
                If resultDate < Date And (Date - resultDate) > 180 Then
    resultDate = DateSerial(Year(Date) + 1, monthNum, dayNum)
End If
                ExtractPickupDate = resultDate
                Exit Function
            End If
            On Error GoTo 0
        End If
    End If
    
    ExtractPickupDate = 0
End Function
' ============================================
' 【v6.0】営業日を加算（特別締切時間対応）
' ============================================
Function AddBusinessDays(startDate As Date, businessDays As Long, _
                         Optional holidays As Object = Nothing) As Date
    Dim currentDate As Date
    Dim daysAdded As Long
    Dim specialValue As String
    
    currentDate = startDate
    daysAdded = 0
    
    Do While daysAdded < businessDays
        currentDate = currentDate + 1
        
        If Weekday(currentDate) <> 1 And Weekday(currentDate) <> 7 Then
            If holidays Is Nothing Then
                daysAdded = daysAdded + 1
            ElseIf Not holidays.Exists(CLng(currentDate)) Then
                daysAdded = daysAdded + 1
            Else
                specialValue = holidays(CLng(currentDate))
                
                If specialValue = "" Then
                    ' 祝日 → スキップ
                Else
                    ' 特別締切日 → 営業日としてカウント
                    daysAdded = daysAdded + 1
                End If
            End If
        End If
    Loop
    
    AddBusinessDays = currentDate
End Function
' ============================================
' 間隔1日ルール：前日が祝日振替出荷日だった場合スキップ
' ============================================
Function ShouldSkipDueToInterval(checkDate As Date, deliveryDays As Collection, _
                                  Optional holidays As Object = Nothing) As Boolean
    Dim twoDaysAgo As Date
    Dim twoDaysAgoIsDeliveryDay As Boolean
    Dim twoDaysAgoIsHoliday As Boolean
    Dim d As Variant
    
    ShouldSkipDueToInterval = False
    twoDaysAgo = checkDate - 2
    
    If Weekday(twoDaysAgo) = 1 Or Weekday(twoDaysAgo) = 7 Then
        Exit Function
    End If
    
    twoDaysAgoIsDeliveryDay = False
    For Each d In deliveryDays
        If CInt(d) = Weekday(twoDaysAgo) Then
            twoDaysAgoIsDeliveryDay = True
            Exit For
        End If
    Next d
    
    If Not twoDaysAgoIsDeliveryDay Then
        Exit Function
    End If
    
    twoDaysAgoIsHoliday = False
    If Not holidays Is Nothing Then
        If holidays.Exists(CLng(twoDaysAgo)) Then
            If holidays(CLng(twoDaysAgo)) = "" Then
                twoDaysAgoIsHoliday = True
            End If
        End If
    End If
    
    If twoDaysAgoIsHoliday Then
        ShouldSkipDueToInterval = True
    End If
End Function

' ============================================
' 次の配送可能日を計算（振替出荷日対応版）
' ============================================
Function GetNextDeliveryDay(baseDate As Date, deliveryDays As Collection, _
                            Optional holidays As Object = Nothing) As Date
    
    Dim checkDate As Date
    Dim maxLoop As Long
    Dim loopCount As Long
    Dim isDeliveryDay As Boolean
    Dim isShifted As Boolean  ' ← 変数名を変更
    Dim i As Long
    
    checkDate = baseDate
    maxLoop = 30
    loopCount = 0
    
    Do While loopCount < maxLoop
        loopCount = loopCount + 1
        
        ' 土日はスキップ
        If Weekday(checkDate) = 1 Or Weekday(checkDate) = 7 Then
            checkDate = checkDate + 1
            GoTo ContinueLoop
        End If
        
        ' 出荷曜日かチェック
        isDeliveryDay = False
        For i = 1 To deliveryDays.count
            If CInt(deliveryDays(i)) = Weekday(checkDate) Then
                isDeliveryDay = True
                Exit For
            End If
        Next i
        
        ' 振替出荷日かチェック
        isShifted = IsShiftedDeliveryDay(checkDate, deliveryDays, holidays)  ' ← 変数名を変更
        
        If isDeliveryDay Then
            ' 通常の出荷曜日
            If IsHoliday(checkDate, holidays) Then
                GetNextDeliveryDay = GetNextBusinessDay(checkDate, holidays)
                Exit Function
            End If
            
            If CheckIntervalRule(checkDate, deliveryDays, holidays) Then
                GetNextDeliveryDay = GetNextBusinessDay(checkDate, holidays)
                Exit Function
            End If
            
            GetNextDeliveryDay = checkDate
            Exit Function
            
        ElseIf isShifted Then  ' ← 変数名を変更
            ' 振替出荷日 → そのまま出荷OK
            ' ※注意：振替出荷日が祝日（GW連休等）の場合、祝日を出荷日として返す可能性あり
            ' 　現状は年1-2回のため手動対応。将来修正する場合はここにIsHolidayチェックを追加
            GetNextDeliveryDay = checkDate
            Exit Function
        End If
        
        checkDate = checkDate + 1
ContinueLoop:
    Loop
    
    GetNextDeliveryDay = baseDate
End Function


' ============================================
' 振替出荷日かどうか判定
' ============================================
Function IsShiftedDeliveryDay(checkDate As Date, deliveryDays As Collection, _
                              Optional holidays As Object = Nothing) As Boolean
    Dim yesterday As Date
    Dim yesterdayIsDeliveryDay As Boolean
    Dim i As Long
    
    IsShiftedDeliveryDay = False
    yesterday = checkDate - 1
    
    ' 前日が出荷曜日かチェック
    yesterdayIsDeliveryDay = False
    For i = 1 To deliveryDays.count
        If CInt(deliveryDays(i)) = Weekday(yesterday) Then
            yesterdayIsDeliveryDay = True
            Exit For
        End If
    Next i
    
    If Not yesterdayIsDeliveryDay Then Exit Function
    
    ' 前日が出荷曜日で、間隔1日ルールに該当する場合 → 今日は振替出荷日
    If CheckIntervalRule(yesterday, deliveryDays, holidays) Then
        IsShiftedDeliveryDay = True
        Exit Function
    End If
    
    ' 前日が出荷曜日で祝日の場合 → 今日は振替出荷日
    If IsHoliday(yesterday, holidays) Then
        IsShiftedDeliveryDay = True
    End If
End Function

' ============================================
' 間隔1日ルールチェック
' ============================================
Function CheckIntervalRule(checkDate As Date, deliveryDays As Collection, _
                           Optional holidays As Object = Nothing) As Boolean
    Dim twoDaysAgo As Date
    Dim twoDaysAgoIsDeliveryDay As Boolean
    Dim i As Long
    
    CheckIntervalRule = False
    twoDaysAgo = checkDate - 2
    
    ' 土日ならルール適用外
    If Weekday(twoDaysAgo) = 1 Or Weekday(twoDaysAgo) = 7 Then
        Exit Function
    End If
    
    ' 2日前が出荷曜日かチェック
    twoDaysAgoIsDeliveryDay = False
    For i = 1 To deliveryDays.count
        If CInt(deliveryDays(i)) = Weekday(twoDaysAgo) Then
            twoDaysAgoIsDeliveryDay = True
            Exit For
        End If
    Next i
    
    If Not twoDaysAgoIsDeliveryDay Then Exit Function
    
    ' 2日前が祝日だったら、振替出荷があったはず → ルール適用
    If IsHoliday(twoDaysAgo, holidays) Then
        CheckIntervalRule = True
    End If
End Function
' ============================================
' 祝日かどうか判定
' ============================================
Function IsHoliday(checkDate As Date, Optional holidays As Object = Nothing) As Boolean
    IsHoliday = False
    
    If holidays Is Nothing Then Exit Function
    
    If holidays.Exists(CLng(checkDate)) Then
        ' 値が空文字なら祝日（特別締切時間があれば営業日）
        If holidays(CLng(checkDate)) = "" Then
            IsHoliday = True
        End If
    End If
End Function

' ============================================
' 次の営業日を取得
' ============================================
Function GetNextBusinessDay(startDate As Date, Optional holidays As Object = Nothing) As Date
    Dim checkDate As Date
    checkDate = startDate + 1
    
    Do While True
        ' 土日をスキップ
        If Weekday(checkDate) = 1 Or Weekday(checkDate) = 7 Then
            checkDate = checkDate + 1
        ' 祝日をスキップ
        ElseIf IsHoliday(checkDate, holidays) Then
            checkDate = checkDate + 1
        Else
            Exit Do
        End If
    Loop
    
    GetNextBusinessDay = checkDate
End Function
' ============================================
' 前の営業日を取得（土日・祝日をスキップして1日戻る）
' ============================================
Function GetPreviousBusinessDay(targetDate As Date, Optional holidays As Object = Nothing) As Date
    Dim checkDate As Date
    checkDate = targetDate - 1

    Do While True
        If Weekday(checkDate) = 1 Or Weekday(checkDate) = 7 Then
            checkDate = checkDate - 1
        ElseIf IsHoliday(checkDate, holidays) Then
            checkDate = checkDate - 1
        Else
            Exit Do
        End If
    Loop

    GetPreviousBusinessDay = checkDate
End Function
Function ExtractDateFromString(text As String) As Date
    Dim monthPos As Long
    Dim dayPos As Long
    Dim monthStr As String
    Dim dayStr As String
    Dim monthNum As Integer
    Dim dayNum As Integer
    Dim currentYear As Integer
    
    monthPos = InStr(text, "月")
    dayPos = InStr(text, "日")
    
    If monthPos > 0 And dayPos > 0 Then
        monthStr = ""
        Dim i As Integer
        For i = monthPos - 1 To 1 Step -1
            If IsNumeric(Mid(text, i, 1)) Then
                monthStr = Mid(text, i, 1) & monthStr
            Else
                Exit For
            End If
        Next i
        
        dayStr = ""
        For i = monthPos + 1 To dayPos - 1
            If IsNumeric(Mid(text, i, 1)) Then
                dayStr = dayStr & Mid(text, i, 1)
            End If
        Next i
        
        On Error Resume Next
        monthNum = CInt(monthStr)
        dayNum = CInt(dayStr)
        On Error GoTo 0
        
        If monthNum > 0 And monthNum <= 12 And dayNum > 0 And dayNum <= 31 Then
    Dim resultDate As Date
    resultDate = DateSerial(Year(Date), monthNum, dayNum)
    If resultDate < Date And (Date - resultDate) > 180 Then
    resultDate = DateSerial(Year(Date) + 1, monthNum, dayNum)
End If
    ExtractDateFromString = resultDate
Else
    ExtractDateFromString = 0
End If
    Else
        ExtractDateFromString = 0
    End If
End Function

Sub CreateEmails(createdFiles As Collection, customerMasterWs As Worksheet, _
                 Optional manufacturerMasterWs As Worksheet = Nothing, _
                 Optional holidays As Object = Nothing, _
                 Optional confirmingListWs As Worksheet = Nothing, _
                 Optional sendDirectly As Boolean = False, _
                 Optional repMasterWs As Worksheet = Nothing)
    Dim OutApp As Object
    Dim i As Long
    Dim customerName As String
    Dim filePath As String
    Dim mailAddresses As String
    Dim mailSubject As String
    Dim mailBody As String
    Dim skippedCustomers As String
    Dim sharedAddress As String
    Dim stockoutInfoList As Collection
    Dim trackingInfoList As Collection
    
    skippedCustomers = ""
    
    ' 共通アドレスを設定
    sharedAddress = g_SharedEmail
    
    On Error Resume Next
    Set OutApp = GetObject(, "Outlook.Application")
    If OutApp Is Nothing Then
        Set OutApp = CreateObject("Outlook.Application")
    End If
    
    If OutApp Is Nothing Then
        Exit Sub
    End If
    On Error GoTo 0
    
    For i = 1 To createdFiles.count
        customerName = createdFiles(i)(0)
        filePath = createdFiles(i)(1)
        
        ' 欠品情報と送り状情報を取得
        Set stockoutInfoList = Nothing
        Set trackingInfoList = Nothing
        
        On Error Resume Next
        If UBound(createdFiles(i)) >= 2 Then
            Set stockoutInfoList = createdFiles(i)(2)
        End If
        If UBound(createdFiles(i)) >= 3 Then
            Set trackingInfoList = createdFiles(i)(3)
        End If
        On Error GoTo 0
        
        ' 担当者名を取得（createdFilesの6番目の要素）
        Dim repNameForMail As String
        repNameForMail = ""
        On Error Resume Next
        If UBound(createdFiles(i)) >= 5 Then
            repNameForMail = CStr(createdFiles(i)(5))
        End If
        On Error GoTo 0
        
        ' メールアドレス取得（担当者別 or 従来）
        If repNameForMail <> "" And Not repMasterWs Is Nothing Then
            mailAddresses = GetRepEmailAddresses(customerName, repNameForMail, repMasterWs)
        Else
            mailAddresses = GetEmailAddresses(customerName, customerMasterWs)
        End If
        
        If mailAddresses = "" Then
            skippedCustomers = skippedCustomers & "・" & customerName & vbCrLf
            GoTo NextCustomer
        End If
        
        Dim OutMail As Object
        Set OutMail = OutApp.CreateItem(0)
        
        If repNameForMail <> "" Then
            mailSubject = "【マツモト産業】納期回答書_" & Format(Date, "mm/dd") & "受注分_" & customerName & "様（" & repNameForMail & "様担当分）"
        Else
            mailSubject = "【マツモト産業】納期回答書_" & Format(Date, "mm/dd") & "受注分_" & customerName & "様"
        End If
        Dim branchSettings As Variant
        branchSettings = GetBranchSettings()
        
        Dim bunnoInfoListForMail As Collection
        Set bunnoInfoListForMail = Nothing
        
        On Error Resume Next
        If UBound(createdFiles(i)) >= 4 Then
            Set bunnoInfoListForMail = createdFiles(i)(4)
        End If
        On Error GoTo 0
        
        Dim bunnoCompletedListForMail As Collection
        Set bunnoCompletedListForMail = Nothing
        
        On Error Resume Next
        If UBound(createdFiles(i)) >= 6 Then
            Set bunnoCompletedListForMail = createdFiles(i)(6)
        End If
        On Error GoTo 0
        
        mailBody = BuildEmailBodyHTML(customerName, branchSettings, stockoutInfoList, _
                              trackingInfoList, bunnoInfoListForMail, manufacturerMasterWs, holidays, confirmingListWs, _
                              bunnoCompletedListForMail)
        
        With OutMail
            .To = mailAddresses
            .Subject = mailSubject
            .HTMLBody = mailBody
            .Attachments.Add filePath
            
            On Error Resume Next
            .SentOnBehalfOfName = sharedAddress
            If Err.Number <> 0 Then
                Err.Clear
                .SentOnBehalfOfName = ""
            End If
            On Error GoTo 0
            
            If sendDirectly Then
                .Send
            Else
                .Display
            End If
        End With
        
NextCustomer:
    Next i
    
    Dim finalMessage As String
    If sendDirectly Then
        finalMessage = "メールを送信しました。"
    Else
        finalMessage = "メールを作成しました。" & vbCrLf & _
                       "内容を確認して送信してください。"
    End If
    
    If skippedCustomers <> "" Then
        finalMessage = finalMessage & vbCrLf & vbCrLf & _
                       "※以下の顧客はメールアドレス未登録のためスキップしました：" & vbCrLf & _
                       skippedCustomers & vbCrLf & _
                       "手動でメール作成してください。"
    End If
    
    MsgBox finalMessage, vbInformation, "メール作成完了"
    
End Sub

Function GetEmailAddresses(customerName As String, customerMasterWs As Worksheet) As String
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim emailList As String
    Dim emailAddress As String
    
    lastRow = customerMasterWs.Cells(customerMasterWs.Rows.count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        If Trim(customerMasterWs.Cells(i, 1).Value) = customerName Then
            emailList = ""
            ' ★ E列（5列目）から開始に変更
                For j = 5 To customerMasterWs.Cells(i, customerMasterWs.Columns.count).End(xlToLeft).Column
                emailAddress = Trim(customerMasterWs.Cells(i, j).Value)
                If emailAddress <> "" Then
                    If emailList = "" Then
                        emailList = emailAddress
                    Else
                        emailList = emailList & "; " & emailAddress
                    End If
                End If
            Next j
            
            GetEmailAddresses = emailList
            Exit Function
        End If
    Next i
    
    GetEmailAddresses = ""
End Function
' ============================================
' 顧客の配送曜日を取得（B列から）
' 戻り値：配送可能な曜日のCollection、未設定ならNothing
' ============================================
Function GetCustomerDeliveryDays(customerName As String, customerMasterWs As Worksheet) As Collection
    Dim lastRow As Long
    Dim i As Long
    Dim deliveryDaysValue As String
    Dim result As Collection
    Dim dayParts As Variant
    Dim j As Long
    Dim dayNum As Integer
    
    ' キャッシュから検索
    If Not g_CustDaysCache Is Nothing Then
        If g_CustDaysCache.Exists(customerName) Then
            Dim cachedDaysValue As String
            cachedDaysValue = g_CustDaysCache(customerName)
            
            If cachedDaysValue <> "" Then
                Set result = New Collection
                cachedDaysValue = Replace(cachedDaysValue, "、", ",")
                dayParts = Split(cachedDaysValue, ",")
                
                For j = LBound(dayParts) To UBound(dayParts)
                    dayNum = ConvertDayNameToNumber(Trim(dayParts(j)))
                    If dayNum > 0 Then
                        result.Add dayNum
                    End If
                Next j
                
                Set GetCustomerDeliveryDays = result
            Else
                Set GetCustomerDeliveryDays = Nothing
            End If
            Exit Function
        End If
    End If
    
    lastRow = customerMasterWs.Cells(customerMasterWs.Rows.count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        If Trim(customerMasterWs.Cells(i, 1).Value) = customerName Then
            deliveryDaysValue = Trim(customerMasterWs.Cells(i, 2).Value)  ' B列
            
            If deliveryDaysValue <> "" Then
                Set result = New Collection
                deliveryDaysValue = Replace(deliveryDaysValue, "、", ",")
dayParts = Split(deliveryDaysValue, ",")
                
                For j = LBound(dayParts) To UBound(dayParts)
                    dayNum = ConvertDayNameToNumber(Trim(dayParts(j)))
                    If dayNum > 0 Then
                        result.Add dayNum
                    End If
                Next j
                
                Set GetCustomerDeliveryDays = result
            Else
                Set GetCustomerDeliveryDays = Nothing
            End If
            Exit Function
        End If
    Next i
    
    Set GetCustomerDeliveryDays = Nothing
End Function

' ============================================
' 曜日名を数字に変換（月=2, 火=3, 水=4, 木=5, 金=6）
' ※VBAのWeekday関数に合わせる（日=1, 月=2, ...）
' ============================================
Function ConvertDayNameToNumber(dayName As String) As Integer
    Select Case dayName
        Case "月", "月曜", "月曜日"
            ConvertDayNameToNumber = 2
        Case "火", "火曜", "火曜日"
            ConvertDayNameToNumber = 3
        Case "水", "水曜", "水曜日"
            ConvertDayNameToNumber = 4
        Case "木", "木曜", "木曜日"
            ConvertDayNameToNumber = 5
        Case "金", "金曜", "金曜日"
            ConvertDayNameToNumber = 6
        Case Else
            ConvertDayNameToNumber = 0
    End Select
End Function
' ============================================
' HTMLメール本文を生成
' ============================================
Function BuildEmailBodyHTML(customerName As String, branchSettings As Variant, _
                            stockoutInfoList As Collection, trackingInfoList As Collection, _
                            Optional bunnoInfoList As Collection = Nothing, _
                            Optional manufacturerMasterWs As Worksheet = Nothing, _
                            Optional holidays As Object = Nothing, _
                            Optional confirmingListWs As Worksheet = Nothing, _
                            Optional bunnoCompletedList As Collection = Nothing) As String
    Dim html As String
    Dim hasStockout As Boolean
    Dim hasTracking As Boolean
    
    hasStockout = False
    hasTracking = False
    
    Dim hasBunno As Boolean
    hasBunno = False
    
    If Not stockoutInfoList Is Nothing Then
        If stockoutInfoList.count > 0 Then hasStockout = True
    End If
    If Not trackingInfoList Is Nothing Then
        If trackingInfoList.count > 0 Then hasTracking = True
    End If
    If Not bunnoInfoList Is Nothing Then
        If bunnoInfoList.count > 0 Then hasBunno = True
    End If
    
    ' HTML開始
    html = "<html><head><style>" & _
           "body { font-family: 'メイリオ', 'Meiryo', sans-serif; font-size: 14px; line-height: 1.6; }" & _
           ".section { margin: 15px 0; padding: 10px; background-color: #f8f8f8; border-left: 4px solid #c0a040; }" & _
           ".section-title { font-weight: bold; color: #333; margin-bottom: 8px; }" & _
           ".stockout { border-left-color: #cc0000; }" & _
           ".stockout .section-title { color: #cc0000; }" & _
           ".tracking { border-left-color: #0066cc; }" & _
           ".tracking-link { color: #0066cc; font-weight: bold; }" & _
           ".item { margin-left: 20px; font-size: 13px; }" & _
           "</style></head><body>"
    
    ' 本文
    html = html & HtmlEscape(customerName) & " 御中<br><br>"
    html = html & "いつもお世話になっております。<br>"
    html = html & "マツモト産業㈱" & branchSettings(0) & "です。<br><br>"
    html = html & "ご注文ありがとうございます。<br>"
    html = html & "納期回答書をお送りいたします。<br><br>"
    
    ' 送り状情報セクション
    If hasTracking Then
        html = html & "<div class='section tracking'>"
        html = html & "<div class='section-title'>■ 送り状番号のご連絡</div>"
        
        ' 商品ごとの送り状セットを作成
        Dim productToTrackingSetHtml As Object
        Set productToTrackingSetHtml = CreateObject("Scripting.Dictionary")
        Dim tItem As Variant
        Dim productKeyHtml As String
        
        For Each tItem In trackingInfoList
            productKeyHtml = tItem(0) & "|" & tItem(1) & "|" & tItem(2)
            
            If Not productToTrackingSetHtml.Exists(productKeyHtml) Then
                productToTrackingSetHtml.Add productKeyHtml, CreateObject("Scripting.Dictionary")
            End If
            
            Dim tKeyHtml As String
            tKeyHtml = tItem(3) & "|" & tItem(4)
            If Not productToTrackingSetHtml(productKeyHtml).Exists(tKeyHtml) Then
                productToTrackingSetHtml(productKeyHtml).Add tKeyHtml, Array(tItem(3), tItem(4))
            End If
        Next tItem
        
        ' 送り状セット文字列でグループ化
        Dim trackingSetToProductsHtml As Object
        Set trackingSetToProductsHtml = CreateObject("Scripting.Dictionary")
        Dim pKeyHtml As Variant
        Dim trackingSetKeyHtml As String
        Dim tKeyVarHtml As Variant
        Dim trackingKeysHtml() As String
        Dim keyCountHtml As Long
        Dim tempKeyHtml As String
        Dim sortIHtml As Long, sortJHtml As Long
        
        For Each pKeyHtml In productToTrackingSetHtml.keys
            ' 送り状キーをソートして一意の文字列を作成
            keyCountHtml = productToTrackingSetHtml(pKeyHtml).count
            ReDim trackingKeysHtml(0 To keyCountHtml - 1)
            sortIHtml = 0
            For Each tKeyVarHtml In productToTrackingSetHtml(pKeyHtml).keys
                trackingKeysHtml(sortIHtml) = CStr(tKeyVarHtml)
                sortIHtml = sortIHtml + 1
            Next tKeyVarHtml
            
            ' バブルソート
            For sortIHtml = 0 To keyCountHtml - 2
                For sortJHtml = sortIHtml + 1 To keyCountHtml - 1
                    If trackingKeysHtml(sortIHtml) > trackingKeysHtml(sortJHtml) Then
                        tempKeyHtml = trackingKeysHtml(sortIHtml)
                        trackingKeysHtml(sortIHtml) = trackingKeysHtml(sortJHtml)
                        trackingKeysHtml(sortJHtml) = tempKeyHtml
                    End If
                Next sortJHtml
            Next sortIHtml
            
            trackingSetKeyHtml = Join(trackingKeysHtml, "||")
            
            If Not trackingSetToProductsHtml.Exists(trackingSetKeyHtml) Then
                trackingSetToProductsHtml.Add trackingSetKeyHtml, New Collection
            End If
            trackingSetToProductsHtml(trackingSetKeyHtml).Add Array(pKeyHtml, productToTrackingSetHtml(pKeyHtml))
        Next pKeyHtml
        
        ' 送り状セットごとに表示
        Dim setKeyVarHtml As Variant
        Dim productListHtml As Collection
        Dim productDataHtml As Variant
        Dim trackingDictHtml As Object
        Dim trackingItemHtml As Variant
        Dim trackingUrl As String
        Dim canDirect As Boolean
        Dim productPartsHtml As Variant
        Dim isMultiTrackingHtml As Boolean
        
        For Each setKeyVarHtml In trackingSetToProductsHtml.keys
            Set productListHtml = trackingSetToProductsHtml(setKeyVarHtml)
            
            ' 最初の商品から送り状情報を取得
            Set trackingDictHtml = productListHtml(1)(1)
            isMultiTrackingHtml = (trackingDictHtml.count >= 2)
            
            ' 送り状を表示
            For Each tKeyVarHtml In trackingDictHtml.keys
                trackingItemHtml = trackingDictHtml(tKeyVarHtml)
                trackingUrl = GetTrackingUrl(CStr(trackingItemHtml(0)), CStr(trackingItemHtml(1)))
                canDirect = CanDirectTrack(CStr(trackingItemHtml(0)))
                
                If canDirect And trackingUrl <> "" Then
                    html = html & "<a href='" & trackingUrl & "' class='tracking-link'>" & _
                           HtmlEscape(CStr(trackingItemHtml(0))) & "  " & trackingItemHtml(1) & "</a><br>"
                Else
                    html = html & "<span class='tracking-link'>" & HtmlEscape(CStr(trackingItemHtml(0))) & "  " & trackingItemHtml(1) & "</span><br>"
                    If trackingUrl <> "" Then
                        html = html & "<div class='item'>→ <a href='" & trackingUrl & "' style='color: #0066cc;'>追跡ページ</a>（番号を入力してください）</div>"
                    End If
                End If
            Next tKeyVarHtml
            
            ' 商品を表示
            For Each productDataHtml In productListHtml
                productPartsHtml = Split(productDataHtml(0), "|")
                html = html & "<div class='item'>・" & HtmlEscape(CStr(productPartsHtml(0))) & "  " & HtmlEscape(CStr(productPartsHtml(1))) & "  x" & productPartsHtml(2) & "</div>"
            Next productDataHtml
            
            ' 複数送り状の場合は注釈
            If isMultiTrackingHtml Then
                html = html & "<div style='margin-left: 20px; color: #666; font-size: 12px; font-style: italic;'>※別々の場所からの出荷になります</div>"
            End If
            
            html = html & "<br>"
        Next setKeyVarHtml
        
        html = html & "</div>"
    End If
    
    ' 欠品情報セクション
    If hasStockout Then
        html = html & "<div style='margin: 15px 0; padding: 12px; background-color: #fff0f0; border-left: 4px solid #cc0000;'>"
        html = html & "<div style='font-weight: bold; color: #cc0000; font-size: 15px; margin-bottom: 8px;'>【注意】欠品中の商品について</div>"
        html = html & "<div style='margin-bottom: 10px; color: #cc0000;'>下記商品は現在欠品中です。ご迷惑をおかけし申し訳ございません。</div>"
        
        Dim stockoutItem As Variant
        Dim stockoutDelivery As String
        Dim stockoutApprox As String
        
        For Each stockoutItem In stockoutInfoList
            stockoutDelivery = ""
            stockoutApprox = ""
            On Error Resume Next
            stockoutDelivery = stockoutItem(3)
            stockoutApprox = stockoutItem(4)
            On Error GoTo 0
            
            ' 表示優先順位：アバウト納期 > 確定納期 > 入荷次第ご連絡
            If stockoutApprox <> "" Then
                html = html & "<div style='margin-left: 20px; color: #cc0000; font-weight: bold;'>・" & _
                       HtmlEscape(CStr(stockoutItem(0))) & "  " & HtmlEscape(CStr(stockoutItem(1))) & "  x" & stockoutItem(2) & _
                       " → " & HtmlEscape(stockoutApprox) & "</div>"
            ElseIf stockoutDelivery = "" Or stockoutDelivery = "欠品中" Or stockoutDelivery = "確認中" Then
                html = html & "<div style='margin-left: 20px; color: #cc0000; font-weight: bold;'>・" & _
                       HtmlEscape(CStr(stockoutItem(0))) & "  " & HtmlEscape(CStr(stockoutItem(1))) & "  x" & stockoutItem(2) & _
                       " → 入荷次第ご連絡</div>"
            Else
                stockoutDelivery = Replace(stockoutDelivery, "（欠品）", "")
                html = html & "<div style='margin-left: 20px; color: #cc0000; font-weight: bold;'>・" & _
                       HtmlEscape(CStr(stockoutItem(0))) & "  " & HtmlEscape(CStr(stockoutItem(1))) & "  x" & stockoutItem(2) & _
                       " → " & HtmlEscape(stockoutDelivery) & "</div>"
            End If
        Next stockoutItem
        
        html = html & "</div>"
    End If
        
    ' 分納情報セクション
    If hasBunno Then
        ' 未定/確認中があるかチェック
        Dim hasBunnoMiteiForMail As Boolean
        hasBunnoMiteiForMail = False
        Dim bunnoItemMailCheck As Variant
        Dim bunnoDetailMailCheck As Collection
        
        For Each bunnoItemMailCheck In bunnoInfoList
            Set bunnoDetailMailCheck = bunnoItemMailCheck(3)
            
            If HasBunnoKakuninchu(bunnoDetailMailCheck) Then
                hasBunnoMiteiForMail = True
                Exit For
            End If
        Next bunnoItemMailCheck
        
        html = html & "<div style='margin: 15px 0; padding: 12px; background-color: #e8f0ff; border-left: 4px solid #0066cc;'>"
        html = html & "<div style='font-weight: bold; color: #0066cc; font-size: 15px; margin-bottom: 8px;'>■ 分納のご連絡</div>"
        html = html & "<div style='margin-bottom: 10px;'>下記商品は分納にてお届けいたします。</div>"
        
        ' 未定がある場合は注釈を追加
        If hasBunnoMiteiForMail Then
            html = html & "<div style='margin-bottom: 10px; color: #cc0000;'>※一部納期未定のためご迷惑をおかけいたします。確定次第ご連絡いたします。</div>"
        End If
        
        Dim bunnoItemHtml As Variant
        Dim bunnoDetailHtml As Collection
        Dim bunnoLineHtml As Variant
        Dim bunnoCounterHtml As Long
        Dim bunnoDateHtml As String
        Dim daysToAddHtml As Long
        
        For Each bunnoItemHtml In bunnoInfoList
            html = html & "<div style='margin-left: 10px; font-weight: bold; color: #003366;'>・" & _
                   HtmlEscape(CStr(bunnoItemHtml(0))) & " " & HtmlEscape(CStr(bunnoItemHtml(1))) & " x" & bunnoItemHtml(2) & "</div>"
            
            Set bunnoDetailHtml = bunnoItemHtml(3)
            daysToAddHtml = GetDeliveryDaysToAdd(CStr(bunnoItemHtml(5)), manufacturerMasterWs)
            bunnoCounterHtml = 0
            
            ' 同じ日付があるかチェック
            Dim hasSameDateHtml As Boolean
            hasSameDateHtml = CheckSameDateInBunno(bunnoDetailHtml)
            
            ' 【v12.0】注番・明細を取得
            Dim orderNumHtml As String
            Dim detailNumHtml As String
            orderNumHtml = ""
            detailNumHtml = ""
            On Error Resume Next
            orderNumHtml = bunnoItemHtml(6)
            detailNumHtml = bunnoItemHtml(7)
            On Error GoTo 0
            
            For Each bunnoLineHtml In bunnoDetailHtml
                bunnoCounterHtml = bunnoCounterHtml + 1
                
                ' 【v12.1】保存済みの計算済み納期を使用（インデックス3）
                bunnoDateHtml = ""
                If UBound(bunnoLineHtml) >= 3 Then
                    bunnoDateHtml = CStr(bunnoLineHtml(3))
                End If
                
                If bunnoDateHtml = "" Then
                    Dim isRosenbinHtml As Boolean
                    isRosenbinHtml = False
                    On Error Resume Next
                    isRosenbinHtml = CBool(bunnoItemHtml(8))
                    On Error GoTo 0
                    bunnoDateHtml = CalculateBunnoDate(CStr(bunnoLineHtml(1)), CBool(bunnoItemHtml(4)), _
                                                       daysToAddHtml, holidays, _
                                                       confirmingListWs, orderNumHtml, detailNumHtml, _
                                                       isRosenbinHtml)
                End If
                                
                ' 場所があれば追加
                Dim locationTextHtml As String
                locationTextHtml = ""
                If UBound(bunnoLineHtml) >= 2 Then
                    If bunnoLineHtml(2) <> "" Then
                        locationTextHtml = "（" & bunnoLineHtml(2) & "）"
                    End If
                End If
                
                ' 確認中または○月○旬予定の場合は赤字
                Dim isUncertain As Boolean
                isUncertain = (bunnoDateHtml = "確認中") Or _
                              (InStr(bunnoDateHtml, "予定") > 0 And InStr(bunnoDateHtml, "出荷") = 0 And InStr(bunnoDateHtml, "配達") = 0)
                
                ' 元データが未定/欠品/確認中だったかチェック
                Dim originalDateHtml As String
                originalDateHtml = CStr(bunnoLineHtml(1))
                
                If isUncertain Then
                    html = html & "<div style='margin-left: 30px; color: #cc0000; font-weight: bold;'>" & _
                           ToCircledNumber(bunnoCounterHtml) & HtmlEscape(CStr(bunnoLineHtml(0))) & " → " & _
                           HtmlEscape(bunnoDateHtml) & HtmlEscape(locationTextHtml) & "（確定次第ご連絡）</div>"
                ElseIf originalDateHtml = "未定" Or InStr(originalDateHtml, "欠品") > 0 Or InStr(originalDateHtml, "確認中") > 0 Then
                    ' 元は未定だったが今は確定 → オレンジ色
                    html = html & "<div style='margin-left: 30px; color: #c86400; font-weight: bold;'>" & _
                           ToCircledNumber(bunnoCounterHtml) & HtmlEscape(CStr(bunnoLineHtml(0))) & " → " & _
                           HtmlEscape(bunnoDateHtml) & HtmlEscape(locationTextHtml) & "</div>"
                Else
                    html = html & "<div style='margin-left: 30px;'>" & _
                           ToCircledNumber(bunnoCounterHtml) & HtmlEscape(CStr(bunnoLineHtml(0))) & " → " & _
                           HtmlEscape(bunnoDateHtml) & HtmlEscape(locationTextHtml) & "</div>"
                End If
            Next bunnoLineHtml
            
            ' 同じ日付がある場合は注釈を追加
            If hasSameDateHtml Then
                html = html & "<div style='margin-left: 20px; color: #666; font-size: 12px; font-style: italic;'>※別々の場所からの出荷になります</div>"
            End If
        Next bunnoItemHtml
        
        html = html & "</div>"
    End If
    
    ' 分納完了の通知
    Dim hasBunnoCompletedMail As Boolean
    hasBunnoCompletedMail = False
    If Not bunnoCompletedList Is Nothing Then
        If bunnoCompletedList.count > 0 Then hasBunnoCompletedMail = True
    End If
    
    If hasBunnoCompletedMail Then
        html = html & "<div style='margin: 15px 0; padding: 12px; background-color: #e8f5e9; border-left: 4px solid #28a745;'>"
        html = html & "<div style='font-weight: bold; color: #28a745; font-size: 15px; margin-bottom: 8px;'>■ 分納完了のご連絡</div>"
        html = html & "<div style='margin-bottom: 10px;'>分納でご注文いただいた商品は全て出荷が完了しました。</div>"
        
        Dim bcItemMail As Variant
        For Each bcItemMail In bunnoCompletedList
            html = html & "<div style='margin-left: 10px; font-weight: bold; color: #1b5e20;'>・" & _
                   HtmlEscape(CStr(bcItemMail(0))) & "  " & HtmlEscape(CStr(bcItemMail(1))) & _
                   "  x" & bcItemMail(2) & "</div>"
        Next bcItemMail
        
        html = html & "</div>"
    End If
    
    ' 確認中の注記
    html = html & "<br>※納期が「確認中」の商品については、<br>"
    html = html & "　メーカー確認後あらためてご連絡いたします。<br><br>"
    
    ' 締め
    html = html & "ご確認よろしくお願いいたします。<br>"
    html = html & "<div style='margin-top: 25px; color: #666666; font-size: 13px;'>"
    html = html & "マツモト産業株式会社<br>"
    html = html & branchSettings(0)
    html = html & "</div>"
    
    html = html & "</body></html>"
    
    BuildEmailBodyHTML = html
End Function

Sub 送付履歴ファイル修正()
    Dim filePath As String
    Dim wb As Workbook
    Dim ws As Worksheet
    
    filePath = ThisWorkbook.Path & "\送付履歴.xlsx"
    
    If Dir(filePath) = "" Then
        MsgBox "送付履歴.xlsxが見つかりません", vbExclamation
        Exit Sub
    End If
    
    Set wb = Workbooks.Open(filePath)
    
    On Error Resume Next
    Set ws = wb.Sheets("Sheet1")
    If Not ws Is Nothing Then
        ws.Name = "送付履歴"
        MsgBox "Sheet1を「送付履歴」にリネームしました", vbInformation
    End If
    On Error GoTo 0
    
    On Error Resume Next
    Set ws = wb.Sheets("確認中一覧")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Call CreateConfirmingSheet(wb)
        MsgBox "「確認中一覧」シートを追加しました", vbInformation
    Else
        ' ステータス列があるかチェック
        If ws.Cells(1, 8).Value <> "ステータス" Then
            Call AddStatusColumn(ws)
            MsgBox "確認中一覧に「ステータス」列を追加しました", vbInformation
        End If
        
        ' 入力規則を更新
        Call UpdateConfirmingValidation(ws)
        
        ' 確定納期→受注納期に列名変更
        If ws.Cells(1, 9).Value = "確定納期" Then
            ws.Cells(1, 9).Value = "受注納期"
        End If
    End If
    
    ' ===== 送付履歴テーブルのマイグレーション（8列→9列） =====
    Set ws = wb.Sheets("送付履歴")
    Dim tblHistory As ListObject
    On Error Resume Next
    Set tblHistory = ws.ListObjects("送付履歴テーブル")
    On Error GoTo 0
    
    If Not tblHistory Is Nothing Then
        If tblHistory.ListColumns.count = 8 Then
            ' B列に「受注日」列を挿入
            ws.Columns("B").Insert Shift:=xlToRight
            ws.Cells(1, 2).Value = "受注日"
            
            ' 列幅調整
            ws.Columns("A").ColumnWidth = 17
            ws.Columns("B").ColumnWidth = 12
            ws.Columns("C").ColumnWidth = 25
            ws.Columns("D").ColumnWidth = 15
            ws.Columns("E").ColumnWidth = 8
            ws.Columns("F").ColumnWidth = 20
            ws.Columns("G").ColumnWidth = 47
            ws.Columns("H").ColumnWidth = 22
            ws.Columns("I").ColumnWidth = 17.88
            
            ' 受注日列の表示形式
            If tblHistory.ListRows.count > 0 Then
                On Error Resume Next
                tblHistory.ListColumns("受注日").DataBodyRange.NumberFormat = "mm/dd"
                On Error GoTo 0
            End If
            
            MsgBox "送付履歴に「受注日」列を追加しました", vbInformation
        End If
    End If
    
    ' ===== 確認中テーブルのマイグレーション（10列→11列） =====
    Set ws = wb.Sheets("確認中一覧")
    Dim tblConfirm As ListObject
    On Error Resume Next
    Set tblConfirm = ws.ListObjects("確認中テーブル")
    On Error GoTo 0
    
    If Not tblConfirm Is Nothing Then
        If tblConfirm.ListColumns.count = 10 Then
            ' B列に「受注日」列を挿入
            ws.Columns("B").Insert Shift:=xlToRight
            ws.Cells(1, 2).Value = "受注日"
            
            ' 列幅調整
            ws.Columns("A").ColumnWidth = 17
            ws.Columns("B").ColumnWidth = 12
            ws.Columns("C").ColumnWidth = 25
            ws.Columns("D").ColumnWidth = 15
            ws.Columns("E").ColumnWidth = 8
            ws.Columns("F").ColumnWidth = 20
            ws.Columns("G").ColumnWidth = 47
            ws.Columns("H").ColumnWidth = 13
            ws.Columns("I").ColumnWidth = 12
            ws.Columns("J").ColumnWidth = 18
            ws.Columns("K").ColumnWidth = 17.88
            
            ' 受注日列の表示形式
            If tblConfirm.ListRows.count > 0 Then
                On Error Resume Next
                tblConfirm.ListColumns("受注日").DataBodyRange.NumberFormat = "mm/dd"
                On Error GoTo 0
            End If
            
            MsgBox "確認中一覧に「受注日」列を追加しました", vbInformation
        End If
    End If
        
    wb.Save
    wb.Close
    
    MsgBox "準備完了！これで納期回答書作成マクロが使えます。", vbInformation
End Sub

Sub AddStatusColumn(ws As Worksheet)
    ' 8列目（問合せ状況の右）に列を挿入
    ws.Columns(8).Insert Shift:=xlToRight
    
    ' ヘッダーを設定
    ws.Cells(1, 8).Value = "ステータス"
    
    ' 列幅を設定
    ws.Columns("H").ColumnWidth = 12
End Sub
Sub UpdateConfirmingValidation(ws As Worksheet)
    Dim tbl As ListObject
    Dim i As Long
    
    On Error Resume Next
    Set tbl = ws.ListObjects("確認中テーブル")
    On Error GoTo 0
    
    If tbl Is Nothing Then
        Exit Sub
    End If
    
    ' 行がある場合のみループ
    If tbl.ListRows.count > 0 Then
        For i = 1 To tbl.ListRows.count
            On Error Resume Next
            tbl.ListRows(i).Range(1, 8).Validation.Delete
            tbl.ListRows(i).Range(1, 8).Validation.Add Type:=xlValidateList, _
                AlertStyle:=xlValidAlertStop, Formula1:="未,済,回答待ち,除外"
            On Error GoTo 0
        Next i
    End If
    
    ' テーブルスタイルをグレー系に変更
    tbl.TableStyle = "TableStyleMedium1"
    
End Sub

Sub CreateConfirmingSheet(wb As Workbook)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim tblRange As Range
    
    Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.count))
    ws.Name = "確認中一覧"
    
    With ws
        .Cells(1, 1).Value = "送付日時"
        .Cells(1, 2).Value = "受注日"
        .Cells(1, 3).Value = "顧客名"
        .Cells(1, 4).Value = "受発注伝票"
        .Cells(1, 5).Value = "明細"
        .Cells(1, 6).Value = "メーカー名"
        .Cells(1, 7).Value = "品名"
        .Cells(1, 8).Value = "問合せ状況"
        .Cells(1, 9).Value = "ステータス"
        .Cells(1, 10).Value = "受注納期"
        .Cells(1, 11).Value = "送付者"
        
        .Cells(2, 1).Value = ""
        
        Set tblRange = .Range("A1:K2")
        Set tbl = .ListObjects.Add(xlSrcRange, tblRange, , xlYes)
        tbl.Name = "確認中テーブル"
        tbl.TableStyle = "TableStyleMedium1"
        
        tbl.ListRows(1).Delete
        
        .Columns("A").ColumnWidth = 17
        .Columns("B").ColumnWidth = 12
        .Columns("C").ColumnWidth = 25
        .Columns("D").ColumnWidth = 15
        .Columns("E").ColumnWidth = 8
        .Columns("F").ColumnWidth = 20
        .Columns("G").ColumnWidth = 47
        .Columns("H").ColumnWidth = 13
        .Columns("I").ColumnWidth = 12
        .Columns("J").ColumnWidth = 18
        .Columns("K").ColumnWidth = 17.88
    End With
    
    ws.Columns(1).NumberFormat = "mm/dd hh:nn"
    ws.Columns(2).NumberFormat = "mm/dd"
End Sub

' ============================================
' ファイルが使用中かチェック
' ============================================
Function IsFileOpen(filePath As String) As Boolean
    Dim fileNum As Integer
    On Error Resume Next
    fileNum = FreeFile
    Open filePath For Binary Access Read Write Lock Read Write As #fileNum
    Close #fileNum
    IsFileOpen = (Err.Number <> 0)
    On Error GoTo 0
End Function

' ============================================
' 送り状番号と運送会社を抽出（複数対応）
' 戻り値：Collection of Array(運送会社, 番号)
' ============================================
Function ExtractTrackingInfo(comment As String) As Collection
    Dim results As Collection
    Set results = New Collection
    
    Dim carriers As Variant
    Dim searchText As String
    Dim i As Long
    Dim startPos As Long
    Dim carrierName As String
    Dim afterCarrier As String
    Dim trackingNum As String
    Dim j As Long
    Dim c As String
    Dim foundPositions As Object
    Set foundPositions = CreateObject("Scripting.Dictionary")
    
    ' 対応する運送会社キーワード
    carriers = Array("ヤマト", "クロネコ", "佐川", "西濃", "福山", "郵便", "ゆうパック", _
                     "日通", "トナミ", "セイノー", "SSX", "JPロジ", "ＪＰロジ", _
                     "第一貨物", "第一")
    
    searchText = comment
    
    ' 各運送会社を検索
    For i = LBound(carriers) To UBound(carriers)
        startPos = 1
        
        Do While startPos <= Len(searchText)
            startPos = InStr(startPos, searchText, carriers(i))
            If startPos = 0 Then Exit Do
            
            ' 同じ位置で既に見つけた運送会社があればスキップ
            If foundPositions.Exists(startPos) Then
                startPos = startPos + 1
                GoTo ContinueSearch
            End If
            
            carrierName = carriers(i)
            
            ' 運送会社名の後ろを取得
            afterCarrier = Mid(searchText, startPos + Len(carrierName))
            
            ' 先頭のコロン（半角・全角）やスペースをスキップ
            Do While Len(afterCarrier) > 0
                c = Left(afterCarrier, 1)
                If c = ":" Or c = "：" Or c = " " Or c = "　" Then
                    afterCarrier = Mid(afterCarrier, 2)
                Else
                    Exit Do
                End If
            Loop
            
            ' 数字部分を抽出（半角・全角両対応）
            trackingNum = ""
            For j = 1 To Len(afterCarrier)
                c = Mid(afterCarrier, j, 1)
                If IsNumericChar(c) Then
                    trackingNum = trackingNum & ToHalfWidthNum(c)
                ElseIf c = "-" Or c = "－" Then
                    ' ハイフンはスキップ
                ElseIf trackingNum <> "" Then
                    Exit For
                End If
            Next j
            
            ' 10桁以上なら送り状番号として追加
            If Len(trackingNum) >= 10 Then
                results.Add Array(GetCarrierFullName(carrierName), trackingNum)
                foundPositions.Add startPos, True
            End If
            
            startPos = startPos + Len(carrierName)
ContinueSearch:
        Loop
    Next i
    
    Set ExtractTrackingInfo = results
End Function

' ============================================
' 数字かどうかを判定（半角・全角両対応）
' ============================================
Function IsNumericChar(c As String) As Boolean
    IsNumericChar = (c >= "0" And c <= "9") Or (c >= "０" And c <= "９")
End Function

' ============================================
' 全角数字を半角に変換
' ============================================
Function ToHalfWidthNum(c As String) As String
    Dim code As Long
    code = AscW(c)
    
    ' 全角数字（０～９）の範囲: &HFF10 ～ &HFF19
    If code >= &HFF10 And code <= &HFF19 Then
        ToHalfWidthNum = ChrW(code - &HFF10 + Asc("0"))
    Else
        ToHalfWidthNum = c
    End If
End Function

' ============================================
' 運送会社略称を正式名称に変換
' ============================================
Function GetCarrierFullName(shortName As String) As String
    Select Case shortName
        Case "ヤマト", "クロネコ"
            GetCarrierFullName = "ヤマト運輸"
        Case "佐川"
            GetCarrierFullName = "佐川急便"
        Case "西濃"
            GetCarrierFullName = "西濃運輸"
        Case "福山"
            GetCarrierFullName = "福山通運"
        Case "郵便", "ゆうパック"
            GetCarrierFullName = "日本郵便"
        Case "日通"
            GetCarrierFullName = "日本通運"
        Case "トナミ"
            GetCarrierFullName = "トナミ運輸"
        Case "セイノー", "SSX"
            GetCarrierFullName = "セイノースーパーエクスプレス"
        Case "JPロジ", "ＪＰロジ"
            GetCarrierFullName = "JPロジスティクス"
        Case "第一貨物", "第一"
            GetCarrierFullName = "第一貨物"
        Case Else
            GetCarrierFullName = shortName
    End Select
End Function
' ============================================
' 確認中一覧の古いデータを削除
' ============================================
Sub CleanOldConfirmingList(ws As Worksheet, daysToKeep As Long)
    Dim tbl As ListObject
    Dim cutoffDate As Date
    cutoffDate = Date - daysToKeep
    
    On Error Resume Next
    Set tbl = ws.ListObjects("確認中テーブル")
    On Error GoTo 0
    
    If tbl Is Nothing Or tbl.ListRows.Count = 0 Then Exit Sub
    
    On Error Resume Next
    tbl.AutoFilter.ShowAllData
    On Error GoTo 0
    
    ' === 配列一括読み取り ===
    Dim tblData As Variant
    tblData = tbl.DataBodyRange.Value
    Dim totalRows As Long
    totalRows = UBound(tblData, 1)
    Dim colCount As Long
    colCount = tbl.ListColumns.Count
    
    ' === 残す行を収集 ===
    Dim keepRows As Collection
    Set keepRows = New Collection
    Dim i As Long
    
    For i = 1 To totalRows
        Dim sentDate As Date
        sentDate = 0
        On Error Resume Next
        sentDate = CDate(tblData(i, 1))
        On Error GoTo 0
        
        If sentDate = 0 Or sentDate >= cutoffDate Then
            keepRows.Add i
        End If
    Next i
    
    ' === 削除が必要な場合のみ書き戻し ===
    If keepRows.Count < totalRows Then
        If keepRows.Count = 0 Then
            tbl.DataBodyRange.Delete
        Else
            Dim result() As Variant
            ReDim result(1 To keepRows.Count, 1 To colCount)
            Dim r As Long, j As Long
            r = 0
            Dim ki As Variant
            For Each ki In keepRows
                r = r + 1
                For j = 1 To colCount
                    result(r, j) = tblData(CLng(ki), j)
                Next j
            Next ki
            
            tbl.DataBodyRange.Delete
            On Error Resume Next
            If tbl.ListRows.Count = 0 Then
                tbl.ListRows.Add
            End If
            On Error GoTo 0
            tbl.Resize tbl.Range.Resize(keepRows.Count + 1, colCount)
            tbl.DataBodyRange.Value = result
        End If
    End If
End Sub
' ============================================
' 運送会社の追跡URLを生成
' ============================================
Function GetTrackingUrl(carrierName As String, trackingNum As String) As String
    ' ハイフン・スペースを除去（全運送会社共通）
    trackingNum = Replace(trackingNum, "-", "")
    trackingNum = Replace(trackingNum, " ", "")
    trackingNum = Replace(trackingNum, "　", "")
    
    If InStr(carrierName, "ヤマト") > 0 Then
        GetTrackingUrl = "https://member.kms.kuronekoyamato.co.jp/parcel/detail?pno=" & trackingNum
    ElseIf InStr(carrierName, "佐川") > 0 Then
        GetTrackingUrl = "https://k2k.sagawa-exp.co.jp/p/web/okurijosearch.do?okurijoNo=" & trackingNum
    ElseIf InStr(carrierName, "西濃") > 0 And InStr(carrierName, "スーパー") = 0 Then
        ' 西濃運輸（セイノースーパーエクスプレスではない）
        GetTrackingUrl = "https://track.seino.co.jp/cgi-bin/gnpquery.pgm?GNPNO1=" & trackingNum
    ElseIf InStr(carrierName, "福山") > 0 Or InStr(carrierName, "福通") > 0 Then
        GetTrackingUrl = "https://corp.fukutsu.co.jp/situation/tracking_no_hunt/" & trackingNum
    ElseIf InStr(carrierName, "郵便") > 0 Or InStr(carrierName, "ゆうパック") > 0 Then
        GetTrackingUrl = "https://trackings.post.japanpost.jp/services/srv/search/?requestNo1=" & trackingNum
    ElseIf InStr(carrierName, "日通") > 0 Or InStr(carrierName, "日本通運") > 0 Then
        GetTrackingUrl = "https://lp-trace.nittsu.co.jp/web/webarpaa702.srv?LANG=JP&officeselect2=&denpyoNo1=" & trackingNum
    ElseIf InStr(carrierName, "トナミ") > 0 Then
        GetTrackingUrl = "https://trc1.tonami.co.jp/trc/search3/excSearch3"
    ElseIf InStr(carrierName, "セイノー") > 0 Or InStr(carrierName, "SSX") > 0 Or InStr(carrierName, "スーパー") > 0 Then
        GetTrackingUrl = "http://inquire.trc.ssx.seino.co.jp/"
    ElseIf InStr(carrierName, "JPロジ") > 0 Or InStr(carrierName, "ＪＰロジ") > 0 Then
        GetTrackingUrl = "https://www.jp-logistics.jp/fwexphp/inquiry/chase/init"
    ElseIf InStr(carrierName, "第一貨物") > 0 Then
        GetTrackingUrl = "https://www.daiichi-kamotsu.co.jp/chase/contact_num/"
    Else
        GetTrackingUrl = ""
    End If
End Function
' ============================================
' 直接リンク可能な運送会社かどうか判定
' ============================================
Function CanDirectTrack(carrierName As String) As Boolean
    If InStr(carrierName, "ヤマト") > 0 Then
        CanDirectTrack = True
    ElseIf InStr(carrierName, "佐川") > 0 Then
        CanDirectTrack = True
    ElseIf InStr(carrierName, "西濃") > 0 And InStr(carrierName, "スーパー") = 0 Then
        CanDirectTrack = True
    ElseIf InStr(carrierName, "福山") > 0 Or InStr(carrierName, "福通") > 0 Then
        CanDirectTrack = True
    ElseIf InStr(carrierName, "郵便") > 0 Or InStr(carrierName, "ゆうパック") > 0 Then
        CanDirectTrack = True
    ElseIf InStr(carrierName, "日通") > 0 Or InStr(carrierName, "日本通運") > 0 Then
        CanDirectTrack = True
    Else
        CanDirectTrack = False
    End If
End Function
' ============================================
' 数字を丸数字に変換
' ============================================
Function ToCircledNumber(num As Long) As String
    Dim circled As Variant
    circled = Array("", "①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩")
    
    If num >= 1 And num <= 10 Then
        ToCircledNumber = circled(num)
    Else
        ToCircledNumber = CStr(num)
    End If
End Function
' ============================================
' 【修正版】分納情報を抽出（場所対応）
' 入力例：分納:700m 12/17 滋賀、300m 12/17 東京
' 戻り値：Collection of Array(数量, 日付文字列, 場所)
' ============================================
Function ExtractBunnoInfo(comment As String) As Collection
    Dim result As Collection
    Set result = New Collection
    
    Dim startPos As Long
    Dim bunnoText As String
    Dim parts As Variant
    Dim i As Long
    Dim part As String
    Dim qty As String
    Dim qtyOriginal As String
    Dim dateStr As String
    Dim locationStr As String
    Dim tokens As Variant
    Dim j As Long
    Dim c As String
    
    ' 「分納:」または「分納：」を探す
    startPos = InStr(comment, "分納:")
    If startPos = 0 Then startPos = InStr(comment, "分納：")
    
    If startPos = 0 Then
        Set ExtractBunnoInfo = result
        Exit Function
    End If
    
    ' 分納:以降を取得
    bunnoText = Mid(comment, startPos + 3)
    
    ' 全角数字を半角に変換
    bunnoText = ConvertToHalfWidth(bunnoText)
    
    ' 他のコメントがある場合は分納部分だけ取得
    Dim endPos As Long
    endPos = InStr(bunnoText, "  ")
    If endPos > 0 Then bunnoText = Left(bunnoText, endPos - 1)
    endPos = InStr(bunnoText, vbLf)
    If endPos > 0 Then bunnoText = Left(bunnoText, endPos - 1)
    
    ' 「,」と「、」を統一
    bunnoText = Replace(bunnoText, "、", ",")
    
    ' カンマで分割
    parts = Split(bunnoText, ",")
    
    For i = LBound(parts) To UBound(parts)
        part = Trim(parts(i))
        If part = "" Then GoTo NextPart
        
        ' 全角スペースを半角に
        part = Replace(part, "　", " ")
        
        ' スペースで分割してトークン化
        tokens = Split(part, " ")
        
        qty = ""
        dateStr = ""
        locationStr = ""
        
        Dim tokenIdx As Long
        For tokenIdx = LBound(tokens) To UBound(tokens)
            Dim token As String
            token = Trim(tokens(tokenIdx))
            If token = "" Then GoTo NextToken
            
            ' 括弧内の場所を先に抽出（トークンから括弧部分を除去）
            Dim bracketLocation As String
            bracketLocation = ExtractLocationFromToken(token)
            If bracketLocation <> "" And locationStr = "" Then
                locationStr = bracketLocation
            End If
            
            ' 数量判定（数字で始まる＋単位）
            If qty = "" And StartsWithNumber(token) Then
                ' 単位の後に日付が連結されているかチェック
                Dim splitResult As Variant
                splitResult = SplitQtyAndDate(token)
                
                If IsArray(splitResult) Then
                    qty = splitResult(0)
                    If dateStr = "" And splitResult(1) <> "" Then
                        dateStr = splitResult(1)
                    End If
                Else
                    qty = token
                End If
            ' 日付判定（スラッシュ含む or 未定 or ○月○旬予定）
            ElseIf dateStr = "" And IsDateToken(token) Then
                dateStr = token
            ' それ以外は場所
            ElseIf locationStr = "" And token <> "" Then
                locationStr = token
            End If
NextToken:
        Next tokenIdx
        
        ' 日付がなければ未定
        If dateStr = "" Then dateStr = "未定"
        
        ' 【v11.8追加】日付を正規化
        dateStr = NormalizeBunnoDate(dateStr)
        
        ' 元の数量文字列を保持
        qtyOriginal = qty
        
        ' 数値チェック用に単位を除去
        Dim qtyNum As String
        qtyNum = qty
        qtyNum = Replace(qtyNum, "個", "")
        qtyNum = Replace(qtyNum, "本", "")
        qtyNum = Replace(qtyNum, "台", "")
        qtyNum = Replace(qtyNum, "枚", "")
        qtyNum = Replace(qtyNum, "kg", "")
        qtyNum = Replace(qtyNum, "ｋｇ", "")
        qtyNum = Replace(qtyNum, "KG", "")
        qtyNum = Replace(qtyNum, "m", "")
        qtyNum = Replace(qtyNum, "ｍ", "")
        qtyNum = Replace(qtyNum, "M", "")
        qtyNum = Replace(qtyNum, "セット", "")
        qtyNum = Replace(qtyNum, "缶", "")
        qtyNum = Replace(qtyNum, "箱", "")
        
        If IsNumeric(qtyNum) And qtyNum <> "" Then
            result.Add Array(qtyOriginal, dateStr, locationStr)
        End If
NextPart:
    Next i
    
    Set ExtractBunnoInfo = result
End Function
' ============================================
' トークンから括弧内の場所を抽出（トークン自体も修正）
' ============================================
Function ExtractLocationFromToken(ByRef token As String) As String
    Dim openPos As Long
    Dim closePos As Long
    Dim location As String
    
    ExtractLocationFromToken = ""
    
    ' 全角括弧
    openPos = InStr(token, "（")
    closePos = InStr(token, "）")
    
    If openPos > 0 And closePos > openPos Then
        location = Mid(token, openPos + 1, closePos - openPos - 1)
        token = Left(token, openPos - 1)  ' トークンから括弧部分を除去
        ExtractLocationFromToken = location
        Exit Function
    End If
    
    ' 半角括弧
    openPos = InStr(token, "(")
    closePos = InStr(token, ")")
    
    If openPos > 0 And closePos > openPos Then
        location = Mid(token, openPos + 1, closePos - openPos - 1)
        token = Left(token, openPos - 1)
        ExtractLocationFromToken = location
        Exit Function
    End If
End Function

' ============================================
' 数字で始まるかチェック
' ============================================
Function StartsWithNumber(text As String) As Boolean
    If Len(text) = 0 Then
        StartsWithNumber = False
        Exit Function
    End If
    Dim firstChar As String
    firstChar = Left(text, 1)
    StartsWithNumber = IsNumeric(firstChar)
End Function

' ============================================
' 日付トークンかどうか判定（v11.8修正：欠品中・確認中対応）
' ============================================
Function IsDateToken(token As String) As Boolean
    IsDateToken = False
    
    ' 未定
    If token = "未定" Then
        IsDateToken = True
        Exit Function
    End If
    
    ' 欠品中・確認中系は未定扱い
    If InStr(token, "欠品") > 0 Or InStr(token, "確認中") > 0 Then
        IsDateToken = True
        Exit Function
    End If
    
    ' スラッシュ含む（12/20形式）
    If InStr(token, "/") > 0 Or InStr(token, "／") > 0 Then
        IsDateToken = True
        Exit Function
    End If
    
    ' ○月○旬予定 形式
    If InStr(token, "月") > 0 And InStr(token, "予定") > 0 Then
        IsDateToken = True
        Exit Function
    End If
End Function
' ============================================
' 【v11.8新規】分納の日付文字列を正規化
' 「12/22出荷」→「12/22」、「欠品中納期確認中」→「未定」
' ============================================
Function NormalizeBunnoDate(dateStr As String) As String
    Dim result As String
    result = Trim(dateStr)
    
    ' 欠品中・確認中系は未定扱い
    If InStr(result, "欠品") > 0 Or InStr(result, "確認中") > 0 Then
        NormalizeBunnoDate = "未定"
        Exit Function
    End If
    
    ' 「出荷」「着」等の余分な文字を除去
    result = Replace(result, "出荷", "")
    result = Replace(result, "着", "")
    result = Trim(result)
    
    ' 「予定」は○旬以外では除去
    If InStr(result, "上旬") = 0 And InStr(result, "中旬") = 0 And InStr(result, "下旬") = 0 Then
        result = Replace(result, "予定", "")
        result = Trim(result)
    End If
    
    NormalizeBunnoDate = result
End Function

' ============================================
' 全角を半角に変換（数字・スラッシュ・コロン）
' ============================================
Function ConvertToHalfWidth(text As String) As String
    Dim i As Long
    Dim c As String
    Dim code As Long
    Dim result As String
    
    result = ""
    
    For i = 1 To Len(text)
        c = Mid(text, i, 1)
        code = AscW(c)
        
        ' 全角数字（０～９）→ 半角
        If code >= &HFF10 And code <= &HFF19 Then
            result = result & ChrW(code - &HFF10 + Asc("0"))
        ' 全角スラッシュ → 半角
        ElseIf c = "／" Then
            result = result & "/"
        ' 全角コロン → 半角
        ElseIf c = "：" Then
            result = result & ":"
        Else
            result = result & c
        End If
    Next i
    
    ConvertToHalfWidth = result
End Function
' ============================================
' 【新規】分納に未定があるかチェック
' ============================================
Function HasBunnoMitei(bunnoInfo As Collection, _
                       Optional confirmingListWs As Worksheet = Nothing, _
                       Optional orderNumber As String = "", _
                       Optional detailNumber As String = "") As Boolean
    Dim item As Variant
    Dim confirmedDate As Date
    
    HasBunnoMitei = False
    
       
    If bunnoInfo Is Nothing Then Exit Function
    If bunnoInfo.count = 0 Then Exit Function
    
    confirmedDate = 0
    If Not confirmingListWs Is Nothing And orderNumber <> "" And detailNumber <> "" Then
        confirmedDate = GetConfirmedDeliveryDate(confirmingListWs, orderNumber, detailNumber)
    End If
    
    For Each item In bunnoInfo
    If item(1) = "未定" Or InStr(item(1), "欠品") > 0 Or InStr(item(1), "確認中") > 0 Or InStr(item(1), "予定") > 0 Then
        If confirmedDate > 0 Then
            ' 確定済み → 次のitemへ
        Else
            HasBunnoMitei = True
            Exit Function
        End If
    End If
Next item
End Function
' ============================================
' 【v12.2新規】計算済み納期に「確認中」があるかチェック
' ※「確定次第ご連絡」表示用
' ============================================
Function HasBunnoKakuninchu(bunnoDetail As Collection) As Boolean
    Dim item As Variant
    Dim calcDate As String
    
    HasBunnoKakuninchu = False
    
    If bunnoDetail Is Nothing Then Exit Function
    If bunnoDetail.count = 0 Then Exit Function
    
    For Each item In bunnoDetail
        calcDate = ""
        
        ' 計算済み納期(item(3))を見る
        If UBound(item) >= 3 Then
            calcDate = CStr(item(3))
        End If
        
        ' 「確認中」があれば未確定
        If calcDate = "確認中" Then
            HasBunnoKakuninchu = True
            Exit Function
        End If
        
        ' 「○旬予定」があれば未確定（「出荷予定」「配達予定」は除く）
        If InStr(calcDate, "予定") > 0 Then
            If InStr(calcDate, "出荷") = 0 And InStr(calcDate, "配達") = 0 Then
                HasBunnoKakuninchu = True
                Exit Function
            End If
        End If
    Next item
End Function
' ============================================
' 【新規】分納の日付を納期表示用に変換
' ============================================
Function CalculateBunnoDate(dateStr As String, isShipRule As Boolean, _
                            daysToAdd As Long, _
                            Optional holidays As Object = Nothing, _
                            Optional confirmingListWs As Worksheet = Nothing, _
                            Optional orderNumber As String = "", _
                            Optional detailNumber As String = "", _
                            Optional isRosenbin As Boolean = False) As String
    Dim bunnoDate As Date
    Dim monthNum As Integer
    Dim dayNum As Integer
    Dim slashPos As Long
    Dim today As Date
    
    today = Date
    
    ' 未定の場合 → 確認中一覧の受注納期をチェック
    If dateStr = "未定" Then
        If Not confirmingListWs Is Nothing And orderNumber <> "" And detailNumber <> "" Then
            Dim confirmedDate As Date
            confirmedDate = GetConfirmedDeliveryDate(confirmingListWs, orderNumber, detailNumber)
            
            If confirmedDate > 0 Then
                If isShipRule Then
                    If confirmedDate <= today Then
                        CalculateBunnoDate = Format(confirmedDate, "m月d日") & "出荷済み"
                    Else
                        CalculateBunnoDate = Format(confirmedDate, "m月d日") & "出荷予定"
                    End If
                ElseIf isRosenbin Then
                    Dim deliveryDateRosen As Date
                    deliveryDateRosen = AddBusinessDays(confirmedDate, WorksheetFunction.Max(daysToAdd - 1, 0), holidays)
                    If deliveryDateRosen <= today Then
                        CalculateBunnoDate = Format(deliveryDateRosen, "m月d日") & "出荷済み"
                    Else
                        CalculateBunnoDate = Format(deliveryDateRosen, "m月d日") & "出荷予定"
                    End If
                Else
                    Dim deliveryDate As Date
                    deliveryDate = AddBusinessDays(confirmedDate, daysToAdd, holidays)
                    
                    If deliveryDate <= today Then
                        CalculateBunnoDate = Format(deliveryDate, "m月d日") & "配達済み"
                    Else
                        CalculateBunnoDate = Format(deliveryDate, "m月d日") & "配達予定"
                    End If
                End If
                Exit Function
            End If
        End If
        
        CalculateBunnoDate = "確認中"
        Exit Function
    End If
    
    If InStr(dateStr, "予定") > 0 Then
    ' まず確認中一覧をチェック
    If Not confirmingListWs Is Nothing And orderNumber <> "" And detailNumber <> "" Then
        Dim confirmedDateForJun As Date
        confirmedDateForJun = GetConfirmedDeliveryDate(confirmingListWs, orderNumber, detailNumber)
        
        If confirmedDateForJun > 0 Then
            If isShipRule Then
                If confirmedDateForJun <= Date Then
                    CalculateBunnoDate = Format(confirmedDateForJun, "m月d日") & "出荷済み"
                Else
                    CalculateBunnoDate = Format(confirmedDateForJun, "m月d日") & "出荷予定"
                End If
            ElseIf isRosenbin Then
                Dim deliveryDateForJunRosen As Date
                deliveryDateForJunRosen = AddBusinessDays(confirmedDateForJun, WorksheetFunction.Max(daysToAdd - 1, 0), holidays)
                If deliveryDateForJunRosen <= Date Then
                    CalculateBunnoDate = Format(deliveryDateForJunRosen, "m月d日") & "出荷済み"
                Else
                    CalculateBunnoDate = Format(deliveryDateForJunRosen, "m月d日") & "出荷予定"
                End If
            Else
                Dim deliveryDateForJun As Date
                deliveryDateForJun = AddBusinessDays(confirmedDateForJun, daysToAdd, holidays)
                
                If deliveryDateForJun <= Date Then
                    CalculateBunnoDate = Format(deliveryDateForJun, "m月d日") & "配達済み"
                Else
                    CalculateBunnoDate = Format(deliveryDateForJun, "m月d日") & "配達予定"
                End If
            End If
            Exit Function
        End If
    End If
    
    CalculateBunnoDate = dateStr
    Exit Function
End If
    
    dateStr = Replace(dateStr, "／", "/")
    slashPos = InStr(dateStr, "/")
    
    If slashPos = 0 Then
        CalculateBunnoDate = "確認中"
        Exit Function
    End If
    
    On Error Resume Next
    monthNum = CInt(Left(dateStr, slashPos - 1))
    dayNum = CInt(Mid(dateStr, slashPos + 1))
    On Error GoTo 0
    
    If monthNum < 1 Or monthNum > 12 Or dayNum < 1 Or dayNum > 31 Then
        CalculateBunnoDate = "確認中"
        Exit Function
    End If
    
    bunnoDate = DateSerial(Year(today), monthNum, dayNum)
    
    If bunnoDate < today And (today - bunnoDate) > 180 Then
        bunnoDate = DateSerial(Year(today) + 1, monthNum, dayNum)
    End If
    
    If isShipRule Then
        If bunnoDate <= today Then
            CalculateBunnoDate = Format(bunnoDate, "m月d日") & "出荷済み"
        Else
            CalculateBunnoDate = Format(bunnoDate, "m月d日") & "出荷予定"
        End If
    ElseIf isRosenbin Then
        Dim deliveryDateCalcRosen As Date
        deliveryDateCalcRosen = AddBusinessDays(bunnoDate, WorksheetFunction.Max(daysToAdd - 1, 0), holidays)
        If deliveryDateCalcRosen <= today Then
            CalculateBunnoDate = Format(deliveryDateCalcRosen, "m月d日") & "出荷済み"
        Else
            CalculateBunnoDate = Format(deliveryDateCalcRosen, "m月d日") & "出荷予定"
        End If
    Else
        Dim deliveryDateCalc As Date
        deliveryDateCalc = AddBusinessDays(bunnoDate, daysToAdd, holidays)
        
        If deliveryDateCalc <= today Then
            CalculateBunnoDate = Format(deliveryDateCalc, "m月d日") & "配達済み"
        Else
            CalculateBunnoDate = Format(deliveryDateCalc, "m月d日") & "配達予定"
        End If
    End If
End Function
    
' ============================================
' コメント（社内）から着日を抽出
' 例：「@@12/20」→ 12月20日のDate
' ============================================
Function ExtractArrivalDateFromInternal(comment As String) As Date
    Dim startPos As Long
    Dim afterMarker As String
    Dim slashPos As Long
    Dim monthNum As Integer
    Dim dayNum As Integer
    Dim resultDate As Date
    Dim i As Long
    Dim c As String
    Dim dateStr As String
    
    ExtractArrivalDateFromInternal = 0
    
    ' 「@@」を探す（半角・全角両対応）
    startPos = InStr(comment, "@@")
    If startPos = 0 Then startPos = InStr(comment, "＠＠")
    If startPos = 0 Then Exit Function
    
    ' 「@@」の後ろを取得
    afterMarker = Mid(comment, startPos + 2)
    
    ' 全角を半角に変換
    afterMarker = ConvertToHalfWidth(afterMarker)
    
    ' 数字とスラッシュを抽出
    dateStr = ""
    For i = 1 To Len(afterMarker)
        c = Mid(afterMarker, i, 1)
        If IsNumeric(c) Or c = "/" Then
            dateStr = dateStr & c
        ElseIf dateStr <> "" Then
            Exit For
        End If
    Next i
    
    ' 日付をパース
    slashPos = InStr(dateStr, "/")
    If slashPos = 0 Then Exit Function
    
    On Error Resume Next
    monthNum = CInt(Left(dateStr, slashPos - 1))
    dayNum = CInt(Mid(dateStr, slashPos + 1))
    On Error GoTo 0
    
    If monthNum < 1 Or monthNum > 12 Or dayNum < 1 Or dayNum > 31 Then
        Exit Function
    End If
    
    resultDate = DateSerial(Year(Date), monthNum, dayNum)
    
    ' 過去の日付で180日以上前なら来年
    If resultDate < Date And (Date - resultDate) > 180 Then
        resultDate = DateSerial(Year(Date) + 1, monthNum, dayNum)
    End If
    
    ExtractArrivalDateFromInternal = resultDate
End Function
' ============================================
' 数量と日付が連結されている場合に分割
' 例：「1個12/19」→ Array("1個", "12/19")
' ============================================
Function SplitQtyAndDate(token As String) As Variant
    Dim units As Variant
    Dim i As Long
    Dim unitPos As Long
    Dim unitLen As Long
    Dim afterUnit As String
    
    units = Array("個", "本", "台", "枚", "セット", "缶", "箱", "kg", "ｋｇ", "KG", "m", "ｍ", "M")
    
    For i = LBound(units) To UBound(units)
        unitPos = InStr(token, units(i))
        If unitPos > 0 Then
            unitLen = Len(units(i))
            afterUnit = Mid(token, unitPos + unitLen)
            
            ' 単位の後に何かあれば分割
            If Len(afterUnit) > 0 Then
                ' 日付っぽいかチェック（数字で始まるか「未定」）
                If StartsWithNumber(afterUnit) Or afterUnit = "未定" Then
                    SplitQtyAndDate = Array(Left(token, unitPos + unitLen - 1), afterUnit)
                    Exit Function
                End If
            End If
        End If
    Next i
    
    ' 分割不要
    SplitQtyAndDate = token
End Function
' ============================================
' コメントからアバウト納期を抽出
' 入力例：「欠品中 1月上旬予定」「欠品中 12/20頃」
' 戻り値：「1月上旬入荷予定」「12月20日頃入荷予定」、なければ空文字
' ============================================
Function ExtractApproxDelivery(comment As String) As String
    Dim afterStockout As String
    Dim startPos As Long
    Dim monthNum As Integer
    Dim dayNum As Integer
    Dim slashPos As Long
    Dim monthPos As Long
    Dim monthStr As String
    Dim dayStr As String
    Dim dayPos As Long
    Dim i As Long
    
    ExtractApproxDelivery = ""
    
    ' 「欠品中」の後ろを取得
    startPos = InStr(comment, "欠品中")
    If startPos = 0 Then Exit Function
    
    afterStockout = Trim(Mid(comment, startPos + 3))
    If afterStockout = "" Then Exit Function
    
    ' 全角を半角に変換
    afterStockout = ConvertToHalfWidth(afterStockout)
    
    ' パターン1：「○月上旬予定」「○月中旬」「○月下旬予定」
    If InStr(afterStockout, "月") > 0 And _
       (InStr(afterStockout, "上旬") > 0 Or InStr(afterStockout, "中旬") > 0 Or InStr(afterStockout, "下旬") > 0) Then
        
        monthPos = InStr(afterStockout, "月")
        monthStr = ""
        
        For i = monthPos - 1 To 1 Step -1
            If IsNumeric(Mid(afterStockout, i, 1)) Then
                monthStr = Mid(afterStockout, i, 1) & monthStr
            Else
                Exit For
            End If
        Next i
        
        If monthStr <> "" Then
            monthNum = CInt(monthStr)
            
            If InStr(afterStockout, "上旬") > 0 Then
                ExtractApproxDelivery = monthNum & "月上旬入荷予定"
            ElseIf InStr(afterStockout, "中旬") > 0 Then
                ExtractApproxDelivery = monthNum & "月中旬入荷予定"
            ElseIf InStr(afterStockout, "下旬") > 0 Then
                ExtractApproxDelivery = monthNum & "月下旬入荷予定"
            End If
        End If
        Exit Function
    End If
    
    ' パターン2：「○/○頃」「○月○日頃」
    If InStr(afterStockout, "頃") > 0 Then
        slashPos = InStr(afterStockout, "/")
        
        If slashPos > 0 Then
            On Error Resume Next
            monthNum = CInt(Left(afterStockout, slashPos - 1))
            
            dayStr = ""
            For i = slashPos + 1 To Len(afterStockout)
                If IsNumeric(Mid(afterStockout, i, 1)) Then
                    dayStr = dayStr & Mid(afterStockout, i, 1)
                Else
                    Exit For
                End If
            Next i
            dayNum = CInt(dayStr)
            On Error GoTo 0
            
            If monthNum >= 1 And monthNum <= 12 And dayNum >= 1 And dayNum <= 31 Then
                ExtractApproxDelivery = monthNum & "月" & dayNum & "日頃入荷予定"
            End If
        ElseIf InStr(afterStockout, "月") > 0 And InStr(afterStockout, "日") > 0 Then
            monthPos = InStr(afterStockout, "月")
            dayPos = InStr(afterStockout, "日")
            
            On Error Resume Next
            monthStr = ""
            For i = monthPos - 1 To 1 Step -1
                If IsNumeric(Mid(afterStockout, i, 1)) Then
                    monthStr = Mid(afterStockout, i, 1) & monthStr
                Else
                    Exit For
                End If
            Next i
            monthNum = CInt(monthStr)
            
            dayStr = ""
            For i = monthPos + 1 To dayPos - 1
                If IsNumeric(Mid(afterStockout, i, 1)) Then
                    dayStr = dayStr & Mid(afterStockout, i, 1)
                End If
            Next i
            dayNum = CInt(dayStr)
            On Error GoTo 0
            
            If monthNum >= 1 And monthNum <= 12 And dayNum >= 1 And dayNum <= 31 Then
                ExtractApproxDelivery = monthNum & "月" & dayNum & "日頃入荷予定"
            End If
        End If
        Exit Function
    End If
    
    ' パターン3：「○月末」「○月末予定」
    If InStr(afterStockout, "月末") > 0 Then
        monthPos = InStr(afterStockout, "月")
        monthStr = ""
        
        For i = monthPos - 1 To 1 Step -1
            If IsNumeric(Mid(afterStockout, i, 1)) Then
                monthStr = Mid(afterStockout, i, 1) & monthStr
            Else
                Exit For
            End If
        Next i
        
        If monthStr <> "" Then
            monthNum = CInt(monthStr)
            ExtractApproxDelivery = monthNum & "月末入荷予定"
        End If
        Exit Function
    End If
End Function
' ============================================
' 欠品中テキストを除外（アバウト納期も含めて）
' 例：「欠品中 1月上旬予定」→ 全部除去
' ============================================
Function RemoveStockoutText(text As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim c As String
    Dim checkText As String
    Dim patternEnd As Long
    Dim hasPattern As Boolean
    
    RemoveStockoutText = text
    
    startPos = InStr(text, "欠品中")
    If startPos = 0 Then Exit Function
    
    ' 「欠品中」の後ろを確認
    endPos = startPos + 3
    
    ' スペースをスキップ
    Do While endPos <= Len(text)
        c = Mid(text, endPos, 1)
        If c <> " " And c <> "　" Then Exit Do
        endPos = endPos + 1
    Loop
    
    ' 納期パターンがあれば終端まで探す
    If endPos <= Len(text) Then
        hasPattern = False
        checkText = Mid(text, endPos)
        
        If InStr(checkText, "月") > 0 Or InStr(checkText, "/") > 0 Then
            patternEnd = 0
            
            If InStr(checkText, "予定") > 0 Then
                patternEnd = InStr(checkText, "予定") + 2
                hasPattern = True
            ElseIf InStr(checkText, "頃") > 0 Then
                patternEnd = InStr(checkText, "頃") + 1
                hasPattern = True
            ElseIf InStr(checkText, "上旬") > 0 Then
                patternEnd = InStr(checkText, "上旬") + 2
                hasPattern = True
            ElseIf InStr(checkText, "中旬") > 0 Then
                patternEnd = InStr(checkText, "中旬") + 2
                hasPattern = True
            ElseIf InStr(checkText, "下旬") > 0 Then
                patternEnd = InStr(checkText, "下旬") + 2
                hasPattern = True
            ElseIf InStr(checkText, "月末") > 0 Then
                patternEnd = InStr(checkText, "月末") + 2
                hasPattern = True
            End If
            
            If hasPattern Then
                endPos = endPos + patternEnd - 1
            End If
        End If
    End If
    
    ' 欠品中（＋アバウト納期）を除去
    RemoveStockoutText = Left(text, startPos - 1) & Mid(text, endPos)
    RemoveStockoutText = Trim(RemoveStockoutText)
End Function
' ============================================
' 同じ注番から保管場所を取得
' ============================================
Function GetStoragePlaceFromSameOrder(sourceWs As Worksheet, cols As Object, _
                                       orderNumber As String, currentRow As Long) As String
    Dim lastRow As Long
    Dim i As Long
    Dim rowOrderNum As String
    Dim rowStoragePlace As String
    
    GetStoragePlaceFromSameOrder = ""
    
    If orderNumber = "" Then Exit Function
    If Not cols.Exists("保管場所") Then Exit Function
    
    ' キャッシュから検索
    If Not g_StorageCache Is Nothing Then
        If g_StorageCache.Exists(orderNumber) Then
            GetStoragePlaceFromSameOrder = g_StorageCache(orderNumber)
            Exit Function
        End If
    End If
    
    ' フォールバック：シートスキャン
    lastRow = sourceWs.Cells(sourceWs.Rows.Count, cols("受発注伝票")).End(xlUp).Row
    
    For i = 7 To lastRow
        If i = currentRow Then GoTo NextRow
        
        rowOrderNum = Trim(g_SourceData(i, cols("受発注伝票")))
        
        If rowOrderNum = orderNumber Then
            rowStoragePlace = Trim(g_SourceData(i, cols("保管場所")))
            
            If rowStoragePlace <> "" Then
                GetStoragePlaceFromSameOrder = rowStoragePlace
                Exit Function
            End If
        End If
NextRow:
    Next i
End Function
' ============================================
' 顧客の保持日数を取得（C列）
' ============================================
Function GetRetentionDays(customerName As String, customerMasterWs As Worksheet) As Long
    Dim lastRow As Long
    Dim i As Long
    Dim retentionValue As Variant
    
    GetRetentionDays = 0
    
    ' キャッシュから検索
    If Not g_CustRetentionCache Is Nothing Then
        If g_CustRetentionCache.Exists(customerName) Then
            GetRetentionDays = g_CustRetentionCache(customerName)
            Exit Function
        End If
    End If
    
    ' フォールバック：シートスキャン
    If customerMasterWs Is Nothing Then Exit Function
    
    lastRow = customerMasterWs.Cells(customerMasterWs.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        If Trim(customerMasterWs.Cells(i, 1).Value) = customerName Then
            retentionValue = customerMasterWs.Cells(i, 3).Value
            If IsNumeric(retentionValue) And retentionValue <> "" Then
                If CLng(retentionValue) > 0 Then
                    GetRetentionDays = CLng(retentionValue)
                End If
            End If
            Exit Function
        End If
    Next i
End Function
' ============================================
' 顧客の路線便フラグを取得（D列）
' ============================================
Function IsRouteDelivery(customerName As String, customerMasterWs As Worksheet) As Boolean
    Dim lastRow As Long
    Dim i As Long
    
    IsRouteDelivery = False
    
    ' キャッシュから検索
    If Not g_CustRouteCache Is Nothing Then
        If g_CustRouteCache.Exists(customerName) Then
            IsRouteDelivery = g_CustRouteCache(customerName)
            Exit Function
        End If
    End If
    
    ' フォールバック：シートスキャン
    If customerMasterWs Is Nothing Then Exit Function
    
    lastRow = customerMasterWs.Cells(customerMasterWs.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        If Trim(customerMasterWs.Cells(i, 1).Value) = customerName Then
            If Trim(customerMasterWs.Cells(i, 4).Value) <> "" Then
                IsRouteDelivery = True
            End If
            Exit Function
        End If
    Next i
End Function
' ============================================
' 2つの日付間の営業日を数える（開始日を含まない）
' ============================================
Function CountBusinessDaysBetween(startDate As Date, endDate As Date, _
                                   Optional holidays As Object = Nothing) As Long
    Dim currentDate As Date
    Dim count As Long
    
    count = 0
    currentDate = startDate
    
    Do While currentDate < endDate
        currentDate = currentDate + 1
        
        ' 土日はスキップ
        If Weekday(currentDate) <> 1 And Weekday(currentDate) <> 7 Then
            ' 祝日チェック
            If holidays Is Nothing Then
                count = count + 1
            ElseIf Not holidays.Exists(CLng(currentDate)) Then
                count = count + 1
            Else
                ' 特別締切日（値あり）は営業日としてカウント
                If holidays(CLng(currentDate)) <> "" Then
                    count = count + 1
                End If
            End If
        End If
    Loop
    
    CountBusinessDaysBetween = count
End Function
' ============================================
' 【担当者別分割送信】得意先担当者セルの値をパース
' 「田中、鈴木」→ Collection("田中","鈴木")
' ============================================
Function ParseRepNames(cellValue As String) As Collection
    Dim result As Collection
    Set result = New Collection
    
    Dim normalized As String
    normalized = Trim(cellValue)
    
    If normalized = "" Then
        Set ParseRepNames = result
        Exit Function
    End If
    
    ' 区切り文字を半角カンマに統一（、 ・ 対応）
    normalized = Replace(normalized, "、", ",")
    normalized = Replace(normalized, "・", ",")
    ' 「様」を除去（「柏原様首藤様」→「柏原,首藤」対応）
    normalized = Replace(normalized, "様", ",")
    
    Dim parts() As String
    parts = Split(normalized, ",")
    
    Dim i As Long
    Dim partValue As String
    For i = LBound(parts) To UBound(parts)
        partValue = Trim(parts(i))
        If partValue <> "" Then
            result.Add partValue
        End If
    Next i
    
    Set ParseRepNames = result
End Function

' ============================================
' Collection内に指定の担当者名が含まれるか判定
' ============================================
Function ContainsRep(repNames As Collection, targetRep As String) As Boolean
    Dim item As Variant
    
    ContainsRep = False
    
    If repNames Is Nothing Then Exit Function
    
    For Each item In repNames
        If CStr(item) = targetRep Then
            ContainsRep = True
            Exit Function
        End If
    Next item
End Function

' ============================================
' 担当者マスターシートにその顧客が存在するか判定
' 存在すれば担当者別分割ON
' ============================================
Function IsSplitByRep(customerName As String, repMasterWs As Worksheet) As Boolean
    Dim lastRow As Long
    Dim i As Long
    
    IsSplitByRep = False
    
    If repMasterWs Is Nothing Then Exit Function
    
    lastRow = repMasterWs.Cells(repMasterWs.Rows.count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        If Trim(repMasterWs.Cells(i, 1).Value) = customerName Then
            IsSplitByRep = True
            Exit Function
        End If
    Next i
End Function

' ============================================
' 担当者マスターから指定顧客の担当者名リストを取得（重複排除）
' ============================================
Function GetRepList(customerName As String, repMasterWs As Worksheet) As Collection
    Dim result As Collection
    Set result = New Collection
    
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    
    If repMasterWs Is Nothing Then
        Set GetRepList = result
        Exit Function
    End If
    
    Dim lastRow As Long
    lastRow = repMasterWs.Cells(repMasterWs.Rows.count, 1).End(xlUp).Row
    
    Dim i As Long
    Dim repN As String
    For i = 2 To lastRow
        If Trim(repMasterWs.Cells(i, 1).Value) = customerName Then
            repN = Trim(repMasterWs.Cells(i, 2).Value)
            ' 末尾の「様」を除去
            If Len(repN) > 1 And Right(repN, 1) = "様" Then
                repN = Left(repN, Len(repN) - 1)
            End If
            If repN <> "" And Not seen.Exists(repN) Then
                seen.Add repN, True
                result.Add repN
            End If
        End If
    Next i
    
    Set GetRepList = result
End Function

' ============================================
' 担当者マスターから指定顧客・担当者のメールアドレスを取得
' C列以降を「;」区切りで返す（GetEmailAddressesと同じ形式）
' ============================================
Function GetRepEmailAddresses(customerName As String, repName As String, _
                               repMasterWs As Worksheet) As String
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim emailList As String
    Dim emailAddress As String
    
    GetRepEmailAddresses = ""
    
    If repMasterWs Is Nothing Then Exit Function
    
    lastRow = repMasterWs.Cells(repMasterWs.Rows.count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        Dim masterRepN As String
        masterRepN = Trim(repMasterWs.Cells(i, 2).Value)
        If Len(masterRepN) > 1 And Right(masterRepN, 1) = "様" Then
            masterRepN = Left(masterRepN, Len(masterRepN) - 1)
        End If
        If Trim(repMasterWs.Cells(i, 1).Value) = customerName And _
           masterRepN = repName Then
            emailList = ""
            For j = 3 To repMasterWs.Cells(i, repMasterWs.Columns.count).End(xlToLeft).Column
                emailAddress = Trim(repMasterWs.Cells(i, j).Value)
                If emailAddress <> "" Then
                    If emailList = "" Then
                        emailList = emailAddress
                    Else
                        emailList = emailList & "; " & emailAddress
                    End If
                End If
            Next j
            
            GetRepEmailAddresses = emailList
            Exit Function
        End If
    Next i
End Function

' ============================================
' 担当者フィルタ判定（CreateDeliveryReport/ByOrderNumbers共通）
' repName="" → 常にTrue（フィルタなし）
' repName="__OTHER__" → 未登録担当者が1人でもいればTrue
' repName=具体名 → その名前がセルに含まれていればTrue
' ============================================
Function ShouldIncludeForRep(cellRepValue As String, repName As String, _
                              registeredRepList As Collection) As Boolean
    If repName = "" Then
        ShouldIncludeForRep = True
        Exit Function
    End If
    
    Dim cellRepNames As Collection
    Set cellRepNames = ParseRepNames(cellRepValue)
    
    If repName = "__OTHER__" Then
        ' 空欄 → その他に含める
        If cellRepNames.count = 0 Then
            ShouldIncludeForRep = True
            Exit Function
        End If
        ' 未登録の担当者が1人でもいれば含める
        ShouldIncludeForRep = False
        Dim item As Variant
        For Each item In cellRepNames
            If Not ContainsRep(registeredRepList, CStr(item)) Then
                ShouldIncludeForRep = True
                Exit Function
            End If
        Next item
    Else
        ' 特定担当者：リストに含まれていればTrue
        ShouldIncludeForRep = ContainsRep(cellRepNames, repName)
    End If
End Function
' ============================================
' HTML特殊文字エスケープ
' ============================================
Function HtmlEscape(text As String) As String
    HtmlEscape = text
    HtmlEscape = Replace(HtmlEscape, "&", "&amp;")
    HtmlEscape = Replace(HtmlEscape, "<", "&lt;")
    HtmlEscape = Replace(HtmlEscape, ">", "&gt;")
    HtmlEscape = Replace(HtmlEscape, """", "&quot;")
End Function
