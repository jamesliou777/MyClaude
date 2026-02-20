Attribute VB_Name = "RCColumnAPI"
' ============================================================
'  RC 柱交互作用曲線計算器 - Excel VBA 介面模組
'  呼叫 WPF 應用程式提供的本機 HTTP API (localhost:5050)
'
'  使用方式：
'    1. 啟動 RCColumnCalculator.exe（視窗標題列會顯示 API 位址）
'    2. 在 Excel 填入斷面與載重資料（見下方說明）
'    3. 執行 CalcPMCurve 巨集
'
'  Excel 工作表「InputData」格式（A欄：標籤，B欄：值）：
'    B2  : fc  (kgf/cm²)
'    B3  : fy  (kgf/cm²)
'    B4  : Es  (kgf/cm²)
'    外輪廓頂點：從第7列起，A欄=X，B欄=Y（空列結束）
'    空心頂點：從第20列起，A欄=X，B欄=Y（空列結束，無空心則留空）
'    鋼筋：從第33列起，A欄=X，B欄=Y，C欄=面積cm²（空列結束）
'    載重：從第50列起，A欄=Pu(tf)，B欄=Mux(tf·m)，C欄=Muy(tf·m)
' ============================================================

Option Explicit

Private Const API_URL  As String = "http://localhost:5050"
Private Const SHEET_IN As String = "InputData"
Private Const SHEET_OUT As String = "Results"

' ────────────────────────────────────────────────────────────
' 主程式：讀取 Excel 資料 → 呼叫 API → 寫入結果
' ────────────────────────────────────────────────────────────
Sub CalcPMCurve()
    ' 1. 確認 API 是否運行
    If Not PingAPI() Then
        MsgBox "無法連線至 RC 柱計算器 API" & vbCrLf & _
               "請先啟動 RCColumnCalculator.exe", vbExclamation, "API 未啟動"
        Exit Sub
    End If

    ' 2. 讀取輸入資料
    Dim wsIn As Worksheet
    On Error GoTo ErrNoSheet
    Set wsIn = ThisWorkbook.Worksheets(SHEET_IN)
    On Error GoTo 0

    Dim fc As Double, fy As Double, Es As Double
    fc = CDbl(wsIn.Range("B2").Value)
    fy = CDbl(wsIn.Range("B3").Value)
    Es = CDbl(wsIn.Range("B4").Value)

    ' 讀取外輪廓（第7列起）
    Dim outerJson As String
    outerJson = ReadXYTable(wsIn, 7, "A", "B")

    ' 讀取空心（第20列起）
    Dim hollowJson As String
    hollowJson = ReadXYTable(wsIn, 20, "A", "B")

    ' 讀取鋼筋（第33列起）：X, Y, 面積
    Dim rebarsJson As String
    rebarsJson = ReadRebarTable(wsIn, 33)

    ' 讀取載重（第50列起）：Pu, Mux, Muy
    Dim loadsJson As String
    loadsJson = ReadLoadTable(wsIn, 50)

    ' 3. 組合 JSON 請求
    Dim reqJson As String
    reqJson = "{"
    reqJson = reqJson & """fc"":" & fc & ","
    reqJson = reqJson & """fy"":" & fy & ","
    reqJson = reqJson & """Es"":" & Es & ","
    reqJson = reqJson & """outer"":[" & outerJson & "],"
    reqJson = reqJson & """hollow"":[" & hollowJson & "],"
    reqJson = reqJson & """rebars"":[" & rebarsJson & "],"
    reqJson = reqJson & """loads"":[" & loadsJson & "]"
    reqJson = reqJson & "}"

    ' 4. 呼叫 API
    Dim respJson As String
    respJson = PostRequest(API_URL & "/api/pmcurve", reqJson)

    If respJson = "" Then
        MsgBox "API 無回應，請確認 WPF 程式已啟動", vbExclamation
        Exit Sub
    End If

    ' 5. 解析並寫入結果
    If InStr(respJson, """ok"":true") = 0 Then
        Dim errMsg As String
        errMsg = ExtractJsonStr(respJson, "error")
        MsgBox "計算失敗：" & errMsg, vbExclamation
        Exit Sub
    End If

    WriteResults respJson
    MsgBox "計算完成！結果已寫入「" & SHEET_OUT & "」工作表。", vbInformation

    Exit Sub
ErrNoSheet:
    MsgBox "找不到工作表「" & SHEET_IN & "」，請建立後再試。", vbExclamation
End Sub

' ────────────────────────────────────────────────────────────
' 測試：確認 API 是否運行
' ────────────────────────────────────────────────────────────
Sub TestConnection()
    If PingAPI() Then
        MsgBox "API 連線成功！" & vbCrLf & API_URL & "/api/ping", vbInformation
    Else
        MsgBox "無法連線至 API：" & API_URL & vbCrLf & _
               "請先啟動 RCColumnCalculator.exe", vbExclamation
    End If
End Sub

' ────────────────────────────────────────────────────────────
' 使用範例資料（矩形 50×60 空心，14根鋼筋）進行測試
' ────────────────────────────────────────────────────────────
Sub TestWithSampleData()
    If Not PingAPI() Then
        MsgBox "請先啟動 RCColumnCalculator.exe", vbExclamation
        Exit Sub
    End If

    Dim req As String
    req = "{"
    req = req & """fc"":280,""fy"":4200,""Es"":2040000,"
    req = req & """outer"":[[0,0],[50,0],[50,60],[0,60]],"
    req = req & """hollow"":[[15,15],[35,15],[35,45],[15,45]],"
    req = req & """rebars"":["
    req = req & "{""x"":5.27,""y"":5.27,""area"":5.07},"
    req = req & "{""x"":18.42,""y"":5.27,""area"":5.07},"
    req = req & "{""x"":31.58,""y"":5.27,""area"":5.07},"
    req = req & "{""x"":44.73,""y"":5.27,""area"":5.07},"
    req = req & "{""x"":44.73,""y"":17.63,""area"":5.07},"
    req = req & "{""x"":44.73,""y"":30,""area"":5.07},"
    req = req & "{""x"":44.73,""y"":42.36,""area"":5.07},"
    req = req & "{""x"":44.73,""y"":54.73,""area"":5.07},"
    req = req & "{""x"":31.58,""y"":54.73,""area"":5.07},"
    req = req & "{""x"":18.42,""y"":54.73,""area"":5.07},"
    req = req & "{""x"":5.27,""y"":54.73,""area"":5.07},"
    req = req & "{""x"":5.27,""y"":42.36,""area"":5.07},"
    req = req & "{""x"":5.27,""y"":30,""area"":5.07},"
    req = req & "{""x"":5.27,""y"":17.63,""area"":5.07}"
    req = req & "],"
    req = req & """loads"":[{""Pu"":100,""Mux"":-17.32,""Muy"":10}]"
    req = req & "}"

    Dim resp As String
    resp = PostRequest(API_URL & "/api/pmcurve", req)

    WriteResults resp
    MsgBox "範例計算完成！結果已寫入「" & SHEET_OUT & "」工作表。", vbInformation
End Sub

' ════════════════════════════════════════════════════════════
'  輔助函式
' ════════════════════════════════════════════════════════════

' ── HTTP GET Ping ────────────────────────────────────────────
Private Function PingAPI() As Boolean
    On Error GoTo Fail
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    http.Open "GET", API_URL & "/api/ping", False
    http.setRequestHeader "Content-Type", "application/json"
    http.send
    PingAPI = (http.Status = 200)
    Exit Function
Fail:
    PingAPI = False
End Function

' ── HTTP POST ───────────────────────────────────────────────
Private Function PostRequest(url As String, body As String) As String
    On Error GoTo Fail
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.send body
    If http.Status = 200 Then
        PostRequest = http.responseText
    Else
        PostRequest = "{""ok"":false,""error"":""HTTP " & http.Status & """}"
    End If
    Exit Function
Fail:
    PostRequest = ""
End Function

' ── 從工作表讀取 X,Y 座標表（回傳 JSON 陣列元素字串）───────────
Private Function ReadXYTable(ws As Worksheet, startRow As Long, _
                              colX As String, colY As String) As String
    Dim result As String
    Dim r As Long
    r = startRow
    Do While ws.Range(colX & r).Value <> "" And r < startRow + 50
        Dim x As Double, y As Double
        x = CDbl(ws.Range(colX & r).Value)
        y = CDbl(ws.Range(colY & r).Value)
        If result <> "" Then result = result & ","
        result = result & "[" & x & "," & y & "]"
        r = r + 1
    Loop
    ReadXYTable = result
End Function

' ── 從工作表讀取鋼筋表（X, Y, area）──────────────────────────────
Private Function ReadRebarTable(ws As Worksheet, startRow As Long) As String
    Dim result As String
    Dim r As Long
    r = startRow
    Do While ws.Cells(r, 1).Value <> "" And r < startRow + 100
        Dim x As Double, y As Double, a As Double
        x = CDbl(ws.Cells(r, 1).Value)
        y = CDbl(ws.Cells(r, 2).Value)
        a = CDbl(ws.Cells(r, 3).Value)
        If result <> "" Then result = result & ","
        result = result & "{""x"":" & x & ",""y"":" & y & ",""area"":" & a & "}"
        r = r + 1
    Loop
    ReadRebarTable = result
End Function

' ── 從工作表讀取載重表（Pu, Mux, Muy）────────────────────────────
Private Function ReadLoadTable(ws As Worksheet, startRow As Long) As String
    Dim result As String
    Dim r As Long
    r = startRow
    Do While ws.Cells(r, 1).Value <> "" And r < startRow + 50
        Dim Pu As Double, Mux As Double, Muy As Double
        Pu  = CDbl(ws.Cells(r, 1).Value)
        Mux = CDbl(ws.Cells(r, 2).Value)
        Muy = CDbl(ws.Cells(r, 3).Value)
        If result <> "" Then result = result & ","
        result = result & "{""Pu"":" & Pu & ",""Mux"":" & Mux & ",""Muy"":" & Muy & "}"
        r = r + 1
    Loop
    ReadLoadTable = result
End Function

' ── 寫入結果到 Results 工作表 ────────────────────────────────────
Private Sub WriteResults(respJson As String)
    ' 取得或建立 Results 工作表
    Dim wsOut As Worksheet
    On Error Resume Next
    Set wsOut = ThisWorkbook.Worksheets(SHEET_OUT)
    On Error GoTo 0
    If wsOut Is Nothing Then
        Set wsOut = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsOut.Name = SHEET_OUT
    End If
    wsOut.Cells.Clear

    Dim r As Long
    r = 1

    ' ── 標題 ──
    wsOut.Cells(r, 1).Value = "RC 柱 P-M 交互作用曲線計算結果"
    wsOut.Cells(r, 1).Font.Bold = True
    wsOut.Cells(r, 1).Font.Size = 14
    r = r + 2

    ' ── 斷面基本資料 ──
    wsOut.Cells(r, 1).Value = "=== 斷面基本資料 ==="
    wsOut.Cells(r, 1).Font.Bold = True
    r = r + 1

    Dim fields(7, 1) As String
    fields(0, 0) = "毛斷面積 Ag (cm²)"   : fields(0, 1) = ExtractJsonNum(respJson, "section", "Ag")
    fields(1, 0) = "空心面積 Ah (cm²)"   : fields(1, 1) = ExtractJsonNum(respJson, "section", "Ah")
    fields(2, 0) = "淨斷面積 Anet (cm²)": fields(2, 1) = ExtractJsonNum(respJson, "section", "Anet")
    fields(3, 0) = "鋼筋面積 Ast (cm²)" : fields(3, 1) = ExtractJsonNum(respJson, "section", "Ast")
    fields(4, 0) = "鋼筋比 ρg (%)"      : fields(4, 1) = ExtractJsonNum(respJson, "section", "rhoG_pct")
    fields(5, 0) = "Po (tf)"            : fields(5, 1) = ExtractJsonNum(respJson, "section", "Po_tf")
    fields(6, 0) = "φPo,max (tf)"       : fields(6, 1) = ExtractJsonNum(respJson, "section", "phiPo_max_tf")
    fields(7, 0) = "β₁"                 : fields(7, 1) = ExtractJsonNum(respJson, "section", "beta1")

    Dim i As Integer
    For i = 0 To 7
        wsOut.Cells(r, 1).Value = fields(i, 0)
        wsOut.Cells(r, 2).Value = CDbl(fields(i, 1))
        r = r + 1
    Next i
    r = r + 1

    ' ── 載重組合安全檢核 ──
    wsOut.Cells(r, 1).Value = "=== 載重組合安全檢核 ==="
    wsOut.Cells(r, 1).Font.Bold = True
    r = r + 1
    wsOut.Cells(r, 1).Value = "編號"
    wsOut.Cells(r, 2).Value = "Pu (tf)"
    wsOut.Cells(r, 3).Value = "Mux (tf·m)"
    wsOut.Cells(r, 4).Value = "Muy (tf·m)"
    wsOut.Cells(r, 5).Value = "Mu (tf·m)"
    wsOut.Cells(r, 6).Value = "α (°)"
    wsOut.Cells(r, 7).Value = "結果"
    wsOut.Rows(r).Font.Bold = True
    r = r + 1

    ' 解析 loadResults 陣列
    Dim loadSection As String
    loadSection = ExtractJsonArray(respJson, "loadResults")
    Dim items() As String
    items = SplitJsonObjects(loadSection)

    Dim j As Integer
    For j = 0 To UBound(items)
        If items(j) <> "" Then
            wsOut.Cells(r, 1).Value = CLng(ExtractField(items(j), "idx"))
            wsOut.Cells(r, 2).Value = CDbl(ExtractField(items(j), "Pu_tf"))
            wsOut.Cells(r, 3).Value = CDbl(ExtractField(items(j), "Mux_tfm"))
            wsOut.Cells(r, 4).Value = CDbl(ExtractField(items(j), "Muy_tfm"))
            wsOut.Cells(r, 5).Value = CDbl(ExtractField(items(j), "Mu_tfm"))
            wsOut.Cells(r, 6).Value = CDbl(ExtractField(items(j), "alpha_deg"))
            Dim isSafe As Boolean
            isSafe = (InStr(items(j), """safe"":true") > 0)
            wsOut.Cells(r, 7).Value = IIf(isSafe, "✓ 安全", "✗ 不安全")
            wsOut.Cells(r, 7).Font.Color = IIf(isSafe, RGB(0, 128, 0), RGB(200, 0, 0))
            wsOut.Cells(r, 7).Font.Bold = True
            r = r + 1
        End If
    Next j
    r = r + 1

    ' ── 各角度平衡點 ──
    wsOut.Cells(r, 1).Value = "=== 各角度平衡點 ==="
    wsOut.Cells(r, 1).Font.Bold = True
    r = r + 1
    wsOut.Cells(r, 1).Value = "α (°)"
    wsOut.Cells(r, 2).Value = "cb (cm)"
    wsOut.Cells(r, 3).Value = "Pn_b (tf)"
    wsOut.Cells(r, 4).Value = "Mn_b (tf·m)"
    wsOut.Cells(r, 5).Value = "φPn_b (tf)"
    wsOut.Cells(r, 6).Value = "φMn_b (tf·m)"
    wsOut.Cells(r, 7).Value = "φ"
    wsOut.Rows(r).Font.Bold = True
    r = r + 1

    Dim bpSection As String
    bpSection = ExtractJsonArray(respJson, "balancePoints")
    Dim bpItems() As String
    bpItems = SplitJsonObjects(bpSection)

    For j = 0 To UBound(bpItems)
        If bpItems(j) <> "" Then
            wsOut.Cells(r, 1).Value = CDbl(ExtractField(bpItems(j), "angle"))
            wsOut.Cells(r, 2).Value = CDbl(ExtractField(bpItems(j), "cb"))
            wsOut.Cells(r, 3).Value = CDbl(ExtractField(bpItems(j), "Pn_b"))
            wsOut.Cells(r, 4).Value = CDbl(ExtractField(bpItems(j), "Mn_b"))
            wsOut.Cells(r, 5).Value = CDbl(ExtractField(bpItems(j), "phiPn_b"))
            wsOut.Cells(r, 6).Value = CDbl(ExtractField(bpItems(j), "phiMn_b"))
            wsOut.Cells(r, 7).Value = CDbl(ExtractField(bpItems(j), "phi"))
            r = r + 1
        End If
    Next j

    ' 自動調整欄寬
    wsOut.Columns("A:G").AutoFit
    wsOut.Activate
    wsOut.Cells(1, 1).Select
End Sub

' ════════════════════════════════════════════════════════════
'  JSON 簡易解析工具函式
' ════════════════════════════════════════════════════════════

' 從 JSON 取得字串值：{"key":"value"} → value
Private Function ExtractJsonStr(json As String, key As String) As String
    Dim pattern As String
    pattern = """" & key & """:"""
    Dim pos As Long
    pos = InStr(json, pattern)
    If pos = 0 Then Exit Function
    pos = pos + Len(pattern)
    Dim endPos As Long
    endPos = InStr(pos, json, """")
    If endPos = 0 Then Exit Function
    ExtractJsonStr = Mid(json, pos, endPos - pos)
End Function

' 從 JSON 取得數值（簡化，適用於扁平物件）
Private Function ExtractJsonNum(json As String, section As String, key As String) As String
    ' 先找 section 的範圍，再找 key
    Dim secStart As Long
    Dim secLabel As String
    secLabel = """" & section & """{"
    secStart = InStr(json, secLabel)
    If secStart = 0 Then
        ' 嘗試 "section": {
        secLabel = """" & section & """: {"
        secStart = InStr(json, secLabel)
    End If
    Dim subJson As String
    If secStart > 0 Then
        Dim braceOpen As Long, braceClose As Long
        braceOpen = InStr(secStart, json, "{")
        Dim depth As Integer
        depth = 1
        Dim k As Long
        k = braceOpen + 1
        Do While depth > 0 And k <= Len(json)
            Dim ch As String
            ch = Mid(json, k, 1)
            If ch = "{" Then depth = depth + 1
            If ch = "}" Then depth = depth - 1
            k = k + 1
        Loop
        subJson = Mid(json, braceOpen, k - braceOpen)
    Else
        subJson = json
    End If
    ExtractJsonNum = ExtractField(subJson, key)
End Function

' 從 JSON 物件取得欄位值（數值或布林）
Private Function ExtractField(obj As String, key As String) As String
    Dim pattern As String
    pattern = """" & key & """:"
    Dim pos As Long
    pos = InStr(obj, pattern)
    If pos = 0 Then Exit Function
    pos = pos + Len(pattern)
    ' 跳過空白
    Do While Mid(obj, pos, 1) = " "
        pos = pos + 1
    Loop
    ' 找到值的結尾（逗號或大括號）
    Dim endPos As Long
    endPos = pos
    Do While endPos <= Len(obj)
        Dim c As String
        c = Mid(obj, endPos, 1)
        If c = "," Or c = "}" Or c = "]" Then Exit Do
        endPos = endPos + 1
    Loop
    ExtractField = Trim(Mid(obj, pos, endPos - pos))
End Function

' 取出 JSON 陣列內容字串：[...] → ...
Private Function ExtractJsonArray(json As String, key As String) As String
    Dim pattern As String
    pattern = """" & key & """["
    Dim pos As Long
    pos = InStr(json, pattern)
    If pos = 0 Then
        pattern = """" & key & """: ["
        pos = InStr(json, pattern)
    End If
    If pos = 0 Then Exit Function
    pos = InStr(pos, json, "[")
    Dim depth As Integer
    depth = 1
    Dim k As Long
    k = pos + 1
    Do While depth > 0 And k <= Len(json)
        Dim ch As String
        ch = Mid(json, k, 1)
        If ch = "[" Then depth = depth + 1
        If ch = "]" Then depth = depth - 1
        k = k + 1
    Loop
    ExtractJsonArray = Mid(json, pos + 1, k - pos - 2)
End Function

' 將 JSON 陣列字串分割為各個物件
Private Function SplitJsonObjects(arr As String) As String()
    Dim results() As String
    ReDim results(0)
    Dim depth As Integer
    depth = 0
    Dim start As Long
    start = 1
    Dim count As Long
    count = 0
    Dim i As Long
    For i = 1 To Len(arr)
        Dim c As String
        c = Mid(arr, i, 1)
        If c = "{" Then
            If depth = 0 Then start = i
            depth = depth + 1
        ElseIf c = "}" Then
            depth = depth - 1
            If depth = 0 Then
                ReDim Preserve results(count)
                results(count) = Mid(arr, start, i - start + 1)
                count = count + 1
            End If
        End If
    Next i
    If count = 0 Then
        ReDim results(0)
        results(0) = ""
    End If
    SplitJsonObjects = results
End Function
