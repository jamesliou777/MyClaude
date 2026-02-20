Attribute VB_Name = "RCColumnAPI"
' =============================================================
' RC Column P-M Curve Calculator — Excel VBA Integration Module
' 用途：透過 HTTP API 呼叫 WPF 桌面應用程式計算 RC 柱交互作用曲線
' API 端點：http://localhost:5050/api/pmcurve
' =============================================================
Option Explicit

' ─────────────────────────────────────────────────────────────
' 主要巨集：讀取 Excel 輸入，呼叫 API，寫入計算結果
' ─────────────────────────────────────────────────────────────
Sub CalcPMCurve()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' --- 讀取基本參數 ---
    Dim fc   As Double  ' f'c (kgf/cm²)
    Dim fy   As Double  ' fy  (kgf/cm²)
    Dim Es   As Double  ' Es  (kgf/cm²)
    Dim cc   As Double  ' 保護層厚度 (cm)
    Dim stirDia As Double ' 箍筋直徑 (cm)

    fc      = ws.Range("B2").Value
    fy      = ws.Range("B3").Value
    Es      = ws.Range("B4").Value
    cc      = ws.Range("B5").Value
    stirDia = ws.Range("B6").Value

    ' --- 讀取外輪廓頂點 (XY 表格，起始列 B9) ---
    Dim outerPts As String
    outerPts = ReadXYTable(ws, 9, "B")

    ' --- 讀取空心區域 (XY 表格，起始列 B20，若無則空白) ---
    Dim hollowPts As String
    hollowPts = ReadXYTable(ws, 20, "B")

    ' --- 讀取鋼筋配置 (起始列 B32：no, x, y) ---
    Dim rebarArr As String
    rebarArr = ReadRebarTable(ws, 32, "B")

    ' --- 讀取載重組合 (起始列 B45：Pu, Mux, Muy) ---
    Dim loadArr As String
    loadArr = ReadLoadTable(ws, 45, "B")

    ' --- 組合 JSON 輸入 ---
    Dim inputJson As String
    inputJson = "{"
    inputJson = inputJson & """fc"":" & fc & ","
    inputJson = inputJson & """fy"":" & fy & ","
    inputJson = inputJson & """Es"":" & Es & ","
    inputJson = inputJson & """cc"":" & cc & ","
    inputJson = inputJson & """stirrupDia"":" & stirDia & ","
    inputJson = inputJson & """outer"":[" & outerPts & "],"
    If Len(hollowPts) > 0 Then
        inputJson = inputJson & """hollow"":[" & hollowPts & "],"
    End If
    inputJson = inputJson & """rebars"":[" & rebarArr & "],"
    inputJson = inputJson & """loads"":[" & loadArr & "]"
    inputJson = inputJson & "}"

    ' --- 呼叫 API ---
    Dim response As String
    response = PostRequest("http://localhost:5050/api/pmcurve", inputJson)

    If Len(response) = 0 Then
        MsgBox "API 未回應，請確認 WPF 應用程式已啟動。", vbExclamation
        Exit Sub
    End If

    ' --- 解析回應並寫入結果 ---
    Call WriteResults(ws, response)

    MsgBox "計算完成！", vbInformation
End Sub

' ─────────────────────────────────────────────────────────────
' 測試：連線確認
' ─────────────────────────────────────────────────────────────
Sub TestConnection()
    Dim result As String
    result = PingAPI()
    If Len(result) > 0 Then
        MsgBox "API 連線成功！" & vbCrLf & result, vbInformation
    Else
        MsgBox "API 連線失敗，請確認 WPF 應用程式已啟動於 port 5050。", vbExclamation
    End If
End Sub

' ─────────────────────────────────────────────────────────────
' 測試：使用範例資料（50×60 矩形空心柱，14 根鋼筋）
' ─────────────────────────────────────────────────────────────
Sub TestWithSampleData()
    ' 外輪廓：50×60 矩形（cm）
    Dim outer As String
    outer = "[0,0],[50,0],[50,60],[0,60]"

    ' 空心：20×30 矩形，置中
    Dim hollow As String
    hollow = "[15,15],[35,15],[35,45],[15,45]"

    ' 鋼筋：#8 (no=8, area=5.07 cm²)，沿外緣均布
    ' 四角 + 各邊中點，保護層 + 箍筋後鋼筋中心約 5 cm
    Dim rebars As String
    rebars = ""
    rebars = rebars & "{""no"":8,""x"":5,""y"":5},"
    rebars = rebars & "{""no"":8,""x"":25,""y"":5},"
    rebars = rebars & "{""no"":8,""x"":45,""y"":5},"
    rebars = rebars & "{""no"":8,""x"":45,""y"":30},"
    rebars = rebars & "{""no"":8,""x"":45,""y"":55},"
    rebars = rebars & "{""no"":8,""x"":25,""y"":55},"
    rebars = rebars & "{""no"":8,""x"":5,""y"":55},"
    rebars = rebars & "{""no"":8,""x"":5,""y"":30}"

    ' 載重組合
    Dim loads As String
    loads = ""
    loads = loads & "{""Pu"":300,""Mux"":50,""Muy"":80},"
    loads = loads & "{""Pu"":150,""Mux"":30,""Muy"":40},"
    loads = loads & "{""Pu"":500,""Mux"":10,""Muy"":10}"

    ' 組合 JSON
    Dim inputJson As String
    inputJson = "{"
    inputJson = inputJson & """fc"":280,"
    inputJson = inputJson & """fy"":4200,"
    inputJson = inputJson & """Es"":2040000,"
    inputJson = inputJson & """cc"":4,"
    inputJson = inputJson & """stirrupDia"":1.27,"
    inputJson = inputJson & """outer"":[" & outer & "],"
    inputJson = inputJson & """hollow"":[" & hollow & "],"
    inputJson = inputJson & """rebars"":[" & rebars & "],"
    inputJson = inputJson & """loads"":[" & loads & "]"
    inputJson = inputJson & "}"

    ' 呼叫 API
    Dim response As String
    response = PostRequest("http://localhost:5050/api/pmcurve", inputJson)

    If Len(response) = 0 Then
        MsgBox "API 未回應，請確認 WPF 應用程式已啟動。", vbExclamation
        Exit Sub
    End If

    ' 輸出至 ActiveSheet
    Call WriteResults(ActiveSheet, response)
    MsgBox "範例計算完成！", vbInformation
End Sub

' =============================================================
' HTTP 輔助函式
' =============================================================

' Ping /api/ping，回傳回應字串
Function PingAPI() As String
    On Error Resume Next
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    http.Open "GET", "http://localhost:5050/api/ping", False
    http.Send
    If http.Status = 200 Then
        PingAPI = http.responseText
    Else
        PingAPI = ""
    End If
    On Error GoTo 0
End Function

' POST JSON 至指定 URL，回傳回應字串
Function PostRequest(url As String, jsonBody As String) As String
    On Error Resume Next
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.Send jsonBody
    If http.Status = 200 Then
        PostRequest = http.responseText
    Else
        PostRequest = ""
    End If
    On Error GoTo 0
End Function

' =============================================================
' 讀取 Excel 表格輔助函式
' =============================================================

' 讀取 XY 頂點表格，格式：startCol=X, startCol+1=Y，遇空白停止
' 回傳 "[x1,y1],[x2,y2],..." 字串
Function ReadXYTable(ws As Worksheet, startRow As Long, startCol As String) As String
    Dim result As String
    Dim r As Long
    Dim colX As Long, colY As Long
    colX = ws.Columns(startCol).Column
    colY = colX + 1

    result = ""
    r = startRow
    Do While ws.Cells(r, colX).Value <> ""
        Dim x As Double, y As Double
        x = ws.Cells(r, colX).Value
        y = ws.Cells(r, colY).Value
        If Len(result) > 0 Then result = result & ","
        result = result & "[" & x & "," & y & "]"
        r = r + 1
    Loop
    ReadXYTable = result
End Function

' 讀取鋼筋表格，格式：no, x, y（遇空白停止）
' 回傳 JSON 物件陣列字串
Function ReadRebarTable(ws As Worksheet, startRow As Long, startCol As String) As String
    Dim result As String
    Dim r As Long
    Dim colNo As Long
    colNo = ws.Columns(startCol).Column

    result = ""
    r = startRow
    Do While ws.Cells(r, colNo).Value <> ""
        Dim no As Integer
        Dim rx As Double, ry As Double
        no = CInt(ws.Cells(r, colNo).Value)
        rx = ws.Cells(r, colNo + 1).Value
        ry = ws.Cells(r, colNo + 2).Value
        If Len(result) > 0 Then result = result & ","
        result = result & "{""no"":" & no & ",""x"":" & rx & ",""y"":" & ry & "}"
        r = r + 1
    Loop
    ReadRebarTable = result
End Function

' 讀取載重表格，格式：Pu, Mux, Muy（遇空白停止）
' 回傳 JSON 物件陣列字串
Function ReadLoadTable(ws As Worksheet, startRow As Long, startCol As String) As String
    Dim result As String
    Dim r As Long
    Dim colPu As Long
    colPu = ws.Columns(startCol).Column

    result = ""
    r = startRow
    Do While ws.Cells(r, colPu).Value <> ""
        Dim pu As Double, mux As Double, muy As Double
        pu  = ws.Cells(r, colPu).Value
        mux = ws.Cells(r, colPu + 1).Value
        muy = ws.Cells(r, colPu + 2).Value
        If Len(result) > 0 Then result = result & ","
        result = result & "{""Pu"":" & pu & ",""Mux"":" & mux & ",""Muy"":" & muy & "}"
        r = r + 1
    Loop
    ReadLoadTable = result
End Function

' =============================================================
' 結果寫入 Excel
' =============================================================
Sub WriteResults(ws As Worksheet, jsonResp As String)
    ' 找到輸出起始欄（D 欄 = 4）
    Const OUT_COL As Integer = 5   ' E 欄起

    ' ── 標題 ──
    With ws.Cells(1, OUT_COL)
        .Value = "RC 柱 P-M 計算結果"
        .Font.Bold = True
        .Font.Size = 12
    End With

    ' ── 斷面基本資訊 ──
    Dim r As Long
    r = 2
    ws.Cells(r, OUT_COL).Value = "斷面資訊"
    ws.Cells(r, OUT_COL).Font.Bold = True
    r = r + 1

    ws.Cells(r, OUT_COL).Value = "Ag (cm²)"
    ws.Cells(r, OUT_COL + 1).Value = ExtractJsonNum(jsonResp, "Ag")
    r = r + 1

    ws.Cells(r, OUT_COL).Value = "Ah (cm²)"
    ws.Cells(r, OUT_COL + 1).Value = ExtractJsonNum(jsonResp, "Ah")
    r = r + 1

    ws.Cells(r, OUT_COL).Value = "Ast (cm²)"
    ws.Cells(r, OUT_COL + 1).Value = ExtractJsonNum(jsonResp, "Ast")
    r = r + 1

    ws.Cells(r, OUT_COL).Value = "ρg (%)"
    ws.Cells(r, OUT_COL + 1).Value = ExtractJsonNum(jsonResp, "rhoG")
    r = r + 1

    ws.Cells(r, OUT_COL).Value = "塑性中心 pcX (cm)"
    ws.Cells(r, OUT_COL + 1).Value = ExtractJsonNum(jsonResp, "pcX")
    r = r + 1

    ws.Cells(r, OUT_COL).Value = "塑性中心 pcY (cm)"
    ws.Cells(r, OUT_COL + 1).Value = ExtractJsonNum(jsonResp, "pcY")
    r = r + 1

    r = r + 1

    ' ── 載重檢核結果 ──
    ws.Cells(r, OUT_COL).Value = "載重組合檢核"
    ws.Cells(r, OUT_COL).Font.Bold = True
    r = r + 1

    ' 標題列
    ws.Cells(r, OUT_COL).Value = "Pu (tf)"
    ws.Cells(r, OUT_COL + 1).Value = "Mux (tf·m)"
    ws.Cells(r, OUT_COL + 2).Value = "Muy (tf·m)"
    ws.Cells(r, OUT_COL + 3).Value = "φPn (tf)"
    ws.Cells(r, OUT_COL + 4).Value = "φMn (tf·m)"
    ws.Cells(r, OUT_COL + 5).Value = "Ratio"
    ws.Cells(r, OUT_COL + 6).Value = "狀態"

    Dim hdr As Integer
    For hdr = 0 To 6
        ws.Cells(r, OUT_COL + hdr).Font.Bold = True
        ws.Cells(r, OUT_COL + hdr).Interior.Color = RGB(68, 114, 196)
        ws.Cells(r, OUT_COL + hdr).Font.Color = RGB(255, 255, 255)
    Next hdr
    r = r + 1

    ' 解析 loadResults 陣列
    Dim loadsJson As String
    loadsJson = ExtractJsonArray(jsonResp, "loadResults")

    Dim loadItems() As String
    loadItems = SplitJsonObjects(loadsJson)

    Dim i As Integer
    For i = 0 To UBound(loadItems)
        If Len(Trim(loadItems(i))) = 0 Then GoTo NextLoad

        Dim item As String
        item = loadItems(i)

        Dim pu_v   As Double: pu_v   = ExtractJsonNum(item, "Pu")
        Dim mux_v  As Double: mux_v  = ExtractJsonNum(item, "Mux")
        Dim muy_v  As Double: muy_v  = ExtractJsonNum(item, "Muy")
        Dim phin_v As Double: phin_v = ExtractJsonNum(item, "phiPn")
        Dim phim_v As Double: phim_v = ExtractJsonNum(item, "phiMn")
        Dim ratio  As Double: ratio  = ExtractJsonNum(item, "ratio")
        Dim safe   As String: safe   = ExtractJsonStr(item, "safe")

        ws.Cells(r, OUT_COL).Value     = pu_v
        ws.Cells(r, OUT_COL + 1).Value = mux_v
        ws.Cells(r, OUT_COL + 2).Value = muy_v
        ws.Cells(r, OUT_COL + 3).Value = phin_v
        ws.Cells(r, OUT_COL + 4).Value = phim_v
        ws.Cells(r, OUT_COL + 5).Value = ratio

        If safe = "true" Then
            ws.Cells(r, OUT_COL + 6).Value = "OK"
            ws.Cells(r, OUT_COL + 6).Interior.Color = RGB(198, 239, 206)
            ws.Cells(r, OUT_COL + 6).Font.Color = RGB(0, 97, 0)
        Else
            ws.Cells(r, OUT_COL + 6).Value = "NG"
            ws.Cells(r, OUT_COL + 6).Interior.Color = RGB(255, 199, 206)
            ws.Cells(r, OUT_COL + 6).Font.Color = RGB(156, 0, 6)
        End If

        r = r + 1
NextLoad:
    Next i

    r = r + 1

    ' ── 平衡點資料 ──
    ws.Cells(r, OUT_COL).Value = "平衡點 (各方位角)"
    ws.Cells(r, OUT_COL).Font.Bold = True
    r = r + 1

    ws.Cells(r, OUT_COL).Value = "α (°)"
    ws.Cells(r, OUT_COL + 1).Value = "cb (cm)"
    ws.Cells(r, OUT_COL + 2).Value = "Pn_b (tf)"
    ws.Cells(r, OUT_COL + 3).Value = "Mn_b (tf·m)"
    ws.Cells(r, OUT_COL + 4).Value = "φPn_b (tf)"
    ws.Cells(r, OUT_COL + 5).Value = "φMn_b (tf·m)"

    For hdr = 0 To 5
        ws.Cells(r, OUT_COL + hdr).Font.Bold = True
        ws.Cells(r, OUT_COL + hdr).Interior.Color = RGB(68, 114, 196)
        ws.Cells(r, OUT_COL + hdr).Font.Color = RGB(255, 255, 255)
    Next hdr
    r = r + 1

    Dim balJson As String
    balJson = ExtractJsonArray(jsonResp, "balancePoints")

    Dim balItems() As String
    balItems = SplitJsonObjects(balJson)

    For i = 0 To UBound(balItems)
        If Len(Trim(balItems(i))) = 0 Then GoTo NextBal

        Dim bitem As String
        bitem = balItems(i)

        ws.Cells(r, OUT_COL).Value     = ExtractJsonNum(bitem, "alpha")
        ws.Cells(r, OUT_COL + 1).Value = ExtractJsonNum(bitem, "cb")
        ws.Cells(r, OUT_COL + 2).Value = ExtractJsonNum(bitem, "Pn_b")
        ws.Cells(r, OUT_COL + 3).Value = ExtractJsonNum(bitem, "Mn_b")
        ws.Cells(r, OUT_COL + 4).Value = ExtractJsonNum(bitem, "phiPn_b")
        ws.Cells(r, OUT_COL + 5).Value = ExtractJsonNum(bitem, "phiMn_b")

        r = r + 1
NextBal:
    Next i

    ' 自動調整欄寬
    ws.Columns(OUT_COL).AutoFit
    ws.Columns(OUT_COL + 1).AutoFit
    ws.Columns(OUT_COL + 2).AutoFit
    ws.Columns(OUT_COL + 3).AutoFit
    ws.Columns(OUT_COL + 4).AutoFit
    ws.Columns(OUT_COL + 5).AutoFit
    ws.Columns(OUT_COL + 6).AutoFit
End Sub

' =============================================================
' JSON 解析輔助函式（不依賴外部程式庫）
' =============================================================

' 從 JSON 字串中提取字串值 "key":"value"
Function ExtractJsonStr(json As String, key As String) As String
    Dim pattern As String
    pattern = """" & key & """:"""
    Dim pos As Long
    pos = InStr(json, pattern)
    If pos = 0 Then
        ExtractJsonStr = ""
        Exit Function
    End If
    pos = pos + Len(pattern)
    Dim endPos As Long
    endPos = InStr(pos, json, """")
    If endPos = 0 Then
        ExtractJsonStr = ""
        Exit Function
    End If
    ExtractJsonStr = Mid(json, pos, endPos - pos)
End Function

' 從 JSON 字串中提取數值 "key":number
Function ExtractJsonNum(json As String, key As String) As Double
    Dim pattern As String
    pattern = """" & key & """:"
    Dim pos As Long
    pos = InStr(json, pattern)
    If pos = 0 Then
        ExtractJsonNum = 0
        Exit Function
    End If
    pos = pos + Len(pattern)
    ' 跳過空白
    Do While Mid(json, pos, 1) = " "
        pos = pos + 1
    Loop
    ' 收集數字字元（含負號、小數點、指數）
    Dim numStr As String
    numStr = ""
    Dim ch As String
    Do
        ch = Mid(json, pos, 1)
        If ch = "-" Or ch = "+" Or ch = "." Or ch = "e" Or ch = "E" Or _
           (ch >= "0" And ch <= "9") Then
            numStr = numStr & ch
            pos = pos + 1
        Else
            Exit Do
        End If
    Loop
    If Len(numStr) > 0 Then
        ExtractJsonNum = CDbl(numStr)
    Else
        ExtractJsonNum = 0
    End If
End Function

' 提取 JSON 陣列內容 "key":[...]
Function ExtractJsonArray(json As String, key As String) As String
    Dim pattern As String
    pattern = """" & key & """["
    Dim pos As Long
    pos = InStr(json, pattern)
    If pos = 0 Then
        ' 嘗試帶空格的格式
        pattern = """" & key & """: ["
        pos = InStr(json, pattern)
    End If
    If pos = 0 Then
        ExtractJsonArray = ""
        Exit Function
    End If
    ' 找到開頭的 [
    Dim startPos As Long
    startPos = InStr(pos + Len(pattern) - 1, json, "[")
    If startPos = 0 Then
        ExtractJsonArray = ""
        Exit Function
    End If

    ' 找對應的 ]
    Dim depth As Integer
    depth = 0
    Dim p As Long
    For p = startPos To Len(json)
        Dim c As String
        c = Mid(json, p, 1)
        If c = "[" Then
            depth = depth + 1
        ElseIf c = "]" Then
            depth = depth - 1
            If depth = 0 Then
                ExtractJsonArray = Mid(json, startPos + 1, p - startPos - 1)
                Exit Function
            End If
        End If
    Next p
    ExtractJsonArray = ""
End Function

' 將 JSON 物件陣列字串拆分成個別物件
' 輸入："{...},{...},..." 回傳物件字串陣列
Function SplitJsonObjects(arrContent As String) As String()
    Dim result() As String
    Dim count As Integer
    count = 0
    ReDim result(0)

    If Len(Trim(arrContent)) = 0 Then
        SplitJsonObjects = result
        Exit Function
    End If

    Dim depth As Integer
    depth = 0
    Dim startPos As Long
    startPos = 0
    Dim p As Long

    For p = 1 To Len(arrContent)
        Dim c As String
        c = Mid(arrContent, p, 1)
        If c = "{" Then
            If depth = 0 Then startPos = p
            depth = depth + 1
        ElseIf c = "}" Then
            depth = depth - 1
            If depth = 0 And startPos > 0 Then
                ReDim Preserve result(count)
                result(count) = Mid(arrContent, startPos, p - startPos + 1)
                count = count + 1
                startPos = 0
            End If
        End If
    Next p

    If count = 0 Then
        ReDim result(0)
        result(0) = ""
    End If
    SplitJsonObjects = result
End Function
