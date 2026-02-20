Attribute VB_Name = "RCColumnAPI"
' =============================================================
' RC Column P-M Curve Calculator - Excel VBA Integration Module
' Calls WPF desktop app via HTTP API to compute P-M interaction curve
' API endpoint: http://localhost:5050/api/pmcurve
' =============================================================
Option Explicit

' ---------------------------------------------------------------
' Main macro: read Excel input, call API, write results
' ---------------------------------------------------------------
Sub CalcPMCurve()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' --- Basic parameters ---
    Dim fc      As Double   ' f'c (kgf/cm2)
    Dim fy      As Double   ' fy  (kgf/cm2)
    Dim Es      As Double   ' Es  (kgf/cm2)
    Dim cc      As Double   ' concrete cover (cm)
    Dim stirDia As Double   ' stirrup diameter (cm)

    fc      = ws.Range("B2").Value
    fy      = ws.Range("B3").Value
    Es      = ws.Range("B4").Value
    cc      = ws.Range("B5").Value
    stirDia = ws.Range("B6").Value

    ' --- Outer polygon vertices (XY table starting at row 9, col B) ---
    Dim outerPts As String
    outerPts = ReadXYTable(ws, 9, "B")

    ' --- Hollow polygon vertices (XY table starting at row 20, col B; leave empty if none) ---
    Dim hollowPts As String
    hollowPts = ReadXYTable(ws, 20, "B")

    ' --- Rebar layout (starting at row 32, col B: no, x, y) ---
    Dim rebarArr As String
    rebarArr = ReadRebarTable(ws, 32, "B")

    ' --- Load combinations (starting at row 45, col B: Pu, Mux, Muy) ---
    Dim loadArr As String
    loadArr = ReadLoadTable(ws, 45, "B")

    ' --- Build JSON input ---
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

    ' --- Call API ---
    Dim response As String
    response = PostRequest("http://localhost:5050/api/pmcurve", inputJson)

    If Len(response) = 0 Then
        MsgBox "API not responding. Please make sure the WPF application is running.", vbExclamation
        Exit Sub
    End If

    ' --- Parse response and write results ---
    Call WriteResults(ws, response)

    MsgBox "Calculation complete!", vbInformation
End Sub

' ---------------------------------------------------------------
' Test: ping API to verify connection
' ---------------------------------------------------------------
Sub TestConnection()
    Dim result As String
    result = PingAPI()
    If Len(result) > 0 Then
        MsgBox "API connection OK!" & vbCrLf & result, vbInformation
    Else
        MsgBox "API connection failed. Please make sure the WPF app is running on port 5050.", vbExclamation
    End If
End Sub

' ---------------------------------------------------------------
' Test: run with hardcoded sample data (50x60 hollow column, 8 rebars)
' ---------------------------------------------------------------
Sub TestWithSampleData()
    ' Outer boundary: 50x60 rectangle (cm)
    Dim outer As String
    outer = "[0,0],[50,0],[50,60],[0,60]"

    ' Hollow: 20x30 rectangle, centered
    Dim hollow As String
    hollow = "[15,15],[35,15],[35,45],[15,45]"

    ' Rebars: #8 (no=8, area=5.07 cm2), along perimeter, cover ~5 cm
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

    ' Load combinations
    Dim loads As String
    loads = ""
    loads = loads & "{""Pu"":300,""Mux"":50,""Muy"":80},"
    loads = loads & "{""Pu"":150,""Mux"":30,""Muy"":40},"
    loads = loads & "{""Pu"":500,""Mux"":10,""Muy"":10}"

    ' Build JSON
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

    ' Call API
    Dim response As String
    response = PostRequest("http://localhost:5050/api/pmcurve", inputJson)

    If Len(response) = 0 Then
        MsgBox "API not responding. Please make sure the WPF application is running.", vbExclamation
        Exit Sub
    End If

    Call WriteResults(ActiveSheet, response)
    MsgBox "Sample calculation complete!", vbInformation
End Sub

' =============================================================
' HTTP helper functions
' =============================================================

' Ping /api/ping and return response text
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

' POST JSON to url, return response text
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
' Excel table reader helpers
' =============================================================

' Read XY vertex table: startCol=X, startCol+1=Y, stop at blank
' Returns "[x1,y1],[x2,y2],..." string
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

' Read rebar table: no, x, y per row, stop at blank
' Returns JSON object array string
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

' Read load table: Pu, Mux, Muy per row, stop at blank
' Returns JSON object array string
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
' Write results to Excel (starting at column E)
' =============================================================
Sub WriteResults(ws As Worksheet, jsonResp As String)
    Const OUT_COL As Integer = 5   ' Column E

    ' Title
    With ws.Cells(1, OUT_COL)
        .Value = "RC Column P-M Curve Results"
        .Font.Bold = True
        .Font.Size = 12
    End With

    ' Section info
    Dim r As Long
    r = 2
    ws.Cells(r, OUT_COL).Value = "Section Properties"
    ws.Cells(r, OUT_COL).Font.Bold = True
    r = r + 1

    ws.Cells(r, OUT_COL).Value = "Ag (cm2)"
    ws.Cells(r, OUT_COL + 1).Value = ExtractJsonNum(jsonResp, "Ag")
    r = r + 1

    ws.Cells(r, OUT_COL).Value = "Ah (cm2)"
    ws.Cells(r, OUT_COL + 1).Value = ExtractJsonNum(jsonResp, "Ah")
    r = r + 1

    ws.Cells(r, OUT_COL).Value = "Ast (cm2)"
    ws.Cells(r, OUT_COL + 1).Value = ExtractJsonNum(jsonResp, "Ast")
    r = r + 1

    ws.Cells(r, OUT_COL).Value = "rhoG (%)"
    ws.Cells(r, OUT_COL + 1).Value = ExtractJsonNum(jsonResp, "rhoG")
    r = r + 1

    ws.Cells(r, OUT_COL).Value = "Plastic Centroid X (cm)"
    ws.Cells(r, OUT_COL + 1).Value = ExtractJsonNum(jsonResp, "pcX")
    r = r + 1

    ws.Cells(r, OUT_COL).Value = "Plastic Centroid Y (cm)"
    ws.Cells(r, OUT_COL + 1).Value = ExtractJsonNum(jsonResp, "pcY")
    r = r + 1

    r = r + 1

    ' Load check results
    ws.Cells(r, OUT_COL).Value = "Load Combination Check"
    ws.Cells(r, OUT_COL).Font.Bold = True
    r = r + 1

    ' Header row
    ws.Cells(r, OUT_COL).Value     = "Pu (tf)"
    ws.Cells(r, OUT_COL + 1).Value = "Mux (tf.m)"
    ws.Cells(r, OUT_COL + 2).Value = "Muy (tf.m)"
    ws.Cells(r, OUT_COL + 3).Value = "phiPn (tf)"
    ws.Cells(r, OUT_COL + 4).Value = "phiMn (tf.m)"
    ws.Cells(r, OUT_COL + 5).Value = "Ratio"
    ws.Cells(r, OUT_COL + 6).Value = "Status"

    Dim hdr As Integer
    For hdr = 0 To 6
        ws.Cells(r, OUT_COL + hdr).Font.Bold = True
        ws.Cells(r, OUT_COL + hdr).Interior.Color = RGB(68, 114, 196)
        ws.Cells(r, OUT_COL + hdr).Font.Color = RGB(255, 255, 255)
    Next hdr
    r = r + 1

    ' Parse loadResults array
    Dim loadsJson As String
    loadsJson = ExtractJsonArray(jsonResp, "loadResults")

    Dim loadItems() As String
    loadItems = SplitJsonObjects(loadsJson)

    Dim i As Integer
    For i = 0 To UBound(loadItems)
        If Len(Trim(loadItems(i))) = 0 Then GoTo NextLoad

        Dim item As String
        item = loadItems(i)

        ws.Cells(r, OUT_COL).Value     = ExtractJsonNum(item, "Pu")
        ws.Cells(r, OUT_COL + 1).Value = ExtractJsonNum(item, "Mux")
        ws.Cells(r, OUT_COL + 2).Value = ExtractJsonNum(item, "Muy")
        ws.Cells(r, OUT_COL + 3).Value = ExtractJsonNum(item, "phiPn")
        ws.Cells(r, OUT_COL + 4).Value = ExtractJsonNum(item, "phiMn")
        ws.Cells(r, OUT_COL + 5).Value = ExtractJsonNum(item, "ratio")

        Dim safe As String
        safe = ExtractJsonStr(item, "safe")

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

    ' Balance points
    ws.Cells(r, OUT_COL).Value = "Balance Points (per angle)"
    ws.Cells(r, OUT_COL).Font.Bold = True
    r = r + 1

    ws.Cells(r, OUT_COL).Value     = "alpha (deg)"
    ws.Cells(r, OUT_COL + 1).Value = "cb (cm)"
    ws.Cells(r, OUT_COL + 2).Value = "Pn_b (tf)"
    ws.Cells(r, OUT_COL + 3).Value = "Mn_b (tf.m)"
    ws.Cells(r, OUT_COL + 4).Value = "phiPn_b (tf)"
    ws.Cells(r, OUT_COL + 5).Value = "phiMn_b (tf.m)"

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

    ' Auto-fit columns
    Dim col As Integer
    For col = OUT_COL To OUT_COL + 6
        ws.Columns(col).AutoFit
    Next col
End Sub

' =============================================================
' JSON parsing helpers (no external library needed)
' =============================================================

' Extract string value:  "key":"value"
Function ExtractJsonStr(json As String, key As String) As String
    Dim pattern As String
    pattern = """" & key & """:"""
    Dim pos As Long
    pos = InStr(json, pattern)
    If pos = 0 Then ExtractJsonStr = "": Exit Function
    pos = pos + Len(pattern)
    Dim endPos As Long
    endPos = InStr(pos, json, """")
    If endPos = 0 Then ExtractJsonStr = "": Exit Function
    ExtractJsonStr = Mid(json, pos, endPos - pos)
End Function

' Extract numeric value:  "key":number
Function ExtractJsonNum(json As String, key As String) As Double
    Dim pattern As String
    pattern = """" & key & """:"
    Dim pos As Long
    pos = InStr(json, pattern)
    If pos = 0 Then ExtractJsonNum = 0: Exit Function
    pos = pos + Len(pattern)
    Do While Mid(json, pos, 1) = " "
        pos = pos + 1
    Loop
    Dim numStr As String
    numStr = ""
    Dim ch As String
    Do
        ch = Mid(json, pos, 1)
        If ch = "-" Or ch = "+" Or ch = "." Or ch = "e" Or ch = "E" _
           Or (ch >= "0" And ch <= "9") Then
            numStr = numStr & ch
            pos = pos + 1
        Else
            Exit Do
        End If
    Loop
    If Len(numStr) > 0 Then ExtractJsonNum = CDbl(numStr) Else ExtractJsonNum = 0
End Function

' Extract array content:  "key":[...]
Function ExtractJsonArray(json As String, key As String) As String
    Dim pattern As String
    pattern = """" & key & """["
    Dim pos As Long
    pos = InStr(json, pattern)
    If pos = 0 Then
        pattern = """" & key & """: ["
        pos = InStr(json, pattern)
    End If
    If pos = 0 Then ExtractJsonArray = "": Exit Function

    Dim startPos As Long
    startPos = InStr(pos + Len(pattern) - 1, json, "[")
    If startPos = 0 Then ExtractJsonArray = "": Exit Function

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

' Split JSON object array string "{...},{...},..." into individual object strings
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
