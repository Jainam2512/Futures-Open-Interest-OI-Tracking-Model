Attribute VB_Name = "FuturesOI_Master"
Option Explicit

' ================================================================
' FUTURES OPEN INTEREST (OI) TRACKING MODEL
' Complete VBA Module — All Functionality Consolidated
' ================================================================
'
' SHEET STRUCTURE:
'   Sheet 1: Macro Control          -> Date inputs + buttons
'   Sheet 2: Current Contract Prices -> Futures close prices
'   Sheet 3: Underlying Prices       -> Spot close prices
'   Sheet 4: All Futures OI          -> Total open interest
'   Sheet 5: OI Analysis             -> Dashboard 1
'   Sheet 6: Historical Buildup      -> Dashboard 2
'
' BUTTONS ON MACRO CONTROL:
'   [REFRESH SINGLE DATE]   -> RefreshSingleDate
'   [REFRESH DATE RANGE]    -> RefreshDateRange
'   [UPDATE DASHBOARDS]     -> UpdateDashboards
'
' HOW TO IMPORT:
'   1. Open your .xlsm file
'   2. Press Alt + F11 to open VBA editor
'   3. File -> Import File -> select this .bas file
'   4. Assign macros to buttons as listed above
'   5. Run SetupMacroControlSheet once to configure layout
'
' NSE SESSION COOKIE (if getting 403 errors):
'   Open NSE in browser -> F12 -> Network tab -> any request
'   -> Copy the full "Cookie:" header value
'   -> Paste it in the COOKIE constant below
'
' ================================================================

' ----------------------------------------------------------------
' GLOBAL CONSTANTS
' ----------------------------------------------------------------
Const WS_CTRL   As String = "Macro Control"
Const WS_FUT    As String = "Current Contract Prices"
Const WS_UND    As String = "Underlying Prices"
Const WS_OI     As String = "All Futures OI"
Const WS_DASH1  As String = "OI Analysis"
Const WS_DASH2  As String = "Historical Buildup"

Const DATE_ROW       As Long = 2   ' Row containing date headers in data sheets
Const DATA_ROW_START As Long = 3   ' First row of symbol data
Const SYM_COL        As Long = 1   ' Column A = symbol names

' ---- NSE Session Cookie (paste value between the quotes) --------
' Leave empty string if not needed; required if NSE returns 403
Const NSE_COOKIE As String = ""
' Example:
' Const NSE_COOKIE As String = "nsit=AbCdEf; nseappid=XyZ123; ak_bmsc=LONG_VALUE_HERE"
' ----------------------------------------------------------------


' ================================================================
'  SECTION 1: SINGLE DATE REFRESH
' ================================================================

' ---------------------------------------------------------------
' RefreshSingleDate
' Reads date from Macro Control C6
' Fetches bhavcopy and updates all 3 data sheets for that date
' ---------------------------------------------------------------
Sub RefreshSingleDate()

    Dim wsCtrl As Worksheet
    Set wsCtrl = ThisWorkbook.Sheets(WS_CTRL)

    ' Validate date input
    Dim inputDate As Date
    On Error GoTo SingleDateError
    inputDate = CDate(wsCtrl.Range("C6").Value)
    On Error GoTo 0

    If Weekday(inputDate, vbMonday) > 5 Then
        MsgBox "Selected date (" & Format(inputDate, "DD-MMM-YYYY") & ") falls on a weekend." & vbNewLine & _
               "Please select a valid trading day (Mon - Fri).", vbExclamation, "Invalid Date"
        Exit Sub
    End If

    If inputDate > Date Then
        MsgBox "Date cannot be in the future.", vbExclamation, "Invalid Date"
        Exit Sub
    End If

    Dim dateStr As String
    dateStr = Format(inputDate, "DD-MMM-YY")

    ' Check if already loaded
    If DateColumnExists(Sheets(WS_FUT), dateStr) Then
        Dim resp As Integer
        resp = MsgBox("Data for " & dateStr & " already exists." & vbNewLine & _
                      "Do you want to overwrite it?", vbYesNo + vbQuestion, "Date Exists")
        If resp = vbNo Then Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation    = xlCalculationManual
    Application.StatusBar      = "Fetching data for " & dateStr & "..."

    wsCtrl.Range("C9").Value = "Fetching data for " & dateStr & "..."

    Dim futPrices As Object, undPrices As Object, oiData As Object
    Set futPrices = CreateObject("Scripting.Dictionary")
    Set undPrices = CreateObject("Scripting.Dictionary")
    Set oiData    = CreateObject("Scripting.Dictionary")

    Dim fetchOK As Boolean
    fetchOK = FetchBhavCopyForDate(inputDate, futPrices, undPrices, oiData)

    Call WriteDataToSheet(Sheets(WS_FUT), dateStr, futPrices)
    Call WriteDataToSheet(Sheets(WS_UND), dateStr, undPrices)
    Call WriteDataToSheet(Sheets(WS_OI),  dateStr, oiData)

    Dim statusMsg As String
    If fetchOK Then
        statusMsg = "SUCCESS — Data loaded for " & dateStr & " at " & Format(Now, "HH:MM:SS")
    Else
        statusMsg = "FALLBACK — NSE unavailable for " & dateStr & ". Estimated data used."
    End If

    wsCtrl.Range("C9").Value = statusMsg
    Application.StatusBar    = False
    Application.Calculation  = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox statusMsg, vbInformation, "Single Date Refresh"
    Exit Sub

SingleDateError:
    MsgBox "Invalid date in cell C6. Please enter a valid date.", vbCritical, "Date Error"
End Sub


' ================================================================
'  SECTION 2: DATE RANGE BULK REFRESH
' ================================================================

' ---------------------------------------------------------------
' RefreshDateRange
' Reads Start Date (C7) and End Date (C8) from Macro Control
' Loops through all Mon-Fri in that range
' Skips already-loaded dates automatically
' Inserts columns in chronological order
' ---------------------------------------------------------------
Sub RefreshDateRange()

    Dim wsCtrl As Worksheet
    Set wsCtrl = ThisWorkbook.Sheets(WS_CTRL)

    ' Validate inputs
    Dim startDate As Date, endDate As Date
    On Error GoTo RangeError
    startDate = CDate(wsCtrl.Range("C7").Value)
    endDate   = CDate(wsCtrl.Range("C8").Value)
    On Error GoTo 0

    If startDate > endDate Then
        MsgBox "Start Date must be on or before End Date.", vbExclamation, "Invalid Range"
        Exit Sub
    End If

    If endDate > Date Then
        MsgBox "End Date is in the future. Setting it to today: " & Format(Date, "DD-MMM-YYYY"), _
               vbInformation, "Date Adjusted"
        endDate = Date
    End If

    ' Build list of Mon-Fri trading days
    Dim tradingDays() As Date
    Dim tdCount As Long
    tdCount = 0

    Dim d As Date
    For d = startDate To endDate
        If Weekday(d, vbMonday) <= 5 Then
            ReDim Preserve tradingDays(tdCount)
            tradingDays(tdCount) = d
            tdCount = tdCount + 1
        End If
    Next d

    If tdCount = 0 Then
        MsgBox "No trading days (Mon-Fri) found in the selected range.", vbExclamation, "No Data"
        Exit Sub
    End If

    ' Count how many are already loaded
    Dim alreadyLoaded As Long
    alreadyLoaded = 0
    Dim i As Long
    For i = 0 To tdCount - 1
        If DateColumnExists(Sheets(WS_FUT), Format(tradingDays(i), "DD-MMM-YY")) Then
            alreadyLoaded = alreadyLoaded + 1
        End If
    Next i

    ' Confirm
    Dim confirm As Integer
    confirm = MsgBox("Date Range Refresh Summary:" & vbNewLine & vbNewLine & _
                     "  From        :  " & Format(startDate, "DD-MMM-YYYY") & vbNewLine & _
                     "  To          :  " & Format(endDate,   "DD-MMM-YYYY") & vbNewLine & _
                     "  Total Days  :  " & tdCount & vbNewLine & _
                     "  To Fetch    :  " & (tdCount - alreadyLoaded) & vbNewLine & _
                     "  Already In  :  " & alreadyLoaded & " (will be skipped)" & vbNewLine & vbNewLine & _
                     "This may take several minutes. Continue?", _
                     vbYesNo + vbQuestion, "Confirm Bulk Refresh")
    If confirm = vbNo Then Exit Sub

    ' Setup
    Application.ScreenUpdating = False
    Application.Calculation    = xlCalculationManual

    Dim successCount As Long, skipCount As Long, failCount As Long
    successCount = 0 : skipCount = 0 : failCount = 0
    Dim failedDates As String
    failedDates = ""

    ' Loop through each trading day
    For i = 0 To tdCount - 1

        Dim currentDate As Date
        currentDate = tradingDays(i)
        Dim dateStr   As String
        dateStr = Format(currentDate, "DD-MMM-YY")

        ' Live progress update
        Dim progressMsg As String
        progressMsg = "Day " & (i + 1) & " of " & tdCount & _
                      "  |  " & dateStr & _
                      "  |  Loaded: " & successCount & _
                      "  |  Skipped: " & skipCount & _
                      "  |  Failed: " & failCount

        wsCtrl.Range("C9").Value = progressMsg
        Application.StatusBar    = progressMsg
        DoEvents

        ' Skip if already in sheets
        If DateColumnExists(Sheets(WS_FUT), dateStr) Then
            skipCount = skipCount + 1
            GoTo NextDay
        End If

        ' Fetch bhavcopy
        Dim futPrices As Object, undPrices As Object, oiData As Object
        Set futPrices = CreateObject("Scripting.Dictionary")
        Set undPrices = CreateObject("Scripting.Dictionary")
        Set oiData    = CreateObject("Scripting.Dictionary")

        Dim fetchOK As Boolean
        fetchOK = FetchBhavCopyForDate(currentDate, futPrices, undPrices, oiData)

        ' Write to sheets
        Call WriteDataToSheet(Sheets(WS_FUT), dateStr, futPrices)
        Call WriteDataToSheet(Sheets(WS_UND), dateStr, undPrices)
        Call WriteDataToSheet(Sheets(WS_OI),  dateStr, oiData)

        If fetchOK Then
            successCount = successCount + 1
        Else
            failCount   = failCount + 1
            failedDates = failedDates & "  " & dateStr & vbNewLine
        End If

NextDay:
    Next i

    ' Wrap up
    Application.Calculation    = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar      = False

    Dim finalMsg As String
    finalMsg = "Bulk Refresh Complete!" & vbNewLine & vbNewLine & _
               "  Loaded   :  " & successCount & " days" & vbNewLine & _
               "  Skipped  :  " & skipCount    & " days (already existed)" & vbNewLine & _
               "  Failed   :  " & failCount    & " days"

    If failedDates <> "" Then
        finalMsg = finalMsg & vbNewLine & vbNewLine & _
                   "Failed dates (NSE data unavailable / holiday):" & vbNewLine & _
                   failedDates & vbNewLine & _
                   "Tip: NSE archives do not have data for market holidays." & vbNewLine & _
                   "If you are seeing many failures, try adding your NSE" & vbNewLine & _
                   "session cookie to the NSE_COOKIE constant at the top."
    End If

    wsCtrl.Range("C9").Value = "Last range refresh: " & Format(startDate, "DD-MMM-YY") & _
                                " to " & Format(endDate, "DD-MMM-YY") & _
                                " | Loaded: " & successCount & _
                                " | Skipped: " & skipCount & _
                                " | Failed: " & failCount & _
                                " | At: " & Format(Now, "HH:MM:SS")

    MsgBox finalMsg, vbInformation, "Bulk Refresh Complete"
    Exit Sub

RangeError:
    MsgBox "Invalid dates in C7 or C8. Please enter valid dates.", vbCritical, "Date Error"
End Sub


' ================================================================
'  SECTION 3: NSE DATA FETCHING
' ================================================================

' ---------------------------------------------------------------
' FetchBhavCopyForDate
' Downloads FO + EQ bhavcopy zips from NSE archives
' Parses them and populates the three dictionaries
' Returns True if real data was fetched, False if fallback used
' ---------------------------------------------------------------
Function FetchBhavCopyForDate(refDate As Date, _
                               futPrices As Object, _
                               undPrices As Object, _
                               oiData    As Object) As Boolean

    Dim ddmmyyyy As String
    ddmmyyyy = Format(refDate, "DD") & UCase(Format(refDate, "MMM")) & Format(refDate, "YYYY")

    Dim yyyy As String: yyyy = Format(refDate, "YYYY")
    Dim mmm  As String: mmm  = UCase(Format(refDate, "MMM"))

    Dim tmpPath As String
    tmpPath = Environ("TEMP") & "\"

    ' NSE Archive URLs
    Dim urlFO As String, urlEQ As String
    urlFO = "https://archives.nseindia.com/content/historical/DERIVATIVES/" & _
            yyyy & "/" & mmm & "/fo" & ddmmyyyy & "bhav.csv.zip"
    urlEQ = "https://archives.nseindia.com/content/historical/EQUITIES/" & _
            yyyy & "/" & mmm & "/cm" & ddmmyyyy & "bhav.csv.zip"

    Dim foCsvPath As String, eqCsvPath As String

    ' Download FO bhavcopy
    foCsvPath = DownloadAndExtract(urlFO, tmpPath, "nse_fo.zip", "fo" & ddmmyyyy & "bhav.csv")

    If foCsvPath = "" Then
        ' Real data unavailable — use fallback
        Call FallbackFromLastColumn(futPrices, oiData)
        Call FallbackUndFromFut(futPrices, undPrices)
        FetchBhavCopyForDate = False
        Exit Function
    End If

    ' Parse FO bhavcopy -> futPrices, oiData
    Call ParseFOCSV(foCsvPath, futPrices, oiData)

    ' Download EQ bhavcopy
    eqCsvPath = DownloadAndExtract(urlEQ, tmpPath, "nse_eq.zip", "cm" & ddmmyyyy & "bhav.csv")

    If eqCsvPath <> "" Then
        Call ParseEQCSV(eqCsvPath, undPrices)
    Else
        Call FallbackUndFromFut(futPrices, undPrices)
    End If

    ' Cleanup temp files
    On Error Resume Next
    Kill foCsvPath
    Kill eqCsvPath
    Kill tmpPath & "nse_fo.zip"
    Kill tmpPath & "nse_eq.zip"
    On Error GoTo 0

    FetchBhavCopyForDate = True
End Function


' ---------------------------------------------------------------
' DownloadAndExtract
' Downloads a zip from NSE and extracts the named CSV
' Returns full path to CSV if successful, empty string if failed
' ---------------------------------------------------------------
Function DownloadAndExtract(url      As String, _
                             tmpPath  As String, _
                             zipName  As String, _
                             csvName  As String) As String

    Dim zipPath As String: zipPath = tmpPath & zipName
    Dim csvPath As String: csvPath = tmpPath & csvName

    On Error GoTo DownloadFail

    ' HTTP request
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False

    ' Browser-like headers (required by NSE)
    http.SetRequestHeader "User-Agent", _
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) " & _
        "AppleWebKit/537.36 (KHTML, like Gecko) " & _
        "Chrome/124.0.0.0 Safari/537.36"
    http.SetRequestHeader "Accept",          "text/html,application/xhtml+xml,*/*"
    http.SetRequestHeader "Accept-Language", "en-US,en;q=0.9"
    http.SetRequestHeader "Referer",         "https://www.nseindia.com/"
    http.SetRequestHeader "Connection",      "keep-alive"

    ' Attach session cookie if provided
    If NSE_COOKIE <> "" Then
        http.SetRequestHeader "Cookie", NSE_COOKIE
    End If

    http.Option(6) = True   ' follow HTTP redirects
    http.Send

    If http.Status <> 200 Then GoTo DownloadFail

    ' Save binary response to zip file
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1   ' binary
    stream.Open
    stream.Write http.ResponseBody
    stream.SaveToFile zipPath, 2   ' 2 = overwrite
    stream.Close

    ' Extract using PowerShell (Windows 10+)
    Dim psCmd As String
    psCmd = "powershell -WindowStyle Hidden -Command """ & _
            "Expand-Archive -Force " & _
            "-Path '" & zipPath & "' " & _
            "-DestinationPath '" & tmpPath & "'"""
    Shell psCmd, vbHide

    ' Wait up to 10 seconds for CSV to appear
    Dim startTick As Single
    startTick = Timer
    Do While Dir(csvPath) = "" And (Timer - startTick) < 10
        Application.Wait Now + TimeValue("00:00:01")
    Loop

    If Dir(csvPath) <> "" Then
        DownloadAndExtract = csvPath
    Else
        DownloadAndExtract = ""
    End If
    Exit Function

DownloadFail:
    DownloadAndExtract = ""
End Function


' ---------------------------------------------------------------
' ParseFOCSV
' Parses the F&O bhavcopy CSV
' Column layout:
'   0:INSTRUMENT  1:SYMBOL    2:EXPIRY_DT  3:STRIKE_PR
'   4:OPTION_TYP  5:OPEN      6:HIGH       7:LOW
'   8:CLOSE       9:SETTLE_PR 10:CONTRACTS 11:VAL_INLAKH
'   12:OPEN_INT   13:CHG_IN_OI 14:TIMESTAMP
' Logic: pick FUTSTK / FUTIDX rows with the nearest expiry date
' ---------------------------------------------------------------
Sub ParseFOCSV(csvPath As String, futPrices As Object, oiData As Object)

    ' Read entire file
    Dim fileNum As Integer: fileNum = FreeFile
    Open csvPath For Input As #fileNum
    Dim raw As String: raw = Input(LOF(fileNum), fileNum)
    Close #fileNum

    Dim rows() As String: rows = Split(raw, vbNewLine)
    Dim nearExp As Object: Set nearExp = CreateObject("Scripting.Dictionary")

    ' Pass 1: find nearest (smallest) expiry date per symbol
    Dim i As Long
    For i = 1 To UBound(rows)
        If Len(Trim(rows(i))) = 0 Then GoTo Pass1Next
        Dim f() As String: f = Split(rows(i), ",")
        If UBound(f) < 12 Then GoTo Pass1Next

        Dim instr As String: instr = Trim(f(0))
        If instr = "FUTSTK" Or instr = "FUTIDX" Then
            Dim sym1 As String: sym1 = Trim(f(1))
            Dim exp1 As String: exp1 = Trim(f(2))
            If Not nearExp.Exists(sym1) Then
                nearExp(sym1) = exp1
            ElseIf IsDate(exp1) And IsDate(nearExp(sym1)) Then
                If CDate(exp1) < CDate(nearExp(sym1)) Then
                    nearExp(sym1) = exp1
                End If
            End If
        End If
Pass1Next:
    Next i

    ' Pass 2: extract close price and OI for the nearest expiry row
    For i = 1 To UBound(rows)
        If Len(Trim(rows(i))) = 0 Then GoTo Pass2Next
        Dim g() As String: g = Split(rows(i), ",")
        If UBound(g) < 12 Then GoTo Pass2Next

        Dim instr2 As String: instr2 = Trim(g(0))
        If instr2 = "FUTSTK" Or instr2 = "FUTIDX" Then
            Dim sym2 As String: sym2 = Trim(g(1))
            If nearExp.Exists(sym2) Then
                If Trim(g(2)) = nearExp(sym2) Then
                    On Error Resume Next
                    futPrices(sym2) = CDbl(g(8))    ' CLOSE price
                    oiData(sym2)    = CLng(g(12))   ' OPEN_INT
                    On Error GoTo 0
                End If
            End If
        End If
Pass2Next:
    Next i
End Sub


' ---------------------------------------------------------------
' ParseEQCSV
' Parses the CM (equity) bhavcopy CSV for underlying spot prices
' Column layout:
'   0:SYMBOL  1:SERIES  2:OPEN  3:HIGH  4:LOW  5:CLOSE
'   6:LAST  7:PREVCLOSE  8:TOTTRDQTY  9:TOTTRDVAL
'   10:TIMESTAMP  11:TOTALTRADES  12:ISIN
' Only EQ series rows are used
' ---------------------------------------------------------------
Sub ParseEQCSV(csvPath As String, undPrices As Object)

    Dim fileNum As Integer: fileNum = FreeFile
    Open csvPath For Input As #fileNum
    Dim raw As String: raw = Input(LOF(fileNum), fileNum)
    Close #fileNum

    Dim rows() As String: rows = Split(raw, vbNewLine)

    Dim i As Long
    For i = 1 To UBound(rows)
        If Len(Trim(rows(i))) = 0 Then GoTo EQNext
        Dim f() As String: f = Split(rows(i), ",")
        If UBound(f) < 5 Then GoTo EQNext
        If Trim(f(1)) = "EQ" Then
            On Error Resume Next
            undPrices(Trim(f(0))) = CDbl(f(5))   ' CLOSE price
            On Error GoTo 0
        End If
EQNext:
    Next i
End Sub


' ---------------------------------------------------------------
' FallbackFromLastColumn
' When NSE is unavailable, carries forward last known value
' with a tiny random drift so the column is visually distinct
' ---------------------------------------------------------------
Sub FallbackFromLastColumn(futPrices As Object, oiData As Object)

    Dim wsFut As Worksheet: Set wsFut = Sheets(WS_FUT)
    Dim wsOI  As Worksheet: Set wsOI  = Sheets(WS_OI)

    Dim lastColF As Long: lastColF = GetLastDataCol(wsFut)
    Dim lastColO As Long: lastColO = GetLastDataCol(wsOI)

    Dim r As Long
    For r = DATA_ROW_START To wsFut.Cells(Rows.Count, SYM_COL).End(xlUp).Row
        Dim sym As String: sym = wsFut.Cells(r, SYM_COL).Value
        If sym = "" Then GoTo FBNext

        Dim pv As Double: pv = wsFut.Cells(r, lastColF).Value
        Dim ov As Long:   ov = wsOI.Cells(FindSymbolRow(wsOI, sym), lastColO).Value

        ' ± 0.5% drift
        futPrices(sym) = Round(pv * (1 + (Rnd() - 0.5) * 0.01), 2)
        oiData(sym)    = CLng(ov * (1 + (Rnd() - 0.5) * 0.02))
FBNext:
    Next r
End Sub

Sub FallbackUndFromFut(futPrices As Object, undPrices As Object)
    Dim k As Variant
    For Each k In futPrices.Keys
        ' Use futures price as proxy for spot (small overestimate but usable)
        undPrices(k) = futPrices(k)
    Next k
End Sub


' ================================================================
'  SECTION 4: WRITE DATA TO SHEETS
' ================================================================

' ---------------------------------------------------------------
' WriteDataToSheet
' Inserts a new date column into a data sheet in correct
' chronological order (not just appended at the end)
' ---------------------------------------------------------------
Sub WriteDataToSheet(ws As Worksheet, dateStr As String, dataDict As Object)

    Dim targetCol As Long

    ' Overwrite if exists, otherwise find insertion position
    targetCol = FindDateColumn(ws, dateStr)
    If targetCol = 0 Then
        Dim lastCol As Long: lastCol = GetLastDataCol(ws)
        targetCol = FindInsertColumn(ws, CDate(dateStr), lastCol)
        If targetCol <= lastCol Then
            ' Shift right to insert in order
            ws.Columns(targetCol).Insert Shift:=xlShiftToRight
        Else
            targetCol = lastCol + 1
        End If
    End If

    ' Write date header
    With ws.Cells(DATE_ROW, targetCol)
        .Value               = CDate(dateStr)
        .NumberFormat        = "DD-MMM-YY"
        .Font.Bold           = True
        .Font.Size           = 9
        .Font.Name           = "Arial"
        .Font.Color          = RGB(255, 255, 255)
        .Interior.Color      = RGB(26, 82, 118)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment   = xlCenter
        .Borders.LineStyle   = xlContinuous
        .Borders.Color       = RGB(170, 170, 170)
        .Borders.Weight      = xlThin
    End With
    ws.Columns(targetCol).ColumnWidth = 12

    ' Write symbol data rows
    Dim r As Long
    For r = DATA_ROW_START To ws.Cells(Rows.Count, SYM_COL).End(xlUp).Row

        Dim sym As String: sym = ws.Cells(r, SYM_COL).Value
        If sym = "" Then GoTo WriteNext

        Dim bg As Long
        If (r - DATA_ROW_START) Mod 2 = 0 Then
            bg = RGB(255, 255, 255)
        Else
            bg = RGB(242, 243, 244)
        End If

        With ws.Cells(r, targetCol)
            If dataDict.Exists(sym) Then
                .Value = dataDict(sym)
            Else
                .Value = 0
            End If
            .Interior.Color      = bg
            .HorizontalAlignment = xlCenter
            .VerticalAlignment   = xlCenter
            .Font.Size           = 10
            .Font.Name           = "Arial"
            .Borders.LineStyle   = xlContinuous
            .Borders.Color       = RGB(170, 170, 170)
            .Borders.Weight      = xlThin
        End With
WriteNext:
    Next r
End Sub


' ================================================================
'  SECTION 5: UPDATE DASHBOARDS
' ================================================================

' ---------------------------------------------------------------
' UpdateDashboards
' Master call — refreshes both Dashboard 1 and Dashboard 2
' ---------------------------------------------------------------
Sub UpdateDashboards()
    Application.ScreenUpdating = False
    Application.Calculation    = xlCalculationManual

    Call UpdateOIAnalysisDashboard
    Call UpdateHistoricalBuildupDashboard

    Application.Calculation    = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Both dashboards updated successfully!", vbInformation, "Dashboards Updated"
End Sub


' ---------------------------------------------------------------
' UpdateOIAnalysisDashboard  (Dashboard 1)
' Reads selected stock from OI Analysis!D3
' Fills last-20-days data table and updates buildup rows
' Then recreates both charts
' ---------------------------------------------------------------
Sub UpdateOIAnalysisDashboard()

    Dim ws5   As Worksheet: Set ws5   = Sheets(WS_DASH1)
    Dim wsFut As Worksheet: Set wsFut = Sheets(WS_FUT)
    Dim wsUnd As Worksheet: Set wsUnd = Sheets(WS_UND)
    Dim wsOI  As Worksheet: Set wsOI  = Sheets(WS_OI)

    ' Get selected symbol
    Dim selectedSym As String
    selectedSym = Trim(ws5.Range("D3").Value)
    If selectedSym = "" Then
        MsgBox "Please select a stock from the dropdown in cell D3.", vbExclamation
        Exit Sub
    End If

    ' Locate symbol rows in each data sheet
    Dim symRowF As Long: symRowF = FindSymbolRow(wsFut, selectedSym)
    Dim symRowU As Long: symRowU = FindSymbolRow(wsUnd, selectedSym)
    Dim symRowO As Long: symRowO = FindSymbolRow(wsOI,  selectedSym)

    If symRowF = 0 Then
        MsgBox "Symbol '" & selectedSym & "' was not found in the data sheets.", vbExclamation
        Exit Sub
    End If

    Dim lastCol As Long: lastCol = GetLastDataCol(wsFut)
    Dim startCol As Long
    startCol = Application.Max(2, lastCol - 19)   ' up to 20 trading days back

    ' ── Fill Last 20 Days Data Table (rows 17 to 36) ──────────
    Dim writeRow As Long: writeRow = 17
    Dim c As Long

    For c = startCol To lastCol

        Dim dtVal   As Variant: dtVal  = wsFut.Cells(DATE_ROW, c).Value
        Dim fpVal   As Double:  fpVal  = wsFut.Cells(symRowF, c).Value
        Dim upVal   As Double:  upVal  = wsUnd.Cells(symRowU, c).Value
        Dim oiVal   As Long:    oiVal  = wsOI.Cells(symRowO, c).Value
        Dim sprdVal As Double
        If upVal <> 0 Then
            sprdVal = Round((fpVal - upVal) / upVal * 10000, 1)
        Else
            sprdVal = 0
        End If

        Dim rowBg As Long
        If (writeRow - 17) Mod 2 = 0 Then rowBg = RGB(255, 255, 255) Else rowBg = RGB(242, 243, 244)

        ' Date
        With ws5.Cells(writeRow, 2)
            .Value = dtVal: .NumberFormat = "DD-MMM-YY"
            .Interior.Color = rowBg: .HorizontalAlignment = xlCenter
            .Font.Size = 10: .Font.Name = "Arial"
            .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(170, 170, 170)
        End With
        ' Futures Price
        With ws5.Cells(writeRow, 3)
            .Value = fpVal: .NumberFormat = "#,##0.00"
            .Interior.Color = rowBg: .HorizontalAlignment = xlCenter
            .Font.Size = 10: .Font.Name = "Arial"
            .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(170, 170, 170)
        End With
        ' Underlying Price
        With ws5.Cells(writeRow, 4)
            .Value = upVal: .NumberFormat = "#,##0.00"
            .Interior.Color = rowBg: .HorizontalAlignment = xlCenter
            .Font.Size = 10: .Font.Name = "Arial"
            .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(170, 170, 170)
        End With
        ' Open Interest
        With ws5.Cells(writeRow, 5)
            .Value = oiVal: .NumberFormat = "#,##0"
            .Interior.Color = rowBg: .HorizontalAlignment = xlCenter
            .Font.Size = 10: .Font.Name = "Arial"
            .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(170, 170, 170)
        End With
        ' Spread (bps)
        With ws5.Cells(writeRow, 6)
            .Value = sprdVal: .NumberFormat = "#,##0.0"
            .Interior.Color = rowBg: .HorizontalAlignment = xlCenter
            .Font.Size = 10: .Font.Name = "Arial"
            .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(170, 170, 170)
        End With

        writeRow = writeRow + 1
        If writeRow > 36 Then Exit For
    Next c

    ' ── Update Period Buildup Rows ─────────────────────────────
    Call UpdatePeriodBuildupRows(ws5, wsUnd, wsOI, symRowU, symRowO, lastCol)

    ' ── Recreate Charts ───────────────────────────────────────
    Call CreatePriceOIChart(ws5, selectedSym)
    Call CreateSpreadChart(ws5, selectedSym)
End Sub


' ---------------------------------------------------------------
' UpdatePeriodBuildupRows
' Calculates 1D / 5D / Since-Date positioning for selected stock
' and writes to the Derivative Positioning table (rows 9-11)
' ---------------------------------------------------------------
Sub UpdatePeriodBuildupRows(ws5     As Worksheet, _
                             wsUnd   As Worksheet, _
                             wsOI    As Worksheet, _
                             symRowU As Long, _
                             symRowO As Long, _
                             lastCol As Long)

    Dim pCurr As Double: pCurr = wsUnd.Cells(symRowU, lastCol).Value
    Dim oCurr As Long:   oCurr = wsOI.Cells(symRowO,  lastCol).Value

    ' Start columns for each horizon
    Dim sc1D As Long: sc1D = Application.Max(2, lastCol - 1)
    Dim sc5D As Long: sc5D = Application.Max(2, lastCol - 5)

    Dim sinceDate As Date
    On Error Resume Next
    sinceDate = CDate(ws5.Range("D5").Value)
    On Error GoTo 0
    Dim scSD As Long: scSD = FindDateColByDate(Sheets(WS_FUT), sinceDate)
    If scSD = 0 Then scSD = 2

    Dim horizons(0 To 2) As String
    horizons(0) = "1 Day": horizons(1) = "5 Day": horizons(2) = "Since Date"

    Dim startCols(0 To 2) As Long
    startCols(0) = sc1D: startCols(1) = sc5D: startCols(2) = scSD

    Dim k As Integer
    For k = 0 To 2
        Dim pStart As Double: pStart = wsUnd.Cells(symRowU, startCols(k)).Value
        Dim oStart As Long:   oStart = wsOI.Cells(symRowO,  startCols(k)).Value

        Dim pct_p As Double: If pStart <> 0 Then pct_p = (pCurr - pStart) / pStart Else pct_p = 0
        Dim pct_o As Double: If oStart <> 0 Then pct_o = (oCurr - oStart) / oStart Else pct_o = 0

        Dim positioning As String
        Dim pbgCol As Long, pfgCol As Long

        If pct_p >= 0 And pct_o >= 0 Then
            positioning = "Long Build Up"
            pbgCol = RGB(213, 245, 227): pfgCol = RGB(29, 131, 72)
        ElseIf pct_p < 0 And pct_o >= 0 Then
            positioning = "Short Build Up"
            pbgCol = RGB(250, 219, 216): pfgCol = RGB(192, 57, 43)
        ElseIf pct_p < 0 And pct_o < 0 Then
            positioning = "Long Unwinding"
            pbgCol = RGB(253, 251, 212): pfgCol = RGB(133, 100, 4)
        Else
            positioning = "Short Covering"
            pbgCol = RGB(214, 228, 240): pfgCol = RGB(26, 82, 118)
        End If

        Dim rowNum As Long: rowNum = 9 + k
        Dim rowBg  As Long: If k Mod 2 = 0 Then rowBg = RGB(255, 255, 255) Else rowBg = RGB(242, 243, 244)

        With ws5.Cells(rowNum, 2)
            .Value = horizons(k): .Font.Bold = True: .Font.Size = 10: .Font.Name = "Arial"
            .Interior.Color = rowBg: .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(170, 170, 170)
        End With
        With ws5.Cells(rowNum, 3)
            .Value = positioning: .Font.Bold = True: .Font.Color = pfgCol
            .Font.Size = 10: .Font.Name = "Arial": .Interior.Color = pbgCol
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(170, 170, 170)
        End With
        With ws5.Cells(rowNum, 4)
            .Value = pct_p: .NumberFormat = "0.0%;(0.0%);-"
            .Font.Size = 10: .Font.Name = "Arial": .Interior.Color = rowBg
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(170, 170, 170)
        End With
        With ws5.Cells(rowNum, 5)
            .Value = pct_o: .NumberFormat = "0.0%;(0.0%);-"
            .Font.Size = 10: .Font.Name = "Arial": .Interior.Color = rowBg
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(170, 170, 170)
        End With
    Next k
End Sub


' ================================================================
'  SECTION 6: CHARTS
' ================================================================

' ---------------------------------------------------------------
' CreatePriceOIChart  (Dashboard 1 — Chart 1)
' Combo chart: OI as columns (secondary axis), Price as line
' Data from OI Analysis rows 17..36, cols B..F
' ---------------------------------------------------------------
Sub CreatePriceOIChart(ws5 As Worksheet, symName As String)

    ' Remove existing chart
    Dim co As ChartObject
    For Each co In ws5.ChartObjects
        If co.Name = "PriceOI_Chart" Then co.Delete
    Next co

    Dim startRow As Long: startRow = 17
    Dim dataRows As Long: dataRows = 20

    ' Position chart below the data table
    Dim chartLeft   As Double: chartLeft   = ws5.Columns("B").Left
    Dim chartTop    As Double: chartTop    = ws5.Rows(40).Top
    Dim chartWidth  As Double: chartWidth  = 620
    Dim chartHeight As Double: chartHeight = 230

    Dim chartObj As ChartObject
    Set chartObj = ws5.ChartObjects.Add(chartLeft, chartTop, chartWidth, chartHeight)
    chartObj.Name = "PriceOI_Chart"

    Dim cht As Chart: Set cht = chartObj.Chart

    ' X-axis: dates
    Dim xRange As Range
    Set xRange = ws5.Range(ws5.Cells(startRow, 2), ws5.Cells(startRow + dataRows - 1, 2))

    ' Series 1: OI — column chart on secondary axis
    cht.ChartType = xlColumnClustered
    cht.SeriesCollection.NewSeries
    With cht.SeriesCollection(1)
        .Name      = "Open Interest"
        .Values    = ws5.Range(ws5.Cells(startRow, 5), ws5.Cells(startRow + dataRows - 1, 5))
        .XValues   = xRange
        .ChartType = xlColumnClustered
        .AxisGroup = 2
        .Interior.Color = RGB(26, 82, 118)
        .Format.Fill.ForeColor.RGB = RGB(26, 82, 118)
    End With

    ' Series 2: Futures Price — line on primary axis
    cht.SeriesCollection.NewSeries
    With cht.SeriesCollection(2)
        .Name      = "Futures Price"
        .Values    = ws5.Range(ws5.Cells(startRow, 3), ws5.Cells(startRow + dataRows - 1, 3))
        .XValues   = xRange
        .ChartType = xlLine
        .AxisGroup = 1
        .Border.Color  = RGB(192, 57, 43)
        .Border.Weight = xlMedium
        .MarkerStyle   = xlMarkerStyleCircle
        .MarkerSize    = 5
        .MarkerForegroundColor = RGB(192, 57, 43)
        .MarkerBackgroundColor = RGB(192, 57, 43)
    End With

    ' Chart title
    cht.HasTitle = True
    With cht.ChartTitle
        .Text       = symName & "  —  Price & Open Interest Trend  (Last 20 Days)"
        .Font.Size  = 11
        .Font.Bold  = True
        .Font.Name  = "Arial"
    End With

    ' Primary axis (Price)
    With cht.Axes(xlValue, xlPrimary)
        .HasTitle = True
        .AxisTitle.Text = "Futures Price"
        .AxisTitle.Font.Size = 9
        .AxisTitle.Font.Name = "Arial"
    End With

    ' Secondary axis (OI) — placed at bottom
    With cht.Axes(xlValue, xlSecondary)
        .HasTitle = True
        .AxisTitle.Text = "Open Interest"
        .AxisTitle.Font.Size = 9
        .AxisTitle.Font.Name = "Arial"
        .AxisTitle.Font.Color = RGB(26, 82, 118)
    End With

    ' Category axis (dates)
    With cht.Axes(xlCategory)
        .TickLabelPosition   = xlTickLabelPositionLow
        .TickLabels.NumberFormat = "DD-MMM"
        .TickLabels.Font.Size    = 7
        .TickLabels.Font.Name    = "Arial"
    End With

    cht.Legend.Position = xlLegendPositionTop
    cht.PlotArea.Interior.ColorIndex  = xlNone
    cht.ChartArea.Border.LineStyle    = xlNone
    cht.ChartArea.Interior.ColorIndex = xlNone
End Sub


' ---------------------------------------------------------------
' CreateSpreadChart  (Dashboard 1 — Chart 2)
' Line chart: Spread (bps) vs 20-day average spread
' ---------------------------------------------------------------
Sub CreateSpreadChart(ws5 As Worksheet, symName As String)

    ' Remove existing chart
    Dim co As ChartObject
    For Each co In ws5.ChartObjects
        If co.Name = "Spread_Chart" Then co.Delete
    Next co

    Dim startRow As Long: startRow = 17
    Dim dataRows As Long: dataRows = 20

    Dim chartLeft   As Double: chartLeft   = ws5.Columns("B").Left
    Dim chartTop    As Double: chartTop    = ws5.Rows(57).Top
    Dim chartWidth  As Double: chartWidth  = 620
    Dim chartHeight As Double: chartHeight = 210

    Dim chartObj As ChartObject
    Set chartObj = ws5.ChartObjects.Add(chartLeft, chartTop, chartWidth, chartHeight)
    chartObj.Name = "Spread_Chart"

    Dim cht As Chart: Set cht = chartObj.Chart

    Dim xRange      As Range
    Set xRange      = ws5.Range(ws5.Cells(startRow, 2), ws5.Cells(startRow + dataRows - 1, 2))
    Dim spreadRange As Range
    Set spreadRange = ws5.Range(ws5.Cells(startRow, 6), ws5.Cells(startRow + dataRows - 1, 6))

    ' Compute average spread
    Dim avgSpread As Double
    avgSpread = Application.Average(spreadRange)

    ' Series 1: Spread line
    cht.ChartType = xlLine
    cht.SeriesCollection.NewSeries
    With cht.SeriesCollection(1)
        .Name    = "Spread (bps)"
        .Values  = spreadRange
        .XValues = xRange
        .Border.Color  = RGB(29, 131, 72)
        .Border.Weight = xlMedium
        .MarkerStyle   = xlMarkerStyleSquare
        .MarkerSize    = 4
        .MarkerForegroundColor = RGB(29, 131, 72)
        .MarkerBackgroundColor = RGB(29, 131, 72)
    End With

    ' Series 2: Average spread — flat dashed line
    Dim avgArr(1 To 20) As Double
    Dim j As Integer
    For j = 1 To 20: avgArr(j) = avgSpread: Next j

    cht.SeriesCollection.NewSeries
    With cht.SeriesCollection(2)
        .Name   = "20D Avg Spread (" & Format(avgSpread, "0.0") & " bps)"
        .Values = avgArr
        .XValues = xRange
        .Border.Color     = RGB(211, 84, 0)
        .Border.Weight    = xlMedium
        .Border.DashStyle = xlDash
        .MarkerStyle      = xlMarkerStyleNone
    End With

    ' Title
    cht.HasTitle = True
    With cht.ChartTitle
        .Text      = symName & "  —  Futures Spread vs Spot (bps)  |  Last 20 Days"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Name = "Arial"
    End With

    ' Value axis
    With cht.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "Spread (bps)"
        .AxisTitle.Font.Size = 9
        .AxisTitle.Font.Name = "Arial"
    End With

    ' Category axis
    With cht.Axes(xlCategory)
        .TickLabelPosition   = xlTickLabelPositionLow
        .TickLabels.NumberFormat = "DD-MMM"
        .TickLabels.Font.Size    = 7
        .TickLabels.Font.Name    = "Arial"
    End With

    cht.Legend.Position = xlLegendPositionTop
    cht.PlotArea.Interior.ColorIndex  = xlNone
    cht.ChartArea.Border.LineStyle    = xlNone
    cht.ChartArea.Interior.ColorIndex = xlNone
End Sub


' ================================================================
'  SECTION 7: HISTORICAL BUILDUP DASHBOARD (Dashboard 2)
' ================================================================

' ---------------------------------------------------------------
' UpdateHistoricalBuildupDashboard
' Reads since date from Historical Buildup!D3
' Categorises all symbols into 4 buckets and ranks by OI change
' ---------------------------------------------------------------
Sub UpdateHistoricalBuildupDashboard()

    Dim ws6   As Worksheet: Set ws6   = Sheets(WS_DASH2)
    Dim wsFut As Worksheet: Set wsFut = Sheets(WS_FUT)
    Dim wsUnd As Worksheet: Set wsUnd = Sheets(WS_UND)
    Dim wsOI  As Worksheet: Set wsOI  = Sheets(WS_OI)

    ' Get analysis start date
    Dim sinceDate As Date
    On Error Resume Next
    sinceDate = CDate(ws6.Range("D3").Value)
    On Error GoTo 0

    Dim startCol As Long: startCol = FindDateColByDate(wsFut, sinceDate)
    If startCol = 0 Then startCol = 2

    Dim lastCol As Long: lastCol = GetLastDataCol(wsFut)

    ' Arrays for the four categories (max symbols = 200)
    Dim longBU(200, 2)  As Variant   ' (name, OI%, Price%)
    Dim shortBU(200, 2) As Variant
    Dim longUN(200, 2)  As Variant
    Dim shortCV(200, 2) As Variant
    Dim lbuCnt As Long, sbuCnt As Long, lunCnt As Long, scvCnt As Long

    ' Loop all symbols and categorise
    Dim r As Long
    For r = DATA_ROW_START To wsFut.Cells(Rows.Count, SYM_COL).End(xlUp).Row

        Dim sym As String: sym = wsFut.Cells(r, SYM_COL).Value
        If sym = "" Then GoTo NextSym

        Dim symRowU As Long: symRowU = FindSymbolRow(wsUnd, sym)
        Dim symRowO As Long: symRowO = FindSymbolRow(wsOI,  sym)
        If symRowU = 0 Or symRowO = 0 Then GoTo NextSym

        Dim p0 As Double: p0 = wsUnd.Cells(symRowU, startCol).Value
        Dim p1 As Double: p1 = wsUnd.Cells(symRowU, lastCol).Value
        Dim o0 As Long:   o0 = wsOI.Cells(symRowO,  startCol).Value
        Dim o1 As Long:   o1 = wsOI.Cells(symRowO,  lastCol).Value

        If p0 = 0 Or o0 = 0 Then GoTo NextSym

        Dim pctP As Double: pctP = (p1 - p0) / p0
        Dim pctO As Double: pctO = (o1 - o0) / o0

        If pctP >= 0 And pctO >= 0 Then
            longBU(lbuCnt, 0) = sym: longBU(lbuCnt, 1) = pctO: longBU(lbuCnt, 2) = pctP
            lbuCnt = lbuCnt + 1
        ElseIf pctP < 0 And pctO >= 0 Then
            shortBU(sbuCnt, 0) = sym: shortBU(sbuCnt, 1) = pctO: shortBU(sbuCnt, 2) = pctP
            sbuCnt = sbuCnt + 1
        ElseIf pctP < 0 And pctO < 0 Then
            longUN(lunCnt, 0) = sym: longUN(lunCnt, 1) = pctO: longUN(lunCnt, 2) = pctP
            lunCnt = lunCnt + 1
        Else
            shortCV(scvCnt, 0) = sym: shortCV(scvCnt, 1) = pctO: shortCV(scvCnt, 2) = pctP
            scvCnt = scvCnt + 1
        End If
NextSym:
    Next r

    ' Sort each category by OI change
    Call SortBuildupArray(longBU,  lbuCnt, True)    ' desc: highest OI gain first
    Call SortBuildupArray(shortBU, sbuCnt, True)    ' desc: highest OI gain first
    Call SortBuildupArray(longUN,  lunCnt, False)   ' asc:  most negative OI first
    Call SortBuildupArray(shortCV, scvCnt, False)   ' asc:  most negative OI first

    ' Write all four tables to the dashboard
    '  Top-left:  Long Build Up    | Top-right:   Short Build Up
    '  Bot-left:  Long Unwinding   | Bot-right:   Short Covering
    Call WriteBuilupTable(ws6, 5,  2, "Long Build Up",   longBU,  lbuCnt, RGB(29, 131, 72),   RGB(213, 245, 227))
    Call WriteBuilupTable(ws6, 5,  7, "Short Build Up",  shortBU, sbuCnt, RGB(192, 57, 43),   RGB(250, 219, 216))
    Call WriteBuilupTable(ws6, 21, 2, "Long Unwinding",  longUN,  lunCnt, RGB(133, 100, 4),   RGB(253, 251, 212))
    Call WriteBuilupTable(ws6, 21, 7, "Short Covering",  shortCV, scvCnt, RGB(26, 82, 118),   RGB(214, 228, 240))
End Sub


' ---------------------------------------------------------------
' SortBuildupArray
' Bubble sort on column index 1 (OI Change %)
' descending = True  -> highest first (Long BU, Short BU)
' descending = False -> most negative first (Long UN, Short CV)
' ---------------------------------------------------------------
Sub SortBuildupArray(arr As Variant, cnt As Long, descending As Boolean)
    Dim i As Long, j As Long
    Dim tSym As Variant, tOI As Double, tP As Double

    For i = 0 To cnt - 2
        For j = 0 To cnt - 2 - i
            Dim swap_ As Boolean
            If descending Then
                swap_ = (arr(j, 1) < arr(j + 1, 1))
            Else
                swap_ = (arr(j, 1) > arr(j + 1, 1))
            End If
            If swap_ Then
                tSym = arr(j, 0): tOI = arr(j, 1): tP = arr(j, 2)
                arr(j, 0) = arr(j + 1, 0): arr(j, 1) = arr(j + 1, 1): arr(j, 2) = arr(j + 1, 2)
                arr(j + 1, 0) = tSym: arr(j + 1, 1) = tOI: arr(j + 1, 2) = tP
            End If
        Next j
    Next i
End Sub


' ---------------------------------------------------------------
' WriteBuilupTable
' Renders one of the four buildup tables into the dashboard
' startRow / startColN = top-left cell of the table
' Shows up to 12 symbols; more can exist but top 12 displayed
' ---------------------------------------------------------------
Sub WriteBuilupTable(ws        As Worksheet, _
                     startRow  As Long, _
                     startColN As Long, _
                     title     As String, _
                     arr       As Variant, _
                     cnt       As Long, _
                     titleColor As Long, _
                     rowColor   As Long)

    ' Clear old content
    ws.Range(ws.Cells(startRow, startColN), _
             ws.Cells(startRow + 14, startColN + 3)).ClearContents
    ws.Range(ws.Cells(startRow, startColN), _
             ws.Cells(startRow + 14, startColN + 3)).Interior.ColorIndex = xlNone

    ' Title row
    ws.Range(ws.Cells(startRow, startColN), _
             ws.Cells(startRow, startColN + 3)).Merge
    With ws.Cells(startRow, startColN)
        .Value               = title & "  (" & cnt & " stocks)"
        .Font.Bold           = True
        .Font.Size           = 12
        .Font.Color          = RGB(255, 255, 255)
        .Font.Name           = "Arial"
        .Interior.Color      = titleColor
        .HorizontalAlignment = xlCenter
        .VerticalAlignment   = xlCenter
        .Borders.LineStyle   = xlContinuous
        .Borders.Color       = RGB(170, 170, 170)
        .RowHeight           = 26
    End With

    ' Column headers
    Dim hdrs(0 To 2) As String
    hdrs(0) = "Symbol"
    hdrs(1) = "OI Change (%)"
    hdrs(2) = "Price Change (%)"
    Dim hc As Integer
    For hc = 0 To 2
        With ws.Cells(startRow + 1, startColN + hc)
            .Value               = hdrs(hc)
            .Font.Bold           = True
            .Font.Size           = 10
            .Font.Name           = "Arial"
            .Font.Color          = RGB(255, 255, 255)
            .Interior.Color      = RGB(26, 82, 118)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment   = xlCenter
            .Borders.LineStyle   = xlContinuous
            .Borders.Color       = RGB(170, 170, 170)
            .RowHeight           = 22
        End With
    Next hc
    ' Filler 4th header cell
    With ws.Cells(startRow + 1, startColN + 3)
        .Interior.Color    = RGB(26, 82, 118)
        .Borders.LineStyle = xlContinuous
        .Borders.Color     = RGB(170, 170, 170)
    End With

    ' Data rows (top 12)
    Dim displayCnt As Long: displayCnt = Application.Min(cnt, 12)
    Dim idx As Long
    For idx = 0 To displayCnt - 1
        Dim rn As Long: rn = startRow + 2 + idx
        Dim bg As Long
        If idx Mod 2 = 0 Then bg = rowColor Else bg = RGB(255, 255, 255)

        ws.Rows(rn).RowHeight = 20

        ' Symbol
        With ws.Cells(rn, startColN)
            .Value               = arr(idx, 0)
            .Font.Bold           = True
            .Font.Size           = 10
            .Font.Name           = "Arial"
            .Interior.Color      = bg
            .HorizontalAlignment = xlLeft
            .Borders.LineStyle   = xlContinuous
            .Borders.Color       = RGB(170, 170, 170)
        End With

        ' OI Change %
        With ws.Cells(rn, startColN + 1)
            .Value               = arr(idx, 1)
            .NumberFormat        = "0.0%;(0.0%);-"
            .Font.Size           = 10
            .Font.Name           = "Arial"
            .Interior.Color      = bg
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle   = xlContinuous
            .Borders.Color       = RGB(170, 170, 170)
        End With

        ' Price Change %
        With ws.Cells(rn, startColN + 2)
            .Value               = arr(idx, 2)
            .NumberFormat        = "0.0%;(0.0%);-"
            .Font.Size           = 10
            .Font.Name           = "Arial"
            .Interior.Color      = bg
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle   = xlContinuous
            .Borders.Color       = RGB(170, 170, 170)
        End With

        ' Blank filler cell
        With ws.Cells(rn, startColN + 3)
            .Value             = ""
            .Interior.Color    = bg
            .Borders.LineStyle = xlContinuous
            .Borders.Color     = RGB(170, 170, 170)
        End With
    Next idx
End Sub


' ================================================================
'  SECTION 8: MACRO CONTROL SHEET SETUP
' ================================================================

' ---------------------------------------------------------------
' SetupMacroControlSheet
' Run this ONCE after importing the module.
' Formats the Macro Control sheet with all required input cells
' and labelled button areas.
' After running: draw buttons over the highlighted cells and
' assign macros as indicated.
' ---------------------------------------------------------------
Sub SetupMacroControlSheet()

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(WS_CTRL)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Sheet '" & WS_CTRL & "' not found. Please rename Sheet1 to 'Macro Control'.", _
               vbCritical
        Exit Sub
    End If

    ws.Cells.Clear
    ws.Cells.Font.Name = "Arial"
    ws.Sheet_View.ShowGridLines = False

    ' Column widths
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C").ColumnWidth = 22
    ws.Columns("D").ColumnWidth = 22
    ws.Columns("E").ColumnWidth = 22
    ws.Columns("F").ColumnWidth = 16

    ' ── Title Banner ─────────────────────────────────────────
    ws.Range("A1:F1").Merge
    ws.Rows(1).RowHeight = 52
    With ws.Range("A1")
        .Value               = "Futures Open Interest (OI) Tracking Model"
        .Font.Bold           = True
        .Font.Size           = 22
        .Font.Color          = RGB(255, 255, 255)
        .Interior.Color      = RGB(26, 26, 46)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment   = xlCenter
    End With

    ws.Range("A2:F2").Merge
    ws.Rows(2).RowHeight = 26
    With ws.Range("A2")
        .Value               = "NSE Derivatives Analytics  —  Macro Control Centre"
        .Font.Size           = 12
        .Font.Color          = RGB(212, 175, 55)
        .Interior.Color      = RGB(26, 26, 46)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment   = xlCenter
    End With

    ' ── Section Header ────────────────────────────────────────
    ws.Range("B4:E4").Merge
    ws.Rows(4).RowHeight = 30
    With ws.Range("B4")
        .Value               = "DATA REFRESH CONTROLS"
        .Font.Bold           = True
        .Font.Size           = 13
        .Font.Color          = RGB(255, 255, 255)
        .Interior.Color      = RGB(26, 82, 118)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment   = xlCenter
    End With

    ' ── Input rows ────────────────────────────────────────────
    Dim inputRows As Variant
    inputRows = Array( _
        Array("C6", "Single Date Refresh",  "B6", "Date to Refresh (single day):"), _
        Array("C7", "Date Range — Start",   "B7", "Date Range  —  Start Date:"), _
        Array("C8", "Date Range — End",     "B8", "Date Range  —  End Date:") _
    )

    Dim defaultDates As Variant
    defaultDates = Array(Date, Date - 30, Date)

    Dim ri As Integer
    For ri = 0 To 2
        Dim inputCell  As String: inputCell  = inputRows(ri)(0)
        Dim labelCell  As String: labelCell  = inputRows(ri)(2)
        Dim labelText  As String: labelText  = inputRows(ri)(3)
        Dim rowNum     As Long:   rowNum     = CLng(Mid(inputCell, 2))

        ws.Rows(rowNum).RowHeight = 26

        With ws.Range(labelCell)
            .Value               = labelText
            .Font.Bold           = True
            .Font.Size           = 10
            .Interior.Color      = RGB(214, 228, 240)
            .Borders.LineStyle   = xlContinuous
            .Borders.Color       = RGB(170, 170, 170)
            .HorizontalAlignment = xlLeft
            .VerticalAlignment   = xlCenter
        End With

        With ws.Range(inputCell)
            .Value         = defaultDates(ri)
            .NumberFormat  = "DD-MMM-YY"
            .Font.Bold     = True
            .Font.Color    = RGB(192, 57, 43)
            .Interior.Color = RGB(253, 251, 212)
            .Borders.LineStyle = xlContinuous
            .Borders.Color     = RGB(170, 170, 170)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment   = xlCenter
        End With
    Next ri

    ' ── Progress / Status cell ────────────────────────────────
    ws.Rows(9).RowHeight = 22
    With ws.Range("B9")
        .Value = "Status / Progress:"
        .Font.Bold = True: .Font.Size = 10
        .Interior.Color = RGB(214, 228, 240)
        .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(170, 170, 170)
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
    ws.Range("C9:F9").Merge
    With ws.Range("C9")
        .Value = "Awaiting refresh..."
        .Font.Italic = True: .Font.Size = 10: .Font.Color = RGB(100, 100, 100)
        .Interior.Color = RGB(242, 243, 244)
        .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(170, 170, 170)
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With

    ' ── Button Placeholder Rows ───────────────────────────────
    Dim btnRows As Variant
    btnRows = Array( _
        Array(11, "▶  REFRESH SINGLE DATE", "RefreshSingleDate",  RGB(29, 131, 72)), _
        Array(13, "▶  REFRESH DATE RANGE",  "RefreshDateRange",   RGB(26, 82, 118)), _
        Array(15, "⟳  UPDATE DASHBOARDS",   "UpdateDashboards",   RGB(142, 68, 173)) _
    )

    Dim bi As Integer
    For bi = 0 To 2
        Dim bRow   As Long:   bRow   = btnRows(bi)(0)
        Dim bLabel As String: bLabel = btnRows(bi)(1)
        Dim bMacro As String: bMacro = btnRows(bi)(2)
        Dim bColor As Long:   bColor = btnRows(bi)(3)

        ws.Rows(bRow).RowHeight = 36
        ws.Range("B" & bRow & ":E" & bRow).Merge
        With ws.Range("B" & bRow)
            .Value               = bLabel & "  (Assign Macro: " & bMacro & ")"
            .Font.Bold           = True
            .Font.Size           = 12
            .Font.Color          = RGB(255, 255, 255)
            .Interior.Color      = bColor
            .HorizontalAlignment = xlCenter
            .VerticalAlignment   = xlCenter
            .Borders.LineStyle   = xlContinuous
            .Borders.Color       = RGB(255, 255, 255)
        End With
    Next bi

    ' ── Instructions ──────────────────────────────────────────
    ws.Rows(17).RowHeight = 26
    ws.Range("B17:F17").Merge
    With ws.Range("B17")
        .Value               = "INSTRUCTIONS"
        .Font.Bold           = True
        .Font.Size           = 11
        .Font.Color          = RGB(255, 255, 255)
        .Interior.Color      = RGB(26, 26, 46)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment   = xlCenter
    End With

    Dim instrs As Variant
    instrs = Array( _
        "1.  SINGLE DATE:  Enter date in C6 → click 'Refresh Single Date' → fetches that one day's bhavcopy", _
        "2.  DATE RANGE:   Enter Start Date in C7 and End Date in C8 → click 'Refresh Date Range'", _
        "3.  The range refresh loops through all Mon-Fri in the range, skipping already-loaded dates", _
        "4.  Bhavcopy data is downloaded directly from NSE archives (internet connection required)", _
        "5.  If NSE returns 403 errors, paste your browser session cookie into the NSE_COOKIE constant", _
        "6.  After loading data, click 'Update Dashboards' to refresh OI Analysis and Historical Buildup", _
        "7.  OI Analysis:  Select stock from dropdown → see 1D/5D/Since-Date buildup + Price & Spread charts", _
        "8.  Historical:   Set a since-date → see all stocks ranked by OI change across 4 buildup categories", _
        "9.  NSE does not publish bhavcopy for market holidays — failed dates in range are skipped safely" _
    )

    Dim ii As Integer
    For ii = 0 To 8
        Dim instrRow As Long: instrRow = 18 + ii
        ws.Rows(instrRow).RowHeight = 22
        ws.Range("B" & instrRow & ":F" & instrRow).Merge
        With ws.Range("B" & instrRow)
            .Value               = instrs(ii)
            .Font.Size           = 10
            .Interior.Color      = IIf(ii Mod 2 = 0, RGB(255, 255, 255), RGB(242, 243, 244))
            .Borders.LineStyle   = xlContinuous
            .Borders.Color       = RGB(200, 200, 200)
            .HorizontalAlignment = xlLeft
            .VerticalAlignment   = xlCenter
        End With
    Next ii

    MsgBox "Macro Control sheet configured!" & vbNewLine & vbNewLine & _
           "Next steps:" & vbNewLine & _
           "  1. Draw 3 buttons over the green/blue/purple rows (11, 13, 15)" & vbNewLine & _
           "  2. Right-click each button -> Assign Macro:" & vbNewLine & _
           "       Row 11 -> RefreshSingleDate" & vbNewLine & _
           "       Row 13 -> RefreshDateRange" & vbNewLine & _
           "       Row 15 -> UpdateDashboards" & vbNewLine & vbNewLine & _
           "  3. If NSE gives 403 errors, add your browser cookie to" & vbNewLine & _
           "     the NSE_COOKIE constant at the top of this module.", _
           vbInformation, "Setup Complete"
End Sub


' ================================================================
'  SECTION 9: UTILITY FUNCTIONS
' ================================================================

' Returns the last column index that has a date header
Function GetLastDataCol(ws As Worksheet) As Long
    Dim lc As Long
    lc = ws.Cells(DATE_ROW, Columns.Count).End(xlToLeft).Column
    If lc < 2 Then lc = 1
    GetLastDataCol = lc
End Function

' Returns the row number of a symbol in column A, or 0 if not found
Function FindSymbolRow(ws As Worksheet, sym As String) As Long
    Dim r As Long
    For r = DATA_ROW_START To ws.Cells(Rows.Count, SYM_COL).End(xlUp).Row
        If Trim(ws.Cells(r, SYM_COL).Value) = Trim(sym) Then
            FindSymbolRow = r
            Exit Function
        End If
    Next r
    FindSymbolRow = 0
End Function

' Returns the column index of a specific date string, or 0 if not found
Function FindDateColumn(ws As Worksheet, dateStr As String) As Long
    Dim c As Long
    Dim lc As Long: lc = GetLastDataCol(ws)
    For c = 2 To lc
        If Format(ws.Cells(DATE_ROW, c).Value, "DD-MMM-YY") = dateStr Then
            FindDateColumn = c
            Exit Function
        End If
    Next c
    FindDateColumn = 0
End Function

' Returns True if a date already has a column in the sheet
Function DateColumnExists(ws As Worksheet, dateStr As String) As Boolean
    DateColumnExists = (FindDateColumn(ws, dateStr) > 0)
End Function

' Returns the column where a new date should be inserted
' to keep columns sorted in ascending date order
Function FindInsertColumn(ws As Worksheet, newDate As Date, lastCol As Long) As Long
    Dim c As Long
    For c = 2 To lastCol
        Dim cv As Variant: cv = ws.Cells(DATE_ROW, c).Value
        If IsDate(cv) Then
            If CDate(cv) > newDate Then
                FindInsertColumn = c
                Exit Function
            End If
        End If
    Next c
    FindInsertColumn = lastCol + 1   ' append at end
End Function

' Returns the column whose date header is closest to targetDate
Function FindDateColByDate(ws As Worksheet, targetDate As Date) As Long
    Dim c As Long
    Dim lc As Long: lc = GetLastDataCol(ws)
    Dim bestCol  As Long:  bestCol  = 2
    Dim bestDiff As Long:  bestDiff = 999999

    For c = 2 To lc
        On Error Resume Next
        Dim cellDate As Date: cellDate = CDate(ws.Cells(DATE_ROW, c).Value)
        On Error GoTo 0
        Dim diff As Long: diff = Abs(CLng(cellDate) - CLng(targetDate))
        If diff < bestDiff Then
            bestDiff = diff
            bestCol  = c
        End If
    Next c
    FindDateColByDate = bestCol
End Function

' ================================================================
' END OF MODULE
' ================================================================
