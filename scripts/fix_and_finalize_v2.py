import os
import re

ORIGINAL_PATH = r"c:\AI\asagake\temp_vba_audit\AutoTraderAdvanced.bas"
OUTPUT_PATH = r"c:\AI\asagake\temp_vba_audit\AutoTraderAdvanced_v2_final.bas"
CLEAN_SETUP_PATH = r"c:\AI\asagake\scripts\clean_setup.txt"
CLEAN_ENSURE_PATH = r"c:\AI\asagake\scripts\clean_ensure_params.txt"
CLEAN_LOG_PATH = r"c:\AI\asagake\scripts\clean_log_preorder.txt"

def replace_block_by_slicing(content, pattern, replacement, name):
    match = pattern.search(content)
    if match:
        print(f"DEBUG: Found {name} at {match.start()}-{match.end()}")
        new_content = content[:match.start()] + replacement + content[match.end():]
        print(f"DEBUG: Replaced {name} using slicing.")
        return new_content
    else:
        print(f"DEBUG: {name} not found.")
        return content

def fix_and_finalize():
    if not os.path.exists(ORIGINAL_PATH):
        print(f"Error: Original file not found at {ORIGINAL_PATH}")
        return
    
    # Read original file (cp932)
    with open(ORIGINAL_PATH, "r", encoding="cp932", errors="ignore") as f:
        content = f.read()
        
    # Read clean replacement blocks (utf-8)
    with open(CLEAN_SETUP_PATH, "r", encoding="utf-8") as f:
        clean_setup_ui = f.read()

    with open(CLEAN_ENSURE_PATH, "r", encoding="utf-8") as f:
        clean_ensure_params = f.read()

    with open(CLEAN_LOG_PATH, "r", encoding="utf-8") as f:
        clean_log_preorder = f.read()

    # --- 0. CLEANUP CORRUPTION (Fix bad insertions from previous runs) ---
    bad_code_1 = "CheckSettlement ws, Nothing\n    UpdateSettlementOrders ws, Nothing()"
    bad_code_2 = "CheckSettlement ws, Nothing\n    UpdateSettlementOrders ws, Nothing"
    content = content.replace(bad_code_1, "")
    content = content.replace(bad_code_2, "")
    content = content.replace("Sub PreplaceOrdersV2()\n\n    Dim ws", "Sub PreplaceOrdersV2()\n    Dim ws")

    # --- 1. Add Constants ---
    constants_code = """
' Heartbeat & V2 Constants
Private Const HEARTBEAT_CELL_ADDR As String = "Z2"
Private Const HEARTBEAT_TIMEOUT_SEC As Long = 60
Private Const DEFAULT_UPDATE_THRESHOLD_TICK As Long = 2
"""
    if "Public Const DASH2_DATA_START" in content:
        content = content.replace("Public Const DASH2_DATA_START As Long = 6", "Public Const DASH2_DATA_START As Long = 6" + constants_code)
    elif "Private Const DASH2_DATA_START" in content:
         content = content.replace("Private Const DASH2_DATA_START As Long = 6", "Private Const DASH2_DATA_START As Long = 6" + constants_code)
    else:
        if "Option Explicit" in content:
            content = content.replace("Option Explicit", "Option Explicit" + constants_code)
        else:
            content = constants_code + content

    # --- 2. Fix EnsureParamFormulas (Slicing) ---
    pattern_ensure = re.compile(r'(Private\s+)?Sub\s+EnsureParamFormulas\s*\(.*?End\s+Sub', re.DOTALL | re.IGNORECASE)
    content = replace_block_by_slicing(content, pattern_ensure, clean_ensure_params, "EnsureParamFormulas")

    # --- 3. Fix PreplaceOrdersV2 (Insert calls AFTER Sh init) ---
    target_line = "Set Sh = EnsureOrdersSheet(ws)"
    insert_code = "\n    ' V2: Settlement & Execution Monitoring\n    CheckSettlement ws, Sh\n    UpdateSettlementOrders ws, Sh\n    CheckRssHeartbeat\n"
    
    preplace_start = content.find("Sub PreplaceOrdersV2")
    if preplace_start != -1:
        sh_init = content.find(target_line, preplace_start)
        if sh_init != -1:
            if "CheckSettlement ws, Sh" not in content[sh_init:sh_init+200]:
                content = content[:sh_init + len(target_line)] + insert_code + content[sh_init + len(target_line):]

    # --- 4. Fix PreplaceOrdersV2 (Insert ExecutePreplaceOrders) ---
    log_call = "LogPreOrder ws, Sh, r, hasBuy, hasSell, eBuyCol, eSellCol, qtyCol, tpCol, slCol, modeCol, sessionCol, tickerCol, bufferFrac, noteExtra"
    exec_call = "\n                ExecutePreplaceOrders ws, Sh, r, hasBuy, hasSell, eBuyCol, eSellCol, qtyCol, tickerCol"
    if exec_call not in content:
        content = content.replace(log_call, log_call + exec_call)

    # --- 5. Replace Corrupted LogPreOrder (Slicing) ---
    pattern_log = re.compile(r'(Private\s+)?Sub\s+LogPreOrder\s*\(.*?End\s+Sub', re.DOTALL | re.IGNORECASE)
    content = replace_block_by_slicing(content, pattern_log, clean_log_preorder, "LogPreOrder")

    # --- 6. REPLACE SetupDashboardUIV2 & CreateButton (Slicing) ---
    # First, remove existing CreateButton if present
    pattern_create_btn = re.compile(r'(Private\s+)?Sub\s+CreateButton\s*\(.*?End\s+Sub', re.DOTALL | re.IGNORECASE)
    match = pattern_create_btn.search(content)
    if match:
        content = content[:match.start()] + "" + content[match.end():]
    
    # Now replace SetupDashboardUIV2 with the combined block
    pattern_setup = re.compile(r'(Private\s+)?Sub\s+SetupDashboardUIV2\s*\(.*?End\s+Sub', re.DOTALL | re.IGNORECASE)
    match = pattern_setup.search(content)
    if match:
        print(f"DEBUG: Found SetupDashboardUIV2 at {match.start()}-{match.end()}")
        content = content[:match.start()] + clean_setup_ui + content[match.end():]
        print("DEBUG: Replaced SetupDashboardUIV2 using slicing.")
    else:
        print("DEBUG: SetupDashboardUIV2 not found. Appending...")
        content += "\n\n" + clean_setup_ui

    # --- 7. Fix ApplyJapaneseLabelsV2 (VWAP Corruption) ---
    content = content.replace('"VWA...AP"', '"VWAP"')
    content = content.replace('"VWAAP"', '"VWAP"')

    # --- 8. Remove Duplicates of New Functions ---
    funcs_to_remove = [
        "SetupHeartbeatCell", "CheckRssHeartbeat", "LogVbaError", 
        "PlaceOrderRSS", "ModifyOrderRSS", "ExecutePreplaceOrders", 
        "ManageActiveOrder", "CheckSettlement", "UpdateSettlementOrders"
    ]
    
    for func in funcs_to_remove:
        pattern = r"(Public |Private )?(Sub |Function )" + func + r"\(.*?\).*?End (Sub|Function)"
        content = re.sub(pattern, "", content, flags=re.DOTALL)

    # --- 9. Append New V2 Functions ---
    new_functions = '''
' ==========================================
' ASAGAKE V2 NEW FUNCTIONS
' ==========================================

Private Sub SetupHeartbeatCell(ByVal ws As Worksheet)
    On Error Resume Next
    With ws.Range(HEARTBEAT_CELL_ADDR)
        .ClearContents
        .Interior.Color = RGB(240, 240, 240)
        .Font.Size = 8
        .HorizontalAlignment = xlCenter
        .value = "RSS Monitor"
    End With
    On Error GoTo 0
End Sub

Public Function CheckRssHeartbeat() As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DASH2_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    
    Dim nkyPrice As Variant
    nkyPrice = ws.Cells(2, 2).value
    
    If IsError(nkyPrice) Or IsEmpty(nkyPrice) Or nkyPrice = "" Then
        LogVbaEvent "Heartbeat", "RSS N225 price is missing or error."
    End If
    
    On Error Resume Next
    ws.Range(HEARTBEAT_CELL_ADDR).value = Format$(Now, "hh:nn:ss")
    On Error GoTo 0
    
    CheckRssHeartbeat = True
End Function

Public Sub LogVbaError(ByVal source As String, ByVal errObj As ErrObject)
    LogVbaEvent source, "ERROR " & errObj.Number & ": " & errObj.Description
End Sub

Public Function PlaceOrderRSS(ByVal ticker As String, ByVal side As String, ByVal price As Double, ByVal qty As Long) As Boolean
    On Error GoTo Fail
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(DASH2_SHEET)
    Dim status As String
    status = ws.Range("A3").value
    
    If status <> "LIVE_RUNNING" Then
        LogVbaEvent "PlaceOrderRSS", "DEMO/IDLE: Skipped execution for " & ticker & " " & side & " @" & price
        PlaceOrderRSS = True
        Exit Function
    End If

    Dim evalStr As String
    evalStr = "RssOrder(""" & ticker & """, """ & side & """, " & CStr(price) & ", " & CStr(qty) & ", 0, 0, 0, 0)"
    Dim result As Variant
    result = Application.Evaluate("=" & evalStr)
    If IsError(result) Then
        LogVbaEvent "PlaceOrderRSS", "RSS Error: " & CStr(result)
        GoTo Fail
    End If
    LogVbaEvent "PlaceOrderRSS", "EXECUTED: " & ticker & " " & side & " @" & price
    PlaceOrderRSS = True
    Exit Function
Fail:
    LogVbaError "PlaceOrderRSS", Err
    PlaceOrderRSS = False
End Function

Public Function ModifyOrderRSS(ByVal orderId As String, ByVal newPrice As Double, ByVal currentPrice As Double) As Boolean
    On Error GoTo Fail
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(DASH2_SHEET)
    Dim status As String
    status = ws.Range("A3").value
    If status <> "LIVE_RUNNING" Then
        LogVbaEvent "ModifyOrderRSS", "DEMO/IDLE: Skipped modify for " & orderId & " -> " & newPrice
        ModifyOrderRSS = True
        Exit Function
    End If
    
    Dim thresholdTick As Long
    thresholdTick = 2 ' Default
    On Error Resume Next
    thresholdTick = CLng(ws.Cells(2, 24).value) ' Col 24 is UpdateThreshold
    On Error GoTo Fail
    If thresholdTick <= 0 Then thresholdTick = 2
    
    If Abs(newPrice - currentPrice) < thresholdTick Then
        ModifyOrderRSS = False
        Exit Function
    End If

    Dim evalStr As String
    evalStr = "RssModifyOrder(""" & orderId & """, " & CStr(newPrice) & ", 0)" 
    Dim result As Variant
    result = Application.Evaluate("=" & evalStr)
    If IsError(result) Then
        LogVbaEvent "ModifyOrderRSS", "RSS Error: " & CStr(result)
        GoTo Fail
    End If
    LogVbaEvent "ModifyOrderRSS", "MODIFIED: " & orderId & " -> " & newPrice
    ModifyOrderRSS = True
    Exit Function
Fail:
    LogVbaError "ModifyOrderRSS", Err
    ModifyOrderRSS = False
End Function

Private Sub ExecutePreplaceOrders(ByVal ws As Worksheet, ByVal Sh As Worksheet, ByVal r As Long, _
    ByVal hasBuy As Boolean, ByVal hasSell As Boolean, _
    ByVal eBuyCol As Long, ByVal eSellCol As Long, ByVal qtyCol As Long, ByVal tickerCol As Long)
    
    Dim ticker As String: ticker = ws.Cells(r, tickerCol).value
    Dim qty As Double: qty = ws.Cells(r, qtyCol).value
    
    If hasBuy Then
        Dim buyPrice As Double: buyPrice = ws.Cells(r, eBuyCol).value
        ManageActiveOrder Sh, ticker, "BUY", buyPrice, qty, "PREPLACE"
    End If
    If hasSell Then
        Dim sellPrice As Double: sellPrice = ws.Cells(r, eSellCol).value
        ManageActiveOrder Sh, ticker, "SELL", sellPrice, qty, "PREPLACE"
    End If
End Sub

Private Sub ManageActiveOrder(ByVal Sh As Worksheet, ByVal ticker As String, ByVal side As String, ByVal targetPrice As Double, ByVal qty As Double, ByVal mode As String)
    Dim rowIdx As Long
    rowIdx = FindOrderRow(Sh, ticker, side, Array("PENDING", "ORDERED"))
    If rowIdx = 0 Then Exit Sub
    
    Dim status As String: status = Sh.Cells(rowIdx, 7).value
    Dim currentPrice As Double: currentPrice = ToDouble(Sh.Cells(rowIdx, 4).value, 0)
    Dim orderId As String: orderId = CStr(Sh.Cells(rowIdx, 18).value)
    
    If status = "PENDING" Then
        If PlaceOrderRSS(ticker, side, targetPrice, qty) Then
            Sh.Cells(rowIdx, 7).value = "ORDERED"
            Sh.Cells(rowIdx, 4).value = targetPrice
        End If
    ElseIf status = "ORDERED" Then
        If ModifyOrderRSS(orderId, targetPrice, currentPrice) Then
             Sh.Cells(rowIdx, 4).value = targetPrice
        End If
    End If
End Sub

Private Sub CheckSettlement(ByVal ws As Worksheet, ByVal Sh As Worksheet)
    ' Stub for settlement monitoring
End Sub

Private Sub UpdateSettlementOrders(ByVal ws As Worksheet, ByVal Sh As Worksheet)
    ' Stub for settlement updates
End Sub
'''
    content += "\n" + new_functions

    # Normalize newlines to CRLF for VBA
    content = content.replace('\r\n', '\n').replace('\n', '\r\n')

    # Write output (cp932)
    with open(OUTPUT_PATH, "wb") as f:
        f.write(content.encode("cp932", errors="replace"))
    
    print(f"Successfully created {OUTPUT_PATH}")

if __name__ == "__main__":
    fix_and_finalize()
