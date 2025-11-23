# ASAGAKE V2 VBA Module Deployment Guide

The `AutoTraderAdvanced` module has been successfully patched with V2 features, including:
- **Heartbeat Monitor**: Checks RSS updates in cell `Z2`.
- **Order Correction Threshold**: New parameter in Column 24 ("注文訂正閾値(tick)").
- **Live/Demo Mode**: `PlaceOrderRSS` and `ModifyOrderRSS` now respect the "LIVE_RUNNING" status.
- **Pre-order Execution**: `ExecutePreplaceOrders` handles pending pre-orders.
- **Fixes**: Resolved "Too many line continuations" error.

## Steps to Deploy

1.  **Open `ASAGAKE.xlsm`**.
2.  **Open VBA Editor** (`Alt + F11`).
3.  **Remove Old Module**:
    *   Right-click on `AutoTraderAdvanced` in the Project Explorer.
    *   Select **Remove AutoTraderAdvanced...**
    *   Select **No** when asked to export.
4.  **Import New Module**:
    *   Right-click on the `Modules` folder (or the project name).
    *   Select **Import File...**
    *   Navigate to: `c:\AI\asagake\temp_vba_audit\AutoTraderAdvanced_v2_final.bas`
    *   Click **Open**.
5.  **Compile**:
    *   Go to **Debug** menu -> **Compile VBAProject**.
    *   Ensure there are no errors.
6.  **Run Setup Macro**:
    *   In the VBA Editor, find `Sub SetupDashboardUIV2` in the new module.
    *   Click inside the sub and press `F5` (Run).
    *   *Alternatively*, run it from Excel: `Alt + F8` -> `SetupDashboardUIV2` -> **Run**.

## Verification

After running `SetupDashboardUIV2`:
*   **Check Cell Z2**: It should be labeled "RSS Monitor" and eventually show a timestamp (if RSS is active).
*   **Check Column 24 (Row 1)**: It should have the header "注文訂正閾値(tick)".
*   **Check Parameter**: Ensure cell `X2` (Row 2, Col 24) has a default value (e.g., 2).

## Troubleshooting
If you see "Variable not defined" or other errors, please report the specific line number and error message.
