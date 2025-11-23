import win32com.client
import os
from pathlib import Path

def safe_export_vba():
    output_dir = r"c:\AI\asagake\temp_vba_audit"
    target_module = "AutoTraderAdvanced"
    
    try:
        # Connect to active Excel instance
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            # Fallback to creating new instance if not running (though user likely has it open)
            excel = win32com.client.Dispatch("Excel.Application")
            # We won't open the workbook here to avoid conflicts, assuming user has it open.
            # If not, we can't easily guess the path if we just started Excel.
            # But let's assume the user has it open as per context.
        
        wb = None
        for w in excel.Workbooks:
            if w.Name == "ASAGAKE.xlsm":
                wb = w
                break
        
        if wb is None:
            print("Error: ASAGAKE.xlsm not found in active Excel instance.")
            # Try opening it read-only? No, safer to ask user.
            return

        print(f"Connected to {wb.Name}")
        
        # Find component
        comp = None
        for c in wb.VBProject.VBComponents:
            if c.Name == target_module:
                comp = c
                break
        
        if comp:
            output_path = os.path.join(output_dir, target_module + ".bas")
            comp.Export(output_path)
            print(f"Successfully exported {target_module} to {output_path}")
        else:
            print(f"Module {target_module} not found.")

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    safe_export_vba()
