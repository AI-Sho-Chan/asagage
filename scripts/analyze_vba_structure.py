import re

def analyze():
    path = r"c:\AI\asagake\temp_vba_audit\AutoTraderAdvanced_v2_final.bas"
    try:
        with open(path, "r", encoding="cp932", errors="replace") as f:
            lines = f.readlines()
    except FileNotFoundError:
        print("File not found.")
        return

    print(f"Total lines: {len(lines)}")
    
    subs = []
    functions = []
    ontime_calls = []
    
    sub_pattern = re.compile(r'^\s*(Private\s+|Public\s+)?Sub\s+(\w+)', re.IGNORECASE)
    func_pattern = re.compile(r'^\s*(Private\s+|Public\s+)?Function\s+(\w+)', re.IGNORECASE)
    
    for i, line in enumerate(lines):
        if sub_pattern.match(line):
            subs.append((i+1, line.strip()))
        if func_pattern.match(line):
            functions.append((i+1, line.strip()))
        if "Application.OnTime" in line:
            ontime_calls.append((i+1, line.strip()))
            
    print("\n--- Subroutines ---")
    for ln, text in subs:
        print(f"{ln}: {text}")
        
    print("\n--- Functions ---")
    for ln, text in functions:
        print(f"{ln}: {text}")
        
    print("\n--- OnTime Calls ---")
    for ln, text in ontime_calls:
        print(f"{ln}: {text}")

if __name__ == "__main__":
    analyze()
