# Excel Automation Add-in

This project is an Excel automation add-in that integrates VBA with a Python backend. It enables context-aware Excel cell operations like entering formulas or navigating to specific cells, based on the spreadsheet's data and layout. The add-in is designed to enhance productivity and reduce manual work by generating Excel instructions through an AI model.


---

## Demos (My hands are off the keyboard!)

### 1. It can fill by intelligently understanding the worksheet

`Fill this column appropriately starting from D2 till D5`

![Excel Automation Demo](/Screen-Recording-2025-07-15-135508_1.gif)


### 2. It is super useful for formula based tasks

`Fill this cell with the correct formula/value`

![Excel Automation Demo](/Screen-Recording-2025-07-15-135618.gif)

### 3. Or filling entire tables!

`Fill this table completely using 4 NBA teams in this column`

![Excel Automation Demo](/Screen-Recording-2025-07-15-135316.gif)

## Features

- Contextual AI-driven Excel commands
- Lightweight communication via flag-based signaling

## Architecture

1. **VBA (Frontend)**:
   - Triggers Python script via `WScript.Shell` and `forms` for loading screen.

2. **Python (Backend)**:
   - Reads current Excel state and executes user's automation tasks.
   
---

## Polling, polling, polling...

Due to VBA’s lack of an intuitive async or callback support:

- **Race conditions**: `xlwings` could not access Excel during `Application.Wait`.
- **No native callbacks**: VBA required a workaround for communication.
- **Polling with flags**: To coordinate Python-VBA interaction, simple file-based flags were used to indicate stages like "loading started" and "done".

This workaround ensured non-blocking communication across two runtimes.

---

## Potential Improvements

- Add more error handling in both VBA and Python
- Replace file-based flags
- Enhance UI/UX with progress bars
- Support for undo actions, multiple sheets, or more command types

---

## Code References

**VBA Add In Code**
```vba
 Public Sub MyMacro(ByRef control As Office.IRibbonControl)
    Dim shell As Object
    Dim pythonPath As String
    Dim scriptPath As String
    Dim command As String

    ' Clean up any old flags from previous runs
    If Dir("C:\temp\done.txt") <> "" Then Kill "C:\temp\done.txt"
    If Dir("C:\temp\show_loading.txt") <> "" Then Kill "C:\temp\show_loading.txt"

    ' Paths
    Set shell = CreateObject("WScript.Shell")
    pythonPath = "C:\Users\mosai\OneDrive\Desktop\excelTester\.venv\Scripts\pythonw.exe"
    scriptPath = "C:\Users\mosai\OneDrive\Desktop\excelTester\.venv\excelClick.py"
    command = """" & pythonPath & """ """ & scriptPath & """"

    ' Asynchronous Run
    shell.Run command, 1, False

    
    ' Do While Dir("C:\temp\show_loading.txt") = ""
    '    DoEvents
    ' Loop

    
    frmLoading.Show vbModeless


    Do While Dir("C:\temp\done.txt") = ""
        DoEvents
    Loop

    Unload frmLoading

    On Error Resume Next
    f = FreeFile()
    Open "C:\temp\result.txt" For Input As #f
        If Err.Number = 0 Then
        Line Input #f, result
        Close #f
    Else
    End If
    On Error GoTo 0
    MsgBox result & " instructions will be executed.", vbInformation, "Execution Summary"
    If Dir("C:\temp\flag.txt") <> "" Then Kill "C:\temp\flag.txt"
    
    ' MsgBox "All instructions executed!", vbOKOnly, AutoExcel
End Sub

