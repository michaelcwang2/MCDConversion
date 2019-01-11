Attribute VB_Name = "Module2"
Sub ProcessMCDs()

' ProcessMCDs Macro

    ' Declare Variables
        ' Strings
            Dim pFileName As String
            Dim pFilePath As String
        ' Workbooks
            Dim pWorkbook As Workbook
    
    ' Setup File Path
        pFilePath = ActiveWorkbook.Path & "\MCD\"
        
    ' Find First MCD in Folder (.xls)
        pFileName = Dir(pFilePath & "*.xls")
    
    ' Loop For MCDConversion and Finding Next MCD in Folder (.xls)
        Do While pFileName <> ""
            Set pWorkbook = Workbooks.Open(pFilePath & pFileName)
            MCDConversion pWorkbook
            pWorkbook.Close SaveChanges:=True
            
            ' Get Next File Name
                pFileName = Dir()
        Loop
        
    ' Objective Complete - Notify
        MsgBox "Done"
    
End Sub

Sub MCDConversion(pWorkbook As Workbook)
Attribute MCDConversion.VB_ProcData.VB_Invoke_Func = " \n14"

' MCDConversion Macro

    ' Declare Variables
        'Doubles
            Dim TotalPartWeight As Double
        'Integers
            Dim Iterator As Integer
            Dim Multiplier1K As Integer
            Dim NumSubstanceMasses As Integer
        'Ranges
            Dim TempRange As Range
        'Strings
            Dim FileName As String
            Dim InternalPN As String
            Dim Manufacturer As String
            Dim ManufacturerPN As String
            Dim PartWeightString As String
            Dim TempString As String
        'Variants
            Dim TempStringArray As Variant
    
    ' Turn Off Alerts
        Application.DisplayAlerts = False

    ' Delete Sheets (Excluding Chemical Data)
        For Iterator = 1 To ActiveWorkbook.Worksheets.Count
        If Sheets(Iterator).Name <> "Chemical Data" Then Sheets(Iterator).Select Replace:=False
        Next Iterator
        ActiveWindow.SelectedSheets.Delete

    ' Get Variables
        ' Set Multiplier
            Multiplier1K = 1000
        ' Get FileName
            FileName = ActiveWorkbook.Name
        ' Set InternalPN
            TempStringArray = Split(FileName, " ")
            InternalPN = TempStringArray(0)
        ' Set Manufacturer Name
            TempStringArray = Split(FileName, InternalPN)
            TempString = TempStringArray(1)
            TempStringArray = Split(TempString, ".")
            Manufacturer = TempStringArray(0)
            Manufacturer = Right(Manufacturer, (Len(Manufacturer) - 1))

    ' Variable: Part Weight
        PartWeightString = Range("H7")
        TempStringArray = Split(PartWeightString, " ")
        PartWeightString = TempStringArray(4)
        TotalPartWeight = CDbl(PartWeightString)
        TotalPartWeight = TotalPartWeight * Multiplier1K
    
    ' Variable: Manufacturer Part Number
        ManufacturerPN = Range("C3")
        TempStringArray = Split(ManufacturerPN, " ")
        ManufacturerPN = TempStringArray(3)

    ' Remove Silicon Expert Header
        Rows("1:8").Select
        Selection.Delete Shift:=xlUp
        
    ' Unmerge All Cells
        ActiveSheet.Cells.UnMerge
        
    ' Remove Item/Subitem Name and Mass
        Columns("A:B").Select
        Selection.Delete Shift:=xlToLeft
    
    ' Delete PPM
        Columns("F:F").Select
        Selection.Delete Shift:=xlToLeft
    
    ' Move CAS Numbers
        Columns("D:D").Select
        Selection.Cut
        Columns("J:J").Select
        ActiveSheet.Paste
    
    ' Move Substance Mass
        Columns("E:E").Select
        Selection.Cut
        Columns("G:G").Select
        ActiveSheet.Paste

    ' Count Substance Masses for Multiplier Range
        Set TempRange = Range("G:G")
        NumSubstanceMasses = WorksheetFunction.CountA(TempRange)
    
    ' Set Multiplier g to mg
        Range("D1").Select
        ActiveCell.FormulaR1C1 = Multiplier1K
        Range("D1").Select
        Selection.Copy
    
    ' Substance Mass
        Range("G1:G" & NumSubstanceMasses).Select
        Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlMultiply, _
            SkipBlanks:=False, Transpose:=False

    ' Weight
        Range("B1:B" & NumSubstanceMasses).Select
        Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlMultiply, _
            SkipBlanks:=False, Transpose:=False

    ' Remove Multiplier from Sheet
        Range("D1").Select
        Application.CutCopyMode = False
        Selection.ClearContents
    
    ' Grab Greensoft Header (MCD Header.xlsx Must be Open)
        Windows("MCD Header.xlsx").Activate
        Range("A1:M5").Select
        Selection.Copy
        Windows(FileName).Activate
        Rows("1:1").Select
        Selection.Insert Shift:=xlDown
    
    ' Place Data in Header
        Range("B2") = Manufacturer
        Range("B3") = ManufacturerPN
        Range("D2") = InternalPN
        Range("B4") = TotalPartWeight
    
    ' Delete the Silicon Expert Logo
        ActiveSheet.Pictures.Delete
    
End Sub

