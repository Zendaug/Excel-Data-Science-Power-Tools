Attribute VB_Name = "SaveAs"
Sub SaveAsCSV()
    PathName = ActiveWorkbook.Path & "\" & ActiveSheet.Name & ".csv"
    If Left(PathName, 1) = "\" Then PathName = Right(PathName, Len(PathName) - 1)
    
    ActiveSheet.Copy
    ActiveSheet.UsedRange.Value = ActiveSheet.UsedRange.Value 'Break links; replace all formulas with values
    ActiveWorkbook.SaveAs Filename:=PathName, FileFormat:=xlCSV, CreateBackup:=False
    
    Call CopyTextToClipboard(PathName)
    ActiveWorkbook.Close
End Sub

Sub SaveAsXLSX()
    PathName = ActiveWorkbook.Path & "\" & ActiveSheet.Name & ".xlsx"
    
    ActiveSheet.Copy
    ActiveSheet.UsedRange.Value = ActiveSheet.UsedRange.Value 'Break links; replace all formulas with values
    ActiveWorkbook.SaveAs Filename:=PathName, CreateBackup:=False
    
    Call CopyTextToClipboard(PathName)
    ActiveWorkbook.Close
End Sub

Sub SaveAsTAB()
    PathName = ActiveWorkbook.Path & "\" & ActiveSheet.Name & ".txt"
    
    ActiveSheet.Copy
    ActiveSheet.UsedRange.Value = ActiveSheet.UsedRange.Value 'Break links; replace all formulas with values
    ActiveWorkbook.SaveAs Filename:=PathName, FileFormat:=xlText, CreateBackup:=False
    
    Call CopyTextToClipboard(PathName)
    ActiveWorkbook.Close
End Sub

Sub SaveAsCSV_R()
    Call SaveAsCSV
    
    Open ActiveWorkbook.Path & "\" & ActiveSheet.Name & ".R" For Output As 1
    
    Print #1, "load.package <- function(x)"
    Print #1, "{"
    Print #1, "  if (!require(x,character.only = TRUE))"
    Print #1, "  {"
    Print #1, "    install.packages(x,dep=TRUE)"
    Print #1, "    if(!require(x,character.only = TRUE)) stop(""Package not found"")"
    Print #1, "  }"
    Print #1, "}"
    Print #1, ""
    Print #1, "dataset <- read.csv(""" & ActiveSheet.Name & ".csv"", header=TRUE)"
    Close #1
End Sub

Sub SaveAsCSV_Python()
    Call SaveAsCSV
    
    Open ActiveWorkbook.Path & "\" & ActiveSheet.Name & ".py" For Output As 1
    
    Print #1, "import pandas as pd"
    Print #1, "dataset = pd.read_csv('" & ActiveSheet.Name & ".csv" & "')"
    Print #1, ""
    Print #1, "# Extract specific record (y) from column ""x"""
    Print #1, "#dataset.loc[y][""x""]"
    Print #1, ""
    Print #1, "# Extract column ""x"""
    Print #1, "#dataset[""x""]"
    Close #1
End Sub


Sub ReadfromXLS()
    Open ActiveWorkbook.Path & "\" & ActiveSheet.Name & ".R" For Output As 1
    
    Print #1, "load.package <- function(x)"
    Print #1, "{"
    Print #1, "  if (!require(x,character.only = TRUE))"
    Print #1, "  {"
    Print #1, "    install.packages(x,dep=TRUE)"
    Print #1, "    if(!require(x,character.only = TRUE)) stop(""Package not found"")"
    Print #1, "  }"
    Print #1, "}"
    Print #1, ""
    Print #1, "load.package (""openxlsx"")"
    Print #1, ""
    Print #1, "dataset <- read.xlsx(xlsxFile = """ & ActiveWorkbook.Name & """, sheet = """ & ActiveSheet.Name & """, startRow = 1, colNames = TRUE, "
    Print #1, "          rowNames = FALSE, detectDates = FALSE, skipEmptyRows = TRUE,"
    Print #1, "          skipEmptyCols = TRUE, rows = NULL, cols = NULL, check.names = FALSE,"
    Print #1, "          namedRegion = NULL, na.strings = ""NA"", fillMergedCells = FALSE)"
    
    Close #1
    MsgBox "'" & ActiveWorkbook.Path & "\" & ActiveSheet.Name & ".R" & "' created."
End Sub

Sub ReadfromXLS_python()
    Open ActiveWorkbook.Path & "\" & ActiveSheet.Name & ".py" For Output As 1
    
    Print #1, "import pandas as pd"
    Print #1, ""
    Print #1, "dataset = pd.read_excel('" & ActiveWorkbook.Name & "', sheet_name = '" & ActiveSheet.Name & "', header = 0, index_col=None)"
        
    Close #1
    MsgBox "'" & ActiveWorkbook.Path & "\" & ActiveSheet.Name & ".py" & "' created."
End Sub

