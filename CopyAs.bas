Attribute VB_Name = "CopyAs"
' This library is a set of macros for copying cells to the clipboard

'Handle 64-bit and 32-bit Office
#If VBA7 Then
  Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
  Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As Long
  Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As LongPtr, _
    ByVal dwBytes As LongPtr) As Long
  Declare PtrSafe Function CloseClipboard Lib "User32" () As Long
  Declare PtrSafe Function OpenClipboard Lib "User32" (ByVal hwnd As LongPtr) As Long
  Declare PtrSafe Function EmptyClipboard Lib "User32" () As Long
  Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, _
    ByVal lpString2 As Any) As Long
  Declare PtrSafe Function SetClipboardData Lib "User32" (ByVal wFormat _
    As LongPtr, ByVal hMem As LongPtr) As Long
#Else
  Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
  Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
  Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, _
    ByVal dwBytes As Long) As Long
  Declare Function CloseClipboard Lib "User32" () As Long
  Declare Function OpenClipboard Lib "User32" (ByVal hwnd As Long) As Long
  Declare Function EmptyClipboard Lib "User32" () As Long
  Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, _
    ByVal lpString2 As Any) As Long
  Declare Function SetClipboardData Lib "User32" (ByVal wFormat _
    As Long, ByVal hMem As Long) As Long
#End If

Public Const GHND = &H42
Public Const CF_TEXT = 1
Public Const MAXSIZE = 4096

Function ClipBoard_SetData(MyString As String)
'PURPOSE: API function to copy text to clipboard
'SOURCE: www.msdn.microsoft.com/en-us/library/office/ff192913.aspx

Dim hGlobalMemory As Long, lpGlobalMemory As Long
Dim hClipMemory As Long, x As Long

'Allocate moveable global memory
  hGlobalMemory = GlobalAlloc(GHND, Len(MyString) + 1)

'Lock the block to get a far pointer to this memory.
  lpGlobalMemory = GlobalLock(hGlobalMemory)

'Copy the string to this global memory.
  lpGlobalMemory = lstrcpy(lpGlobalMemory, MyString)

'Unlock the memory.
  If GlobalUnlock(hGlobalMemory) <> 0 Then
    MsgBox "Could not unlock memory location. Copy aborted."
    GoTo OutOfHere2
  End If

'Open the Clipboard to copy data to.
  If OpenClipboard(0&) = 0 Then
    MsgBox "Could not open the Clipboard. Copy aborted."
    Exit Function
  End If

'Clear the Clipboard.
  x = EmptyClipboard()

'Copy the data to the Clipboard.
  hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)

OutOfHere2:
  If CloseClipboard() = 0 Then
    MsgBox "Could not close Clipboard."
  End If

End Function

Sub CopyTextToClipboard(Optional txt = "This was copied to the clipboard using VBA!")
'PURPOSE: Copy a given text to the clipboard (using Windows API)
'SOURCE: www.TheSpreadsheetGuru.com
'NOTES: Must have above API declaration and ClipBoard_SetData function in your code

'Dim txt As String

'Place text into the Clipboard
   ClipBoard_SetData (txt)

'Notify User
  'MsgBox "There is now text copied to your clipboard!", vbInformation

End Sub

Sub CopyToDataFrame()
' Copy the selected cells in a format that can be pasted into R as a dataframe
    temp_data = selection.Value
    r_text = "x <- data.frame("
    For a = 1 To UBound(temp_data, 2)
        If IsNumeric(temp_data(1, a)) = False Then
            If InStr(1, temp_data(1, a), " ", vbTextCompare) > 0 Then
                temp_data(1, a) = """" & temp_data(1, a) & """"
            End If
        End If
        r_text = r_text & temp_data(1, a) & " = c("
        For b = 2 To UBound(temp_data, 1)
            If IsNumeric(temp_data(b, a)) = False Then
                If Left(temp_data(b, a), 1) = "=" Then
                    temp_data(b, a) = Right(temp_data(b, a), Len(temp_data(b, a)) - 1) 'Allow for functions, such as =mean(1,2,3)
                Else
                    temp_data(b, a) = """" & temp_data(b, a) & """"
                End If
            End If
            If temp_data(b, a) = Null Or temp_data(b, a) = "" Then temp_data(b, a) = "NA"
            r_text = r_text & temp_data(b, a)
            If b < UBound(temp_data, 1) Then r_text = r_text & ", "
        Next
        r_text = r_text & ")"
        If a < UBound(temp_data, 2) Then r_text = r_text & "," & Chr(13) & Chr(10)
    Next
    r_text = r_text & Chr(13) & Chr(10) & ")"
    ClipBoard_SetData (r_text)
End Sub


Sub CopyToPythonDictList()
' Copy the selected cells in a format that can be pasted into Python as a combined dictionary/list
    temp_data = selection.Value
    p_text = "x = {"
    For a = 1 To UBound(temp_data, 2)
        If IsNumeric(temp_data(1, a)) = False Then
            If InStr(1, temp_data(1, a), " ", vbTextCompare) > 0 Then
                temp_data(1, a) = """" & temp_data(1, a) & """"
            End If
        End If
        p_text = p_text & temp_data(1, a) & " : ["
        For b = 2 To UBound(temp_data, 1)
            If IsNumeric(temp_data(b, a)) = False Then
                If Left(temp_data(b, a), 1) = "=" Then
                    temp_data(b, a) = Right(temp_data(b, a), Len(temp_data(b, a)) - 1) 'Allow for functions, such as =mean(1,2,3)
                Else
                    temp_data(b, a) = """" & temp_data(b, a) & """"
                End If
            End If
            If temp_data(b, a) = Null Or temp_data(b, a) = "" Then temp_data(b, a) = "np.NaN"
            p_text = p_text & temp_data(b, a)
            If b < UBound(temp_data, 1) Then p_text = p_text & ", "
        Next
        p_text = p_text & "]"
        If a < UBound(temp_data, 2) Then p_text = p_text & "," & Chr(10)
    Next
    p_text = p_text & "}"
    ClipBoard_SetData (p_text)
End Sub

Sub CopyToPythonDictList_pandas()
' Copy the selected cells in a format that can be pasted into Python as a pandas dataframe
    temp_data = selection.Value
    p_text = "x = pd.DataFrame(data = {"
    For a = 1 To UBound(temp_data, 2)
        If IsNumeric(temp_data(1, a)) = False Then
            If InStr(1, temp_data(1, a), " ", vbTextCompare) > 0 Then
                temp_data(1, a) = """" & temp_data(1, a) & """"
            End If
        End If
        p_text = p_text & temp_data(1, a) & " : ["
        For b = 2 To UBound(temp_data, 1)
            If IsNumeric(temp_data(b, a)) = False Then
                If Left(temp_data(b, a), 1) = "=" Then
                    temp_data(b, a) = Right(temp_data(b, a), Len(temp_data(b, a)) - 1) 'Allow for functions, such as =mean(1,2,3)
                Else
                    temp_data(b, a) = """" & temp_data(b, a) & """"
                End If
            End If
            If temp_data(b, a) = Null Or temp_data(b, a) = "" Then temp_data(b, a) = "np.NaN"
            p_text = p_text & temp_data(b, a)
            If b < UBound(temp_data, 1) Then p_text = p_text & ", "
        Next
        p_text = p_text & "]"
        If a < UBound(temp_data, 2) Then p_text = p_text & "," & Chr(10)
    Next
    p_text = p_text & "})"
    ClipBoard_SetData (p_text)
End Sub
