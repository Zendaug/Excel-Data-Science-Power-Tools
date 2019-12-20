Attribute VB_Name = "MiscString"
Function ReverseName(full_name As String)
' Reverses a name putting the surname first (e.g., from "John Smith" to "Smith, John")
    
    For a = 1 To Len(full_name)
        If a > Len(full_name) Then Exit For
        temp = InStr(a, full_name, " ")
        If temp = 0 Then
            Exit For
        Else
            space_pos = temp
            a = space_pos + 1
        End If
    Next
    
    ReverseName = Right(full_name, Len(full_name) - space_pos) & ", " & Left(full_name, space_pos - 1)
    
End Function

Function WEEKS(start_week, skip_week) As String
    WEEKS = ""
    For a = 1 To 13
        If (start_week + a - 1) <> skip_week Then
            If a > 1 Then WEEKS = WEEKS & ","
            WEEKS = WEEKS & start_week + a - 1
        End If
    Next
End Function

Function PAYCODE(category, instance)
    If category = "Ongoing" Then
        PAYCODE = "Ongoing staff member - delete this line"
    ElseIf category = "Normal" Then
        If instance = 1 Then PAYCODE = "TE"
        If instance > 1 Then PAYCODE = "TF"
    ElseIf category = "PhD" Then
        If instance = 1 Then PAYCODE = "TG"
        If instance > 1 Then PAYCODE = "TH"
    End If
End Function


Function CONCATENATEMULTIPLE(Ref As Range, Optional Separator As String = ", ") As String
    ' Concatenates a selected range and returns it as text
    Dim Cell As Range
    Dim Result As String
    For Each Cell In Ref
        Result = Result & Cell.Value & Separator
    Next Cell
    CONCATENATEMULTIPLE = Left(Result, Len(Result) - Len(Separator))
End Function

Function disp_array(temp_arr)
    ' Returns the contents of an array as text
    temp = ARRAY_FLATTEN(temp_arr)
    disp_array = ""
    For a = 1 To UBound(temp)
        disp_array = disp_array & temp(a) & " "
    Next
End Function

Function CELLS_TO_LINE(cells_range As Range, Optional sep = " ", Optional max_length = 80)
    ' Returns the nominated cells as text
    temp = cells_range.Value
    temp_line = ""
    CELLS_TO_LINE = ""
    For a = 1 To UBound(temp, 1)
        temp_line = temp_line & temp(a, 1) & sep
        If Len(temp_line) + Len(sep) > max_length Then
            CELLS_TO_LINE = CELLS_TO_LINE & Chr(10)
            temp_line = ""
        End If
        CELLS_TO_LINE = CELLS_TO_LINE & temp(a, 1)
        If a < UBound(temp, 1) Then CELLS_TO_LINE = CELLS_TO_LINE & sep
    Next
End Function

