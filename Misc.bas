Attribute VB_Name = "Misc"
Public Sub ExportModules()
' A sub that will export all of the modules for upload to Github
    Dim bExport As Boolean
    Dim wkbSource As Excel.Workbook
    Dim szSourceWorkbook As String
    Dim szExportPath As String
    Dim szFileName As String
    Dim cmpComponent As VBIDE.VBComponent

    ''' The code modules will be exported in a folder named.
    ''' VBAProjectFiles in the Documents folder.
    ''' The code below create this folder if it not exist
    ''' or delete all files in the folder if it exist.
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Export Folder not exist"
        Exit Sub
    End If
    
    On Error Resume Next
        Kill FolderWithVBAProjectFiles & "\*.*"
    On Error GoTo 0

    ''' NOTE: This workbook must be open in Excel.
    'szSourceWorkbook = ActiveWorkbook.Name
    'Set wkbSource = Application.Workbooks(szSourceWorkbook)

    Set wkbSource = Application.Workbooks("Custom.xlam")
    
    If wkbSource.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to export the code"
    Exit Sub
    End If
    
    Dim sFolder As String
    ' Open the select folder prompt
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then ' if OK is pressed
            sFolder = .SelectedItems(1)
        End If
    End With
    
    If sFolder <> "" Then ' if a file was chosen
        szExportPath = sFolder & "\"
    Else
        Exit Sub
    End If
    
    For Each cmpComponent In wkbSource.VBProject.VBComponents
        
        bExport = True
        szFileName = cmpComponent.Name

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                szFileName = szFileName & ".cls"
            Case vbext_ct_MSForm
                szFileName = szFileName & ".frm"
            Case vbext_ct_StdModule
                szFileName = szFileName & ".bas"
            Case vbext_ct_Document
                ''' This is a worksheet or workbook object.
                ''' Don't try to export.
                bExport = False
        End Select
        
        If bExport Then
            ''' Export the component to a text file.
            cmpComponent.Export szExportPath & szFileName
            
        ''' remove it from the project if you want
        '''wkbSource.VBProject.VBComponents.Remove cmpComponent
        
        End If
   
    Next cmpComponent

    MsgBox "Export is ready"
End Sub


Function FolderWithVBAProjectFiles() As String
    Dim WshShell As Object
    Dim FSO As Object
    Dim SpecialPath As String

    Set WshShell = CreateObject("WScript.Shell")
    Set FSO = CreateObject("scripting.filesystemobject")

    SpecialPath = WshShell.SpecialFolders("MyDocuments")

    If Right(SpecialPath, 1) <> "\" Then
        SpecialPath = SpecialPath & "\"
    End If
    
    If FSO.FolderExists(SpecialPath & "VBAProjectFiles") = False Then
        On Error Resume Next
        MkDir SpecialPath & "VBAProjectFiles"
        On Error GoTo 0
    End If
    
    If FSO.FolderExists(SpecialPath & "VBAProjectFiles") = True Then
        FolderWithVBAProjectFiles = SpecialPath & "VBAProjectFiles"
    Else
        FolderWithVBAProjectFiles = "Error"
    End If
    
End Function

Function UNIQUE(data_range, Optional unique_case = 0)
' Identifies unique values, based on the nominated case (unique_case)
' If "unique_case" is set to 0, it returns the number of unique cases
' Works either from a worksheet range, or 1-dimensional VBA Array

    Dim temp_array()
    If TypeName(data_range) = "Range" Then
        temp_data = ARRAY_FLATTEN(data_range.Value)
        n_rows = UBound(temp_data, 1)
        n_cols = 0
        ReDim temp_array(1 To n_rows)
    ElseIf TypeName(data_range) = "Variant()" Then
        Debug.Print "It's a variant"
        temp_data = ARRAY_FLATTEN(data_range)
        n_rows = UBound(temp_data, 1)
        n_cols = 0
        ReDim temp_array(1 To n_rows)
    Else
        temp_data = data_range
    End If
    
    c = 0
    
    If n_cols > 0 Then
        For a = 1 To n_rows
            Debug.Print a & " of " & n_rows
            For b = 1 To n_cols
                If temp_data(a, b) <> "" And temp_data <> "NA" Then
                    If c > 0 Then
                        found_match = False
                        For d = 1 To c
                            If temp_array(d) = temp_data(a, b) Then
                                found_match = True
                                Exit For
                            End If
                        Next
                        If found_match = False Then
                            temp_array(c + 1) = temp_data(a, b)
                            c = c + 1
                        End If
                    Else
                        temp_array(1) = temp_data(a, b)
                        c = 1
                    End If
                End If
            Next
        Next
    Else
        For a = 1 To n_rows
            If temp_data(a) <> "" Then
                If c > 0 Then
                    found_match = False
                    For d = 1 To c
                        If temp_array(d) = temp_data(a) Then
                            found_match = True
                            Exit For
                        End If
                    Next
                    If found_match = False Then
                        temp_array(c + 1) = temp_data(a)
                        c = c + 1
                    End If
                Else
                    temp_array(1) = temp_data(a)
                    c = 1
                End If
            End If
        Next
    End If
    
    If unique_case = 0 Then
        UNIQUE = c
    Else
        UNIQUE = temp_array(unique_case)
    End If
End Function


Function UNIQUE_SEP(data_range, Optional sep = ",", Optional unique_case = 0)
' Identifies unique values, and allows for separators
' Works either from a worksheet range, or 1-dimensional VBA Array

' Figure out whether a worksheet range or array
    Dim temp_array()
    If TypeName(data_range) = "Range" Then
        temp_data = data_range.Value
        n_rows = UBound(temp_data, 1)
        n_cols = UBound(temp_data, 2)
        ReDim temp_array(1 To n_rows * n_cols)
    ElseIf TypeName(data_range) = "Variant()" Then
        temp_data = data_range
        n_rows = UBound(temp_data, 1)
        n_cols = 0
        ReDim temp_array(1 To n_rows)
    End If

' Flatten the array
    temp_data2 = ARRAY_FLATTEN(temp_data)

' Count the number of total elements in the new array
    elements_n = 0
    For a = 1 To UBound(temp_data2, 1)
        arr_split = Split(temp_data2(a), sep)
        elements_n = elements_n + UBound(arr_split) + 1
    Next

' Parse the array; insert new elements into it
    Dim temp_data3()
    ReDim temp_data3(1 To elements_n)
    element_pos = 0
    For a = 1 To UBound(temp_data2, 1)
        arr_split = Split(temp_data2(a), sep)
        For b = LBound(arr_split) To UBound(arr_split)
            element_pos = element_pos + 1
            temp_data3(element_pos) = Trim(arr_split(b))
        Next
    Next

    n_rows = UBound(temp_data3, 1)
    ReDim temp_array(1 To n_rows)
    c = 0 ' Number of hits
    
' Count the number of unique values
    For a = 1 To n_rows
        If temp_data3(a) <> "" Then
            If c > 0 Then
                found_match = False
                For d = 1 To c
                    If temp_array(d) = temp_data3(a) Then
                        found_match = True
                        Exit For
                    End If
                Next
                If found_match = False Then
                    temp_array(c + 1) = temp_data3(a)
                    c = c + 1
                End If
            Else
                temp_array(1) = temp_data3(a)
                c = 1
            End If
        End If
    Next
    
    If unique_case = 0 Then
        UNIQUE_SEP = c
    Else
        UNIQUE_SEP = temp_array(unique_case)
    End If
End Function

Function MATCH_ROW(target, data_range As Range, Optional instance = 1)
    temp_data = data_range.Value
    
    n_rows = UBound(temp_data, 1)
    n_cols = UBound(temp_data, 2)
    n_inst = 0
    
    For a = 1 To n_rows
        For b = 1 To n_cols
            If temp_data(a, b) = target Then
                n_inst = n_inst + 1
                If n_inst = instance Then
                    MATCH_ROW = a
                    Exit Function
                End If
            End If
        Next
    Next
    
    MATCH_ROW = "Not found"
End Function
