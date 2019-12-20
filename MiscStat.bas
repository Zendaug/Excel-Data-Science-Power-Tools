Attribute VB_Name = "MiscStat"
Function BLAU_CLUSTER(group_range, cluster_range, cluster, Optional ignore = "")
' Calculates Blau's index for the groups (e.g., teams) within a larger cluster (e.g., departments)
' group_range is a categorical variable representing group membership (e.g., teams)
' cluster_range is a broader clustering variable (e.g., departments)
' cluster is the cluster for which Blau's index is calculated for
' Works either from a worksheet range, array formulas, or 1-dimensional VBA Array
' Ignores the "ignore" value

    Dim temp_array(), temp_array2()
    If TypeName(group_range) = "Range" Then ' From Range
        temp_array = ARRAY_FLATTEN(group_range.Value)
    ElseIf TypeName(group_range) = "Variant()" Then ' From Array formulas
        temp_array = ARRAY_FLATTEN(group_range)
    Else ' From regular array
        temp_array = group_range
    End If
        
    group_array = cluster_range.Value
            
    ReDim temp_array2(1 To UNIQUE(temp_array), 1 To 2)
    Dim tot: tot = 0
    For a = 1 To UNIQUE(temp_array)
        temp_array2(a, 1) = UNIQUE(temp_array, a)
        temp_array2(a, 2) = 0
        For b = 1 To UBound(temp_array)
            If temp_array(b) = temp_array2(a, 1) And temp_array2(a, 1) <> ignore And group_array(b, 1) = cluster Then
                temp_array2(a, 2) = temp_array2(a, 2) + 1
                tot = tot + 1
            End If
        Next
    Next

    BLAU_CLUSTER = 1
    
    For a = 1 To UBound(temp_array2, 1)
        If temp_array2(a, 1) <> ignore Then
            BLAU_CLUSTER = BLAU_CLUSTER - (temp_array2(a, 2) / tot) ^ 2
        End If
    Next

End Function

Function BLAU(group_range, Optional ignore = "")
' Calculate's Blau's index.
' group_range is a categorical variable representing group membership
' Works either from a worksheet range, array formulas, or 1-dimensional VBA Array
' Ignores the "ignore" value

    Dim temp_array(), temp_array2()
    If TypeName(group_range) = "Range" Then ' From Range
        temp_array = ARRAY_FLATTEN(group_range.Value)
    ElseIf TypeName(group_range) = "Variant()" Then ' From Array formulas
        temp_array = ARRAY_FLATTEN(group_range)
    Else ' From regular array
        temp_array = group_range
    End If
        
    ReDim temp_array2(1 To UNIQUE(temp_array), 1 To 2)
    Dim tot: tot = 0
    For a = 1 To UNIQUE(temp_array)
        temp_array2(a, 1) = UNIQUE(temp_array, a)
        temp_array2(a, 2) = 0
        For b = 1 To UBound(temp_array)
            If temp_array(b) = temp_array2(a, 1) And temp_array2(a, 1) <> ignore Then
                temp_array2(a, 2) = temp_array2(a, 2) + 1
                tot = tot + 1
            End If
        Next
    Next
    
    BLAU = 1
    
    For a = 1 To UBound(temp_array2, 1)
        If temp_array2(a, 1) <> ignore Then
            BLAU = BLAU - (temp_array2(a, 2) / tot) ^ 2
        End If
    Next
        
End Function

Function pval(p_val, Optional pLessTen = False)
    ' Automatically generates asterisks based on the p value
    pval = ""
    If p_val < 0.1 And p_val >= 0.05 And pLessTen = True Then pval = "(*)"
    If p_val < 0.05 And p_val >= 0.01 Then pval = "*"
    If p_val < 0.01 And p_val >= 0.001 Then pval = "**"
    If p_val < 0.001 Then pval = "***"
End Function

Function COEF(eff, se, p_val, Optional decimals = 2, Optional pLessTen = False)
    dec = "."
    If decimals > 0 Then
        For a = 1 To decimals
            dec = dec & "0"
        Next
    End If

    COEF = WorksheetFunction.Text(eff, dec) & pval(p_val, pLessTen) & " (" & _
    WorksheetFunction.Text(se, dec) & ")"
End Function

Function PARTIAL(x1_y, x2_y, x1_x2)
    ' Returns the partial correlation between X1 and Y
    PARTIAL = (x1_y - (x1_x2 * x2_y)) / (Sqr(1 - x1_x2 ^ 2) * Sqr(1 - x2_y ^ 2))
End Function

Function SEMIPARTIAL(x1_y, x2_y, x1_x2)
    ' Returns the semipartial correlation between X1 and Y
    SEMIPARTIAL = (x1_y - (x1_x2 * x2_y)) / Sqr(1 - x1_x2 ^ 2)
End Function
Function r_adj(r, Optional rel_x = 1, Optional rel_y = 1)
    ' Returns r Pearson correlation, adjusted for reliability of X and Y
    If r <> "" Then
        r_adj = r / Sqr(rel_x * rel_y)
    ElseIf r = 1 Then
        r_adj = 1
    Else
        r_adj = ""
    End If
End Function

Function VAR_NULLDIST(scale_pts)
    ' Returns the variance of a null distribution, based on the number of scale points
    VAR_NULLDIST = (scale_pts ^ 2 - 1) / 12
End Function

Function RESCALE(x, dataset As Range, new_min, new_max)
    ' Rescales the value of x based on the desired new range, defined by a new minimum and new maximum

    old_min = WorksheetFunction.Min(dataset.Value)
    old_max = WorksheetFunction.Max(dataset.Value)

    x = x - old_min
    
    x = (x * (new_max - new_min)) / (old_max - old_min)
    RESCALE = x + new_min
    
End Function

Function PHI_RATIO()
    ' Returns phi (the Golden Ratio)
    PHI_RATIO = (1 + Sqr(5)) / 2
End Function
