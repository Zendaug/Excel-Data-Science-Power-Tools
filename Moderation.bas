Attribute VB_Name = "Moderation"
Function ICC_2(ICC1, group_size)
    ICC_2 = (group_size * ICC1) / (1 + (group_size - 1) * ICC1)
End Function

Function GRADE(mark)
    If mark = 0 Then GRADE = "NA"
    If mark > 0 And mark < 50 Then GRADE = "N"
    If mark >= 50 And mark < 60 Then GRADE = "P"
    If mark >= 60 And mark < 70 Then GRADE = "C"
    If mark >= 70 And mark < 80 Then GRADE = "D"
    If mark >= 80 Then GRADE = "HD"
End Function


Function MODERATE(mark, Optional assessment = 50, Optional pass = 50)
    MODERATE = False
    If mark >= 45 And mark < 50 Then MODERATE = True
    If mark = 58 Or mark = 59 Or mark = 68 Or mark = 69 Or mark = 78 Or mark = 79 Then MODERATE = True
    If assessment < pass Then MODERATE = True
End Function

' Adjust the moderated mark up or down based on whether a set of criteria have been fulfilled
Function ADJUST(mark, criteria As Boolean)
    ADJUST = mark

    If mark > 45 And mark < 50 Then
        If criteria = True Then
            ADJUST = 50
        Else
            ADJUST = 45
        End If
    End If
    
    If mark = 58 Or mark = 59 Then
        If criteria = True Then
            ADJUST = 60
        Else
            ADJUST = 57
        End If
    End If
    
    If mark = 68 Or mark = 69 Then
        If criteria = True Then
            ADJUST = 70
        Else
            ADJUST = 67
        End If
    End If
    
    If mark = 78 Or mark = 79 Then
        If criteria = True Then
            ADJUST = 80
        Else
            ADJUST = 77
        End If
    End If
End Function


