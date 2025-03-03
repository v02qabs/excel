Attribute VB_Name = "Module1"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
    Dim sunday As Integer
    Weekday ("2022/01/01")
    sunday = Weekday("2000/01/01")
    If sunday = 1 Then
        Cells(1, 1) = "“ú"
    Else
        Cells(1, 1) = "-"
    

        End If
        
End Sub
