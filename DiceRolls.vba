Attribute VB_Name = "DiceRolls"
Function RandomBetween(Min As Integer, Max As Integer) As Integer
    RandomBetween = Int((Max - Min + 1) * Rnd + Min)
End Function

Function TextRight(cell As Range) As Integer
    Dim index As Integer
    Dim text As String
    
    text = cell.Value
    
    index = InStr(text, "d")
    
    If index < 0 Then
        TextRight = ""
    Else
        TextRight = Mid(text, index + 1)
    End If
    
End Function

Function Sort(dice_values() As Integer, n As Integer) As Integer()
    Dim i As Integer
    Dim j As Integer
    Dim temp As Integer
    
    If n = 0 Then
        Sort = dice_values
    Else
        For i = 0 To n - 1
            For j = 0 To n - 2
                If dice_values(j) < dice_values(j + 1) Then
                    temp = dice_values(j)
                    dice_values(j) = dice_values(j + 1)
                    dice_values(j + 1) = temp
                End If
            Next j
        Next i
    End If
    
    Sort = dice_values
End Function

Function ConcatArray(dices_values() As Integer, n As Integer) As String
    Dim i As Integer
    Dim result As String
    
    result = ""
    
    ' -1 for the array length
    For i = LBound(dices_values) To n - 1
    
        If i = n - 1 Then
            result = result & dices_values(i)
        Else
            result = result & dices_values(i) & ", "
        End If
    Next i
    
    ' Just to personalize the view
    result = " [ " & result & " ]"
    
    ConcatArray = Trim(result)
End Function

Function Sum(dices_values() As Integer, n As Integer) As Integer
    Dim sum_total
    
    For i = 0 To n - 1
        sum_total = sum_total + dices_values(i)
    Next i
    
    ' Just for returnig the total of sum
    Sum = sum_total
End Function

Function Max(dices_values() As Integer, n As Integer) As Integer
    ' First, this function sort the array to get the first (high)

    sorted_array = Sort(dices_values(), n)
    
    ' After this, just get the first value (the highest)
    Max = sorted_array(0)
End Function

Function Min(dices_values() As Integer, n As Integer) As Integer
    ' First, this function sort the array to get the last value (low)
    sorted_array = Sort(dices_values(), n)
    
    ' If the array isn't empty, they get the last value
    If n > 0 Then
        Min = sorted_array(n - 1)
    Else
        Min = sorted_array(0)
    End If
End Function

Function Vlookup(search_value As Variant, table As Range, return_column As Integer) As Variant
    Dim i As Integer
    Dim num_lines As Integer
    
    ' Receive the line numbers of the table
    num_lines = table.Rows.Count
    
    For i = 1 To num_lines
        ' Verifica se o valor na primeira coluna corresponde ao valor procurado
        If table.Cells(i, 1).Value = search_value Then
            ' Return the value os spedified column
             Vlookup = table.Cells(i, return_column).Value
            Exit Function
        End If
    Next i
    
    ' If not search the value, just return 0
    Vlookup = 0
End Function

Sub DiceRoll()

' Starting with random number
Randomize

' Getting the quantity of dices to roll and the value of dice
Dim qtd_dices As Integer
qtd_dices = Int(Range("L45").Value)

' If the user don't write a number will be placed 1
If qtd_dices = 0 Then
    qtd_dices = 1
End If

Dim value_dice As Integer
value_dice = Int(TextRight(Range("L43")))
ReDim dices_values(qtd_dices) As Integer

Dim sorted_dices() As Integer

' Config to lookup value in table
Dim tableExpertise As Range
Dim tableAttributes As Range
Dim result As Variant
Dim search_value As Variant
Dim return_column As Integer ' For the column that a want to return the value

' Definitions of variable
search_value = Range("L49").Value
Set tableExpertise = Range("R2:V30")
return_column = 5

' The value of expertise (perícia)
resultExpertise = Vlookup(search_value, tableExpertise, return_column)

' The value of attribute
' To the pure attibute
search_valueA = Range("M49").Value ' A variable exclusive to the table attributes
Set tableAttributes = Range("L2:P9")

resultAttribute = Vlookup(search_valueA, tableAttributes, return_column)

If (IsEmpty(qtd_dices) Or qtd_dices = 0) Then
    dice_valors = RandomBetween(1, value_dice)
Else
    For i = 1 To qtd_dices
        dices_values(i - 1) = RandomBetween(1, value_dice)
    Next i
End If

roll_type = Range("L47").Value
Dim flat_dmg As Variant
flat_dmg = Range("M43").Value

' Organizing the array in ascend numbers
sorted_dices = Sort(dices_values(), Int(qtd_dices))

If Not IsEmpty(roll_type) Then
    If roll_type = "Dano" Then
        If IsEmpty(search_value) And IsEmpty(search_valueA) Then
            write_string = Sum(sorted_dices, Int(qtd_dices)) & " <- " & ConcatArray(sorted_dices, Int(qtd_dices))
        ElseIf Not IsEmpty(search_valueA) Then
            write_string = Sum(sorted_dices, Int(qtd_dices)) + flat_dmg + resultAttribute & " <- " & ConcatArray(sorted_dices, Int(qtd_dices)) & " + " & flat_dmg & " + " & resultAttribute
        Else
            write_string = Sum(sorted_dices, Int(qtd_dices)) + flat_dmg & " <- " & ConcatArray(sorted_dices, Int(qtd_dices)) & " + " & flat_dmg
        End If
    ElseIf roll_type = "Desvantagem" Then
        If IsEmpty(search_value) And IsEmpty(search_valueA) Then
            write_string = Min(sorted_dices, Int(qtd_dices)) & " <- " & ConcatArray(sorted_dices, Int(qtd_dices))
        ElseIf Not IsEmpty(search_valueA) Then
            write_string = Min(sorted_dices, Int(qtd_dices)) + resultAttribute & " <- " & ConcatArray(sorted_dices, Int(qtd_dices)) & " + " & resultAttribute
        Else
            write_string = Min(sorted_dices, Int(qtd_dices)) + resultExpertise & " <- " & ConcatArray(sorted_dices, Int(qtd_dices)) & " + " & resultExpertise
        End If
    ElseIf roll_type = "Vantagem" Then
        If IsEmpty(search_value) And IsEmpty(search_valueA) Then
            write_string = Max(sorted_dices, Int(qtd_dices)) & " <- " & ConcatArray(sorted_dices, Int(qtd_dices))
        ElseIf Not IsEmpty(search_valueA) Then
            write_string = Max(sorted_dices, Int(qtd_dices)) + resultAttribute & " <- " & ConcatArray(sorted_dices, Int(qtd_dices)) & " + " & resultAttribute
        Else
            write_string = Max(sorted_dices, Int(qtd_dices)) + resultExpertise & " <- " & ConcatArray(sorted_dices, Int(qtd_dices)) & " + " & resultExpertise
        End If
    End If
Else
    If IsEmpty(search_value) And IsEmpty(search_valueA) Then
        write_string = Sum(sorted_dices, qtd_dices) & " <- " & ConcatArray(sorted_dices, Int(qtd_dices))
    ElseIf Not IsEmpty(search_valueA) Then
        write_string = Sum(sorted_dices, Int(qtd_dices)) + resultAttribute & " <- " & ConcatArray(sorted_dices, Int(qtd_dices)) & " + " & resultAttribute
    Else
        write_string = Sum(sorted_dices, qtd_dices) + resultExpertise & " <- " & ConcatArray(sorted_dices, Int(qtd_dices)) & " + " & resultExpertise
    End If
End If

Range("N43").Value = write_string
End Sub
