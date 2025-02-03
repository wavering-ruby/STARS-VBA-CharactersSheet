Attribute VB_Name = "DiceRolls"
Function random_between(min As Integer, max As Integer) As Integer
    random_between = Int((max - min + 1) * Rnd + min)
End Function

Function text_right(cell As Range)
    Dim index As Integer
    Dim text As String
    
    text = cell.Value
    
    index = InStr(text, "d")
    
    If index < 0 Then
        text_right = ""
    Else
        text_right = Mid(text, index + 1)
    End If
    
End Function

Function concat_array(dices_values() As Integer) As String
    Dim i As Integer
    Dim result As String
    
    result = ""
    
    ' -1 for the array length
    For i = LBound(dices_values) To UBound(dices_values) - 1
        result = result & dices_values(i) & " "
    Next i
    
    ' Just to personalize the view
    result = "Random dices rolled: [ " & result & "]"
    
    concat_array = Trim(result)
End Function


Sub dice_roll()

' Starting with random number
Randomize

' Getting the quantity of dices to roll and the value of dice
qtd_dices = Range("L45").Value
Dim value_dice As Integer
value_dice = Int(text_right(Range("L43")))
ReDim dices_values(qtd_dices) As Integer

' Debug message
If IsEmpty(qtd_dices) Then
    MsgBox "O valor é 0"
Else
    MsgBox "O valor na L45 é " & qtd_dices
End If

If (IsEmpty(qtd_dices) Or qtd_dices = 0) Then
    dice_valors = random_between(1, value_dice)
Else
    For i = 1 To qtd_dices
        dices_values(i - 1) = random_between(1, value_dice)
    Next i
End If

' Debug message to view the stored values in the array
For i = 0 To qtd_dices - 1
    MsgBox "O índice " & i & " contém o valor: " & dices_values(i)
Next i

MsgBox concat_array(dices_values)
End Sub
