# Excel Functions

---

# VBA Functions

## DiceRolls (Sub)

This function makes the mechanic of roll dice in a Excel file. Enabling to the player use just one Sheet and one Excel file to player with your characters. Reduzing unnecessary processing. 

```
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
```

## RandomBetween

This function is core for the project, because this function draws random values ​​that will be stored in an array with the length determined by the quantity of pieces that user wants to roll.

``` VBA
Function RandomBetween(Min As Integer, Max As Integer) As Integer
    RandomBetween = Int((Max - Min + 1) * Rnd + Min)
End Function
```

## TextRight

In the Excel, are defined what dice the user wants (d2, d3, d4, d6, d8, d10, d12, d20, d100) and this function gets the right side of the string and return as int to use a Max parameter in the function [RandomBetween](#randombetween) 

```VBA
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
```

---
