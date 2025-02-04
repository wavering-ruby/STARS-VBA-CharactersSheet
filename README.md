# Excel Functions

---

# VBA Functions

## DiceRolls (Sub)

This function makes the mechanic of roll dice in a Excel file. Enabling to the player use just one Sheet and one Excel file to player with your characters. Reduzing unnecessary processing. 

```VBA
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
Dim table As Range
Dim result As Variant
Dim search_value As Variant
Dim return_column As Integer ' For the column that a want to return the value

' Definitions of variable
search_value = Range("L49").Value
Set table = Range("R2:V30")
return_column = 5

' The value of expertise (perícia)
result = Vlookup(search_value, table, return_column)

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
        If IsEmpty(flat_dmg) Then
            write_string = Sum(sorted_dices, Int(qtd_dices)) & " <- " & ConcatArray(sorted_dices, Int(qtd_dices))
        Else
            write_string = Sum(sorted_dices, Int(qtd_dices)) + flat_dmg & " <- " & ConcatArray(sorted_dices, Int(qtd_dices)) & " + " & flat_dmg
        End If
    ElseIf roll_type = "Desvantagem" Then
        If IsEmpty(search_value) Then
            write_string = Min(sorted_dices, Int(qtd_dices)) & " <- " & ConcatArray(sorted_dices, Int(qtd_dices))
        Else
            write_string = Min(sorted_dices, Int(qtd_dices)) + result & " <- " & ConcatArray(sorted_dices, Int(qtd_dices)) & " + " & result
        End If
    ElseIf roll_type = "Vantagem" Then
        If IsEmpty(search_value) Then
            write_string = Max(sorted_dices, Int(qtd_dices)) & " <- " & ConcatArray(sorted_dices, Int(qtd_dices))
        Else
            write_string = Max(sorted_dices, Int(qtd_dices)) + result & " <- " & ConcatArray(sorted_dices, Int(qtd_dices)) & " + " & result
        End If
    End If
Else
    If IsEmpty(search_value) Then
        write_string = Sum(sorted_dices, qtd_dices) & " <- " & ConcatArray(sorted_dices, Int(qtd_dices))
    Else
        write_string = Sum(sorted_dices, qtd_dices) + result & " <- " & ConcatArray(sorted_dices, Int(qtd_dices)) & " + " & result
    End If
End If

Range("N43").Value = write_string
End Sub
```

---

## RandomBetween

This function is core for the project, because this function draws random values ​​that will be stored in an array with the length determined by the quantity of pieces that user wants to roll.

``` VBA
Function RandomBetween(Min As Integer, Max As Integer) As Integer
    RandomBetween = Int((Max - Min + 1) * Rnd + Min)
End Function
```

---

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

## 

# Versions

## 4.0.

### 4.1.
- Now you can roll dices in the characters sheet;
- Added calculator of reactions (blocking and dodging) and passive perception;
- Fixed the PP Limit and added sum with Power attribute if the character's subclass is "Path of Agony".
