# Excel Functions

## Age Range

This function takes an integer input from a cell, which represents the age of the character, and uses conditional comparisons to determine the character's age group.

```Excel
=IF(AND(G6>=8, G6<=12), "Child", IF(AND(G6>=13, G6<=17), "Teenager", IF(AND(G6>=18, G6<=24), "Young Adult", IF(AND(G6>=25, G6<=39), "Adult", IF(AND(G6>=40, G6<=64), "Middle Aged", IF(G6>=65, "Senior", "Baby"))))))
```

- **G6:** The age of the character.
- **Output:** Returns one of the following strings based on the age:
    - "Child" (ages 8-12)
    - "Teenager" (ages 13-17)
    - "Young Adult" (ages 18-24)
    - "Adult" (ages 25-39)
    - "Middle Aged" (ages 40-64)
    - "Senior" (ages 65+)
    - "Baby" (if age is below 8)

---

## Total Attributes

This function is used to calculate the total of a character’s attributes by summing values from multiple columns.

```Excel
=Básico!$M3 + N3 + Básico!$O3
```

- Básico!$M3: Attribute value from the "Básico" sheet, cell M3.
- N3: Local value in the current sheet for the attribute.
- Básico!$O3: Attribute value from the "Básico" sheet, cell O3.
- Output: The sum of the three values, representing the total attribute value.

---

## HP Calculation (PV)

This function calculates a character's Health Points (HP or PV). The formula includes variables like class, Constitution attribute, and Chaotic Exposure percentage. It also accounts for the "Robusto" subclass of Duelists, which modifies the calculation when the character reaches a certain threshold.

```Excel
=IFERROR(IF(AND(G20="Robusto", G22>=15%), (VLOOKUP($G$18, Classe, 2, TRUE) + (P8*2)) + ((G22/5%) * (VLOOKUP($G$18, Classe, 5, TRUE) + (P8*2))), (VLOOKUP($G$18, Classe, 2, TRUE) + P8) + (((G22-5%)/5%) * (VLOOKUP($G$18, Classe, 5, TRUE) + P8))), 0)
```

- G20: The subclass of the character (e.g., "Robusto").
- G22: The percentage of Chaotic Exposure for the character.
- $G$18: The character's class.
- Classe: A table that contains class-related values (e.g., HP values).
- P8: Constitution attribute value.
- Output: The character's HP (PV) based on their class, Constitution, and Chaotic Exposure. For "Robusto" subclasses, the formula adjusts when exposure exceeds 15%, doubling the Constitution bonus.

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

' The value of expertise (per cia)
result = Vlookup(search_value, table, return_column)

' Debug message
'If IsEmpty(qtd_dices) Then
    'MsgBox "O valor   " & qtd_dices
'Else
    'MsgBox "O valor na L45   " & qtd_dices
'End If

If (IsEmpty(qtd_dices) Or qtd_dices = 0) Then
    dice_valors = RandomBetween(1, value_dice)
Else
    For i = 1 To qtd_dices
        dices_values(i - 1) = RandomBetween(1, value_dice)
    Next i
End If

' Debug message to view the stored values in the array
'For i = 0 To qtd_dices - 1
'    MsgBox "O  ndice " & i & " cont m o valor: " & dices_values(i)
'Next i

'MsgBox Sort(dices_values, Int(qtd_dices))

roll_type = Range("L47").Value
Dim flat_dmg As Variant
flat_dmg = Range("M43").Value

' Organizing the array in ascend numbers
sorted_dices = Sort(dices_values(), Int(qtd_dices))

'MsgBox Sorted_dices(0)

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

In the Excel, are defined what dice the user wants (d2, d3, d4, d6, d8, d10, d12, d20, d100) and this function gets the right side of the string and return as int to use a Max parameter in the function [RandomBetween](#randombetween). This function is called once.

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

## Sort

This is one of the most important function in VBA code, because all the element in the array is ordered in descending, and with this, the [Min](#min) and [Max](#max) function get his value more easile and [ConcatArray](#concatarray) gets a tidy array.

```VBA
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
```

---

## ConcatArray

Return a string that is used to write in a cell all the dices rolled randomly. This function is needed to show in a beatiful way the array elements to the user.

```VBA
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
```

---

## Sum

A function that return all the values summed of a array. This function it's used to the "Damage" dices.

```VBA
Function Sum(dices_values() As Integer, n As Integer) As Integer
    Dim sum_total
    
    For i = 0 To n - 1
        sum_total = sum_total + dices_values(i)
    Next i
    
    ' Just for returnig the total of sum
    Sum = sum_total
End Function
```

---

## Max

Get's the bigger element in a array. Because the array it's actually sorted, the function just return the first value of the array.

```VBA
Function Max(dices_values() As Integer, n As Integer) As Integer
    sorted_array = Sort(dices_values(), n)
    Max = sorted_array(0)
End Function
```

---

## Min

Get's the lowest element in a array.

```VBA
Function Min(dices_values() As Integer, n As Integer) As Integer
    sorted_array = Sort(dices_values(), n)
    
    If n > 0 Then
        Min = sorted_array(n - 1)
    Else
        Min = sorted_array(0)
    End If
End Function
```

---

## Vlookup

This is a code that imit a VLOOKUP function in the Excel. Utilizing a table defined in [DiceRolls](#dicerolls-(sub)) this code search for a string value in a determined range of cells and return a value As Variant for the function that is the result of Expertise values determined by the user.

```VBA
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

```

---

# Versions

## 2.0.
- :envelope_with_arrow: [Download](./Ficha-de-STARS-2.0.xlsx)
- Updated to the new book version.

## 3.0.
- :envelope_with_arrow: [Download](./Ficha-de-STARS-3.0.xlsx)
- Updated to the new book version.

## 3.1.
- :envelope_with_arrow: [Download](./Ficha-de-STARS-3.1.xlsx)
- Updated some cells.

## 3.2.
- :envelope_with_arrow: [Download](./Ficha-de-STARS-3.2.xlsx)
- Fixed the calculate of PV, PP and PS;
- Fixed the calculate of equipment cost;
- Added a template for the Power Abilities on the first page.

## 3.3.
- :envelope_with_arrow: [Download](./Ficha-de-STARS-3.3.xlsx)
- Added PV calculate for Robust;
- Adapted the formulas of equipment cost to old versions of Excel;
- Added displacement when selected the subclass "Elite Trooper".

## 3.4.
- :envelope_with_arrow: [Download](./Ficha-de-STARS-3.4.xlsx)
- Fixed PV, PP and PS points at 5%;
- Added dropdown list for class names.

## 3.5.
- :envelope_with_arrow: [Download](./Ficha-de-STARS-3.5.xlsx)
- Fixed the power DT;

## 4.0.
- :envelope_with_arrow: [Download](./Ficha-de-STARS-4.0.xlsx)
- Fixed the calculation of displacement and now it's possible to change the value from "m" to squares;
- Added calculator for resistance.

## 4.1.
- :envelope_with_arrow: [Download](./Ficha-de-STARS-4.1.xlsm)
- Now you can roll dice on the character sheet;
- Added calculator for reactions (blocking and dodging) and passive perception;
- Fixed the PP Limit and added a sum with Power attribute if the character's subclass is "Path of Agony".
