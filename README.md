# About the Repository

This is a fan made to the fandom o STARS RPG. Here, you will have a lot of Character's Sheet automatically in Excel to improve your session and organize you character way efficient. Feel free to open a Issue in the case you want a change in the Sheet. If you have a question, please, send me a e-mail and I will answer all your question. **Remember that's it a brazillian (:br:) repository, so the Sheet will be in pt-br**, if you want a translated sheet send me a e-mail.

Have fun and **STARS THE BEST RPG FOREVER**

---

# Summary

├─ [Excel Functions](#excel-functions) <br>
│  ├─ [Age Range](#age-range) <br>
│  ├─ [Total Attributes](#total-attributes) <br>
│  ├─ [HP Calculation (PV)](#hp-calculation-pv) <br>
│  ├─ [PP Calculation (PP)](#pp-calculation-pp) <br>
│  └─ [SP Calculation (PS)](#sp-calculation-ps) <br>
├─ [VBA Functions](#vba-functions) <br>
│  ├─ [Dice Rolls](#dicerolls-sub) <br>
│  ├─ [RandomBetween](#randombetween) <br>
│  ├─ [TextRight](#textright) <br>
│  ├─ [Sort](#sort) <br>
│  ├─ [ConcatArray](#concatarray) <br>
│  ├─ [Sum](#sum) <br>
│  ├─ [Max](#max) <br>
│  ├─ [Min](#min) <br>
│  └─ [Vlookup](#vlookup) <br>
├─ [Versions](#versions) <br>
│  ├─ [2.0.](#20) <br>
│  ├─ [3.0.](#30) <br>
│  ├─ [3.1.](#31) <br>
│  ├─ [3.2.](#32) <br>
│  ├─ [3.3.](#33) <br>
│  ├─ [3.4.](#34) <br>
│  ├─ [3.5.](#35) <br>
│  ├─ [4.0.](#40) <br>
│  ├─ [4.1.](#41) <br>
│  └─ [4.2.](#42) <br>

---

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

## PP Calculation (PP)

This function calculates a character's Health Points (HP or PV). The formula uses variables such as class, Constitution attribute, and Chaotic Exposure percentage. It dynamically adjusts the calculation based on the character's class and exposure level.

```Excel
=SEERRO((PROCV($G$18;Classe;3;VERDADEIRO) + P6) + (((G22 - 5%) / 5%) * (PROCV($G$18;Classe;6;VERDADEIRO) + P6)); 0)
```

- G18: The character's class.
- Classe: A table containing class-related values (e.g., base PP and modifiers).
- P6: Power attribute value.
- G22: The percentage of Chaotic Exposure for the character.
- Output: The character's Power Points (PP) based on their class, Power attribute, and Chaotic Exposure.

---

## SP Calculation (PS)

In construction.

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

This function is core to the project because it generates random values that are stored in an array, with the length determined by the number of pieces the user wants to roll.

``` VBA
Function RandomBetween(Min As Integer, Max As Integer) As Integer
    RandomBetween = Int((Max - Min + 1) * Rnd + Min)
End Function
```

**E.g.:** if the user write d10 in the "Básico" sheet, the maximum value will be "10" because 10 is to the right of "d". This is only possible thanks to [TextRight](#textright) function.

Continuing with the example: 10 is now the maximum value of the die, and the minimum is always 1 (the die value can't be <= 0). The calculation involves a Rnd variable, which returns a random number to the function that multiplies the difference between the maximum and minimum values.

The Rnd function always returns a number >= 0 and < 1.

If the value of Rnd is 0.2, the formula will be "(8) * 1.2". Remember, the Int() function in VBA rounds the value down. So, the final result is: 9.

---

## TextRight

In Excel, the user defines which die they want (d2, d3, d4, d6, d8, d10, d12, d20, d100), and this function extracts the right side of the string and returns it as an integer to use as the Max parameter in the [RandomBetween](#randombetween). function. This function is called once.

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

**E.g.:** when the user writes "d10", this function selects everything to the right of the index where "d" is located. When it returns to the main function, the right side of "d" will be converted to an integer (note that the return type of the function is "As Integer").

---

## Sort

This is one of the most important functions in VBA code because all the elements in the array are ordered in descending order. With this, the [Min](#min) and [Max](#max) function their values more easily, and [ConcatArray](#concatarray) produces a tidy array.

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

**E.g.:** the Sort function will check every single element in the array and will change the values at each index to order them in descending order. For example:

[2, 4, 6, 3]

The array will be ordered as follows:

[6, 5, 4, 2]

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
- :envelope_with_arrow: [Download](./Character's%20Sheet/Ficha-de-STARS-2.0.xlsx)
- Updated to the new book version.

---

## 3.0.
- :envelope_with_arrow: [Download](./Character's%20Sheet/Ficha-de-STARS-3.0.xlsx)
- Updated to the new book version.

---

## 3.1.
- :envelope_with_arrow: [Download](./Character's%20Sheet/Ficha-de-STARS-3.1.xlsx)
- Updated some cells.

---

## 3.2.
- :envelope_with_arrow: [Download](./Character's%20Sheet/Ficha-de-STARS-3.2.xlsx)
- Fixed the calculate of PV, PP and PS;
- Fixed the calculate of equipment cost;
- Added a template for the Power Abilities on the first page.

---

## 3.3.
- :envelope_with_arrow: [Download](./Character's%20Sheet/Ficha-de-STARS-3.3.xlsx)
- Added PV calculate for Robust;
- Adapted the formulas of equipment cost to old versions of Excel;
- Added displacement when selected the subclass "Elite Trooper".

---

## 3.4.
- :envelope_with_arrow: [Download](./Character's%20Sheet/Ficha-de-STARS-3.4.xlsx)
- Fixed PV, PP and PS points at 5%;
- Added dropdown list for class names.

---

## 3.5.
- :envelope_with_arrow: [Download](./Character's%20Sheet/Ficha-de-STARS-3.5.xlsx)
- Fixed the power DT;

---

## 4.0.
- :envelope_with_arrow: [Download](./Character's%20Sheet/Ficha-de-STARS-4.0.xlsx)
- Fixed the calculation of displacement and now it's possible to change the value from "m" to squares;
- Added calculator for resistance.

---

## 4.1.
- :envelope_with_arrow: [Download](./Character's%20Sheet/Ficha-de-STARS-4.1.xlsm)
- Now you can roll dice on the character sheet;
- Added calculator for reactions (blocking and dodging) and passive perception;
- Fixed the PP Limit and added a sum with Power attribute if the character's subclass is "Path of Agony".

---

## 4.2. 
- :envelope_with_arrow: [Download](./Character's%20Sheet/Ficha-de-STARS-4.2.xlsm)
- Added a new tab for the "Passagem" powers.
- In Construction.

---

# Bibliography
- [Int](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/int-fix-functions).
- [Rnd](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/rnd-function).
- [Randomize](https://learn.microsoft.com/pt-br/office/vba/language/reference/user-interface-help/randomize-statement).

---
