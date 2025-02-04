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
- Updated to the new book version.

## 3.1.
- Updated some cells.

## 3.2. 
- Fixed the calculate of PV, PP and PS;
- Fixed the calculate of equipament cust;
- Added a template for the Power Abilities on the first page.

## 3.3.
- Added PV calculate for Robust;
- Adaptade the formules of equipamente cust to old versions of Excel;
- Added deslocation when selected the subclasse "Elite Trooper".

## 3.4. 
- Fixed PV, PP and PS points at 5%
- Added suspense list for classes names.

## 3.5 
- Fixed the power DT;

## 4.0.
- Fixed the calculate of deslocation e now it's possible to change the value in "m" to squares;
- Added calculator for resistence.

## 4.1.
- Now you can roll dices in the characters sheet;
- Added calculator of reactions (blocking and dodging) and passive perception;
- Fixed the PP Limit and added sum with Power attribute if the character's subclass is "Path of Agony".
