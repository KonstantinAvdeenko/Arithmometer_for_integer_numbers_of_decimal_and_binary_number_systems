Dim Char As Single 
Private Sub Command1_Click()         ‘Процедура нажатия на кнопку перевода из 10 системы счисления в 2
If IsNumeric(Text1) = False Then       ‘Проверка на введенное число
MsgBox ("Введите число")
Exit Sub
End If
Dim a As Single, rez() As Single, rezult() As Single, sss As String
a = Text1
k = 1
Label1:
ReDim Preserve rez(1 To k) As Single       ‘этот блок заносит цифры двоичного числа в массив rez. Высчитывается оно по стандартному алгоритму остатков от деления
c = a \ 2          ‘запоминаем целую часть от деления на 2
x = c * 2          ‘умножаем на 2
rez(k) = a – x     ‘разность этого и изначального будет 0 или 1
a = c              ‘полученую целую часть записываем как изначальную
If a = 0 Then GoTo Label2     ‘и идём к циклу 2.Окончанию
k = k + 1                     ‘иначе увеличиваем размерность массива k
If a = 1 Then                 ‘Проверяем,если полученое знач=1 тогда 
ReDim Preserve rez(1 To k) As Single     
‘элемент массива,то есть цифра двоичного числа будет =1
rez(k) = 1
GoTo Label2
End If
GoTo Label1
Label2:
ReDim Preserve rezult(1 To k) As Single   ‘По скольку полученый массив это зеркальное отображение результата,то нужно его обратить 
For i = 1 To k
rezult(k + 1 - i) = rez(i) ‘Массив заполняется в обратном порядке
Next i
For i = 1 To k
sss = sss & rezult(i)         ‘Процедура вывода результата
Next i
Label3 = "Результат: " & sss
End Sub
 
Private Sub Command2_Click()      ‘Процедура нажатия на кнопку перевода из 2 системы счисления в 10
If IsNumeric(Text1) = False Then       ‘Проверка на число
MsgBox ("Введите число")
Exit Sub
End If
For i = 1 To Len(Text1)        ‘ Проверка на 0 или 1
Char = Mid(Text1, i, 1)        
If Char <> "0" Then
If Char <> "1" Then
MsgBox ("Введённое значение должно содержать только 0 или 1")
Exit Sub
End If
End If
Next i
Dim Dvo As String, rezult As Single    
Dvo = Text1
‘перевод в десятичный код осуществляется по простому алгоритму. Это сумма произведений i-ой цифры числа на 2 в степени последний индекс –i 
For i = 1 To Len(Dvo)         
rezult = rezult + Mid(Dvo, i, 1) * (2 ^ (Len(Dvo) - i))
Next i
Label3 = "Результат: " & result      
End Sub

Private Sub Command3_Click()       ‘Эта процедура-смесь первой и второй
‘Смысл действия с двоичными: Переводит в десятичные, осуществляет арифметические действия и переводит обратно 
If IsNumeric(Text1) = False Then
MsgBox ("Введите число")
Exit Sub
End If
If IsNumeric(Text2) = False Then
MsgBox ("Введите число")
Exit Sub
End If
 
For i = 1 To Len(Text1)
Char = Mid(Text1, i, 1)
If Char <> "0" Then
If Char <> "1" Then
MsgBox ("Введённое значение должно содержать только 0 или 1")
Exit Sub
End If
End If
Next i
For i = 1 To Len(Text2)
Char = Mid(Text2, i, 1)
If Char <> "0" Then
If Char <> "1" Then
MsgBox ("Введённое значение должно содержать только 0 или 1")
Exit Sub
End If
End If
Next i
Dim Dvo1 As String, Dvo2 As String, rezult As Single
Dim rez1 As Single, rez2 As Single, rez3 As Single
Dvo1 = Text1
Dvo2 = Text2
For i = 1 To Len(Dvo1)
rez1 = rez1 + Mid(Dvo1, i, 1) * (2 ^ (Len(Dvo1) - i))
Next i
For i = 1 To Len(Dvo2)
rez2 = rez2 + Mid(Dvo2, i, 1) * (2 ^ (Len(Dvo2) - i))
Next i
‘ Арифметические действия с числами выбираемые флажком на форме
If Option4 = True Then
rez3 = rez1 + rez2
End If
If Option5 = True Then
rez3 = rez1 - rez2
End If
If Option6 = True Then
rez3 = rez1 * rez2
End If
If Option7 = True Then
rez3 = rez1 / rez2
End If
 
Dim a As Single, rez() As Single, rezultat() As Single, sss As String
‘ Перевод результата арифметического действия в 2 и 10 системы счисления
a = rez3
k = 1
Label1:
ReDim Preserve rez(1 To k) As Single
c = a \ 2
x = c * 2
rez(k) = a - x
a = c
If a = 0 Then GoTo Label2
k = k + 1
If a = 1 Then
ReDim Preserve rez(1 To k) As Single
rez(k) = 1
GoTo Label2
End If
GoTo Label1
Label2:
ReDim Preserve rezultat(1 To k) As Single
For i = 1 To k
rezultat(k + 1 - i) = rez(i)
Next i
For i = 1 To k
sss = sss & rezultat(i)
Next i
Label3 = "Результат: " & sss
 
End Sub
Private Sub Command4_Click() ‘Кнопка выхода
If MsgBox("Вы желаете выйти?;)", vbOKCancel, "exit") = vbOK Then
Unload Form1
Exit Sub
End If
End Sub

 ‘Процедуры обрабатывающие комбинации нажатых кнопок и активных полей
Private Sub Option1_Click()
If Option1 = True Then
Command2.Enabled = False
Command1.Enabled = True
Command3.Enabled = False
Text1.Enabled = True
Text2.Enabled = False
Frame1.Enabled = False
End If
End Sub
Private Sub Option2_Click()
If Option2 = True Then
Command2.Enabled = True
Command1.Enabled = False
Command3.Enabled = False
Text1.Enabled = True
Text2.Enabled = False
Frame1.Enabled = False
End If
End Sub
Private Sub Option3_Click()
If Option3 = True Then
Command2.Enabled = False
Command1.Enabled = False
Command3.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Frame1.Enabled = True
Option4 = True
End If
End Sub
