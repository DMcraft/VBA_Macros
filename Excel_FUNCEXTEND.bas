Attribute VB_Name = "FUNCEXTEND"
Function СУММАПРОПИСЬЮ(n As Double, Optional padezh As Byte = 0) As String
 
 Dim Nums1, Nums2, Nums3, Nums4 As Variant
 
 If padezh = 0 Then
    Nums1 = Array("", "один ", "два ", "три ", "четыре ", "пять ", "шесть ", "семь ", "восемь ", "девять ")
    Nums2 = Array("", "десять ", "двадцать ", "тридцать ", "сорок ", "пятьдесят ", "шестьдесят ", "семьдесят ", "восемьдесят ", "девяносто ")
    Nums3 = Array("", "сто ", "двести ", "триста ", "четыреста ", "пятьсот ", "шестьсот ", "семьсот ", "восемьсот ", "девятьсот ")
    Nums4 = Array("", "одна ", "две ", "три ", "четыре ", "пять ", "шесть ", "семь ", "восемь ", "девять ")
    Nums5 = Array("десять ", "одиннадцать ", "двенадцать ", "тринадцать ", "четырнадцать ", "пятнадцать ", "шестнадцать ", "семнадцать ", "восемнадцать ", "девятнадцать ")
 Else
    Nums1 = Array("", "одном ", "двух ", "трёх ", "четырёх ", "пяти ", "шести ", "семи ", "восьми ", "девяти ")
    Nums2 = Array("", "десяти ", "двадцати ", "тридцати ", "сорока ", "пятидесяти ", "шестидесяти ", "семидясяти ", "восьмидесяти ", "девяноста ")
    Nums3 = Array("", "ста ", "двухста ", "трёхста ", "четырёхста ", "пятиста ", "шестиста ", "семиста ", "восьмиста ", "девяноста ")
    Nums4 = Array("", "одной ", "двух ", "трёх ", "четырёх ", "пяти ", "шести ", "семи ", "восьми ", "девяти ")
    Nums5 = Array("десяти ", "одиннадцати ", "двенадцати ", "тринадцати ", "четырнадцати ", "пятнадцати ", "шестнадцати ", "семнадцати ", "восемнадцати ", "девятнадцати ")
 End If
 
 If n <= 0 Then
   СУММАПРОПИСЬЮ = "ноль"
   Exit Function
 End If
 'разделяем число на разряды, используя вспомогательную функцию Class
 ed = Class(n, 1)
 dec = Class(n, 2)
 sot = Class(n, 3)
 tys = Class(n, 4)
 dectys = Class(n, 5)
 sottys = Class(n, 6)
 mil = Class(n, 7)
 decmil = Class(n, 8)
 'проверяем миллионы
 Select Case decmil
   Case 1
     mil_txt = Nums5(mil) & "миллионов "
     GoTo www
   Case 2 To 9
     decmil_txt = Nums2(decmil)
 End Select
 Select Case mil
   Case 1
     mil_txt = Nums1(mil) & "миллион "
   Case 2, 3, 4
     mil_txt = Nums1(mil) & "миллиона "
   Case 5 To 20
     mil_txt = Nums1(mil) & "миллионов "
 End Select
www:
 sottys_txt = Nums3(sottys)
 'проверяем тысячи
 Select Case dectys
   Case 1
     tys_txt = Nums5(tys) & "тысяч "
     GoTo eee
   Case 2 To 9
     dectys_txt = Nums2(dectys)
 End Select
 Select Case tys
   Case 0
     If dectys > 0 Then tys_txt = Nums4(tys) & "тысяч "
   Case 1
     tys_txt = Nums4(tys) & "тысяча "
   Case 2, 3, 4
     tys_txt = Nums4(tys) & "тысячи "
   Case 5 To 9
     tys_txt = Nums4(tys) & "тысяч "
 End Select
 If dectys = 0 And tys = 0 And sottys <> 0 Then sottys_txt = sottys_txt & " тысяч "
eee:
 sot_txt = Nums3(sot)
 'проверяем десятки
 Select Case dec
   Case 1
     ed_txt = Nums5(ed)
     GoTo rrr
   Case 2 To 9
     dec_txt = Nums2(dec)
 End Select
 ed_txt = Nums1(ed)
rrr:
 'формируем итоговую строку
 СУММАПРОПИСЬЮ = decmil_txt & mil_txt & sottys_txt & dectys_txt & tys_txt & sot_txt & dec_txt & ed_txt
End Function
'вспомогательная функция для выделения из числа разрядов
Private Function Class(M, I)
  Class = Int(Int(M - (10 ^ I) * Int(M / (10 ^ I))) / 10 ^ (I - 1))
End Function

' удаление гиперсылок в документе
Sub удаление_гиперсылок()
 For Each sh In ThisWorkbook.Worksheets
 sh.Hyperlinks.Delete
 Next
End Sub

' добавление_строк Макрос
Sub добавление_строк()
Do While ActiveCell.Value > ""
r = ActiveCell.Row + 1
c = ActiveCell.Column
Cells(r, c).Activate
    Selection.Insert Shift:=xlDown
    Selection.Insert Shift:=xlDown
    Selection.Insert Shift:=xlDown
    Selection.Insert Shift:=xlDown
r = ActiveCell.Row + 4
c = ActiveCell.Column
Cells(r, c).Activate
Loop

End Sub

'Обработака регулярных выражений
' https://newtechaudit.ru/obrabotka-teksta-v-excel-regulyarnye-vyrazheniya/
Public Function RegExp(Text As String, Pattern As String) As String
    On Error GoTo ErrorHandler
    Set newObj = CreateObject("VBScript.RegExp")
    newObj.Pattern = Pattern
    newObj.Global = True
    If newObj.Test(Text) Then
        Set matches = newObj.Execute(Text)
        RegExp = matches.Item(0)

        Exit Function
    End If
ErrorHandler:
    RegExp = CVErr(xlErrValue)
End Function



