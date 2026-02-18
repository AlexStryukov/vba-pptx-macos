Attribute VB_Name = "FormulaModule"

' ==========================================================
' ОСНОВНАЯ ФУНКЦИЯ ДЛЯ ВСТАВКИ ФОРМУЛЫ
' Можно вызывать из других макросов
' ==========================================================
Function CreateFormula(mathText As String, posX As Single, posY As Single) As Shape
    Dim templateShape As Shape
    Dim newShape As Shape
    Dim sld As Slide
    
    ' 1. Ищем шаблон на 1-м слайде
    On Error Resume Next
    Set templateShape = ActivePresentation.Slides(1).Shapes("MathTemplate")
    On Error GoTo 0
    
    If templateShape Is Nothing Then
        MsgBox "Ошибка: Шаблон 'MathTemplate' не найден на Слайде 1!"
        Exit Function
    End If
    
    Set sld = Application.ActiveWindow.View.Slide
    
    ' 2. Клонируем через буфер обмена
    templateShape.Copy
    MacWait 0.1
    
    On Error Resume Next
    Set newShape = sld.Shapes.Paste(1)
    On Error GoTo 0
    
    If newShape Is Nothing Then Set newShape = sld.Shapes(sld.Shapes.Count)
    
    ' 3. Настраиваем объект
    With newShape
        .Left = posX
        .Top = posY
        .TextFrame.TextRange.Text = mathText
        .TextFrame.TextRange.Font.Size = 24
        .Select
    End With
    
    MacWait 0.2
    
    ' 4. Компилируем в Professional вид
    On Error Resume Next
    Application.CommandBars.ExecuteMso "EquationProfessional"
    On Error GoTo 0
    
    ' Возвращаем объект, чтобы с ним можно было работать дальше
    Set CreateFormula = newShape
End Function

' ==========================================================
' ПРИМЕР ИСПОЛЬЗОВАНИЯ (На месте)
' ==========================================================
Sub ExampleUsage()
    ' Пример создания одной формулы
    CreateFormula "x = \sqrt(a^2 + b^2)", 100, 100
    
    ' Пример создания второй формулы
    CreateFormula "E = mc^2", 100, 200
    
    ' Снимаем выделение в конце
    Application.ActiveWindow.Selection.Unselect
End Sub

' Вспомогательная функция ожидания для Mac
Sub MacWait(Seconds As Single)
    Dim endT As Single
    endT = Timer + Seconds
    Do While Timer < endT
        DoEvents
    Loop
End Sub
