Function ColNoToLetter(columnNumber As Integer) As String
    Dim columnLetter As String
    Dim dividend As Integer
    Dim modulus As Integer
    
    columnLetter = ""
    dividend = columnNumber
        
    Do While dividend > 0
        modulus = (dividend - 1) Mod 26
        columnLetter = Chr(65 + modulus) & columnLetter
        dividend = Int((dividend - modulus) / 26)
    Loop
        
    ColNoToLetter = columnLetter
End Function
