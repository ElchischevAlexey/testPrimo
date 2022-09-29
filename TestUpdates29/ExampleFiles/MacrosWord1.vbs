Sub Replace_Example2()

Dim NewString As String
Dim MyString As String
Dim FindString As String
Dim ReplaceString As String

MyString = "India is a developing country and India is the Asian Country"
FindString = "India"
ReplaceString = "Bharath"

NewString = Replace(MyString, FindString, ReplaceString, Start:=34)

'MsgBox NewString

End Sub
