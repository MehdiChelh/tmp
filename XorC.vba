Option Explicit 
Sub test() 
     'this sub is only present to demonstrate use of the function!
     'it is not required to use the function.
    Dim r As Range, retVal, sKey As String 
    sKey = Application.InputBox("Enter your key", "Key entry", "My Key", , , , , 2) 
    retVal = MsgBox("This is the key you entered:" & vbNewLine & Chr$(34) & sKey & Chr$(34) & vbNewLine & _ 
    "Please confirm OK or Cancel to exit", vbOKCancel, "Confirm Key") 
    If retVal = vbCancel Then Exit Sub 
    For Each r In Sheets("Sheet1").UsedRange 
        If r.Interior.ColorIndex = 6 Then 
            r.Value = XorC(r.Value, sKey) 
        End If 
    Next r 
End Sub 
 
Function XorC(ByVal sData As String, ByVal sKey As String) As String 
    Dim l As Long, i As Long, byIn() As Byte, byOut() As Byte, byKey() As Byte 
    Dim bEncOrDec As Boolean 
     'confirm valid string and key input:
    If Len(sData) = 0 Or Len(sKey) = 0 Then XorC = "Invalid argument(s) used": Exit Function 
     'check whether running encryption or decryption (flagged by presence of "xxx" at start of sData):
    If Left$(sData, 3) = "xxx" Then 
        bEncOrDec = False 'decryption
        sData = Mid$(sData, 4) 
    Else 
        bEncOrDec = True 'encryption
    End If 
     'assign strings to byte arrays (unicode)
    byIn = sData 
    byOut = sData 
    byKey = sKey 
    l = LBound(byKey) 
    For i = LBound(byIn) To UBound(byIn) - 1 Step 2 
        byOut(i) = ((byIn(i) + Not bEncOrDec) Xor byKey(l)) - bEncOrDec 'avoid Chr$(0) by using bEncOrDec flag
        l = l + 2 
        If l > UBound(byKey) Then l = LBound(byKey) 'ensure stay within bounds of Key
    Next i 
    XorC = byOut 
    If bEncOrDec Then XorC = "xxx" & XorC 'add "xxx" onto encrypted text
End Function 
