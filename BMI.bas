Attribute VB_Name = "Module1"
Sub Height_and_Weight()


'Declare the variables

Dim Height As Variant
Dim Weight As Variant


'Request Input Data

Height = InputBox("Enter Height in Meters")
Weight = InputBox("Enter Weight")

'Verify Data

If IsNumeric(Height) And IsNumeric(Weight) Then
    Range("M17").Value = CDbl(Height)
    Range("M18").Value = CDbl(Weight)
    
    MsgBox "Weight and Height have been added correctly", vbInformation
Else
    MsgBox "Error: Please enter valid numeric values.", vbCritical
End If

End Sub
