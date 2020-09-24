Attribute VB_Name = "modRGB"
Public Sub GetRGB(ByVal Collective As Long, red As Long, green As Long, blue As Long)
If Collective < 0 Then Collective = RGB(105, 105, 255) 'System color replacer
Dim x As Long
x = Int(Collective / 65536)
blue = x
Collective = Collective - x * 65536
x = Int(Collective / 256)
green = x
red = Collective - x * 256
End Sub
