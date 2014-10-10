Attribute VB_Name = "UnLoadAllFrms"
Public Sub UnloadAll()
Dim f As Integer
    f = Forms.Count
    Do While f > 0
        Unload Forms(f - 1)
        If f = Forms.Count Then Exit Do
        f = f - 1
    Loop
 End Sub

