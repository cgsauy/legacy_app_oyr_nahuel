Attribute VB_Name = "ModFormularios"
Option Explicit

Public Sub CentroForm(Formulario As Form)
On Error Resume Next
    With Formulario
        .Move (Screen.Width - .Width) / 2, (Screen.Height - .Height) / 2
    End With
End Sub

Public Sub ObtengoSeteoForm(Formulario As Form, Optional LeftIni As Currency = 0, Optional TopIni As Currency = 0, _
                                            Optional WidthIni As Currency = 0, Optional HeightIni As Currency = 0)
    If LeftIni = 0 Then LeftIni = Formulario.Left
    If TopIni = 0 Then TopIni = Formulario.Top
    If WidthIni = 0 Then WidthIni = Formulario.Width
    If HeightIni = 0 Then HeightIni = Formulario.Height
    
    'Busco si tengo seteada la última posición y tamaño del formulario
    'Sino le marco yo los tamaños iniciales. ------------------------------------------
    Formulario.Left = GetSetting(App.Title, "Settings", "AA" & Formulario.Name & "Left", LeftIni)
    Formulario.Top = GetSetting(App.Title, "Settings", "AA" & Formulario.Name & "Top", TopIni)
    Formulario.Width = GetSetting(App.Title, "Settings", "AA" & Formulario.Name & "Width", WidthIni)
    Formulario.Height = GetSetting(App.Title, "Settings", "AA" & Formulario.Name & "Height", HeightIni)
    
End Sub

Public Sub GuardoSeteoForm(Formulario As Form)
    'Guarda la posicion y tamaño del formulario, si su estado es normal.
    If Formulario.WindowState <> vbMinimized And Formulario.WindowState <> vbMaximized Then
        SaveSetting App.Title, "Settings", "AA" & Formulario.Name & "Left", Formulario.Left
        SaveSetting App.Title, "Settings", "AA" & Formulario.Name & "Top", Formulario.Top
        SaveSetting App.Title, "Settings", "AA" & Formulario.Name & "Width", Formulario.Width
        SaveSetting App.Title, "Settings", "AA" & Formulario.Name & "Height", Formulario.Height
    End If
End Sub

Public Sub BorroSeteoForm(Formulario As Form, bLeft As Boolean, bTop As Boolean, bWidth As Boolean, bHeight As Boolean)
    'Borra de la registry el valor de posicion o tamaño del formulario.
    If bLeft Then DeleteSetting App.Title, "Settings", "AA" & Formulario.Name & "Left"
    If bTop Then DeleteSetting App.Title, "Settings", "AA" & Formulario.Name & "Top"
    If bWidth Then DeleteSetting App.Title, "Settings", "AA" & Formulario.Name & "Width"
    If bHeight Then DeleteSetting App.Title, "Settings", "AA" & Formulario.Name & "Height"
End Sub

Public Sub GuardoSeteoControl(ByVal sPropControl As String, sValor As String)
    On Error Resume Next
    SaveSetting App.Title, "Settings", "AA" & sPropControl, sValor
End Sub

Public Function ObtengoSeteoControl(Key As String, Optional sDefault As String) As Variant
    On Error Resume Next
    ObtengoSeteoControl = 0
    ObtengoSeteoControl = GetSetting(App.Title, "Settings", "AA" & Key, sDefault)
End Function

Public Function FormActivo(Nombre As String) As Boolean

Dim f As Form

    FormActivo = False
    For Each f In Forms
        If f.Name = Trim(Nombre) Then
            FormActivo = True: Exit Function
        End If
    Next
    
End Function

Public Function FormActivoCaption(Nombre As String) As Boolean

Dim f As Form

    FormActivoCaption = False
    For Each f In Forms
        If f.Caption = Trim(Nombre) Then
            FormActivoCaption = True
            f.SetFocus
            Exit Function
        End If
    Next
    
End Function
