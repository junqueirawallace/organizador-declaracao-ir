Attribute VB_Name = "ModuloNavegacao"
Option Explicit

' Macro opcional para Excel desktop.
' O Excel para iPad não executa macros VBA.
Public Sub MoverIconeParaPosicao()
    Dim shp As Shape
    Dim ws As Worksheet
    Dim nomeIconeProcurado As String
    Dim novaPosicaoX As Double
    Dim novaPosicaoY As Double

    Set ws = ActiveSheet
    nomeIconeProcurado = "IconeInicio" ' ajuste ao nome da forma
    novaPosicaoX = 20
    novaPosicaoY = 20

    On Error Resume Next
    Set shp = ws.Shapes(nomeIconeProcurado)
    On Error GoTo 0

    If shp Is Nothing Then
        MsgBox "Não encontrei a forma '" & nomeIconeProcurado & "' na aba '" & ws.Name & "'.", vbExclamation
        Exit Sub
    End If

    shp.Left = novaPosicaoX
    shp.Top = novaPosicaoY
    MsgBox "Ícone reposicionado com sucesso.", vbInformation
End Sub
