Attribute VB_Name = "M�dulo1"

Private Sub BuscaArquivo_Click()
ActiveSheet.Unprotect
Cells(1, 2) = ""
Cells(2, 2) = ""
Cells(3, 2) = ""
Filename = Application.GetOpenFilename("Arquivos txt(*.txt),", , "Selecione um arquivo")
If Filename = False Then
MsgBox "Nenhum arquivo foi selecionado!", vbCritical, "Falha na sele��o de arquivo"
Filename = ""
ActiveSheet.Protect
Exit Sub
End If
Cells(1, 2) = Replace(Filename, "http:", "")

Dr = Left(Cells(1, 2), InStrRev(Cells(1, 2), "\"))
With Application.FileDialog(msoFileDialogFolderPicker)
.AllowMultiSelect = False
.Show
If .SelectedItems.Count > 0 Then
Cells(2, 2) = .SelectedItems(1) & "\FET-2018-PONTUAL.csv" ' Nome padr�o para envio ao DICI
Else
x = MsgBox("Nenhum destino foi selecionado!" & Chr(13) & "O destino ser� o mesmo da origem.", vbOKCancel, "Falha na sele��o do destino")
If x <> vbOK Then Cells(1, 2) = "": Exit Sub
Cells(2, 2) = Dr & "FET-2018-PONTUAL.csv"
End If
End With
If Dir(Cells(2, 2)) = "FET-2018-PONTUAL.csv" Then
    x = MsgBox(Cells(2, 2) & " j� existe. Deseja reprocess�-lo?", vbOKCancel, "Arquivo j� existe!")
    If x <> vbOK Then
        Cells(1, 2) = ""
        Cells(2, 2) = ""
    Exit Sub
    End If
End If
ActiveSheet.Protect
Converter_Click
End Sub


Private Sub Converter_Click()
inicio = Now()
On Error GoTo Erro
Open Cells(1, 2) For Input As 1
Open Cells(2, 2) For Output As 2
i = 0
While Not EOF(1)
Line Input #1, Linha
Linha2 = Replace(Linha, vbTab, ";")
Print #2, Linha2
i = i + 1
Wend
ActiveSheet.Unprotect
Cells(3, 2) = "Foram processadas " & i & " linhas em " & DateDiff("s", inicio, Now()) & " segundos."
ActiveSheet.Protect
Close 1: Close 2
Exit Sub
Erro:
Close 1: Close 2
MsgBox "Nenhum processamento foi realizado", , "Opera��o cancelada"
End Sub
Function TestaExistenciaArquivo(ByVal caminhoArquivo As String)
    Dim retorno As Boolean
    Set FSO = New FileSystemObject
    retorno = FSO.FileExists(caminhoArquivo)
    TestaExistenciaArquivo = retorno
End Function

