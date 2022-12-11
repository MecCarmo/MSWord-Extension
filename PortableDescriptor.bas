Attribute VB_Name = "PortableDescriptor"
Sub Main()
    UserForm1.Show
End Sub

Sub AddImg() 'Adiciona uma imagem
    
    Dim intChoice As Integer
    Dim strPath As String
    

    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False

    intChoice = Application.FileDialog(msoFileDialogOpen).Show

    If intChoice <> 0 Then

        strPath = Application.FileDialog( _
            msoFileDialogOpen).SelectedItems(1)
    End If

    Selection.InlineShapes.AddPicture FileName:= _
        strPath, LinkToFile:=False, _
        SaveWithDocument:=True
        
End Sub

Sub getalt() 'Caixa de input para cada Alt text das imagens
    Dim oshp As InlineShape
    Dim Message, Title, Default
    Message = "Enter the image description"
    Title = "InputBox Demo"
    Default = "Image description"
    
    For Each oshp In ActiveDocument.InlineShapes
        oshp.Select
    If oshp.AlternativeText <> "" Then
        
    Else
        oshp.AlternativeText = InputBox(Message, Title, Default)
        
    End If
    Next
End Sub


Sub ConvertToTxt() 'Converter para .txt
    With Application.Dialogs(wdDialogFileSaveAs)
      .Name = Split(ActiveDocument.FullName, ".doc")(0) & ".txt"
      .Format = wdFormatTextLineBreaks
      .AddToMru = False
      .Show
    End With
End Sub


Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro2"
'
' Macro2 Macro
    Main
'
'
End Sub
