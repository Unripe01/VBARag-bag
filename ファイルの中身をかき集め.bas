Option Explicit

Sub �����W��()
    Dim app As Application
    Dim masterBook As Workbook
    Dim masterSheet As Worksheet
    Dim w As Worksheet
    
    Set app = Application
    Set masterBook = ActiveWorkbook
    Set masterSheet = ActiveSheet
    
    app.ScreenUpdating = False
    masterBook.Worksheets.Add Before:=masterSheet
    Set w = masterBook.Sheets(1)
    w.Range("A1").Value = "�w�b�_"
    w.Range("A2").Value = "�w�b�_�Q"

        

    '* �t�@�C����I��
    Dim files As Variant
    Set files = Get�t�@�C��
    If (files Is Nothing) Then
        Exit Sub
    End If
    
    Dim file As Variant
    Dim targetWorkbook As Workbook
    Dim w1 As Worksheet
    Dim r1 As Range
    Dim r2 As Range
    
    For Each file In files
        
        Set targetWorkbook = app.Workbooks.Open(file, , False)
        Set w1 = targetWorkbook.Sheets(1)
        Set r1 = w1.Range(w1.Range("A3"), w1.Cells(w1.Rows.Count, 13).End(xlUp))
        
        '''masterSheet.Activate
        Set r2 = masterSheet.Range("A" & masterSheet.Rows.Count).End(xlUp).Offset(1, 0)
        
        r1.Copy r2
        
        targetWorkbook.Close False
    Next file
    
    app.ScreenUpdating = True
End Sub


Function Get�t�@�C��() As FileDialogSelectedItems

    
    Dim var_fname As Variant
    Dim var_bookname As Variant
    Dim strPrompt As String
    Dim str_rtn As String
    
    '�t�@�C���I���̃_�C�A���O�N��
    With Application.FileDialog(msoFileDialogFilePicker)
    
    
        With .Filters
            '�u�t�@�C���̎�ށv���N���A,excel�����\��
            .Clear
            .Add "Excel�u�b�N", "*.xls; *.xlsx; *.xlsm", 1
        End With
    
        '�\������t�H���_���w��
        .InitialFileName = "C:\"
    
        '�\������A�C�R���̑傫�����w��
        .InitialView = msoFileDialogViewLargeIcons
        .AllowMultiSelect = True
        .ButtonName = "�I��"
        .Title = "�t�@�C���I��(�����t�@�C����)"
        
        
        If .Show = -1 Then '�L���ȃ{�^�����N���b�N���ꂽ
            
            Set Get�t�@�C�� = .SelectedItems

        Else   '[�L�����Z��]�{�^�����N���b�N���ꂽ
            'MsgBox "�L�����Z������܂���"
        End If
       
    End With

End Function

