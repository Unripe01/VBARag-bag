Attribute VB_Name = "Module2"
Option Explicit

Sub �����W��()
Attribute �����W��.VB_ProcData.VB_Invoke_Func = " \n14"
    
    Application.ScreenUpdating = False
    Sheets.Add Before:=ActiveSheet
    Sheets(1).Select
    Range("A1").Value = "�w�b�_"
    Range("A2").Value = "�w�b�_�Q"

    Dim app As Application
    Dim masterBook As Workbook
    Dim masterSheet As Worksheet
    
    Set app = Application
    Set masterBook = ActiveWorkbook
    Set masterSheet = ActiveSheet
    

    '* �t�@�C����I��
    Dim files As Variant
    Set files = Get�t�@�C��
    If (files Is Nothing) Then
        Exit Sub
    End If
    
    Dim file As Variant
    Dim targetWorkbook As Workbook
    
    For Each file In files
        
        Set targetWorkbook = app.Workbooks.Open(file)
        targetWorkbook.Sheets(1).Range(Range("A3"), Cells(Rows.Count, 13).End(xlUp)).Copy
        
        masterSheet.Activate
        masterSheet.Range("A" & Rows.Count).End(xlUp).Offset(1, 0).PasteSpecial
        
        targetWorkbook.Close
    Next file
    
    Application.ScreenUpdating = True
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
