Attribute VB_Name = "Module2"
Option Explicit

Sub かき集め()
Attribute かき集め.VB_ProcData.VB_Invoke_Func = " \n14"
    
    Application.ScreenUpdating = False
    Sheets.Add Before:=ActiveSheet
    Sheets(1).Select
    Range("A1").Value = "ヘッダ"
    Range("A2").Value = "ヘッダ２"

    Dim app As Application
    Dim masterBook As Workbook
    Dim masterSheet As Worksheet
    
    Set app = Application
    Set masterBook = ActiveWorkbook
    Set masterSheet = ActiveSheet
    

    '* ファイルを選ぶ
    Dim files As Variant
    Set files = Getファイル
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


Function Getファイル() As FileDialogSelectedItems

    
    Dim var_fname As Variant
    Dim var_bookname As Variant
    Dim strPrompt As String
    Dim str_rtn As String
    
    'ファイル選択のダイアログ起動
    With Application.FileDialog(msoFileDialogFilePicker)
    
    
        With .Filters
            '「ファイルの種類」をクリア,excelだけ表示
            .Clear
            .Add "Excelブック", "*.xls; *.xlsx; *.xlsm", 1
        End With
    
        '表示するフォルダを指定
        .InitialFileName = "C:\"
    
        '表示するアイコンの大きさを指定
        .InitialView = msoFileDialogViewLargeIcons
        .AllowMultiSelect = True
        .ButtonName = "選択"
        .Title = "ファイル選択(複数ファイル可)"
        
        
        If .Show = -1 Then '有効なボタンがクリックされた
            
            Set Getファイル = .SelectedItems

        Else   '[キャンセル]ボタンがクリックされた
            'MsgBox "キャンセルされました"
        End If
       
    End With

End Function
