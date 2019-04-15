Option Explicit

Sub かき集め()
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
    w.Range("A1").Value = "ヘッダ"
    w.Range("A2").Value = "ヘッダ２"

        

    '* ファイルを選ぶ
    Dim files As Variant
    Set files = Getファイル
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

