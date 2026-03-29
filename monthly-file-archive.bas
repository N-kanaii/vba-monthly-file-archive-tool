Attribute VB_Name = "Module1"
Option Explicit

Sub 指定のHPへ年度別にアップロード()

Application.ScreenUpdating = False
Application.EnableEvents = False

    Dim fso As FileSystemObject ' FileSystemObjectインスタンス
    Set fso = New FileSystemObject
    
    Dim folder1 As String
    Dim folder2 As String
    Dim folder3 As String
    Dim folder4 As String
    Dim folder5 As String
    Dim folder6 As String
    Dim folder7 As String
  
    
    Dim todayFolderName As String
    Dim nowFolderName As String
    Dim nendo As String
    
    Dim itemname As String
    Dim fullpath As String
    
    Dim i As Long
    Dim lastRowNum As Long
    
    
    '月、年度取得
    todayFolderName = Format(DateAdd("m", -1, Date), "yyyy.m") '翌月に前月分を処理するため-1
    nowFolderName = Format(Date, "yyyy.m月") '現在の月
    nendo = Format(DateAdd("m", -4, Date), "yyyy年度")  '年度フォルダ


 With Sheets("Master")
        '最終行数を取得します。
        lastRowNum = .Cells(.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRowNum '２行目から開始のため

        itemname = .Cells(i, 1).Value
        fullpath = .Cells(i, 2).Value



    '●フォルダ作成
    folder1 = fullpath
   
    If fso.FolderExists(folder1) = True Then
    
       Else
       fso.CreateFolder folder1
       
    End If
    
    '\階層フォルダ作成
    
    folder2 = folder1 & "\" & itemname
   
    If fso.FolderExists(folder2) = True Then
    
       Else
       fso.CreateFolder folder2
       
    End If
    
    
    '過去分フォルダ作成
     folder3 = folder2 & "\過去分"
   
    If fso.FolderExists(folder3) = True Then
    
       Else
       fso.CreateFolder folder3
       
    End If
    
    '過去分→年度
    
      folder4 = folder3 & "\" & nendo
   
    If fso.FolderExists(folder4) = True Then
    
       Else
       fso.CreateFolder folder4
       
    End If
     
     '過去分→年度→月フォルダ
      folder5 = folder4 & "\" & todayFolderName
   
    If fso.FolderExists(folder5) = True Then
    
       Else
       fso.CreateFolder folder5
       
    End If
    
    
    
    '●過去分へ移動
    
Dim 移動元フォルダ As Folder
Set 移動元フォルダ = fso.GetFolder(folder2)    'フォルダ作成
Dim 移動先フォルダ As Folder
Set 移動先フォルダ = fso.GetFolder(folder5) '過去分へ移動させる

' 移動元フォルダのすべてのファイルをループ
Dim kakofile As file
        For Each kakofile In 移動元フォルダ.Files

           ' 条件に合致するファイルか判定
            If kakofile.Name Like "*.xlsm" Then

            ' 移動を実行
            kakofile.Move 移動先フォルダ.Path & "\"  'あれば止まる

            End If

        Next
    

    '●当月分処理

  folder6 = "原本ファイルの格納先" & itemname
  folder7 = "リネームのためのフォルダ" & "\一時保管BOX"
             
    
Dim 原本元フォルダ As Folder
Set 原本元フォルダ = fso.GetFolder(folder6) '当月フォルダ
Dim 原本リネーム先フォルダ As Folder
Set 原本リネーム先フォルダ = fso.GetFolder(folder7) '一時保管BOX

' 移動元フォルダのすべてのファイルをループ
        Dim imafile As file
        For Each imafile In 原本元フォルダ.Files



        ' 条件に合致するファイルか判定
        If imafile.Name Like "*.xlsm" Then
        
              
        ' コピーを実行
           imafile.Copy 原本リネーム先フォルダ.Path & "\"
     
       
            End If

            Next
      
      '●原本リネーム先フォルダの中をリネーム

' すべてのファイルをループ
        Dim imafile2 As file
        For Each imafile2 In 原本リネーム先フォルダ.Files

    ' 条件に合致するファイルか判定
    If imafile2.Name Like "*.xlsm" Then
        
    ' リネーム
    Name imafile2 As folder7 & "\" & nowFolderName & "　" & imafile2.Name
      
     
       
    End If

    Next
    
  
 '●リネームした当月分を移動
    
Dim 当月移動元フォルダ As Folder
Set 当月移動元フォルダ = fso.GetFolder(folder7)
Dim 当月移動先フォルダ As Folder
Set 当月移動先フォルダ = fso.GetFolder(folder2)

' 移動元フォルダのすべてのファイルをループ
Dim nowfile As file
             For Each nowfile In 当月移動元フォルダ.Files
    


    ' 条件に合致するファイルか判定
    If nowfile.Name Like "*.xlsm" Then
        

        
        ' 移動を実行
    nowfile.Move 当月移動先フォルダ.Path & "\"
    
    End If
    
             Next
    


Next i

            End With
    
   MsgBox "対応完了"
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

