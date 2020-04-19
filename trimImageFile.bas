Sub 画像の取り込み()
    
    Dim strPath As String
        strPath = ""
    
        With Application.FileDialog(msoFileDialogFolderPicker)
            If .Show = True Then
                strPath = .SelectedItems(1)
            End If
        End With
    
    Dim objFSO As Object
    Dim objFiles As Object
    Dim objFile As Object
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set objFiles = objFSO.GetFolder(strPath).Files
    
    Dim obj画像 As Object '画像をもとに生成されたシェイプオブジェクト
    
        For Each objFile In objFiles
        With objFile
            
            DoEvents
            
            If InStr(.Name, ".png") > 0 _
            Or InStr(.Name, ".jpg") > 0 _
            Or InStr(.Name, ".jpeg") > 0 _
            Or InStr(.Name, ".gif") > 0 _
            Or InStr(.Name, ".bmp") > 0 Then
            
                Set obj画像 = ActiveSheet.Pictures.Insert(.Path)
                
                Call トリミング(obj画像)
                With obj画像
                    'コピーして画像を作る。
                    .Copy
                    ActiveSheet.PasteSpecial _
                        Format:="図 (PNG)", _
                        Link:=False, _
                        DisplayAsIcon:=False
                        
                    'シェイプを削除
                    .Delete
                End With
                Set obj画像 = Nothing
                
            End If
            
        End With
        Next objFile
        Set objFile = Nothing
        Set objFiles = Nothing
        
    Set objFSO = Nothing
    
    MsgBox "完了"
    
End Sub

Sub トリミング(obj画像 As Object)
'トリミングするサイズをメンテしてください。

    With obj画像.ShapeRange
        '上部
        .LockAspectRatio = msoFalse
        .ScaleHeight 0.8840909091, msoFalse, msoScaleFromTopLeft
        .PictureFormat.Crop.PictureWidth = 445
        .PictureFormat.Crop.PictureHeight = 792
        .PictureFormat.Crop.PictureOffsetX = 0
        .PictureFormat.Crop.PictureOffsetY = -45
        
        '下部
        .LockAspectRatio = msoFalse
        .ScaleHeight 0.9023136247, msoFalse, msoScaleFromTopLeft
        .PictureFormat.Crop.PictureWidth = 445
        .PictureFormat.Crop.PictureHeight = 792
        .PictureFormat.Crop.PictureOffsetX = 0
        .PictureFormat.Crop.PictureOffsetY = -11
        
    End With
End Sub
