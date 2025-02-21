''' 현재 시트의 모든 이미지(Shape) 목록을 배열로 반환하는 함수
Function GetImageList(ws As Worksheet) As Variant
    Dim shp As Shape
    Dim imageList As Object
    
    Set imageList = CreateObject("System.Collections.ArrayList")
    
    ' 모든 Shape 개체를 확인하여 이미지(msoPicture)만 리스트에 추가
    For Each shp In ws.Shapes
        If shp.Type = msoPicture Then
            imageList.Add shp.Name
        End If
    Next shp
    
    ' 배열로 변환하여 반환
    If imageList.Count > 0 Then
        GetImageList = imageList.ToArray()
    Else
        GetImageList = Array() ' 이미지가 없을 경우 빈 배열 반환
    End If
End Function
