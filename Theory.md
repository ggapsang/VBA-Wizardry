## Application Object

```
Sub Application_Object
    Dim strInfo As String
    strInfo = Application.UserName
    strInfo = Application.Version
    strInfo = Application.Path
    strInfo = Application.StartPath
End Sub
```

- .UserName = 사용자 이름
- .Verstion = 현재 설치된 엑셀 버전
- .Path = 엑셀이 설치되어 있는 폴더 경로
- .StartupPath = 시작 폴더의 위치
