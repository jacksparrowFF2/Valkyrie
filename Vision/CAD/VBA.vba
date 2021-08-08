' 声明坐标变量
Dim cc(0 To 2) As double
cc(0) = 1000
cc(1) = 1000
CC(2) = 0
' 开始循环
For i = 1 to 1000 step 10 
    call thisDrawing.ModelSpace.AddCircle(cc,i*10)
Next i
' 显示整个图形
ZoomExtents
