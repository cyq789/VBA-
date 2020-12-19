## 引用对象
### 引用格式
``` vb
Application.Workbooks("工作簿1").Worksheets("Sheets1").Range("A1")
```
### 对象名称
``` vb
Workbooks(1) '当前打开的第一个工作簿
Workbooks("工作簿233"） '打开名为工作簿233的工作簿2
'活动单元格，与下述等效
ActiveCell 
Application.ActiveCell 
ActiveWindow.ActiveCell 
Application.ActiveWindow.ActiveCell
Worksheets(1) '当前工作簿中的第一张工作表
```

## 对象的属性和方法

```vb
' 属性：set sheet1's name as abc 
Work("Sheets1").Name = "abc"

' 方法：Select A1
Range(A1).Select
```
## 操作工作簿
### 如何打开一个工作簿文件
用录制宏功能，从excel→文件打开，可得到如下代码，不需要自己调试
``` vb
Sub 打开文件()
    Range("I7").Select
    Windows("录入表单.xlsx").Activate
    ActiveWindow.WindowState = xlNormal
	
	'以下这句为关键语句
    Workbooks.Open Filename:="C:\Users\TB468XF\Desktop\学习资料\excel\Tool\录入表单.xlsx"
    
	Windows("lesson1.xlsx").Activate
End Sub
```
除Filename参数外，还有14个可决定如何Open的参数

### 将工作簿切换为活动工作簿
``` vb
Work("Sheets1").Activate
```

##操作工作表

