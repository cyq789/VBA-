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
    Workbooks.Open Filename:="C:\Users\TB745823749857\Desktop\学习资料\excel\Tool\录入表单.xlsx"
    
	Windows("lesson1.xlsx").Activate
End Sub
```
除Filename参数外，还有14个可决定如何Open的参数

### 将工作簿切换为活动工作簿
``` vb
Workbooks("工作簿1").Activate
```

## 操作工作表
### 引用sheets的方法
``` vb
Worksheets(2) '索引号
Worksheets("Sheets1") 'Name
Name '属性窗口中的代码名称（Name），在性质上不同于标签名称，只能在属性窗口修改
```
### 获取sheets数量 在复制表、移动表时很有用
``` vb
Worksheets.Count
```

### sheets和worksheets的区别
Sheets 包括工作表、图标、MS宏表、MS对话框
Worksheets是sheets中的一种

## 操作单元格

### 引用单元格
``` vb
Sub 引用单个固定区域()
	Activesheet.Range("A1:A10").Value = 200
	Dim n as String
	n = "B1:B10"
	Activesheet.Range(n) = 100
End Sub

Sub 引用多个不连续的单元格区域()
	Range("A1:A4,B6:E10,C2:F4").Select
```

Cells()属性只能引用单个单元格，但由于行列可指定，使用起来最为灵活
```
Activesheet.Cells(3, 4).Value = 20
Activesheet.Cells(3, "D").Value = 20
Range("B3:F9").Cells(2,3).Value = 20 '在B3到F9区域中的（2,3）位置，属于相对位置
```
### 引用整行、整列
``` vb
Activesheet.Rows.Select '选择所有行
Activesheet.Rows("3:5") '活动单元表的第三行到第五行
Activesheet.Rows(3) '活动单元表的第三行到第五行
Activesheet.Rows("3:3") '活动单元表的第三行到第五行

Rows("3:10").Rows("1:1")  '指定区域的第一行
```

``` vb
Activesheet.Columns.Select
Activesheet.Columns("F:G").Select
Activesheet.Columns(6).Select
Columns("F:G").Columns("B:B").Select 'B:G列中的第二列
```

### Selection
简便写法

### Offset属性引用相对位置
``` vb
Sub 引用相对位置单元格()
	ActiveCell.offset(4,0).Value = 500 '活动单元格下方第四行的单元格
End Sub
```
Offset属性可用于单元格，也可用于区域

### 修改单元格中保存的数据
``` vb
Sub 在A1:D5单元格中输入数据()
	Range("A1:D5").Value = "Excel VBA其实很简单"
End Sub
```
区分属性Text\Value\Formula
- Formula用于读取公式
- Value用于读取值
- Text读取表面值，用日期的例子会比较好理解 P85

### 包含单元格个数
``` vb
Sub 单元格个数()
	MsgBox Range("A1:D5").Count 
End Sub

Activesheet.UsedRange.Rows.Count
Activesheet.UsedRange.Columns.Count
```
### 获取单元格位置
```
Range("H1").Value = ActiveCell.Row
Range("H1").Value = ActiveCell.Address
```
### 复制、剪切单元格区域 
```
' 简便写法
Range("A1").Copy Range("C1")
Range("A1").Cut Range("C1")

## 操作Excel应用程序
- ScreenUpdating
- DisplayAlerts
- WorksheetFunction属性使用工作表函数（并不是所有工作表函数都能调用）
``` vb
Sub 统计个数()
	Dim Mycount As Integer
	Mycount = Application.WorksheetFunction.Countif(Range("A1:B50"),">1000")
	MsgBox Mycount
End Sub
```
```
Sub 全屏显示()
	Application.DisplayFullScreen = True
End Sub 
```




