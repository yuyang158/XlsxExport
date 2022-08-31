# 配置导出工具使用方法

## 导出工具的使用

### 执行导出

1. 运行ExportRelease文件夹中的ExcelExport.exe可以执行导出操作。
2. 导出文件保存在Export文件夹中
3. 导出文件确认无误后需要拷贝到Unity工程中的Assets/Res/Xlsx目录

### 如何添加一个新的导出Excel文件

1. 使用文本编辑工具打开ExportRelease\Config\Export.json
2. 在里面按照Json格式添加新的Excel文件的文件名

### 一个Excel文件是否支持多sheet？

*支持*

## 导出文件的配置

### 基础格式

每个Excel表单均包含三行文件头

	| id | value |
	| --- | ---|
	| int | string |
	| 说明1 | 说明2 |
	| cs | cs |

	* 第一行为英文键名
	* 第二行为类型
	* 第三行为内容说明
	* 第四行导出列为服务器(s)或客户端(c)使用

### 支持类型

#### 基础类型

1. int 整数类型
2. number 浮点数类型
3. string 字符类型
4. translate 多语言文本类型
5. link 外链类型，内容填外链表的id，第一行天对应表单名

#### 复杂类型

1. 数组类型：将第一行的名称配置为一样即可
2. Key-Value类型：将第一行的名称配置为A.B。例如：item.id, item.count

#### 额外处理

1. 表单名及列名支持ignore工程，即使用ignore开头的列及表单不予导出。用于处理公式等。
2. 表单名使用大写开头驼峰命名：CeShi
3. 列名使用小写开头驼峰命名：ceShi
# 配置导出工具使用方法

## 导出工具的使用

### 执行导出

1. 运行ExportRelease文件夹中的ExcelExport.exe可以执行导出操作。
2. 导出文件保存在Export文件夹中
3. 导出文件确认无误后需要拷贝到Unity工程中的Assets/Res/Xlsx目录

### 如何添加一个新的导出Excel文件

1. 使用文本编辑工具打开ExportRelease\Config\Export.json
2. 在里面按照Json格式添加新的Excel文件的文件名

### 一个Excel文件是否支持多sheet？

*支持*

## 导出文件的配置

### 基础格式

每个Excel表单均包含三行文件头

	| id | value |
	| --- | ---|
	| int | string |
	| 说明1 | 说明2 |
	| cs | cs |

	* 第一行为英文键名
	* 第二行为类型
	* 第三行为内容说明
	* 第四行导出列为服务器(s)或客户端(c)使用

### 支持类型

#### 基础类型

1. int 整数类型
2. number 浮点数类型
3. string 字符类型
4. translate 多语言文本类型
5. link 外链类型，内容填外链表的id，第一行天对应表单名

#### 复杂类型

1. 数组类型：将第一行的名称配置为一样即可
2. Key-Value类型：将第一行的名称配置为A.B。例如：item.id, item.count

#### 额外处理

1. 表单名及列名支持ignore工程，即使用ignore开头的列及表单不予导出。用于处理公式等。
2. 表单名使用大写开头驼峰命名：TestCase
3. 列名使用小写开头驼峰命名：testCase
