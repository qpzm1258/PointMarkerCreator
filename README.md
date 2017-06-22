# 程序简介
本程序按照template.dotx模板的书签插入数据和图片

本程序基于.NET Framework4.5和Microsoft Office 2010/2013/2016 API
# 使用说明（中文）
## 运行程序后会有相应的使用提示,中文如下

本程序基于.NET Framework4.5和Microsoft Office 2010/2013/2016 API

环境设置:
[txt 设置]
请将txt文件放置在: 程序所在目录\point\point.txt
txt的格式为(每行一个点，[]表示必填参数，实际文件内无[]，下同):[点名],[X],[Y],[H]

[图片设置]
请将概略点位图放在: 程序所在目录\point\imagedir\big\目录下
请将点位略图放在: 程序所在目录\point\imagedir\middle\目录下
请将点位详细图放在: 程序所在目录\point\imagedir\small\目录下
每个图片目录的图片必须重名为: [点名].jpg

[文档设置]
文档将会生成在: 程序所在目录\point\point.docx
警告！！！该文件每次运行将会被覆盖，且不会有任何提示，请做好备份和保存。

[比例设置]
请输入比例尺(1:1000即使输入1000,1:500输入500):


请确认你的文件放置在正确位置了，并按任意键开始

## 程序运行如下:

[点名],[X],[Y],[H]  
&ensp;&emsp;&emsp;&emsp; .  
&ensp;&emsp;&emsp;&emsp; .  
&ensp;&emsp;&emsp;&emsp; .  
[点名],[X],[Y],[H]  
&ensp;&emsp;&emsp;&emsp; .  
&ensp;&emsp;&emsp;&emsp; .  
&ensp;&emsp;&emsp;&emsp; .  

### 如果生成成功
将显示:Document created successfully !

### 如果生成失败
则显示相关信息

# 一些问题的解决方案:
### 异常1、System.AccessVoilationException:尝试读取和写入保护内存……  
**原因:** office版本与dll版本不匹配，请确认office版本是2010/2013/2016  
**解决方案:** 卸载现有版本的office，安装指定版本。  

### 异常2、服务器出现意外情况……  
**原因:** 由于office自动更新导致office在注册表中存在多个键值。  
**解决方案:** 控制面板->程序与功能->卸载程序，找到office，右键点击选择更改，然后按提示修复office  

如果在使用中发现其他新的异常或者有新的需求，请联系本人，在提出需求或者修复前，网管必须请客吃饭  

# 跟新日志
### 更新日志:v1.3.0.0  
使用分页符换页，禁用输入检查  

### 更新日志:v1.4.0.0  
修复图片长宽定义无效的错误。  

### 更新日志:v1.5.0.0  
添加图幅编号生成，必须输入比例尺。  

### 更新日志:v1.6.0.0  
修改调用顺序，先进行设置选项显示再创建winword进程。  

