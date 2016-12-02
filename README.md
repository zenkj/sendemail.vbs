# sendemail.vbs
创建一个数据文件 data.xlsx，内容如下：

|姓名 | 邮箱 | 主题 | 内容 | 附件 |
|----|----|----|----|----|
|张三|abc@gmail.com|腐败邀请|来老地方轰趴|d:\abc.dat;c:\def.txt|

执行sendemail.vbs脚本，然后选择这个data.xlsx文件，确保附件中的文件存在，之后将会使用outlook发送邮件。
需要确保outlook提前配置好。
