# JNU-Word-auto-generated
暨南大学勤工助学考核表一键生成（包含签到表）
##  说明
毕设写完了真的没事干，这是以前挖的第二个坑，现在把它填上  
用于自动生成排班表，不用人工做表。  
其中txt文件存的是个人信息，一共4行，分别是：  
```
小明.txt 
  ├── 姓名
  ├── 手机号码
  ├── 银行卡号
  └── 工时 
```
module和module2文件是表格模板，将它们和主程序放在同一目录下即可  
以下为项目结构（其中main.py, 所有txt文件, module.docx, module2.docx是必须文件，剩下的docx是运行程序后生成的文件）  
```
JNU-Word-auto-generated 
  ├── main.py
  ├── 小明.txt
  ├── 小李.txt
  ├── module.docx
  ├── module2.docx
  ├── 小明5月工作考核表.docx（运行程序后生成）
  ├── 小李5月工作考核表.docx（运行程序后生成）
  └── 总表.docx（运行程序后生成）
```
- 适用于日期安排在每月1-25号的工时表生成，且不考虑是否为工作日  
##  立项理由
工时表人工排序真不是人干的，眼睛不够用，还容易抄错，不如一键生成，大家都别做表了  
这大概是大学生涯最后一次按自己的意愿写代码了，且行且珍惜  
##  特别鸣谢
救不出塞尔达的林克  

![957e2aa182b8318ad06d7a8105b1aa51](https://user-images.githubusercontent.com/85060372/168338404-09027ec1-9435-48ce-a59e-539c76d2b0fb.jpg)
