# 实现自动合并word的python脚本文件

- 保留所有文档固有的样式
- 保留文档的顺序

## 使用方法

**【注意】** 只能合并`.docx`文件，若需要合并的文档中存在`.doc`文件，需先手动将其转换为`.docx`文件才能使用此脚本

- 在该项目的根目录下创建`files`文件夹，将需要合并的`.docx`文件放进去
- 运行`merge.py`，即可得到合并结果文件`merge_result.docx`

## 特性说明

### 文档分页

每个文档自动从新一页开始，两个文档之间会插入一个分页符

### 文档合并顺序

文档的合并顺序与在资源管理器中的排序一致，因此对排序有强要求的需求，最好对文档进行编号

