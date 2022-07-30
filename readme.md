# doc-markdown-converter

环境要求: python3

安装依赖

```
pip3 install -r ./requirements.txt 
```

## 使用方法

```
python main.py xxx.docx
```

## 原理解释

使用`mammoth`库先将docx文件转换成html, 再使用`markdownify`将html转换成markdown格式.
