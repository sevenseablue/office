## word的表格拷贝到excel中

####说明
- 一个表格拷贝到一个sheet中

- 如果列数不一样， 以最多的列数为准

- 如果行数不一样， 以最多的行数为准
####环境
```shell script
python3
pip install python-docx
pip install docx
pip install xlsxwriter
```

####执行语句
```python
python word_tb_to_excel.py /home/wangdawei/Documents/word1.docx
```
结果在同名的excel文件中