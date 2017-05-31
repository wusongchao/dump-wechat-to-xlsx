# dump-wechat-to-xlsx

![Alt text](https://github.com/wusongchao/dump-wechat-to-xlsx/raw/master/xiaoguo.png)

need: `xlsxwriter`

It's rubbish, since the user have to use sqlcipher to decrypt the `EnMicroMsg.db` first(you can search many methods to do so in the Internet), then export these three tables: `rcontact`, `message`, `ImgInfo2` to csv.

The Default path of csv files and out dir are:
```python
EXPORTED_DATA_PATH = "ExportedData/"
XLSX_OUT_PATH = "out/"
```

then run the main.py
