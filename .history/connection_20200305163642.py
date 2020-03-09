from impala.dbapi import connect
import pandas as df
from impala.util import as_pandas
from impala.dbapi import connect
# HIVE的方式和impala一样，只是修改host ,port
# 参考这个https://blog.csdn.net/Xiblade/article/details/82318294
# C:\Users\admin\AppData\Local\Programs\Python\Python35\Lib\site-packages\thrift_sasl
# 执行数据库连接后，再次出现问题

# TypeError: can’t concat str to bytes
# 定位到错误的最后一条，在init.py第94行

# ...
# header = struct.pack(">BI", status, len(body))
# self._trans.write(header + body)

# 修改为
# ...
# header = struct.pack(">BI", status, len(body))
# if(type(body) is str):
#     body = body.encode()
# self._trans.write(header + body)

conn = connect(host='10.80.22.229', port=21050, database='dw', timeout=5)
cursor = conn.cursor()
sql = '''
select * from  app.vendor_verify
limit 10
'''
cursor.execute(sql)
data = as_pandas(cursor)
print("msg")
print(data)
# data.to_excel('province_data.xlsx',index=None)
# data.to_csv('province_data.csv',index=None,encoding='gbk')
