
import pandas as pd

df = pd.read_excel('./客户信息总表(清洗).xlsx')
df.to_sql(name='app01_customer', con='mysql+pymysql://root:123456@localhost:3306/customertable?charset=utf8',
          if_exists='replace',index=False)