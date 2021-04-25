import pandas as pd

df = pd.read_excel('./2021年1月陈皮酒订单.xlsx')
df.to_sql(name='chenpijiudingdan', con='mysql+pymysql://root:123456@localhost:3306/ChenPiJiu?charset=utf8',
          if_exists='replace',
          index=False)
