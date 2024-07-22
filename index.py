import time
import requests
import pandas as pd

# 读取 Excel 文件
# df = pd.read_excel('file.xlsx')
# # 筛选包含"发图"的列
# keywords = ['发图/来图', '计划下达', '合同签订']
# filtered_columns = df.columns[df.apply(lambda col: col.astype(str).str.contains('|'.join(keywords)).any())]
# filtered_df = df[filtered_columns]
# 显示数据
# print(df)

# def download_file(url, local_filename):
#     # 发送 HTTP GET 请求获取文件内容
#     with requests.get(url, stream=True) as r:
#         r.raise_for_status()  # 检查请求是否成功
#         # 打开一个本地文件进行写入
#         with open(local_filename, 'wb') as f:
#             for chunk in r.iter_content(chunk_size=8192): 
#                 f.write(chunk)  # 将文件内容写入本地文件
#     return local_filename

# # 示例用法
# url = 'http://oa.cisdi.com.cn:8080/weaver/weaver.file.FileDownload?f_weaver_belongto_userid=18694&f_weaver_belongto_usertype=0&fileid=4321544&download=1&requestid=3415521&desrequestid=0&fromrequest=1'
# local_filename = 'file.xlsx'
# download_file(url, local_filename)
# print(f"Downloaded {local_filename}")

# def task():
#     print("Task is running")
#     url = "http://ccis-oamgate.cisdi.com.cn:7777/api/v1/todolist"

#     payload = "{\"pageNum\":1,\"pageSize\":13,\"input\":\"\",\"type\":\"\",\"typeName\":\"\"}"
#     headers = {
#     'Accept': 'application/json, text/plain, */*',
#     'Content-Type': 'application/Json; charset=UTF-8',
#     'token': 'eyJhbGciOiJIUzI1NiJ9.eyJqdGkiOiJ7XCJ1c2VySWRcIjpcIjAxNDYzMkBjaXNkaS5jb20uY25cIn0iLCJzdWIiOiLnmbvlvZV0b2tlbiIsImlzcyI6IumXqOaItyIsImlhdCI6MTcyMTM1NDgxMywiZXhwIjoxNzIxMzgzOTEzfQ.ebY4gZY72s-gtPZabJL-yGIfJoQGUxexw-2wQ5iWxH8',
#     'Cookie': 'OAMAuthnCookie_ccis-oamgate.cisdi.com.cn:7777=1c82d8a5105dd3c467f0b2819fccf967012911b0%7EVcFBPnASueboBby1RBx5wrxChUkyEEEpyi1RQPefur6GonW1lwF7EEBjk%2FA8eApb8K1sfPwVow%2FIyHxApxPLkGRhONLoncJ6EXQHdpomeMrTkT%2Bjz4e%2BRJ9veOhetJ4q5U35wwI6vYsZP3J5XUAqKT5pxRpxwQ6YOeaC%2FrxcpnS28qSk8b0fA5aE%2BgyWTDXAF1Hz1pGN5sIP1hpAd%2F4WCyhNJn%2Bx%2BzxH%2BHBeUPQ3Aeo%2FZ%2F8CXd8gx12g34A8lNuL8N4JcQvGAkH%2BV%2BU4Q5Z6T%2Fxry6O4m%2BdvGKkvuk9K3OLVaEjyXh7GuQwufoVov%2BN%2BRCTa62QvmG3J8GkzQcVSbdJ3GdG6rgMm3PGKjNvGWYU0v5q4Yf1gscDAFV4OlsqgzMuVxRzFXsi0rAqZ3SPwh0FR3xdcPhGXnhynEsy82%2FeWNrr8OzqABfCVaRtR%2B6z57i%2FlMQj5Ng0eXzz7LSz4qRV0Uim8mZuYqm8pqYCxIRZDU%2F15Q1HSOtG1QHtn4ybtb7u16MC3WMxvn6q5EZ5%2B75WhQM4IQFlFLGFT0ivHNGF2MNEkulzDs7grKLXk673WzlidebA3aECt2MSdZCkoP3YuVNRTL6cyywZX1pP2v0E6cUrJ%2FKjnvdcntdSBWsDZ9g54ERmRfAMiXX96gtXcDfTGhzzg5HGv%2BaFx9Sj5QkHSWSYsDeHShy8TkmcTeGYBeAbeysk3uGZNkJWjGpyTthCMsyETXRk7tdDacDeADDUNrFAK5EOMYK7BjUXu9du%2FURdlSwOecb0AzYyKAGgccA%3D%3D; token=eyJhbGciOiJIUzI1NiJ9.eyJqdGkiOiJ7XCJ1c2VySWRcIjpcIjAxNDYzMkBjaXNkaS5jb20uY25cIn0iLCJzdWIiOiLnmbvlvZV0b2tlbiIsImlzcyI6IumXqOaItyIsImlhdCI6MTcyMTM1NDgxMywiZXhwIjoxNzIxMzgzOTEzfQ.ebY4gZY72s-gtPZabJL-yGIfJoQGUxexw-2wQ5iWxH8; ccis_sso_token=eyJhbGciOiJIUzI1NiJ9.eyJqdGkiOiJ7XCJ1c2VySWRcIjpcIjAxNDYzMkBjaXNkaS5jb20uY25cIn0iLCJzdWIiOiJjY2lz6aqM6K-BdG9rZW4iLCJpc3MiOiLpl6jmiLciLCJpYXQiOjE3MjEzNTQ4MTMsImV4cCI6MTcyMTM4MzkxM30.0hiJMbd32AG1aWvkG2FzEAzMigymBnRKzO_kwjqp_eg; oid=82BD2971CBFC45B3BF10B6E2819EF85F; formatedUserName=014632@cisdi.com.cn; last_login_user=014632; sto-id-24862=GAGEFIAKGBBO; workCode=014632; userId=014632@cisdi.com.cn; portal_session=NjY2; PRD=tnM0nABM9OHtNL0oxeXWPIet:S'
#     }

#     response = requests.request("POST", url, headers=headers, data=payload)

#     print(response.text)

# while True:
#     task()
#     time.sleep(60)  # 等待 60 秒