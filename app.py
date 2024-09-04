import os
import json
import base64
import requests

# 定义要创建的文件夹路径
folder_path = './files'
 
# 如果文件夹不存在，则创建文件夹
if not os.path.exists(folder_path):
    os.makedirs(folder_path)

def getFileList():
    url = "http://10.108.0.105:9999/asset-management-platform/api/apiDevelop/external/custom/api/oa/pa/zg/yjjh/list?appKey=2090000000008399559&appSecret=2206a222c4fd459f806a9255ee9332fa&enablePage=true&currPage=1&pageSize=10"
    payload = json.dumps({
    "project_number": 1
    })
    headers = {
    'Content-Type': 'application/json'
    }

    response = requests.request("POST", url, headers=headers, data=payload)
    return json.loads(response.text)

def outPutFile(requestid):
    url = "http://oa.cisdi.com.cn:8080/api/sdzb/getFileDate"
    payload = json.dumps({
        "requestid": requestid
    })
    headers = {
    'Content-Type': 'application/json'
    }

    response = requests.request("POST", url, headers=headers, data=payload)
    dataTemp = json.loads(response.text)
    # 假设这是你的Base64编码的字符串
    base64_string = dataTemp['resultData'][0]['base64']
    
    
    
    # 指定要创建的文件名
    file_name = dataTemp['resultData'][0]['filename']
    
    # 将解码后的数据写入文件
    if (not os.path.exists('./files/' + file_name)):
        # 解码Base64字符串
        decoded_data = base64.b64decode(base64_string)
        with open('./files/' + file_name, 'wb') as file:
            file.write(decoded_data)

fileList = getFileList()
print(fileList['data'])
fileList = fileList['data']['rowList']
for item in fileList:
    print(item['requestname'])
    outPutFile(item['requestid'])