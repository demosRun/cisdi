import requests

def download_file(url, local_filename):
    # 发送 HTTP GET 请求获取文件内容
    with requests.get(url, stream=True) as r:
        r.raise_for_status()  # 检查请求是否成功
        # 打开一个本地文件进行写入
        with open(local_filename, 'wb') as f:
            for chunk in r.iter_content(chunk_size=8192): 
                f.write(chunk)  # 将文件内容写入本地文件
    return local_filename

# 示例用法
url = 'http://oa.cisdi.com.cn:8080/weaver/weaver.file.FileDownload?f_weaver_belongto_userid=18694&f_weaver_belongto_usertype=0&fileid=4321544&download=1&requestid=3415521&desrequestid=0&fromrequest=1'
local_filename = 'file.xlsx'
download_file(url, local_filename)
print(f"Downloaded {local_filename}")
