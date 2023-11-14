
from flask import Flask, request
import pygsheets
import json
from datetime import datetime

app = Flask(__name__) 
app.config["DEBUG"] = True #會啟動Flask的Debug模式
app.config["JSON_AS_ASCII"] = False #解決Flask中文亂碼的問題


@app.route("/get", methods=["GET"])
def get_test():
    header = request.headers
    print(header)
    return "Get Success!!"

@app.route("/print", methods=["POST"])
def body_print():
    body = request.get_json()

    print (body)
    return "Print Success"

@app.route("/api_support", methods=["POST"])
def api_support():
    inputbody_dict = request.get_json()

    
    gc = pygsheets.authorize(service_file = "test.json")

    sht = gc.open_by_url('test')

    wks = sht.worksheet_by_title("統計")

    key_list = wks.get_col(1)
    value_list_key = ["key","attachment"]

    for i in range(len(inputbody_dict)):
        #print(inputbody_dict[i])
        value_list = []
        for k in value_list_key:
            value_list.append(inputbody_dict[i][k])
        
        row_num = 0
        try:
            row_num = key_list.index(inputbody_dict[i]["key"])+1
        except:
            row_num = key_list.index("")+1
        
        if inputbody_dict[i]["key"] in key_list:
            wks.update_row(index = row_num, values = value_list)
            status = "Replace"
            #print("Replace")
        else:
            try:
                wks.append_table(values = value_list)
            except:
                #print("Append")
                status = "Append"
        
        header = request.headers["X-Real-Ip"]
    
        now_time = datetime.now()

        log = f'{now_time} 狀態:{status} 行數:{row_num} 來源:{header}'
        
        #log寫入文件
        path = 'api-support.log'

        with open(path, 'a') as file:
            file.write(log+'\n')


    return "OK!!"


@app.route("/api_presales", methods=["POST"])
def api_presales():
    inputbody_dict = request.get_json()


    gc = pygsheets.authorize(service_file = "test.json")

    sht = gc.open_by_url('test')

    wks = sht.worksheet_by_title("統計")
    
    
    key_list = wks.get_col(1)
    row_num = 0
    value_list = []

    try:
        row_num = key_list.index(inputbody_dict["編號"])+1
    except:
        row_num = key_list.index("")+1

    #整理POST進來的資料
    value_list_key = ['1','2','3','4']

    for i in value_list_key:
        value_list.append(inputbody_dict[i])

    value_list.append(f'=PROPER(REGEXREPLACE(P{row_num}, "\.", " "))')
    value_list.append(f'=REGEXREPLACE(REGEXREPLACE(REGEXREPLACE(G{row_num}, "@highercloud.com.tw", ""),"\[",""),"\]","")')
    value_list.append(f'=M{row_num}-E{row_num}')
    value_list.append(inputbody_dict['4'])
    ########整理##########

    print(row_num)

    status = ""
    
    if inputbody_dict["編號"] in key_list:
        wks.update_row(index = row_num, values = value_list)
        status = "Replace"
        #print("Replace")
    else:
        try:
            wks.append_table(values = value_list)
        except:
            #print("Append")
            status = "Append"

    
    header = request.headers["X-Real-Ip"]
    
    now_time = datetime.now()

    log = f'{now_time} 狀態:{status} 行數:{row_num} 來源:{header}'
    
    #log寫入文件
    path = 'api-presales.log'
    with open(path, 'a') as file:
        file.write(log+'\n')    

    print(log)
    return "OK!!"

if __name__ == "__main__": #如果以主程式執行
    app.run(host="0.0.0.0", port=2000) #則立刻啟動伺服器 #run是Flask的Function，後面可以接host跟port。



