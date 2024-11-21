import pygsheets
import discord
import datetime
import time
import requests
import pandas

gcredit = pygsheets.authorize(service_file='./pygsheet-api.json')
webhook_url = "WEBHOOK_URL_HERE"

testcase_list = ["前端測試紀錄 01"]
testcase_handler = "常數表"




while 1:

    sheet = gcredit.open_by_url('GOOGLE_SHEET_LINK')
    handler = sheet.worksheet_by_title(testcase_handler).get_all_records()

    print("[Connect] Google Sheet")    
    
    for test in testcase_list:

        print(f"[Search] {test}")
        
        cur = sheet.worksheet_by_title(test)
        current = cur.get_all_records()
        
        people = {}
        for i in range(len(handler)): people[handler[i]["組員"]] = []

        for testcase in range(1 , len(current)):

            print(f"[Checking] testcase {testcase}")

            name = current[testcase]['通知 @']
            if current[testcase]['已通知'] == "FALSE" and len(name) == 3:
                cur.update_value(f"N{testcase + 2}" , "TRUE")
                cur.update_value(f"O{testcase + 2}" , datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S"))
                people[name].append(testcase)
                

        print("[Done] Scan testcase")

        


        for person in range(len(handler)):

            print("[Checking]" , handler[person]["組員"])
        
            if len(people[handler[person]["組員"]]) == 0:
                continue

            data = {
                "username":"軟工測試報告通知",
                "avatar_url": "https://www.shutterstock.com/image-vector/chat-bot-icon-virtual-smart-600nw-2478937555.jpg",
                "content":f"<@{handler[person]["Discord User ID"]}>"
                }

            data["embeds"] = [
            {
                "description": "[報告書URL](GOOGLE_SHEET_LINK)",
                
                "title": "軟工測試報告通知",
                "color":0xfe7e06,
                "fields":[
                    
                ]
            }
            ]

            for field in people[handler[person]["組員"]]:

                data['embeds'][0]['fields'].append({
                    "name":f"{current[field]['target']} : {current[field]["測試內容"]}",
                    "value":f"{current[field]["測試結果"]}"
                })

            result = requests.post(webhook_url, json=data)
            print(f"[Webhook]")

        print(f"[Done] {test}")

    print("[Done] All Pages Have been searched")


    time.sleep(60)