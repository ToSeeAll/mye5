# -*- coding: UTF-8 -*-
import os

import requests as req
import json, sys, time

path = sys.path[0] + r'/1.txt'
num1 = 0


def gettoken(refresh, id, secret):
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    data = {'grant_type': 'refresh_token',
            'refresh_token': refresh,
            "client_id": id,
            'client_secret': secret,
            'redirect_uri': 'http://localhost:53682/'
            }
    html = req.post('https://login.microsoftonline.com/common/oauth2/v2.0/token', data=data, headers=headers)
    jsontxt = json.loads(html.text)
    refresh = jsontxt['refresh_token']
    access_token = jsontxt['access_token']
    with open(path, 'w+') as f:
        f.write(refresh)
    return access_token


def main():
    fo = open(path, "r+")
    refresh = fo.read()
    fo.close()
    global num1
    id = os.environ['CONFIG_ID']
    secret = os.environ['CONFIG_KEY']
    localtime = time.asctime(time.localtime(time.time()))
    access_token = gettoken(refresh, id, secret)
    headers = {
        'Authorization': access_token,
        'Content-Type': 'application/json'
    }
    try:
        if req.get(r'https://graph.microsoft.com/v1.0/me/drive/root', headers=headers).status_code == 200:
            num1 += 1
            print("1调用成功" + str(num1) + '次')
        if req.get(r'https://graph.microsoft.com/v1.0/me/drive', headers=headers).status_code == 200:
            num1 += 1
            print("2调用成功" + str(num1) + '次')
        if req.get(r'https://graph.microsoft.com/v1.0/drive/root', headers=headers).status_code == 200:
            num1 += 1
            print('3调用成功' + str(num1) + '次')
        if req.get(r'https://graph.microsoft.com/v1.0/users ', headers=headers).status_code == 200:
            num1 += 1
            print('4调用成功' + str(num1) + '次')
        if req.get(r'https://graph.microsoft.com/v1.0/me/messages', headers=headers).status_code == 200:
            num1 += 1
            print('5调用成功' + str(num1) + '次')
        if req.get(r'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules',
                   headers=headers).status_code == 200:
            num1 += 1
            print('6调用成功' + str(num1) + '次')
        if req.get(r'https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messages/delta',
                   headers=headers).status_code == 200:
            num1 += 1
            print('7调用成功' + str(num1) + '次')
        if req.get(r'https://graph.microsoft.com/v1.0/me/drive/root/children', headers=headers).status_code == 200:
            num1 += 1
            print('8调用成功' + str(num1) + '次')
        if req.get(r'https://api.powerbi.com/v1.0/myorg/apps', headers=headers).status_code == 200:
            num1 += 1
            print('8调用成功' + str(num1) + '次')
        if req.get(r'https://graph.microsoft.com/v1.0/me/mailFolders', headers=headers).status_code == 200:
            num1 += 1
            print('9调用成功' + str(num1) + '次')
        if req.get(r'https://graph.microsoft.com/v1.0/me/outlook/masterCategories', headers=headers).status_code == 200:
            num1 += 1
            print('10调用成功' + str(num1) + '次')
            print('此次运行结束时间为 :', localtime)
    except:
        print("pass")
        pass


if __name__ == '__main__':
    for _ in range(3):
        main()
