# -*- coding:utf-8 -*-

import os
import pprint
import sys

from PIL import Image, ImageDraw, ImageFont
from config import host as HOST, appId as APPID, clientIp as CLIENTIP, clientId as CLIENTID, requestId as REQUESTID, \
    code_height, code_width, name_width, name_height, name_font, code_font, name_size, code_size, name_fill, code_fill, \
    course_height, course_width, course_font, course_size, course_fill, char_len
from config import normal_template as NORMAL_TEMPLATE
from config import excellent_template as EXCELLENT_TEMPLATE
from config import course_excellent_template as COURSE_EXCELLENT_TEMPLATE
from config import normal_list_file as NORMAL_LIST_FILE
from config import excellent_list_file as EXCELLENT_LIST_FILE
from config import course_excellent_list_file as COURSE_EXCELLENT_LIST_FILE

import requests
import pandas as pd


def del_file(path):
    ls = os.listdir(path)
    for i in ls:
        c_path = os.path.join(path, i)
        if os.path.isdir(c_path):
            del_file(c_path)
        else:
            os.remove(c_path)

class CertiAutomaker(object):
    def __init__(self):
        # 七牛云参数配置信息
        self.host = HOST
        self.appid = APPID
        self.clientip = CLIENTIP
        self.clientid = CLIENTID
        self.requestid = REQUESTID
        self.uploadUrl = None
        self.data_form = None
        self.normal_list_file = NORMAL_LIST_FILE.decode("utf-8")
        # self.normal_list_file = NORMAL_LIST_FILE
        self.excellent_list_file = EXCELLENT_LIST_FILE.decode("utf-8")
        # self.excellent_list_file = EXCELLENT_LIST_FILE
        self.course_excellent_list_file = COURSE_EXCELLENT_LIST_FILE
        self.normal_certi_template = NORMAL_TEMPLATE
        self.excellent_certi_template = EXCELLENT_TEMPLATE
        self.course_excellent_certi_template = COURSE_EXCELLENT_TEMPLATE
        self.data_normal = None
        self.data_excellent = None

    def set_upload_url(self):
        print("开始连接七牛云")
        url = 'http://{0}/cstore/api/v2/uploadPolicy?appId={1}&clientIp={2}&clientId={3}&requestId={4}'.format(HOST,APPID,CLIENTIP,CLIENTID,REQUESTID)
        print(url)
        r = requests.get(url)
        data_dict = r.json()
        if data_dict['code'] != 0:
            print("连接七牛云失败")
            print(data_dict['message'])
        else:
            print("连接成功，正在获取上传所需appid,token,bcketid,key,token参数")
            policyList = data_dict['data']['policyList']
            formFields = policyList[0]['formFields']
            data_form = {}
            for filed in formFields:
                data_form[filed['key']] = filed['value']
            self.uploadUrl = policyList[0]['uploadUrl']
            self.data_form = data_form

    def render(self, is_normal):
        # 0 表示优秀学员
        if is_normal == 0:
            filename = self.excellent_list_file
        # 1 表示普通学员
        elif is_normal == 1:
            filename = self.normal_list_file
        # 2 表示课程证书学员
        elif is_normal == 2:
            filename = self.course_excellent_list_file
        else:
            filename = self.normal_list_file

        datas = pd.read_excel(filename, encoding="utf-8")

        datas['download_url'] = None
        for index, row in datas.iterrows():
            phone, name, code, course = row[0], row[1], row[2], row[3]
            phone = str(phone)[:11]
            if len(phone) == 11:
                download_url = self.set_water_text(self.normal_certi_template if is_normal == 1 else self.course_excellent_certi_template,
                                                   code, name, phone,course, is_normal)
                datas['download_url'][index] = download_url
            else:
                print("手机号格式非法")
                datas['download_url'][index] = ""
        if is_normal == 0:
            self.data_excellent = datas
        # 1 表示普通学员
        elif is_normal == 1:
            self.data_normal = datas
        # 2 表示课程证书学员
        elif is_normal == 2:
            self.data_course = datas

    # 设置文字水印
    def set_water_text(self, imagefile, code, name, phone, course, is_normal):
        img = Image.open(imagefile)
        draw = ImageDraw.Draw(img)
        (img_x, img_y) = img.size
        namefont = ImageFont.truetype(name_font, name_size)
        codefont = ImageFont.truetype(code_font, code_size)
        coursefont = ImageFont.truetype(course_font, course_size)
        zh_name = name
        if len(zh_name) == 3:
            nw = name_width * (img_x / 25)
        elif len(zh_name) == 4:
            nw = (name_width-0.5) * (img_x / 25)
        else:
            nw = (name_width + 0.7) * (img_x / 25)

        #调整课程名字大小
        if len(course) < char_len:
            cw = (13 - 0.5 * len(course)) * (img_x / 25)
            coursefont = ImageFont.truetype(course_font, course_size)
        elif len(course) == char_len:
            cw = (0.6 + 8.5) * (img_x / 25)
            size = 70
            coursefont = ImageFont.truetype(course_font, size)
        elif len(course) == char_len+1:
            cw = (0.6 + 8.5) * (img_x / 25)
            size = 60
            coursefont = ImageFont.truetype(name_font, size)
        elif len(course) == char_len+2:
            cw = (0.6 + 8.5) * (img_x / 25)
            size = 60
            coursefont = ImageFont.truetype(name_font, size)
        elif len(course) == char_len+3:
            cw = (0.6 + 8.5) * (img_x / 25)
            size = 55
            coursefont = ImageFont.truetype(name_font, size)
        elif len(course) == char_len+4:
            cw = (0.6 + 8.5) * (img_x / 25)
            size = 55
            coursefont = ImageFont.truetype(name_font, size)
        elif len(course) == char_len+5:
            cw = (0.6 + 8.5) * (img_x / 25)
            size = 55
            coursefont = ImageFont.truetype(name_font, size)
        elif len(course) == char_len+6:
            cw = (0.6 + 8.5) * (img_x / 25)
            size = 40
            coursefont = ImageFont.truetype(name_font, size)
        else:
            cw = (0.6 + 8.5) * (img_x / 25)
            size = 40
            coursefont = ImageFont.truetype(name_font, size)

        draw.text((nw, name_height), zh_name, font=namefont, fill=name_fill)
        # draw course
        if is_normal ==2:
            draw.text((cw, course_height), course, font=coursefont, fill=course_fill)
        # draw code
        draw.text((code_width, code_height), str(code), font=codefont, fill=code_fill)

        # img.show(imagefile)

        address = "./{person_type}/{phone}.jpg".format(
            person_type="person_normal" if is_normal == 1 else "person_excellent",
            phone=phone)
        address = address
        print("正在将证书{0}保存到本地.............".format(phone))
        img.save(address, "jpeg")
        print("保存到本地成功")
        return self.upload(address, phone)

    def upload(self, address,phone):
        self.set_upload_url()
        print("正在将本地证书上传到七牛云..............")
        files = {'file': open(address, 'rb')}
        r1 = requests.post(self.uploadUrl, files=files, data=self.data_form)
        # num = 0
        # while r1.status_code == 614:  # 文件已经存在
        #     num = num+1
        #     files = {'file': ("{phone}({count}).jpg".format(phone=phone,count=num), open(address, 'rb'))}
        #     r1 = requests.post(self.uploadUrl, files=files, data=self.data_form)
        download_url = r1.json()['data']['downloadUrl']
        print("上传成功")
        return download_url

    def to_csv(self):
        datas_combine = pd.concat([self.data_excellent, self.data_normal, self.data_course])
        datas_combine.to_csv("certificate.csv", encoding="gbk", index=False)

    def remove_tmp(self):
        del_path = r'./person_excellent'
        del_file(del_path)
        del_path = r'./person_normal'
        del_file(del_path)

if __name__ == '__main__':
    certiautomaker = CertiAutomaker()
    print("开始清除上次生成的证书图片缓存")
    certiautomaker.remove_tmp()
    print("开始上传优秀证书")
    certiautomaker.render(is_normal=0)
    print("开始上传普通证书")
    certiautomaker.render(is_normal=1)
    print("开始上传课程证书")
    certiautomaker.render(is_normal=2)
    print("正在生成下载链接")
    certiautomaker.to_csv()
    print("正在清除程序缓存")