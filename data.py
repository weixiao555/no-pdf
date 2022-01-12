# --coding:utf-8--
import random
import re
import base64
import json
import os
import sys
import time
import xlrd
import xlwt
import requests
from fake_useragent import UserAgent
from xlutils.copy import copy as xl_copy
from selenium import webdriver
from PIL import Image
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By


def encrypt(string):
    param = string.encode('utf-8')
    encryption = base64.b64encode(param)
    result = encryption.decode('utf-8')
    return result


class MyCookie:
    print_screen_path = 'resource/login_picture/print_screen.png'
    save_path = 'resource/login_picture/save.png'

    # 初始化driver
    def __init__(self):  # 传入图片所在的文件夹
        chrome_options = Options()
        chrome_options.add_argument('--headless')
        self.driver = webdriver.Chrome(options=chrome_options)
        self.driver.get('http://wzjcjg.jzyglxt.com/')
        self.driver.maximize_window()

    # 通过selenium登录来获取cookie,# 外部只需调用这个方法即可
    def selenium_login(self, _user, _pass):
        self.driver.find_element(By.ID, 'userName').send_keys(_user)
        self.driver.find_element(By.ID, 'passWord').send_keys(_pass)
        self.get_yzm_png()
        # (原本out()后面加个'i',故意弄错测试)
        yzm = self.out()
        self.driver.find_element(By.ID, 'yzm').send_keys(yzm)
        while True:
            self.driver.find_element(By.ID, 'login').click()
            time.sleep(2)
            result = self.is_element_exist('//*[@id="layui-layer2"]/div[3]/a')
            if not result:
                print('成功登录')
                time.sleep(5)
                cookies = self.driver.get_cookies()
                time.sleep(5)
                try:
                    pass_0 = cookies[0]['value']
                    user0 = cookies[1]['value']
                    xc = cookies[2]['value']
                    asp = cookies[3]['value']
                    cookie = 'ASP.NET_SessionId=' + asp + ';.xCookie=' + xc + ';user=' + user0 + ';pass=' + pass_0
                except Exception:
                    print('cookie获取失败!!')
                    print('1.请检查"账号"或者"密码"是否错误!!')
                    print('2.请检查网络是否有问题!!')
                    print('3.该网站服务端响应异常!!')
                    print('请重新运行!!')
                    sys.exit(1)
                self.driver.close()
                return cookie
                # break
            else:
                print('验证码错误')
                time.sleep(2)
                self.driver.find_element(By.CLASS_NAME, 'layui-layer-btn0').click()
                self.driver.find_element(By.XPATH, '//*[@id="mainBody"]/div[2]/div[1]/div[3]/img').click()
                self.get_yzm_png()
                yzm2 = self.out()
                self.driver.find_element(By.ID, 'yzm').clear()
                self.driver.find_element(By.ID, 'yzm').send_keys(yzm2)
        # print('退出循环')
        # time.sleep(120)

    # 打码接口
    @staticmethod
    def base64_api(uname, pwd, img):
        with open(img, 'rb') as f:
            base64_data = base64.b64encode(f.read())
            b64 = base64_data.decode()
        data = {"username": uname, "password": pwd, "image": b64}
        result = json.loads(requests.post("http://api.ttshitu.com/base64", json=data).text)
        if result['success']:
            return result["data"]["result"]
        else:
            print('打码平台欠费,无法解析验证码,请先充值')
            print('请前往"http://www.ttshitu.com/"充值')
            # return result["message"]
        # return ""

    # 打码后的输出结果
    def out(self):
        img_path = self.save_path
        result = MyCookie.base64_api(uname='tangsan', pwd='future827', img=img_path)
        print("验证码:" + result)
        return result

    # 问题所在,始终为False
    # 判断验证码是否正确(判断元素是否存在)
    def is_element_exist(self, element):
        flag = True
        try:
            self.driver.find_element(By.XPATH, element)
            return flag
        except Exception:
            flag = False
            return flag

    # 获取验证码图片
    def get_yzm_png(self):
        self.driver.save_screenshot(self.print_screen_path)
        image_element = self.driver.find_element(By.XPATH, '//*[@id="mainBody"]/div[2]/div[1]/div[3]/img')  # 定位验证码
        location = image_element.location  # 获取验证码x,y轴坐标
        # print(location)
        size = image_element.size  # 获取验证码的长宽
        # print(size)
        rangle = (int(location['x']), int(location['y']), int(location['x'] + size['width']),
                  int(location['y'] + size['height']))  # 写成我们需要截取的位置坐标
        i = Image.open(self.print_screen_path)  # 打开截图
        frame4 = i.crop(rangle)  # 使用Image的crop函数，从截图中再次截取我们需要的区域
        frame4.save(self.save_path)  # 保存我们接下来的验证码图片 进行打码


class MyParam:
    def __init__(self):
        self.page = 1  # 默认只有一页数据 少于1000条,如果有多页,调用的地方循环

    # 获取委托单唯一号，类别，监管等级编号，委托单编号，检测结果等的参数
    def get_few_param(self, cookie, project_number):
        headers = {
            'Cookie': '{}'.format(cookie),
            "Host": "wzjcjg.jzyglxt.com",
            "Origin": "http://wzjcjg.jzyglxt.com",
            "Referer": "http://wzjcjg.jzyglxt.com/WebList/EasyUiIndex?FormDm=WTDGL&FormStatus=0&FormParam=PARAM--{}|ALL|".format(
                project_number),
            "User-Agent": UserAgent().chrome,
        }
        param = '{"isbackground":false,"encode":false,"FormDm":"WTDGL","FormStatus":"0",' \
                '"FormParam":' + '"PARAM--{0}|ALL|"'.format(project_number) + ',"FormHidden":"",' \
                                                                              '"FormFilter":"","FormZd":"","CheckSession":"","Log":""}'
        encryption = encrypt(param)  # 调用的base64_service
        form_data = {
            "param": '{}'.format(encryption),
            "page": self.page,  # 如果有超过一千条数据的请求,还是先拿到form_data,page就不写死,调用的地方循环
            "rows": "1000",
            "filterRules": "[]"
        }
        return {'headers': headers, 'form_data': form_data}

    # 获取"查看"的数据的参数
    def get_more_param(self, cookie, project_number):
        result = self.get_few_param(cookie, project_number)
        headers = result['headers']
        form_data = result['form_data']
        index_headers = self.get_index_headers()  # index页面的headers
        input_data_list = self.get_input(headers, form_data)  # index页面的form_data
        param_list = self.get_param(index_headers, input_data_list)
        return {'headers': headers, 'form_data': param_list}

    # 查看 ==1==，自己生成datainput/Index页面所需参数模板（里面包含查看的加密参数）
    def get_input(self, headers, form_data):
        time.sleep(0.5)
        response = requests.post('http://wzjcjg.jzyglxt.com/WebList/SearchEasyUiFormData', headers=headers,
                                 data=form_data)
        dictionaries = eval(response.text)['rows']
        # print('dictionaries', dictionaries)
        input_data_list = []
        for element in dictionaries:  # 只保留需要的,防止后面"查看"获取一些无用的数据,浪费时间,请求一个查看需要1~2秒的时间,比较久
            if element['SYXMBH'] in "FS/HNT_JC/HNT_ZT/SJX_JC/SJX_ZT/GYC/GHJ_JC/GHJ_ZT/SN/SZ/SPB/KS/HSA/TG/TYH":
                companycode = element['YTDWBH']
                jydbh = element['RECID']
                random_number = str(random.uniform(0, 1))
                syxmbh = element['SYXMBH']
                t1_tablename = "M_BY,M_D_" + syxmbh + ",M_" + syxmbh
                t2_tablename = "S_BY,S_D_" + syxmbh + ",S_" + syxmbh
                zdzdtable = "XTZD_BY,DWZD_" + syxmbh + ",ZDZD_" + syxmbh
                fieldparam = "M_BY,SYXMBH," + syxmbh + "|M_BY,JG_XMDH," + 'M5' + " |M_BY,SCWTS,1" + "|M_BY,SCWTSDZ,"
                input_data = {
                    'zdzdtable': zdzdtable,  # 此处为项目类别
                    't1_tablename': t1_tablename,  # 此处为项目类别
                    't1_pri': 'RECID',  # 都一样
                    't1_title': '委托单',  # 都一样
                    'LX': 'W',  # 都一样
                    'rownum': 2,  # 都一样
                    't2_tablename': t2_tablename,  # 此处为项目类别
                    'fieldparam': fieldparam,  # 细微不一样
                    't2_pri': 'BYZBRECID,RECID',  # 都一样
                    't2_title': '记录',  # 都一样
                    't3_tablename': '',  # 都一样
                    't3_pri': '',  # 都一样
                    't3_title': '',  # 都一样
                    'type': 'leftright',  # 都一样
                    '_': random_number,
                    'companycode': companycode,
                    'sylbzdzd': syxmbh,  # 此处为项目类别
                    'individualZdzdtable': 'DATAZDZD_INDIVIDUAL',  # 都一样
                    'syxmbh': syxmbh,  # 此处为项目类别
                    't2_syxmbh': '',  # 都一样
                    't2_syxmdh': '',  # 都一样
                    't2_syxmmc': '',  # 都一样
                    'jydbh': jydbh,
                    't2_order': 'len(zh),zh',  # 都一样
                    't2_orderseq': ' asc,asc',  # 都一样
                    'individualProjectZdzdtable': 'DATAZDZD_INDIVIDUALSYXM',  # 都一样
                    'view': 'true',  # 都一样
                }
                input_data_list.append(input_data)
        return input_data_list

    # 请求Index页面的请求头
    def get_index_headers(self):
        headers = {
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/'
                      'signed-exchange;v=b3;q=0.9',
            'Accept-Encoding': 'gzip, deflate',
            'Accept-Language': 'zh-CN,zh-HK;q=0.9,zh;q=0.8,en;q=0.7,en-GB;q=0.6,en-US;q=0.5',
            'Connection': 'keep-alive',
            'Cookie': 'user=huangxiao; pass=88888; ASP.NET_SessionId=ul2incwbbx4ogegu1d5judm3; .xCookie=8ABA5154C1AAAE72E'
                      'D9D8961E27FC0711C9346D082BBEFB1EA5D5EA95ADD7513E7F0B71D13D308CC9DF51CD7F3245BE2C63BCFFDCE7CC27322B'
                      '893BF15B0737984E63DDB565B250104C2882DBD0EF6EC5319C7FF804162BB736BD17689D2E91F312001AA1242499EF62B8'
                      'D2F83F42A5C5F0ACE42229CE84787E268DC382C1262F59A4ECE8CB4608E43B50B98447A444B',
            'DNT': "1",
            'Host': 'wzjcjg.jzyglxt.com',
            'Referer': 'http://wzjcjg.jzyglxt.com/user/mainnew',
            'Upgrade-Insecure-Requests': '1',
            'User-Agent': UserAgent().chrome,
        }
        return headers

    # 查看==2==  查看的加密参数就来自这个页面
    def get_param(self, headers, form_data):
        url = 'http://wzjcjg.jzyglxt.com/datainput/Index'
        param_list = []
        print('正在获取参数...')
        for i in range(len(form_data)):
            print("\r%s" % (len(form_data) - i), end='')
            while True:
                try:
                    # time.sleep(0.1)
                    response = requests.get(url=url, headers=headers, data=form_data[i], timeout=1)
                    html = response.text
                    # print(html)
                    find_value = re.compile(r'<input type="hidden" id="param" name="param" value="(.*?)"')
                    param_value = re.findall(find_value, html)[0]
                    param = {'param': param_value}
                    param_list.append(param)
                    break
                except Exception:
                    print('获取参数失败,正在重新获取参数!!')
        print('\r已获取全部参数！')
        return param_list


class MyGet:
    def __init__(self):
        self.few_data_list = []
        self.more_data_list = []

    # 获取委托单唯一号，类别，监管等级编号，委托单编号，检测结果等粗略数据,无关紧要的数据
    def get_few_data(self, headers, form_data):
        response = requests.post('http://wzjcjg.jzyglxt.com/WebList/SearchEasyUiFormData', headers=headers,
                                 data=form_data)
        data = eval(response.text)['rows']
        print('本次一共有：' + str(len(data)) + '个数据')
        project_name = data[0]['GCMC']  # 顺便把工程名称获取了
        for i in data:
            if i['SYXMBH'] in "FS/HNT_JC/HNT_ZT/SJX_JC/SJX_ZT/GYC/GHJ_JC/GHJ_ZT/SN/SZ/SPB/KS/HSA/TG/TYH":  # TYH：混凝土检测养护
                self.few_data_list.append([
                    i['RECID'],  # 委托单唯一号          0
                    i['SYXMBH'],  # 类别               1
                    i['WTDBH'],  # 监管登记编号          2////////
                    i['SLWTDBH'],  # 委托单编号         3
                    i['JCJGMS']  # 检测结果             4
                ])
        print('一共有：' + str(len(self.few_data_list)) + ' 个有效数据')
        return {'project_name': project_name, 'few_data_list': self.few_data_list}

    def get_more_data(self, headers, param_list):
        print('正在请求页面...')
        for i in range(len(param_list)):
            print("\r%s" % (len(param_list) - i), end='')
            time.sleep(0.1)
            while True:
                try:
                    response = requests.post('http://wzjcjg.jzyglxt.com/DataInput/SearchData', headers=headers,
                                             data=param_list[i],
                                             timeout=3)
                    self.more_data_list.append(response.json()['data'])
                    break
                except Exception:
                    print('请求页面失败,正在重新请求页面!!')
        return self.more_data_list


class MyParse:
    def __init__(self):
        self.all_data = {
            "FS": [],  # 防水
            "HNT_JC": [],  # 混凝土试块汇总表标养（基础）
            "HNT_ZT": [],  # 混凝土试块汇总表标养（主体）
            "SJX_JC": [],  # 砂浆试块试验报告汇总附表（基础）
            "SJX_ZT": [],  # 砂浆试块试验报告汇总附表（主体）
            "GYC": [],  # 钢材原材汇总表
            "GHJ_JC": [],  # 钢筋焊接试验报告汇总表（基础）
            "GHJ_ZT": [],  # 钢筋焊接试验报告汇总表（主体）
            "SN": [],  # 水泥试验报告汇总附表
            "SZ": [],  # 砖试验报告汇总附表
            "SPB": [],  # 砼、砂浆配合比汇总表
            "KS": [],  # 混凝土抗渗试块试验报告汇总表
            "HSA": [],  # 砂、石试验报告汇总表
            "TG": []  # 电工套管汇总表
        }

    # 外部只需调用这个方法即可
    def sort_service(self, few_data_list, more_data_list):
        for i in range(len(few_data_list)):
            self.sort(few_data_list[i], more_data_list[i])
        return self.all_data

    # 拿到的接口数据分类  "查看"的接口数据
    def sort(self, one_few_data, one_more_data):  # 传入的为一条数据
        kind = one_more_data['t1hidden'][0]['defval']
        # if sort in "FS/HNT_JC/HNT_ZT/SJX_JC/SJX_ZT/GYC/GHJ_JC/GHJ_ZT/SN/SZ/SPB/KS/HSA/TG/TYH":
        #     print(sort)
        if kind == 'FS':
            self.save_fs_data(one_few_data, one_more_data)
        elif kind == 'HNT' or kind == 'TYH':
            self.save_hnt_data(one_few_data, one_more_data)
        elif kind == 'SJX':
            self.save_sjx_data(one_few_data, one_more_data)
        elif kind == 'GYC':
            self.save_gyc_data(one_few_data, one_more_data)
        elif kind == 'GHJ':
            self.save_ghj_data(one_few_data, one_more_data)
        elif kind == 'SN':
            self.save_sn_data(one_few_data, one_more_data)
        elif kind == 'SZ':
            self.save_sz_data(one_few_data, one_more_data)
        elif kind == 'HSA':
            self.save_hsa_data(one_few_data, one_more_data)
        elif kind == 'SPB':
            self.save_spb_data(one_few_data, one_more_data)
        elif kind == 'KS':
            self.save_ks_data(one_few_data, one_more_data)
        elif kind == 'TG':
            self.save_tg_data(one_few_data, one_more_data)
        else:
            pass

    def save_fs_data(self, one_few_data, one_more_data):  # 防水
        for x in range(len(one_more_data['t2datas'])):
            data = [
                one_more_data['t2datas'][x]['t2data'][5]['zdval'],  # 材料名称
                '',  # 规格型号
                one_more_data['t2datas'][x]['t2data'][18]['zdval'].strip('有限公司'),  # 生产厂家
                '500',  # 施工面积
                one_more_data['t2datas'][x]['t2data'][3]['zdval'],  # 进场日期（要求试验日期）
                '',  # 试验报告日期
                '',  # 试验报告单编号
                one_few_data[4],  # 检测结果
                one_more_data['t2datas'][x]['t2data'][2]['zdval'],  # 主要使用部位
                one_more_data['t2datas'][x]['t2data'][0]['zdval'],  # 报审表编号
                one_few_data[2],  # 监管登记编号
                one_few_data[3],  # 委托单编号
                one_few_data[0]  # 委托单编号
            ]
            self.all_data["FS"].append(data)

    def save_sjx_data(self, one_few_data, one_more_data):  # 砂浆
        for x in range(len(one_more_data['t2datas'])):
            data = [
                one_more_data['t2datas'][x]['t2data'][2]['zdval'],  # 工程部位
                one_more_data['t2datas'][x]['t2data'][7]['zdval'],  # 设计标号
                one_more_data['t2datas'][x]['t2data'][4]['zdval'],  # 试块制作日期（制作日期）
                one_more_data['t2datas'][x]['t2data'][5]['zdval'],  # 龄期
                '',  # 试验报告单编号
                '',  # 实测强度值(Mu)
                '',  # 达到设计强度(%)
                '',  # 水泥出厂日期
                '',  # 水泥批号
                '',  # 配合比出单日期
                one_more_data['t2datas'][x]['t2data'][0]['zdval'],  # 报审表编号
                one_few_data[2],  # 监管登记编号
                one_few_data[3],  # 委托单编号
                one_few_data[0]  # RECID
            ]
            jc_string = '基础'
            # book = xl_copy(xlrd.open_workbook(save_path, formatting_info=True))
            if jc_string in one_more_data['t2datas'][0]['t2data'][2]['zdval']:
                self.all_data["SJX_JC"].append(data)
                # sheet = book.get_sheet("砂浆试块试验报告汇总附表（基础）")
                # count = save_data_length['SJX_JC'] + 2
                # save_data_length['SJX_JC'] += 1
            else:
                self.all_data["SJX_ZT"].append(data)

    def save_sn_data(self, one_few_data, one_more_data):  # 水泥
        for x in range(len(one_more_data['t2datas'])):
            data = [
                '',  # 水泥标号
                '200',  # 数量
                one_more_data['t2datas'][x]['t2data'][7]['zdval'].strip('有限公司'),  # 生产厂家及批号
                one_more_data['t2datas'][x]['t2data'][10]['zdval'],  # 出厂日期
                '',  # 安定性试验单编号
                '',  # 安定性试验结果
                '',  # 强度试验报告日期
                '',  # 强度试验单编号
                '',  # 强度试验结果
                '',  # 主要使用部位
                one_more_data['t2datas'][x]['t2data'][0]['zdval'],  # 报审表编号
                one_few_data[2],  # 监管登记编号
                one_few_data[3],  # 委托单编号
                one_few_data[0]  # RECID
            ]
        self.all_data["SN"].append(data)

    def save_hsa_data(self, one_few_data, one_more_data):  # 砂、石试验
        for x in range(len(one_more_data['t2datas'])):
            data = [
                '',  # 品种
                '',  # 规格
                '',  # 试验报告单编号
                '',  # 结论
                one_more_data['t2datas'][x]['t2data'][3]['zdval'],  # 进场时间
                '100',  # 数量
                one_few_data[2],  # 监管登记编号
                one_few_data[3],  # 委托单编号
                one_few_data[0]  # RECID
            ]
        self.all_data["HSA"].append(data)

    def save_spb_data(self, one_few_data, one_more_data):  # 砼、砂浆配合比
        for x in range(len(one_more_data['t2datas'])):
            qd = re.sub(".*?砂浆", '', one_more_data['t2datas'][x]['t2data'][5]['zdval'])
            zl = re.sub("M.*?", '', one_more_data['t2datas'][x]['t2data'][5]['zdval'])
            data = [
                qd,  # 强度
                zl,  # 种类
                one_more_data['t2datas'][x]['t2data'][2]['zdval'],  # 使用部位
                '',  # 试验报告单编号
                '',  # 报告日期
                '重量配合比',  # 项目
                '',  # 水泥
                '',  # 水
                '',  # 砂
                '',  # 灰膏
                '',  # 外加剂
                one_few_data[2],  # 监管登记编号
                one_few_data[3],  # 委托单编号
                one_few_data[0]  # RECID
            ]
            self.all_data["SPB"].append(data)

    def save_gyc_data(self, one_few_data, one_more_data):  # 钢材原材
        for x in range(len(one_more_data['t2datas'])):
            data = [
                '',  # 规格品种
                one_more_data['t2datas'][x]['t2data'][10]['zdval'],  # 规格品种
                one_more_data['t2datas'][x]['t2data'][9]['zdval'],  # 进场数量（吨）
                one_more_data['t2datas'][x]['t2data'][7]['zdval'].strip('集团有限公司'),  # 生产厂家
                '',  # 质保单出厂日期
                one_more_data['t2datas'][x]['t2data'][8]['zdval'],  # 质保书炉批号
                one_more_data['t2datas'][x]['t2data'][3]['zdval'],  # 进场日期
                '',  # 试验报告单日期
                '',  # 试验报告单编号
                one_few_data[4],  # 检测结果
                one_more_data['t2datas'][x]['t2data'][2]['zdval'],  # 主要使用部位
                one_more_data['t2datas'][x]['t2data'][0]['zdval'],  # 报审表编号
                one_few_data[2],  # 监管登记编号
                one_few_data[3],  # 委托单编号
                one_few_data[0]  # RECID
            ]
            self.all_data["GYC"].append(data)

    def save_ghj_data(self, one_few_data, one_more_data):  # 钢筋焊接试验
        for x in range(len(one_more_data['t2datas'])):
            data = [
                '',  # 规格品种
                one_more_data['t2datas'][x]['t2data'][9]['zdval'],  # 规格品种
                one_more_data['t2datas'][x]['t2data'][4]['zdval'],  # 焊接类型
                'E55',  # 使用何种焊条(剂)
                '300',  # 焊接头数量(个)
                # 后面加上的，最开始没有
                one_more_data['t2datas'][x]['t2data'][3]['zdval'],  # 进场日期
                '',  # 试验报告单日期
                '',  # 试验报告单编号
                one_few_data[4],  # 检测结果
                one_more_data['t2datas'][x]['t2data'][2]['zdval'],  # 主要使用部位
                one_more_data['t2datas'][x]['t2data'][0]['zdval'],  # 报审表编号
                one_few_data[2],  # 监管登记编号
                one_few_data[3],  # 委托单编号
                one_few_data[0]  # RECID
            ]
            jc_string = '基础'
            # book = xl_copy(xlrd.open_workbook(save_path, formatting_info=True))
            if jc_string in one_more_data['t2datas'][0]['t2data'][2]['zdval']:
                self.all_data["GHJ_JC"].append(data)
            else:
                self.all_data["GHJ_ZT"].append(data)

    def save_hnt_data(self, one_few_data, one_more_data):  # 混凝土试块
        for x in range(len(one_more_data['t2datas'])):
            data = [
                one_more_data['t2datas'][x]['t2data'][2]['zdval'],  # 工程部位
                one_more_data['t2datas'][x]['t2data'][6]['zdval'],  # 设计标号
                '',  # 浇捣数量
                one_more_data['t2datas'][x]['t2data'][3]['zdval'],  # 试块成形日期
                '28',  # 龄期(默认就是28,因为有的读出来是0)
                '',  # 试验报告单编号
                '',  # 实测强度
                '',  # 检测结果
                '',  # 水泥出厂日期
                '',  # 水泥批号
                '',  # 配合比出单日期
                one_more_data['t2datas'][x]['t2data'][0]['zdval'],  # 报审表编号
                one_few_data[2],  # 监管登记编号
                one_few_data[3],  # 委托单编号
                one_few_data[0]  # RECID
            ]
            jc_string = '基础'
            # book = xl_copy(xlrd.open_workbook(save_path, formatting_info=True))
            if jc_string in one_more_data['t2datas'][0]['t2data'][2]['zdval']:
                self.all_data["HNT_JC"].append(data)
            else:
                self.all_data["HNT_ZT"].append(data)

    def save_ks_data(self, one_few_data, one_more_data):  # 混凝土抗渗试块试验
        for x in range(len(one_more_data['t2datas'])):
            data = [
                one_more_data['t2datas'][x]['t2data'][2]['zdval'],  # 工程部位
                one_more_data['t2datas'][x]['t2data'][6]['zdval'],  # 设计等级
                '',  # 浇捣数量
                one_more_data['t2datas'][x]['t2data'][3]['zdval'],  # 试块成形日期
                one_more_data['t2datas'][x]['t2data'][4]['zdval'],  # 龄期
                '',  # 试验报告日期
                '',  # 试验报告单编号
                one_few_data[4],  # 检测结果
                '',  # 水泥出厂日期
                '',  # 水泥批号
                '',  # 配合出单日期
                one_more_data['t2datas'][x]['t2data'][0]['zdval'],  # 报审表编号
                one_few_data[2],  # 监管登记编号
                one_few_data[3],  # 委托单编号
                one_few_data[0]  # 委托单编号
            ]
            self.all_data["KS"].append(data)

    def save_sz_data(self, one_few_data, one_more_data):  # 砖试验报告
        for x in range(len(one_more_data['t2datas'])):
            data = [
                one_more_data['t2datas'][x]['t2data'][4]['zdval'],  # 种类
                one_more_data['t2datas'][x]['t2data'][10]['zdval'].strip('有限公司'),  # 生产厂家
                '10',  # 进场数量
                one_more_data['t2datas'][x]['t2data'][5]['zdval'],  # 强度等级
                one_more_data['t2datas'][x]['t2data'][3]['zdval'],  # 进场日期
                '',  # 试验报告日期
                '',  # 试验报告单编号
                '',  # 实测强度值
                one_more_data['t2datas'][x]['t2data'][2]['zdval'],  # 主要使用部位
                one_more_data['t2datas'][x]['t2data'][0]['zdval'],  # 报审表编号
                one_few_data[2],  # 监管登记编号
                one_few_data[3],  # 委托单编号
                one_few_data[0]  # 委托单编号
            ]
            self.all_data["SZ"].append(data)

    def save_tg_data(self, one_few_data, one_more_data):  # 电工套管试验报告
        for x in range(len(one_more_data['t2datas'])):
            data = [
                one_more_data['t2datas'][x]['t2data'][4]['zdval'],  # 材料名称
                one_more_data['t2datas'][x]['t2data'][5]['zdval'],  # 规格及型号
                one_more_data['t2datas'][x]['t2data'][9]['zdval'],  # 数量
                one_more_data['t2datas'][x]['t2data'][3]['zdval'],  # 进场日期
                '',  # 试验报告单编号
                one_few_data[4],  # 检测结果
                '',  # 部位
                one_more_data['t2datas'][x]['t2data'][0]['zdval'],  # 报审表编号
                one_few_data[2],  # 监管登记编号
                one_few_data[3],  # 委托单编号
                one_few_data[0]  # 委托单编号
            ]
            self.all_data["TG"].append(data)


class MySave:
    web_json_path = 'resource/json/web_json.json'

    def __init__(self):
        self.save_data_length = {
            "FS": 0,
            "HNT_JC": 0,
            "HNT_ZT": 0,
            "SJX_JC": 0,
            "SJX_ZT": 0,
            "GYC": 0,
            "GHJ_JC": 0,
            "GHJ_ZT": 0,
            "SN": 0,
            "SZ": 0,
            "SPB": 0,
            "KS": 0,
            "HSA": 0,
            "TG": 0
        }

    # 外部只需要调用此方法即可
    def my_save(self, all_data, project_name, save_path):  # json数据,项目名，保存路径
        self.save_all_data_json(all_data)
        self.create_xls(project_name, save_path)
        self.mange_xls(all_data, save_path)

    # 保存为json,后续和白描合并需要用
    def save_all_data_json(self, all_data):
        b = json.dumps(all_data, ensure_ascii=False)
        all_json = open(self.web_json_path, 'w', encoding='utf-8')
        all_json.write(b)
        all_json.close()

    def create_xls(self, project_name, save_path):
        style_bold = xlwt.easyxf('font: color-index Black, bold on,height 200')
        book = xlwt.Workbook(encoding="utf-8", style_compression=0)
        sheets = [
            book.add_sheet("防水", cell_overwrite_ok=True),
            book.add_sheet("砂浆试块试验报告汇总附表（基础）", cell_overwrite_ok=True),
            book.add_sheet("砂浆试块试验报告汇总附表（主体）", cell_overwrite_ok=True),
            book.add_sheet("水泥试验报告汇总附表", cell_overwrite_ok=True),
            book.add_sheet("砂、石试验报告汇总表", cell_overwrite_ok=True),
            book.add_sheet("砼、砂浆配合比汇总表", cell_overwrite_ok=True),
            book.add_sheet("钢材原材汇总表", cell_overwrite_ok=True),
            book.add_sheet("钢筋焊接试验报告汇总表（基础）", cell_overwrite_ok=True),
            book.add_sheet("钢筋焊接试验报告汇总表（主体）", cell_overwrite_ok=True),
            book.add_sheet("混凝土试块汇总表标养（基础）", cell_overwrite_ok=True),
            book.add_sheet("混凝土试块汇总表标养（主体）", cell_overwrite_ok=True),
            book.add_sheet("混凝土抗渗试块试验报告汇总表", cell_overwrite_ok=True),
            book.add_sheet("砖试验报告汇总附表", cell_overwrite_ok=True),
            book.add_sheet("电工套管汇总表", cell_overwrite_ok=True),
        ]
        title_list = [
            ['序号', '材料名称', '规格型号', '生产厂家', '施工面积(M²)', '进场日期', '试验报告日期', '试验报告单编号', '检测结果',
             '主要使用部位', '报审表编号', '监管登记编号', '委托单编号'],  # 防水
            ['序号', '工程部位', '设计标号', '试块制作日期', '龄期(天)', '试验报告单编号', '实测强度值(MPa)', '达到设计强度(%)',
             '水泥出厂日期', '水泥批号', '配合比出单日', '报审表编号', '监管登记编号', '委托单编号'],
            ['序号', '工程部位', '设计标号', '试块制作日期', '龄期(天)', '试验报告单编号', '实测强度值(MPa)', '达到设计强度(%)',
             '水泥出厂日期', '水泥批号', '配合比出单日', '报审表编号', '监管登记编号', '委托单编号'],
            ['序号', '水泥标号', '数量(吨)', '生产厂家及批号', '出厂日期', '安定性试验单编号', '安定性试验结果', '强度试验报告日期',
             '强度试验单编号', '强度试验结果', '主要使用部位', '报审表编号', '监管登记编号', '委托单编号'],
            ['序号', '品种', '规格(mm)', '试验报告单编号', '结论', '进场日期', '数量(m³)', '监管登记编号', '委托单编号'],
            ['序号', '强度', '种类', '使用部位', '试验报告单编号', '报告日期', '项目', '水泥', '水', '砂', '灰膏', '外加剂',
             '监管登记编号', '委托单编号'],  # 砼、砂浆配合比汇总表
            ['序号', '规格品种', '', '进场数量(吨)', '生产厂家', '质保单出厂日期', '质保单炉批号', '进场日期', '试验报告日期',
             '试验报告单编号', '试验结果', '主要使用部位', '报审表编号', '监管登记编号', '委托单编号'],  # 钢材原材汇总表
            ['序号', '规格品种', '', '焊接类型', '使用何种焊条(剂)', '焊接头数量(个)', '进场日期', '试验报告日期', '试验报告单编号', '试验结果',
             '主要使用部位', '报审表编号', '监管登记编号', '委托单编号'],  # 钢筋焊接试验报告汇总表（基础）
            ['序号', '规格品种', '', '焊接类型', '使用何种焊条(剂)', '焊接头数量(个)', '进场日期', '试验报告日期', '试验报告单编号', '试验结果',
             '主要使用部位', '报审表编号', '监管登记编号', '委托单编号'],  # 钢筋焊接试验报告汇总表（主体）
            ['序号', '工程部位', '设计标号', '浇捣数量(m³)', '试块成形日期', '龄期(天)', '试验报告单编号', '实测强度值(Mpa)',
             '检测结果(%)', '水泥出厂日期', '水泥批号', '配合比出单日期', '报审表编号', '监管登记编号', '委托单编号'],  # 混凝土试块汇总表标养（基础）
            ['序号', '工程部位', '设计标号', '浇捣数量(m³)', '试块成形日期', '龄期(天)', '试验报告单编号', '实测强度值(Mpa)',
             '检测结果(%)', '水泥出厂日期', '水泥批号', '配合比出单日期', '报审表编号', '监管登记编号', '委托单编号'],  # 混凝土试块汇总表标养（主体）
            ['序号', '工程部位', '设计标号', '浇捣数量(m³)', '试块成形日期', '龄期(天)', '试验报告日期', '试验报告单编号',
             '检测结果', '水泥出厂日期', '水泥批号', '配合比出单日期', '报审表编号', '监管登记编号', '委托单编号'],  # 混凝土抗渗试块试验报告汇总表
            ['序号', '种类', '生产厂家', '进场数量(万块)', '强度等级(Mu)', '进场日期', '试验报告日期', '试验报告单编号',
             '实测强度(Mu)', '主要使用部位', '报审表编号', '监管登记编号', '委托单编号'],  # 砖试验报告汇总附表
            ['序号', '材料名称', '规格及型号', '数量', '进场日期', '试验报告单编号', '试验结果', '部位', '报审表编号',
             '监管登记编号', '委托单编号'],  # 电工套管汇总表
        ]
        for i in range(len(sheets)):
            sheets[i].write(0, 0, project_name, style_bold)
            for j in range(len(title_list[i])):
                sheets[i].write(1, j, title_list[i][j], style_bold)
        book.save(save_path)

    def mange_xls(self, all_data, save_path):
        for sort in all_data:  # all_data是所有的数据(无解析)
            if sort == 'FS':
                self.save_fs_xls(all_data[sort], save_path)
            elif sort in ('HNT_JC', "HNT_ZT"):
                self.save_hnt_xls(sort, all_data[sort], save_path)
            elif sort in ('SJX_JC', "SJX_ZT"):
                self.save_sjx_xls(sort, all_data[sort], save_path)
            elif sort == 'GYC':
                self.save_gyc_xls(all_data[sort], save_path)
            elif sort in ('GHJ_JC', "GHJ_ZT"):
                self.save_ghj_xls(sort, all_data[sort], save_path)
            elif sort == 'SN':
                self.save_sn_xls(all_data[sort], save_path)
            elif sort == 'SZ':
                self.save_sz_xls(all_data[sort], save_path)
            elif sort == 'HSA':
                self.save_hsa_xls(all_data[sort], save_path)
            elif sort == 'SPB':
                self.save_spb_xls(all_data[sort], save_path)
            elif sort == 'KS':
                self.save_ks_xls(all_data[sort], save_path)
            elif sort == 'TG':
                self.save_tg_xls(all_data[sort], save_path)
            else:
                pass

    def save_fs_xls(self, data, save_path):  # 防水
        book = xl_copy(xlrd.open_workbook(save_path, formatting_info=True))
        sheet = book.get_sheet("防水")
        for _data in data:  # data相当于是防水所有的数据,循环写入
            count = self.save_data_length['FS'] + 2
            self.save_data_length['FS'] += 1
            for i in range(len(_data) - 1):
                sheet.write(count, 0, count - 1)
                sheet.write(count, i + 1, _data[i])
        book.save(save_path)

    def save_sjx_xls(self, sort, data, save_path):  # 砂浆
        book = xl_copy(xlrd.open_workbook(save_path, formatting_info=True))
        for _data in data:  # data就是砂浆的所有数据(基础/主体)
            if sort == "SJX_JC":
                sheet = book.get_sheet("砂浆试块试验报告汇总附表（基础）")
                count = self.save_data_length['SJX_JC'] + 2
                self.save_data_length['SJX_JC'] += 1
            else:
                sheet = book.get_sheet("砂浆试块试验报告汇总附表（主体）")
                count = self.save_data_length['SJX_ZT'] + 2
                self.save_data_length['SJX_ZT'] += 1
            for i in range(len(_data) - 1):
                sheet.write(count, 0, count - 1)
                sheet.write(count, i + 1, _data[i])
        book.save(save_path)

    def save_sn_xls(self, data, save_path):  # 水泥
        book = xl_copy(xlrd.open_workbook(save_path, formatting_info=True))
        sheet = book.get_sheet("水泥试验报告汇总附表")
        for _data in data:  # data就是水泥的所有数据
            count = self.save_data_length['SN'] + 2
            self.save_data_length['SN'] += 1
            for i in range(len(_data) - 1):
                sheet.write(count, 0, count - 1)
                sheet.write(count, i + 1, _data[i])
        book.save(save_path)

    def save_hsa_xls(self, data, save_path):  # 砂、石试验
        book = xl_copy(xlrd.open_workbook(save_path, formatting_info=True))
        sheet = book.get_sheet("砂、石试验报告汇总表")
        for _data in data:  # data就是砂,石试验的所有数据
            count = self.save_data_length['HSA'] + 2
            self.save_data_length['HSA'] += 1
            for i in range(len(_data) - 1):
                sheet.write(count, 0, count - 1)
                sheet.write(count, i + 1, _data[i])
        book.save(save_path)

    def save_spb_xls(self, data, save_path):  # 砼、砂浆配合比
        book = xl_copy(xlrd.open_workbook(save_path, formatting_info=True))
        sheet = book.get_sheet("砼、砂浆配合比汇总表")
        for _data in data:  # data就是砼,砂浆配合比的所有数据
            count = self.save_data_length['SPB'] + 2
            self.save_data_length['SPB'] += 1
            for i in range(len(_data) - 1):
                sheet.write(count, 0, count - 1)
                sheet.write(count, i + 1, _data[i])
        book.save(save_path)

    def save_gyc_xls(self, data, save_path):  # 钢材原材
        book = xl_copy(xlrd.open_workbook(save_path, formatting_info=True))
        sheet = book.get_sheet("钢材原材汇总表")
        for _data in data:  # data就是钢材原材的所有数据
            count = self.save_data_length['GYC'] + 2
            self.save_data_length['GYC'] += 1
            for i in range(len(_data) - 1):
                sheet.write(count, 0, count - 1)
                sheet.write(count, i + 1, _data[i])
        book.save(save_path)

    def save_ghj_xls(self, sort, data, save_path):  # 钢筋焊接试验
        book = xl_copy(xlrd.open_workbook(save_path, formatting_info=True))
        for _data in data:  # data就是钢筋焊接的所有数据(基础/主体)
            if sort == "GHJ_JC":
                sheet = book.get_sheet("钢筋焊接试验报告汇总表（基础）")
                count = self.save_data_length['GHJ_JC'] + 2
                self.save_data_length['GHJ_JC'] += 1
            else:
                sheet = book.get_sheet("钢筋焊接试验报告汇总表（主体）")
                count = self.save_data_length['GHJ_ZT'] + 2
                self.save_data_length['GHJ_ZT'] += 1
            for i in range(len(_data) - 1):
                sheet.write(count, 0, count - 1)
                sheet.write(count, i + 1, _data[i])
        book.save(save_path)

    def save_hnt_xls(self, sort, data, save_path):  # 混凝土试块
        book = xl_copy(xlrd.open_workbook(save_path, formatting_info=True))
        for _data in data:  # data就是混凝土试块的所有数据(基础/主体)
            if sort == "HNT_JC":
                sheet = book.get_sheet("混凝土试块汇总表标养（基础）")
                count = self.save_data_length['HNT_JC'] + 2
                self.save_data_length['HNT_JC'] += 1
            else:
                sheet = book.get_sheet("混凝土试块汇总表标养（主体）")
                count = self.save_data_length['HNT_ZT'] + 2
                self.save_data_length['HNT_ZT'] += 1
            for i in range(len(_data) - 1):
                sheet.write(count, 0, count - 1)
                sheet.write(count, i + 1, _data[i])
        book.save(save_path)

    def save_ks_xls(self, data, save_path):  # 混凝土抗渗试块试验
        book = xl_copy(xlrd.open_workbook(save_path, formatting_info=True))
        sheet = book.get_sheet("混凝土抗渗试块试验报告汇总表")
        for _data in data:  # data就是抗渗的所有数据
            count = self.save_data_length['KS'] + 2
            self.save_data_length['KS'] += 1
            for i in range(len(_data) - 1):
                sheet.write(count, 0, count - 1)
                sheet.write(count, i + 1, _data[i])
        book.save(save_path)

    def save_sz_xls(self, data, save_path):  # 砖试验报告
        book = xl_copy(xlrd.open_workbook(save_path, formatting_info=True))
        sheet = book.get_sheet("砖试验报告汇总附表")
        for _data in data:  # data就是砖试验的所有数据
            count = self.save_data_length['SZ'] + 2
            self.save_data_length['SZ'] += 1
            for i in range(len(_data) - 1):
                sheet.write(count, 0, count - 1)
                sheet.write(count, i + 1, _data[i])
        book.save(save_path)

    def save_tg_xls(self, data, save_path):  # 电工套管试验报告
        book = xl_copy(xlrd.open_workbook(save_path, formatting_info=True))
        sheet = book.get_sheet("电工套管汇总表")
        for _data in data:  # data就是电工套管的所有数据
            count = self.save_data_length['TG'] + 2
            self.save_data_length['TG'] += 1
            for i in range(len(_data) - 1):
                sheet.write(count, 0, count - 1)
                sheet.write(count, i + 1, _data[i])
        book.save(save_path)


if __name__ == '__main__':
    _user = 'huangxiao'
    _pass = '88888'
    project_number = 'G069338'

    cookies = MyCookie()  # 通过selenium生成cookie
    cookie = cookies.selenium_login(_user, _pass)

    param = MyParam()  # 生成请求页面的参数
    few = param.get_few_param(cookie, project_number)
    few_headers = few['headers']
    few_form_data = few['form_data']
    more = param.get_more_param(cookie, project_number)
    more_headers = more['headers']
    more_form_data = more['form_data']

    get = MyGet()  # 获取数据"所有"和"查看"的数据
    few_data = get.get_few_data(few_headers, few_form_data)
    project_name = few_data['project_name']
    few_data_list = few_data['few_data_list']
    more_data_list = get.get_more_data(more_headers, more_form_data)

    parse = MyParse()
    all_data = parse.sort_service(few_data_list, more_data_list)  # 合并数据,"所有"和"查看"

    mySave = MySave()
    save_path = r'resource/excel/《{}》[{}]汇总表.xls'.format(project_name, project_number)
    mySave.my_save(all_data, project_name, save_path)  # 保存为xls
