import requests
import re
import json
import xlsxwriter
import threading


class cbghandle(object):
    def __init__(self, serStart, serLenth, sheetName):
        self.headers = {
            "connection": "close",
            "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3",
            "accept-encoding": "gzip, deflate, br",
            "accept-language": "zh-CN,zh;q=0.9,en;q=0.8",
            "cache-control": "max-age=0",
            "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.169 Safari/537.36"
        }
        #self.url = "https://jianghu.cbg.163.com/cgi/api/query?view_loc=equip_list&search_type=role&order_by=&page="
        #self.url = "https://jianghu.cbg.163.com/cgi/api/query?search_type=role&order_by=&page="
        self.url = "https://jianghu.cbg.163.com/cgi/api/query?search_type=role&serverid="
        self.thisPage = 1
        self.soup = ""
        self.result = [1]
        self.is_last_page = False  # 识别是否到底
        self.OrdList = []  # 用于存储id
        self.serverid = ""  # 服务器id
        self.ordersn = ""  # 账号id
        self.area_name = ""  # "服务器": usitem["area_name"],
        self.format_equip_name = ""  # "门派": usitem["format_equip_name"],
        self.server_name = ""  # "区服": usitem["server_name"],
        self.price = ""  # "价格": usitem["price"]
        self.basic_attrs = ""  # 修为
        self.level_desc = ""  # 等级
        self.count = 0  # 计数器
        self.workbook = ""
        self.worksheet = ""
        self.row = 0
        self.s = ""
        # 判断各类物品
        self.res_feng = "无"
        self.res_chui = "无"
        self.res_man = "无"
        self.res_juan = "无"
        self.res_tao = "无"
        self.res_deng = "无"
        self.res_qian = "无"
        self.res_han = "无"
        self.res_rui = "无"
        self.res_wan = "无"
        self.res_hua = "无"
        self.res_tian = "无"
        self.res_chang = "无"
        self.res_yhuo = "无"
        self.res_bjing = "无"
        self.res_hxie = "无"
        self.res_chuchen = "否"
        self.res_yyun = "否"
        self.res_lfeng = "否"
        self.res_jhong = "否"
        self.res_zyou = "否"
        self.res_wuhua = "否"
        self.urlstr = ""
        self.serLenth = serLenth
        self.serStart = serStart
        self.sheetName = sheetName
        self.res_caicao = ""
        self.res_wakuang = ""
        self.res_famu = ""
        self.res_shenghuo = ""
        self.res_dazao = ""
        self.res_miyao = ""
        self.platform = ""  # 平台
        self.fairShow = ""  # 公示


    def sendUrl(self):
        # 发送请求
        print("正在采集第" +
              str(self.thisPage) +
              "页藏宝阁数据...服务器id为：" +
              str(self.serStart))
        response = requests.get(self.url + str(self.serStart) + "&order_by=&page=" +
                                str(self.thisPage), headers=self.headers)
        response.encoding = 'unicode-escape'
        res = json.loads(response.text)
        print(res)
        if len(res["result"]) == 0:
            self.serStart += 1
            self.thisPage = 0
        if res["paging"]["is_last_page"] == True:
            self.is_last_page = True
        self.result = res["result"]
        self.getOrdersn()

    def getOrdersn(self):
        # 获取账号id 写入OrdList
        lengths = len(self.result)
        res = self.result
        for index in range(0, lengths):
            print("正在写入藏宝阁第" + str(self.thisPage) +
                  "页第" + str(index) + "条数据...")
            self.ordersn = res[index]["game_ordersn"]
            self.serverid = res[index]["serverid"]
            self.area_name = res[index]["area_name"]
            self.format_equip_name = res[index]["format_equip_name"]
            self.server_name = res[index]["server_name"]
            self.price = res[index]["price"]
            self.basic_attrs = res[index]["other_info"]["basic_attrs"]
            self.level_desc = res[index]["level_desc"]

            self.fairShow = ""
            self.platform = ""
            if res[index]["pass_fair_show"] == 0:
                self.fairShow = "公示中"
            else:
                self.fairShow = "在售"
            if res[index]["platform_type"] == 1:
                self.platform = "iPhone"
            else:
                self.platform = "Android"
            # 获取账号详细参数
            self.getUserInfo()

    def shenghuojin(self):
        # 生活技能
        pat_caicao = re.compile(
            r'"1": {"level": (.*?), "name": "采草", "icon": "life_caiyao_icon_black2', re.S)
        self.res_caicao = re.findall(pat_caicao, self.s)
        if len(self.res_caicao) == 0:
            self.res_caicao = [0]

        pat_wakuang = re.compile(
            r'"3": {"level": (.*?), "name": "挖矿", "icon": "life_wakuang_icon_black2', re.S)
        self.res_wakuang = re.findall(pat_wakuang, self.s)

        if len(self.res_wakuang) == 0:
            self.res_wakuang = [0]

        pat_famu = re.compile(
            r'"2": {"level": (.*?), "name": "伐木", "icon": "life_fawu_icon_black2', re.S)
        self.res_famu = re.findall(pat_famu, self.s)

        if len(self.res_famu) == 0:
            self.res_famu = [0]

        pat_shenghuo = re.compile(
            r'anqi_icon_black2"}, "3": {"level": (.*?), "name": "生活装备", "icon": "life_gongju_icon_black2', re.S)
        self.res_shenghuo = re.findall(pat_shenghuo, self.s)

        if len(self.res_shenghuo) == 0:
            self.res_shenghuo = [0]

        pat_miyao = re.compile(
            r'ngju_icon_black2"}, "2": {"level": (.*?), "name": "秘药炼制", "icon": "life_duyao_icon_black2"', re.S)
        self.res_miyao = re.findall(pat_miyao, self.s)

        if len(self.res_miyao) == 0:
            self.res_miyao = [0]

        pat_dazao = re.compile(
            r'yao_icon_black2"}, "4": {"level": (.*?), "name": "打造台制作", "icon": "life_dazao_icon_black2"', re.S)
        self.res_dazao = re.findall(pat_dazao, self.s)

        if len(self.res_dazao) == 0:
            self.res_dazao = [0]

    def checkW(self):

        # 判断各类信息
        pat_zyou = re.compile(r'"name": "紫游"', re.S)
        res_zyou = re.findall(pat_zyou, self.s)
        if len(res_zyou) == 1:
            self.res_zyou = "是"
        else:
            self.res_zyou = "否"

        pat_chuchen = re.compile(r'"name": "出尘"', re.S)
        res_chuchen = re.findall(pat_chuchen, self.s)
        if len(res_chuchen) == 1:
            self.res_chuchen = "是"
        else:
            self.res_chuchen = "否"

        pat_yyun = re.compile(r'"name": "月韵"', re.S)
        res_yyun = re.findall(pat_yyun, self.s)
        if len(res_yyun) == 1:
            self.res_yyun = "是"
        else:
            self.res_yyun = "否"

        pat_lfeng = re.compile(r'"name": "流风"', re.S)
        res_lfeng = re.findall(pat_lfeng, self.s)
        if len(res_lfeng) == 1:
            self.res_lfeng = "是"
        else:
            self.res_lfeng = "否"

        pat_jhong = re.compile(r'"name": "惊鸿"', re.S)
        res_jhong = re.findall(pat_jhong, self.s)
        if len(res_jhong) == 1:
            self.res_jhong = "是"
        else:
            self.res_jhong = "否"



        pat_wuhua = re.compile(r'五花马', re.S)
        res_wuhua = re.findall(pat_wuhua, self.s)
        if len(res_wuhua) == 1:
            self.res_wuhua = "有"
        else:
            self.res_wuhua = "无"

        pat_feng = re.compile(r'风盈香', re.S)
        res_feng = re.findall(pat_feng, self.s)
        if len(res_feng) == 1:
            self.res_feng = "有"
        else:
            self.res_feng = "无"

        pat_chui = re.compile(r'垂玉', re.S)
        res_chui = re.findall(pat_chui, self.s)
        if len(res_chui) == 1:
            self.res_chui = "有"
        else:
            self.res_chui = "无"

        pat_man = re.compile(r'蔓萝纤', re.S)
        res_man = re.findall(pat_man, self.s)
        if len(res_man) == 1:
            self.res_man = "有"
        else:
            self.res_man = "无"

        pat_juan = re.compile(r'卷游尘', re.S)
        res_juan = re.findall(pat_juan, self.s)
        if len(res_juan) == 1:
            self.res_juan = "有"
        else:
            self.res_juan = "无"

        pat_tao = re.compile(r'桃花驹', re.S)
        res_tao = re.findall(pat_tao, self.s)
        if len(res_tao) == 1:
            self.res_tao = "有"
        else:
            self.res_tao = "无"

        pat_deng = re.compile(r'灯如昼', re.S)
        res_deng = re.findall(pat_deng, self.s)
        if len(res_deng) == 1:
            self.res_deng = "有"
        else:
            self.res_deng = "无"

        pat_qian = re.compile(r'流光·乾坤一掷', re.S)
        res_qian = re.findall(pat_qian, self.s)
        if len(res_qian) == 1:
            self.res_qian = "有"
        else:
            self.res_qian = "无"

        pat_han = re.compile(r'流光·寒彻', re.S)
        res_han = re.findall(pat_han, self.s)
        if len(res_han) == 1:
            self.res_han = "有"
        else:
            self.res_han = "无"

        pat_rui = re.compile(r'流光·瑞云', re.S)
        res_rui = re.findall(pat_rui, self.s)
        if len(res_rui) == 1:
            self.res_rui = "有"
        else:
            self.res_rui = "无"

        pat_wan = re.compile(r'流光·万钧', re.S)
        res_wan = re.findall(pat_wan, self.s)
        if len(res_wan) == 1:
            self.res_wan = "有"
        else:
            self.res_wan = "无"

        pat_hua = re.compile(r'流光·花楹', re.S)
        res_hua = re.findall(pat_hua, self.s)
        if len(res_hua) == 1:
            self.res_hua = "有"
        else:
            self.res_hua = "无"

        pat_tian = re.compile(r'流光·天外', re.S)
        res_tian = re.findall(pat_tian, self.s)
        if len(res_tian) == 1:
            self.res_tian = "有"
        else:
            self.res_tian = "无"

        pat_chang = re.compile(r'流光·长生', re.S)
        res_chang = re.findall(pat_chang, self.s)
        if len(res_chang) == 1:
            self.res_chang = "有"
        else:
            self.res_chang = "无"

        pat_yhuo = re.compile(r'悠游·萤火', re.S)
        res_yhuo = re.findall(pat_yhuo, self.s)
        if len(res_yhuo) == 1:
            self.res_yhuo = "有"
        else:
            self.res_yhuo = "无"

        pat_bjing = re.compile(r'悠游·冰晶', re.S)
        res_bjing = re.findall(pat_bjing, self.s)
        if len(res_bjing) == 1:
            self.res_bjing = "有"
        else:
            self.res_bjing = "无"

        pat_hxie = re.compile(r'悠游·花谢', re.S)
        res_hxie = re.findall(pat_hxie, self.s)
        if len(res_hxie) == 1:
            self.res_hxie = "有"
        else:
            self.res_hxie = "无"

    def getUserInfo(self):
        self.count += 1
        userUrl = "https://jianghu.cbg.163.com/cgi/api/get_equip_detail"
        data = {"serverid": str(self.serverid), "ordersn": str(self.ordersn)}
        response = requests.post(userUrl, data=data, headers=self.headers)
        response.encoding = 'unicode-escape'
        self.s = response.text.encode(
            'utf-8').decode('unicode_escape')  # 转换成中文
        self.checkW()
        self.row += 1
        self.worksheet.write(self.row, 0, self.area_name)  # 第4行的第1列设置值为35.5
        self.worksheet.write(
            self.row, 1, self.format_equip_name)  # 第4行的第1列设置值为35.5
        self.worksheet.write(self.row, 2, self.server_name)  # 第4行的第1列设置值为35.5
        self.worksheet.write(self.row, 3, self.level_desc)  # 第4行的第1列设置值为35.5
        self.worksheet.write(
            self.row, 4, self.basic_attrs[0][1])  # 第4行的第1列设置值为35.5
        self.worksheet.write(self.row, 5, self.price / 100)  # 第4行的第1列设置值为35.5
        # 秘籍 特技
        self.worksheet.write(self.row, 6, self.basic_attrs[1][1])  # 秘籍
        self.worksheet.write(self.row, 7, self.basic_attrs[2][1])  # 特技
        # 状态 平台
        self.worksheet.write(self.row, 8, self.fairShow)
        self.worksheet.write(self.row, 9, self.platform)

        self.worksheet.write(self.row, 27, self.res_feng)
        self.worksheet.write(self.row, 28, self.res_chui)
        self.worksheet.write(self.row, 29, self.res_man)
        self.worksheet.write(self.row, 30, self.res_juan)
        self.worksheet.write(self.row, 31, self.res_wuhua)
        self.worksheet.write(self.row, 10, self.res_tao)
        self.worksheet.write(self.row, 11, self.res_deng)
        self.worksheet.write(self.row, 12, self.res_qian)
        self.worksheet.write(self.row, 13, self.res_han)
        self.worksheet.write(self.row, 14, self.res_rui)
        self.worksheet.write(self.row, 15, self.res_wan)
        self.worksheet.write(self.row, 16, self.res_hua)
        self.worksheet.write(self.row, 17, self.res_tian)
        self.worksheet.write(self.row, 18, self.res_chang)
        self.worksheet.write(self.row, 19, self.res_yhuo)
        self.worksheet.write(self.row, 20, self.res_bjing)
        self.worksheet.write(self.row, 21, self.res_hxie)
        self.worksheet.write(self.row, 22, self.res_yyun)
        self.worksheet.write(self.row, 23, self.res_chuchen)
        self.worksheet.write(self.row, 24, self.res_lfeng)
        self.worksheet.write(self.row, 25, self.res_jhong)
        self.worksheet.write(self.row, 26, self.res_zyou)


        # 生活技能
        # self.worksheet.write(self.row, 29, self.res_caicao[0])  # 特技
        # self.worksheet.write(self.row, 30, self.res_wakuang[0])  # 特技
        # self.worksheet.write(self.row, 31, self.res_famu[0])  # 特技
        # self.worksheet.write(self.row, 32, self.res_shenghuo[0])  # 特技
        # self.worksheet.write(self.row, 33, self.res_miyao[0])  # 特技
        # self.worksheet.write(self.row, 34, self.res_dazao[0])  # 特技
        item = {}
        resch = []

    def run(self):
        self.creatSheet()
        self.worksheet.write(0, 0, "服务器")  # 第4行的第1列设置值为35.5
        self.worksheet.write(0, 1, "门派")  # 第4行的第1列设置值为35.5
        self.worksheet.write(0, 2, "区服")  # 第4行的第1列设置值为35.5
        self.worksheet.write(0, 3, "等级")  # 第4行的第1列设置值为35.5
        self.worksheet.write(0, 4, "修为")  # 第4行的第1列设置值为35.5
        self.worksheet.write(0, 5, "价格")  # 第4行的第1列设置值为35.5

        self.worksheet.write(0, 6, "金秘笈")
        self.worksheet.write(0, 7, "金紫色特技")
        self.worksheet.write(0, 8, "状态")
        self.worksheet.write(0, 9, "平台")

        self.worksheet.write(0, 10, "桃花驹")
        self.worksheet.write(0, 11, "灯如昼")
        self.worksheet.write(0, 12, "流光·乾坤一掷")
        self.worksheet.write(0, 13, "流光·寒彻")
        self.worksheet.write(0, 14, "流光·瑞云")
        self.worksheet.write(0, 15, "流光·万钧")
        self.worksheet.write(0, 16, "流光·花楹")
        self.worksheet.write(0, 17, "流光·天外")
        self.worksheet.write(0, 18, "流光·长生")
        self.worksheet.write(0, 19, "悠游·萤火")
        self.worksheet.write(0, 20, "悠游·冰晶")
        self.worksheet.write(0, 21, "悠游·花谢")
        self.worksheet.write(0, 22, "月韵")
        self.worksheet.write(0, 23, "出尘")
        self.worksheet.write(0, 24, "流风")
        self.worksheet.write(0, 25, "惊鸿")
        self.worksheet.write(0, 26, "紫游")
        self.worksheet.write(0, 27, "风盈香")
        self.worksheet.write(0, 28, "垂玉")
        self.worksheet.write(0, 29, "蔓萝纤")
        self.worksheet.write(0, 30, "卷游尘")
        self.worksheet.write(0, 31, "五花马")

        # self.worksheet.write(0, 29, "采集")
        # self.worksheet.write(0, 30, "挖矿")
        # self.worksheet.write(0, 31, "伐木")
        # self.worksheet.write(0, 32, "生活装备")
        # self.worksheet.write(0, 33, "炼药")
        # self.worksheet.write(0, 34, "打造台")
        while self.serStart < self.serLenth:
            self.sendUrl()
            self.thisPage += 1
        self.closeSheet()

    def creatSheet(self):
        self.workbook = xlsxwriter.Workbook(self.sheetName+'.xlsx')  # 创建一个excel文件
        self.worksheet = self.workbook.add_worksheet(
            u'sheet1')  # 在文件中创建一个名为TEST的sheet,不加名字默认为sheet1
    def closeSheet(self):
        self.workbook.close()


class myThread (threading.Thread):
    def __init__(self, serStart, serLenth, sheetName):
        threading.Thread.__init__(self)
        self.serStart = serStart
        self.serLenth = serLenth
        self.sheetName = sheetName
    def run(self):
        cbghandle(self.serStart, self.serLenth, self.sheetName).run()


if __name__ == "__main__":
  #  cbghandle(1, 163).run()
    thread1 = myThread(1, 7,"cbg1")
    thread2 = myThread(7, 22,"cbg2")
    thread3 = myThread(22, 40,"cbg3")
    thread4 = myThread(40, 70,"cbg4")
    thread5 = myThread(70, 120, "cbg5")
    thread6 = myThread(120, 160, "cbg6")
    thread1.start()
    thread2.start()
    thread3.start()
    thread4.start()
    thread5.start()
    thread6.start()
    thread1.join()
    thread2.join()
    thread3.join()
    thread4.join()
    thread5.join()
    thread6.join()
