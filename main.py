# -*- coding:utf-8 -*-
import json
import requests
import sys
import xlwt
import time


def excel(name, content):
    workbook = xlwt.Workbook(encoding='utf-8')  # 新建工作簿
    sheet1 = workbook.add_sheet(name)  # 新建sheet
    sheet1.write(0, 0, "单号")  # 第1行第1列数据
    sheet1.write(0, 1, "快递公司")  # 第1行第2列数据
    sheet1.write(0, 2, "最新时间")
    sheet1.write(0, 3, "揽收时间")
    sheet1.write(0, 4, "签收时间")
    sheet1.write(0, 5, "最新内容")
    sheet1.write(0, 6, "状态")

    index = 0
    for i in content:
        # print('写入的是',i)
        index += 1
        sheet1.write(index, 0,i['单号'])  # 第1行第1列数据
        sheet1.write(index, 1,i['快递类型'])  # 第1行第1列数据
        sheet1.write(index, 2,i['最新时间'])  # 第1行第1列数据
        sheet1.write(index, 3,i['揽收时间'])  # 第1行第1列数据
        sheet1.write(index, 4,i['签收时间'])  # 第1行第1列数据
        sheet1.write(index, 5,i['最新内容'])  # 第1行第1列数据
        sheet1.write(index, 6,i['状态'])  # 第1行第1列数据
    workbook.save(name + '.xls')  # 保存


def search_by_kuaidi100(company, num):
    state = {
        '0': '在途',
        '1': '揽件',
        '2': '疑难',
        '3': '签收',
        '4': '退签',
        '5': '派件',
        '6': '退回',
    }

    head = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.186 Safari/537.36'
    }

    url = 'http://www.kuaidi100.com/query?type=' + company + '&postid=' + num

    jsonContext = json.loads(requests.get(url, headers=head).text)

    time = ''  # 最新时间
    context = ''  # 内容
    tracking = ''  # 状态
    carrier = ''  # 快递类型
    ls_time = ''

    try:
        time = jsonContext['data'][0]['time']  # 最新时间
        context = jsonContext['data'][0]['context']  # 内容
        tracking = state[jsonContext['state']]  # 状态
        carrier = jsonContext['com']  # 快递类型
        ls_time = jsonContext['data'][-1]['time']

    finally:
        if tracking == '签收':
            qs = time
        else:
            qs = ''
        dic = {
            '单号': num,
            '状态': tracking,
            '快递类型': carrier,
            '最新时间': time,
            '最新内容': context,
            '揽收时间': ls_time,
            '签收时间': qs
        }
        return dic


def main():
    print('请输入圆通单号,回车后按q开始查询下:\n')

    qlist = []

    while True:
        q = input()
        if q == 'q':
            break
        else:
            if q:
                qlist.append(q)
    pro = 0
    ret = []
    for i in qlist:
        pro += 1
        jd = (pro / len(qlist)) * 100
        ret.append(search_by_kuaidi100('yuantong',i))
        sys.stdout.write('\r')
        sys.stdout.write("%s%% |%s" % (int(jd), int(jd) * '#'))
        sys.stdout.flush()

    print('\n查询完成,正在写入excel...')
    name=str(int(time.time()))
    excel(name,ret)
    input('写入完成:' + name +'.xls' + ',按任意键退出')


if __name__ == '__main__':
    main()
