# -*- coding: utf-8 -*-
import tkinter as tk
import requests
import xlwt
import json
import tkinter.messagebox as messagebox
from fake_useragent import UserAgent
import jieba
import wordcloud
import tempfile
import re
import matplotlib.pyplot as plt
import time
import numpy as np
import xlrd
import os
import threading
current_file_path = os.path.abspath(__file__)

stop_event = threading.Event()


def fetch_comments_in_thread():
    ua = UserAgent()
    headers={'User-Agent':ua.random}
    try:
        r=requests.get(f"https://c.y.qq.com/splcloud/fcgi-bin/smartbox_new.fcg?key={MAINTK.URL.get()}",headers=headers)
    except:
        print("请求失败，请检查网络")
        for btn_i in [MAINTK.btn_save,MAINTK.btn_read,MAINTK.btn_cloud,MAINTK.btn_n_t_bar,MAINTK.btn_n_t_bar2,MAINTK.btn_best10]:
            btn_i.config(state=tk.NORMAL)
        stop_event.clear()
        return 0
    serach=json.loads(r.text)
    it_list=serach.get('data').get('song').get('itemlist',[])
    if len(it_list)>0:
        topid=it_list[0].get('id')
        print(f"找到歌手[{it_list[0].get('singer')}]的歌曲[{it_list[0].get('name')}]")
    else:
        print("找不到歌曲")
        return 0
    _=''
    common_comments=[]
    url=f"https://c.y.qq.com/base/fcgi-bin/fcg_global_comment_h5.fcg?biztype=1&cmd=8&topid={topid}&pagenum={0}&pagesize=25"
    ua = UserAgent()
    headers={'User-Agent':ua.random}
    try:
        r=requests.get(url,headers=headers)
    except:
        print("请求失败，请检查网络")
        for btn_i in [MAINTK.btn_save,MAINTK.btn_read,MAINTK.btn_cloud,MAINTK.btn_n_t_bar,MAINTK.btn_n_t_bar2,MAINTK.btn_best10]:
            btn_i.config(state=tk.NORMAL)
        stop_event.clear()
        return 0
    
    data=json.loads(r.text)
    DATA.commenttotal=int(data.get('comment').get('commenttotal'))
    print(f"共{DATA.commenttotal}条评论")
    pages=int(MAINTK.GET_number.get())//25 if MAINTK.GET_number.get()!='all' else DATA.commenttotal//25+1
    for page in range(0,pages):
        url=f"https://c.y.qq.com/base/fcgi-bin/fcg_global_comment_h5.fcg?biztype=1&cmd=8&topid={topid}&pagenum={str(page)}&pagesize=25"
        ua = UserAgent()
        headers={'User-Agent':ua.random}
        try:
            r=requests.get(url,headers=headers)
        except:
            print("网络波动，中断请求")
            for btn_i in [MAINTK.btn_save,MAINTK.btn_read,MAINTK.btn_cloud,MAINTK.btn_n_t_bar,MAINTK.btn_n_t_bar2,MAINTK.btn_best10]:
                btn_i.config(state=tk.NORMAL)
            stop_event.clear()
            break
        data=json.loads(r.text)
        DATA.commenttotal=int(data.get('comment').get('commenttotal'))
        __d=data.get('comment').get('commentlist')if data.get('comment').get('commentlist')!=None else []
        common_comments+=__d
        for i in __d:
            _+=('"'+i.get('rootcommentcontent','_').replace('"',"'")+'" ')
        if len(__d)==0:
            break
        print(f'已获取{len(common_comments)}条评论')
        DATA.common_comments=np.array(common_comments,dtype=dict)
        if stop_event.is_set():break
    MAINTK.comment_list.set(_)
    if stop_event.is_set():
        print(f'获取中断，共获取{len(common_comments)}条评论')
    else:
        print(f'获取完成，共获取{len(common_comments)}条评论')
    MAINTK.btn_get.config(text="获取评论(输入歌曲名)", command=btn.get)
    for btn_i in [MAINTK.btn_save,MAINTK.btn_read,MAINTK.btn_cloud,MAINTK.btn_n_t_bar,MAINTK.btn_n_t_bar2,MAINTK.btn_best10]:
        btn_i.config(state=tk.NORMAL)
    stop_event.clear()


class __btn():
    def get(self):
        thread = threading.Thread(target=fetch_comments_in_thread)
        thread.start()
        MAINTK.btn_get.config(text="中断", command=self.stop_get)
        for btn in [MAINTK.btn_save,MAINTK.btn_read,MAINTK.btn_cloud,MAINTK.btn_n_t_bar,MAINTK.btn_n_t_bar2,MAINTK.btn_best10]:
            btn.config(state=tk.DISABLED)
    def stop_get(self):
        stop_event.set()
    def save(self):
        workbook=xlwt.Workbook(encoding='utf-8')
        worksheet=workbook.add_sheet('评论')
        worksheet.write(0,0,'昵称')
        worksheet.write(0,1,'时间')
        worksheet.write(0,2,'点赞数')
        worksheet.write(0,3,'评论')
        n=1
        for i in range(0,len(DATA.common_comments)):
            worksheet.write(n,0,DATA.common_comments[i].get('nick','NULL'))
            worksheet.write(n,1,time.strftime('%Y年%m月%d日%H:%M:%S',time.localtime(int(DATA.common_comments[i].get('time',0)))))
            worksheet.write(n,2,DATA.common_comments[i].get('praisenum','NULL'))
            worksheet.write(n,3,DATA.common_comments[i].get('rootcommentcontent','NULL'))
            n+=1
        workbook.save(os.path.join(current_file_path,"../",MAINTK.FILE_NAME.get()+'.xls'))
        print(MAINTK.FILE_NAME.get()+'.xls'+"保存成功")
    def read(self):
        workbook:xlrd.book.Book=xlrd.open_workbook(os.path.join(current_file_path,"../",MAINTK.FILE_NAME.get()+'.xls'))
        sheet:xlrd.book.sheet.Sheet=workbook.sheets()[0]
        DATA.common_comments=[]
        for k in range(1,sheet.nrows):
            i=sheet.row_values(k)
            DATA.common_comments.append({'nick':i[0],"time":time.mktime(time.strptime(i[1],'%Y年%m月%d日%H:%M:%S')),'praisenum':i[2],'rootcommentcontent':i[3]})
        DATA.commenttotal=len(DATA.common_comments)
        _=''
        for i in DATA.common_comments:
            _+=('"'+i.get('rootcommentcontent','_').replace('"',"'")+'" ')
        DATA.common_comments=np.array(DATA.common_comments,dtype=dict)
        MAINTK.comment_list.set(_)
        print(f'共加载{len(DATA.common_comments)}条评论')

    def cloud(self):
        with tempfile.TemporaryFile(mode='wb+') as file:
            file.name='_.jpeg'
            
            wordcloud.wordcloud.WordCloud(width=1000,font_path="msyh.ttc",height=700).generate(
                " ".join((
                    i for i in jieba.lcut(('_'.join((re.sub('([em].*[/em])?','',i.get('rootcommentcontent','_')) for i in DATA.common_comments))).replace(" ","_"))if len(i)>=2)
                )).to_file(file)
            img = plt.imread(file)
            plt.imshow(img)
            plt.show()
    def n_t_bar(self):
        plt.rcParams['font.family'] = ['sans-serif']
        plt.rcParams['font.sans-serif'] = ['SimHei']
        fig=plt.figure()
        
        time_max=time.localtime(int(DATA.common_comments[0].get('time',0)))
        time_min=time.localtime(int(DATA.common_comments[-1].get('time',0)))
        dic={}
        for i in DATA.common_comments:
            mon_lo=time.localtime(int(i.get('time',0))).tm_year*12+time.localtime(int(i.get('time',0))).tm_mon-1
            if mon_lo not in dic.keys():
                dic.update({mon_lo:[i]})
            else:
                dic[mon_lo].append(i)
        x=[f"{ym//12}年{ym%12+1}月" for ym in dic.keys()]
        y=[len(l) for l in dic.values()]
        plt.title('按月的评论数分析')
        plt.xlabel('月次')
        plt.ylabel('评论数')
        plt.bar(x,y)
        plt.show()
    def n_t_bar2(self):
        plt.rcParams['font.family'] = ['sans-serif']
        plt.rcParams['font.sans-serif'] = ['SimHei']
        fig=plt.figure()
        lct=int(time.time())//86400*86400
        dic={i:[] for i in range(lct-86400*29,lct+1,86400)}
        for i in DATA.common_comments:
            day_lo=int(i.get('time',0))//86400*86400
            if day_lo not in dic.keys():
                ...
            else:
                dic[day_lo].append(i)
        x=[time.strftime('%d',time.localtime(da)) for da in dic.keys()]
        y=[len(l) for l in dic.values()]
        plt.title('一个月内的评论数分析')
        plt.xlabel('日')
        plt.ylabel('评论数')
        plt.bar(x,y)
        plt.show()
    def best10(self):
        b10=sorted(DATA.common_comments,key=lambda i:int(i.get('praisenum','0')),reverse=True)[:10]
        _info=""
        for i in b10:
            _info+=f"{i.get('praisenum')}赞\t{time.strftime('%Y年%m月%d日%H:%M:%S',time.localtime(int(i.get('time',0))))}\t{i.get('nick')}:\n\t{i.get('rootcommentcontent')}\n\n"
        messagebox.showinfo('评论top10',_info)
    def serach(self):
        ...
        ua = UserAgent()
        headers={'User-Agent':ua.random}
        topid=requests.get("https://c.y.qq.com/splcloud/fcgi-bin/smartbox_new.fcg?key={}",headers=headers)
        
class __DATA():
    def __init__(self):
        self.comment_comments:np.array=np.array(())
        self.word=[]
        self.word_counter={}
        self.commenttotal=0

class tk_SF():
    def __init__(self):
        self.root= tk.Tk()
        self.root.title('QQ音乐评论爬取和分析')
        self.root.resizable(width=None, height=None)
        self.URL=tk.StringVar()

        self.comment_list=tk.StringVar()
        tk.Listbox(self.root,listvariable=self.comment_list,width=50,height=19).grid(column=0,row=0,columnspan=6,rowspan=11)


        tk.Entry(self.root,textvariable=self.URL,width=30).grid(column=6,row=0)
        self.btn_get=tk.Button(self.root,text='获取评论(输入歌曲名)',command=btn.get)
        self.btn_get.grid(column=6,row=1)
        tk.Label(self.root,text='获取的评论数("all"表示全部)',width=20).grid(column=6,row=2)
        self.GET_number=tk.Spinbox(self.root,values=['all']+[i for i in range(25,1001,25)],width=20)
        self.GET_number.grid(column=6,row=3)
        
        self.btn_save=tk.Button(self.root,text='将评论保存为xls文件',command=btn.save)
        self.btn_save.grid(column=6,row=4)
        self.btn_read=tk.Button(self.root,text='从xls文件中加载',command=btn.read)
        self.btn_read.grid(column=6,row=5)
        tk.Label(self.root,text='操作的文件名',width=20).grid(column=6,row=6)
        self.FILE_NAME=tk.StringVar()
        tk.Entry(self.root,textvariable=self.FILE_NAME,width=20).grid(column=6,row=7)

        self.btn_cloud=tk.Button(self.root,text='词云图显示',command=btn.cloud)
        self.btn_cloud.grid(column=6,row=8)

        self.btn_n_t_bar=tk.Button(self.root,text='按月的评论数分析',command=btn.n_t_bar)
        self.btn_n_t_bar.grid(column=6,row=9)
        self.btn_n_t_bar2=tk.Button(self.root,text='一个月内的评论数分析',command=btn.n_t_bar2)
        self.btn_n_t_bar2.grid(column=6,row=10)
        self.btn_best10=tk.Button(self.root,text='评论点赞top10',command=btn.best10)
        self.btn_best10.grid(column=6,row=11)
        self.init()

    def init(self):
        self.URL.set('晴天')
        self.FILE_NAME.set('comments')
    def mainloop(self):
        self.root.mainloop()

if __name__=="__main__":
    btn=__btn()
    DATA=__DATA()
    MAINTK=tk_SF()
    MAINTK.mainloop()
