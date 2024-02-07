import openpyxl
import os
import json
class Procrsser:
    def __init__(self,path,savepath,workbookpath):
        self.path=path
        self.savepath=savepath
        self.workbook =openpyxl.load_workbook(workbookpath)
        self.sheet = self.workbook['标注']
    def filter(self,string):
        res=''
        res2=''
        res3=''
        for i in string:
            if((i<='\u9fff' and i>='u4e00') or i==','):
                res+=i
            elif(i.isalpha()):
                res2+=i
            elif(i.isdigit()):
                res3+=i
        
        return [res,res2,res3]
    def process(self):
        files = [os.path.join(self.path, file) for file in os.listdir(self.path)]
        for i in files:
            num=self.filter(i)[2]
            self.sheet['A3']=num
            with open(i+'\\文章.txt','r',encoding='utf-8') as f:
                title=f.readline().strip()
                if(len(title)>15):
                    title=''
                    print('*这篇文章没有标题')
                print(title)
                self.sheet['B3']=title
            with open(i+'\\选项.txt','r',encoding='utf-8') as f:
                question=f.readline().strip()
                self.sheet['C3']=self.filter(question)[0]
                self.sheet['D3']=self.filter(question)[1]
                optionA=f.readline().strip()
                self.sheet['E3']=self.filter(optionA)[1]
                self.sheet['F3']=self.filter(optionA)[0]
                optionB=f.readline().strip()
                self.sheet['E4']=self.filter(optionB)[1]
                self.sheet['F4']=self.filter(optionB)[0]
                optionC=f.readline().strip()
                self.sheet['E5']=self.filter(optionC)[1]
                self.sheet['F5']=self.filter(optionC)[0]
                optionD=f.readline().strip()
                self.sheet['E6']=self.filter(optionD)[1]
                self.sheet['F6']=self.filter(optionD)[0]
            self.workbook.save(self.savepath+'\\数据'+num+'.xlsx')

if __name__ == '__main__':
    with open('config.json','r',encoding='utf-8') as file:
        file_content = file.read()
        data=json.loads(file_content)
    p=Procrsser(data['path'],data['savepath'],data['workbookpath'])
    p.process()