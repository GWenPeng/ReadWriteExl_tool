import  xlrd
from xlutils.copy import copy
def   ReadWriteExl(* position,sheet_num,filename):
        #相加的位置 position  sheet_num 第几张表
        excel=xlrd.open_workbook(filename)
        excel.sheet_names()
        sheet = excel.sheet_by_index(sheet_num)
        c=[]
        Str='{"Code":"SUNGL","Dept":"","HaveDinner":"","MarketType":""},{"Code":"SUNCT","Dept":"","HaveDinner":"","MarketType":""},{"Code":"SUNDPT","Dept":"","HaveDinner":"","MarketType":""},{"Code":"SUNPKG","Dept":"","HaveDinner":"",MarketType:""},{"Code":"GetinHouse","Dept":"","HaveDinner":"","MarketType":""},{"Code":"Nights","Dept":"","HaveDinner":"","MarketType":""}'

        # c= sheet.row_values(1)
        workbook = xlrd.open_workbook(filename) #打开原文件
        workbooknew = copy(workbook)  #copy元文件
        # 打开第二个sheet
        ws = workbooknew.get_sheet(sheet_num)
        for num in range(1,sheet.nrows):
            # sheet.nrows 为最大行数 从第二行开始读入数据并写在第E列
            c=sheet.row_values(num)
            dict = {'Code': "", 'Dept': "", 'HaveDinner': "", 'MarketType': ""}
            for po in range(0,len(position)-1):
                for key in dict:
                    dict[key] = c[po]
            ws.write(num, 4, '['+str(dict)+','+Str+']') #在第二sheet E列第二行开始写入
        workbooknew.save('HotelParameter.xls')


if __name__ == '__main__':
    ReadWriteExl(0,1,2,3,sheet_num=1,filename="HotelParameter-zh-CN(1).xlsx")