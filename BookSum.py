# -*- coding:utf-8 -*-
import configparser
import os
import sys
import PySimpleGUI as sg
import pandas as pd


class BookSum:
    def __init__(self):
        # 获取文件路径
        self.type = None
        # 获取配置文件(如果有)

    def appPath(self, relativepath):
        """Returns the base application path."""
        if hasattr(sys, 'frozen'):
            basePath = os.path.dirname(sys.executable)
            # Handles PyInstaller
        else:
            basePath = os.path.dirname(__file__)
        return os.path.join(basePath, relativepath)

    def createINI(self):
        # 避免打包—复制的复杂操作，直接将新建一个ini文件再写入内容。
        content = '''[book]
                    ！# params | descript  | default 
                    ！# header | 需要获取的数据开始的行数(从0开始计数，包括标签行)   |  2
                    ！# cleanFlag | 判断数据结束所用的列标签 (根据这列数据的情况去除合计行，因为合计行没有序号)  | 序号  
                    ！# groupFlag |数据汇总分组依据变量所在列标签   |  姓名
                    ！# sheetName | 汇总表名称   | 全年汇总
                    ！#--------特别注意，cleanFlag与groupFlag要与表格保持一致，包括空格数量等（如此处的姓名中有两个空格）-----------#
                    ！header = 2
                    ！cleanFlag = 序号
                    ！groupFlag = 姓  名
                    ！sheetName = 全年汇总
                    '''
        # 为了缩进好看点，字符串有几段多了空格，此处处理下
        content = content.split('！')
        content = [x.rstrip(' ') for x in content]
        with open('config.ini', 'w+', encoding='utf-8') as file:
          file.writelines(content)
        sg.popup('ini文件创建成功', font=("微软雅黑", 12), title='提示')
        # 若重新创建ini文件，则需要重新读取ini文件
        self.config()

    def gui(self):
        '''
        简单GUI
        :return: filepath  读取到的文件地址+文件名
        '''
        sg.change_look_and_feel('Light Green 1')
        layout = [
            [sg.Text('生成ini配置文件：第一次使用或ini文件格式被破坏时使用。', size=(45, 1), font=("微软雅黑", 12)), sg.Button(
                button_text='生成ini文件', key='createINI', font=("微软雅黑", 10))],
            [sg.Text('建议使用xlxs文件格式，否则无法保存结果到同一工作簿当中',
                     size=(50, 1), font=("微软雅黑", 12))],
            [sg.Text('请选择文件所在路径', size=(15, 1), font=("微软雅黑", 12), auto_size_text=False, justification='left'),
             sg.InputText('表格路径', font=("微软雅黑", 12)), sg.FileBrowse(button_text='浏览', font=("微软雅黑", 10))],
            [sg.Submit(button_text='  提 交 ', font=("微软雅黑", 10), auto_size_button=True, pad=[5, 5]), sg.Cancel(
                button_text='  退 出  ', key='Cancel', font=("微软雅黑", 10), auto_size_button=True, pad=[5, 5])]
        ]
        window = sg.Window(
            '工作簿汇总工具', default_element_size=(40, 3)).Layout(layout)
        # TODO这里面的逻辑需要优化
        while True:
            button, values = window.Read()
            if button == 'createINI':
                self.createINI()
            elif button in (None, 'Cancel'):
                break
                return False
            else:
                filepath = values['浏览']
                type = filepath.split('.')[-1]
                if type in ('xlsx', 'xls'):
                    self.type = type
                    return filepath
                else:
                    sg.popup('所选文件非Excel工作簿类型文件，请重试',
                             font=("微软雅黑", 12), title='提示')

    def config(self):
        '''
        通过configparser模块读取相关配置文件
        '''
        try:
            config = configparser.ConfigParser()
            config.read(self.appPath('config.ini'), encoding="utf-8")
            self.header = config.getint('book', 'header')
            self.sheetName = config.get('book', 'sheetName')
            self.cleanFlag = config.get('book', 'cleanFlag')
            self.groupFlag = config.get('book', 'groupFlag')
            return True
        except configparser.NoSectionError:
            sg.popup('ini文件不存在或配置文件格式被破坏。使用前请调整或重新生成ini文件',font=("微软雅黑", 12),title='提示')
            return False

    def readfile(self,filePath):
        '''
        利用pandas读取工作表数据
        return：sheet_datas : 工作簿中所有工作表对象 list
        '''
        # 1.读取文件
        salary_book = pd.read_excel(filePath, header=self.header, sheet_name=None, na_values=[0],
                                    keep_default_na=False)
        # 2.获取sheet名合集
        sheet_names = salary_book.keys()
        # 3.判断是否已经汇总过
        if self.sheetName in sheet_names:
            sg.popup('该工作簿中已存在汇总工作表，若要重新汇总请删除该表。',
                     font=("微软雅黑", 12), title='提示')
            return False
        else:
            sheet_datas = [salary_book[x] for x in sheet_names]
            return sheet_datas

    def dataclean(self, sheet_datas):
        '''
        去除非数据行（行末注释、合计等行）
        将表内的空值、NaN值替换成0
        :param sheet_datas: 读取到的所有数据 list
        :return: sheet_datas_cleaned: 清洗过的数据集  list
        '''
        sheet_datas_cleaned = []
        cleanFlag = self.cleanFlag
        for i in range(sheet_datas.__len__()):
            sheet_data = sheet_datas[i].copy()
            # 无序号行则删除（通过标志行进行判断）
            sheet_data = sheet_data[[isinstance(
                x, int) for x in sheet_data[cleanFlag]]]
            sheet_data.replace(to_replace='', value=0, inplace=True)
            sheet_data.fillna(0, inplace=True)
            # sheet_data = sheet_data[sheet_data['序号'].isin([x for x in range(1, 50)])]
            sheet_datas_cleaned.append(sheet_data)
        return sheet_datas_cleaned

    def sumby(self, sheet_datas_cleaned):
        '''
        将各表数据汇总在一起，并通过特定列分组汇总。
        :param sheet_datas_cleaned: 清洗过的数据集  list
        :return: final_datas: 合并、分组、汇总后的数据 pd.DataFrame
        '''
        datas_concated = pd.concat([x for x in sheet_datas_cleaned])
        datas_grouped = datas_concated.groupby(
            by=self.groupFlag, as_index=True, sort=False)
        # 此处不显示所有列，但是加上columns参数就可以显示
        datas_grouped_sum = datas_grouped.sum(cloumns=None)
        # 删除不需要的‘序号’汇总列 标识列
        final_datas = datas_grouped_sum.drop(columns=[self.cleanFlag], axis=1)
        return final_datas

    def filewriter(self, final_datas,filePath):
        if self.type == 'xlsx':
            # 利用openpyxl模块的append模式添加数据到原表中(仅支持xlsx文件)
            filewriter = pd.ExcelWriter(filePath, mode='a', engine='openpyxl')
            final_datas.to_excel(
                filewriter, sheet_name=self.sheetName, encoding='utf-8')
            filewriter.save()
            filewriter.close()
        elif self.type == 'xls':
            filewriter = pd.ExcelWriter('{}.xls'.format(self.sheetName))
            final_datas.to_excel(
                filewriter, sheet_name=self.sheetName, encoding='utf-8')
            filewriter.save()
            filewriter.close()
        sg.popup("数据汇总完成，请打开表格查看{}工作表(xlsx)或程序同级目录下的{}工作簿(xls)".format(
            self.sheetName, self.sheetName), font=("微软雅黑", 12), title='提示')
        return

# 主程序入口
if __name__ == '__main__':
    sample = BookSum()
    sample.config()
    filePath = sample.gui()
    if filePath:
        data = sample.readfile(filePath)
        if data:
            data_cleaned = sample.dataclean(data)
            final_datas = sample.sumby(data_cleaned)
            sample.filewriter(final_datas,filePath)
