
#导入模块
import wx
import wx.xrc
import Zjang_XLS_arithmetic_core as zxc
import time
import json
import os
import wx.grid 
import openpyxl
import datetime
import pandas as pd


# 读取配置文件
def read_config():
    with open("123XLS.json","r",encoding = "utf-8") as f:
        return json.load(f)

config = read_config()

VERSION = config['version']

DEBUG = bool(config['debug'])

LANGUAGE = config['language']

SRA = config['show_row_add']
#show row add
SRC = config['show_col_add']
#show col add
#
# OUTPUT_DIR = config['output_dir']

SI = config['init_subject']

OUTPUT_PATH = config['output_path']

DECIMAL_PLACES_NUM = config['decimal_places_num']

USING_ROUND_WAY = config['using_round_way']

#Students Row Init
SCI = config['init_students_col']
SCI = str(SCI)

#Subject Arrange Init
SRI = config['init_subjects_row']
SRI = str(SRI)

#SNW = Show None Value
SNV = config["None_value_show"]

#NG = new grid
NGR = config['new_grid_row']
NGC = config['new_grid_col']

CORE = config['core']
CORES = config['cores']

ROUND_WAYS = [
    "ROUND_HALF_UP",
    "ROUND_UP",
    "ROUND_DOWN",
    "ROUND_HALF_DOWN",
    "ROUND_HALF_EVEN",
]
#['四舍五入round','向上取整round_down','向下取整round_up','银行家舍入round_half_even','五舍六入round_half_down'],


#验证日志文件
try:
    f=open("crash_log.txt","r",encoding="utf-8")
    f.close()
except:
    f=open("crash_log.txt","w",encoding="utf-8")
    f.close()

# 读取语言包
def get_text():
    with open("languages_packs.json", "r", encoding="utf-8") as f:
        return json.load(f)


# 获取配置中的语言对应的索引
def get_language_index(text_pack, language):
    for i, item in enumerate(text_pack):
        if item['Language'] == language:
            return i
    return None
def get_languages():
    text = get_text()
    num = len(text)
    language_list = []
    for i in range(num):
        language_list.append(text[i]['Language'])
    return language_list

config = read_config()
text_pack = get_text()
language_index = get_language_index(text_pack, config['language'])
TEXT = text_pack[language_index]
LANGUAGES = get_languages()
ERROR_MESSAGE = TEXT["Main_Messages"][0]
FINALLY_MESSAGE = TEXT["Main_Messages"][1]    
def get_time():
    timestamp = time.time()
    local_time = time.localtime(timestamp)
    time_str = "["+time.strftime('%Y-%m-%d %H:%M:%S', local_time)+"]"
    return time_str

def error_message(e,message=ERROR_MESSAGE):
    f=open("crash_log.txt","a",encoding='utf-8')
    f.write(f"{get_time()}{str(e)}\n")
    f.close()
    wx.MessageBox(message,'',wx.OK | wx.ICON_ERROR)

class Options(wx.Frame):
    def __init__(self, parent, id):
        super().__init__(parent, title='123Excel',size = (500,600))

        panel = wx.Panel(self)
        self.language_setting_text = wx.StaticText(panel, pos=(10, 10), label=TEXT["Options_Texts"][3])
        self.language_setting = wx.ComboBox(panel, choices=get_languages(), pos=(10, 40), size=(150, 30))
        self.language_setting.SetValue(LANGUAGE)

        self.sub_row_init_text = wx.StaticText(panel, pos=(10, 130), label='学科行初始化[subjects_row init]')
        self.sub_row_init = wx.TextCtrl(panel, pos=(10, 150), size=(150, 30))
        self.sub_row_init.SetValue(str(SRI))

        self.stu_row_init_text = wx.StaticText(panel, pos=(10, 190), label='学生列初始化[students_arrange init]')
        self.stu_row_init = wx.TextCtrl(panel, pos=(10, 210), size=(150, 30))
        self.stu_row_init.SetValue(str(SCI))

        self.subject_init_text = wx.StaticText(panel, pos=(10, 70), label='科目初始化[subject init]')
        self.subject_init = wx.TextCtrl(panel, pos=(10, 90), size=(150, 30))
        self.subject_init.SetValue(str(SI))

        self.decimal_places_num_text = wx.StaticText(panel, pos=(10, 250), label='保留小数位数[decimal_places_num]')
        self.decimal_places_num = wx.TextCtrl(panel, pos=(10, 270), size=(150, 30))
        self.decimal_places_num.SetValue(str(DECIMAL_PLACES_NUM))

        self.round_way_text = wx.StaticText(panel, pos=(10, 310), label='取舍方式[round_way]')
        self.round_way = wx.ComboBox(panel, choices=['四舍五入round', '向上取整round_down', '向下取整round_up',
                                                     '银行家舍入round_half_even', '五舍六入round_half_down'],
                                     pos=(10, 330), size=(150, 30))
        round_way_index = ROUND_WAYS.index(USING_ROUND_WAY)
        round_way_text = ['四舍五入round', '向上取整round_down', '向下取整round_up', '银行家舍入round_half_even',
                          '五舍六入round_half_down']
        self.round_way.SetValue(round_way_text[round_way_index])

        self.debug_text = wx.StaticText(panel, pos=(10, 370), label=TEXT["Options_Texts"][5])
        self.debug = wx.ComboBox(panel, choices=['True', 'False'], pos=(10, 390), size=(150, 30))
        self.debug.SetValue(str(DEBUG))

        self.line_ = wx.StaticLine(panel, pos=(250, 0), size=(1, 600), style=wx.LI_VERTICAL)

        self.tips = wx.StaticText(panel, pos=(260, 10), label="高级功能")

        self.new_grid_row_text = wx.StaticText(panel, pos=(260, 70), label='新建表格行数')
        self.new_grid_row = wx.TextCtrl(panel, pos=(260, 90), size=(150, 30))
        self.new_grid_row.SetValue(str(NGR))

        self.new_grid_col_text = wx.StaticText(panel, pos=(260, 130), label='新建表格列数')
        self.new_grid_col = wx.TextCtrl(panel, pos=(260, 150), size=(150, 30))
        self.new_grid_col.SetValue(str(NGC))

        self.show_none_value_text = wx.StaticText(panel, pos=(260, 190), label='显示空值')
        self.show_none_value = wx.TextCtrl(panel, pos=(260, 210), size=(150, 30))
        self.show_none_value.SetValue(str(SNV))

        self.show_grid_add_col_text = wx.StaticText(panel, pos=(260, 250), label='显示时补充列数')
        self.show_grid_add_col = wx.TextCtrl(panel, pos=(260, 270), size=(150, 30))
        self.show_grid_add_col.SetValue(str(SRC))

        self.show_grid_add_row_text = wx.StaticText(panel, pos=(260, 310), label='显示时补充行数')
        self.show_grid_add_row = wx.TextCtrl(panel, pos=(260, 330), size=(150, 30))
        self.show_grid_add_row.SetValue(str(SRA))

        self.save_config = wx.Button(panel, label=TEXT["Options_Texts"][1], pos=(10, 520), size=(100, 40))
        self.save_config.Bind(wx.EVT_BUTTON, self.save)

        self.core = wx.StaticText(panel, label="内核", pos=(260, 370))
        self.core_setting = wx.ComboBox(panel, choices=CORES, pos=(260, 390), size=(150, 30))
        self.core_setting.SetValue(CORES[CORES.index(CORE)])

    def save(self, event):
        round_way_text = ['四舍五入round', '向上取整round_down', '向下取整round_up', '银行家舍入round_half_even',
                          '五舍六入round_half_down']
        change_config = {
            'version': VERSION,
            'output_path': OUTPUT_PATH,

            'None_value_show': SNV,
            'language': self.language_setting.GetValue(),
            'debug': self.debug.GetValue(),
            'using_round_way': ROUND_WAYS[round_way_text.index(self.round_way.GetValue())],
            'decimal_places_num': self.decimal_places_num.GetValue(),
            'init_subjects_row': self.sub_row_init.GetValue(),
            'init_students_col': self.stu_row_init.GetValue()
            , 'init_subject': self.subject_init.GetValue()
            , 'new_grid_row': self.new_grid_row.GetValue()
            , 'new_grid_col': self.new_grid_col.GetValue()
            , 'show_col_add': self.show_grid_add_col.GetValue()
            , 'show_row_add': self.show_grid_add_row.GetValue()
            , 'show_none_value': self.show_none_value.GetValue()
            , 'core': self.core_setting.GetValue()
            ,'cores':['123Excel','pandas']

        }

        with open("123XLS.json", "w", encoding="utf-8") as f:
            json.dump(change_config, f, ensure_ascii=False, indent=4)
            wx.MessageBox("___⩗___", '提示', wx.OK | wx.ICON_INFORMATION)


class Main(wx.Frame):
    def __init__(self,parent,id):
        super(Main, self).__init__(parent, id, size=(1920, 1080),title='123Excel')
        panel = wx.Panel(self)
        self.text_input_path =  wx.TextCtrl(panel,pos=(0,0),size=(220,20))
        self.button_enter_path = wx.Button(panel,label=TEXT["Main_Buttons"][0],pos=(220,0),size=(80,20))
        self.button_enter_path.Bind(wx.EVT_BUTTON,self.open_file)

        self.input_path_box = wx.GenericDirCtrl(panel, wx.ID_ANY, wx.EmptyString, wx.Point( 0,20 ), wx.Size( 300,1060 ), wx.DIRCTRL_3D_INTERNAL|wx.SUNKEN_BORDER, wx.EmptyString, 0)
        self.input_path_box .Bind(wx.EVT_DIRCTRL_SELECTIONCHANGED, self.get_path)
        self.subject_input = wx.TextCtrl(panel,pos=(420,10),size=(60,20))
        self.subject_input.SetValue(SI)

        self.input_subject_text = wx.StaticText(panel,label=TEXT["Main_Texts"][1],pos=(480,10),size=(80,20))
        self.students_arrange = wx.TextCtrl(panel,pos=(580,10),size=(60,20))
        self.students_arrange.SetValue(SCI)

        self.input_students_arrange_text = wx.StaticText(panel,label=TEXT["Main_Texts"][4],pos=(640,10),size=(80,20))
        self.subjects_row= wx.TextCtrl(panel,pos=(740,10),size=(60,20))
        self.subjects_row.SetValue(SRI)

        self.input_subjects_row_text = wx.StaticText(panel,label=TEXT["Main_Texts"][6],pos=(800,10),size=(80,20))

        self.sheet = wx.ComboBox(panel,pos=(900,10),size=(60,20))
        self.input_sheet_text = wx.StaticText(panel,label='sheet',pos=(960,10),size=(80,20))
#
        
        self.get_max_score = wx.Button(panel,label=TEXT["Main_Buttons"][3],pos=(420,40),size=(80,20))
        self.get_max_score.Bind(wx.EVT_BUTTON,self.max_score)

        self.get_min_score = wx.Button(panel,label=TEXT["Main_Buttons"][5],pos=(520,40),size=(80,20))
        self.get_min_score.Bind(wx.EVT_BUTTON,self.min_score)

        self.get_avg = wx.Button(panel,label=TEXT["Main_Buttons"][4],pos=(620,40),size=(80,20))
        self.get_avg.Bind(wx.EVT_BUTTON,self.avg)

        self.zidingyi = wx.Button(panel,label=TEXT["Main_Buttons"][8],pos=(720,40),size=(80,20))
        self.zidingyi.Bind(wx.EVT_BUTTON,self.zidingyi_)
        self.max_zdy = wx.TextCtrl(panel,pos=(820,40),size=(80,20))
        self.max_zdy.SetValue('60')
        self.max_zdy_text = wx.StaticText(panel,label=">",pos=(910,40))
        self.min_zdy = wx.TextCtrl(panel,pos=(920,40),size=(80,20))
        self.min_zdy.SetValue('100')

        # self.grid = wx.grid.Grid(panel,pos=(300, 80), size=(1280, 900))

        self.return_box = wx.ListBox(panel,pos=(1580,50),size=(400,1030),style=wx.LC_REPORT)

        self.save_grid = wx.Button(panel,label="保存",pos=(1220,40),size=(80,20))
        self.save_grid.Bind(wx.EVT_BUTTON,self.save_grid_)

        self.Setting = wx.Button(panel,label="设置",pos=(1830,0),size=(80,50))
        self.Setting.Bind(wx.EVT_BUTTON,self.setting)

        self.Copy = wx.Button(panel,label="复制",pos=(1750,0),size=(80,50))
        self.Copy.Bind(wx.EVT_BUTTON,self.copy)

        self.create_grid = wx.Button(panel,label="创建新表",pos=(300,0),size=(80,20))
        self.create_grid.Bind(wx.EVT_BUTTON,self.create_grid_)



    def copy(self,event):
        selected_value = []
        for i in range(self.return_box.GetCount()):
            selected_value.append(self.return_box.GetString(i))
        selected_value = '   '.join(selected_value)
        if not selected_value:
            wx.MessageBox("啥也没有!", "Error", wx.OK | wx.ICON_ERROR)
            return
        if wx.TheClipboard.Open():
            wx.TheClipboard.SetData(wx.TextDataObject(selected_value))
            wx.TheClipboard.Close()
            wx.MessageBox("复制成功!", "提示", wx.OK | wx.ICON_INFORMATION)

    def setting(self,event):
        app = wx.App()
        frame = Options(parent=None, id=-1)
        frame.Show()
        app.MainLoop()
    def open_file(self,event):
        global file_path
        try:
            file_path=self.text_input_path.GetValue()  
            value = zxc.read_excel(excel_file=file_path,sheet=self.sheet.GetValue())
            if self.grid:
                self.grid.Destroy()
            self.grid = wx.grid.Grid(pos=(300, 80), size=(1280, 900))
            self.grid.CreateGrid(round(eval(f'{len(value)}{SRA}') ), round(eval(f'{len(value[0])}{SRC}')))
            for i in range(len(value)):
                for j in range(len(value[0])):
                    if value[i][j] == None or str(value[i][j]) == 'nan':
                        self.grid.SetCellValue(i, j, SNV)
                    else:
                        self.grid.SetCellValue(i, j, str(value[i][j]))
            self.return_box.Insert(f"{file_path}>>>", self.return_box.GetCount())
            self.Refresh()
        except Exception as e:
            error_message(e)

    def create_grid_(self,event):
        self.grid.CreateGrid(int(NGR),int(NGC))
        self.Refresh()
        self.return_box.Insert("创建成功", self.return_box.GetCount())


    def get_path(self,event):
        panel = wx.Panel(self)
        try:
            input_path_box = getattr(self, 'input_path_box', None)
            if input_path_box is None:
                raise AttributeError("path_box is not initialized")

            path = input_path_box.GetPath()

            if path and self.text_input_path.GetValue() != path:
                self.text_input_path.SetValue(path)
        except Exception as e:
            error_message(e)

        try:
            excel_file = pd.ExcelFile(path)
            sheet_names = excel_file.sheet_names
            wx.ComboBox.Set(self.sheet, sheet_names)
            self.sheet.SetSelection(0)
            file_path = self.text_input_path.GetValue()
            value = zxc.read_excel(excel_file=file_path,output_dir=OUTPUT_PATH ,sheet=self.sheet.GetValue())
            try:
                self.grid.Destroy()
            except Exception as e:
                pass
            self.grid = wx.grid.Grid(panel,pos=(300, 80), size=(1280, 900))
            self.grid.CreateGrid(round(eval(f'{len(value)}{SRA}')), round(eval(f'{len(value[0])}{SRC}')))
            for i in range(len(value)):
                for j in range(len(value[0])):
                    if value[i][j] == None or str(value[i][j]) == 'nan':
                        self.grid.SetCellValue(i, j, SNV)
                    else:
                        self.grid.SetCellValue(i, j, str(value[i][j]))
            self.return_box.Insert(f"{file_path}>>>", self.return_box.GetCount())
            self.Refresh()
        except Exception  as e:
            f = open("crash_log.txt", "a", encoding='utf-8')
            f.write(f"{get_time()}{str(e)}\n")
            f.close()



    def max_score(self,event):
        try:
            self.return_box.Clear()
            file_path = self.text_input_path.GetValue()
            sub_max_value = self.subject_input.GetValue()
            student_lie_value = int(self.students_arrange.GetValue())
            result = zxc.get_max(zxc.read_excel(file_path,OUTPUT_PATH,f"{self.sheet.GetValue()}"),
                                    sub_max_value,
                                    student_lie_value,
                                    int(self.subjects_row.GetValue()))
            for i in result:
                self.return_box.Insert(i, self.return_box.GetCount())
        
        except Exception as e:   
            error_message(e)

    def min_score(self,event):
        try:
            self.return_box.Clear()
            file_path = self.text_input_path.GetValue()
            sub_min_value = self.subject_input.GetValue()
            student_lie_value = int(self.students_arrange.GetValue())
            result = zxc.get_min(zxc.read_excel(file_path,OUTPUT_PATH,self.sheet.GetValue()),
                                    sub_min_value,
                                    student_lie_value,
                                    int(self.subjects_row.GetValue()))
            for i in result:
                self.return_box.Insert(i, self.return_box.GetCount())
        except Exception as e:   
            error_message(e)
    
    def avg(self,event):
        try:
            self.return_box.Clear()
            file_path = self.text_input_path.GetValue()
            sub_avg_value = self.subject_input.GetValue()
            result = zxc.get_avg(zxc.read_excel(file_path,OUTPUT_PATH,self.sheet.GetValue()),
                                    sub_avg_value,int(self.subjects_row.GetValue())
            )
            
            self.return_box.Insert(result, self.return_box.GetCount())
        except Exception as e:   
            error_message(e)

    def zidingyi_(self,event):
        try:
            self.return_box.Clear()
            file_path = self.text_input_path.GetValue()
            sub_zdy_value = self.subject_input.GetValue()
            sub_zdy_max = self.max_zdy.GetValue()
            sub_zdy_min = self.min_zdy.GetValue()
            result = zxc.zidingyi(zxc.read_excel(file_path,OUTPUT_PATH,self.sheet.GetValue()),
                                    sub_zdy_value,
                                    int(self.students_arrange.GetValue()),
                                    sub_zdy_max,
                                    sub_zdy_min,
                                    int(self.subjects_row.GetValue())
            )
            for i in result:
                self.return_box.Insert(i, self.return_box.GetCount())
        except Exception as e:   
            error_message(e)

    def save_grid_(self,event):
        try:
            def write_to_excel(data, filename):
                wb = openpyxl.Workbook()
                ws = wb.active
                for row in data:
                    ws.append(row)
                wb.save(filename)

            data = []
            for i in range(self.grid.GetNumberRows()):
                row = []
                for j in range(self.grid.GetNumberCols()):
                    cell_value = self.grid.GetCellValue(i, j)
                    row.append(cell_value)
                data.append(row)

            # 获取当前用户的下载文件夹路径
            downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
            # 拼接完整的文件路径
            file_path = os.path.join(downloads_path, f'123excel-[{datetime.datetime.now().strftime("%Y%m%d%H%M%S")}].xlsx')

            write_to_excel(data, file_path)
            wx.MessageBox(FINALLY_MESSAGE, '提示', wx.OK | wx.ICON_INFORMATION)
            self.return_box.Insert(f"{file_path}>>>save \n 保存至当前用户的下载文件夹", self.return_box.GetCount())

        except Exception as e:
            error_message(e)
        # try:
        #     def write_to_excel(data, filename):
        #         wb = openpyxl.Workbook()
        #         ws = wb.active
        #         for row in data:
        #             ws.append(row)
        #         wb.save(filename)
        #
        #     data = []
        #     for i in range(self.grid.GetNumberRows()):
        #         row = []
        #         for j in range(self.grid.GetNumberCols()):
        #             cell_value = self.grid.GetCellValue(i, j)
        #             row.append(cell_value)
        #         data.append(row)
        #     file_path = OUTPUT_DIR + '/' +f'{}.xlsx'
        #     write_to_excel(data, file_path)
        #     wx.MessageBox(FINALLY_MESSAGE, '提示', wx.OK | wx.ICON_INFORMATION)
        #     self.return_box.Insert(f"{file_path}>>>save", self.return_box.GetCount())
        #
        # except Exception as e:
        #     error_message(e)


#         """
# 123开发软件实属不易，
# 如果您觉得软件还不错，可以考虑赞助一下！
# 联系方式：
# 123的outlook邮箱：zhangbingxi123@outlook.com
# 123的微信：zhbx114514
# 123的QQ：3829187270
# 123的B站：https://space.bilibili.com/3537114710936221?spm_id_from=333.1007.0.0
# Zjang Studio官方Q群：670186875
# 你可以赞助给123经费，也可以在github或者gitee上留下star
# 如果你大发慈悲留下经费，可以在赞助时备注一下你的名字，下次更新将会在鸣谢名单中显示
# 并且可以抢先体验新版本
#         ""","赞助")

if __name__ == '__main__':
    app = wx.App()
    frame = Main(None, -1)
    frame.Show()
    app.MainLoop()