r"""


"""

#导入模块
import wx
import wx.xrc
import Zjang_XLS_arithmetic_core as zxc
import time
import json
import os
import wx.grid 
import openpyxl

# 读取配置文件
def read_config():
    with open("123XLS.json","r",encoding = "utf-8") as f:
        config = json.load(f)
    return config 

config = read_config()

VERSION = config['version']

DEBUG = bool(config['debug'])

LANGUAGE = config['language']

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
        text_pack = json.load(f)
    return text_pack

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

class Main(wx.Frame):
    def __init__(self,parent,id):
        super(Main, self).__init__(parent, id, size=(2000, 900),title='123Excel')
        panel = wx.Panel(self)
        self.text_input_path =  wx.TextCtrl(panel,pos=(10,10),size=(300,20))
        self.button_enter_path = wx.Button(panel,label=TEXT["Main_Buttons"][0],pos=(320,10),size=(80,20))
        self.button_enter_path.Bind(wx.EVT_BUTTON,self.open_file)

        self.input_path_box = wx.GenericDirCtrl(panel, wx.ID_ANY, wx.EmptyString, wx.Point( 10,40 ), wx.Size( 300,750 ), wx.DIRCTRL_3D_INTERNAL|wx.SUNKEN_BORDER, wx.EmptyString, 0)
        self.input_path_box .Bind(wx.EVT_DIRCTRL_SELECTIONCHANGED, self.get_path)
        self.subject_input = wx.TextCtrl(panel,pos=(420,10),size=(60,20))

        self.input_subject_text = wx.StaticText(panel,label=TEXT["Main_Texts"][1],pos=(480,10),size=(80,20))
        self.students_arrange = wx.TextCtrl(panel,pos=(580,10),size=(60,20))
        self.students_arrange.SetValue(SCI)

        self.input_students_arrange_text = wx.StaticText(panel,label=TEXT["Main_Texts"][4],pos=(640,10),size=(80,20))
        self.subjects_row= wx.TextCtrl(panel,pos=(740,10),size=(60,20))
        self.subjects_row.SetValue(SRI)

        self.input_subjects_row_text = wx.StaticText(panel,label=TEXT["Main_Texts"][6],pos=(800,10),size=(80,20))
        self.sheet = wx.TextCtrl(panel,pos=(900,10),size=(60,20))
        self.sheet.SetValue('sheet1')
        self.input_sheet_text = wx.StaticText(panel,label='sheet',pos=(960,10),size=(80,20))

        
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
        self.max_zdy_text = wx.StaticText(panel,label=">",pos=(910,40),size=(80,20))
        self.min_zdy = wx.TextCtrl(panel,pos=(920,40),size=(80,20))
        self.min_zdy.SetValue('100')

        self.grid = wx.grid.Grid(panel,pos=(330, 80), size=(1200, 750))

        self.return_box = wx.ListBox(panel,pos=(1580,80),size=(450,750),style=wx.LC_REPORT)

        self.only_excel_mode = wx.Button(panel,label="◈",pos=(1060,40),size=(20,20))
        self.only_excel_mode.Bind(wx.EVT_BUTTON,self.only_excel_mode_)

        self.save_grid = wx.Button(panel,label="保存",pos=(1220,40),size=(80,20))
        self.save_grid.Bind(wx.EVT_BUTTON,self.save_grid_)

        controls = [
            self.text_input_path,
            self.button_enter_path,
            self.input_path_box ,
            self.subject_input,
            self.input_subject_text,
            self.students_arrange,
            self.input_students_arrange_text,
            self.subjects_row,
            self.input_subjects_row_text,
            self.sheet,
            self.input_sheet_text,
            self.get_max_score,
            self.grid,
            self.return_box,
            self.get_min_score,
            self.get_avg,

        ]
           
    def open_file(self,event):
        global file_path
        try:
            file_path=self.text_input_path.GetValue()  
            value = zxc.read_excel(file_path,self.sheet.GetValue())
            self.grid.CreateGrid(len(value), len(value[0]))
            for i in range(len(value)):
                for j in range(len(value[0])):
                    if value[i][j] == None:
                        self.grid.SetCellValue(i, j, SNV)
                    else:
                        self.grid.SetCellValue(i, j, str(value[i][j]))
            self.return_box.Insert(f"{file_path}>>>", self.return_box.GetCount())
        except Exception as e:
            error_message(e)

    def get_path(self,event):
        try:
            input_path_box = getattr(self, 'input_path_box', None)
            if input_path_box is None:
                raise AttributeError("path_box is not initialized")

            path = input_path_box.GetPath()
            if path and self.text_input_path.GetValue() != path:
                self.text_input_path.SetValue(path)
        except Exception as e:
            error_message(e)
    
    def max_score(self,event):
        try:
            sub_max_value = self.subject_input.GetValue()
            student_lie_value = int(self.students_arrange.GetValue())
            result = zxc.get_max(zxc.read_excel(file_path,OUTPUT_PATH,f"{self.sheet.GetValue()}.xml"),
                                    sub_max_value,
                                    student_lie_value,
                                    int(self.subjects_row.GetValue()))
            for i in result:
                self.return_box.Insert(i, self.return_box.GetCount())
            
            wx.MessageBox(FINALLY_MESSAGE, '提示', wx.OK | wx.ICON_INFORMATION)
        
        except Exception as e:   
            error_message(e)

    def min_score(self,event):
        try:
            sub_min_value = self.subject_input.GetValue()
            student_lie_value = int(self.students_arrange.GetValue())
            result = zxc.get_min(zxc.read_excel(file_path,OUTPUT_PATH,f"{self.sheet.GetValue()}.xml"),
                                    sub_min_value,
                                    student_lie_value,
                                    int(self.subjects_row.GetValue()))
            for i in result:
                self.return_box.Insert(i, self.return_box.GetCount())
        except Exception as e:   
            error_message(e)
    
    def avg(self,event):
        try:
            sub_avg_value = self.subject_input.GetValue()
            result = zxc.get_avg(zxc.read_excel(file_path,OUTPUT_PATH,f"{self.sheet.GetValue()}.xml"),
                                    sub_avg_value,int(self.subjects_row.GetValue())
            )
            
            self.return_box.Insert(result, self.return_box.GetCount())
        except Exception as e:   
            error_message(e)

    def zidingyi_(self,event):
        try:
            sub_zdy_value = self.subject_input.GetValue()
            sub_zdy_max = self.max_zdy.GetValue()
            sub_zdy_min = self.min_zdy.GetValue()
            result = zxc.zidingyi(zxc.read_excel(file_path,OUTPUT_PATH,f"{self.sheet.GetValue()}.xml"),
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
    
    def only_excel_mode_(self,event):
        try:
            path = zxc.get_file_path(file_path,OUTPUT_PATH)
            value = zxc.read_excel(file_path,OUTPUT_PATH,f"{self.sheet.GetValue()}.xml")
            app = wx.App()
            frame = only_Excel_mode(None, -1, value, path)
            frame.Show()
            app.MainLoop()
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
            write_to_excel(data, file_path+'[123excel].xlsx')
            wx.MessageBox(FINALLY_MESSAGE, '提示', wx.OK | wx.ICON_INFORMATION)
            self.return_box.Insert(f"{file_path}>>>save", self.return_box.GetCount())

        except Exception as e:
            error_message(e)

    

class Options(wx.Frame):
    def __init__(self, parent, id):
        super().__init__(parent, id, f'123 Excel-{TEXT["Options_Texts"][0]}', size=(400, 600))
        
        panel = wx.Panel(self)
        self.language_setting_text = wx.StaticText(panel,pos=(10,10),  label=TEXT["Options_Texts"][3])
        self.language_setting = wx.ComboBox(panel, choices=get_languages(), pos=(10, 40), size=(150, 30))
        self.language_setting.SetValue(LANGUAGE)

        self.sub_row_init_text = wx.StaticText(panel,pos=(10,130), label='学科行初始化[subjects_row init]')
        self.sub_row_init = wx.TextCtrl(panel, pos=(10, 150), size=(150, 30))
        self.sub_row_init.SetValue(str(SRI))

        self.stu_row_init_text = wx.StaticText(panel,pos=(10,190), label='学生列初始化[students_arrange init]')
        self.stu_row_init = wx.TextCtrl(panel, pos=(10, 210), size=(150, 30))
        self.stu_row_init.SetValue(str(SCI))

        self.decimal_places_num_text = wx.StaticText(panel,pos=(10,250), label='保留小数位数[decimal_places_num]')
        self.decimal_places_num = wx.TextCtrl(panel, pos=(10, 270), size=(150, 30))
        self.decimal_places_num.SetValue(str(DECIMAL_PLACES_NUM))

        self.round_way_text = wx.StaticText(panel,pos=(10,310), label='取舍方式[round_way]')
        self.round_way = wx.ComboBox(panel, choices=['四舍五入round','向上取整round_down','向下取整round_up','银行家舍入round_half_even','五舍六入round_half_down'], pos=(10, 330), size=(150, 30))
        round_way_index = ROUND_WAYS.index(USING_ROUND_WAY)
        round_way_text = ['四舍五入round','向上取整round_down','向下取整round_up','银行家舍入round_half_even','五舍六入round_half_down']
        self.round_way.SetValue(round_way_text[round_way_index])

        self.debug_text = wx.StaticText(panel,pos=(10,370), label=TEXT["Options_Texts"][5])
        self.debug = wx.ComboBox(panel, choices=['True','False'], pos=(10, 390), size=(150, 30))
        self.debug.SetValue(str(DEBUG))

        self.save_config = wx.Button(panel,label = TEXT["Options_Texts"][1], pos=(10, 520), size=(100, 40))
        self.save_config.Bind(wx.EVT_BUTTON, self.save)
    
    def save(self,event):
        # try:
            round_way_text = ['四舍五入round','向上取整round_down','向下取整round_up','银行家舍入round_half_even','五舍六入round_half_down']
            change_config = {
                'version':VERSION,
                'output_path':OUTPUT_PATH,
                'None_value_show':SNV,
                'language':self.language_setting.GetValue(),
                'debug':self.debug.GetValue(),
                'using_round_way':ROUND_WAYS[round_way_text.index(self.round_way.GetValue())],
                'decimal_places_num':self.decimal_places_num.GetValue(),
                'init_subjects_row':self.sub_row_init.GetValue(),
                'init_students_col':self.stu_row_init.GetValue()
            }


            with open("123XLS.json","w",encoding = "utf-8") as f:
                json.dump(change_config,f,ensure_ascii=False,indent=4)
                wx.MessageBox("___⩗___", '提示', wx.OK | wx.ICON_INFORMATION)
                

        # except Exception as e:   
        #     error_message(e)






      
class only_Excel_mode(wx.Frame):
    def __init__(self, parent, id , value = [], path=" "):
        super(only_Excel_mode, self).__init__(parent, id, size=(2000, 1200),title=f'{path}>>>')
        panel = wx.Panel(self)
        self.grid = wx.grid.Grid(panel,pos=(0, 0), size=(2400, 1000))
        self.grid.CreateGrid(len(value), len(value[0]))
        for i in range(len(value)):
            for j in range(len(value[0])):
                if value[i][j] == None:
                    self.grid.SetCellValue(i, j, SNV)
                else:
                    self.grid.SetCellValue(i, j, str(value[i][j]))


class Laucher(wx.Frame):
    def __init__(self, parent, id):
        super(Laucher, self).__init__(parent, id, size=(1000, 600),title='123 Excel')
        panel = wx.Panel(self) 
        self.Lauch =wx.Button(panel,label = TEXT["Laucher_Buttons"][0], pos=(10, 360), size=(100, 40))
        self.Update_log = wx.Button(panel,label = TEXT["Laucher_Buttons"][2], pos=(10,410), size=(100, 40))
        self.Convert = wx.Button(panel,label = TEXT["Laucher_Buttons"][3], pos=(10,460), size=(100, 40))
        self.Setting = wx.Button(panel,label = TEXT["Laucher_Buttons"][1], pos=(10,510), size=(100, 40))
        self.text=wx.StaticText(panel, label=TEXT["Laucher_Texts"][0], pos=(10, 40), size=(100, 40))
        self.font = wx.Font(30, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL)
        self.text2=wx.StaticText(panel, label=TEXT["Laucher_Texts"][1], pos=(10, 80), size=(500, 40))

        self.button_about = wx.Button(panel, label=u'关于\nAbout', pos=(880, 510), size=(100, 40))
        self.qqq=wx.BitmapButton(panel, bitmap=wx.Bitmap('qqq.png'), pos=(820, 510), size=(40, 40))
        self.text.SetFont(self.font)
        self.Lauch.Bind(wx.EVT_BUTTON, self.Lauch_)
        self.Update_log.Bind(wx.EVT_BUTTON, self.Update_log_)
        self.Convert.Bind(wx.EVT_BUTTON, self.Convert_)
        self.Setting.Bind(wx.EVT_BUTTON, self.Setting_)
        self.button_about.Bind(wx.EVT_BUTTON, self.About_)
        self.qqq.Bind(wx.EVT_BUTTON, self.zanzhu_) 

        controls = [
            self.Lauch,
            self.Update_log,
            self.Convert,
            self.Setting,
            self.button_about,
            self.qqq,
        ]
    
    def Lauch_(self, event):
        app = wx.App()
        frame = Main(parent=None, id=-1)
        frame.Show()
        app.MainLoop()

    def Update_log_(self, event):
        os.system('start https://gitee.com/loser123zbx/123-xls')

    def Convert_(self, event):
        pass

    def Setting_(self, event):
        app = wx.App()
        frame = Options(parent=None, id=-1)
        frame.Show()
        app.MainLoop()

    def About_(self, event):
        os.system('start https://gitee.com/loser123zbx/123-xls')
    def zanzhu_(self, event):
                wx.MessageBox(
        """
123开发软件实属不易，
如果您觉得软件还不错，可以考虑赞助一下！
联系方式：
123的outlook邮箱：zhangbingxi123@outlook.com
123的微信：zhbx114514
123的QQ：3829187270
123的B站：https://space.bilibili.com/3537114710936221?spm_id_from=333.1007.0.0
Zjang Studio官方Q群：670186875
你可以赞助给123经费，也可以在github或者gitee上留下star
如果你大发慈悲留下经费，可以在赞助时备注一下你的名字，下次更新将会在鸣谢名单中显示
并且可以抢先体验新版本
        ""","赞助")

if __name__ == '__main__':
    app = wx.App()
    frame = Laucher(None, -1)
    frame.Show()
    app.MainLoop()