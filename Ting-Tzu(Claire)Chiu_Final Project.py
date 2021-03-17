
import re, wx, sqlite3, datetime, xlwt



def CreateDataBase():
    # Create a connection to sqlite
    konn = sqlite3.connect('VisitorLog.db')

    # Open a cursor
    a = konn.cursor()

    # Execute any SQL statement
    # Create table named "FirstSignUp" if it does not already exist.
    a.execute("create table if not exists SignUp (First, Last, email, Phone, Host, Date)")
    

    # Add entries
    konn.commit()









class VisitorSys(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, parent=None, title="Visitor System", size=(600,300))
        panel = wx.Panel(self)
        self.label = wx.StaticText(panel,label = "Check in Status : " , pos=(270,10), size=(100,50), style= wx.ALIGN_LEFT)
        self.label.SetFont(wx.Font(20, wx.FONTFAMILY_TELETYPE, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL))
        
        wx.StaticText(parent=panel, label="First Name",  size = (100,20),  pos=(10,10))
        wx.StaticText(parent=panel, label="Last Name" ,  size = (100,20),  pos=(10,40))
        wx.StaticText(parent=panel, label="Email"     ,  size = (100,20),  pos=(10,70))
        wx.StaticText(parent=panel, label="Phone"     ,  size = (100,20),  pos=(10,100))
        wx.StaticText(parent=panel, label="Host"      ,  size = (100,20),  pos=(10,130))
        
        self.FirstName = wx.TextCtrl(parent=panel, size = (100,25), pos=(120,10))
        self.LastName  = wx.TextCtrl(parent=panel, size = (100,25), pos=(120,40))
        self.Email     = wx.TextCtrl(parent=panel, size = (100,25), pos=(120,70))
        self.Phone     = wx.TextCtrl(parent=panel, size = (100,25), pos=(120,100))

        employe        = ['Andy Pai', 'Claire Chiu', 'Robert Chang', 'Mandy Lu', 'Charlie Hwang']
        self.combo     = wx.ComboBox(parent=panel ,value = "Andy Pai", size = (100,30), choices = employe,  pos=(120,130))

        
        self.submit_btn      = wx.Button(parent=panel, label="Submit", pos=(20,200))
        self.download_btn    = wx.Button(parent=panel, label="Download List", pos=(140,200))
        self.read_btn        = wx.Button(parent=panel, label="Read Log file", pos=(260,200))
        
        self.message         = wx.StaticText(parent=panel,pos=(270,50), size = (100,20), style= wx.ALIGN_LEFT)

	
        self.Bind(wx.EVT_BUTTON, self.Submit,       self.submit_btn)
        self.Bind(wx.EVT_BUTTON, self.DownLoad,     self.download_btn)
        self.Bind(wx.EVT_BUTTON, self.ReadDataBase, self.read_btn)



    def ReadDataBase(self, event):
        # Create a connection to sqlite
        konn = sqlite3.connect('VisitorLog.db')

        # Open a cursor
        a = konn.cursor()

        alllines = a.execute("select * from SignUp")
#        print(alllines)
        l_xl = []
        for i in alllines:
            l_xl.append(i)

        sort_l = sorted(l_xl, key = lambda x : x[5], reverse = True)

        # Print it
        for i in sort_l:
            print(i)
        a.close()


    def DownLoad(self, event):
        konn = sqlite3.connect('VisitorLog.db')
        a = konn.cursor()
        alllines = a.execute("select * from SignUp")
        l_xl = []
        for i in alllines:
            l_xl.append(i)

        workbook  = xlwt.Workbook()
        sheet = workbook.add_sheet("Visitor Log")


        header_l = ['First Nmae', 'Last Name', 'email address', 'Phone', 'Host', 'Time']
        for i,v in enumerate(header_l):         # write header for excel in row 0
            sheet.write(0, i, v)

        
        for i in range(len(l_xl)):                  #row
            for j in range(len(l_xl[i])) :          #col          
                sheet.write(i+1, j, l_xl[i][j])

        workbook.save('Visitor_Log.xls')
#        print(l_xl)


        

    def Submit(self, event):
        if self.__Info_Check():
            l_obj   = [self.FirstName, self.LastName, self.Email, self.Phone, self.combo]
            l_txt   = []

            for i in l_obj:
                l_txt.append(i.GetValue())
#            print(l_txt)
            self.__WriteIntoDB(l_txt)
            self.message.SetLabel("Welcome")
            font = wx.Font(20, wx.FONTFAMILY_ROMAN, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)
            self.message.SetForegroundColour((0,255,0))
            
        else:
#            print("not correct")
            self.message.SetLabel("Please input Valid Info")
            font = wx.Font(20, wx.FONTFAMILY_ROMAN, wx.FONTSTYLE_ITALIC, wx.FONTWEIGHT_BOLD)
            self.message.SetForegroundColour((255,0,0))
            
        self.message.SetFont(font)



    def __Name_Check(self, name):
        str_name = name.GetValue()
        count     = 0
        for i in str_name:
            if i.isdigit():
                count +=1
        if count > 0:
            return False
        else:
            return True


    def __Email_Check(self):
        email_name = self.Email.GetValue()
        email_check = re.compile(r'(\w+@\w+\.(com|net|org|edu))')
        
        if email_check.search(email_name):
            return True
        else:
            return False
        
            
    def __Phone_Check(self):
        phone = self.Phone.GetValue()
        count     = 0
        for i in phone:
            if i.isdigit():
                count +=1
        if count ==  len(phone) and count > 0:
            return True
        else:
            return False


    def __Info_Check(self):
        if self.__Name_Check(self.FirstName) and self.__Name_Check(self.LastName) and self.__Email_Check() and self.__Phone_Check():
            return True
        else:
            return False


    def __WriteIntoDB(self, l_visit):
        konn = sqlite3.connect('VisitorLog.db')
        # Open a cursor
        a = konn.cursor()

        # Add entries
        a.execute("insert into SignUp (First, Last, email, Phone, Host, Date) values (?, ?, ?, ?, ?, ?)", \
                  (l_visit[0], l_visit[1], l_visit[2], l_visit[3], l_visit[4], datetime.datetime.today() ))
        konn.commit()
                    





if __name__ == '__main__':

    CreateDataBase()
    app = wx.App()
    frame = VisitorSys()
    frame.Show()
    app.MainLoop()















