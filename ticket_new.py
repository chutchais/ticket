import win32gui,win32con,win32api,win32ui
# import re
import pyautogui
import re, traceback
import time
import sys
import tempfile
import argparse



class WindowMgr:
    """Encapsulates some calls to the winapi for window management"""
    settingFile =''

    def __init__ (self):
        """Constructor"""
        self.hwnd = None

    def find_window(self,title):
        try:
            self.hwnd = win32gui.FindWindow(None, title)
            assert self.hwnd
            return self.hwnd
        except:
            pyautogui.alert(text='Not found program name ' + title + '\n' 
                            'Please open program before excute script', title='Unable to open program', button='OK')
            print ('Not found program')
            return None


    def set_onTop(self,hwnd):
        win32gui.SetForegroundWindow(hwnd)
        return win32gui.GetWindowRect(hwnd)



    def Maximize(self,hwnd):
        win32gui.ShowWindow(hwnd,win32con.SW_RESTORE)#, win32con.SW_MAXIMIZE

    def get_mouseXY(self):
        return win32gui.GetCursorPos()

    def set_mouseXY(self):
        import os.path
        import json
        x,y,w,h = win32gui.GetWindowRect(self.hwnd)
        print ('Current Window X : %s  Y: %s' %(x,y))
        fname = 'setting.json'
        if os.path.isfile(fname) :
            dict = eval(open(fname).read())
            x1 = dict['x']
            y1 = dict['y']
            print ('Setting X : %s  Y: %s' %(x1,y1))
        win32api.SetCursorPos((x+x1,y+y1))
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x+x1, y+y1, 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x+x1, y+y1, 0, 0)
        print ('Current Mouse X %s' % self.get_mouseXY()[0])
        print ('Current Mouse Y %s' % self.get_mouseXY()[1])


    def saveFirstDataPos(self):
        x,y,w,h = win32gui.GetWindowRect(self.hwnd)
        print ('Window X : %s  Y: %s' %(x,y))
        x1,y1 = self.get_mouseXY()
        print ('Mouse X : %s  Y: %s' %(x1,y1))
        data={}
        data['x'] = x1-x
        data['y'] = y1-y
        # f = open("setting.json", "w")
        # self.settingFile
        f = open(settingFile, "w")
        f.write(str(data))

        f.close()

    def wait(self,seconds=1,message=None):
        """pause Windows for ? seconds and print
an optional message """
        win32api.Sleep(seconds*1000)
        if message is not None:
            win32api.keybd_event(message, 0,0,0)
            time.sleep(.05)
            win32api.keybd_event(message,0 ,win32con.KEYEVENTF_KEYUP ,0)

    def typer(self,stringIn=None):
        PyCWnd = win32ui.CreateWindowFromHandle(self.hwnd)
        for s in stringIn :
            if s == "\n":
                self.hwnd.SendMessage(win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
                self.hwnd.SendMessage(win32con.WM_KEYUP, win32con.VK_RETURN, 0)
            else:
                print ('Ord %s' % ord(s))
                PyCWnd.SendMessage(win32con.WM_CHAR, ord(s), 0)
        PyCWnd.UpdateWindow()


#     import win32ui
# import time

    def WindowExists(windowname):
        try:
            win32ui.FindWindow(None, windowname)

        except win32ui.error:
            return False
        else:
            return True

def get_data(ticket):
    import urllib3
    http = urllib3.PoolManager()
    url = 'http://svr-lcb1app:8080/e-Ticket/GETdata.php?barcode=' + ticket
    r = http.request('GET', url)
    print(r.data)
    if r.status == 200:
        str = r.data.decode("utf-8")
        data={}
        print(str)
        print(len(str))
        if len(str)>0 :
            tmp = str.split('|')
            
            data['barcode']= tmp[0]
            data['container']= tmp[1]
            data['bl']= tmp[2]
            data['status']= r.status
            data['description']='OK'
            data['url'] = url
        else:
            data['status']= r.status
            data['description']= 'Not found barcode :' + ticket
            data['url'] = url
    else :
        data={}
        data['status']= r.status
        data['description'] = 'Unable to access Ticket web server'
        data['url'] = url

    print (data)
    return data

def get_data2():
    import urllib3
    http = urllib3.PoolManager()
    url = 'http://svr-lcb1app:8080/e-Ticket/GETdata2.php'
    r = http.request('GET', url)
    print('Data returned %s' % r.data)

    if r.status == 200:
        str = r.data.decode("utf-8")
        data={}
        print(str)
        print(len(str))
        if len(str)>0 :
            tmp = str.split('|')
            
            data['barcode']= tmp[0]
            data['container']= tmp[1]
            data['bl']= tmp[2]
            data['status']= r.status
            data['description']='OK'
            data['url'] = url
        else:
            data['status']= 0
            data['barcode']= ''
            data['description']= 'No Barcode details' 
            data['url'] = url
    else :
        data={}
        data['status']= 9999
        data['description'] = 'Unable to access Ticket web server'
        data['url'] = url

    # print (data)
    return data

def fill_data(hwnd,ticket_dict):
    # print (ticket_dict['barcode'])
    secs_between_keys=0.05
    if ticket_dict['description'].strip()=='OK':
        hwnd.wait(0,0x09)
        hwnd.wait(0,0x09)
        hwnd.wait(0,0x09)
        pyautogui.typewrite('3', interval=secs_between_keys) #Full Out
        #hwnd.wait(0,0x09)
        if ticket_dict['container'].strip() !='':
            pyautogui.typewrite(ticket_dict['container'].strip(), interval=secs_between_keys)
        else:
            hwnd.wait(0,0x09)
            hwnd.wait(0,0x09)
        hwnd.wait(0,0x09)
        pyautogui.typewrite('M', interval=secs_between_keys)
        hwnd.wait(0,0x09)
        pyautogui.typewrite(ticket_dict['bl'].strip(), interval=secs_between_keys)
        hwnd.wait(0,0x09)
        pyautogui.typewrite('LOCAL' ,interval=secs_between_keys)
    else:
        pyautogui.alert(text=ticket_dict['description'].strip(), title='Unable to get barcode details', button='OK')



def main():
    try:
        ldir = tempfile.mkdtemp()
        parser = argparse.ArgumentParser()
        parser.add_argument('-d', '--directory', default=ldir)
        args = parser.parse_args()
        tmpDir = args.directory
        print (tmpDir)

        fname = tmpDir + 'setting.json'
        settingFile = fname
        print (fname)
        # regex = "Untitled - Notepad"
        # regex = "Microsoft Excel - Book1"
        regex = "Session A - [24 x 80]"
        state_left = win32api.GetKeyState(0x01)  # Left button down = 0 or 1. Button up = -127 or -128
        state_right = win32api.GetKeyState(0x02)  # Right button down = 0 or 1. Button up = -127 or -128

        import os.path
        
        w = WindowMgr()
        h = w.find_window(regex)
        # if h == None :
        #     sys.exit()
        

        if os.path.isfile(fname) :
            # w.set_mouseXY()
            print ('Configuration is existing')
        else:
            pos = w.set_onTop(h)
            w.Maximize(h)
            while True:
                a = win32api.GetKeyState(0x01)
                b = win32api.GetKeyState(0x02)

                if a != state_left:  # Button state changed
                    state_left = a
                    print(a)
                    if a < 0:
                        print('Left Button Pressed')
                    else:
                        print('Left Button Released')
                        w.saveFirstDataPos()
                        print (w.get_mouseXY())
                        break

                time.sleep(0.001)

        prev_barcode=''
        curr_barcode=''
        while True:
            ticket_info = get_data2()
            if ticket_info['status'] == 200 :
                curr_barcode = ticket_info['barcode']
                if curr_barcode != prev_barcode :
                    h = w.find_window(regex)
                    pos = w.set_onTop(h)
                    w.Maximize(h)
                    w.set_mouseXY()
                    fill_data (w,ticket_info)
                    prev_barcode = curr_barcode
            time.sleep(5)
            # ticket_number = pyautogui.prompt(text='Please scan Ticket number :', title='Scan Ticket Number' , default='')
            # if ticket_number == 'quit' or ticket_number == None  :
            #     print ('See you ,Bye Bye..')
            #     break
            # else :
            #     #----Older version----
            #     h = w.find_window(regex)
            #     pos = w.set_onTop(h)
            #     w.Maximize(h)
            #     w.set_mouseXY()
            #     ticket_info = get_data(ticket_number)
            #     fill_data (w,ticket_info)

            #     print ('Finished...')



        



    except:
        f = open(tmpDir + "log.txt", "w")
        f.write(traceback.format_exc())
        print(traceback.format_exc())






main()

#2558