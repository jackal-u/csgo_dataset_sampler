import time,os,sys,oss2,shutil
import zipfile
from pyautogui import position
import win32gui
import win32api
import win32con
import win32ui
from win32gui import GetDesktopWindow,GetWindowRect
import pygame
from pynput import mouse
from pynput import keyboard as pk
from PIL import Image
from PIL import ImageGrab
import json
import numpy as np
import _thread as thread
import cv2
##pynput.keyboard.KeyCode.from_char('a')
'''
press alt to begin/end recording
图像文件夹+json
'''
pre_class2key=None
def on_click(x, y, button, pressed):
    if str(button) in key_names:
        idx=key_names.index(str(button))
        key_state_np[idx]=int(pressed)
        time_list[idx]=time.time()
def on_scroll(x,y,dx,dy):
    idx=key_names.index('scroll')
    key_state_np[idx]=int(4 if dy>0 else 7)
    time_list[idx]=time.time()
def on_press(key):
    global recording,pre_class2key,upload
    key=str(key)
    if key[:7]=='Key.alt':
        if not recording and (not thread_num): #若不在录制中且不在保存中
            recording=True
            ssound.play()
        else:
            recording=False
            esound.play()
    elif key[:7]=='Key.esc':
        if not recording and (not thread_num): upload=True
    else:
        if key[1]=='\\' and (key in special_dict): name=special_dict[key]
        else: name=key.lower()
        name=name[:min(len(name),8)]
        if name in key_names1:
            idx=key_names.index(name)
            key_state_np[idx]=1
            time_list[idx]=time.time()
        elif name in key_names2 and (name!=pre_class2key):
            idx=key_names.index(name)
            key_state_np[idx]=2
            time_list[idx]=time.time()
            pre_class2key=name
def on_release(key):
    global pre_class2key
    key=str(key)
    if key[1]=='\\' and (key in special_dict): name=special_dict[key]
    else: name=key.lower()
    name=name[:min(len(name),8)]
    if name in key_names1:
        idx=key_names.index(name)
        key_state_np[idx]=0
        time_list[idx]=time.time()
    elif name in key_names2:
        pre_class2key=None

begin=[552,92];wh=[928,736]
dimensions=[]
#hdesktop = GetDesktopWindow()
#dimensions = GetWindowRect(hdesktop)
width = win32api.GetSystemMetrics(win32con.SM_CXVIRTUALSCREEN)
height = win32api.GetSystemMetrics(win32con.SM_CYVIRTUALSCREEN)
left = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
top = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)
recording=False
upload=False
keyboard_listener=pk.Listener(on_press=on_press,on_release=on_release)
mouse_listener=mouse.Listener(on_click=on_click,on_scroll=on_scroll)
pygame.init()
ssound = pygame.mixer.Sound('./assets/alert_start.wav')
esound = pygame.mixer.Sound('./assets/alert_end.wav')
keyboard_listener.start()
mouse_listener.start()

key_names=['Button.left','Button.right','scroll',"'w'","'a'","'s'","'d'","'r'","'q'","'e'","'b'",\
           "'1'","'2'","'3'","'4'","'5'","'6'","'7'","'8'","'9'","'0'",'key.shif','key.spac','key.ctrl']#所有按键
key_names1=["'w'","'a'","'s'","'d'",'key.shif','key.ctrl']#连按按键
key_names2=["'r'","'q'","'e'","'b'","'1'","'2'","'3'","'4'","'5'","'6'","'7'","'8'","'9'","'0'",'key.spac']#单次按键
special_dict={"'\\x17'":"'w'","'\\x01'":"'a'","'\\x13'":"'s'","'\\x04'":"'d'"}
n_action=len(key_names)
key_state_np=np.int_(np.zeros(n_action))
time_list=[0]*n_action
step_save_num=200
thread_num=0
frame_rates=[]


def beginer():#设置录制区域，此代码没用到
    global begin,wh,dimensions
    dimensions=[]
    prexy=[];aftxy=[]
    for i in range(2):
        if i==0:
            txt=input('设置起始位置：鼠标放在起始位置，任意键选定')
            x, y = position()
            print('鼠标坐标 : (%s, %s)' % (x, y))
            prexy=[x,y]
        else:
            txt=input('设置终止位置：鼠标放在终止位置，任意键选定')
            x, y = position()
            print('鼠标坐标 : (%s, %s)' % (x, y))
            aftxy=[x,y]
    best_wh=[abs(aftxy[0]-prexy[0]),abs(aftxy[1]-prexy[1])]
    center=[(prexy[0]+aftxy[0])//2,(prexy[1]+aftxy[1])//2]
    wh=[round(best_wh[0]/32)*32,round(best_wh[1]/32)*32]
    begin=[center[0]-wh[0]//2,center[1]-wh[1]//2]
    if not (wh[0]<1 or wh[1]<1):
        dimensions+=begin
        dimensions+=[begin[0]+wh[0],begin[1]+wh[1]]
    else: dimensions=[]
    print('按alt开始与停止:')
    now=time.strftime('%y%m%d%H%M',time.localtime())


hwin = win32gui.GetDesktopWindow()
def capture_screenshots():
    # 创建设备描述表
    desktop_dc = win32gui.GetWindowDC(hwin)
    img_dc = win32ui.CreateDCFromHandle(desktop_dc)
    mem_dc = img_dc.CreateCompatibleDC()
    #print(width,height,left,top,img_dc)#全屏时：1920 1080 0 0
    # 创建位图对象
    screenshot = win32ui.CreateBitmap()
    screenshot.CreateCompatibleBitmap(img_dc, width, height)
    mem_dc.SelectObject(screenshot)
    # 截图至内存设备描述表
    mem_dc.BitBlt((0, 0), (width, height), img_dc, (left, top), win32con.SRCCOPY)
    # 将截图保存到列表中
    bmp_info = screenshot.GetInfo()
    bmp_str = screenshot.GetBitmapBits(True)
    img = np.frombuffer(bmp_str, dtype='uint8')
    img.shape = (bmp_info['bmHeight'], bmp_info['bmWidth'], 4)
    # 释放内存
    win32gui.DeleteObject(screenshot.GetHandle())
    mem_dc.DeleteDC()
    img_dc.DeleteDC()
    win32gui.ReleaseDC(hwin, desktop_dc)
    return img

json_file=''
def img_capture(dimensions):
    print('\n\n按alt键开始录制，esc键上传录制结果',end='')
    global key_state_np,time_list,game_frames,count,thread_num,json_file,upload
    begin_time=time.time()
    folder_path='./result/'+'{:0<20}'.format(begin_time).replace('.','-')
    json_path=folder_path+'/'+'{:0<20}'.format(begin_time).replace('.','-')+'.json'
    key_state_np=np.int_(np.zeros(n_action))
    time_list=[0]*n_action
    game_frames=[]
    pre_time=0
    count=0
    while True:
        if recording: 
            if count:#得不是第一帧
                game_frames[count-1].append((key_state_np,time_list,position()))
                if count%step_save_num==0: #每隔step_save_num，清空内存，新开线程写入文件
                    frame_rate=step_save_num/(time.time()-pre_time)
                    frame_rates.append(frame_rate)
                    print(frame_rate,end='')
                    frames=game_frames.copy()
                    count=0;game_frames=[];thread_num+=1
                    pre_time=time.time()
                    thread.start_new_thread( frames2json, (frames,folder_path,json_path))
            else: pre_time=time.time() #记录起始时间
            print()
            print(key_state_np,position(),end='')
            ft=time.time()
            key_state_np=np.int_(key_state_np==1);time_list=[0]*n_action
            #img=ImageGrab.grab(dimensions) #这个达19帧
            img=capture_screenshots() #这个确实好一些21帧可以
            game_frames.append([img,ft])
            count+=1
        elif count:
            print()
            thread_num+=1
            thread.start_new_thread( frames2json, (game_frames[:-1],folder_path,json_path))
            break
        elif upload:
            break
    json_file=json_path

def frames2json(frames,folder,jsonfile):#每隔一段时间就存图片入磁盘
    global thread_num
    def opencv_save(img,quality=50):
        #img = cv2.cvtColor(img,cv2.COLOR_RGB2BGR) #用ImageGrab时才用
        img=cv2.resize(img, (1920, 1080))
        img_path=folder+'/{}.jpg'.format('{:0<20}'.format(frame[1]).replace('.','-'))
        cv2.imwrite(img_path,img,[int(cv2.IMWRITE_JPEG_QUALITY), quality])
        return img_path
    '''def PIL_save(img,quality=75):#不如opencv_save
        img = Image.fromarray(img).resize((1920, 1080),Image.BILINEAR).convert('BGR')
        img_path=folder+'/{}.jpg'.format('{:0<20}'.format(frame[1]).replace('.','-'))
        img.save(img_path,quality=quality)'''
    if not os.path.isdir(folder):
        os.mkdir(folder)
        f=open(jsonfile,'w');f.write('{}');f.close()
    frame_dict={}
    for frame in frames:
        img=np.array(frame[0])
        img_path=opencv_save(img)
        #img_path=PIL_save(img) #不如opencv_save
        frame_dict[frame[1]]={'img':img_path,'mouse_pos':str(list(frame[2][2])),'action':str(list(frame[2][0])),'act_time':str(frame[2][1])}
    with open(jsonfile, 'r+') as f:
        frames_json = json.load(f)
        frames_json.update(frame_dict)
    with open(jsonfile, 'w') as f:
        json.dump(frames_json, f, indent=4)
    f.close()
    thread_num-=1
    print('save {} frames'.format(len(frames)),end='')
        
def time_check(jsonfile):#检查并保证时间对齐
    print('\ntime checking...')
    frames_json = json.load(open(jsonfile))
    pre_key=list(frames_json)[0]
    any_wrong_flag=False
    for key in list(frames_json)[1:]:
        act_list=eval(frames_json[key]["action"])
        time_list=eval(frames_json[key]["act_time"])
        pre_act_list=eval(frames_json[pre_key]["action"])
        pre_time_list=eval(frames_json[pre_key]["act_time"])
        wrong_flag=False
        for i in range(len(time_list)):
            t=time_list[i]
            if t and t<eval(key):
                wrong_flag=True
                pre_act_list[i]=act_list[i]
                pre_time_list[i]=t
                act_list[i]=int(pre_act_list[i]==1)
                time_list[i]=0
        if wrong_flag:
            print('*{}:{}'.format(key,frames_json[key]["act_time"]))
            frames_json[pre_key]["action"]=str(pre_act_list)
            frames_json[pre_key]["act_time"]=str(pre_time_list)
            frames_json[key]["action"]=str(act_list)
            frames_json[key]["act_time"]=str(time_list)
        pre_key=key
    if not any_wrong_flag: print('无错误')
    with open(jsonfile, 'w') as f:
        json.dump(frames_json, f, indent=4)
    f.close()

# 阿里云账号AccessKey拥有所有API的访问权限，风险很高。强烈建议您创建并使用RAM用户进行API访问或日常运维，请登录RAM控制台创建RAM用户。
auth = oss2.Auth('LTAI5tFks9mrco621mv897GU', 'OCPCoZZibZkHq075Us3x3XkENgnqLM')
# yourEndpoint填写Bucket所在地域对应的Endpoint。以华东1（杭州）为例，Endpoint填写为https://oss-cn-hangzhou.aliyuncs.com。
# 填写Bucket名称。
bucket = oss2.Bucket(auth, 'oss-cn-beijing.aliyuncs.com', 'csgoai')
def zip_upload(result_path='.\\result'):
    def get_file_paths(result_root):
        folder_paths = [];zip_paths=[];file_names=[]
        for root, dirs, files in os.walk(result_root):
            for folder in dirs:
                folder_paths.append(os.path.join(root, folder))
            for file in files:
                if file[-3:]=='zip':
                    file_names.append(file)
                    zip_paths.append(os.path.join(root, file))
        return folder_paths,zip_paths,file_names
    def zip_folder(folder_path, output_path):
        print('正在压缩文件夹：'+folder_path+'...',end=' ')
        zip_file = zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED)
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                zip_file.write(os.path.join(root, file), file)
        zip_file.close()
        print('得到'+folder_path+'.zip',end=' ')
        shutil.rmtree(folder_path)
        print('已删除原文件夹')
    def percentage(consumed_bytes, total_bytes):
        if total_bytes:
            rate = int(100 * (float(consumed_bytes) / float(total_bytes)))
            print('\r{0}% '.format(rate), end='')
            sys.stdout.flush()
    folder_paths,_,_=get_file_paths(result_path)
    if folder_paths:
        print('\n压缩以下文件夹：')
        for folder in folder_paths:
            print(folder)
        for folder in folder_paths:
            zip_folder(folder,folder+'.zip')
    else:print('无文件夹需压缩')
    _,zip_paths,file_names=get_file_paths(result_path)
    if zip_paths:
        print('\n上传{}个文件：'.format(len(zip_paths)),end='')
        for i in range(len(zip_paths)):
            print('\n{}上传：'.format(file_names[i]))
            bucket.put_object_from_file(file_names[i],zip_paths[i], progress_callback=percentage)
            os.rename(zip_paths[i],zip_paths[i]+'-uploaded')
    else: print('无压缩文件需上传')
    #以下测试使用：
    #print()
    #for obj in oss2.ObjectIterator(bucket):
    #    print(obj.key)

def cac_demensions(w_final=1920,h_final=1080,cut_width=True):
    global dimensions,width,height,left,top
    hdesktop = GetDesktopWindow()
    dimensions = GetWindowRect(hdesktop)
    center=(int((dimensions[0]+dimensions[2])/2),int((dimensions[1]+dimensions[3])/2))
    good_w=dimensions[3]*w_final/h_final
    dimensions=list(dimensions)
    if cut_width and dimensions[2]>good_w:
        dimensions[0]=center[0]-int(good_w/2)
        dimensions[2]=center[0]+int(good_w/2)
        width=int(good_w)
        left=dimensions[0]
    
if __name__=='__main__':
    try:
        while True:
            #beginer() #弃用
            path='.\\result'
            if not os.path.isdir(path):
                os.mkdir(path)
            if not dimensions:
                cac_demensions(cut_width=True)
                print(dimensions,[width,height,left,top])
            if not thread_num:#目前已无保存进程，继续下一次录制
                if frame_rates:
                    print('平均帧率：',sum(frame_rates)/len(frame_rates))
                '''while json_file:
                    c=input('\n继续检查并修正时间顺序? y/[n]')
                    if c=='y': time_check(json_file)
                    else: break'''
                img_capture(dimensions)
                pre_thread_num=0
                if upload:
                    zip_upload(path)
                    upload=False
            else:
                if pre_thread_num and thread_num!=pre_thread_num: print()
                pre_thread_num=thread_num
                print('\r请稍等，{}个保存进程执行中...'.format(thread_num),end='')
    except Exception as error:
        print("出错了 (╯ŏ‿ŏ)╯ ┴ ┴ ")

os.system('pause')