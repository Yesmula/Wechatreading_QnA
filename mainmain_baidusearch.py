import win32gui,time,urllib,json,requests,re,base64,os
from PIL import ImageGrab
from PIL import Image
from bs4 import BeautifulSoup
from baidusearch.baidusearch import search
def get_window_pos(window_name):
    window = win32gui.FindWindow(None, window_name)
    if window:
        return win32gui.GetWindowRect(window)
    else:
        return None

def screenshot_beginning(window_offset=(0,0,0,0)):
    im0bbox=(100, 880, 400, 940)#开始游戏四个字
    zipped0=zip(im0bbox,window_offset)
    mapped0=map(sum, zipped0)
    sum0=tuple(mapped0)
    im0=ImageGrab.grab(sum0)
    #im0.save('t.png')

def screenshot_answering(pic_name,window_offset=(0,0,0,0)):
    im1bbox=(50, 290, 460, 720)
    im2bbox=(50, 450, 460, 720)
    zipped1=zip(im1bbox,window_offset)
    mapped1=map(sum, zipped1)
    sum1=tuple(mapped1)
    zipped2=zip(im2bbox,window_offset)
    mapped2=map(sum, zipped2)
    sum2=tuple(mapped2)
    im1=ImageGrab.grab(sum1)#第几题+题目+选项
    im2=ImageGrab.grab(sum2)#选项
    im1.save('%ss_1.png'%pic_name)
    im2.save('%ss_2.png'%pic_name)
#baidu ocr
def get_file_content_as_base64(path, urlencoded=False):
    with open(path, "rb") as f:
        content = base64.b64encode(f.read()).decode("utf8")
        if urlencoded:
            content = urllib.parse.quote_plus(content)
    return content

def get_window_offset(position):
    if position==None:
        print('Cannot get position!')
        window_offset=(0,0,0,0)
    else:
        window_offset=(position[0]*1.25,position[1]*1.25,position[0]*1.25,position[1]*1.25)
    return window_offset

def get_access_token():
    url = "https://aip.baidubce.com/oauth/2.0/token"
    params = {"grant_type": "client_credentials", "client_id": API_KEY, "client_secret": SECRET_KEY}
    return str(requests.post(url, params=params).json().get("access_token"))

def getsentence(gg):
    ggnum=gg['words_result_num']
    gsentence=''.join(gg['words_result'][i]['words'] for i in range(ggnum))
    return gsentence

def get_picture_dic_response(pic_location):
    url = "https://aip.baidubce.com/rest/2.0/ocr/v1/general_basic?access_token=" + get_access_token()
    payload='image='+get_file_content_as_base64(r"%s"%pic_location,True).replace('/', '%2F').replace('=', '%3D').replace('+', '%2B')
    headers = {'Content-Type': 'application/x-www-form-urlencoded','Accept': 'application/json'}
    response = requests.request("POST", url, headers=headers, data=payload)
    dic_response=json.loads(response.text)
    return dic_response

def get_question(qdic,qnum_options):
    qquestion=''.join(qdic['words_result'][i]['words'] for i in range(1,qdic['words_result_num']-qnum_options))
    return qquestion

def get_search_results_baidu(search_term):
    print('Searching...')
    results = search(search_term)
    final_result=''
    for result in results:
        result=(result['title']+result['abstract']).replace(" ", "").replace("\n", "")
        final_result+=result
    return final_result

def get_coincidence(gchoices,gresults):
    sm=[]
    for i in range(len(gchoices)):
        p=gresults.count(gchoices[i])**4
        sm.append(p)
    sm_sort=sm.copy()
    sm_sort.sort()
    if all(si == 0 for si in sm) or sm_sort[-1]==sm_sort[-2]:
        print('Search by character...')
        sm=[]
        for i in range(len(gchoices)):
            k=0
            for j in gchoices[i]:
                character_position=gresults.find(j)
                k+=gresults.count(j)*(character_position+2)**(-i)
            p=k**4
            sm.append(p)
            if all(si == 0 for si in sm):
                print('Sorry, I cannot decide.')
                return 0
    sumsm=sum(sm)
    for i in range(len(sm)):
        sm[i]=sm[i]/sumsm
    return sm

def compare_images(image1_path, image2_path):
    image1 = Image.open(image1_path)
    image2 = Image.open(image2_path)
    if image1.mode != image2.mode:
        return 0
    if image1.size != image2.size:
        return 0
    pairs = zip(image1.getdata(), image2.getdata())
    if len(image1.getbands()) == 1:
        # for gray-scale jpegs
        dif = sum(abs(p1 - p2) for p1, p2 in pairs)
    else:
        dif = sum(abs(c1 - c2) for p1, p2 in pairs for c1, c2 in zip(p1, p2))

    ncomponents = image1.size[0] * image1.size[1] * 3
    difference = (dif / 255.0 * 100) / ncomponents
    similarity = 100 - difference
    return similarity / 100

def get_two_digit_array(array0):
    try:
        arr=[]
        for i in range(len(array0)):
            arr.append(round(array0[i], 2))
        return arr
    except:
        return ['The parray is wrong.']
def main():
    tottime=500
    k=0
    os.makedirs('./Wechatreading_QnA',exist_ok = True)
    for i in range(tottime):
        screenshot_answering('./Wechatreading_QnA/%d'%i,window_offset)
        if i>0:
            pic_coincidence=compare_images('./Wechatreading_QnA/%ds_1.png'%i,'./Wechatreading_QnA/%ds_1.png'%(i-1))
            if pic_coincidence<0.95:
                k=0
                time.sleep(1)
                screenshot_answering('./Wechatreading_QnA/%d'%i,window_offset)
                b_1_dic=get_picture_dic_response('./Wechatreading_QnA/%ds_1.png'%i)
                b_1_dic_0=b_1_dic['words_result'][0]['words']
                if b_1_dic_0.find('第')>=0 or b_1_dic_0.find('题')>=0 or True:
                    print('pic_coincidence=%.3f,识别'%pic_coincidence)
                    b_2_dic=get_picture_dic_response('./Wechatreading_QnA/%ds_2.png'%i)
                    num_options=b_2_dic['words_result_num']
                    question=get_question(b_1_dic,num_options)
                    choices=[]
                    for i in range(b_2_dic['words_result_num']):
                        choices.append(b_2_dic['words_result'][i]['words'])
                    print(question,choices)
                    search_results=get_search_results_baidu(question)
                    print(search_results)
                    parray=get_two_digit_array(get_coincidence(choices,search_results))
                    print(parray)
                else:
                    print('pic_coincidence=%.3f,未获得题目,不识别'%pic_coincidence)
                    continue
            else:
                os.remove('./Wechatreading_QnA/%ds_1.png'%(i-1))
                os.remove('./Wechatreading_QnA/%ds_2.png'%(i-1))
                if i%5==0:
                    print('pic_coincidence=%.3f,不识别'%pic_coincidence)
                    k+=1
                    if k>1:
                        print('I Wil Quit In 3s.')
                        time.sleep(3)
                        return 0
        time.sleep(1)

API_KEY = ""
SECRET_KEY = ""
try:
    print('Welcome to Wechatreading_QnA!')
    #position=get_window_pos('微信读书')
    position=(0,0,0,0)
    window_offset=get_window_offset(position)
    if position[0]<0:
        print('Please Show Applet Window! I Will Quit In 3s.')
        time.sleep(3)
    else:
        print('The Window Position:',position)
        print('window_offset=',window_offset)
        main()
except:
    print('Please Open Wechatreading Applet. I Wil Quit In 3s.')
    time.sleep(3)



