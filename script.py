from json.decoder import JSONDecodeError
from typing import Dict
import aiohttp
import xlsxwriter
import requests, asyncio
from openpyxl import load_workbook

           
def ifIter(v):
    dictionary = {1:1}
    lists = [0]
    if type(v) == type(dictionary) or type(v) == type(lists):
        return 1
    else:
        return 0
    
    
def ifList(v):
    lists = [0]
    if type(v) == type(lists) and len(v)>0 :
        return 1
    else:
        return 0
  
        
def ifDict(v):
    dictionary = {1:1}
    if type(v) == type(dictionary) and len(v)>0:
        return 1
    else:
        return 0
    
    
def printDict(v0):
    if ifDict(v0):
        for k1,v1 in v0.items():
            return(k1, v1)
    elif ifList(v0):
        for v1 in v0:
            return(v1)
    else:
        return(v0)
         
         
def random_chr(n,m=0,k=1, binary=1):
    if binary:
        rand=''
        if m==-1:
            n,m=m,n
        for i in range(k):
            rand+= chr(n+i)
        return rand
    return chr(n)
    
    
def url_creator(point, url_d={}, debug=0):
    if url_d!={}:
        
                    
        pass
    else:
        
        mw= "mw-tst.itsmartflex.com/" 
        #mw = "mw-apitst.vodafone.ua/" #mw_v
        
        phone_rp = "380662584192"
        phone_сs= "380662583046"
        phone_b= "380662583045"
        phone_o= "380662584418"
        product_id = "?product.id="
        siebel_token = "?siebelToken=177C50397B39937400431DC33E982F3D"
        
        point_s = "MYVF-SETTINGS"
        point_l = "MYVF-LINKS"
                
        get_token = "uaa/oauth/token?grant_type=client_credentials"
        relat_phone = "customer/api/customerManagement/v3/customer/"
        counters_i_bonus = "prepaybalance/tmf-api/prepayBalanceManagement/v4/bucket"
        offer = "qualification/api/productOfferingQualificationManagement/v1/productOfferingQualification/"
        setting_i_links = "entity/api/functions/"
        
        proto = "https"
        slash = "://"
        
    if (point=="token"):
        url = proto + slash + mw + get_token
    if (point=="relPhone"):
        url = proto + slash + mw + relat_phone + phone_rp
    if (point=="countMain" or point=="countDpi"):
        url = proto + slash + mw + counters_i_bonus + product_id + phone_сs
    if (point=="bonus"):
        url = proto + slash + mw + counters_i_bonus + '/' + phone_b
    if (point=="offer"):
        url = proto + slash + mw + offer + phone_o + siebel_token
    if (point=="settings"):
        url = proto + slash + mw + setting_i_links + point_s
    if (point=="links"):
        url = proto + slash + mw + setting_i_links + point_l
        
    log_print(debug, 2, 'url_creator ',url, point)
    
    return url


def header_creator(point, token='', debug=0, fil_type = 0, filler= ''):
 
    log_print(debug, 2, 'heador_creator', filler)
    filler0 = filler if fil_type>=0 else ''
    filler1 = filler if fil_type>=1 else ''
    filler2 = filler if fil_type>=2 else ''
    filler3 = filler if fil_type>=3 else ''
    filler4 = filler if fil_type>=4 else ''
    filler5 = filler if fil_type>=5 else ''
    filler6 = filler if fil_type>=6 else ''
    filler7 = filler if fil_type>=7 else ''
    lang = ['Accept-Language'+filler6 , 'ru'+filler0]
    cont = ['Content-Type'+filler5 , "application/json"+filler1]
    auth = ['Authorization'+filler4 , 'Basic aW50ZXJuYWw6aW50ZXJuYWw='+filler2 ]
    auth2 = ['Authorization'+filler4 , 'Bearer ' + token +filler2 ]
    prof = ['Profile' + filler7] 
    if (point=="token"):
        headers = {lang[0] : lang[1],
                   cont[0] : cont[1],
                   auth[0] : auth[1]  }
    if (point=="settings" or point=="links"):
        headers = {lang[0] : lang[1],
                   cont[0] : cont[1],
                   auth2[0] : auth2[1] }
    if (point=="relPhone"):
        headers = {prof[0] : 'RELATED-PARTY' + filler3,
                   lang[0] : lang[1],
                   cont[0] : cont[1],
                   auth2[0] : auth2[1] }
    if (point=="countMain"):
        headers = {prof[0]: 'COUNTERS' + filler3,
                   lang[0] : lang[1],
                   cont[0] : cont[1],
                   auth2[0] : auth2[1] }
    if (point=="countDpi"):
        headers = {prof[0]: 'COUNTERS-DATA' + filler3,
                   lang[0] : lang[1],
                   cont[0] : cont[1],
                   auth2[0] : auth2[1] }
    if (point=="bonus"):
        headers = {prof[0]: 'BONUS' + filler3,
                   lang[0] : lang[1],
                   cont[0] : cont[1],
                   auth2[0] : auth2[1] }
    if (point=="offer"):
        headers = {prof[0]: 'ACTIVE-OFFER' + filler3,
                   lang[0] : lang[1],
                   cont[0] : cont[1],
                   auth2[0] : auth2[1] }
    
    log_print(debug, 4,'header_creator2',point)
    log_print(debug, 5,'header_creator3',headers)
    return headers
         
              
def data_creator(s, data = {}):
    if data:    
        for d in data:
            data = {d +  s : data[d] + s} 
    else:
            data = {s : s} 
    return data
 
 
def manageRequest(point, verb , token=None, headers={}, data={}, debug=0, log={} ):
    if token==None:
        url = url_creator('token', {}, debug)
        headers = header_creator('token', '', debug)
        r = _requests(url, 'post', headers, data, debug)
        token = r.json()["access_token"]
        log_print(debug, 4,'manageRequest token', token)
    
    url = url_creator(point, {}, debug)
    headers = header_creator(point, token, debug)
    r = _requests(url, verb, headers, data, debug)    
    return _print(r, verb, url, point, debug, log)    


def many_req_gen(verbs, points, token, data={}, debug=0 ):
    
    for point in points:
        log_print(debug, 1, 'many_rew_gen0', point, 'gen' )
        for verb in verbs:
            log_print(debug, 2, 'many_rew_gen1', verb, 'gen' )
            url = url_creator(point, {}, debug)
            headers = header_creator(point, token, debug)        
            x = (verb, url, point, headers, data, debug)
            log_print(debug, 3,'many_rew_gen2', x)
            log_print(debug, 4,'many_rew_gen3', 'yield')
            yield x


def create_filler(n, s='a'):
    st = ''
    for i in range(n):
        st+= s
    return st


def log_print(debug, n, *message):
    if debug>= n:
        print(*message)

 
def gen_u(point, quirks, debug=0):
    if quirks['change'] == 'url':
        if quirks['mode'] == 'content':
            filler = quirks['chr'] * quirks['content_num']
            for i in range(quirks['step']):
                url = url_creator(point, {}, debug) + filler
                filler = filler + quirks['chr']* quirks['content_num']
                log_print(debug, 4, 'gen_u', url)
                yield url
                
        elif quirks['mode'] == 'change':
            for i in quirks['stop'] - quirks['start']:
                filler =  chr( [quirks['start'] +i ] ) 
                url = url_creator(point, {}, debug) + filler
                log_print(debug, 4, 'gen_u', url)
                yield url
    else:
        while True:
            yield url_creator(point, {}, debug)
 
 
 
def gen_h(point, token, quirks, debug=0):
    log_print(debug, 4, 'gen_h 0', quirks)
    if quirks['change'] == 'header':
        log_print(debug, 4, 'gen_h 1')
        if quirks['mode'] == 'content':
            log_print(debug, 4, 'gen_h 2')
            filler = quirks['chr'] * quirks['content_num']
            for i in range(quirks['step']):
                log_print(debug, 4, 'gen_h 3')
                header = header_creator(point, token, debug, quirks['change_type'], filler)
                log_print(debug, 4, 'gen_h3.5 lenof filler', len(filler))
                filler = filler + quirks['chr']* quirks['content_num']
                log_print(debug, 4, 'gen_h 4',header)
                yield header
                
        elif quirks['mode'] == 'change':
            for i in quirks['stop'] - quirks['start']:
                filler =  chr( [quirks['start'] +i ] ) 
                header = header_creator(point, token, debug, quirks['change_type'], filler)
                log_print(debug, 4, 'gen_h6', header)
                yield header
    else:
        while True:
            yield header_creator(point, token, debug)


def gen_d(quirks, data={}, debug=0):
    if quirks['change'] == 'data':
        log_print(debug, 4, 'gen_d 0')
        if quirks['mode'] == 'content':
            log_print(debug, 4, 'gen_d 1')
            filler = quirks['chr'] * quirks['content_num']
            for i in range(quirks['step']):
                log_print(debug, 4, 'gen_d 2')
                data = data_creator(filler, data) 
                filler = filler + quirks['chr']* quirks['content_num']
                log_print(debug, 4, 'gen_d 4', data)
                yield data
                
        elif quirks['mode'] == 'change':
            for i in quirks['stop'] - quirks['start']:
                data = data_creator(filler, data) 
                filler =  chr( [quirks['start'] +i ] ) 
                log_print(debug, 4, 'gen_d', data)
                yield data
    else:
        while True:
            yield data
     
           
def many_req_quirks_gen(verbs, points, token, quirks, data, debug ):
    
    log_print(debug, 4,' many_req_quirks_gen quirks mode')
    for point in points:
        for verb in verbs:
            subgen_h = gen_h(point, token, quirks, debug)
            subgen_u = gen_u(point, quirks, debug)
            subgen_d = gen_d(quirks, data, debug)
            if quirks['mode'] == 'change':
                n = (quirks['stop_chr'] - quirks['start_chr'] )/ quirks['step_chr']
            if quirks['mode'] == 'content':
                n = quirks['step']
                 
            for i in range(n):
                header = next(subgen_h)
                url = next(subgen_u)
                data = next(subgen_d)
                x = (verb, url, point, header, data, i)

                log_print(debug, 4, ' many_req_quirks_gen', i,n)
            yield x                       
   
       
async def _requests(verbs, points, token, quirks=None, headers = {}, data= {}, debug=0):
    if quirks is not None:
        gen = many_req_quirks_gen(verbs, points, token, quirks, data, debug )
        log_print(debug,2,'_requests1')
    else:
        gen = many_req_gen(verbs, points, token, data, debug )
        log_print(debug,2,'_requests2')
    log_print(debug,2,'_requests3')
    async with aiohttp.ClientSession() as session:
        log_print(debug,2,'_requests session')
        q=0
        try:
            while x:= next(gen):
                q+=1
                verb, url, point, headers, data, filler = x #next(gen)
                async with session.request(verb, url, headers=headers, data=data) as response:
                # async with session.post(url, headers=headers) as response:
                    status =  response.status
                    log_print(debug, 1, '_requests, in session', status)
                    try:
                        resp = await response.json()
                    except aiohttp.client_exceptions.ContentTypeError:
                        log_print(debug, 1, '_requests not a json, q', q)
                        resp = await response.text()

                    if status>=200 and status<300:
                        log[0][(point, url, verb, filler)] = (status, headers, data, resp)
                    else:
                        log[1][(point, url, verb, filler)] = (status, headers, data, resp)
            await session.close()
        except StopIteration as e:
            log_print(debug, 1,'_requests', e, q,filler)


def fetchToken(url=None, headers=None, mw = None):
    if mw is None:
        mw = 'mw-tst.itsmartflex.com'
    elif mw == 'vodafone':
        mw = 'mw-apitst.vodafone.ua' #mw-tst.itsmartflex.com
    
    if url is None:
        url = 'https://' + mw + '/uaa/oauth/token?grant_type=client_credentials'
        url2 = 'https://' + mw + '/uaa/oauth/token?grant_type=password'
    if headers is None:
        basic1 = 'aW50ZXJuYWw6aW50ZXJuYWw='
        basic2= 'c3ZjLXNpZWJlbC1tdzoxMjM0='
        headers = {"Authorization": "Basic " + basic1,"Content-Type": "application/json"}
        headers2 = {"Authorization": "Basic c3ZjLXNpZWJlbC1tdzoxMjM0=","Content-Type": "application/x-www-form-urlencode"}
    # data = {'username' :'380952240016'}
    print(url)
    print(headers)
    r = requests.post(url, headers = headers)#, data = data) 
    print(r.status_code)
    json_response = r.json()
    return json_response['access_token']


def fetchPass(token = None, url=None, headers=None, mw = None):
    if mw is None:
        # mw = 'mw-apitst.vodafone.ua'
        mw = 'mw-tst.itsmartflex.com'
    if url is None:
        # url = 'https://' + mw + '/uaa/oauth/token?grant_type=client_credentials'
        url = 'https://' + mw + '/uaa/oauth/token?grant_type=password'
    if token is None:
        token = 'eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9.eyJzY29wZSI6WyJtc191YWFfY3JlZGVudGlhbF90eXBlX3Bhc3N3b3JkX2p3dCIsIm9wZW5pZCJdLCJleHAiOjE2MTE3NjI5NzksImFkZGl0aW9uYWxEZXRhaWxzIjp7Im1zaXNkbiI6IjM4MDk1MjI0MDAxNiJ9LCJhdXRob3JpdGllcyI6WyJST0xFX1NZU1RFTSJdLCJqdGkiOiI3NmRiNTEyMS0zZjJmLTQzZTMtYTJjZi1lMTA4NGNjM2YzZTciLCJ0ZW5hbnQiOiJYTSIsImNsaWVudF9pZCI6InN2Yy1zaWViZWwtbXcifQ.d50VVTF50QF-nqvxrmzcXWg9Ih0qRdTlPYdfhi_RKzgDL_RjDtu_sIiv6FWLKGyAgYcpVD-ukL2F8BU53OGAqgEbJo0zJLzVXEIUDcdc-zxZufPf-x6MtFhwSo4Xk6v22esSgPtU-LD2Kk8wfr4_yA0BChu4OXS2uXwvy7qaQS7h889kIyO2pSeHBJdtS0-EjrP4cFebd32hFZ9J_25ockKjJ17y08ubYQh_YMrbOzHlWrfmbYhMVN71uadgOGfRnwkSWzVJ-GThcFasXTuK7ONKa748sCQWG_q4R7QEzZbvDnFk2Dw-VXyb73kOJQKOJY3QblDTVmkqF-yg8DMKGQ'

    if headers is None:
        # headers = {"Authorization": "Basic aW50ZXJuYWw6aW50ZXJuYWw=","Content-Type": "application/json"}
        # –header ‘Content-Type: application/x-www-form-urlencoded’
        headers = {"Authorization": "Basic YXBwLW15dm9kYWZvbmUtbXc6UGg1ZDg2QnpCNg==","Content-Type": "application/x-www-form-urlencoded", 'Profile': 'MYVODAFONE'}
        headers2 = {"Authorization": "Basic c3ZjLXNpZWJlbC1tdzoxMjM0’=","Content-Type": "application/x-www-form-urlencode"}
    data = {'username' :'380952240016', 'password' : token}

    print(url)
    print(headers)
    r = requests.post(url, headers = headers, data = data) 
    print(r.status_code)
    try:
        json_response = r.json()['access_token']
    except:
        json_response = r.text
        print(json_response)
    return json_response


def print_elem(dic, key):
    if type(dic) == type(dictionary) and dic.get(key, 0):
        return key
    else:
        return key + ' ' + dic[key]


def isType(e):  #-1 None, 0 if empty array, 1 if str int float, 2 if list, 3 if dict
    if e is None:
        return -1
    if e == [] or e == {} or e == ():
        return 0
    if str(type(e)) == "<class 'bool'>" or str(type(e)) == "<class 'int'>" or str(type(e)) == "<class 'float'>" or str(type(e)) == "<class 'str'>":
        return 1 
    if str(type(e)) == "<class 'list'>":
        return 2 
    if str(type(e)) == "<class 'dict'>":
        return 3  


def is_int(i):
    try:
        int(i)
        return 1 
    except:
        return 0
    
    
def none_safe_str(s):
    if s is None:
        return s
    else:
        return str(s)
    
    
def checkArray(el): #0 str, 1 list, 2 dict, 3 listdict, 4 listlist, 5 dictlist, 6 dictdict
    if isType(el)<2: #isType(): -1 None, 0 if empty array, 1 if str int float, 2 if list, 3 if dict
        return 0
    for e in el:
        # print('checkarray0',isType(e),type(e))
        # print(e)
        if isType(el)==2: #list
            if isType(e)>=2: #list or dict
                return isType(e) + 1
            else: return 1
        if isType(el)==3: #dict
            if isType(el[e])>=2: #list or dict
                return isType(el[e]) + 3
            else: return 2
    return 0        
        
        
def parse_string(s):
    for c in s:
        if c=='[':
            pass
        if c=='{':
            print()
        # if 


def parse_str(s,n, pre=' ', dic=0):
    if isType(s)<2:
        yield (n,s,dic)
    if isType(s)==2:
        for c in s:
            yield from parse_str(c, n+1, ' ',dic)
        print()
    if isType(s)==3:
        for c in s:
            yield (n+1,c,dic)
            yield from parse_str(s[c], n, pre+' ', dic+1)
        dic -=1
    
    
example = [{'a':[1,2],'b':3},{'w':[6,{'q':4}],'e':{'r':5}}]
g = parse_str(resp, 1)
print(next(g))


point = 'settings'
verb ='get'
url = url_creator(point)
token = fetchToken()
header = header_creator(point, token)
r = requests.request(method = verb, url= url, headers=header)
resp = r.json()
resp =_
example = [{'a':[1,2],'b':{3}},{'w':{'q':4}}]
print(str(example))
print(repr(example))



q = 3.21
type(q)
token = fetchToken()
header = header_creator(point, token = token)
url = url_creator(point)
r = requests.request(verb, url, headers = header )
r
r.json()
resp = r.json()
filename = 'testings.xlsx'
example = [{'a':[1,2],'b':3},{'w':[6,{'q':4}],'e':{'r':{'r':5,'t':6},'t':[6,7,8]},'y':9 },5]



workbook = xlsxwriter.Workbook(filename)
sett = workbook.add_worksheet('settings')

write_to_excel(resp, filename,verbose=3, new_worksheet='testings',x0=2,y0=0)
workbook,worksheet,i = write_to_excel(resp,verbose=3, current_worksheet=sett, workbook=workbook,x0=0,y0=3)
workbook.close()



def write_to_excel(resp, filename='test.xlsx', verbose=10, workbook='', current_worksheet='',new_worksheet='testings',x0=0,y0=0):
    '''Функция для красивого вывода ответа на запрос в файл екселя
    '''
    g = parse_str(resp,0)
    # next(g)
    if workbook=='':
        workbook = xlsxwriter.Workbook(filename)
        need_close = 1
    else:
        need_close =0
        
        
    if current_worksheet=='':
        sett = workbook.add_worksheet(new_worksheet)
    else:
        sett = current_worksheet
    # bold = workbook.add_format({'bold': True})
    border = workbook.add_format()
    border.set_border(style=1)

    border_no_top = workbook.add_format()
    border_no_top.set_bottom()
    border_no_top.set_left()
    border_no_top.set_right()
    border_no_topright = workbook.add_format()
    border_no_topright.set_border(style =3)
    border_no_topright.set_bottom()
    border_no_topright.set_left()
    border_no_left = workbook.add_format()
    border_no_left.set_border(style =3)
    border_no_left.set_bottom()
    border_no_left.set_top()
    border_no_left.set_right()
    border_no_right = workbook.add_format()
    border_no_right.set_border(style =3)
    border_no_right.set_bottom()
    border_no_right.set_top()
    border_no_right.set_left()
    border_no_bottom = workbook.add_format()
    border_no_bottom.set_top()
    border_no_bottom.set_left()
    border_no_bottom.set_right()
    border_no_bottomleft = workbook.add_format()
    border_no_bottomleft.set_border(style =3)
    border_no_bottomleft.set_top()
    border_no_bottomleft.set_right()

    try:
        i=0
        d_prev=0
        s_prev=''
        q=0
        border_prev = ''
        while True:
            s,n,d =next(g)
            d+=y0
            if i==0:
                d_prev=d
                n_prev=n
            if q==0:
                i =x0
                # d+=y0
                d_prev=d
                i_prev=i
            print('n=',n,'| s=',s,' |d=',d)
            if d>d_prev:
                i-=1
                print('-1', s)
            
            if n == n_prev and q>0: #если уровень списка одинаков для текущего и предыдущего
                if d==d_prev: #если уровень словаря одинаков для текущего и предыдущего
                    if border_prev == 'no left': #если для предыдущей клеточки установлена граница "без левой стороны", то устанавливаем "без нижней и левой"
                        sett.write(i_prev, d_prev, str(s_prev),border_no_bottomleft)
                    else: #иначе устанавливаем только  "без нижней"
                        sett.write(i_prev, d_prev, str(s_prev),border_no_bottom)

                    sett.write(i, d, str(s),border_no_top)
                    border_prev = 'no top' #устанавливаем "без верхней " для текущей
                    q+=1
                elif d>d_prev:
                    if border_prev == 'no top': #если для предыдущей клеточки установлена граница "без правой стороны", то устанавливаем "без верхней и правой"
                        sett.write(i_prev, d_prev, str(s_prev),border_no_topright)
                    else: #иначе устанавливаем только "без правой"
                        sett.write(i_prev, d_prev, str(s_prev),border_no_right)
                    sett.write(i, d, str(s),border_no_left)
                    border_prev = 'no left'
                    q+=1
                else:
                    sett.write(i, d, str(s))
                    border_prev = ''
                    q+=1
            else:
                sett.write(i, d, str(s))
                border_prev = ''
                q+=1
            
            d_prev = d
            i_prev = i
            n_prev = n
            s_prev = s
            i+=1
    except Exception as e:
        print(e)
    finally:
        if need_close:
            workbook.close()
    return (workbook,sett,i-x0)

filename = 'setting1.xlsx'

       
def read_settings(filename, sheet_name='settings', diapasone=('A1', 'C40'), validate_function=None, **kwargs ):
    
    wb = load_workbook(filename = filename)
    if sheet_name == '__active':
        sheet = wb.active
    else: 
        sheet = wb.get_sheet_by_name(sheet_name)
    try:
        cells = sheet[diapasone[0]: diapasone[1]]
    except Exception as e:
        printlog('Неверный диапазон значений для парсинга')
    c = [0 for i in range(len(cells) )]
    parsed_settings = {'diapasone' : [], 'verbose_level' : 0, 'url_mw' : {}, 'url_siebel' : {}, 'how to check response' : '', 'exclude or inlude fields': '',
                            'field of response': [] ,'show type of field in response' : '', 'enter after each element of list': '', 'response in borders':'', 'columns' : {}, 'token-mw' :'', 'token-siebel' :''}

    columns_width = 1
    exclude_or_include = 0        
    i = 0
    for *c, in cells:  
        if c[0].value is not None and c[0].value != '':
            columns_width = 1
            if c[1].value is not None and c[1].value != '':
                exclude_or_include = 0        
            elif exclude_or_include == 1:
                parsed_settings['field of response'].append(c[2].value)
        elif columns_width == 1:
            parsed_settings['columns'][c[1].value] = c[2].value            
    
        if c[0].value == 'начальная точка': parsed_settings['diapasone'].append(c[1].value)
        if c[0].value == 'конечная точка': parsed_settings['diapasone'].append(c[1].value)
        if c[0].value == 'уровень раскрытия ответа запроса': parsed_settings['verbose_level'].append(c[1].value)   
        if c[0].value == 'url для токена миддлваре': url_for_token_mv = c[1].value
        if c[0].value == 'генерация токена миддлваре': gen_for_token_mv = c[1].value
        if c[0].value == 'url для токена сибеля':  url_for_token_siebel = c[1].value            
        if c[0].value == 'генерация токена сибеля':  gen_for_token_siebel = c[1].value            
        if c[0].value == 'проверка определения способа подстановки токена': url_for_token_check = c[1].value
        if c[0].value == 'токен миддлваре': 
            if c[1].value == 'из точки':
                parsed_settings['url_mw']['from'] = 'from_point_and_url'            
                parsed_settings['url_mw']['from_point'] = c[2].value            
                parsed_settings['url_mw']['from_url'] = url_for_token_mv           
                parsed_settings['url_mw']['how_gen'] = gen_for_token_mv     
                parsed_settings['url_mw']['how_check_url'] = url_for_token_check     
            if c[1].value == 'взять из клетки':
                parsed_settings['url_mw']['from'] = 'from_cell'            
                parsed_settings['url_mw']['token'] = c[2].value            
            if c[1].value == 'ничего не делать':
                parsed_settings['url_mw']['from'] = 'nothing'            
        if c[0].value == 'токен сибель': 
            if c[1].value == 'из точки':
                parsed_settings['url_siebel']['from'] = 'from_point_and_url'            
                parsed_settings['url_siebel']['from_point'] = c[2].value            
                parsed_settings['url_siebel']['from_url'] = url_for_token_siebel           
                parsed_settings['url_siebel']['how_gen'] = gen_for_token_siebel           
                parsed_settings['url_siebel']['how_check_url'] = url_for_token_check     
            if c[1].value == 'взять из клетки':
                parsed_settings['url_siebel']['from'] = 'from_cell'            
                parsed_settings['url_siebel']['token'] = c[2].value            
            if c[1].value == 'ничего не делать':
                parsed_settings['url_siebel']['from'] = 'nothing'            

        if c[0].value == 'старый токен': parsed_settings['verbose_level'].append(c[1].value)            
        if c[0].value == 'проверка совпадения ответа': 
            if c[1].value == 'простая': 
                parsed_settings['how to check response'].append('simple')            
            if c[1].value == 'сложная': 
                parsed_settings['how to check response'].append(c[1].value)            
                if c[2].value == 'исключить': 
                    parsed_settings['exclude or inlude fields'].append('exclude')            
                    exclude_or_include =1        
                if c[2].value == 'включить': 
                    parsed_settings['exclude or inlude fields'].append('include')    
                    exclude_or_include =1        
        if c[0].value == 'вывод типов полей в ответе': parsed_settings['show type of field in response'] = c[1].value            
        if c[0].value == 'за каждым элементом списка ответов перевод строки ': parsed_settings['enter after each element of list'] = c[1].value            
        if c[0].value == 'ответ в рамочки': parsed_settings['response in borders'] = c[1].value            
        if c[0].value == 'размер столбцов': 
            columns_width = 1
            parsed_settings['columns'][c[1].value] = c[2].value            
            
    return parsed_settings
        


        
def read_data(filename, sheet_name='__active', diapasone=('A1', 'I40'), validate_function=None, **kwargs ):
    
    wb = load_workbook(filename = filename)
    if sheet_name == '__active':
        sheet = wb.active
    else: 
        sheet = wb.get_sheet_by_name(sheet_name)
    try:
        cells = sheet[diapasone[0]: diapasone[1]]
    except Exception as e:
        printlog('Неверный диапазон значений для парсинга')
    c = [0 for i in range(len(cells) )]
    parsed_data = {'points' : [], 'verbs' : {}, 'headers_name' : {}, 'headers_value' : {}, 'url' : {}, 'post_data':{}, 'response' : {}, 'response_code' : {}, 'token' :0}

    i = 0
    for *c, in cells:  
        i +=1
        if i==1:
            continue
        if c[1].value is not None and is_int(c[0].value) :
            parsed_data['points'].append(c[1].value)
            last_point = c[1].value
            parsed_data['verbs'][c[1].value] = []
            parsed_data['verbs'][c[1].value].append(c[2].value)
            parsed_data['headers_name'][c[1].value] = []
            parsed_data['headers_name'][c[1].value].append(c[3].value)
            parsed_data['headers_value'][c[1].value] = {}
            parsed_data['headers_value'][c[1].value][c[3].value] = c[4].value
            parsed_data['url'][c[1].value] = []
            parsed_data['url'][c[1].value].append(c[5].value)
            parsed_data['post_data'][c[1].value] = {}
            parsed_data['post_data'][c[1].value][c[2].value] = c[6].value
            parsed_data['response_code'][c[1].value] = {}
            parsed_data['response_code'][c[1].value][c[2].value] = c[7].value
            parsed_data['response'][c[1].value] = {}
            parsed_data['response'][c[1].value][c[2].value] = c[8].value
            # if validate_function is not None:
                # if not validate_function(parsed_data):
                    # raise Error #ToDo
            
        elif last_point is not None:
            if c[2].value is not None:
                parsed_data['verbs'][last_point].append(c[2].value)
            if c[3].value is not None:
                parsed_data['headers_name'][last_point].append(c[3].value)
                if c[4].value is not None:
                    if  c[4].value == 'Bearer' or c[4].value == 'Bearer ':
                        if new_token == 1:
                            token = fetchToken()
                        parsed_data['headers_value'][last_point][c[3].value] = c[4].value + ' ' + token
                        
                    else:
                        parsed_data['headers_value'][last_point][c[3].value] = c[4].value
            print(1)
            if c[5].value is not None:
                parsed_data['url'][last_point].append(c[5].value)
            print(2)
            if c[6].value is not None:
                parsed_data['post_data'][last_point].append(c[6].value)
            print(3)
            if c[7].value is not None:
                parsed_data['response_code'][last_point][c[2].value] = c[7].value
            print(4)
            if c[8].value is not None:
                parsed_data['response'][last_point][c[2].value] = c[8].value
            print(5)
    return parsed_data
        

parsed_data = read_data(filename)        
parsed_data
filename = 'setting1.xlsx'
r = requests.request(verb, *parsed_data['url'][point], headers=parsed_data['headers_value'][point])
        
parsed_data['points']
parsed_data['verbs'] 
parsed_data['headers_name'] 
parsed_data['headers_value']
parsed_data['url'] 
parsed_data['post_data']
parsed_data['response']
parsed_data['response_code']

token = fetchToken()
point = 'token'
url = parsed_data['url'][point]
url
header = parsed_data['headers_value'][point]
header


r = requests.request(verb, *url, headers=header)
r

r = requests.request(verb, *parsed_data['url'][point], headers=parsed_data['headers_value'][point])
worksheets = []
row = []
simple = 0
ws_name = []
for point in points_new:
    ws_name.append(point)
    worksheets.append(workbook.add_worksheet(point) )
    row.append(0)
parsed_data['headers_value'][point]
import xlsxwriter

write_response_data(parsed_data , filename='response.xlsx')

    

def write_2_excel_parsed_data(parsed_data, filename='response.xlsx', sheet_name='response', **kwargs ):
    workbook = xlsxwriter.Workbook(filename = filename)
    worksheet = workbook.add_worksheet(sheet_name)
    if settings is not None and settings.get('columns_width',0):
        for c, w in settings['columns_width'].items():
            worksheet.set_column(int(c), int(c), int(w))
    counter = 1
    point_counter = 1  
    for point in parsed_data['points']:
        worksheet.write( counter, point_counter, none_safe_str(point))
        verb_counter = counter
        for verb in parsed_data['verbs'][point]:
            worksheet.write( counter, point_counter+5, none_safe_str( parsed_data['post_data'][point].get(verb,'')) )
            worksheet.write( verb_counter, point_counter+1, str(verb))
            try:
                if verb=='post' and parsed_data['post_data'][point].get(verb,''):
                    r = requests.request(verb, *parsed_data['url'][point], headers=parsed_data['headers_value'][point], data=parsed_data['post_data'][point][verb])
                else:
                    r = requests.request(verb, *parsed_data['url'][point], headers=parsed_data['headers_value'][point])
                resp = r.json()
            except Exception as e:
                print(e)
                resp = r.text
                worksheet.write( verb_counter, point_counter+6, str(r) )

            workbook,worksheet,resp_counter = write_to_excel(resp,verbose=3, current_worksheet=worksheet, workbook=workbook,x0=verb_counter,y0=point_counter+7)
            worksheet.write( verb_counter, point_counter+7, none_safe_str(resp) )
            verb_counter += resp_counter+1
        worksheet.write( counter, point_counter+2, *parsed_data['url'][point])
        header_counter = counter
        for h in parsed_data['headers_name'][point]:
            worksheet.write( header_counter, point_counter+3, none_safe_str(h))
            worksheet.write( header_counter, point_counter+4, none_safe_str(parsed_data['headers_value'][point][h]))
            header_counter += 1
        counter = max(counter+1, verb_counter, header_counter) + 1
    workbook.close()
    
    
write_response_data(parsed_data )




q=0     
for c in  x:
    q+=1
    if ifIter(c):
        print(type(c),len(c), q)
        if ifDict(c):
            for k0,v0 in c.items():
                print(k0)
                if q ==4:
                    printDict(v0)
                else:
                    print(v0)        
                    
        elif ifList(c):
            for v0 in c:
                if q ==4:
                    printDict(v0)
                else:
                    print(v0)        
        else:
            print(c)    
        print()
        # print(ifiter)
    else:
        print(type(c))
        print(c)
    print()

log

import os
cwd = os.getcwd()
os.chdir("/path/to/your/folder")
os.listdir('.')
cwd

def prompt(s):
    print(s)
    ret = input()
    return ret
        

def user_interface():
    
    print('Интерфейс для программы тестирования')
    print('''Режимы работы: 
    1. считывание данных из файла,
    2. вывод текущих точек,
    3. вывод подробной информации о конкретной точке
    4. вывод всей информации
    5. тестирование точки''')
    

def init():
    
    s = int(user_interface() )
    
    
    
    
    token = fetchToken()
    # quirks = {'mode': 'change', 'change': 'url', 'change_type': 1, 'start_chr':'0', 'step_chr':'1', 'stop_chr':'150' }
    quirks = {'mode': 'content', 'content': 'header','change_type': 1,'step':10, 'chr':'a', 'content_num': 100}
    verbs = ['get', 'head', 'options', 'patch','connect', 'delete', 'post', 'put', 'trace']
    points = ['token', 'settings', 'links', 'relPhone', 'countMain', 'countDpi', 'bonus', 'offer']
    log = [{}, {}]
    asyncio.run( _requests(verbs, points, token,  debug=2 ))
    write_to_excel('test')
