from json.decoder import JSONDecodeError
import aiohttp
import requests, asyncio

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
        
#         if url_d['point']==token:
                    
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

# loop = asyncio.get_event_loop()
# loop.run_until_complete(asyncio.gather(
#     factorial("A", 2),
#     factorial("B", 3),
#     factorial("C", 4),
# ))
# loop.close()




import csv
filename = 'ok.xlsx'
with open(filename, 'wb') as outfile:
    writer = csv.writer(outfile)
    # to get tabs use csv.writer(outfile, dialect='excel-tab')
    writer.writerows(log[0])

from pandas import DataFrame, ExcelWriter

myDF = DataFrame(log[0])
writer = ExcelWriter(filename)
myDF.to_excel(writer)
writer.save()

import xlsxwriter

def write_to_excel(s='', debug=0):
    
    filename = ['ok' +s+ '.xlsx', 'err' +s+ '.xlsx']
    for i in range(2):    
        workbook = xlsxwriter.Workbook(filename[i])
        worksheet = workbook.add_worksheet()
        q=0
        for j in log[i]:
            (a,b,c,d) = j 
            (e,f,g,h) = log[i][j]
            worksheet.write(q,1, str(a))
            worksheet.write(q,2, str(b))
            worksheet.write(q,3, str(c))
            worksheet.write(q,4, str(d))
            worksheet.write(q,5, str(e))
            worksheet.write(q,6, str(f))
            worksheet.write(q,7, str(g))
            worksheet.write(q,8, str(h))

            # print(log[0][i], 'd,e,f', d,e, a)
            q+=1
            log_print(debug, 4, q)
            # if q>20:
            #     break
        workbook.close()

points
dictionary = {0:0}

def print_elem(dic, key):
    if type(dic) == type(dictionary) and dic.get(key, 0):
        return key
    else:
        return key + ' ' + dic[key]

def write2_to_excel(s='', simple = 0, debug=0):
    
    s='test'
    simple = 0
    debug=0
    
    # filename = ['ok' +s+ '.xlsx', 'err' +s+ '.xlsx']
    # for i in range(2):  

    
workbook = xlsxwriter.Workbook('test.xlsx')
worksheets = []
row = []
ws_name = []
for p in points:
    ws_name.append(p)
    worksheets.append(workbook.add_worksheet(p) )
    row.append(0)
    
for j in log[i]:
    temp = [ 0 for i in range(8)] 

    (temp[0], temp[1], temp[2], temp[3]) = j 
    (temp[4], temp[5], temp[6], temp[7]) = log[i][j]
    i = 0
    t = ws_name.index(temp[0])
    for n in range(7):
        worksheets[t].write( row[t], n+1, str(temp[n]))
    
    workbook,worksheet = write_to_excel(resp,verbose=3, current_worksheet=sett, workbook=workbook,x0=0,y0=3)

    row[t]+=1
    # print(log[0][i], 'd,e,f', d,e, a)
    # q+=1
    log_print(debug, 4, q)
    # if q>2:
    #     break
workbook.close()

token = fetchToken()

verbs = ['get', 'post']
point = 'links'
header = header_creator(point,token)
url = url_creator(point)
r = requests.get(url,headers = header)
resp = r.json()
describe = 5
filename = 'setting.xlsx'

example = [1,2,[2,3, [4, {5: {3:4,1:2,5:[6,7]}}, [ {7:8, 9:[10,11] },6 ], 4], 5], {1:2, 2:[3,4, {5:6} ] }, 8]

def parse_str(s,n=0, dic=0, verbose=10):
    if isType(s)<2 or verbose<=0:
        yield (s,n,dic)
    if isType(s)==2 and verbose>0:
        n+=1
        verbose-=1
        for c in s:
            yield from parse_str(c, n,dic,verbose)
        n-=1
        print()
    if isType(s)==3 and verbose>0:
        dic +=1
        verbose-=1
        for c in s:
            yield (c,n,dic-1)
            yield from parse_str(s[c], n, dic,verbose)
        dic -=1
example = [1,2,[12,13, [24, {35: {43:4,41:2,45:[56,57]}}, [ {67:8, 69:[70,11] },56 ], 44], 5], {31:2, 32:[43,4, {55:6} ] }, 28]

workbook = xlsxwriter.Workbook(filename)
sett = workbook.add_worksheet('settings')

write_to_excel(resp, filename,verbose=3, new_worksheet='testings',x0=2,y0=0)
workbook,worksheet,i = write_to_excel(resp,verbose=3, current_worksheet=sett, workbook=workbook,x0=0,y0=3)
workbook.close()



def write_to_excel(resp, filename='test.xlsx', verbose=10, workbook='', current_worksheet='',new_worksheet='testings',x0=0,y0=0):
    '''Функция для красивого вывода ответа на запрос в файл екселя
    '''
    g = parse_str(resp,verbose=verbose)
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

filename
resp
debug = 0
describe = 4
workbook = xlsxwriter.Workbook(filename)
sett = workbook.add_worksheet('settings')
counter = 0
point_counter = 0
for point in points:
    counter += 1
    point_counter += 1
    sett.write(point_counter + 1, 0, counter)
    sett.write(point_counter + 1, 1, point)
    
    heads_counter = 0
    for h,v in header_creator(point, token = token).items():
        sett.write(point_counter + heads_counter + 1, 3, h)
        sett.write(point_counter + heads_counter + 1, 4, v)
        print(h,v,' counter + heads_counter', counter, heads_counter)
        
        heads_counter+=1
        
    sett.write(point_counter + 1, 5, url)
    
    verbs_counter = 0
    for verb in verbs:
        sett.write(point_counter + verbs_counter + 1, 2, verb)
        header = header_creator(point, token = token)
        url = url_creator(point)
        r = requests.request(verb, url, headers = header )
        try:
            resp = r.json()
        except Exception as e:
            print(e)
            resp = r.text
        sett.write(point_counter + verbs_counter + 1, 6, r.status_code)
        # sett.write(i+1+j, 7, str(resp))
        # l = 0
        # if ifLD(resp): #респ - словарь или список
            # sett.write(i+1+j+l, 8, 'list or dict')
        workbook,sett,resp_counter = write_to_excel(resp,verbose=2, current_worksheet=sett, workbook=workbook,x0=point_counter + verbs_counter + 1,y0=8)
        # verbs_counter = writing_structure_in_excel(sett, resp, describe, point_counter=point_counter, verbs_counter=verbs_counter, debug=debug)
        verbs_counter+=resp_counter+1
    point_counter = point_counter  + max(verbs_counter, heads_counter)                        

workbook.close()



def writing_structure_in_excel(sett, resp, verbose, point_counter=0, verbs_counter=0, debug=0):
    # try:
    resp_counter = 0
    sett.write(point_counter + verbs_counter + 1, 7, str(checkArray(resp)) + ' ' + str(type(resp)) )
    
    if verbose==0:
        sett.write(point_counter + verbs_counter + 1, 8, str(resp))
        
    if verbose>=1:
        if isType(resp)<2: #просто строка
            sett.write(point_counter + verbs_counter + 1, 8, str(resp))
        # if checkArray(resp)<3: #список или словарь
        if verbose<3:
            if isType(resp)==2:
                for res in resp:
                    sett.write(point_counter + verbs_counter +resp_counter+ 1, 7, str(checkArray(res)) + ' vv ' + str(type(res)) )
                    sett.write(point_counter + verbs_counter + resp_counter + 1, 8, str(res))
                    log_print(debug, 4, '02')
                    resp_counter +=1
            if isType(resp)==3:
                for res in resp:
                    sett.write(point_counter + verbs_counter +resp_counter+ 1, 7, str(checkArray(res)) + ' vv ' + str(type(res)) )
                    sett.write(point_counter + verbs_counter + resp_counter + 1, 8, str(res))
                    sett.write(point_counter + verbs_counter + resp_counter + 1, 9, str(resp[res]))
                    resp_counter +=1
        if verbose>=3:
            resp_counter = excel_writing_element(sett, resp, verbose, 8, point_counter+ verbs_counter, resp_counter)
            
    if verbose>=4:
                    
        if checkArray(resp)>=3: #список словарей, список списков, словарь списков или словарь словарей
            if isType(resp)==2: #если список, который внутри содержит список или словарь
                                #список словарей, список списков,
                for res in resp:
                            
                    if isType(res)==2: #если список списков, который внутри содержит список или словарь
                        #resp list, res list, res1 res2
                        for res1 in res:
                            resp_counter = excel_writing_element(sett, res1, verbose, 8, point_counter+ verbs_counter, resp_counter)

                            
                    if isType(res)==3: #если список словарей, который внутри содержит список или словарь
                            #resp list, res dict, res1 res2
                        for res1 in res:
                            sett.write(point_counter + verbs_counter + resp_counter + 1, 8, str(res1)) #вывести ключи элементов resp[res]
                            resp_counter = excel_writing_element(sett, res[res1], verbose, 9, point_counter+ verbs_counter, resp_counter)


            if isType(resp)==3:#если словарь, который внутри содержит список или словарь
                                        #список словарей, список списков,
                for res in resp:
                    sett.write(point_counter + verbs_counter +resp_counter+ 1, 7, str(checkArray(res)) + ' ' + str(type(res)) )
                    sett.write(point_counter + verbs_counter + resp_counter + 1, 8, str(res)) #вывести ключи элементов resp[res]
                            
                    if isType(resp[res])==2: #если словарь списков, который внутри содержит список или словарь
                        #resp dict, res list, res1 res2
                        for res1 in resp[res]:
                            
                            resp_counter = excel_writing_element(sett, res1, verbose, 9, point_counter+ verbs_counter, resp_counter)
                            
                    if isType(resp[res])==3: #если словарь словарей, который внутри содержит список или словарь
                            #resp dict, res dict, res1 res2
                        for res1 in resp[res]:
                            sett.write(point_counter + verbs_counter + resp_counter + 1, 9, str(res1)) #вывести ключи элементов resp[res]

                            resp_counter = excel_writing_element(sett, resp[res][res1], verbose, 10, point_counter + verbs_counter, resp_counter)
    # except Exception as e:
    #     print(e)
    verbs_counter += resp_counter+1
    # print('verbose', verbose)
    return verbs_counter


def excel_writing_element(sett, res1, verbose, index, counters, resp_counter):
    try:
        log_print(debug, 4, '10')
            
        if isType(res1)<2: #если ни список, ни словарь
            log_print(debug, 4, '11')
            
            sett.write(counters + resp_counter + 1, index, str(res1))
            log_print(debug, 4, '12')
            #resp dict, res list, res1 str
    except Exception as e:
        print('|', res1, '|')
        print(e) 
    if verbose<3:
        sett.write(counters + resp_counter + 1, index, str(res1))
    else:
            
        if isType(res1)==2: #если список 
            #resp dict, res list, res1 list, res2
            try:
                for res2 in res1:
                    if verbose>4:
                        sett.write(counters + resp_counter + 1, index, str(res2))
                        resp_counter = verbose_writing(sett, res2, index, counters, resp_counter)
                    else:
                        sett.write(counters + resp_counter + 1, index, str(res2))
                        resp_counter +=1
            except Exception as e:
                sett.write(counters + resp_counter + 1, index, str(res1))
                print(' excel_writing_element', e)
                
        if isType(res1)==3: #если словарь 
            #resp dict, res list, res1 dict, res2
            try:
                for res2 in res1:
                    sett.write(counters +resp_counter+ 1, 7, str(checkArray(res1[res2])) + ' vvv ' + str(type(res1[res2]))+ str(len(res1[res2])) )
                    
                    if verbose>4:
                        sett.write(counters + resp_counter + 1, index, str(res2))
                        resp_counter = verbose_writing(sett, res1[res2], index, counters, resp_counter)
                    else:
                        
                        sett.write(counters + resp_counter + 1, index, str(res2))
                        sett.write(counters + resp_counter + 1, index+1, str(res1[res2]))
                        resp_counter +=1
            except Exception as e:
                sett.write(counters + resp_counter + 1, index, str(res1))
                print(' excel_writing_element', e)
    resp_counter +=1
    return resp_counter

def verbose_writing(sett, res2, index, counters, resp_counter):
    if checkArray(res2)<3:
        sett.write(counters + resp_counter + 1, index+1, str(res2))
    if checkArray(res2)==3:
        for lis in res2:
            for k,v in lis.items():
                sett.write(counters + resp_counter + 1, index+1, str(k))
                sett.write(counters + resp_counter + 1, index+2, str(v))
                resp_counter +=1
    if checkArray(res2)==4:
        for lis in res2:
            for k in lis:
                sett.write(counters + resp_counter + 1, index+1, str(k))
                resp_counter +=1
    if checkArray(res2)==5:
        for k,lis in res2.items():
            sett.write(counters + resp_counter + 1, index+1, str(k))
            for v in lis:
                sett.write(counters + resp_counter + 1, index+2, str(v))
                resp_counter +=1
    if checkArray(res2)==3:
        for kd, dic in res2.items():
            sett.write(counters + resp_counter + 1, index+1, str(kd))
            for k,v in dic.items():
                sett.write(counters + resp_counter + 1, index+2, str(k))
                sett.write(counters + resp_counter + 1, index+3, str(v))
                resp_counter +=1
    return resp_counter



q = 3.21
type(q)
token = fetchToken()
header = header_creator(point, token = token)
url = url_creator(point)
r = requests.request(verb, url, headers = header )
r
r.json()
resp = r.json()



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
        

def deconv(el):
    try:
        print('log0')
        print(type(el), isType(el), len(el))
        print(el)
        if checkArray(el):
            print('log1')
            if isType(el)==3:
                print('log2', type(el), el.keys(), 'log2')
                yield el.keys()
                c = 0
                print('log2.5')
                for e in el:
                    print('log3')
                    if not c:
                        if checkArray(el[e]):
                            print('log4')
                            # print(el[e])
                            g = deconv(el[e])
                            yield from next(g)
                            
                    c += 1
        if checkArray(el):

            if isType(el)>=2:
                print('log11')
                for e in el:
                    print('log12')
                    print(type(e), checkArray(e))
                    print(e)
                    if checkArray(e):
                        print('log13')
                        g = deconv(e)
                        yield from next(g)
                    else:
                        yield  e
        if not checkArray(el):
            print('log20')
            yield ('str', el)

    # if isType(el)==
    except Exception as e:
        print('log9')
        print(e)

def unlist(el):
    for e in el:
        pass
def dict2list(dic):
    l = []
    for d in dic:
        l.append(dic[d])
    return l
      

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

def deconv2(el ,i):
    print('__0')
    if checkArray(el)==3 or checkArray(el)==4:
        for e in el:
            # g = deconv2(el)
            i = i+1
            print('__0.5',i)
            yield from deconv2(el, i)
    print('__1')

    if checkArray(el)>=5:
        for e in el:
            yield el.keys()
            # g = deconv2(el[e])
            i = i+1
            print('__1.5',i)
            if i<10:
                yield from next(deconv2(el[e]), i)
            else:
                yield el
                
    print('__2')

    if checkArray(el)<3:
        yield el.keys()
    print('__3')
    
def deconv3(el, l):
    if isType(el)==3:
        l.append( list(el) )
        deconv3( dict2list(el), l)
    elif isType(el)==2:
        if checkArray(el)>=2:

          
d = {'key': ' AnimGetGift2_2', 'groups': ['MYVF-WIDGET'], 'link': 'Неправильно!'}        
list(d)
dict2list(d)

gen = deconv2(resp, 0)

while True:
    # try:
    l = next(gen)
    # except Exception as e:
        # print(e)
    # print('s=', s)
    print('l=', l)
    for i in l:
        print(i)



def gen():
    yield from x, y

next(z)    
z= gen()


        
filename = 'setting1.xlsx'

from openpyxl import load_workbook

wb = load_workbook(filename = filename)
sheet = wb.active

cells = sheet['A1': 'M40']
cells.values
c = [0 for i in range(9)]
type(c)
t = (c[0], c[1], c[2])

points_new = []
verbs_new = {}
headers_name = {}
headers_value = {}
url_new = {}
data_new = {}
response_new = {}
response_code = {}
new_token =0

i = 0
for *c, in cells:  
    # print(c[0].value, c[1].value, c[2].value, c[3].value, c[4].value, c[5].value, c[6].value, c[7].value)
    i +=1
    if i==1:
        continue
    if c[1].value is not None:
        points_new.append(c[1].value)
        last_point = c[1].value
        verbs_new[c[1].value] = []
        verbs_new[c[1].value].append(c[2].value) #verbs[point]
        headers_name[c[1].value] = []
        headers_name[c[1].value].append(c[3].value) #headers[point]
        headers_value[c[1].value] = {}
        headers_value[c[1].value][c[3].value] = c[4].value #headers_value[point][header_name]
        url_new[c[1].value] = []
        url_new[c[1].value].append(c[5].value) #url[point]
        if c[6].value is not None:
            data_new[c[1].value] = []
            data_new[c[1].value].append(c[6].value) #data[point]
            print('data type', type(c[6].value))
        response_code[c[1].value] = {}
        response_code[c[1].value][c[2].value] = c[7].value  #response_code[point][verb]
        response_new[c[1].value] = {}
        response_new[c[1].value][c[2].value] = c[8].value   #response_code[point][verb]
        
    else:
        if c[2].value is not None:
            verbs_new[last_point].append(c[2].value)
        if c[3].value is not None:
            headers_name[last_point].append(c[3].value)
            if c[4].value is not None:
                if  c[4].value == 'Bearer' or c[4].value == 'Bearer ':
                    if new_token == 1:
                        token = fetchToken()
                    headers_value[last_point][c[3].value] = c[4].value + ' ' + token
                    
                else:
                    headers_value[last_point][c[3].value] = c[4].value
                    
        if c[5].value is not None:
            url_new[last_point].append(c[5].value)
        if c[6].value is not None:
            data_new[c[1].value].append(c[6].value)
            print('data type', type(c[6].value))
        if c[7].value is not None:
            response_code[last_point][c[2].value] = c[6].value
        if c[8].value is not None:
            response_new[last_point][c[2].value] = c[7].value

points_new
verbs_new 
headers_name 
headers_value
url_new 
response_new
response_code
data_new
token = fetchToken()

worksheets = []
row = []
simple_worksheets = 0
ws_name = [] #worksheets name
if simple_worksheets:
    for point in points_new:
        ws_name.append(point)
        worksheets.append(workbook.add_worksheet(point) )
        row.append(0)

workbook = xlsxwriter.Workbook('response.xlsx')
worksheet = workbook.add_worksheet('response')
i = 1
j = 1    
print('ok1')
for point in points_new:
    worksheet.write( i, j, point)
    i1 = i
    print('ok2')
    for verb in verbs_new[point]:
        print('ok21')
        
        print(verb)
        
        worksheet.write( i1, j+1, verb)
        worksheet.write( i1, j+6, response_code[point].get(verb, 0) )
        worksheet.write( i1, j+7, response_new[point].get(verb, 0) )
        print('ok22')
        
        
        i1 = i1 + 1
    print('ok23')
    worksheet.write( i, j+2, *url_new[point])
    print('ok24')
    if data_new.get(point,0)
        worksheet.write( i, j+5, *data_new.get(point,'') )
    
    i2 = i
    print('ok3')
    for h in headers_name[point]:
        print(h)
        worksheet.write( i2, j+3, h)
        worksheet.write( i2, j+4, headers_value[point][h])
        i2 = i2 + 1
    print('ok4')
    i3 = i
  
    print('i, i+2, i1,i2,i3', i, i+2, i1,i2,i3)
    i = max(i+1, i1, i2, i3) + 1
workbook.close()

type(response_new['links']['get'])

if ifLD(response_new['links']['get']):
    for i in response_new['links']:
        print(i)     

def ifLD(ld):
    if ifDict(ld) or ifList(ld):
        return 1
    return 0
ok =0 
for point in points_new:
    for verb in  verbs_new[point]:
        # print('0')
        # print(verb, url_new[point], headers_value[point])
        r = requests.request(verb, *url_new[point], headers =  headers_value[point])
        # print('1')
        print(r.status_code)
        # print('2')
        if r.text == response_new[point]:
            print('ok')
            ok +=1
        else:
            print()
            print()
            response_new[point]
            print(r.text)
        # requests.re
ok
u = ['https://mw-tst.itsmartflex.com/uaa/oauth/token?grant_type=client_credentials']

h = {'Accept-Language': 'ru', 'Content-Type': 'application/json', 'Authorization': 'Basic aW50ZXJuYWw6aW50ZXJuYWw='}        
req = requests.request('GET', *u, headers = h)


import requests
import xlrd, xlwt
import pandas as pd

file = 'something.xlsx'

# Load spreadsheet
xl = pd.ExcelFile(file)
print(xl.sheet_names)
df1 = xl.parse('Sheet12')
df = pd.read_excel(file)

df.get('token',0)
for d in df:
    print(d) 
    
log[0]
df0 = pd.DataFrame(log[0][0])
df1 = pd.DataFrame(log[1])
df = pd.DataFrame(log)
df
df.to_excel('something.xlsx')
df0.to_excel('something.xlsx')
df1.to_excel('something.xlsx')


q=0
for i in log[0]:
    print()
    print(i)
    # print(log[0][i])
    d={1:2}
    s=[0,0,0,0]
    (s[0], s[1], s[2], s[3] ) = (log[0][i])

    for st in s:
        if type(st)!=type(q):
            print(type(st), len(st))
        if type(st)==type(s):
            for t in st:
                if type(t)==type(s):
                    for y in t:
                        print(y)
                # print(*st)
                elif type(t)==type(d):
                    for y in t:
                        print(y, t[y])
                else:
                    print(t)
            # print(*st)
        elif type(st)==type(d):
            for t in st:
                
                if type(t)==type(s):
                    for y in t:
                        print(y)
                # print(*st)
                elif type(t)==type(d):
                    for y in t:
                        print(y, t[y])
                else:
                    print(t, st[t])
            # print()
        else:
            print(st)

    if q>5:
        break
    q+=1
    
    
q=0
for i in log[0]:
    print(i)
    q+=1
    if q>5:
        break

x =log[0][('settings', 'https://mw-tst.itsmartflex.com/entity/api/functions/MYVF-SETTINGS', 'put', 2)]

i=0
s='q'

def dict_reader(dic, x):
    if len(dic)>0:
        for d in dic:
            print(dic[d])
            if ifDict(dic[d]):
                print('dict', dict_reader(dic[d], x))
            if ifList(dic[d]):
                print('list', list_reader(dic[d],x))
            print(d)
            print()
        return x
    return 0
    
def list_reader(lis, x):
    x+=1
    if len(lis)>0:
        for l in lis:
            if ifDict(l):
                print('dict', dict_reader(l,x ))
            if ifList(l):
                print('list', list_reader(l, x))
            print(l)
            print()
        return x
    return 0

           
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
