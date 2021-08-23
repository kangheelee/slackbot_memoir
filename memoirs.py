import os
import json
import requests
from pandas import json_normalize
import pandas as pd
from datetime import datetime, date, time, timedelta
import time
import openpyxl

# 환경 변수로 슬랙 토큰을 입력 후 사용해주세요.
# Counting Bot
token = 'xoxb-2094939655075-2183607541286-h5VD3ThopzJ70IFvPpfvdOr6'
headers = {"Authorization": 'Bearer ' + token}
# Count
#token = 'xoxb-1675602897633-1875733837124-yUsjP44m7wrrizYbFl0DiTzO'


# 커넥션 에러 뜰 경우에만 사용
#headers = {"user-agent": "크롬 개발자 도구에서 찾으시오."} 

english_table = {"Ireh RYU" : "류이레", "Sujin Kim" : "김수진", "Chimin Ahn" : "안치민", "Jeongyeon Kim" : "김정연", "Daewon Noh" : "노대원", "Ryoo Seojin" : "류서진", "Jueun Lee" : "이주은", "Suryeon Kim" : "김수련", "editor gieun" : "김기은", "Leon Firenze Leem" : "임레온", "KANG HEE LEE" : "이강희", "Rebecca Choi" : "최혜림", "L" : "이유진B", "Ed Chanwoo Kim" : "김찬우", "Soyoung Kim" : "김소영"}

def filter_channel(channel_list, filter:str = '토요일'):
    channels = list()
    for id in channel_list['name']:
        if channel_list['name'][id].find(filter) != -1:
            channels.append(channel_list['name'][id])
    return channels

def get_all_channel():
    URL = 'https://slack.com/api/conversations.list'
    # 파라미터
    params = {
        'Content-Type': 'application/x-www-form-urlencoded',
        'types': 'public_channel, private_channel'
        }

    # API 호출
    res = requests.get(URL, params = params, headers=headers)
    channel_list = json_normalize(res.json()['channels'])

    return channel_list

def find_channel(channel_name:str = '1_공지사항'):
    
    channel_list = get_all_channel()
    channel_id = list(channel_list.loc[channel_list['name'] == channel_name, 'id'])[0]
    return channel_id

# 함수 : get all messages
def get_all_messages(channel:str, start_time:str='0', end_time:str=time.time()):
    URL = 'https://slack.com/api/conversations.history'
    # 파라미터
    params = {
        'Content-Type': 'application/x-www-form-urlencoded',
        'channel' : channel,
        'oldest' : start_time,
        'latest' : end_time
            }
    res = requests.get(URL, params = params, headers=headers)
    conversations = pd.DataFrame(columns = ['ts' , 'user', 'text', 'type'])
    conversations = pd.concat([conversations, json_normalize(res.json()['messages'])], ignore_index=True)
    
    return conversations[['ts','user','text','type']]

# 함수 : user id -> user nickname
# Bug 누락되는 이름 발생
def changetonick(user_id:str):
    URL = 'https://slack.com/api/users.info'
        # 파라미터
    params = {
        'Content-Type': 'application/x-www-form-urlencoded',
        'user' : user_id
            }
    res = requests.get(URL, params = params, headers=headers)
    try:
        user_nick = list(json_normalize(res.json())['user.profile.real_name'])[0]
        return user_nick
    except:
        pass
    return

# 함수 : members
def get_members(channelid):
    URL = 'https://slack.com/api/conversations.members'
        # 파라미터
    params = {
        'Content-Type': 'application/x-www-form-urlencoded',
        'channel' : channelid
        }
    res = requests.get(URL, params = params, headers=headers)
    mem_list = list(json_normalize(res.json())['members'])
    return mem_list[0]

# 함수 : ts - dt
def todatetime(ts:str):
    date_time = datetime.fromtimestamp(float(ts)).strftime('%Y-%m-%d %H:%M')
    return date_time

# 함수 : download excel file
def down_excel(dataframe, title):
    title = '5_output/' + title + '.xlsx'
    dataframe.to_excel(title, sheet_name = 'sheet1')

def load_excel(title):
    title = '5_output/' + title + '.xlsx'
    dataframe = pd.read_excel(title, index_col = 0)
    return dataframe
    
def merge_excel(term_length):
    #dataframe 
    merge_df = load_excel(str(1) + '주차 회고여부')
    for term in range(term_length):
        if term > 1:
            df = load_excel(str(term) + '주차 회고여부')
            merge_df = pd.concat([merge_df, df], axis = 1)
    down_excel(merge_df, '회고체크'+ '_'+ str(term_length-1))

# 함수 : make dataframe
def make_data(channel, oldest, latest):
    df1 = get_all_messages(find_channel(channel), oldest, latest)
    col_ts = pd.DataFrame([todatetime(x) for x in df1['ts']], columns = ['date'])
    col_user = pd.DataFrame([changetonick(y) for y in df1['user']], columns = ['user'])
    del df1['ts']
    del df1['user']
    df1 = pd.concat([col_ts, col_user, df1], axis=1)
    return df1

def filter_completed(df):
    users = list()
    data = df.to_dict()
    for id in data['text']:
        if str(data['text'][id]).find('회고록') != -1 :
            users.append(data['user'][id])
        else :
            if len(data['text'][id]) > 300:    
                if str(data['text'][id]).find('반갑습니다') == -1 and str(data['text'][id]).find('안녕하세요') == -1 :
                    users.append(data['user'][id])
                
    return users

def filter_uncompleted(total_df):
    reasons = dict()
    users = list()
    data = total_df.to_dict()
    for id in data['text']:
        if str(data['text'][id]).find('회고록') != -1 :
            #users.append(data['user'][id])
            continue
        else :
            if len(data['text'][id]) > 300:
                if str(data['text'][id]).find('반갑습니다') == -1 and str(data['text'][id]).find('안녕하세요') == -1 :
                    continue
             #       users.append(data['user'][id])
                else:
                    reasons[data['user'][id]] = '자기소개'
            else:
                reasons[data['user'][id]] = 'under_300'
                
    return reasons

def filter_members(members, filters):
    real_members = []
    
    for member in members:
        is_real = True
        for filter in filters:
            if member.find(filter) != -1:
                is_real = False
        if is_real:
            real_members.append(member)
    return real_members

def eng_to_kor(name):
    for key in english_table:
        if key == name:
            name = english_table[key]
    return name

def count(oldest, latest, late_oldest, late_latest, term):
    print(str(term) + '주차')
    # 필요한 값 : 찾으려는 채널명, oldest, latest, 엑셀 파일 저장명
    
    all_members = []
    
    #### find channels ####
    channel_list = get_all_channel().to_dict()
    off_channel_list = filter_channel(channel_list, '회고록_offline')
    on_channel_list = filter_channel(channel_list, '회고록_online')
    share_channel_list = filter_channel(channel_list, '회고록_shareonly')
    channels = off_channel_list + on_channel_list + share_channel_list
    
    #### make DataFrame ####
    df = pd.DataFrame(columns = ['date' , 'user', 'text', 'type'])
    late_df = pd.DataFrame(columns = ['date' , 'user', 'text', 'type'])
    total_df = pd.DataFrame(columns = ['date' , 'user', 'text', 'type'])
    #### get members, data ####
    for i in range(len(channels)):
        all_members.extend(get_members(find_channel(channels[i])))
        # make_data : get the data from server
        df = pd.concat([df, make_data(channels[i], oldest, latest)], ignore_index=True)
        late_df = pd.concat([late_df, make_data(channels[i], late_oldest, late_latest)], ignore_index=True)
        total_df = pd.concat([total_df, make_data(channels[i], oldest, late_latest)],  ignore_index=True)

    #### get members nick ####
    filters = ['메모어L', '메모어', '메모어R', '이동건', '박세훈', '김상엽', 'Counting Bot', 'FlaskBot', 'Count',  's1375811068', '전수빈']
    all_members = list(set(all_members))
    all_members_nick = [changetonick(member) for member in all_members]
    all_members_nick = [eng_to_kor(member) for member in all_members_nick]
    # filter 운영진, 봇
    all_members_nick = filter_members(all_members_nick, filters)
    
    #### 회고 정상 제출 ####
    users = filter_completed(df)
    user_completed = list(set(users))
    user_completed = filter_members(user_completed, filters)
    user_completed = [eng_to_kor(member) for member in user_completed]
    
    #### 회고 지각 제출 ####
    late_users = filter_completed(late_df)
    late_user_completed = list(set(late_users))
    late_user_completed = filter_members(late_user_completed, filters)
    late_user_completed = [eng_to_kor(member) for member in late_user_completed]

    #### 회고 미제출 ####
    reasons = filter_uncompleted(total_df)
    #print('reasons : ', reasons)
    
    user_uncompleted = [j for j in all_members_nick if j not in user_completed]
    user_uncompleted = [j for j in user_uncompleted if j not in late_user_completed]
    #print('user_uncompleted : ', user_uncompleted)
    fin_df = pd.DataFrame(columns = all_members_nick)
    
    fin_df.loc['회고록' + str(term) + '회'] = 'O'
    for _ in late_user_completed:
        fin_df[_]['회고록' + str(term) + '회'] = 'L'    
    for _ in user_uncompleted:
        if reasons.get(_):
            fin_df[_]['회고록' + str(term) + '회'] = 'X' + reasons[_]
        else:
            fin_df[_]['회고록' + str(term) + '회'] = 'X'
    fin_df = fin_df.transpose()
    
    # 엑셀파일로 저장
    down_excel(fin_df,str(term) + '주차 회고여부')

def find_time(oldest, latest, interval, term_length):
    oldests, latests = list(), list()
    oldests.append(oldest)
    latests.append(latest)

    for i in range(term_length-1):
        oldest = oldest + timedelta(days=interval)
        oldests.append(oldest)
        latest = latest + timedelta(days=interval)
        latests.append(latest)
    return oldests, latests
    
if __name__ == "__main__":

    oldests, latests, late_oldest, late_latest = list(), list(), list(), list()
    oldest = datetime(2021, 6, 17, minute = 0)
    latest = datetime(2021, 6, 21, minute = 10)
    late_oldest = latest
    late_latest = datetime(2021, 6, 24, minute = 0)
     
    term_length = 13
    oldests, latests = find_time(oldest, latest, interval = 7, term_length = term_length)
    late_oldests, late_latests = find_time(late_oldest, late_latest, interval = 7, term_length = term_length)
    
    current_term = 9
    i = 0
    
    for oldest, latest, late_oldest, late_latest in zip(oldests, latests, late_oldests, late_latests):
        i = i + 1
        if i < current_term:
            continue
        oldest = time.mktime(oldest.timetuple())
        latest = time.mktime(latest.timetuple())
        late_oldest = time.mktime(late_oldest.timetuple())
        late_latest = time.mktime(late_latest.timetuple())
        count(oldest, latest, late_oldest, late_latest, i)
        if i == current_term:
            break
    
    merge_excel(current_term+1)    

    