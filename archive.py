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
token = 'xoxb-1675602897633-1854874536133-myMT8wKX6R9l61nDfpWWA6Hh'
headers = {"Authorization": 'Bearer ' + token}
# 커넥션 에러 뜰 경우에만 사용
#headers = {"user-agent": "크롬 개발자 도구에서 찾으시오."} 

english_table = {"Ireh RYU" : "류이레", "Sujin Kim" : "김수진", "Chimin Ahn" : "안치민", "Jeongyeon Kim" : "김정연", "Daewon Noh" : "노대원", "Ryoo Seojin" : "류서진", "Jueun Lee" : "이주은", "Suryeon Kim" : "김수련", "editor gieun" : "김기은", "Leon Firenze Leem" : "임레온", "KANG HEE LEE" : "이강희", "Rebecca Choi" : "최혜림", "L" : "이유진B", "Ed Chanwoo Kim" : "김찬우", "Soyoung Kim" : "김소영" }
kor_table = {"류이레" : "Ireh RYE", "김수진" : "Sujin Kim" , "안치민" : "Chimin Ahn" , "김정연" : "Jeongyeon Kim" , "노대원" : "Daewon Noh", "류서진" : "Ryoo Seojin", "이주은" : "Jueun Lee", "김수련" : "Suryeon Kim" , "김기은" : "editor gieun", "임레온" : "Leon Firenze Leem", "이강희" : "KANG HEE LEE", "최혜림" : "Rebecca Choi", "이유진B" : "L", "김찬우" : "Ed Chanwoo Kim", "김소영" : "Soyoung Kim"}
# 함수 : find channel id

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
    conversations = json_normalize(res.json()['messages'])
    #return conversations[['ts','user','text','type','reply_users']]
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
    title = 'output/' + title + '.xlsx'
    dataframe.to_excel(title, sheet_name = 'sheet1')

def load_excel(title):
    title = 'output/' + title + '.xlsx'
    dataframe = pd.read_excel(title, index_col = 0)
    return dataframe
    
def merge_excel(term_length):
    #dataframe 
    merge_df = load_excel(str(1) + '주차 아카이빙')
    for term in range(term_length):
        if term > 1:
            df = load_excel(str(term) + '주차 아카이빙')
            merge_df = pd.concat([merge_df, df], axis = 1)
    down_excel(merge_df, 'archiving'+ '_' + str(term_length-1))
    
# 함수 : make dataframe
def make_data(channel, oldest, latest):
    df1 = get_all_messages(find_channel(channel), oldest, latest)
    colts = pd.DataFrame([todatetime(x) for x in df1['ts']], columns = ['date'])
    coluser = pd.DataFrame([changetonick(y) for y in df1['user']], columns = ['user'])
    del df1['ts']
    del df1['user']
    df1 = pd.concat([colts, coluser, df1], axis=1)
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

def filter_archived(df, users, filters):
    archives = dict()
    for user in users:
        archives[user] = list()        
    data = df.to_dict()
    for id in data['text']:
        if data['user'][id] not in filters:
            user = eng_to_kor(data['user'][id])
            if str(data['text'][id]).find('회고록') != -1 :
                if len(archives[user]) == 1:
                    archives[user].insert(0, data['text'][id])
                else: 
                    archives[user].append(data['text'][id])
            else :
                if len(data['text'][id]) > 300 and str(data['text'][id]).find('반갑습니다') == -1 and str(data['text'][id]).find('안녕하세요') == -1 :
                    if len(archives[user]) == 1:
                        archives[user].insert(0, data['text'][id])
                    else: 
                        archives[user].append(data['text'][id])
 
    for user in users:
        if len(archives[user]) == 0:
            archives[user].append('X')
    return archives

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

def update_late_submission(user_list, term, archives):
    df = load_excel(str(term-1) + '주차 아카이빙')
    for user in user_list:
        content = str(df['회고록'+str(term-1)+'회'][user])
        if content == 'X':
            df['회고록'+str(term-1)+'회'][user] = 'L' + archives[user][0]
        else :
            df_before = load_excel(str(term-2) + '주차 아카이빙')
            before_content = str(df_before['회고록'+str(term-2)+'회'][user])
            if before_content == 'X':
                df_before['회고록'+str(term-2)+'회'][user] = 'L' + df['회고록'+str(term-1)+'회'][user]
                df['회고록'+str(term-1)+'회'][user] = 'L' + archives[user][0]
                down_excel(df_before, str(term-2) + '주차 아카이빙')
    down_excel(df, str(term-1) + '주차 아카이빙')

def find_late_submission(df, term, users):
    if term == 13:
        return list(set(users))
    late_users = list()
    for user in users:
        if users.count(user) > 1:
            late_users.append(user)
    
    return list(set(late_users))

def update_archive_df(archive_df, archives, users):
    for user in users:
        if len(archives[user]) > 1:
            archive_df[user] = archives[user][1]
        else:
            archive_df[user] = archives[user][0]
    return archive_df

def eng_to_kor(name):
    for key in english_table:
        if key == name:
            name = english_table[key]
    return name

def kor_to_eng(name):
    for key in kor_table:
        if key == name:
            name = kor_table[key]
    return name

def archive(oldest, latest, term):
    print(str(term) + '주차')
    # 자동화 시작
    all_members = []
    # label
    #df = pd.DataFrame(columns = ['date' , 'user', 'text', 'type', 'reply_users'])
    df = pd.DataFrame(columns = ['date' , 'user', 'text', 'type'])

    channel_list = get_all_channel().to_dict()
    sat_channel_list = filter_channel(channel_list, '토요일')
    sun_channel_list = filter_channel(channel_list, '일요일')
    share_channel_list = filter_channel(channel_list, 'shareonly')
    channels = sat_channel_list + sun_channel_list + share_channel_list
     
    for i in range(len(channels)):
        all_members.extend(get_members(find_channel(channels[i])))
        df = pd.concat([df, make_data(channels[i], oldest, latest)], ignore_index=True)
        # make_data : preprocessing data

    filters = ['Count', '운영진', '메모어', '운영진B', '이동건', '박세훈', '김상엽', 'FlaskBot', 'Counting Bot', 's1375811068']
    
    all_members = list(set(all_members))
    # 기간 동안 회고 여부와 댓글 수
    all_members_nick = [changetonick(member) for member in all_members]
    all_members_nick = [eng_to_kor(member) for member in all_members_nick]
    all_members_nick = filter_members(all_members_nick, filters)

    users = filter_completed(df)
    users = [eng_to_kor(member) for member in users]
    archives = filter_archived(df, all_members_nick, filters)


    archive_df = pd.DataFrame(columns = all_members_nick)
    
    archive_df.loc['회고록' + str(term) + '회'] = 'X'
    archive_df = update_archive_df(archive_df, archives, all_members_nick)
    archive_df = archive_df.transpose()

    # 엑셀파일로 저장
    down_excel(archive_df, str(term) + '주차 아카이빙')
    
    late_users = find_late_submission(df, term, users)    
    late_users = filter_members(late_users, filters)
    late_users = [eng_to_kor(member) for member in late_users]
    if len(late_users) > 0 and term > 1:
        update_late_submission(late_users, term, archives)

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

    oldests, latests = list(), list()
    oldest = datetime(2021, 3, 8, minute = 10)
    latest = datetime(2021, 3, 15, minute = 10)
     
    term_length = 13
    oldests, latests = find_time(oldest, latest, interval = 7, term_length = term_length)
    i = 0
    current_term = 8
    for oldest, latest in zip(oldests, latests):
        i = i + 1
        if i < current_term:
            continue
        oldest = time.mktime(oldest.timetuple())
        latest = time.mktime(latest.timetuple())
        
        archive(oldest, latest, i)
        if i == current_term:
            break
    
    merge_excel(current_term + 1)    

    