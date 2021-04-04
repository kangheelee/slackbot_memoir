import os
import json
import requests
from pandas import json_normalize
import pandas as pd
from datetime import datetime, date, time, timedelta
import time
import openpyxl

# 환경 변수로 슬랙 토큰을 입력 후 사용해주세요.
# export SLACK_BOT_TOKEN='xoxb-bla-bla'
#token = os.environ["SLACK_BOT_TOKEN"]ㅣ
#token = 'xoxb-1675602897633-1854874536133-YL2aYkdADMLeIhiLNPTD81yi'
token = 'xoxb-1498144326948-1811806285184-FO3eE7gJzsTeAUIDtU9h31Ji'
# 커넥션 에러 뜰 경우에만 사용
headers = {"user-agent": "크롬 개발자 도구에서 찾으시오."}

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
        'token': token,
        'types': 'public_channel, private_channel'
            }

    # API 호출
    res = requests.get(URL, params = params)
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
        'token': token,
        'channel' : channel,
        'oldest' : start_time,
        'latest' : end_time
            }
    res = requests.get(URL, params = params)
    conversations = json_normalize(res.json()['messages'])
    return conversations[['ts','user','text','type','reply_users']]

# 함수 : user id -> user nickname
# Bug 누락되는 이름 발생
def changetonick(user_id:str):
    URL = 'https://slack.com/api/users.info'
        # 파라미터
    params = {
        'Content-Type': 'application/x-www-form-urlencoded',
        'token': token,
        'user' : user_id
            }
    res = requests.get(URL, params = params)
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
        'token': token,
        'channel' : channelid
            }
    res = requests.get(URL, params = params)
    mem_list = list(json_normalize(res.json())['members'])
    return mem_list[0]

# 함수 : ts - dt
def todatetime(ts:str):
    date_time = datetime.fromtimestamp(float(ts)).strftime('%Y-%m-%d %H:%M')
    return date_time

# 함수 : download excel file
def down_excel(dataframe, title):
    title = title + '.xlsx'
    dataframe.to_excel(title, sheet_name = 'sheet1')

def load_excel(title):
    title = title + '.xlsx'
    dataframe = pd.read_excel(title, index_col = 0)
    return dataframe

def merge_excel(term_length, title):
    #dataframe 
    merge_df = load_excel(title + str(1))
    for term in range(term_length):
        if term > 1:
            df = load_excel(title + str(term))
            merge_df = pd.concat([merge_df, df])
    down_excel(merge_df.transpose(), title)

# 함수 : make dataframe
def make_data(channel):
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
        if len(data['text'][id]) >300:
            users.append(data['user'][id])
    return users


def filter_members(members):
    real_members = []
    filters = ['운영진', '메모어', '운영진B', '이동건', '박세훈', '김상엽', 'FlaskBot', 'Counting Bot']
    for member in members:
        is_real = True
        for filter in filters:
            if member.find(filter) != -1:
                is_real = False
        if is_real:
            real_members.append(member)
    print('real_members : ', real_members)
    return real_members

def count(oldest, latest, term):

    # 자동화 시작
    all_members = []
    # label
    df = pd.DataFrame(columns = ['date' , 'user', 'text', 'type', 'reply_users'])
    
    channel_list = get_all_channel().to_dict()
#    sat_channel_list = filter_channel(channel_list, '토요일')
#    sun_channel_list = filter_channel(channel_list, '일요일')
#    share_channel_list = filter_channel(channel_list, 'shareonly')
#    channels = sat_channel_list + sun_channel_list + share_channel_listx`
    channels = filter_channel(channel_list, '온라인')

    for i in range(len(channels)):
        all_members.extend(get_members(find_channel(channels[i])))
        df = pd.concat([df, make_data(channels[i])], ignore_index=True)
        # make_data : preprocessing data
    all_members = list(set(all_members))
    all_members_nick = [changetonick(z) for z in all_members]
    all_members_nick = filter_members(all_members_nick)
    fin_df = pd.DataFrame(columns = all_members_nick, index = ['댓글 횟수'])
    reply_list = list(df['reply_users'])
    print('reply_list : ', reply_list)
    index0 = list(fin_df.columns)
    index1 = [0 for m in range(len(index0))]
    print('index1 : ', index1)
    print('index0 : ', index0)
    for q in reply_list:
        if type(q) == list:
            for p in q:
                if all_members_nick.count(changetonick(p)) != 0:
                    index1[index0.index(changetonick(p))] += 1
    
    fin_df.loc['댓글 횟수'] = index1
    reply_check = []
    for num in index1:
        reply_check.append('O' if num >= 8 else 'X')
    fin_df.loc['댓글 여부'] = reply_check
    down_excel(fin_df,'reply_count' + str(term))
    
    # 엑셀파일로 저장
    df = pd.DataFrame(columns = all_members_nick, index = ['댓글'+str(term)])
    df.loc['댓글'+str(term)] = reply_check
    down_excel(df,'reply' + str(term))

def find_time(oldest, latest, interval, length):
    oldests, latests = list(), list()
    oldests.append(oldest)
    latests.append(latest)

    for i in range(length-1):
        oldest = oldest + timedelta(days=interval)
        oldests.append(oldest)
        latest = latest + timedelta(days=interval)
        latests.append(latest)
    return oldests, latests
    
if __name__ == "__main__":

    oldests, latests = list(), list()
    oldest = datetime(2020, 12, 4)
    latest = datetime(2021, 1, 4, minute = 10)
    
    term_length = 3
    oldests, latests = find_time(oldest, latest, interval = 28, length = term_length)
    i = 0
    for oldest, latest in zip(oldests, latests):
 
        i = i + 1
        oldest = time.mktime(oldest.timetuple())
        latest = time.mktime(latest.timetuple())
        count(oldest, latest, i)
    title = 'reply_count'
    merge_excel(term_length+1, title)
    title = 'reply'
    merge_excel(term_length+1, title)