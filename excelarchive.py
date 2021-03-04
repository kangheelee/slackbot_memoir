import os
import json
import requests
from pandas import json_normalize
import pandas as pd
from datetime import datetime
import time
import openpyxl

# 환경 변수로 슬랙 토큰을 입력 후 사용해주세요.
# export SLACK_BOT_TOKEN='xoxb-bla-bla'
token = os.environ["SLACK_BOT_TOKEN"]

# 커넥션 에러 뜰 경우에만 사용
headers = {"user-agent": "크롬 개발자 도구에서 찾으시오."}

# 함수 : find channel id
def channelfind(channel_name:str = '1_공지사항'):
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
    channel_id = list(channel_list.loc[channel_list['name'] == channel_name, 'id'])[0]
    return channel_id

# 함수 : get all messages
def get_all_messages(channel:str, oldest:str='0', latest:str=time.time()):
    URL = 'https://slack.com/api/conversations.history'
    # 파라미터
    params = {
        'Content-Type': 'application/x-www-form-urlencoded',
        'token': token,
        'channel' : channel,
        'oldest' : oldest,
        'latest' : latest
            }
    res = requests.get(URL, params = params)
    conversations = json_normalize(res.json()['messages'])
    return conversations[['ts','user','text','type','reply_users']]

# 함수 : user id -> user nickname
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
        user_nick = list(json_normalize(res.json())['user.profile.display_name'])[0]
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

# 함수 : make dataframe
def make_data(channel):
    df1 = get_all_messages(channelfind(channel), oldest, latest) 
    colts = pd.DataFrame([todatetime(x) for x in df1['ts']], columns = ['date'])
    coluser = pd.DataFrame([changetonick(y) for y in df1['user']], columns = ['user'])
    del df1['ts']
    del df1['user']
    df1 = pd.concat([colts, coluser, df1], axis=1)
    return df1


if __name__ == "__main__":
    # 필요한 값 : 찾으려는 채널명, oldest, latest, 엑셀 파일 저장명
    channels = list(map(str, input("찾고 싶은 채널명을 모두 입력해주세요(띄어쓰기로 구분)\n").split()))
    oldest = input("시작 날짜를 알려주세요! 예) 2021-03-01 00:00 \n입력하지 않을 경우 전체기간입니다!\n")
    latest = input("종료 날짜를 알려주세요! 예) 2021-03-01 00:00 \n입력하지 않을 경우 전체기간입니다!\n")
    file_name = input("저장할 엑셀 파일명을 지정해주세요! 예) 메모어_4기_아카이빙\n")

    # 기간 입력 안했을 경우 디폴트값
    if oldest == '':
        oldest = '0'
    if latest == '':
        latest = time.time()

    # 자동화 시작
    all_members = []
    df = pd.DataFrame(columns = ['date' , 'user', 'text', 'type', 'reply_users'])
    for i in range(len(channels)):
        all_members.extend(get_members(channelfind(channels[i])))
        df = pd.concat([df, make_data(channels[i])], ignore_index=True)
    all_members = list(set(all_members))

    # 기간 동안 회고 여부와 댓글 수
    all_members_nick = [ changetonick(z) for z in all_members]
    user_completed = list(set(df['user']))
    user_uncompleted = [j for j in all_members_nick if j not in user_completed]
    fin_df = pd.DataFrame(columns = all_members_nick, index=['회고 여부'])
    fin_df.loc['회고 여부'] = 'O'

    for _ in user_uncompleted:
        fin_df[_]['회고 여부'] = 'X'
    
    reply_list = list(df['reply_users'])
    index0 = list(fin_df.columns)
    index1 = [0 for m in range(len(index0))]

    for q in reply_list:
        if type(q) == list:
            for p in q:
                index1[index0.index(changetonick(p))] += 1

    fin_df.loc['댓글 횟수'] = index1

    # 엑셀파일로 저장
    down_excel(df, file_name)
    down_excel(fin_df,'금주회고여부')