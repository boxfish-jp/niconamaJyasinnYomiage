import win32api
import win32con
import win32gui
import threading
import time
import subprocess
import requests
import json
import websockets
import asyncio
import re
from bs4 import BeautifulSoup

flag = True
speaking = False

proc = subprocess.Popen(["Release\injectioncode.exe"],stdout=subprocess.PIPE)
pipeMess = ["NotChange"]

# 邪神ちゃんが読み上げ中かどうかを判別するために実行するプログラムの立ち上げとログの記録
def Startlogging():
    global pipeMess
    for line in iter(proc.stdout.readline,b''):
        if (not flag) :
           break
        out = line.decode('utf-8')
        pipeMess.append(out.replace('\r\n', ''))

# 記録されたログの一番最後のメッセージを返す関数
def checkLog():
    return pipeMess[-1]

# 邪神ちゃんが読み上げが終わるまで、終了しない(待つ)関数
def nowPlaying():
    if (pipeMess[-1] == "WM_PAINT"):
        while(1):
            logSum = len(pipeMess)
            time.sleep(0.3)
            if (len(pipeMess) == logSum): #0.3秒経ってもログが更新されなかったら、関数を終了
                return
    else:
        time.sleep(0.5)
        nowPlaying()

#コメントの中にurlが含まれているかを判別(邪神ちゃんがurlを読み始めると長いので対策)
def replaceUrls(text):
    # URLを検出する正規表現パターン
    urlPattern = re.compile(r"https?://\S+")

    # 文字列中のURLを置換
    replacedText = re.sub(urlPattern, " url省略", text)

    return replacedText

def sendVoicePeak(message):
    global speaking
    speaking = True
    message = replaceUrls(message)
    #ウインドウハンドルを取得
    VoicePeak = win32gui.FindWindow(None, "VOICEPEAK")  #voicepeakのウインドウハンドルを取得
    JUCE = win32gui.FindWindow(None,"JUCEWindow")       #voicepeak(juce)のウインドウハンドルを取得
    time.sleep(0.05)
    #メッセージを入力
    win32gui.SendMessage(JUCE, win32con.WM_ACTIVATE,2,0)    #ウインドウをアクティブ化
    win32gui.SendMessage(VoicePeak,win32con.WM_SETFOCUS,0,0)    #テキストエリアにフォーカスする
    for s in message:
        print(s);
        win32gui.SendMessage(VoicePeak, win32con.WM_CHAR,ord(s),0)  #読み上げたい文字を送信
        time.sleep(0.05)
    win32gui.SendMessage(VoicePeak,win32con.WM_KILLFOCUS,0,0)   #テキストエリアのフォーカスを外す

    time.sleep(0.5)
    #再生
    win32gui.SendMessage(VoicePeak,win32con.WM_CHAR,32,0)   #スペースを送信して、読み上げを実行
    nowPlaying();

    #入力した文字を削除
    win32gui.SendMessage(VoicePeak,win32con.WM_SETFOCUS,0,0)    #テキストエリアにフォーカスする
    for s in message:
        win32gui.SendMessage(VoicePeak,win32con.WM_KEYDOWN,0x27,0)  #右矢印キーを入力
        win32gui.SendMessage(VoicePeak,win32con.WM_KEYDOWN,8,0)     #バックスペースを入力
        time.sleep(0.05)
    win32gui.SendMessage(VoicePeak,win32con.WM_KILLFOCUS,0,0)
    speaking = False

def setupVoicePeak():
    VoicePeak = win32gui.FindWindow(None, "VOICEPEAK")  #voicepeakのウインドウハンドルを取得
    rect = win32gui.GetWindowRect(VoicePeak)
    nowPos = win32api.GetCursorPos()
    win32api.SetCursorPos((rect[0]+161, rect[1]+167))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, rect[0]+161, rect[1]+167, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, rect[0]+161, rect[1]+167, 0, 0)
    win32api.SetCursorPos(nowPos)
    sendVoicePeak("起動しました")

def nicoCommentGet():
    ### コメントを取得したい放送のURLを指定
    live_id = "co5043209"
    #live_id = "co3000390" # コミュニティIDを指定すると放送中のものを取ってきてくれる
    url = "https://live2.nicovideo.jp/watch/"+live_id

    ### htmlを取ってきてWebSocket接続のための情報を取得
    response = requests.get(url)
    soup = BeautifulSoup(response.text, "html.parser") 
    embedded_data = json.loads(soup.find('script', id='embedded-data')["data-props"])
    urlSystem = embedded_data["site"]["relive"]["webSocketUrl"]

    ### websocketでセッションに送るメッセージ
    message_system_1 = {"type":"startWatching",
                        "data":{"stream":{"quality":"abr",
                                        "protocol":"hls",
                                        "latency":"low",
                                        "chasePlay":False},
                                "room":{"protocol":"webSocket",
                                        "commentable":True},
                                "reconnect":False}}
    message_system_2 ={"type":"getAkashic",
                    "data":{"chasePlay":False}}
    message_system_1 = json.dumps(message_system_1)
    message_system_2 = json.dumps(message_system_2)

    ### コメントセッション用のグローバル変数
    uri_comment = None
    message_comment = None

    ### 視聴セッションとのWebSocket接続関数
    async def connect_WebSocket_system(urlSystem):
        global uri_comment
        global message_comment

        ### 視聴セッションとのWebSocket接続を開始
        async with websockets.connect(urlSystem) as websocket:

            ### 最初のメッセージを送信
            await websocket.send(message_system_1)
            await websocket.send(message_system_2) # これ送らなくても動いてしまう？？
            print("SENT TO THE SYSTEM SERVER: ",message_system_1)
            print("SENT TO THE SYSTEM SERVER: ",message_system_2)

            ### 視聴セッションとのWebSocket接続中ずっと実行
            while True:
                message = await websocket.recv()
                message_dic = json.loads(message)
                print("RESPONSE FROM THE SYSTEM SERVER: ",message_dic)

                ### コメントセッションへ接続するために必要な情報が送られてきたら抽出してグローバル変数へ代入
                if(message_dic["type"]=="room"):
                    uri_comment = message_dic["data"]["messageServer"]["uri"]
                    threadID = message_dic["data"]["threadId"]
                    message_comment = [{"ping": {"content": "rs:0"}},
                                        {"ping": {"content": "ps:0"}},
                                        {"thread": {"thread": threadID,
                                                    "version": "20061206",
                                                    "user_id": "guest",
                                                    "res_from": -150,
                                                    "with_global": 1,
                                                    "scores": 1,
                                                    "nicoru": 0}},
                                        {"ping": {"content": "pf:0"}},
                                        {"ping": {"content": "rf:0"}}]
                    message_comment = json.dumps(message_comment)

                ### pingが送られてきたらpongとkeepseatを送り、視聴権を獲得し続ける
                if(message_dic["type"]=="ping"):
                    pong = json.dumps({"type":"pong"})
                    keepSeat = json.dumps({"type":"keepSeat"})
                    sendVoicePeak(" ") #空白(読み上げられない)を定期的に送ると過疎放送でも安定して読み上げてくれる
                    await websocket.send(pong)
                    await websocket.send(keepSeat)

    ### コメントセッションとのWebSocket接続関数
    async def connect_WebSocket_comment():
        loop = asyncio.get_event_loop()

        global uri_comment
        global message_comment

        ### 視聴セッションがグローバル変数に代入するまで1秒待つ
        await loop.run_in_executor(None, time.sleep, 1)

        ### コメントセッションとのWebSocket接続を開始
        async with websockets.connect(uri_comment) as websocket:

                ### 最初のメッセージを送信
            await websocket.send(message_comment)
            print("SENT TO THE COMMENT SERVER: ",message_comment)

            ### コメントセッションとのWebSocket接続中ずっと実行
            while True:
                message = await websocket.recv()
                message_dic = json.loads(message)
                if "chat" in message_dic:
                    print(message_dic["chat"]["content"])
                    if (start):
                        while(speaking):
                            pass
                        sendVoicePeak(message_dic["chat"]["content"])
                    else:
                        print("false")
                else:
                    print("RESPONSE FROM THE COMMENT SERVER: ",message_dic)

    
    
    asyncio.new_event_loop().run_in_executor(None,wait)
    ### asyncioを用いて上で定義した2つのWebSocket実行関数を並列に実行する
    loop = asyncio.get_event_loop()
    gather = asyncio.gather(
    connect_WebSocket_system(urlSystem),
    connect_WebSocket_comment(),
    )
    loop.run_until_complete(gather)
    
start = False
def wait():
    time.sleep(5)
    while(speaking):
        pass
    sendVoicePeak("今からコメントを読み上げます")
    global start
    start = True

def main():
    time.sleep(5)   #邪神ちゃんが起動するまで待機
    print("wait")
    preLog = ""     #邪神ちゃんのウインドウのログが入る
    while(1):
        nowLog = checkLog()
        # 新しいログが届いたら、出力
        if (preLog != nowLog):
            preLog = nowLog
            print(nowLog)
            # UPLINKというログ(Voicepeakのウインドウメッセージが取得できるようになったことを表す)がきたら処理を抜ける
            if (nowLog == "UPLINK"):
                print("start")
                break
    setupVoicePeak()
    nicoCommentGet()

log = threading.Thread(target=Startlogging) #別スレッドで並列処理させる
log.start()
main()


