import win32api
import win32con
import win32gui
import time
time.sleep(5)
#ウインドウハンドルを取得
VoicePeak = win32gui.FindWindow(None, "VOICEPEAK")  #voicepeakのウインドウハンドルを取得
JUCE = win32gui.FindWindow(None,"JUCEWindow")       #voicepeak(juce)のウインドウハンドルを取得
#メッセージを入力
win32gui.SendMessage(JUCE, win32con.WM_ACTIVATE,2,0)    #ウインドウをアクティブ化
win32gui.SendMessage(VoicePeak,win32con.WM_SETFOCUS,0,0)    #テキストエリアにフォーカスする
win32gui.SendMessage(VoicePeak, win32con.WM_CHAR,ord("邪"),0)  #読み上げたい文字を送信
win32gui.SendMessage(VoicePeak,win32con.WM_KILLFOCUS,0,0)   #テキストエリアのフォーカスを外す

