# [Python]Pythonでカレンダー同期するときのタイムゾーン完全理解

投稿日 2026年03月23日
112 views112 ビュー

## 1. はじめに（問題提起）
OutlookやGoogleカレンダーは、日常的に利用するカレンダーとして非常に便利です。
Pythonを使ってExcelと連携させることで、自動で予定を登録・更新することも可能ですが、
「OutlookとGoogleカレンダーで時間がずれてしまった」という経験はありませんか？

特に海外のチームや複数デバイスで予定を扱う場合、タイムゾーンの扱いを誤ると大きなズレが生じます。
本記事では、Pythonを利用したExcel連携で起こりやすい「時間ズレ」の原因と、正しい処理方法を解説します。

## 2. 結論
OutlookとGoogleカレンダーでは、日時の扱い方の前提が異なるためズレが発生します。

Outlook（Win32COM）：ローカルタイム（PCのタイムゾーン）前提
Googleカレンダー API：UTCまたは明示されたタイムゾーンを厳密に扱う

## 3. 詳細解説
### ■ Outlookの場合
Python から Outlook（Win32COM）に予定を送るとき、datetime が タイムゾーン情報なし（naive） の場合、Outlook は PCのローカルタイム（例：JST） として解釈します。
そのため、Excel → Python → Outlook の連携では、タイムゾーン補正が必要になることがあります。

処理の流れ（例）:

① Excel の datetime（タイムゾーンなし / native）

↓

② Python の datetime（native）

↓

③ Outlook が「これはローカルタイム（JST）だ」と解釈


対策例（Python）:
```
import win32com.client
from datetime import datetime
import pytz

# Excelから取得した naive datetime
excel_dt = datetime(2026, 3, 23, 14, 0)

# JST タイムゾーンを付与
jst = pytz.timezone("Asia/Tokyo")
aware_dt = jst.localize(excel_dt)

# Outlook 予定作成
outlook = win32com.client.Dispatch("Outlook.Application")
appt = outlook.CreateItem(1)  # 1=予定アイテム
appt.Start = aware_dt.strftime("%Y-%m-%d %H:%M:%S")
appt.Subject = "サンプル予定"
appt.Save()
```

### ■ Googleカレンダーの場合
Googleカレンダー API では、datetime は ISO8601形式（タイムゾーン付き） で
送る必要があります。

送った日時はそのまま解釈されるため、naive datetime を送ると UTC とみなされる場合があります。

明示的にタイムゾーンを指定するのが安全です。


処理の流れ（例）:

① Python の datetime（タイムゾーン付き / aware, JST）

↓

② ISO8601 に変換（例: 2026-03-23T14:00:00+09:00）

↓

③ Googleカレンダーが正しく JST として解釈


対策例（Python）:
```
from googleapiclient.discovery import build
from datetime import datetime
import pytz

# JSTで日時を作成
jst = pytz.timezone("Asia/Tokyo")
start_dt = jst.localize(datetime(2026, 3, 23, 14, 0))
end_dt = jst.localize(datetime(2026, 3, 23, 15, 0))

event = {
    "summary": "サンプル予定",
    "start": {"dateTime": start_dt.isoformat(), "timeZone": "Asia/Tokyo"},
    "end": {"dateTime": end_dt.isoformat(), "timeZone": "Asia/Tokyo"},
}

service = build("calendar", "v3", credentials=credentials)
service.events().insert(calendarId="primary", body=event).execute()
```

Googleカレンダーではタイムゾーンを明示すれば、補正は不要です。

## 4. まとめ
Outlook は ローカルタイム前提、Googleカレンダーは UTC or 明示タイムゾーン前提
Python でカレンダー連携を行う場合は、この仕様の違いを理解した上で、naive/aware datetime を正しく扱うことが重要です。

タイムゾーンを誤ると、予定が1時間ずれる、日付が変わるなどのトラブルにつながります。

## 5. 参考・実用
今回の検証をもとに、Excel と Outlook を同期するツールを作成しました。

GUI 操作で簡単に同期可能
EXE形式でそのまま利用可能

→ 興味のある方はこちらから：
Outlook版 https://ikhm-design.booth.pm/items/8106311
Google版 https://ikhm-design.booth.pm/items/8064800
(ご覧いただけると嬉しいです)
###  この記事はQiitaにも掲載しています： https://qiita.com/ikhm-design/items/10dbfbd5c44920b7179c
