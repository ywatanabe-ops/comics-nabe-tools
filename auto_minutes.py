"""
Zoomフォルダ監視 → 自動議事録生成 → Notion保存
"""

import os
import re
import sys
import time
import json
import shutil
import tempfile
import requests
import anthropic
sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")
from base64 import b64encode
from datetime import datetime, timedelta, timezone
from pathlib import Path
from dotenv import load_dotenv
load_dotenv(Path(__file__).parent / ".env")
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# ===== 設定 =====
ZOOM_FOLDER = Path(os.path.expanduser("~")) / "OneDrive" / "ドキュメント" / "Zoom"
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
CLAUDE_API_KEY = os.environ.get("CLAUDE_API_KEY", "")
NOTION_API_KEY = os.environ.get("NOTION_API_KEY", "")
NOTION_DB_ID   = os.environ.get("NOTION_DB_ID", "")
NOTION_PROXY   = "https://notion-proxy.y-watanabe-502.workers.dev"

ZOOM_ACCOUNT_ID  = os.environ.get("ZOOM_ACCOUNT_ID", "")
ZOOM_CLIENT_ID   = os.environ.get("ZOOM_CLIENT_ID", "")
ZOOM_CLIENT_SECRET = os.environ.get("ZOOM_CLIENT_SECRET", "")
ZOOM_USER_ID     = os.environ.get("ZOOM_USER_ID", "me")

PROCESSED_FILE  = Path(__file__).parent / ".processed_files.json"
CLIENT_SECRET   = Path(__file__).parent / "client_secret.json"
TOKEN_FILE      = Path(__file__).parent / "token.json"
CALENDAR_SCOPES = ["https://www.googleapis.com/auth/calendar.readonly"]


def load_processed():
    if PROCESSED_FILE.exists():
        return set(json.loads(PROCESSED_FILE.read_text()))
    return set()

def save_processed(processed: set):
    PROCESSED_FILE.write_text(json.dumps(list(processed)))


# ===== Zoom クラウド録画 API =====

ZOOM_TOKEN_URL   = "https://zoom.us/oauth/token"
ZOOM_API_BASE    = "https://api.zoom.us/v2"
PREFERRED_TYPES  = [
    "shared_screen_with_speaker_view",
    "active_speaker",
    "gallery_view",
    "shared_screen_with_gallery_view",
    "shared_screen",
]

def zoom_get_token() -> str:
    """Server-to-Server OAuth でアクセストークンを取得する。"""
    creds = b64encode(f"{ZOOM_CLIENT_ID}:{ZOOM_CLIENT_SECRET}".encode()).decode()
    resp = requests.post(
        ZOOM_TOKEN_URL,
        params={"grant_type": "account_credentials", "account_id": ZOOM_ACCOUNT_ID},
        headers={"Authorization": f"Basic {creds}"},
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()["access_token"]

def zoom_list_recordings(token: str, days: int = 1) -> list:
    """指定日数分のクラウド録画一覧を取得する。"""
    now = datetime.now(timezone.utc)
    from_date = (now - timedelta(days=days)).strftime("%Y-%m-%d")
    to_date   = now.strftime("%Y-%m-%d")
    headers   = {"Authorization": f"Bearer {token}"}
    meetings  = []
    next_page = None
    while True:
        params = {"from": from_date, "to": to_date, "page_size": 30}
        if next_page:
            params["next_page_token"] = next_page
        resp = requests.get(f"{ZOOM_API_BASE}/users/{ZOOM_USER_ID}/recordings",
                            headers=headers, params=params, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        meetings.extend(data.get("meetings", []))
        next_page = data.get("next_page_token")
        if not next_page:
            break
    return meetings

def zoom_pick_mp4(recording_files: list) -> dict | None:
    """最適な MP4 ファイルを選択する。"""
    mp4s = [f for f in recording_files
            if f.get("file_type") == "MP4"
            and f.get("download_url")
            and f.get("status") == "completed"]
    for ptype in PREFERRED_TYPES:
        for f in mp4s:
            if f.get("recording_type") == ptype:
                return f
    return mp4s[0] if mp4s else None

def zoom_download(token: str, download_url: str, dest: Path) -> None:
    """録画ファイルをダウンロードする。"""
    url = f"{download_url}?access_token={token}"
    resp = requests.get(url, stream=True, timeout=600)
    resp.raise_for_status()
    total = int(resp.headers.get("content-length", 0))
    done  = 0
    with open(dest, "wb") as f:
        for chunk in resp.iter_content(chunk_size=1024 * 1024):
            if chunk:
                f.write(chunk)
                done += len(chunk)
                if total:
                    print(f"\r  {done//(1024*1024)}MB / {total//(1024*1024)}MB ({done/total*100:.0f}%)", end="", flush=True)
    print()

def zoom_poll(days: int = 1, dry_run: bool = False, reprocess: bool = False) -> None:
    """Zoom クラウド録画を確認して未処理のものを議事録化する。"""
    if not all([ZOOM_ACCOUNT_ID, ZOOM_CLIENT_ID, ZOOM_CLIENT_SECRET]):
        print("[ERROR] .env に ZOOM_ACCOUNT_ID / ZOOM_CLIENT_ID / ZOOM_CLIENT_SECRET を設定してください。")
        return

    print("[Zoom] APIに接続中...")
    token = zoom_get_token()
    print("[Zoom] 認証成功")

    meetings = zoom_list_recordings(token, days)
    print(f"[Zoom] {len(meetings)}件の会議を取得")

    processed = load_processed()
    count = 0

    for meeting in meetings:
        uid = f"cloud:{meeting.get('uuid', meeting.get('id', ''))}"
        if uid in processed and not reprocess:
            print(f"[Zoom] スキップ（処理済み）: {meeting.get('topic', '')}")
            continue

        best = zoom_pick_mp4(meeting.get("recording_files", []))
        if not best:
            print(f"[Zoom] MP4なし → スキップ: {meeting.get('topic', '')}")
            continue

        topic      = meeting.get("topic", "会議")
        start_str  = meeting.get("start_time", "")
        try:
            dt = datetime.fromisoformat(start_str.replace("Z", "+00:00"))
        except Exception:
            dt = datetime.now(timezone.utc)

        jst  = dt.astimezone(timezone(timedelta(hours=9)))
        date_str = f"{jst.year}-{jst.month:02d}-{jst.day:02d}"
        time_str = f"{jst.hour:02d}.{jst.minute:02d}.00"
        # parse_folder_name が解釈できる形式のフォルダ名を組み立てる
        folder_name = f"{date_str} {time_str}{topic}"

        print(f"\n[Zoom] 処理: {topic} ({date_str} {jst.hour:02d}:{jst.minute:02d})")
        if dry_run:
            print("  [dry-run] スキップ")
            continue

        tmp_dir = Path(tempfile.mkdtemp())
        try:
            mp4_path = tmp_dir / f"{meeting.get('id', 'rec')}.mp4"
            print("[Zoom] ダウンロード中...")
            zoom_download(token, best["download_url"], mp4_path)
            print(f"[Zoom] ダウンロード完了 ({mp4_path.stat().st_size//(1024*1024)}MB)")

            # フォルダ名から会議情報を組み立てる（既存関数を流用）
            info = parse_folder_name(folder_name)
            process_file_from_info(mp4_path, info)

            processed.add(uid)
            save_processed(processed)
            count += 1
        except Exception as e:
            print(f"[ERROR] {e}")
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)

    print(f"\n[Zoom] 完了: {count}件処理しました。")


# ===== フォルダ名から会議情報を解析 =====
def parse_folder_name(folder_name: str) -> dict:
    """
    形式A（ローカル録画）: "2026-04-09 13.59.53【定例】顧問契約／株式会社ウェルフォート"
    形式B（クラウド録画DL）: "【定例】顧問契約／株式会社ドミニオン　5月 11日 (月曜日)⋅午前11:30～午後12:00"
    """
    info = {"title": folder_name, "date": "", "time": "", "other": "", "our": "弊社"}

    # 形式A: 先頭が "YYYY-MM-DD HH.MM.SS"
    m = re.match(r"(\d{4}-\d{2}-\d{2})\s+(\d{2})\.(\d{2})\.\d{2}", folder_name)
    if m:
        info["date"] = m.group(1)
        info["time"] = f"{m.group(2)}:{m.group(3)}"
        title_match = re.search(r"(【.+】.+)", folder_name)
        if title_match:
            info["title"] = title_match.group(1)
    else:
        # 形式B: "【...】...　5月 11日 (月曜日)⋅午前11:30～午後12:00"
        m2 = re.search(r"(\d+)月\s*(\d+)日", folder_name)
        if m2:
            month = int(m2.group(1))
            day   = int(m2.group(2))
            info["date"] = f"{datetime.now().year}-{month:02d}-{day:02d}"

        m3 = re.search(r"(午前|午後)(\d+):(\d+)", folder_name)
        if m3:
            ampm   = m3.group(1)
            hour   = int(m3.group(2))
            minute = int(m3.group(3))
            if ampm == "午後" and hour != 12:
                hour += 12
            elif ampm == "午前" and hour == 12:
                hour = 0
            info["time"] = f"{hour:02d}:{minute:02d}"

        # タイトル = 日付部分より前
        title_part = re.sub(r"[　\s]+\d+月.*$", "", folder_name).strip()
        if title_part:
            info["title"] = title_part

    # 顧客名（「／」以降）
    if "／" in info["title"]:
        info["other"] = info["title"].split("／")[-1].strip()

    return info


# ===== Googleカレンダー認証 =====
def get_calendar_service():
    if not CLIENT_SECRET.exists():
        return None
    creds = None
    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), CALENDAR_SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(CLIENT_SECRET), CALENDAR_SCOPES)
            creds = flow.run_local_server(port=0)
        TOKEN_FILE.write_text(creds.to_json())
    return build("calendar", "v3", credentials=creds)


# ===== 次回MTG取得 =====
def fetch_next_meeting(title: str) -> dict:
    empty = {"datetime": "未定", "zoomUrl": "未定", "meetingId": "未定", "passcode": "未定"}
    try:
        service = get_calendar_service()
        if not service:
            print("[Calendar] client_secret.json が見つかりません。スキップします。")
            return empty

        search_query = title.split("／")[-1].strip() if "／" in title else title

        # 今日の翌日0時以降で検索
        from datetime import timedelta
        tomorrow = datetime.now(timezone.utc) + timedelta(days=1)
        tomorrow = tomorrow.replace(hour=0, minute=0, second=0, microsecond=0)
        time_min = tomorrow.isoformat()

        print(f"[Calendar] 検索クエリ: '{search_query}' / timeMin: {time_min}")

        # 顧客名で検索して直近1件だけ取得
        result = service.events().list(
            calendarId="primary",
            q=search_query,
            timeMin=time_min,
            singleEvents=True,
            orderBy="startTime",
            maxResults=10
        ).execute()
        items = result.get("items", [])
        print(f"[Calendar] 検索結果: {[e.get('summary') for e in items]}")
        # 社内MTGは除外
        items = [e for e in items if "社内" not in (e.get("summary") or "")]
        if not items:
            print("[Calendar] 次回の予定が見つかりませんでした")
            return empty
        event = items[0]

        start_dt = event["start"].get("dateTime") or event["start"].get("date")
        end_dt   = event["end"].get("dateTime")   or event["end"].get("date")
        s = datetime.fromisoformat(start_dt)
        e = datetime.fromisoformat(end_dt)
        days = ["月", "火", "水", "木", "金", "土", "日"]
        dt_str = f"{s.month}/{s.day}（{days[s.weekday()]}）"
        if "T" in start_dt:
            dt_str += f" {s.strftime('%H:%M')}～{e.strftime('%H:%M')}"

        desc = (event.get("description") or "") + " " + (event.get("location") or "")
        zoom = parse_zoom_from_text(desc)
        print(f"[Calendar] 取得成功: {event.get('summary')} → {dt_str}")
        print(f"[Calendar] ZoomURL: {zoom['zoomUrl']} / ID: {zoom['meetingId']} / Pass: {zoom['passcode']}")
        return {"datetime": dt_str, **zoom}
    except Exception as ex:
        print(f"[Calendar] エラー: {ex}")
        return empty


def fetch_meeting_from_calendar_by_time(file_time: datetime) -> dict:
    """ファイル保存時刻の直前の会議をGoogleカレンダーから検索する。"""
    empty = {"title": "", "date": "", "time": "", "other": "", "our": "弊社"}
    try:
        service = get_calendar_service()
        if not service:
            print("[Calendar] client_secret.json が見つかりません。スキップします。")
            return empty

        if file_time.tzinfo is None:
            file_time = file_time.replace(tzinfo=timezone.utc)
        time_min = (file_time - timedelta(hours=12)).isoformat()
        time_max = (file_time + timedelta(minutes=30)).isoformat()
        jst_time = file_time.astimezone(timezone(timedelta(hours=9)))
        print(f"[Calendar] {jst_time.strftime('%Y-%m-%d %H:%M')} の直前会議を検索中...")

        result = service.events().list(
            calendarId="primary",
            timeMin=time_min,
            timeMax=time_max,
            singleEvents=True,
            orderBy="startTime",
        ).execute()

        candidates = []
        for event in result.get("items", []):
            start_str = event["start"].get("dateTime") or event["start"].get("date")
            try:
                start = datetime.fromisoformat(start_str)
                if start.tzinfo is None:
                    start = start.replace(tzinfo=timezone(timedelta(hours=9)))
                if start <= file_time:
                    candidates.append((event, start))
            except Exception:
                continue

        if not candidates:
            print("[Calendar] 直前の会議が見つかりませんでした")
            return empty

        candidates.sort(key=lambda x: x[1], reverse=True)
        event, start = candidates[0]
        summary = event.get("summary", "会議")
        jst = start.astimezone(timezone(timedelta(hours=9)))
        date_str = f"{jst.year}-{jst.month:02d}-{jst.day:02d}"
        time_str = jst.strftime("%H:%M")
        other = summary.split("／")[-1].strip() if "／" in summary else ""
        print(f"[Calendar] 会議を特定: {summary} ({date_str} {time_str})")
        return {"title": summary, "date": date_str, "time": time_str, "other": other, "our": "弊社"}
    except Exception as e:
        print(f"[Calendar] 時刻検索エラー: {e}")
        return empty


def parse_zoom_from_text(text: str) -> dict:
    url_match = re.search(r"https://[a-z0-9.-]*zoom\.us/j/[^\s<\"&\n]+", text)
    zoom_url = url_match.group(0) if url_match else "未定"
    id_match   = re.search(r"(?:ミーティングID[：:\s]+|Meeting ID[：:\s]+)([0-9 ]+)", text, re.I)
    meeting_id = id_match.group(1).strip() if id_match else "未定"
    pass_match = re.search(r"(?:パスコード[：:\s]+|パスワード[：:\s]+|Passcode[：:\s]+|Password[：:\s]+)([A-Za-z0-9]+)", text, re.I)
    passcode   = pass_match.group(1).strip() if pass_match else "未定"
    return {"zoomUrl": zoom_url, "meetingId": meeting_id, "passcode": passcode}


# ===== MIMEタイプ判定 =====
MIME_MAP = {
    ".mp4": "video/mp4",
    ".m4a": "audio/mp4",
    ".mp3": "audio/mpeg",
    ".wav": "audio/wav",
    ".m4v": "video/mp4",
}

# ===== VTT解析 =====
def parse_vtt(text: str) -> str:
    lines = text.split("\n")
    result = []
    for line in lines:
        line = line.strip()
        if not line or line == "WEBVTT" or re.match(r"^\d+$", line) or "-->" in line:
            continue
        line = re.sub(r"<[^>]+>", "", line)
        result.append(line)
    return "\n".join(result)


# ===== Gemini文字起こし =====
def transcribe_with_gemini(media_path: Path, info: dict) -> str:
    suffix = media_path.suffix.lower()
    mime_type = MIME_MAP.get(suffix, "video/mp4")
    print(f"[Gemini] アップロード中: {media_path.name} ({mime_type})")
    file_size = media_path.stat().st_size

    # Step1: resumableアップロード開始
    init_res = requests.post(
        f"https://generativelanguage.googleapis.com/upload/v1beta/files?key={GEMINI_API_KEY}",
        headers={
            "X-Goog-Upload-Protocol": "resumable",
            "X-Goog-Upload-Command": "start",
            "X-Goog-Upload-Header-Content-Type": mime_type,
            "X-Goog-Upload-Header-Content-Length": str(file_size),
            "Content-Type": "application/json",
        },
        json={"file": {"displayName": media_path.name}}
    )
    upload_url = init_res.headers.get("X-Goog-Upload-URL")
    if not upload_url:
        raise Exception("アップロードURL取得失敗")

    # Step2: ファイルアップロード
    with open(media_path, "rb") as f:
        upload_res = requests.post(
            upload_url,
            headers={
                "Content-Length": str(file_size),
                "X-Goog-Upload-Offset": "0",
                "X-Goog-Upload-Command": "upload, finalize",
            },
            data=f
        )
    file_info = upload_res.json()
    file_uri = file_info.get("file", {}).get("uri")
    if not file_uri:
        raise Exception(f"ファイルURI取得失敗: {file_info}")

    # Step3: 処理待ち
    file_name = file_info["file"]["name"]
    file_state = file_info["file"].get("state", {})
    if isinstance(file_state, dict):
        file_state = file_state.get("name", "")
    print("[Gemini] 解析中...")
    poll_count = 0
    while file_state == "PROCESSING" and poll_count < 60:
        time.sleep(5)
        poll_res = requests.get(
            f"https://generativelanguage.googleapis.com/v1beta/{file_name}?key={GEMINI_API_KEY}"
        )
        poll_data = poll_res.json()
        file_state = poll_data.get("state", {})
        if isinstance(file_state, dict):
            file_state = file_state.get("name", "")
        poll_count += 1
        print(f"[Gemini] {poll_count * 5}秒経過 state={file_state}")

    # Step4: 文字起こし（ストリーミング）
    prompt = (
        f"{info['title']}\n{info['date']}・{info['time']}\n"
        f"先方：{info['other']}\n弊社：{info['our']}\n\n"
        "音声ファイルの文字起こしをお願いします。\n"
        "フォーマット: [MM:SS] 話者: 発言内容\n"
        "話者が変わるたびに明確に区別してください。"
    )
    gen_res = requests.post(
        f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:streamGenerateContent?alt=sse&key={GEMINI_API_KEY}",
        json={
            "contents": [{"role": "user", "parts": [
                {"fileData": {"mimeType": mime_type, "fileUri": file_uri}},
                {"text": prompt}
            ]}],
            "generationConfig": {"maxOutputTokens": 65536}
        },
        stream=True,
        timeout=600,
    )
    if not gen_res.ok:
        raise Exception(f"Gemini APIエラー {gen_res.status_code}: {gen_res.text[:500]}")
    import json as _json
    chunks = []
    for line in gen_res.iter_lines():
        if not line:
            continue
        text = line.decode("utf-8") if isinstance(line, bytes) else line
        if not text.startswith("data: "):
            continue
        payload = text[6:]
        if payload.strip() == "[DONE]":
            break
        try:
            chunk_data = _json.loads(payload)
            candidates = chunk_data.get("candidates", [])
            if not candidates:
                error_info = chunk_data.get("error") or chunk_data.get("promptFeedback")
                if error_info:
                    print(f"\n[Gemini] APIエラー: {error_info}")
                continue
            candidate = candidates[0]
            finish_reason = candidate.get("finishReason", "")
            if finish_reason and finish_reason not in ("STOP", "MAX_TOKENS", ""):
                print(f"\n[Gemini] 生成停止理由: {finish_reason}")
            part_text = candidate.get("content", {}).get("parts", [{}])[0].get("text", "")
            if part_text:
                chunks.append(part_text)
                print(".", end="", flush=True)
        except Exception as e:
            print(f"\n[Gemini] chunk解析エラー: {e} / raw: {text[:200]}")
    print()
    transcript = "".join(chunks)
    if not transcript:
        raise Exception("文字起こし失敗: レスポンスが空です")
    print("[Gemini] 文字起こし完了")
    return transcript


# ===== Claude議事録生成 =====
def generate_minutes_with_claude(transcript: str, info: dict) -> str:
    title  = info["title"]
    date   = info["date"]
    time_  = info["time"]
    other  = info["other"]
    our    = info["our"]
    is_internal = "社内" in title
    next_mtg = info.get("next_mtg", {})
    has_next_mtg = next_mtg.get("datetime", "未定") != "未定"

    if is_internal:
        next_mtg_section = ""
    elif has_next_mtg:
        next_mtg_section = f"""━━━━━━━━━━━━━━━━━
📅次回MTGついて
━━━━━━━━━━━━━━━━━
■日時
{next_mtg['datetime']}

■Web会議URL
{next_mtg['zoomUrl']}
ミーティングID： {next_mtg['meetingId']}
パスコード： {next_mtg['passcode']}

━━━━━━━━━━━━━━━━━"""
    else:
        next_mtg_section = """━━━━━━━━━━━━━━━━━
📅次回MTGついて
━━━━━━━━━━━━━━━━━
下記URLよりご都合の良い日時をご指定いただけますと幸いです。
↓↓ オンラインMTG（30分）の日時調整する ↓↓

https://kyozon.eeasy.jp/Advisory_Agreement_sw

・〇月度の〇回分ご調整をお願いいたします。

・1週間を目安にご対応いただけますと幸いです。

━━━━━━━━━━━━━━━━━"""

    prompt = f"""あなたは議事録作成のエキスパートです。以下の会議情報と文字起こしをもとに、詳細な議事録を作成してください。

【会議情報】
- 会議名: {title}
- 日時: {date} {time_}
- 先方参加者: {other}
- 弊社参加者: {our}

【出力形式】
議事録：{title}

開催概要
| 項目 | 内容 |
|------|------|
| 日時 | {date} {time_} |
| 先方 | {other} |
| 弊社 | {our} |

サマリー
（主要なポイントを3〜5個の箇条書きで）

討議内容

[議題1のタイトル]
（発言・意見・背景・経緯・数値・金額など漏らさず詳細に記載）

決定事項
（決定に至った理由・背景も含めて記載）

保留・継続検討事項
（決定に至らなかった事項）

アクションアイテム
| 担当者 | タスク内容 | 期限 |
|--------|-----------|------|

次回会議
- 日時: {next_mtg.get('datetime', '未定')}
- Zoom URL: {next_mtg.get('zoomUrl', '未定')}
- ミーティングID: {next_mtg.get('meetingId', '未定')}
- パスコード: {next_mtg.get('passcode', '未定')}

顧客向け案内文（Facebook Messenger用）
@[先方参加者の代表者名] 様
Cc：代表、参加メンバー
本日もMTGのお時間をいただきありがとうございました。
━━━━━━━━━━━━━━━━━
📝議事録について
━━━━━━━━━━━━━━━━━
簡単ではございますが、議事録を作成しましたので共有いたします。
[議事録URL]
{next_mtg_section}
引き続き、どうぞ宜しくお願いいたします。

---
【作成ルール】
- 文字起こしの内容を最初から最後まで省略せずすべて反映してください
- 金額・数値・日付・企業名・人名など具体的な情報は必ず正確に記載してください
- 絵文字はMessenger案内文のアイコン（📝📅）と区切り線（━）のみ使用可、その他は使用しないでください
- Markdown記号（#、**など）は使用しないでください
- 情報が不足している場合は「要確認」、未定の場合は「未定」と記載してください
- Messenger案内文の[先方参加者の代表者名]は先方参加者の最初の一人の名前を入れてください
- Messenger案内文の[議事録URL]はそのまま「[議事録URL]」と出力してください（後で手動で入力）
- Messenger案内文は空行を一切入れないでください。各行の間は改行1つのみ。━と本文の間も改行1つのみ。絶対に空行（連続する改行）を挿入しないでください
- 次回MTGのZoom情報（URL・ID・パスコード）は文字起こしから正確に抽出してください
- 「〇〇より」「〇〇氏より」「〇〇から」などの発言者帰属表現は使用しないでください。内容を主語なしで簡潔・明確に記載してください
- 発言を「」で引用する形式は使用しないでください。発言内容は地の文として簡潔にまとめてください

【文字起こし】
{transcript}"""

    print("[Claude] 議事録生成中...")
    client = anthropic.Anthropic(api_key=CLAUDE_API_KEY)
    text = ""
    with client.messages.stream(
        model="claude-sonnet-4-6",
        max_tokens=32000,
        messages=[{"role": "user", "content": prompt}],
        extra_headers={"anthropic-beta": "output-128k-2025-02-19"}
    ) as stream:
        text = stream.get_final_text()
    print("[Claude] 議事録生成完了")
    return text


# ===== Notionブロック変換 =====
def markdown_to_notion_blocks(text: str) -> list:
    blocks = []
    lines = text.split("\n")
    i = 0
    while i < len(lines):
        stripped = lines[i].strip()
        if not stripped:
            i += 1
            continue
        if stripped.startswith("### "):
            blocks.append({"object": "block", "type": "heading_3", "heading_3": {"rich_text": [{"type": "text", "text": {"content": stripped[4:]}}]}})
            i += 1; continue
        if stripped.startswith("## "):
            blocks.append({"object": "block", "type": "heading_2", "heading_2": {"rich_text": [{"type": "text", "text": {"content": stripped[3:]}}]}})
            i += 1; continue
        if stripped.startswith("# "):
            blocks.append({"object": "block", "type": "heading_1", "heading_1": {"rich_text": [{"type": "text", "text": {"content": stripped[2:]}}]}})
            i += 1; continue
        if stripped.startswith("|") and stripped.endswith("|"):
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                table_lines.append(lines[i].strip())
                i += 1
            rows = [l for l in table_lines if not re.match(r"^\|[\s\-|:]+\|$", l)]
            rows = [l[1:-1].split("|") for l in rows]
            rows = [[c.strip() for c in r] for r in rows]
            if rows:
                col_count = max(len(r) for r in rows)
                table_rows = []
                for row in rows:
                    padded = row + [""] * (col_count - len(row))
                    table_rows.append({"object": "block", "type": "table_row", "table_row": {"cells": [[{"type": "text", "text": {"content": c}}] for c in padded]}})
                blocks.append({"object": "block", "type": "table", "table": {"table_width": col_count, "has_column_header": True, "has_row_header": False, "children": table_rows}})
            continue
        if stripped.startswith("- ") or stripped.startswith("* "):
            blocks.append({"object": "block", "type": "bulleted_list_item", "bulleted_list_item": {"rich_text": [{"type": "text", "text": {"content": stripped[2:]}}]}})
            i += 1; continue
        blocks.append({"object": "block", "type": "paragraph", "paragraph": {"rich_text": [{"type": "text", "text": {"content": stripped}}]}})
        i += 1
    return blocks


# ===== Notion保存 =====
def upload_to_notion(minutes_text: str, info: dict):
    title  = info["title"]
    date   = info["date"]
    time_  = info["time"]
    page_title = f"議事録：{title} ({date} {time_})"
    all_blocks = markdown_to_notion_blocks(minutes_text)

    headers = {
        "Authorization": f"Bearer {NOTION_API_KEY}",
        "Notion-Version": "2022-06-28",
        "Content-Type": "application/json",
    }

    print("[Notion] ページ作成中...")
    create_res = requests.post(
        f"{NOTION_PROXY}/v1/pages",
        headers=headers,
        json={
            "parent": {"database_id": NOTION_DB_ID},
            "properties": {
                "ミーティング名": {"title": [{"text": {"content": page_title}}]}
            },
            "children": all_blocks[:100]
        }
    )
    create_data = create_res.json()
    if create_data.get("object") == "error":
        raise Exception(create_data.get("message"))
    page_id = create_data["id"]

    # 残りブロックを100件ずつ追加
    remaining = all_blocks[100:]
    while remaining:
        batch = remaining[:100]
        remaining = remaining[100:]
        requests.patch(
            f"{NOTION_PROXY}/v1/blocks/{page_id}/children",
            headers=headers,
            json={"children": batch}
        )

    page_url = create_data.get("url") or f"https://notion.so/{page_id.replace('-', '')}"
    print(f"[Notion] 保存完了: {page_url}")
    return page_url


# 対応ファイル拡張子
MEDIA_EXTS = {".mp4", ".m4a", ".mp3", ".wav", ".m4v"}
TEXT_EXTS  = {".txt", ".vtt"}


# ===== メイン処理 =====
def process_file_from_info(file_path: Path, info: dict):
    """ファイルパスと会議情報 dict を受け取って議事録化する（内部共通処理）。"""
    try:
        suffix = file_path.suffix.lower()
        if suffix in TEXT_EXTS:
            raw = file_path.read_text(encoding="utf-8", errors="ignore")
            transcript = parse_vtt(raw) if suffix == ".vtt" else raw
            print(f"[テキスト] 文字起こしファイルを読み込みました ({len(transcript)}文字)")
        else:
            transcript = transcribe_with_gemini(file_path, info)

        is_internal = "社内" in info["title"]
        next_mtg = {"datetime": "未定", "zoomUrl": "未定", "meetingId": "未定", "passcode": "未定"}
        if not is_internal:
            next_mtg = fetch_next_meeting(info["title"])
        info["next_mtg"] = next_mtg

        minutes = generate_minutes_with_claude(transcript, info)
        url     = upload_to_notion(minutes, info)
        print(f"完了: {url}")
    except Exception as e:
        print(f"[ERROR] {e}")


def process_file(file_path: Path):
    folder_name = file_path.parent.name
    info = parse_folder_name(folder_name)

    # フォルダ名から会議情報が取れない場合、Googleカレンダーで直前会議を検索
    if not info["date"] or info["title"] == folder_name:
        file_mtime = datetime.fromtimestamp(file_path.stat().st_mtime, tz=timezone.utc)
        cal_info = fetch_meeting_from_calendar_by_time(file_mtime)
        if cal_info["title"]:
            info.update(cal_info)

    print(f"\n{'='*50}")
    print(f"処理開始: {folder_name}")
    print(f"会議名: {info['title']} / 日付: {info['date']} / 先方: {info['other']}")
    process_file_from_info(file_path, info)


class ZoomFolderHandler(FileSystemEventHandler):
    def __init__(self):
        self.processed = load_processed()

    def _handle_file(self, path: Path):
        suffix = path.suffix.lower()
        if suffix not in {".mp4", ".m4v", ".vtt", ".txt"}:
            return
        folder_key = str(path.parent)
        if folder_key in self.processed:
            return
        # ファイルが書き込み完了するまで待つ
        time.sleep(10)
        self.processed.add(folder_key)
        save_processed(self.processed)
        if suffix in {".vtt", ".txt"}:
            print(f"[監視] テキストファイルを使用: {path.name}")
            process_file(path)
        else:
            # MP4/M4V の場合は同フォルダのVTTを優先、なければMP4を使う
            vtt_files = list(path.parent.glob("*.vtt"))
            if vtt_files:
                print(f"[監視] VTTファイルを使用: {vtt_files[0].name}")
                process_file(vtt_files[0])
            else:
                print(f"[監視] MP4ファイルを使用: {path.name}")
                process_file(path)

    def on_created(self, event):
        if not event.is_directory:
            self._handle_file(Path(event.src_path))

    def on_moved(self, event):
        if not event.is_directory:
            self._handle_file(Path(event.dest_path))


if __name__ == "__main__":
    # 環境変数チェック
    missing = [k for k, v in {
        "GEMINI_API_KEY": GEMINI_API_KEY,
        "CLAUDE_API_KEY": CLAUDE_API_KEY,
        "NOTION_API_KEY": NOTION_API_KEY,
        "NOTION_DB_ID":   NOTION_DB_ID,
    }.items() if not v]
    if missing:
        print(f"[ERROR] 環境変数が未設定: {', '.join(missing)}")
        exit(1)

    # クラウド録画ポーリングモード: python auto_minutes.py poll [--days N] [--dry-run] [--reprocess]
    if len(sys.argv) >= 2 and sys.argv[1] == "poll":
        args  = sys.argv[2:]
        days  = 1
        for i, a in enumerate(args):
            if a == "--days" and i + 1 < len(args):
                days = int(args[i + 1])
        zoom_poll(
            days=days,
            dry_run="--dry-run" in args,
            reprocess="--reprocess" in args,
        )
        exit(0)

    # テストモード: python auto_minutes.py test "フォルダ名の一部" [--force]
    if len(sys.argv) >= 2 and sys.argv[1] == "test":
        args = sys.argv[2:]
        force = "--force" in args
        keyword = next((a for a in args if not a.startswith("--")), "")
        # ZoomフォルダからVTT/MP4を検索
        candidates = []
        for folder in sorted(ZOOM_FOLDER.iterdir()):
            if not folder.is_dir():
                continue
            if keyword and keyword not in folder.name:
                continue
            for ext in [".vtt", ".m4a", ".mp4"]:
                files = list(folder.glob(f"*{ext}"))
                if files:
                    candidates.append(files[0])
                    break
        if not candidates:
            print(f"[ERROR] 対象ファイルが見つかりません: {ZOOM_FOLDER}")
            exit(1)
        # 最新のものを使う
        target = candidates[-1]
        folder_key = str(target.parent)
        processed = load_processed()
        if folder_key in processed and not force:
            print(f"[SKIP] 処理済みです: {target.parent.name}")
            print(f"  再実行するには --force オプションをつけてください")
            print(f"  例: python auto_minutes.py test {keyword} --force")
            exit(0)
        print(f"テストモード: {target}")
        process_file(target)
        processed.add(folder_key)
        save_processed(processed)
        exit(0)

    print(f"監視開始: {ZOOM_FOLDER}")
    print("Ctrl+C で停止")

    handler = ZoomFolderHandler()

    # 起動時スキャン：既存の未処理ファイルを処理する
    print("[起動スキャン] 未処理ファイルを確認中...")
    scan_count = 0
    if ZOOM_FOLDER.exists():
        for folder in sorted(ZOOM_FOLDER.iterdir()):
            if not folder.is_dir():
                continue
            folder_key = str(folder)
            if folder_key in handler.processed:
                continue
            target = None
            for ext in [".vtt", ".txt", ".mp4", ".m4a", ".m4v"]:
                files = list(folder.glob(f"*{ext}"))
                if files:
                    target = files[0]
                    break
            if target:
                print(f"[起動スキャン] 未処理: {folder.name}")
                handler.processed.add(folder_key)
                save_processed(handler.processed)
                process_file(target)
                scan_count += 1
    if scan_count == 0:
        print("[起動スキャン] 未処理ファイルなし")
    else:
        print(f"[起動スキャン] {scan_count}件処理しました")

    observer = Observer()
    observer.schedule(handler, str(ZOOM_FOLDER), recursive=True)
    observer.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
