"""
Zoomフォルダ監視 → 自動議事録生成 → Notion保存
"""

import os
import re
import sys
import time
import json
import requests
import anthropic
sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")
from datetime import datetime, timezone
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


# ===== フォルダ名から会議情報を解析 =====
def parse_folder_name(folder_name: str) -> dict:
    """
    例: "2026-04-09 13.59.53【定例】顧問契約／株式会社ウェルフォート"
    """
    info = {"title": folder_name, "date": "", "time": "", "other": "", "our": "弊社"}

    # 日時
    m = re.match(r"(\d{4}-\d{2}-\d{2})\s+(\d{2})\.(\d{2})\.\d{2}", folder_name)
    if m:
        info["date"] = m.group(1)
        info["time"] = f"{m.group(2)}:{m.group(3)}"

    # 会議名（【】以降）
    title_match = re.search(r"(【.+】.+)", folder_name)
    if title_match:
        info["title"] = title_match.group(1)

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
        today = datetime.now().date()
        tomorrow = datetime(today.year, today.month, today.day + 1, tzinfo=timezone.utc)
        time_min = tomorrow.isoformat()

        # 顧客名で検索して直近1件だけ取得
        result = service.events().list(
            calendarId="primary",
            q=search_query,
            timeMin=time_min,
            singleEvents=True,
            orderBy="startTime",
            maxResults=1
        ).execute()
        items = result.get("items", [])
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
        return {"datetime": dt_str, **zoom}
    except Exception as ex:
        print(f"[Calendar] エラー: {ex}")
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

    # Step4: 文字起こし
    prompt = (
        f"{info['title']}\n{info['date']}・{info['time']}\n"
        f"先方：{info['other']}\n弊社：{info['our']}\n\n"
        "音声ファイルの文字起こしをお願いします。\n"
        "フォーマット: [MM:SS] 話者: 発言内容\n"
        "話者が変わるたびに明確に区別してください。"
    )
    gen_res = requests.post(
        f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={GEMINI_API_KEY}",
        json={
            "contents": [{"role": "user", "parts": [
                {"fileData": {"mimeType": mime_type, "fileUri": file_uri}},
                {"text": prompt}
            ]}],
            "generationConfig": {"maxOutputTokens": 65536}
        }
    )
    gen_data = gen_res.json()
    transcript = gen_data.get("candidates", [{}])[0].get("content", {}).get("parts", [{}])[0].get("text", "")
    if not transcript:
        raise Exception(f"文字起こし失敗: {gen_data}")
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
📅次回MTGについて
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
- 次回MTGのZoom情報（URL・ID・パスコード）は文字起こしから正確に抽出してください

【文字起こし】
{transcript}"""

    print("[Claude] 議事録生成中...")
    client = anthropic.Anthropic(api_key=CLAUDE_API_KEY)
    text = ""
    with client.messages.stream(
        model="claude-sonnet-4-5",
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
def process_file(file_path: Path):
    folder_name = file_path.parent.name
    info = parse_folder_name(folder_name)
    print(f"\n{'='*50}")
    print(f"処理開始: {folder_name}")
    print(f"会議名: {info['title']} / 日付: {info['date']} / 先方: {info['other']}")

    try:
        suffix = file_path.suffix.lower()
        if suffix in TEXT_EXTS:
            # テキスト/VTTはそのまま読み込む
            raw = file_path.read_text(encoding="utf-8", errors="ignore")
            transcript = parse_vtt(raw) if suffix == ".vtt" else raw
            print(f"[テキスト] 文字起こしファイルを読み込みました ({len(transcript)}文字)")
        else:
            transcript = transcribe_with_gemini(file_path, info)

        # 次回MTG取得（社内はスキップ）
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


class ZoomFolderHandler(FileSystemEventHandler):
    def __init__(self):
        self.processed = load_processed()

    def _handle_mp4(self, path: Path):
        if path.suffix.lower() not in {".mp4", ".m4v"}:
            return
        folder_key = str(path.parent)
        if folder_key in self.processed:
            return
        # ファイルが書き込み完了するまで待つ
        time.sleep(10)
        self.processed.add(folder_key)
        save_processed(self.processed)
        # 同じフォルダのVTTを優先、なければMP4を使う
        vtt_files = list(path.parent.glob("*.vtt"))
        if vtt_files:
            print(f"[監視] VTTファイルを使用: {vtt_files[0].name}")
            process_file(vtt_files[0])
        else:
            print(f"[監視] MP4ファイルを使用: {path.name}")
            process_file(path)

    def on_created(self, event):
        if not event.is_directory:
            self._handle_mp4(Path(event.src_path))

    def on_moved(self, event):
        if not event.is_directory:
            self._handle_mp4(Path(event.dest_path))


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

    # テストモード: python auto_minutes.py test "フォルダ名の一部"
    if len(sys.argv) >= 2 and sys.argv[1] == "test":
        keyword = sys.argv[2] if len(sys.argv) >= 3 else ""
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
        print(f"テストモード: {target}")
        process_file(target)
        exit(0)

    print(f"監視開始: {ZOOM_FOLDER}")
    print("Ctrl+C で停止")

    handler  = ZoomFolderHandler()
    observer = Observer()
    observer.schedule(handler, str(ZOOM_FOLDER), recursive=True)
    observer.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
