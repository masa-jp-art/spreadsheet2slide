import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from google.colab import auth
from google.auth import default
import time

def create_slide_with_texts(service, row_data):
    """新しいスライドを作成し、カスタマイズされたテキストボックスを配置する"""

    # 各テキストボックスの詳細設定
    text_box_settings = [
        {
            # 1つ目のテキストボックス（タイトル用）
            'width': 600,
            'height': 50,
            'x': 30,
            'y': 30,
            'font_family': 'BIZ UDPMincho',
            'font_size': 30,
            'bold': True,
            'alignment': 'START'  # 左揃え
        },
        {
            # 2つ目のテキストボックス（サブタイトル用）
            'width': 400,
            'height': 240,
            'x': 30,
            'y': 80,
            'font_family': 'BIZ UDPGothic',
            'font_size': 12,
            'bold': False,
            'alignment': 'START',  # 左揃え
            #'bullet': True  # 箇条書きを有効化
        },
        {
            # 3つ目のテキストボックス（本文用）
            'width': 400,
            'height': 80,
            'x': 30,
            'y': 320,
            'font_family': 'BIZ UDPMincho',
            'font_size': 14,
            'bold': True,
            'alignment': 'START'  # 左揃え
        }
    ]

    # スライドの作成
    create_slide_request = {
        'requests': [{
            'createSlide': {
                'slideLayoutReference': {
                    'predefinedLayout': 'BLANK'
                }
            }
        }]
    }

    slide_response = service.presentations().batchUpdate(
        presentationId=PRESENTATION_ID,
        body=create_slide_request
    ).execute()

    slide_id = slide_response.get('replies')[0].get('createSlide').get('objectId')
    requests = []

    # 各テキストボックスの作成
    for i, (text, settings) in enumerate(zip(row_data[:3], text_box_settings)):
        box_id = f'box_{slide_id}_{i}'

        # テキストボックスの作成
        requests.append({
            'createShape': {
                'objectId': box_id,
                'shapeType': 'TEXT_BOX',
                'elementProperties': {
                    'pageObjectId': slide_id,
                    'size': {
                        'width': {'magnitude': settings['width'], 'unit': 'PT'},
                        'height': {'magnitude': settings['height'], 'unit': 'PT'}
                    },
                    'transform': {
                        'scaleX': 1,
                        'scaleY': 1,
                        'translateX': settings['x'],
                        'translateY': settings['y'],
                        'unit': 'PT'
                    }
                }
            }
        })

        # テキストの挿入
        requests.append({
            'insertText': {
                'objectId': box_id,
                'text': str(text)
            }
        })

        # テキストのスタイル設定
        requests.append({
            'updateTextStyle': {
                'objectId': box_id,
                'style': {
                    'fontSize': {'magnitude': settings['font_size'], 'unit': 'PT'},
                    'fontFamily': settings['font_family'],
                    'bold': settings['bold']
                },
                'textRange': {
                    'type': 'ALL'
                },
                'fields': 'fontSize,fontFamily,bold'
            }
        })

        # テキストの配置設定
        requests.append({
            'updateParagraphStyle': {
                'objectId': box_id,
                'style': {
                    'alignment': settings['alignment']
                },
                'textRange': {
                    'type': 'ALL'
                },
                'fields': 'alignment'
            }
        })

    # すべてのリクエストを一括で実行
    if requests:
        body = {'requests': requests}
        service.presentations().batchUpdate(
            presentationId=PRESENTATION_ID,
            body=body
        ).execute()

def process_spreadsheet():
    """スプレッドシートの処理を実行"""
    try:
        # GoogleスプレッドシートのIDとシート名
        SPREADSHEET_ID = '<任意のGoogleスプレッドシートのID>'
        SHEET_NAME = '<任意のスプレッドシートのシート名>'

        # GoogleスライドのプレゼンテーションID
        global PRESENTATION_ID
        PRESENTATION_ID = '<任意のGoogleスライドのID>'

        # Google認証
        auth.authenticate_user()
        creds, _ = default()

        # Googleスプレッドシートに接続
        gc = gspread.authorize(creds)
        worksheet = gc.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)

        # Googleスライドに接続
        slides_service = build('slides', 'v1', credentials=creds)

        # すべての行を取得
        all_rows = worksheet.get_all_values()
        data_rows = all_rows[1:]  # ヘッダー行をスキップ

        print(f"処理を開始します。合計 {len(data_rows)} 行のデータを処理します。")

        for i, row in enumerate(data_rows, 2):
            if not any(row[:3]):
                print(f"行 {i} は空のためスキップします。")
                continue

            try:
                create_slide_with_texts(slides_service, row)
                print(f"行 {i} の処理が完了しました。")
                time.sleep(1)

            except Exception as e:
                print(f"行 {i} の処理中にエラーが発生しました: {str(e)}")
                time.sleep(2)
                continue

        print("すべての処理が完了しました。")

    except Exception as e:
        print(f"処理中にエラーが発生しました: {str(e)}")

if __name__ == "__main__":
    process_spreadsheet()
