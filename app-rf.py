# 必要なライブラリのインポート
from typing import List, Dict, Any
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from google.colab import auth
from google.auth import default
import time
import logging

# ロギングの設定
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

class SlideCreator:
    """Google Slidesでスライドを作成するためのクラス"""

    def __init__(self, presentation_id: str):
        """
        Parameters:
            presentation_id (str): Google SlidesのプレゼンテーションID
        """
        self.presentation_id = presentation_id
        self.text_box_settings = self._initialize_text_box_settings()

    def _initialize_text_box_settings(self) -> List[Dict[str, Any]]:
        """テキストボックスの設定を初期化する"""
        return [
            {
                # タイトル用テキストボックス
                'width': 600, 'height': 50,
                'x': 30, 'y': 30,
                'font_family': 'BIZ UDPMincho',
                'font_size': 30,
                'bold': True,
                'alignment': 'START'
            },
            {
                # サブタイトル用テキストボックス
                'width': 400, 'height': 240,
                'x': 30, 'y': 80,
                'font_family': 'BIZ UDPGothic',
                'font_size': 12,
                'bold': False,
                'alignment': 'START'
            },
            {
                # 本文用テキストボックス
                'width': 400, 'height': 80,
                'x': 30, 'y': 320,
                'font_family': 'BIZ UDPMincho',
                'font_size': 14,
                'bold': True,
                'alignment': 'START'
            }
        ]

    def create_slide(self, service: Any, texts: List[str]) -> None:
        """
        新しいスライドを作成し、テキストボックスを配置する

        Parameters:
            service: Google Slides API service
            texts: 配置するテキストのリスト
        """
        # 空のスライドを作成
        slide_id = self._create_empty_slide(service)

        # テキストボックスを作成して配置
        requests = self._generate_text_box_requests(slide_id, texts)

        # バッチ更新を実行
        if requests:
            self._execute_batch_update(service, requests)

    def _create_empty_slide(self, service: Any) -> str:
        """空のスライドを作成し、スライドIDを返す"""
        create_slide_request = {
            'requests': [{
                'createSlide': {
                    'slideLayoutReference': {
                        'predefinedLayout': 'BLANK'
                    }
                }
            }]
        }

        response = service.presentations().batchUpdate(
            presentationId=self.presentation_id,
            body=create_slide_request
        ).execute()

        return response.get('replies')[0].get('createSlide').get('objectId')

    def _generate_text_box_requests(self, slide_id: str, texts: List[str]) -> List[Dict]:
        """テキストボックス作成のためのリクエストを生成"""
        requests = []

        for i, (text, settings) in enumerate(zip(texts[:3], self.text_box_settings)):
            box_id = f'box_{slide_id}_{i}'

            # シェイプ作成リクエスト
            requests.append(self._create_shape_request(slide_id, box_id, settings))
            # テキスト挿入リクエスト
            requests.append(self._create_text_request(box_id, text))
            # スタイル更新リクエスト
            requests.append(self._create_style_request(box_id, settings))
            # 配置更新リクエスト
            requests.append(self._create_alignment_request(box_id, settings))

        return requests

    def _execute_batch_update(self, service: Any, requests: List[Dict]) -> None:
        """バッチ更新を実行"""
        body = {'requests': requests}
        service.presentations().batchUpdate(
            presentationId=self.presentation_id,
            body=body
        ).execute()

    @staticmethod
    def _create_shape_request(slide_id: str, box_id: str, settings: Dict) -> Dict:
        """シェイプ作成リクエストを生成"""
        return {
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
                        'scaleX': 1, 'scaleY': 1,
                        'translateX': settings['x'],
                        'translateY': settings['y'],
                        'unit': 'PT'
                    }
                }
            }
        }

    @staticmethod
    def _create_text_request(box_id: str, text: str) -> Dict:
        """テキスト挿入リクエストを生成"""
        return {
            'insertText': {
                'objectId': box_id,
                'text': str(text)
            }
        }

    @staticmethod
    def _create_style_request(box_id: str, settings: Dict) -> Dict:
        """スタイル更新リクエストを生成"""
        return {
            'updateTextStyle': {
                'objectId': box_id,
                'style': {
                    'fontSize': {'magnitude': settings['font_size'], 'unit': 'PT'},
                    'fontFamily': settings['font_family'],
                    'bold': settings['bold']
                },
                'textRange': {'type': 'ALL'},
                'fields': 'fontSize,fontFamily,bold'
            }
        }

    @staticmethod
    def _create_alignment_request(box_id: str, settings: Dict) -> Dict:
        """配置更新リクエストを生成"""
        return {
            'updateParagraphStyle': {
                'objectId': box_id,
                'style': {'alignment': settings['alignment']},
                'textRange': {'type': 'ALL'},
                'fields': 'alignment'
            }
        }

class SpreadsheetProcessor:
    """スプレッドシートを処理するクラス"""

    def __init__(self, spreadsheet_id: str, sheet_name: str, presentation_id: str):
        """
        Parameters:
            spreadsheet_id (str): スプレッドシートID
            sheet_name (str): シート名
            presentation_id (str): プレゼンテーションID
        """
        self.spreadsheet_id = spreadsheet_id
        self.sheet_name = sheet_name
        self.slide_creator = SlideCreator(presentation_id)

    def process(self) -> None:
        """スプレッドシートの処理を実行"""
        try:
            # Google認証を実行
            auth.authenticate_user()
            creds, _ = default()

            # APIサービスを初期化
            gc = gspread.authorize(creds)
            slides_service = build('slides', 'v1', credentials=creds)

            # スプレッドシートからデータを取得
            worksheet = gc.open_by_key(self.spreadsheet_id).worksheet(self.sheet_name)
            data_rows = worksheet.get_all_values()[1:]  # ヘッダーをスキップ

            logger.info(f"処理を開始します。合計 {len(data_rows)} 行のデータを処理します。")

            # 各行を処理
            self._process_rows(slides_service, data_rows)

            logger.info("すべての処理が完了しました。")

        except Exception as e:
            logger.error(f"処理中にエラーが発生しました: {str(e)}")

    def _process_rows(self, slides_service: Any, data_rows: List[List[str]]) -> None:
        """データ行を処理"""
        for i, row in enumerate(data_rows, 2):
            if not any(row[:3]):
                logger.info(f"行 {i} は空のためスキップします。")
                continue

            try:
                self.slide_creator.create_slide(slides_service, row)
                logger.info(f"行 {i} の処理が完了しました。")
                time.sleep(1)

            except Exception as e:
                logger.error(f"行 {i} の処理中にエラーが発生しました: {str(e)}")
                time.sleep(2)

def main():
    """メイン処理"""
    # 設定値
    SPREADSHEET_ID = '<GoogleスプレッドシートのID>'
    SHEET_NAME = '<シート名>'
    PRESENTATION_ID = '<GoogleスライドのID>'

    # 処理の実行
    processor = SpreadsheetProcessor(SPREADSHEET_ID, SHEET_NAME, PRESENTATION_ID)
    processor.process()

if __name__ == "__main__":
    main()

