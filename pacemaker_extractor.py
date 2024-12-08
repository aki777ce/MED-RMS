
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import PyPDF2
import re
import csv
import os
from typing import Optional

class DeviceData:
    def __init__(self, 患者ID, 送信日時, **kwargs):
        self.患者ID = 患者ID
        self.送信日時 = 送信日時
        self.デバイスタイプ = kwargs.get('デバイスタイプ', '')
        self.心房リードインピーダンス = kwargs.get('心房リードインピーダンス', None)
        self.心室リードインピーダンス = kwargs.get('心室リードインピーダンス', None)
        self.心房ペーシング閾値 = kwargs.get('心房ペーシング閾値', None)
        self.心房パルス幅 = kwargs.get('心房パルス幅', None)
        self.心室ペーシング閾値 = kwargs.get('心室ペーシング閾値', None)
        self.心室パルス幅 = kwargs.get('心室パルス幅', None)
        self.P波高値 = kwargs.get('P波高値', None)
        self.R波高値 = kwargs.get('R波高値', None)
        self.予測寿命_最小 = kwargs.get('予測寿命_最小', None)
        self.予測寿命_最大 = kwargs.get('予測寿命_最大', None)
        self.ATAF時間パーセント = kwargs.get('ATAF時間パーセント', None)
        self.VT回数 = kwargs.get('VT回数', None)
        self.ASVS = kwargs.get('ASVS', None)
        self.ASVP = kwargs.get('ASVP', None)
        self.APVS = kwargs.get('APVS', None)
        self.APVP = kwargs.get('APVP', None)
        # Micra AV2用パラメータ
        self.キャプチャ閾値 = kwargs.get('心室閾値', None)
        self.AMVS = kwargs.get('AMVS', None)
        self.VSオンリー = kwargs.get('VSオンリー', None)
        self.AMVP = kwargs.get('AMVP', None)
        self.VPオンリー = kwargs.get('VPオンリー', None)
        # Cobalt/Evera用パラメータ
        self.RVコイル = kwargs.get('RVコイル', None)
        self.SVCコイル = kwargs.get('SVCコイル', None)
        self.VF治療回数 = kwargs.get('VF治療回数', None)
        self.VT治療回数 = kwargs.get('VT治療回数', None)
        # LV関連パラメータ
        self.LVインピーダンス = kwargs.get('LVインピーダンス', None)
        self.LVペーシング閾値 = kwargs.get('LVペーシング閾値', None)
        self.LVパルス幅 = kwargs.get('LVパルス幅', None)

    def get_value(self, attr_name: str) -> str:
        """属性値を取得し、Noneまたは"None"の場合は空文字列を返す"""
        value = getattr(self, attr_name)
        if value is None or str(value) == "None":
          return ""
        return str(value)

    def get_formatted_date(self) -> str:
        """送信日時を整形された形式で返す"""
        from datetime import datetime
        date_formats = ['%d-%b-%Y', '%Y.%m.%d']
        for date_format in date_formats:
            try:
                date_obj = datetime.strptime(self.送信日時.split()[0], date_format)
                return date_obj.strftime('%Y年%m月%d日')
            except ValueError:
                continue
        return self.送信日時

    def get_formatted_lifetime(self) -> str:
        """予測寿命を「最小-最大y」の形式で返す"""
        min_life = self.get_value('予測寿命_最小')
        max_life = self.get_value('予測寿命_最大')
        if min_life and max_life:
            return f"{min_life}-{max_life}y"
        return ""

class PacemakerExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ペースメーカーデータ抽出")
        self.root.geometry("1600x900")
        
        # データ保持用変数
        self.pdf_text = ""
        self.extracted_data = []
        
        self.create_widgets()
        
    def create_modern_button(self, parent, text, command):
        """モダンなボタンを作成するヘルパーメソッド"""
        button = tk.Button(parent,
            text=text,
            command=command,
            bg='#1a73e8',  # 背景色
            fg='white',    # 文字色
            font=('Helvetica', 10),
            relief='flat',
            padx=20,
            pady=10,
            cursor='hand2'  # ホバー時のカーソル
        )
        # ホバーエフェクト
        button.bind('<Enter>', lambda e: button.configure(bg='#1557b0'))
        button.bind('<Leave>', lambda e: button.configure(bg='#1a73e8'))
        return button

    def create_widgets(self):
        # メインフレームのスタイル設定
        style = ttk.Style()
        style.configure('Main.TFrame', background='#f0f2f5')
        style.configure('Card.TLabelframe', background='white', relief='flat', borderwidth=0)
        style.configure('Card.TLabelframe.Label', background='white', foreground='#1a73e8', font=('Helvetica', 12, 'bold'))
        
        # Treeviewスタイル
        style.configure('Modern.Treeview',
            background='white',
            fieldbackground='white',
            foreground='#333333',
            rowheight=30,
            font=('Helvetica', 10)
        )
        style.configure('Modern.Treeview.Heading',
            background='#f8f9fa',
            foreground='#1a73e8',
            font=('Helvetica', 10, 'bold')
        )
        
        # メインフレーム
        main_frame = ttk.Frame(self.root, padding="20", style='Main.TFrame')
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # アクションカード（上部ボタン群）
        action_card = ttk.LabelFrame(main_frame, padding="15", text="アクション", style='Card.TLabelframe')
        action_card.grid(row=0, column=0, columnspan=2, pady=(0, 20), sticky=(tk.W, tk.E))
        
        # モダンなボタンの配置
        load_btn = self.create_modern_button(action_card, "PDFファイル選択", self.load_pdf)
        extract_btn = self.create_modern_button(action_card, "データ抽出", self.extract_data)
        export_btn = self.create_modern_button(action_card, "CSV出力", self.export_to_csv)
        
        load_btn.pack(side=tk.LEFT, padx=10)
        extract_btn.pack(side=tk.LEFT, padx=10)
        export_btn.pack(side=tk.LEFT, padx=10)

        # テキストプレビューカード
        preview_card = ttk.LabelFrame(main_frame, text="PDFテキスト", padding="15", style='Card.TLabelframe')
        preview_card.grid(row=1, column=0, padx=(0, 10), sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # テキストエリアのコンテナフレーム
        text_container = ttk.Frame(preview_card)
        text_container.pack(fill=tk.BOTH, expand=True)
        
        # テキストエリアとスクロールバー
        self.text_area = tk.Text(text_container, wrap=tk.WORD, 
                               font=('Helvetica', 10),
                               bg='white', fg='#333333')
        text_scroll = ttk.Scrollbar(text_container, orient=tk.VERTICAL,
                                  command=self.text_area.yview)
        self.text_area.configure(yscrollcommand=text_scroll.set)
        
        # テキストエリアの配置
        self.text_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        text_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # 結果表示カード
        result_card = ttk.LabelFrame(main_frame, text="抽出結果", padding="15", style='Card.TLabelframe')
        result_card.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 結果表示エリアのコンテナフレーム
        tree_container = ttk.Frame(result_card)
        tree_container.pack(fill=tk.BOTH, expand=True)
        
        # 結果表示エリアとスクロールバー
        self.result_tree = ttk.Treeview(tree_container, columns=[
            '患者ID', '送信日時', '心房リードインピーダンス', '心室リードインピーダンス',
            '心房ペーシング閾値', '心房パルス幅', '心室ペーシング閾値', '心室パルス幅',
            'P波高値', 'R波高値', '予測寿命_最小', '予測寿命_最大', 'ATAF時間パーセント', 'VT回数',  # ここを変更
            'AS-VS%', 'AS-VP%', 'AP-VS%', 'AP-VP%', 'AMVS', 'VSオンリー',
            'AMVP', 'VPオンリー', 'RVコイル', 'SVCコイル',
            'VF治療回数', 'VT治療回数', 'LVインピーダンス', 'LVペーシング閾値', 'LVパルス幅'
        ], show='headings', style='Modern.Treeview')
        
        vsb = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=self.result_tree.yview)
        hsb = ttk.Scrollbar(tree_container, orient=tk.HORIZONTAL, command=self.result_tree.xview)
        self.result_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # カラム設定
        column_names = {
            '患者ID': '患者ID',
            '送信日時': '送信日時',
            '心房リードインピーダンス': '心房インピーダンス',
            '心室リードインピーダンス': '心室インピーダンス',
            '心房ペーシング閾値': '心房閾値',
            '心房パルス幅': '心房PW',
            '心室ペーシング閾値': '心室閾値',
            '心室パルス幅': '心室PW',
            'P波高値': 'P波高',
            'R波高値': 'R波高',
            '予測寿命_最小': '予測寿命_最小',
            '予測寿命_最大': '予測寿命_最大',
            'ATAF時間パーセント': 'AT/AF%',  # キーを一致させる
            'VT回数': 'VT回数',
            'AS-VS%': 'AS-VS%',
            'AS-VP%': 'AS-VP%',
            'AP-VS%': 'AP-VS%',
            'AP-VP%': 'AP-VP%',
            'AMVS': 'AMVS',
            'VSオンリー': 'VSオンリー',
            'AMVP': 'AMVP',
            'VPオンリー': 'VPオンリー',
            'RVコイル': 'RVコイル',
            'SVCコイル': 'SVCコイル',
            'VF治療回数': 'VF治療回数',
            'VT治療回数': 'VT治療回数',
            'LVインピーダンス': 'LVインピーダンス',
            'LVペーシング閾値': 'LV閾値',
            'LVパルス幅': 'LV PW'
        }

        self.result_tree.column('#0', width=0, stretch=tk.NO)
        for col in self.result_tree['columns']:
            self.result_tree.column(col, width=150, stretch=True)
            self.result_tree.heading(col, text=column_names[col], anchor=tk.CENTER)

        # コンポーネントの配置
        self.result_tree.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        vsb.grid(row=0, column=1, sticky=(tk.N, tk.S))
        hsb.grid(row=1, column=0, sticky=(tk.E, tk.W))

        # グリッドの重み付け設定
        tree_container.grid_columnconfigure(0, weight=1)
        tree_container.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(1, weight=3)
        main_frame.grid_rowconfigure(1, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(0, weight=1)

        # 初期ウィンドウサイズの設定
        self.root.geometry("1600x900")

    def extract_data(self, text):
        """テキストからデータを抽出するメソッド"""
        lines = text.split('\n')
        for line in lines:
            # 心房データを無視する
            if any(keyword in line for keyword in ["心房", "P波", "RA"]):
                continue
            # 心室データを抽出する
            if any(keyword in line for keyword in ["RV", "R波", "電極インピーダンス", "心室閾値", "電圧/パルス幅設定値"]):
                self.extracted_data.append(line)
        return self.extracted_data

    def load_pdf(self):
        file_paths = filedialog.askopenfilenames(
            filetypes=[("PDFファイル", "*.pdf")],
            title="PDFファイルを選択（複数選択可）"
        )
        
        if not file_paths:
            return
            
        # 既存のテキストをクリアするか確認
        if self.pdf_text:
            if not messagebox.askyesno("確認", 
                "既存のテキストデータをクリアして新しいファイルを読み込みますか？\n" +
                "「いいえ」を選択すると、既存のデータに追加されます。"):
                # 既存のテキストを保持
                current_text = self.pdf_text
            else:
                # テキストをクリア
                current_text = ""
                self.text_area.delete('1.0', tk.END)
        else:
            current_text = ""

        total_files = len(file_paths)
        processed_files = 0
        failed_files = []

        try:
            for file_path in file_paths:
                try:
                    with open(file_path, 'rb') as file:
                        pdf_reader = PyPDF2.PdfReader(file)
                        text = ""
                        for page in pdf_reader.pages:
                            text += page.extract_text()
                        
                        if current_text:
                            current_text += "\n\n" + text
                        else:
                            current_text = text
                        
                        processed_files += 1
                        
                except Exception as e:
                    failed_files.append((file_path, str(e)))
                    continue

            self.pdf_text = current_text
            self.text_area.delete('1.0', tk.END)
            self.text_area.insert('1.0', current_text[:2000] + "\n...(省略)...")

            # 結果メッセージの作成
            if failed_files:
                error_msg = "\n\n読み込みに失敗したファイル:\n" + \
                           "\n".join([f"{os.path.basename(f[0])}: {f[1]}" for f in failed_files])
                messagebox.showwarning("警告", 
                    f"{total_files}個中{processed_files}個のファイルを読み込みました。{error_msg}")
            else:
                messagebox.showinfo("成功", 
                    f"{total_files}個のPDFファイルの読み込みが完了しました")

        except Exception as e:
            messagebox.showerror("エラー", f"ファイルの読み込み中にエラーが発生しました: {str(e)}")

    def extract_data_with_retry(self, section: str, patterns: list, data_type: str) -> tuple:
        """
        複数のパターンを試行してデータを抽出する汎用関数
        Args:
            section: テキストセクション
            patterns: 正規表現パターンのリスト
            data_type: データの種類（デバッグ用）
        Returns:
            抽出されたデータのタプル。マッチしない場合はNoneのタプル
        """
        for pattern in patterns:
            match = re.search(pattern, section, re.DOTALL)
            if match:
                return match.groups()
            
        # どのパターンもマッチしない場合、より緩やかな数値検索を試みる
        if '心房' in data_type or '心室' in data_type:
            # 数値のみを検索（単位や記号を考慮しない）
            simple_numbers = re.findall(r'(\d+\.?\d*)', section)
            if len(simple_numbers) >= 2:
                print(f"警告: {data_type}の値を緩やかな検索で発見: {simple_numbers[:2]}")
                return tuple(simple_numbers[:2])
        
        print(f"警告: {data_type}の値が見つかりませんでした")
        return tuple([None] * len(re.findall(r'\(', patterns[0])))

    def is_valid_data(self, data: DeviceData) -> bool:
        """データが有効（少なくとも1つ以上のパラメータが存在）かをチェック"""
        # 患者IDと送信日時以外のフィールドをチェック
        fields_to_check = [
            data.心房リードインピーダンス, data.心室リードインピーダンス,
            data.心房ペーシング閾値, data.心室ペーシング閾値,
            data.P波高値, data.R波高値,
            data.予測寿命_最小, data.予測寿命_最大,
            data.キャプチャ閾値, data.AMVS, data.VSオンリー,
            data.RVコイル, data.SVCコイル,
            data.LVインピーダンス, data.LVペーシング閾値
        ]
        # 1つでもNoneでない値があればTrue
        return any(field is not None for field in fields_to_check)

    def extract_single_data_set(self, section: str) -> Optional[DeviceData]:
        try:
            # 患者IDと送信日時の抽出（これらは必須）
            id_match = re.search(r'ID：\s*(\d+)', section)
            datetime_match = re.search(r'送信日時：\s*([^\n]+)', section)
            
            if not id_match or not datetime_match:
                return None

            # デバイスタイプの判定を先に行う
            device_type = ""
            if re.search(r'Micra\s*AV', section):
                device_type = "Micra AV2"
            elif re.search(r'Cobalt', section):
                device_type = "Cobalt"
            elif re.search(r'Evera', section):
                device_type = "Evera"

            # データオブジェクトの作成
            data = DeviceData(
                患者ID=id_match.group(1),
                送信日時=datetime_match.group(1),
                デバイスタイプ=device_type
            )
            
            # リードインピーダンスの抽出
            imp_patterns = [
                # 既存のパターン
                r'リードインピーダンス\s+(\d+)\s*Ω\s+(\d+)\s*Ω',
                r'Aペーシング\s*\([^\)]+\)\s*(\d+)\s*Ω.*\nRVペーシング\s*\([^\)]+\)\s*(\d+)\s*Ω',
                r'RA\([^\)]+\)\s+RV\([^\)]+\)\s*\nリードインピーダンス\s+(\d+)\s*Ω\s+(\d+)\s*Ω',
                r'リードインピーダンス.*?(\d+)\s*Ω.*?(\d+)\s*Ω',
                # 電極インピーダンスのパターン
                r'電極インピーダンス\s+(\d+)\s*Ω',  # 単独の電極インピーダンス
                r'電極インピーダンス.*?(\d+)\s*Ω',    # 周辺テキストがある場合の電極インピーダンス
                # ペーシングインピーダンスのパターン追加
                r'ペーシングインピーダンス\s+(\d+)\s*Ω\s+(\d+)\s*Ω\s+(\d+)\s*Ω',  # 3値パターン
                r'ペーシングインピーダンス\s+(\d+)\s*Ω\s+(\d+)\s*Ω',  # 2値パターン
                r'ペーシングインピーダンス\s+(\d+)\s*Ω'  # 1値パターン
            ]

            imp_values = self.extract_data_with_retry(section, imp_patterns, "リード/電極/ペーシングインピーダンス")
            if imp_values:
                if len(imp_values) == 3:
                    # 3つの値がある場合: 心房、心室、LV
                    data.心房リードインピーダンス = imp_values[0]
                    data.心室リードインピーダンス = imp_values[1]
                    data.LVインピーダンス = imp_values[2]
                elif len(imp_values) == 2:
                    # 2つの値がある場合: 心房、心室
                    data.心房リードインピーダンス = imp_values[0]
                    data.心室リードインピーダンス = imp_values[1]
                elif len(imp_values) == 1:
                    # 1つの値のみの場合は心室のデータとして扱う
                    data.心室リードインピーダンス = imp_values[0]

            # ペーシング閾値とパルス幅の抽出
            try:
                # 2値パターン（心房・心室）
                thresh_patterns = [
                    r'ペーシング閾値\s+(\d+\.?\d*)\s*V\s*\((\d+\.?\d*)\s*ms\)\s+(\d+\.?\d*)\s*V\s*\((\d+\.?\d*)\s*ms\)$',
                    r'キャプチャ閾値\s+(\d+\.?\d*)\s*V\s*\((\d+\.?\d*)\s*ms\)\s+(\d+\.?\d*)\s*V\s*\((\d+\.?\d*)\s*ms\)$',
                    r'閾値.*?(\d+\.?\d*)\s*V.*?(\d+\.?\d*)\s*ms.*?(\d+\.?\d*)\s*V.*?(\d+\.?\d*)\s*ms(?!\s+\d)',
                    r'キャプチャ閾値\s+(\d+\.?\d*)\s*ms\s*で\s*(\d+\.?\d*)\s*V\s+(\d+\.?\d*)\s*ms\s*で\s*(\d+\.?\d*)\s*V(?!\s+\d)',
                    r'キャプチャ閾値.*?(\d+\.?\d*)\s*ms\s*で\s*(\d+\.?\d*)\s*V.*?(\d+\.?\d*)\s*ms\s*で\s*(\d+\.?\d*)\s*V(?!\s+\d)'
                ]
                # 1値パターン（心室のみ）
                single_thresh_patterns = [
                    r'キャプチャ閾値\s+(\d+\.?\d*)\s*V\s*\((\d+\.?\d*)\s*ms\)(?!\s+\d)',
                    r'キャプチャ閾値\s+(\d+\.?\d*)\s*ms\s*で\s*(\d+\.?\d*)\s*V(?!\s+\d)',
                    r'キャプチャ閾値.*?(\d+\.?\d*)\s*ms\s*で\s*(\d+\.?\d*)\s*V(?!\s+\d)',
                    r'キャプチャ閾値\s+(\d+\.?\d*)\s*ms\s*で\s*(\d+\.?\d*)\s*V'
                ]

                # 3値パターン
                three_value_patterns = [
                    r'電圧/パルス幅設定値\s+(\d+\.?\d*)\s*V\s*/\s*(\d+\.?\d*)\s*ms\s+(\d+\.?\d*)\s*V\s*/\s*(\d+\.?\d*)\s*ms\s+(\d+\.?\d*)\s*V\s*/\s*(\d+\.?\d*)\s*ms'
                ]
                    
                # 1値パターンで試行
                single_values = self.extract_data_with_retry(section, single_thresh_patterns, "単一キャプチャ閾値")
                if single_values and len(single_values) == 2:
                        if 'msで' in section:
                            data.心室パルス幅, data.心室ペーシング閾値 = single_values
                        else:
                            data.心室ペーシング閾値, data.心室パルス幅 = single_values
                
                else:
                    # 通常のデバイスの場合の処理
                    # まず2値パターンで試行
                    thresh_values = self.extract_data_with_retry(section, thresh_patterns, "ペーシング閾値")
                    
                    if thresh_values and len(thresh_values) == 4:
                        # msでパターンの判定
                        if any('msで' in pattern for pattern in thresh_patterns):
                            data.心房パルス幅, data.心房ペーシング閾値, data.心室パルス幅, data.心室ペーシング閾値 = thresh_values
                        else:
                            data.心房ペーシング閾値, data.心房パルス幅, data.心室ペーシング閾値, data.心室パルス幅 = thresh_values
                    else:
                        # 1値パターン（心室のみ）を試行
                        single_values = self.extract_data_with_retry(section, single_thresh_patterns, "単一キャプチャ閾値")
                        if single_values and len(single_values) == 2:
                            # 心房データは明示的にNoneを設定
                            data.心房ペーシング閾値 = None
                            data.心房パルス幅 = None
                            
                            if 'msで' in section:
                                data.心室パルス幅, data.心室ペーシング閾値 = single_values
                            else:
                                data.心室ペーシング閾値, data.心室パルス幅 = single_values
                        else:
                            # 3値パターンを試行
                            three_values = self.extract_data_with_retry(section, three_value_patterns, "3値パターン")
                            if three_values and len(three_values) == 6:
                                logging.info("3値パターンを検出しました。このデータは別途処理が必要かもしれません。")

            except Exception as e:
                logging.error(f"ペーシング閾値とパルス幅の抽出でエラーが発生: {str(e)}")

            # P/R波高値の抽出
            wave_patterns = [
                r'P/R波高値\s+(\d+\.?\d*)\s*mV\s+(\d+\.?\d*)\s*mV',
                r'P波高値\s+(\d+\.?\d*)\s*mV.*\nR波高値\s+(\d+\.?\d*)\s*mV',
                r'[P波|R波].*?(\d+\.?\d*)\s*mV.*?(\d+\.?\d*)\s*mV'
            ]
            wave_values = self.extract_data_with_retry(section, wave_patterns, "P/R波高値")
            if wave_values and len(wave_values) == 2:
                data.P波高値, data.R波高値 = wave_values

            # 予測寿命の抽出
            life_patterns = [
                r'予測寿命.*?最小値：\s*([\d\.]+)\s*years.*?最大値：\s*([\d\.]+)\s*years',
                r'最小値：\s*([\d\.]+)\s*years.*?最大値：\s*([\d\.]+)\s*years',
                r'予測寿命.*?(\d+\.?\d*)\s*years.*?(\d+\.?\d*)\s*years',
            ]
            life_values = self.extract_data_with_retry(section, life_patterns, "予測寿命")
            if life_values and len(life_values) >= 2:
                data.予測寿命_最小, data.予測寿命_最大 = life_values[:2]

            # VT回数の抽出
            vt_patterns = [
                r'VT \(>150 bpm\)\s+(\d+)\s+\d+',
                r'VT モニタ\s+>150 bpm.*?VT \(>150 bpm\)\s+(\d+)',
                r'VT/VFカウンタ.*?VT \(>150 bpm\)\s+(\d+)',
                r'VT.*?>150 bpm.*?(\d+)'
            ]
            vt_values = self.extract_data_with_retry(section, vt_patterns, "VT回数")
            if vt_values and len(vt_values) >= 1:
                data.VT回数 = vt_values[0]

            # AT/AF時間の抽出
            ataf_patterns = [
                r'AT/AF時間%\s+([<\d\.]+)%',
                r'AT/AF時間.*?([<\d\.]+)%',
                r'AT/AF.*?([<\d\.]+)%',
                r'AT/AF時間\s+=\s+(\d+)\s*(?:sec|min)',
                r'AT/AF\s*(?:時間|総持続時間).*?([<\d\.]+)'
            ]
            ataf_values = self.extract_data_with_retry(section, ataf_patterns, "AT/AF時間")
            if ataf_values and len(ataf_values) >= 1:
                data.ATAF時間パーセント = ataf_values[0]

            # ペーシングモードの抽出
            rate_histogram_patterns = [
                r'レートヒストグラム[\s\S]+?時間%[\s\S]+?AT/AF時間',
                r'時間%\s+総VP[\s\S]+?AT/AF時間',
                r'時間%[\s\S]+?AP-VP'
            ]

            rate_histogram_section = None
            for pattern in rate_histogram_patterns:
                match = re.search(pattern, section)
                if match:
                    rate_histogram_section = match.group(0)
                    break

            if rate_histogram_section:
                mode_patterns = [
                    r'(?:時間%|AT/AF外).*?[\n\r].*?AS-VS\s+([<\d\.]+)%\s*AS-VP\s+([<\d\.]+)%\s*AP-VS\s+([<\d\.]+)%\s*AP-VP\s+([<\d\.]+)%',
                    r'AS-VS\s+([<\d\.]+)%[\s\S]*?AS-VP\s+([<\d\.]+)%[\s\S]*?AP-VS\s+([<\d\.]+)%[\s\S]*?AP-VP\s+([<\d\.]+)%'
                ]

                for pattern in mode_patterns:
                    mode_match = re.search(pattern, rate_histogram_section)
                    if mode_match:
                        data.ASVS = mode_match.group(1)
                        data.ASVP = mode_match.group(2)
                        data.APVS = mode_match.group(3)
                        data.APVP = mode_match.group(4)
                        break

                # バックアップ：個別の値を探す
                if not all([data.ASVS, data.ASVP, data.APVS, data.APVP]):
                    individual_patterns = {
                        'AS-VS': r'AS-VS\s+([<\d\.]+)%',
                        'AS-VP': r'AS-VP\s+([<\d\.]+)%',
                        'AP-VS': r'AP-VS\s+([<\d\.]+)%',
                        'AP-VP': r'AP-VP\s+([<\d\.]+)%'
                    }
                    
                    for mode, pattern in individual_patterns.items():
                        match = re.search(pattern, rate_histogram_section)
                        if match:
                            value = match.group(1)
                            if mode == 'AS-VS' and not data.ASVS:
                                data.ASVS = value
                            elif mode == 'AS-VP' and not data.ASVP:
                                data.ASVP = value
                            elif mode == 'AP-VS' and not data.APVS:
                                data.APVS = value
                            elif mode == 'AP-VP' and not data.APVP:
                                data.APVP = value

            # デバイス特有のパラメータの抽出
            if device_type == "Micra AV2":
                # キャプチャ閾値の抽出（RVペーシング閾値と同義）
                capture_patterns = [
                    r'キャプチャ閾値\s+(\d+\.?\d*)\s*V',
                    r'キャプチャ.*?(\d+\.?\d*)\s*V'
                ]
                capture_values = self.extract_data_with_retry(section, capture_patterns, "心室閾値")
                if capture_values and len(capture_values) >= 1:
                    data.心室閾値 = capture_values[0]
                    data.心室ペーシング閾値 = capture_values[0]

                # Micra特有のペーシングモード抽出
                micra_mode_patterns = {
                    'AMVS': r'AM-VS\s+([<\d\.]+)%',
                    'VSオンリー': r'VSオンリー\s+([<\d\.]+)%',
                    'AMVP': r'AM-VP\s+([<\d\.]+)%',
                    'VPオンリー': r'VPオンリー\s+([<\d\.]+)%'
                }
                
                for mode, pattern in micra_mode_patterns.items():
                    match = re.search(pattern, section)
                    if match:
                        setattr(data, mode, match.group(1))

                    # コイルインピーダンスの抽出
                    coil_patterns = {
                        'RVコイル': r'RV\s*=\s*(\d+)\s*Ω|除細動インピーダンス\s*RV\s*([0-9]+)\s*Ω',
                        'SVCコイル': r'SVC\s*=\s*(\d+)\s*Ω|除細動インピーダンス\s*SVC\s*(\d+)\s*Ω'
                    }
                    
                    for param, pattern in coil_patterns.items():
                        match = re.search(pattern, section)
                        if match:
                            setattr(data, param, match.group(1) or match.group(2))

                # 治療回数の抽出
                therapy_patterns = {
                    'VF治療回数': r'VF\s+(\d+)',
                    'VT治療回数': r'VT\s+(\d+)',
                }
                
                for param, pattern in therapy_patterns.items():
                    match = re.search(pattern, section)
                    if match:
                        setattr(data, param, match.group(1))

            # LVパラメータの抽出（デバイスタイプに関係なく）
            if 'LV' in section:
                lv_patterns = {
                    'LVインピーダンス': r'LV.*?インピーダンス[^\d]*(\d+)\s*Ω',
                    'LVペーシング閾値': r'LV.*?ペーシング閾値\s+([\d\.]+)\s*V',
                    'LVパルス幅': r'LV.*?パルス幅\s+([\d\.]+)\s*ms'
                }
                
                for param, pattern in lv_patterns.items():
                    match = re.search(pattern, section)
                    if match:
                        setattr(data, param, match.group(1))

            # 最後にデータの有効性をチェック
            if not self.is_valid_data(data):
                return None
            
            return data
            
        except Exception as e:
            print(f"データ抽出中にエラーが発生: {str(e)}")
            print(f"問題のあるセクション: {section[:200]}...")
            return None

    def extract_data(self):
        if not self.pdf_text:
            messagebox.showwarning("警告", "PDFファイルを先に読み込んでください")
            return

        # 結果をクリア
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        
        self.extracted_data.clear()
        
        # Quick Look IIで区切ってセクションに分割
        sections = re.split(r'Quick Look II', self.pdf_text)
        
        # 重複チェック用のセット
        seen_data = set()
        
        for section in sections:
            if not section.strip():
                continue
                
            data = self.extract_single_data_set(section)
            if data:
                # 患者IDと送信日時のタプルを作成
                data_key = (data.患者ID, data.送信日時)
                
                # 重複チェック
                if data_key in seen_data:
                    continue
                
                # データの有効性チェック
                if not self.is_valid_data(data):
                    continue
                
                seen_data.add(data_key)
                self.extracted_data.append(data)
                self.result_tree.insert('', tk.END, values=(
                    data.get_value('患者ID'),
                    data.get_formatted_date(),  # 送信日時を整形された形式で取得
                    data.get_value('心房リードインピーダンス'),
                    data.get_value('心室リードインピーダンス'),
                    data.get_value('心房ペーシング閾値'),
                    data.get_value('心房パルス幅'),
                    data.get_value('心室ペーシング閾値'),
                    data.get_value('心室パルス幅'),
                    data.get_value('P波高値'),
                    data.get_value('R波高値'),
                    data.get_value('予測寿命_最小'),
                    data.get_value('予測寿命_最大'),
                    data.get_value('ATAF時間パーセント'),
                    data.get_value('VT回数'),
                    data.get_value('ASVS'),
                    data.get_value('ASVP'),
                    data.get_value('APVS'),
                    data.get_value('APVP'),
                    data.get_value('AMVS'),
                    data.get_value('VSオンリー'),
                    data.get_value('AMVP'),
                    data.get_value('VPオンリー'),
                    data.get_value('RVコイル'),
                    data.get_value('SVCコイル'),
                    data.get_value('VF治療回数'),
                    data.get_value('VT治療回数'),
                    data.get_value('LVインピーダンス'),
                    data.get_value('LVペーシング閾値'),
                    data.get_value('LVパルス幅')
                ))
        
        messagebox.showinfo("成功", f"{len(self.extracted_data)}件のデータを抽出しました")

    def export_to_csv(self):
        if not self.extracted_data:
            messagebox.showwarning("警告", "先にデータを抽出してください")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSVファイル", "*.csv")]
        )
        
        if file_path:
            try:
                with open(file_path, 'w', newline='', encoding='utf-8-sig') as file:
                    writer = csv.writer(file)
                    # ヘッダー行を書き込み
                    writer.writerow([
                        "患者ID", "送信日時", 
                        "心房リードインピーダンス", "心室リードインピーダンス",
                        "心房ペーシング閾値", "心房パルス幅",
                        "心室ペーシング閾値", "心室パルス幅",
                        "P波高値", "R波高値",
                        "予測寿命_最小", "予測寿命_最大",
                        "AT/AF時間%", "VT回数",
                        "AS-VS%", "AS-VP%", "AP-VS%", "AP-VP%",
                        "AMVS", "VSオンリー",
                        "AMVP", "VPオンリー", "RVコイル",
                        "SVCコイル", "VF治療回数", "VT治療回数",
                        "LVインピーダンス", "LVペーシング閾値", "LVパルス幅"
                    ])
                    
                    # データ行を書き込み
                    for data in self.extracted_data:
                        writer.writerow([
                            data.get_value('患者ID'),
                            data.get_value('送信日時'),
                            data.get_value('心房リードインピーダンス'),
                            data.get_value('心室リードインピーダンス'),
                            data.get_value('心房ペーシング閾値'),
                            data.get_value('心房パルス幅'),
                            data.get_value('心室ペーシング閾値'),
                            data.get_value('心室パルス幅'),
                            data.get_value('P波高値'),
                            data.get_value('R波高値'),
                            data.get_value('予測寿命_最小'),
                            data.get_value('予測寿命_最大'),
                            data.get_value('ATAF時間パーセント'),
                            data.get_value('VT回数'),
                            data.get_value('ASVS'),
                            data.get_value('ASVP'),
                            data.get_value('APVS'),
                            data.get_value('APVP'),
                            data.get_value('AMVS'),
                            data.get_value('VSオンリー'),
                            data.get_value('AMVP'),
                            data.get_value('VPオンリー'),
                            data.get_value('RVコイル'),
                            data.get_value('SVCコイル'),
                            data.get_value('VF治療回数'),
                            data.get_value('VT治療回数'),
                            data.get_value('LVインピーダンス'),
                            data.get_value('LVペーシング閾値'),
                            data.get_value('LVパルス幅')
                        ])
                
                messagebox.showinfo("成功", f"{len(self.extracted_data)}件のデータをCSVに出力しました")
            except Exception as e:
                messagebox.showerror("エラー", f"CSVファイルの出力中にエラーが発生しました: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = PacemakerExtractorApp(root)
    root.mainloop()
