import tempfile
import msoffcrypto
from pathlib import Path
import os
import tkinter as tk
from tkinter import filedialog, simpledialog
from PyPDF2 import PdfReader, PdfWriter

# ダイアログボックスでディレクトリのパスを選択
root = tk.Tk()
root.withdraw()  # メインウィンドウを表示しない
file_dir_path = filedialog.askdirectory(title="PWを解除するファイルが入ったフォルダを選択してください。")

# ダイアログボックスでパスワードを入力
password = simpledialog.askstring("パスワード入力", "パスワードを入力してください：", show='*')

file_dir = Path(file_dir_path)
files = list(file_dir.glob("*"))  # All files in the directory

# PW_unlockedフォルダを作成（既に存在する場合はそのまま）
unlocked_dir = file_dir / 'PW_unlocked'
unlocked_dir.mkdir(exist_ok=True)

# ディレクトリ内にファイルがあるか確認
if not files:
    print("指定されたディレクトリにファイルが見つかりませんでした。")
else:
    # ファイルを順次確認
    for file in files:
        if file.suffix in ['.docx', '.xlsx', '.pptx']:
        # パスワード解除したテンポラリファイル作成
            with file.open("rb") as f:
                fd, path = tempfile.mkstemp()
                office_file = msoffcrypto.OfficeFile(f)
                office_file.load_key(password=password)
                with os.fdopen(fd, 'wb') as tf:
                    office_file.decrypt(tf)
                
                # 新しいファイルパスを作成 (PW_unlockedフォルダ内 + 元のファイル名)
                unlocked_file_path = unlocked_dir / file.name
                
                # パスワード解除したファイルを新たに保存
                os.rename(path, unlocked_file_path)

        elif file.suffix == '.pdf':
            with file.open("rb") as f:
                reader = PdfReader(f)
                reader.decrypt(password)

                writer = PdfWriter()
                for i in range(len(reader.pages)):
                    writer.add_page(reader.pages[i])

                unlocked_file_path = unlocked_dir / file.name
                with unlocked_file_path.open('wb') as out:
                    writer.write(out)