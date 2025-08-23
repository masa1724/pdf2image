
# 1. 環境構築
## Pythonの仮想環境を作成（プロジェクトごとに独立した Python 実行環境）
python -m venv .venv

## Pythonの仮想環境を有効化
.venv\Scripts\activate

## 依存パッケージをインストール
python -m pip install -r requirements.txt


# 2. 実行
python src/main.py


# 3. ビルド
## EXE化（distフォルダ配下にexeファイルが出力される）
pyinstaller --onefile --windowed --specpath build --name PDF画像変換 src/main.py

