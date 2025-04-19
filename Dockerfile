FROM python:latest

WORKDIR /mnt

RUN pip install --no-cache-dir openpyxl pandas

# デフォルトのCMDを指定（main.pyを自動実行するよう変更）
CMD ["python3", "main.py"]
