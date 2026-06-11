"""
サンプル売上データ生成スクリプト
---------------------------------
架空の3支店 × 6ヶ月分の売上Excelファイルを input/ フォルダに生成する。
実案件で「支店から毎月届くExcel」を想定したダミーデータ。
"""

import random
from pathlib import Path

import pandas as pd

# 乱数シードを固定:誰がいつ実行しても同じデータが生成されるようにする
# (READMEのスクショと実行結果が一致し、発注者が再現確認できる)
random.seed(42)

# === 設定値はファイル冒頭にまとめる(あとから変更しやすくするため) ===
BRANCHES = ["東京支店", "大阪支店", "福岡支店"]
MONTHS = ["2026-01", "2026-02", "2026-03", "2026-04", "2026-05", "2026-06"]
CATEGORIES = ["家電", "食品", "日用品", "衣料品"]
PRODUCTS = {
    "家電": ["掃除機", "電子レンジ", "ドライヤー"],
    "食品": ["コーヒー豆", "オリーブオイル", "パスタ"],
    "日用品": ["洗剤", "ティッシュ", "シャンプー"],
    "衣料品": ["Tシャツ", "ソックス", "パーカー"],
}
PRICE_RANGE = {  # カテゴリごとの単価レンジ(円)
    "家電": (3000, 30000),
    "食品": (500, 3000),
    "日用品": (200, 1500),
    "衣料品": (1000, 8000),
}

INPUT_DIR = Path("input")  # 生成先フォルダ


def generate_branch_month(branch: str, month: str) -> pd.DataFrame:
    """1支店・1ヶ月分の売上明細(20〜40行)をDataFrameとして生成する"""
    rows = []
    for _ in range(random.randint(20, 40)):
        category = random.choice(CATEGORIES)
        product = random.choice(PRODUCTS[category])
        low, high = PRICE_RANGE[category]
        unit_price = random.randint(low, high)
        quantity = random.randint(1, 15)
        # 日付は当月内のランダムな日(1〜28日に収めて月末日の差異を回避)
        day = random.randint(1, 28)
        rows.append({
            "日付": f"{month}-{day:02d}",
            "支店": branch,
            "カテゴリ": category,
            "商品名": product,
            "単価": unit_price,
            "数量": quantity,
            "売上金額": unit_price * quantity,
        })
    # 日付順に並べ替えて実際の売上台帳らしくする
    return pd.DataFrame(rows).sort_values("日付").reset_index(drop=True)


def main():
    # 生成先フォルダがなければ自動作成(bookscraperと同じ配慮)
    INPUT_DIR.mkdir(exist_ok=True)

    count = 0
    for month in MONTHS:
        for branch in BRANCHES:
            df = generate_branch_month(branch, month)
            # ファイル名例: 売上_東京支店_2026-01.xlsx
            filename = INPUT_DIR / f"売上_{branch}_{month}.xlsx"
            df.to_excel(filename, index=False)
            count += 1
            print(f"生成: {filename}({len(df)}行)")

    print(f"\n完了: {count}ファイルを {INPUT_DIR}/ に生成しました")


if __name__ == "__main__":
    main()