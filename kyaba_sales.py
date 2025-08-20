# -*- coding: utf-8 -*-
"""
キャバクラ売上サンプルデータ生成器
Rose Garden向けTableau分析用データ作成
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random
import os

print("🍸 Rose Garden サンプルデータ生成器を開始...")

try:
    # 必要なライブラリの確認
    print("ライブラリをチェック中...")
    import pandas as pd
    import numpy as np
    print("✅ 必要なライブラリは揃っています")
except ImportError as e:
    print(f"❌ 必要なライブラリが不足: {e}")
    print("次のコマンドでインストールしてください:")
    print("pip install pandas numpy")
    exit()

# =========================================
# 1. 基本設定
# =========================================

# 日本語名前データ
CUSTOMER_NAMES = [
    "田中", "佐藤", "高橋", "渡辺", "伊藤", "山田", "中村", "小林", "加藤", "吉田",
    "山本", "佐々木", "山口", "松本", "井上", "木村", "林", "斎藤", "清水", "森田"
]

CAST_NAMES = [
    "美咲", "麗子", "優香", "愛美", "聖子", "真由美", "由美子", "智子", "恵子", "裕子",
    "美香", "直子", "典子", "良子", "美穂", "千代子", "和子", "洋子", "京子", "幸子"
]

# =========================================
# 2. 顧客データ生成
# =========================================

print("👥 顧客データを生成中...")

customers = []
for i in range(1, 501):  # 500名の顧客
    # 顧客ランクの決定
    rand = random.random()
    if rand < 0.08:
        rank = 'VIP'
        visits = random.randint(8, 25)
        spend = random.randint(800000, 3000000)
        age = random.randint(35, 55)
    elif rand < 0.30:
        rank = '優良'
        visits = random.randint(4, 12)
        spend = random.randint(300000, 1200000)
        age = random.randint(30, 50)
    elif rand < 0.85:
        rank = '一般'
        visits = random.randint(1, 6)
        spend = random.randint(80000, 400000)
        age = random.randint(25, 45)
    else:
        rank = '新規'
        visits = random.randint(1, 3)
        spend = random.randint(30000, 150000)
        age = random.randint(23, 40)
    
    # 登録日
    if rank == '新規':
        reg_date = datetime.now() - timedelta(days=random.randint(1, 90))
    else:
        reg_date = datetime.now() - timedelta(days=random.randint(30, 730))
    
    customer = {
        'customer_id': i,
        'customer_name': f"{random.choice(CUSTOMER_NAMES)}_{i:03d}",
        'customer_rank': rank,
        'registration_date': reg_date.strftime('%Y-%m-%d'),
        'birth_year': datetime.now().year - age,
        'age': age,
        'occupation_category': random.choice(['経営者', 'サラリーマン', '医師', 'IT関係', '金融関係']),
        'total_visits': visits,
        'total_spent': spend,
        'last_visit_date': (datetime.now() - timedelta(days=random.randint(1, 60))).strftime('%Y-%m-%d'),
        'status': 'active'
    }
    customers.append(customer)

customers_df = pd.DataFrame(customers)
print(f"✅ 顧客データ {len(customers_df)} 件生成完了")

# =========================================
# 3. キャストデータ生成
# =========================================

print("⭐ キャストデータを生成中...")

casts = []
cast_types = ['知的系', '癒し系', 'ギャル系', 'お姉さん系', '妹系']

for i in range(1, 31):  # 30名のキャスト
    hire_date = datetime.now() - timedelta(days=random.randint(30, 1095))
    experience_months = max(1, (datetime.now() - hire_date).days // 30)
    
    # 経験に応じた設定
    if experience_months >= 24:
        hourly_rate = random.randint(5000, 8000)
        nominations = random.randint(200, 600)
        rating = round(random.uniform(4.2, 5.0), 2)
    elif experience_months >= 12:
        hourly_rate = random.randint(4000, 6000)
        nominations = random.randint(100, 400)
        rating = round(random.uniform(3.8, 4.8), 2)
    else:
        hourly_rate = random.randint(3000, 4500)
        nominations = random.randint(20, 200)
        rating = round(random.uniform(3.5, 4.5), 2)
    
    cast = {
        'cast_id': i,
        'cast_name': f"{random.choice(CAST_NAMES)}_{i:02d}",
        'hire_date': hire_date.strftime('%Y-%m-%d'),
        'cast_type': random.choice(cast_types),
        'experience_months': experience_months,
        'hourly_rate': hourly_rate,
        'total_nominations': nominations,
        'average_rating': rating,
        'status': 'active'
    }
    casts.append(cast)

casts_df = pd.DataFrame(casts)
print(f"✅ キャストデータ {len(casts_df)} 件生成完了")

# =========================================
# 4. 売上データ生成
# =========================================

print("💰 売上データを生成中...")

sales = []
service_types = ['通常セット', 'プレミアムセット', 'VIPコース', '延長コース', 'ボトルキープ']
payment_methods = ['現金', 'カード', '掛け', '電子マネー']

# 過去1年間のデータ
start_date = datetime.now() - timedelta(days=365)

for i in range(1, 10001):  # 10,000件の売上
    # 日付生成（週末に偏重）
    sale_date = start_date + timedelta(days=random.randint(0, 365))
    weekday = sale_date.weekday()
    
    # 週末の確率を上げる
    if weekday >= 5:  # 土日
        weight = 3
    elif weekday == 4:  # 金曜
        weight = 2
    else:  # 平日
        weight = 1
    
    if random.randint(1, 6) > weight:
        continue
    
    # 時間設定（19:00-26:00、22-24時がピーク）
    hour_weights = [5, 8, 12, 15, 18, 20, 15, 7]  # 19-26時の重み
    hour = random.choices(range(19, 27), weights=hour_weights)[0]
    if hour >= 24:
        hour -= 24
        sale_date += timedelta(days=1)
    
    minute = random.randint(0, 59)
    sale_time = f"{hour:02d}:{minute:02d}:00"
    
    # 顧客とキャストの選択
    customer = customers_df.sample(1).iloc[0]
    cast = casts_df.sample(1).iloc[0]
    
    # サービスタイプの選択（顧客ランクに応じて）
    if customer['customer_rank'] == 'VIP':
        service_type = random.choices(service_types, weights=[1, 3, 4, 1.5, 0.5])[0]
    elif customer['customer_rank'] == '優良':
        service_type = random.choices(service_types, weights=[2, 4, 2.5, 1, 0.5])[0]
    else:
        service_type = random.choices(service_types, weights=[5, 3, 1, 0.8, 0.2])[0]
    
    # 料金計算
    base_prices = {
        '通常セット': 8000,
        'プレミアムセット': 15000,
        'VIPコース': 25000,
        '延長コース': 10000,
        'ボトルキープ': 30000
    }
    
    # 顧客ランクによる倍率
    if customer['customer_rank'] == 'VIP':
        multiplier = random.uniform(2.5, 4.0)
    elif customer['customer_rank'] == '優良':
        multiplier = random.uniform(1.5, 2.8)
    elif customer['customer_rank'] == '一般':
        multiplier = random.uniform(0.8, 1.5)
    else:  # 新規
        multiplier = random.uniform(0.6, 1.2)
    
    base_charge = int(base_prices[service_type] * multiplier)
    drink_charge = random.randint(3000, 20000)
    
    # 指名料（30%の確率）
    nomination_fee = random.randint(2000, 8000) if random.random() < 0.3 else 0
    
    # 延長料金（20%の確率）
    if random.random() < 0.2:
        extension_fee = random.randint(5000, 20000)
        duration = random.randint(120, 360)
    else:
        extension_fee = 0
        duration = random.randint(60, 180)
    
    total_amount = base_charge + drink_charge + nomination_fee + extension_fee
    
    sale = {
        'sale_id': i,
        'customer_id': customer['customer_id'],
        'cast_id': cast['cast_id'],
        'sale_date': sale_date.strftime('%Y-%m-%d'),
        'sale_time': sale_time,
        'service_type': service_type,
        'base_charge': base_charge,
        'drink_charge': drink_charge,
        'nomination_fee': nomination_fee,
        'extension_fee': extension_fee,
        'total_amount': total_amount,
        'payment_method': random.choice(payment_methods),
        'duration_minutes': duration
    }
    sales.append(sale)

sales_df = pd.DataFrame(sales)
print(f"✅ 売上データ {len(sales_df)} 件生成完了")

# =========================================
# 5. 統計情報表示
# =========================================

print("\n" + "="*50)
print("📊 生成データの統計情報")
print("="*50)

print(f"\n👥 顧客統計:")
print(f"総顧客数: {len(customers_df):,}名")
rank_counts = customers_df['customer_rank'].value_counts()
for rank, count in rank_counts.items():
    print(f"  {rank}: {count:,}名 ({count/len(customers_df)*100:.1f}%)")

print(f"\n⭐ キャスト統計:")
print(f"総キャスト数: {len(casts_df):,}名")
print(f"平均時給: ¥{casts_df['hourly_rate'].mean():,.0f}")
print(f"平均評価: {casts_df['average_rating'].mean():.2f}/5.0")

print(f"\n💰 売上統計:")
print(f"総取引数: {len(sales_df):,}件")
print(f"総売上: ¥{sales_df['total_amount'].sum():,}")
print(f"平均単価: ¥{sales_df['total_amount'].mean():,.0f}")

print(f"\nサービス別売上:")
service_stats = sales_df.groupby('service_type')['total_amount'].agg(['count', 'sum'])
for service in service_stats.index:
    count = service_stats.loc[service, 'count']
    total = service_stats.loc[service, 'sum']
    print(f"  {service}: {count:,}件, ¥{total:,}")

# =========================================
# 6. ファイル保存
# =========================================

print(f"\n📁 CSVファイルに保存中...")

try:
    customers_df.to_csv('rose_garden_customers.csv', index=False, encoding='utf-8-sig')
    casts_df.to_csv('rose_garden_casts.csv', index=False, encoding='utf-8-sig')
    sales_df.to_csv('rose_garden_sales.csv', index=False, encoding='utf-8-sig')
    
    print("✅ ファイル保存完了!")
    print("📄 rose_garden_customers.csv")
    print("📄 rose_garden_casts.csv") 
    print("📄 rose_garden_sales.csv")
    
    # ファイルサイズ確認
    for filename in ['rose_garden_customers.csv', 'rose_garden_casts.csv', 'rose_garden_sales.csv']:
        if os.path.exists(filename):
            size = os.path.getsize(filename)
            print(f"   {filename}: {size:,} bytes")

except Exception as e:
    print(f"❌ ファイル保存エラー: {e}")

# =========================================
# 7. Tableau使用ガイド
# =========================================

print(f"\n🎯 Tableauでの使用方法:")
print("1. Tableauを起動")
print("2. [データに接続] → [ファイル] → [テキストファイル]")
print("3. 'rose_garden_sales.csv' を選択")
print("4. [追加] で他のテーブルも追加:")
print("   - rose_garden_customers.csv (customer_id で結合)")
print("   - rose_garden_casts.csv (cast_id で結合)")
print("5. データ型確認後、分析開始!")

print(f"\n推奨される最初のチャート:")
print("📈 月別売上推移 (線グラフ)")
print("🎯 顧客ランク別売上 (円グラフ)")
print("⭐ キャスト別パフォーマンス (棒グラフ)")
print("🕒 時間帯別来店数 (ヒートマップ)")

print(f"\n🎉 サンプルデータ生成完了!")
print(f"現在のフォルダに3つのCSVファイルが作成されました。")
print(f"Tableauでの分析を開始してください!")
