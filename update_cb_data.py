#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AI智能可轉債競拍系統 v3.0 - 資料更新腳本
從Excel檔案更新JSON資料庫

使用方式：
python update_cb_data.py 00_CB代號名稱.xlsx

作者：Rex大叔的168台股投資教室
日期：2025-11-04
"""

import pandas as pd
import numpy as np
import json
from datetime import datetime, timedelta
import sys
import os

def calculate_statistics(df):
    """計算286檔競拍案例的完整統計"""
    
    print("=" * 80)
    print("開始計算統計數據...")
    print("=" * 80)
    
    statistics = {
        "資料來源": "04_所有CB競拍資料庫",
        "資料期間": {
            "開始": df['開標日期'].min().strftime('%Y-%m-%d'),
            "結束": df['開標日期'].max().strftime('%Y-%m-%d'),
            "總筆數": len(df)
        },
        "更新時間": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        "維度統計": {}
    }
    
    print(f"✓ 資料期間: {statistics['資料期間']['開始']} ~ {statistics['資料期間']['結束']}")
    print(f"✓ 總筆數: {statistics['資料期間']['總筆數']}")
    
    # 1. 產業統計
    print("\n計算產業統計...")
    industry_stats = {}
    for industry in df['產業分類'].unique():
        industry_df = df[df['產業分類'] == industry]
        industry_stats[industry] = {
            "樣本數": len(industry_df),
            "平均最低得標": round(industry_df['最低得標'].mean(), 2),
            "平均最低溢價": round(industry_df['最低溢價'].mean() * 100, 2),
            "平均得標價": round(industry_df['平均得標'].mean(), 2),
            "平均溢價": round(industry_df['平均溢價'].mean() * 100, 2)
        }
    statistics['維度統計']['產業'] = industry_stats
    print(f"✓ 完成 {len(industry_stats)} 個產業統計")
    
    # 2. 發行規模統計
    print("\n計算發行規模統計...")
    size_ranges = {
        '<2億': (0, 2),
        '2-5億': (2, 5),
        '5-10億': (5, 10),
        '10-15億': (10, 15),
        '15-20億': (15, 20),
        '>20億': (20, float('inf'))
    }
    
    size_stats = {}
    for label, (min_val, max_val) in size_ranges.items():
        if max_val == float('inf'):
            size_df = df[df['發行規模'] >= min_val]
        else:
            size_df = df[(df['發行規模'] >= min_val) & (df['發行規模'] < max_val)]
        
        if len(size_df) > 0:
            size_stats[label] = {
                "樣本數": len(size_df),
                "平均最低得標": round(size_df['最低得標'].mean(), 2),
                "平均最低溢價": round(size_df['最低溢價'].mean() * 100, 2)
            }
    statistics['維度統計']['發行規模'] = size_stats
    print(f"✓ 完成 {len(size_stats)} 個規模區間統計")
    
    # 3. 股本統計
    print("\n計算股本統計...")
    capital_ranges = {
        '<3億': (0, 3),
        '3-6億': (3, 6),
        '6-10億': (6, 10),
        '10-15億': (10, 15),
        '15-20億': (15, 20),
        '>20億': (20, float('inf'))
    }
    
    capital_stats = {}
    for label, (min_val, max_val) in capital_ranges.items():
        if max_val == float('inf'):
            capital_df = df[df['股本'] >= min_val]
        else:
            capital_df = df[(df['股本'] >= min_val) & (df['股本'] < max_val)]
        
        if len(capital_df) > 0:
            capital_stats[label] = {
                "樣本數": len(capital_df),
                "平均最低得標": round(capital_df['最低得標'].mean(), 2),
                "平均最低溢價": round(capital_df['最低溢價'].mean() * 100, 2)
            }
    statistics['維度統計']['股本'] = capital_stats
    print(f"✓ 完成 {len(capital_stats)} 個股本區間統計")
    
    # 4. 年期統計
    print("\n計算年期統計...")
    years_stats = {}
    for years in sorted(df['年期'].unique()):
        years_df = df[df['年期'] == years]
        years_stats[f"{years}年"] = {
            "樣本數": len(years_df),
            "平均最低得標": round(years_df['最低得標'].mean(), 2),
            "平均最低溢價": round(years_df['最低溢價'].mean() * 100, 2)
        }
    statistics['維度統計']['年期'] = years_stats
    print(f"✓ 完成 {len(years_stats)} 個年期統計")
    
    # 5. 擔保統計
    print("\n計算擔保統計...")
    guarantee_stats = {}
    for guarantee in df['擔保'].unique():
        guarantee_df = df[df['擔保'] == guarantee]
        guarantee_stats[guarantee] = {
            "樣本數": len(guarantee_df),
            "平均最低得標": round(guarantee_df['最低得標'].mean(), 2),
            "平均最低溢價": round(guarantee_df['最低溢價'].mean() * 100, 2)
        }
    statistics['維度統計']['擔保'] = guarantee_stats
    print(f"✓ 完成 {len(guarantee_stats)} 個擔保類型統計")
    
    # 6. 信評統計
    print("\n計算信評統計...")
    rating_stats = {}
    for rating_group in ['2-3分', '4-5分', '6-7分', '8-9分', 'BBB']:
        if rating_group == 'BBB':
            rating_df = df[df['信評'] == 'BBB']
        else:
            min_val = int(rating_group[0])
            max_val = int(rating_group[2])
            rating_df = df[df['信評'].isin([min_val, max_val])]
        
        if len(rating_df) > 0:
            rating_stats[rating_group] = {
                "樣本數": len(rating_df),
                "平均最低得標": round(rating_df['最低得標'].mean(), 2),
                "平均最低溢價": round(rating_df['最低溢價'].mean() * 100, 2)
            }
    statistics['維度統計']['信評'] = rating_stats
    print(f"✓ 完成 {len(rating_stats)} 個信評等級統計")
    
    # 7. 轉換價統計
    print("\n計算轉換價統計...")
    conversion_ranges = {
        '<20元': (0, 20),
        '20-50元': (20, 50),
        '50-100元': (50, 100),
        '100-150元': (100, 150),
        '150-200元': (150, 200),
        '>200元': (200, float('inf'))
    }
    
    conversion_stats = {}
    for label, (min_val, max_val) in conversion_ranges.items():
        if max_val == float('inf'):
            conv_df = df[df['轉換價'] >= min_val]
        else:
            conv_df = df[(df['轉換價'] >= min_val) & (df['轉換價'] < max_val)]
        
        if len(conv_df) > 0:
            conversion_stats[label] = {
                "樣本數": len(conv_df),
                "平均最低得標": round(conv_df['最低得標'].mean(), 2),
                "平均最低溢價": round(conv_df['最低溢價'].mean() * 100, 2)
            }
    statistics['維度統計']['轉換價'] = conversion_stats
    print(f"✓ 完成 {len(conversion_stats)} 個轉換價區間統計")
    
    # 8. 理論價統計
    print("\n計算理論價統計...")
    theoretical_ranges = {
        '<85元': (0, 85),
        '85-90元': (85, 90),
        '90-95元': (90, 95),
        '95-98元': (95, 98),
        '98-100元': (98, 100),
        '100-102元': (100, 102),
        '102-105元': (102, 105),
        '105-110元': (105, 110),
        '>110元': (110, float('inf'))
    }
    
    theoretical_stats = {}
    for label, (min_val, max_val) in theoretical_ranges.items():
        if max_val == float('inf'):
            theo_df = df[df['理論價'] >= min_val]
        else:
            theo_df = df[(df['理論價'] >= min_val) & (df['理論價'] < max_val)]
        
        if len(theo_df) > 0:
            theoretical_stats[label] = {
                "樣本數": len(theo_df),
                "平均最低得標": round(theo_df['最低得標'].mean(), 2),
                "平均最低溢價": round(theo_df['最低溢價'].mean() * 100, 2)
            }
    statistics['維度統計']['理論價'] = theoretical_stats
    print(f"✓ 完成 {len(theoretical_stats)} 個理論價區間統計")
    
    # 9. 市場氛圍
    print("\n計算市場氛圍...")
    latest_date = df['開標日期'].max()
    
    def calculate_market_atmosphere(days, label):
        cutoff_date = latest_date - timedelta(days=days)
        period_df = df[df['開標日期'] >= cutoff_date].copy()
        
        if len(period_df) == 0:
            return None
        
        period_df = period_df.sort_values('開標日期')
        mid_point = len(period_df) // 2
        
        first_half = period_df.iloc[:mid_point]
        second_half = period_df.iloc[mid_point:]
        
        first_avg = first_half['最低溢價'].mean() * 100 if len(first_half) > 0 else 0
        second_avg = second_half['最低溢價'].mean() * 100 if len(second_half) > 0 else 0
        
        change = second_avg - first_avg
        
        if change > 2:
            trend = "強勢上升"
            adjustment = 0.7
        elif change > 1:
            trend = "溫和上升"
            adjustment = 0.4
        elif change > -1:
            trend = "平穩"
            adjustment = 0.0
        elif change > -2:
            trend = "溫和下降"
            adjustment = -0.3
        else:
            trend = "急劇下降"
            adjustment = -0.6
        
        return {
            "案例數": len(period_df),
            "平均最低溢價": round(period_df['最低溢價'].mean() * 100, 2),
            "趨勢": trend,
            "建議調整": adjustment,
            "開始日期": cutoff_date.strftime('%Y-%m-%d')
        }
    
    market_atmosphere = {}
    for days, label in [(30, "近1月"), (90, "近3月"), (180, "近6月"), (365, "近1年")]:
        result = calculate_market_atmosphere(days, label)
        if result:
            market_atmosphere[label] = result
            print(f"  {label}: {result['案例數']}檔, 趨勢={result['趨勢']}")
    
    statistics['市場氛圍'] = market_atmosphere
    print(f"✓ 完成市場氛圍分析")
    
    return statistics


def build_cb_database(cb_names_df):
    """建立CB名稱資料庫"""
    
    print("\n建立CB資料庫...")
    print("=" * 80)
    
    cb_dict = {}
    for _, row in cb_names_df.iterrows():
        cb_code = str(row['代號'])
        cb_dict[cb_code] = {
            "股票代號": int(row['股票代號']),
            "代號": int(row['代號']),
            "名稱": row['名稱'],
            "掛牌最高": float(row['掛牌最高']) if pd.notna(row['掛牌最高']) else None,
            "掛牌最低": float(row['掛牌最低']) if pd.notna(row['掛牌最低']) else None,
            "資金用途": row['資金用途'] if pd.notna(row['資金用途']) else None
        }
    
    print(f"✓ 完成 {len(cb_dict)} 檔CB資料庫建立")
    return cb_dict


def main():
    """主程式"""
    
    if len(sys.argv) < 2:
        print("使用方式: python update_cb_data.py <Excel檔案路徑>")
        print("範例: python update_cb_data.py 00_CB代號名稱.xlsx")
        sys.exit(1)
    
    excel_file = sys.argv[1]
    
    if not os.path.exists(excel_file):
        print(f"錯誤：找不到檔案 {excel_file}")
        sys.exit(1)
    
    print("=" * 80)
    print("AI智能可轉債競拍系統 v3.0 - 資料更新腳本")
    print("=" * 80)
    print(f"\n讀取Excel檔案: {excel_file}")
    
    try:
        # 讀取競拍資料庫
        print("\n讀取 04_所有CB競拍資料庫...")
        df_auction = pd.read_excel(excel_file, sheet_name='04_所有CB競拍資料庫')
        print(f"✓ 讀取 {len(df_auction)} 筆競拍資料")
        
        # 讀取CB名稱資料
        print("\n讀取 00_CB代號名稱過濾及掛牌高低...")
        df_names = pd.read_excel(excel_file, 
                                 sheet_name='00_CB代號名稱過濾及掛牌高低_20251031')
        print(f"✓ 讀取 {len(df_names)} 筆CB資料")
        
        # 計算統計數據
        statistics = calculate_statistics(df_auction)
        
        # 建立CB資料庫
        cb_database = build_cb_database(df_names)
        
        # 整合資料
        print("\n整合資料...")
        print("=" * 80)
        integrated_data = {
            "統計數據": statistics,
            "CB資料庫": cb_database
        }
        
        # 輸出JSON檔案
        output_file = 'cb_data_integrated.json'
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(integrated_data, f, ensure_ascii=False, indent=2)
        
        file_size = os.path.getsize(output_file)
        
        print(f"\n✓ 資料已輸出至: {output_file}")
        print(f"✓ 檔案大小: {file_size:,} bytes ({file_size/1024:.1f} KB)")
        print("\n" + "=" * 80)
        print("✓ 更新完成！")
        print("=" * 80)
        
        print("\n請將以下檔案放在一起：")
        print("  1. index_v3_updated.html")
        print(f"  2. {output_file}")
        print("\n然後用瀏覽器開啟 index_v3_updated.html 即可使用！")
        
    except Exception as e:
        print(f"\n錯誤：{str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()
