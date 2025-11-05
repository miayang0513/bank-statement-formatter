#!/usr/bin/env python3
"""
銀行明細統一格式化工具
支持4家銀行：Monzo, Revolut, Wise, Amex
"""

import pandas as pd
import os
import sys
from pathlib import Path
from datetime import datetime


def read_monzo(file_path):
    """讀取Monzo CSV文件"""
    df = pd.read_csv(file_path)
    
    # 統一格式
    result = []
    for _, row in df.iterrows():
        # 合併日期和時間
        date_str = str(row['Date'])
        time_str = str(row['Time']) if pd.notna(row['Time']) else '00:00:00'
        
        # 解析日期
        try:
            if '/' in date_str:
                date_obj = datetime.strptime(f"{date_str} {time_str}", "%d/%m/%Y %H:%M:%S")
            else:
                date_obj = datetime.strptime(f"{date_str} {time_str}", "%Y-%m-%d %H:%M:%S")
        except:
            continue
        
        # 金額：Amount × -1
        amount = float(row['Amount']) if pd.notna(row['Amount']) else 0
        amount = amount * -1  # 取負值
        
        # 項目名稱：Name + Description
        name = str(row['Name']) if pd.notna(row['Name']) else ''
        description = str(row['Description']) if pd.notna(row['Description']) else ''
        # 合併Name和Description
        if name and description:
            item_name = f"{name} {description}"
        elif name:
            item_name = name
        elif description:
            item_name = description
        else:
            item_name = ''
        
        # Type
        type_val = str(row['Type']) if pd.notna(row['Type']) else ''
        
        # Notes and #tags
        notes = str(row['Notes and #tags']) if pd.notna(row['Notes and #tags']) else ''
        
        result.append({
            'date': date_obj,
            'amount': amount,
            'currency': str(row['Currency']) if pd.notna(row['Currency']) else 'GBP',
            'description': item_name,  # 項目名稱：Name + Description
            'type': type_val,
            'notes': notes,
            'category': str(row['Category']) if pd.notna(row['Category']) else 'General',
            'bank': 'Monzo',
            'transaction_id': str(row['Transaction ID']) if pd.notna(row['Transaction ID']) else ''
        })
    
    return pd.DataFrame(result)


def read_revolut(file_path):
    """讀取Revolut CSV文件"""
    df = pd.read_csv(file_path)
    
    result = []
    for _, row in df.iterrows():
        # 使用Started Date
        date_str = str(row['Started Date'])
        try:
            date_obj = datetime.strptime(date_str, "%Y-%m-%d %H:%M:%S")
        except:
            continue
        
        # 金額：Amount x -1
        amount = float(row['Amount']) if pd.notna(row['Amount']) else 0
        amount = amount * -1  # 取負值
        
        # 描述
        description = str(row['Description']) if pd.notna(row['Description']) else ''
        
        result.append({
            'date': date_obj,
            'amount': amount,
            'currency': str(row['Currency']) if pd.notna(row['Currency']) else 'GBP',
            'description': description,
            'category': 'General',
            'bank': 'Revolut',
            'type': str(row['Type']) if pd.notna(row['Type']) else '',
            'transaction_id': ''
        })
    
    return pd.DataFrame(result)


def read_wise(file_path):
    """讀取Wise CSV文件"""
    df = pd.read_csv(file_path)
    
    result = []
    for _, row in df.iterrows():
        # 處理COMPLETED和REFUNDED狀態的交易
        status = str(row['Status'])
        if status not in ['COMPLETED', 'REFUNDED']:
            continue
        
        # 使用Created on日期
        date_str = str(row['Created on'])
        try:
            date_obj = datetime.strptime(date_str, "%Y-%m-%d %H:%M:%S")
        except:
            continue
        
        # 方向
        direction = str(row['Direction'])
        if direction not in ['IN', 'OUT']:
            continue
        
        # 金額：都取自Source amount (after fees)
        # IN為負數，OUT為正數
        source_amount = float(row['Source amount (after fees)']) if pd.notna(row['Source amount (after fees)']) else 0
        if direction == 'IN':
            amount = -source_amount  # IN為負數
        else:  # OUT
            amount = source_amount  # OUT為正數
        
        currency = str(row['Source currency']) if pd.notna(row['Source currency']) else 'GBP'
        
        # 項目名稱：Target name
        target_name = str(row['Target name']) if pd.notna(row['Target name']) else ''
        
        # Reference
        reference = str(row['Reference']) if pd.notna(row['Reference']) else ''
        
        # Source name
        source_name = str(row['Source name']) if pd.notna(row['Source name']) else ''
        
        result.append({
            'date': date_obj,
            'amount': amount,
            'currency': currency,
            'description': target_name,  # 項目名稱用Target name
            'reference': reference,
            'source_name': source_name,
            'category': str(row['Category']) if pd.notna(row['Category']) else 'General',
            'bank': 'Wise',
            'type': str(row['ID']) if pd.notna(row['ID']) else '',
            'transaction_id': str(row['ID']) if pd.notna(row['ID']) else ''
        })
    
    return pd.DataFrame(result)


def read_amex(file_path):
    """讀取Amex XLSX文件"""
    # 表頭在第6行（索引6）
    df = pd.read_excel(file_path, header=6)
    
    result = []
    for _, row in df.iterrows():
        # 解析日期（格式：DD/MM/YYYY）
        date_val = row['Date']
        if pd.isna(date_val):
            continue
        
        try:
            # 處理日期格式 DD/MM/YYYY
            if isinstance(date_val, str):
                date_obj = datetime.strptime(date_val, "%d/%m/%Y")
            else:
                # 如果是pandas的datetime類型，直接使用
                date_obj = pd.to_datetime(date_val)
        except:
            continue
        
        # 金額：直接使用Amount（不轉換）
        amount = float(row['Amount']) if pd.notna(row['Amount']) else 0
        
        # 描述
        description = str(row['Description']) if pd.notna(row['Description']) else ''
        
        # 類別
        category = str(row['Category']) if pd.notna(row['Category']) else 'General'
        
        # 地址
        address = str(row['Address']) if pd.notna(row['Address']) else ''
        
        result.append({
            'date': date_obj,
            'amount': amount,
            'currency': 'GBP',  # Amex文件通常是GBP
            'description': description,
            'category': category,
            'address': address,
            'bank': 'Amex',
            'type': '',
            'transaction_id': ''
        })
    
    return pd.DataFrame(result)


def format_monzo_output(df):
    """
    格式化Monzo的输出
    格式: | 時間 | 項目名稱 | 空 | 空 | 金額 (需有幣別) | 空 | 空 | 空 | 空 | 空 | 備註 |
    
    輸出邏輯：
    - 時間：Date + Time (YYYY-MM-DD HH:MM:SS格式)
    - 項目名稱：Name + Description
    - 金額：Amount × -1 (需有幣別)
    - 備註：Type - Notes and #tags
    """
    # 格式化日期为字符串 (YYYY-MM-DD HH:MM:SS格式)
    df['date'] = df['date'].dt.strftime('%Y-%m-%d %H:%M:%S')
    
    # 格式化金額（保留2位小數）
    df['amount'] = df['amount'].round(2)
    
    # 創建輸出格式
    output_df = pd.DataFrame({
        '時間': df['date'],
        '項目名稱': df['description'].fillna(''),
        '': '',
        ' ': '',
        '金額': df.apply(lambda row: f"{row['amount']:.2f}", axis=1),
        '  ': '',
        '   ': '',
        '    ': '',
        '     ': '',
        '      ': '',
        '備註': df.apply(lambda row: f"{str(row['type']) if pd.notna(row['type']) else ''} - {str(row['notes']) if pd.notna(row['notes']) else ''}".strip(' -'), axis=1)
    })
    
    return output_df


def format_revolut_output(df):
    """
    格式化Revolut的输出
    格式: | 時間 | 項目名稱 | 空 | 空 | 金額 (需有幣別) | 空 | 空 | 空 | 空 | 空 | 備註 |
    
    輸出邏輯：
    - 時間：Started Date (YYYY-MM-DD HH:MM:SS格式)
    - 項目名稱：Description
    - 金額：Amount x -1 (需有幣別)
    - 備註：Description
    """
    # 格式化日期为字符串 (YYYY-MM-DD HH:MM:SS格式)
    df['date'] = df['date'].dt.strftime('%Y-%m-%d %H:%M:%S')
    
    # 格式化金額（保留2位小數）
    df['amount'] = df['amount'].round(2)
    
    # 創建輸出格式
    output_df = pd.DataFrame({
        '時間': df['date'],
        '項目名稱': df['description'].fillna(''),
        '': '',
        ' ': '',
        '金額': df.apply(lambda row: f"{row['amount']:.2f}", axis=1),
        '  ': '',
        '   ': '',
        '    ': '',
        '     ': '',
        '      ': '',
        '備註': df['description'].fillna('')  # 備註使用Description
    })
    
    return output_df


def format_wise_output(df):
    """
    格式化Wise的输出
    格式: | 時間 | 項目名稱 | 空 | 空 | 金額 (需有幣別) | 空 | 空 | 空 | 空 | 空 | 備註 |
    
    輸出邏輯：
    - 時間：Created on (YYYY-MM-DD HH:MM:SS格式)
    - 項目名稱：Target name
    - 金額：Source amount (after fees) + Source currency (IN為負數，OUT為正數)
    - 備註：Reference + Source name
    """
    # 格式化日期为字符串 (YYYY-MM-DD HH:MM:SS格式)
    df['date'] = df['date'].dt.strftime('%Y-%m-%d %H:%M:%S')
    
    # 格式化金額（保留2位小數）
    df['amount'] = df['amount'].round(2)
    
    # 創建輸出格式
    output_df = pd.DataFrame({
        '時間': df['date'],
        '項目名稱': df['description'].fillna(''),
        '': '',
        ' ': '',
        '金額': df.apply(lambda row: f"{row['amount']:.2f}", axis=1),
        '  ': '',
        '   ': '',
        '    ': '',
        '     ': '',
        '      ': '',
        '備註': df.apply(lambda row: f"{str(row['reference']) if pd.notna(row['reference']) else ''} {str(row['source_name']) if pd.notna(row['source_name']) else ''}".strip(), axis=1)
    })
    
    return output_df


def format_amex_output(df):
    """
    格式化Amex的输出
    格式: | 時間 | 項目名稱 | 空 | 空 | 金額 (需有幣別) | 空 | 空 | 空 | 空 | 空 | 備註 |
    
    輸出邏輯：
    - 時間：Date (YYYY-MM-DD HH:MM:SS格式，時間設為00:00:00)
    - 項目名稱：Description
    - 金額：Amount (需有幣別)
    - 備註：Category + Address
    """
    # 格式化日期为字符串 (YYYY-MM-DD HH:MM:SS格式，時間設為00:00:00)
    df['date'] = df['date'].dt.strftime('%Y-%m-%d 00:00:00')
    
    # 格式化金額（保留2位小數）
    df['amount'] = df['amount'].round(2)
    
    # 創建輸出格式
    output_df = pd.DataFrame({
        '時間': df['date'],
        '項目名稱': df['description'].fillna(''),
        '': '',
        ' ': '',
        '金額': df.apply(lambda row: f"{row['amount']:.2f}", axis=1),
        '  ': '',
        '   ': '',
        '    ': '',
        '     ': '',
        '      ': '',
        '備註': df.apply(lambda row: f"{row['category']} {row['address']}".strip() if pd.notna(row.get('address', '')) and str(row['address']).strip() else row['category'], axis=1)
    })
    
    return output_df


def format_statements(month_dir):
    """格式化指定月份的銀行明細"""
    month_dir = Path(month_dir)
    
    if not month_dir.exists():
        print(f"錯誤：目錄 {month_dir} 不存在")
        return None
    
    all_formatted_outputs = []
    
    # 讀取并格式化Monzo
    monzo_files = list(month_dir.glob('monzo*.csv'))
    if monzo_files:
        print(f"讀取 Monzo 文件: {monzo_files[0].name}")
        try:
            df = read_monzo(monzo_files[0])
            formatted_df = format_monzo_output(df)
            all_formatted_outputs.append(formatted_df)
            print(f"  ✓ Monzo: {len(formatted_df)} 筆交易")
        except Exception as e:
            print(f"  ✗ 警告：讀取Monzo文件失敗: {e}")
    
    # 讀取并格式化Revolut
    revolut_files = list(month_dir.glob('revolut*.csv'))
    if revolut_files:
        print(f"讀取 Revolut 文件: {revolut_files[0].name}")
        try:
            df = read_revolut(revolut_files[0])
            formatted_df = format_revolut_output(df)
            all_formatted_outputs.append(formatted_df)
            print(f"  ✓ Revolut: {len(formatted_df)} 筆交易")
        except Exception as e:
            print(f"  ✗ 警告：讀取Revolut文件失敗: {e}")
    
    # 讀取并格式化Wise
    wise_files = list(month_dir.glob('wise*.csv'))
    if wise_files:
        print(f"讀取 Wise 文件: {wise_files[0].name}")
        try:
            df = read_wise(wise_files[0])
            formatted_df = format_wise_output(df)
            all_formatted_outputs.append(formatted_df)
            print(f"  ✓ Wise: {len(formatted_df)} 筆交易")
        except Exception as e:
            print(f"  ✗ 警告：讀取Wise文件失敗: {e}")
    
    # 讀取并格式化Amex
    amex_files = list(month_dir.glob('amex*.xlsx'))
    if amex_files:
        print(f"讀取 Amex 文件: {amex_files[0].name}")
        try:
            df = read_amex(amex_files[0])
            formatted_df = format_amex_output(df)
            all_formatted_outputs.append(formatted_df)
            print(f"  ✓ Amex: {len(formatted_df)} 筆交易")
        except Exception as e:
            print(f"  ✗ 警告：讀取Amex文件失敗: {e}")
    
    if not all_formatted_outputs:
        print("錯誤：沒有找到任何銀行明細文件")
        return None
    
    # 合併所有格式化後的輸出（不排序，保持各銀行的原始順序）
    combined_df = pd.concat(all_formatted_outputs, ignore_index=True)
    
    # 不排序，保持各銀行數據的原始順序，方便檢查
    # combined_df = combined_df.sort_values('時間')
    
    return combined_df


def main():
    """主函數"""
    if len(sys.argv) < 2:
        print("用法: python formatter.py <月份目錄>")
        print("例如: python formatter.py statements/202510")
        sys.exit(1)
    
    month_dir = sys.argv[1]
    
    print(f"處理月份目錄: {month_dir}")
    print("-" * 50)
    
    df = format_statements(month_dir)
    
    if df is None or df.empty:
        print("錯誤：沒有數據可以輸出")
        sys.exit(1)
    
    # 輸出文件
    output_file = Path(month_dir) / 'combined_statements.csv'
    
    # 使用utf-8-sig編碼（帶BOM），確保Google Spreadsheet能正確識別中文
    # 使用逗號分隔，這是Google Spreadsheet的標準格式
    df.to_csv(output_file, index=False, encoding='utf-8-sig', lineterminator='\n')
    
    print("-" * 50)
    print(f"成功！已生成合併的明細文件: {output_file}")
    print(f"總共 {len(df)} 筆交易")
    print(f"日期範圍: {df['時間'].min()} 到 {df['時間'].max()}")
    print(f"\n💡 提示：可以直接將 {output_file.name} 上傳到 Google Spreadsheet")


if __name__ == '__main__':
    main()

