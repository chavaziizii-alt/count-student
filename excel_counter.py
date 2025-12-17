import os
import glob
import pandas as pd
import argparse
import sys

def count_rows_in_directory(directory):
    # بررسی وجود پوشه
    if not os.path.exists(directory):
        print(f"❌ خطا: پوشه '{directory}' پیدا نشد.")
        return

    # لیست کردن تمام فایل‌های اکسل
    # استفاده از case insensitive برای پسوندها در سیستم‌های لینوکس/مک مهم است
    extensions = ['*.xlsx', '*.xls', '*.XLSX', '*.XLS']
    excel_files = []
    for ext in extensions:
        excel_files.extend(glob.glob(os.path.join(directory, ext)))

    # حذف تکراری‌ها (اگر وجود داشته باشد)
    excel_files = list(set(excel_files))
    
    if not excel_files:
        print("⚠️ هیچ فایل اکسلی در این پوشه یافت نشد.")
        return

    print(f"🔍 تعداد {len(excel_files)} فایل اکسل پیدا شد. شروع پردازش...\n")
    
    grand_total_rows = 0

    for file_path in excel_files:
        file_name = os.path.basename(file_path)
        try:
            # sheet_name=None تمام شیت‌ها را می‌خواند
            xls_dict = pd.read_excel(file_path, sheet_name=None)
            
            file_total = 0
            sheet_info = []

            for sheet_name, df in xls_dict.items():
                count = len(df)
                file_total += count
                sheet_info.append(f"{sheet_name}: {count}")

            grand_total_rows += file_total
            print(f"✅ {file_name} -> {file_total} ردیف")
            # برای دیدن جزئیات هر شیت، خط زیر را از حالت کامنت خارج کنید
            # print(f"   └── {', '.join(sheet_info)}")

        except Exception as e:
            print(f"❌ خطا در خواندن {file_name}: {e}")

    print("-" * 40)
    print(f"🚀 مجموع نهایی رکوردها: {grand_total_rows:,}")
    print("-" * 40)

if __name__ == "__main__":
    # ایجاد قابلیت دریافت ورودی از خط فرمان
    parser = argparse.ArgumentParser(description="شمارش تعداد ردیف‌های فایل‌های اکسل در یک پوشه")
    
    # آرگومان مسیر پوشه (اختیاری - اگر وارد نشود پوشه جاری را می‌گردد)
    parser.add_argument('--path', type=str, default='.', help='مسیر پوشه حاوی فایل‌های اکسل')
    
    args = parser.parse_args()
    
    # اجرا
    target_path = args.path
    # حذف کوتیشن‌های احتمالی اگر کاربر مسیر را با " وارد کرده باشد
    target_path = target_path.strip('"').strip("'")
    
    count_rows_in_directory(target_path)
