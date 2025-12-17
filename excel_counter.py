import os
import glob
import pandas as pd
import argparse

def count_rows_in_directory(directory):
    # بررسی وجود پوشه
    if not os.path.exists(directory):
        print(f"❌ خطا: پوشه '{directory}' پیدا نشد.")
        print("💡 لطفا مطمئن شوید پوشه‌ای با این نام ساخته‌اید.")
        return

    # لیست کردن تمام فایل‌های اکسل
    extensions = ['*.xlsx', '*.xls', '*.XLSX', '*.XLS']
    excel_files = []
    for ext in extensions:
        # ساخت مسیر کامل فایل‌ها
        search_path = os.path.join(directory, ext)
        excel_files.extend(glob.glob(search_path))

    # حذف تکراری‌ها
    excel_files = list(set(excel_files))
    
    if not excel_files:
        print(f"⚠️ هیچ فایل اکسلی در پوشه '{directory}' یافت نشد.")
        return

    print(f"🔍 تعداد {len(excel_files)} فایل اکسل در پوشه '{directory}' پیدا شد. شروع پردازش...\n")
    
    grand_total_rows = 0

    for file_path in excel_files:
        file_name = os.path.basename(file_path)
        try:
            # خواندن تمام شیت‌ها
            xls_dict = pd.read_excel(file_path, sheet_name=None)
            
            file_total = 0
            for sheet_name, df in xls_dict.items():
                count = len(df)
                file_total += count

            grand_total_rows += file_total
            print(f"✅ {file_name} -> {file_total} ردیف")

        except Exception as e:
            print(f"❌ خطا در خواندن {file_name}: {e}")

    print("-" * 40)
    print(f"🚀 مجموع نهایی رکوردها: {grand_total_rows:,}")
    print("-" * 40)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="شمارش تعداد ردیف‌های فایل‌های اکسل")
    
    # 👇👇👇👇👇👇👇👇👇👇👇
    # آدرس فایل را در خط زیر وارد کرده‌ام (داخل student)
    parser.add_argument('--path', type=str, default='student', help='مسیر پوشه')
    # 👆👆👆👆👆👆👆👆👆👆👆
    
    args = parser.parse_args()
    
    # تمیزکاری مسیر ورودی
    target_path = args.path.strip('"').strip("'")
    
    count_rows_in_directory(target_path)
