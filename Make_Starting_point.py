import os
import shutil
import subprocess
import sys
import time
from datetime import datetime, timedelta

# Python 실행 경로 확인
PYTHON_EXECUTABLE = r"c:\Users\김남빈\Desktop\Coding\Itemlist_Update\myenv\Scripts\python.exe"
print(sys.executable)  # 현재 실행 중인 Python의 경로
print(sys.path)        # 패키지 경로 확인

# 경로 설정
source_path = r"C:\Users\김남빈\Desktop\Coding\Itemlist_Update"
destination_path = r"C:\Users\김남빈\OneDrive\★Hangaweemarket\Online\☆Item_List"
log_file_path = os.path.join(source_path, "task_scheduler_log.txt")

# 날짜 계산
yesterday = (datetime.now() - timedelta(days=1)).strftime("%y%m%d")
today = datetime.now().strftime("%y%m%d")

# 이동/복사/삭제 파일 목록
files_to_move = [f"Scraped_{yesterday}.xlsx", f"ScrapedM_{yesterday}.xlsx"]
files_to_copy = [f"Itemlist_{yesterday}.xlsx", f"Purchase_{yesterday}.xlsx"]
files_to_delete = [f"Active_Items_{yesterday}.csv", f"Total_Items_{yesterday}.csv"]

# 실행할 스크립트 목록 (순서 중요)
scripts_to_run = [
    "token_extractor.py",
    "scraping_Purchase_v2.py",
    "scraping_M_final.py",
    "scraping_L_final.py",
    "Update_Excel_Final_v9.py"
]

# 작업 디렉토리 이동
os.chdir(source_path)

# 로그 작성 함수
def log_message(message):
    with open(log_file_path, "a", encoding="utf-8") as log:
        log.write(f"{datetime.now()} - {message}\n")
    print(message)

# 파일 삭제 함수
def delete_files(file_list, path):
    for file_name in file_list:
        file_path = os.path.join(path, file_name)
        if os.path.exists(file_path):
            os.remove(file_path)
            log_message(f"✅ Deleted: {file_path}")
        else:
            log_message(f"⚠️ File not found for deletion: {file_path}")

# 파일 이동 함수
def move_files(file_list, src_path, dest_path):
    for file_name in file_list:
        src_file = os.path.join(src_path, file_name)
        dest_file = os.path.join(dest_path, file_name)
        if os.path.exists(src_file):
            shutil.move(src_file, dest_file)
            log_message(f"✅ Moved: {src_file} → {dest_file}")
        else:
            log_message(f"⚠️ File not found for moving: {src_file}")

# 파일 복사 및 이름 변경 함수
def copy_and_rename_files(file_list, src_path, dest_path, old_date, new_date):
    for file_name in file_list:
        src_file = os.path.join(src_path, file_name)
        dest_file = os.path.join(dest_path, file_name)
        new_file_name = file_name.replace(old_date, new_date)
        renamed_file = os.path.join(src_path, new_file_name)

        if os.path.exists(src_file):
            shutil.copy2(src_file, dest_file)
            log_message(f"✅ Copied: {src_file} → {dest_file}")

            time.sleep(1)  # 파일 시스템 지연 방지
            os.rename(src_file, renamed_file)
            log_message(f"✅ Renamed: {src_file} → {renamed_file}")
        else:
            log_message(f"⚠️ File not found for copying/renaming: {src_file}")

# 스크립트 실행 함수 (모든 스크립트를 순서대로 실행)
def run_scripts(script_list, src_path):
    today_itemlist = os.path.join(src_path, f"Itemlist_{today}.xlsx")
    today_purchase = os.path.join(src_path, f"Purchase_{today}.xlsx")

    # 필수 파일이 존재해야 실행
    if os.path.exists(today_itemlist) and os.path.exists(today_purchase):
        for script in script_list:
            script_path = os.path.join(src_path, script)

            if not os.path.exists(script_path):
                log_message(f"❌ Script not found: {script_path}")
                break

            log_message(f"🚀 Executing script: {script_path}")
            try:
                if not os.path.exists(PYTHON_EXECUTABLE):
                    log_message(f"❌ Python executable not found: {PYTHON_EXECUTABLE}")
                    break

                subprocess.run([PYTHON_EXECUTABLE, script_path], cwd=src_path, check=True)
                log_message(f"✅ Successfully executed: {script_path}")

            except subprocess.CalledProcessError as e:
                log_message(f"❌ Error executing script {script_path}: {e}")
                break
    else:
        log_message("[INFO] Skipping script execution as required files are missing.")

# 실행 흐름
delete_files(files_to_delete, source_path)
move_files(files_to_move, source_path, destination_path)
copy_and_rename_files(files_to_copy, source_path, destination_path, yesterday, today)
run_scripts(scripts_to_run, source_path)
