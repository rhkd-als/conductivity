import os
import msvcrt

def print_progress(current, total, bar_length=30):
    percent = current / total if total else 0
    filled = int(bar_length * percent)
    
    # \033[32m: 초록색 시작, \033[0m: 색상 초기화
    green_bar = "\033[32m" + "█" * filled + "\033[0m"
    remaining_bar = "-" * (bar_length - filled)
    
    print(f"\rProgress: |{green_bar}{remaining_bar}| {percent*100:5.1f}% ({current}/{total})", end="")

def delete_screenshot_files(folder_path):
    folder_path = folder_path.strip().strip('"')
    
    if not os.path.exists(folder_path):
        print(f"❌ 경로를 찾을 수 없습니다: {folder_path}")
        return

    # 먼저 모든 파일 리스트를 수집 (전체 개수 파악용)
    all_files = []
    for root, _, files in os.walk(folder_path):
        for file in files:
            all_files.append(os.path.join(root, file))

    total = len(all_files)
    if total == 0:
        print("⚠️ 스캔할 파일이 없습니다.")
        return

    deleted_count = 0
    skipped_count = 0
    failed_count = 0

    print(f"🔍 스캔 및 삭제 작업 시작...")

    for i, file_path in enumerate(all_files, start=1):
        file_name = os.path.basename(file_path)
        
        if "Screenshot" in file_name:
            try:
                os.remove(file_path)
                deleted_count += 1
            except Exception:
                failed_count += 1
        else:
            skipped_count += 1
        
        # 매 파일마다 프로그레스 바 업데이트
        print_progress(i, total)

    # 결과 출력
    print("\n\n" + "="*40)
    print("✅ 작업 완료!")
    print(f"🗑️  Deleted files: {deleted_count}")
    print(f"⏭️  Skipped files: {skipped_count}")
    if failed_count > 0:
        print(f"❌ Failed to delete: {failed_count}")
    print("="*40)

if __name__ == "__main__":
    while True:
        folder = input("📂 삭제 작업을 진행할 폴더 경로를 입력하세요: ").strip()
        
        delete_screenshot_files(folder)

        print("\n계속하려면 Ctrl+F를 누르시고, 종료하려면 아무 키나 누르세요.")
        if msvcrt.getch() != b'\x06':
            break