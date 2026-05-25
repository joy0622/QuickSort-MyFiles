import os
import shutil
from pathlib import Path

# 확장자별 폴더 매핑 설정
FILE_TYPES = {
    'Documents': ['.pdf', '.docx', '.txt', '.xlsx', '.pptx'],
    'Images': ['.jpg', '.jpeg', '.png', '.gif', '.svg'],
    'Videos': ['.mp4', '.mov', '.avi'],
    'Archives': ['.zip', '.tar', '.rar', '.7z'],
}

def organize_folder(target_path):
    target_dir = Path(target_path)
    
    if not target_dir.exists():
        print(f"Error: {target_path} 경로를 찾을 수 없습니다.")
        return

    for file in target_dir.iterdir():
        if file.is_file():
            extension = file.suffix.lower()
            dest_folder = 'Others' # 분류되지 않은 파일용
            
            for folder, exts in FILE_TYPES.items():
                if extension in exts:
                    dest_folder = folder
                    break
            
            # 목적지 폴더 생성
            (target_dir / dest_folder).mkdir(exist_ok=True)
            
            # 파일 이동
            try:
                shutil.move(str(file), str(target_dir / dest_folder / file.name))
                print(f"Moved: {file.name} -> {dest_folder}")
            except Exception as e:
                print(f"Failed to move {file.name}: {e}")

if __name__ == "__main__":
    # 정리하고 싶은 폴더 경로를 입력하세요
    path_to_clean = input("정리할 폴더의 전체 경로를 입력하세요: ")
    organize_folder(path_to_clean)
