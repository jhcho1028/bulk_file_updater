import os
import shutil
import git
import openpyxl

# 엑셀 파일 경로 및 이슈 템플릿 경로 설정
excel_file = 'repolist.xlsx'
workflows_file = r'D:\Sources\GitHub_TEMP\PR_WIND3\.github\workflows\assign-reviewers.yml'
org_name = 'MY_ORG_NAME'

# 특정 시트에서 레포지토리 목록 가져오기
sheet_name = 'Sheet1'  # 시트 이름 지정
wb = openpyxl.load_workbook(excel_file)
ws = wb[sheet_name]

# 레포지토리 목록 로드 (C열, 1행 헤더 제외)
repos = [row[0].strip() for row in ws.iter_rows(min_row=2, min_col=3, max_col=3, values_only=True) if row[0] is not None]

# 작업 디렉토리 설정 (예: 'D:\Repositories\')
base_dir = r'D:\Sources\GitHub_TEMP'

for repo_name in repos:
    repo_dir = os.path.join(base_dir, repo_name)

    print(f'Processing repository: {repo_name}')

    # 레포지토리 클론 또는 업데이트
    if not os.path.exists(repo_dir):
        repo_url = f'https://github.com/{org_name}/{repo_name}.git'
        repo = git.Repo.clone_from(repo_url, repo_dir)
        print("Clone 완료!")
    else:
        repo = git.Repo(repo_dir)
        repo.remotes.origin.pull()
        print("Pull 완료!")

    # 특정 파일 복사 및 덮어쓰기
    destination_path = os.path.join(repo_dir, '.github', 'workflows')
    os.makedirs(destination_path, exist_ok=True)  # 폴더 생성

    dest_file = os.path.join(destination_path, os.path.basename(workflows_file))
    shutil.copy(workflows_file, dest_file)
    print(f'File copied to {dest_file}')

    # 변경 사항 확인 후 커밋 & 푸시
    repo.git.add(dest_file)  # 특정 파일만 스테이징
    if repo.is_dirty(path=dest_file):  # 변경 사항이 있는 경우만 커밋 & 푸시
        repo.index.commit('Update assign-reviewers.yml')
        repo.remotes.origin.push()
        print("Push 완료!")
    else:
        print(f'No changes detected in {repo_name}, skipping commit.')

print("모든 작업이 완료되었습니다!")
