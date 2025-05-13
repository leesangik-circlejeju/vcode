import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QListWidget, QMessageBox,
    QFileDialog, QProgressBar, QTreeWidget, QTreeWidgetItem, QComboBox, QMenu, QInputDialog
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QIcon, QFont
import git
from git import Repo, GitCommandError

class GitWorker(QThread):
    """Git 작업을 백그라운드에서 처리하는 워커 스레드"""
    progress = pyqtSignal(str)
    error = pyqtSignal(str)
    finished = pyqtSignal()
    files_updated = pyqtSignal(list)

    def __init__(self, repo_path, operation, github_url=None, *args, **kwargs):
        super().__init__()
        self.repo_path = repo_path
        self.operation = operation
        self.github_url = github_url
        self.args = args
        self.kwargs = kwargs

    def run(self):
        try:
            repo = Repo(self.repo_path)
            
            if self.operation == 'upload':
                repo.git.add('.')
                repo.index.commit('Update files')
                origin = repo.remote('origin')
                origin.push()
                self.progress.emit('내컴퓨터 → GIT 업로드 완료!')
                
            elif self.operation == 'download':
                origin = repo.remote('origin')
                origin.pull()
                self.progress.emit('GIT → 내컴퓨터 다운로드 완료!')
                
            elif self.operation == 'sync_local_to_git':
                repo.git.add('.')
                repo.index.commit('Sync local to git')
                origin = repo.remote('origin')
                origin.push()
                self.progress.emit('내컴퓨터 → GIT 동기화 완료! (내컴퓨터 내용으로 GIT 동일화)')

            elif self.operation == 'sync_git_to_local':
                origin = repo.remote('origin')
                origin.fetch()
                repo.git.reset('--hard', 'origin/main')
                self.progress.emit('GIT → 내컴퓨터 동기화 완료! (GIT 내용으로 내컴퓨터 동일화)')

            # 파일 목록 업데이트
            files = [item.a_path for item in repo.index.entries]
            self.files_updated.emit(files)

        except GitCommandError as e:
            self.error.emit(f'Git 오류: {str(e)}')
        except Exception as e:
            self.error.emit(f'오류 발생: {str(e)}')
        finally:
            self.finished.emit()

class GitGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.worker = None

    def initUI(self):
        self.setWindowTitle('Git GUI Control')
        self.setGeometry(100, 100, 900, 600)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # 경로 입력 (내컴퓨터 위치, Github 위치)
        path_form_layout = QHBoxLayout()
        self.local_path_input = QLineEdit('c:/homepage')
        self.github_url_input = QLineEdit('https://github.com/leesangik-circlejeju/vcode.git')
        path_form_layout.addWidget(QLabel('내컴퓨터 위치 :'))
        path_form_layout.addWidget(self.local_path_input)
        path_form_layout.addWidget(QLabel('Github 위치 :'))
        path_form_layout.addWidget(self.github_url_input)
        layout.addLayout(path_form_layout)

        # 상단 Git 동작 콤보박스 + 옵션입력 + 확인 + 안내
        git_action_names = [
            '내컴퓨터 → GIT 업로드',
            'GIT → 내컴퓨터 다운로드',
            '내컴퓨터 → GIT 동기화',
            'GIT → 내컴퓨터 동기화'
        ]
        self.top_git_action_combo = QComboBox()
        self.top_git_action_combo.addItems(git_action_names)
        self.top_git_action_combo.currentIndexChanged.connect(self.update_top_git_action_option)
        self.top_git_action_confirm = QPushButton('확인')
        self.top_git_action_confirm.setMaximumWidth(60)
        self.top_git_action_confirm.clicked.connect(self.handle_top_git_action)
        top_git_action_row = QHBoxLayout()
        top_git_action_row.addWidget(self.top_git_action_combo)
        top_git_action_row.addWidget(self.top_git_action_confirm)
        self.top_git_option_input = QLineEdit()
        self.top_git_option_input.setPlaceholderText('옵션 입력 (예: 명령어, 브랜치명 등)')
        self.top_git_option_label = QLabel('')
        self.top_git_option_label.setStyleSheet('color: #555; font-size: 11px; padding: 2px;')
        layout.addLayout(top_git_action_row)
        layout.addWidget(self.top_git_option_input)
        layout.addWidget(self.top_git_option_label)

        # 안내 메시지 박스 (QLabel)
        self.message_label = QLabel('')
        self.message_label.setStyleSheet('color: #1a237e; background: #f4f4f4; border: 1px solid #ccc; padding: 4px; font-size: 13px;')
        self.message_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        layout.addWidget(self.message_label)

        # 파일 목록 2열 배치 (QTreeWidget)
        file_lists_layout = QHBoxLayout()
        # 왼쪽: 저장소 파일 목록
        left_layout = QVBoxLayout()
        self.file_tree_label = QLabel('저장소 파일 목록:')
        self.file_tree = QTreeWidget()
        self.file_tree.setHeaderLabel('폴더/파일')
        self.file_tree.setSelectionMode(QTreeWidget.SingleSelection)
        self.file_tree.setUniformRowHeights(True)
        self.file_tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.file_tree.customContextMenuRequested.connect(lambda pos: self.show_tree_context_menu(self.file_tree, pos, is_git=False))
        left_layout.addWidget(self.file_tree_label)
        left_layout.addWidget(self.file_tree)
        # 내컴퓨터 관리용 명령어 UI (콤보박스+입력+확인+옵션라벨)
        self.local_action_combo = QComboBox()
        self.local_action_combo.addItems([
            '명령어 선택',
            '새 파일 만들기',
            '새 폴더 만들기',
            '이름 변경',
            '삭제',
            '탐색기에서 열기'
        ])
        self.local_action_combo.currentIndexChanged.connect(self.update_local_terminal_command)
        self.local_option_input = QLineEdit()
        self.local_option_input.setPlaceholderText('옵션 입력 (예: 파일명, 폴더명 등)')
        self.local_option_label = QLabel('')
        self.local_option_label.setStyleSheet('color: #555; font-size: 11px; padding: 2px;')
        self.local_action_button = QPushButton('확인')
        self.local_action_button.clicked.connect(self.handle_local_action)
        local_action_layout = QVBoxLayout()
        local_action_row = QHBoxLayout()
        local_action_row.addWidget(self.local_action_combo)
        local_action_row.addWidget(self.local_action_button)
        local_action_layout.addLayout(local_action_row)
        local_action_layout.addWidget(self.local_option_input)
        local_action_layout.addWidget(self.local_option_label)
        left_layout.addLayout(local_action_layout)
        file_lists_layout.addLayout(left_layout)
        # 오른쪽: Git 파일 목록
        right_layout = QVBoxLayout()
        self.git_file_tree_label = QLabel('Git 파일 목록:')
        self.git_file_tree = QTreeWidget()
        self.git_file_tree.setHeaderLabel('폴더/파일')
        self.git_file_tree.setSelectionMode(QTreeWidget.SingleSelection)
        self.git_file_tree.setUniformRowHeights(True)
        self.git_file_tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.git_file_tree.customContextMenuRequested.connect(lambda pos: self.show_tree_context_menu(self.git_file_tree, pos, is_git=True))
        right_layout.addWidget(self.git_file_tree_label)
        right_layout.addWidget(self.git_file_tree)
        # Git 관리용 명령어 UI (콤보박스+입력+확인+옵션라벨)
        self.git_action_combo = QComboBox()
        self.git_action_combo.addItems([
            '명령어 선택',
            'Git add (추적 시작)',
            'Git untrack (추적 해제)',
            'Git ignore에 추가',
            'Git 파일 삭제',
            'Git 폴더 삭제',
            'Git 이름 바꾸기',
            'Git 폴더 생성'
        ])
        self.git_action_combo.currentIndexChanged.connect(self.update_git_terminal_command)
        self.git_option_input = QLineEdit()
        self.git_option_input.setPlaceholderText('옵션 입력 (예: 파일명, 폴더명 등)')
        self.git_option_label = QLabel('')
        self.git_option_label.setStyleSheet('color: #555; font-size: 11px; padding: 2px;')
        self.git_action_button = QPushButton('확인')
        self.git_action_button.clicked.connect(self.handle_git_action)
        git_action_layout = QVBoxLayout()
        git_action_row = QHBoxLayout()
        git_action_row.addWidget(self.git_action_combo)
        git_action_row.addWidget(self.git_action_button)
        git_action_layout.addLayout(git_action_row)
        git_action_layout.addWidget(self.git_option_input)
        git_action_layout.addWidget(self.git_option_label)
        right_layout.addLayout(git_action_layout)
        file_lists_layout.addLayout(right_layout)
        layout.addLayout(file_lists_layout)

        # 경로 입력 변경 시 자동 갱신
        self.local_path_input.textChanged.connect(self.on_path_changed)
        self.github_url_input.textChanged.connect(self.on_path_changed)
        # 폴더 경로가 바뀔 때마다 상단 명령어 예시도 갱신
        self.local_path_input.textChanged.connect(self.update_top_git_action_option)

        # 프로그램 시작 시 바로 갱신
        self.on_path_changed()

        # 체크박스 신호 연결
        self._connect_tree_checkbox_signals()

    def on_path_changed(self):
        self.update_buttons()
        self.update_file_list()

    def update_buttons(self):
        path = self.local_path_input.text()
        is_valid = bool(path and os.path.exists(path))
        self.top_git_action_combo.setEnabled(is_valid)
        self.top_git_action_confirm.setEnabled(is_valid)
        self.top_git_option_input.setEnabled(is_valid)
        self.local_action_combo.setEnabled(is_valid)
        self.local_action_button.setEnabled(is_valid)
        self.git_action_combo.setEnabled(is_valid)
        self.git_action_button.setEnabled(is_valid)

    def update_top_git_action_option(self):
        idx = self.top_git_action_combo.currentIndex()
        folder = self.local_path_input.text().replace('\\', '/').rstrip('/')
        folder_name = folder.split('/')[-1] if folder else '폴더명'
        if idx == 0:  # 내컴퓨터 → GIT 업로드
            self.top_git_option_input.setText(f'cd "{folder}" && git add . && git commit -m "업로드" && git push')
            self.top_git_option_label.setText(
                f'설명: 내컴퓨터의 변경사항을 GIT 원격 저장소로 업로드합니다.\n'
                f'옵션: cd "{folder}" && git add/commit/push 명령어를 순서대로 입력하세요.\n'
                f'예시: cd "{folder}" && git add . && git commit -m "업로드" && git push\n'
                f'예시(특정 파일): cd "{folder}" && git add "test.txt" && git commit -m "업데이트" && git push\n'
                f'주의: 커밋 메시지는 자유롭게 입력, push 전 인증 필요할 수 있습니다.'
            )
        elif idx == 1:  # GIT → 내컴퓨터 다운로드
            self.top_git_option_input.setText(f'cd "{folder}" && git pull')
            self.top_git_option_label.setText(
                f'설명: GIT 원격 저장소의 내용을 내컴퓨터로 다운로드합니다.\n'
                f'옵션: cd "{folder}" && git pull 또는 git pull origin 브랜치명 형식으로 입력하세요.\n'
                f'예시: cd "{folder}" && git pull\n예시(브랜치): cd "{folder}" && git pull origin main\n'
                f'주의: 로컬 변경사항이 있으면 충돌이 발생할 수 있습니다.'
            )
        elif idx == 2:  # 내컴퓨터 → GIT 동기화
            self.top_git_option_input.setText(f'cd "{folder}" && git add . && git commit -m "동기화" && git push --force')
            self.top_git_option_label.setText(
                f'설명: 내컴퓨터 폴더의 내용으로 GIT 저장소를 동일하게 맞춥니다.\n'
                f'옵션: cd "{folder}" && git add/commit/push --force 명령어를 입력하세요.\n'
                f'예시: cd "{folder}" && git add . && git commit -m "동기화" && git push --force\n'
                f'주의: --force는 원격 저장소 내용을 덮어쓰므로 주의하세요.'
            )
        elif idx == 3:  # GIT → 내컴퓨터 동기화
            self.top_git_option_input.setText(f'cd "{folder}" && git fetch && git reset --hard origin/main')
            self.top_git_option_label.setText(
                f'설명: GIT 저장소의 내용으로 내컴퓨터 폴더를 동일하게 맞춥니다.\n'
                f'옵션: cd "{folder}" && git fetch && git reset --hard origin/브랜치명 형식으로 입력하세요.\n'
                f'예시: cd "{folder}" && git fetch && git reset --hard origin/main\n'
                f'주의: 내컴퓨터의 변경사항은 모두 사라집니다.'
            )
        else:
            self.top_git_option_input.setText('')
            self.top_git_option_label.setText('')

    def handle_top_git_action(self):
        idx = self.top_git_action_combo.currentIndex()
        option = self.top_git_option_input.text().strip()
        if idx < 0:
            self.message_label.setText('동작을 선택하세요.')
            return
        operation_map = {
            0: 'upload',
            1: 'download',
            2: 'sync_local_to_git',
            3: 'sync_git_to_local'
        }
        operation = operation_map.get(idx)
        self.run_git_operation(operation, extra_option=option)

    def run_git_operation(self, operation, extra_option=None):
        repo_path = self.local_path_input.text()
        github_url = self.github_url_input.text()
        if not repo_path:
            self.message_label.setText('저장소 경로를 입력해주세요.')
            return
        # 안내 메시지 초기화
        if operation == 'upload':
            self.message_label.setText('내컴퓨터 → GIT 업로드 중...')
        elif operation == 'download':
            self.message_label.setText('GIT → 내컴퓨터 다운로드 중...')
        elif operation == 'sync_local_to_git':
            self.message_label.setText('내컴퓨터 → GIT 동기화 중...')
        elif operation == 'sync_git_to_local':
            self.message_label.setText('GIT → 내컴퓨터 동기화 중...')
        # 워커 스레드 시작
        self.worker = GitWorker(repo_path, operation, github_url)
        self.worker.progress.connect(self.update_status)
        self.worker.error.connect(self.show_error)
        self.worker.finished.connect(self.operation_finished)
        self.worker.files_updated.connect(self.update_file_list)
        self.worker.start()

    def update_status(self, message):
        self.message_label.setText(message)

    def show_error(self, error_message):
        self.message_label.setText(f'<span style="color:red;">{error_message}</span>')

    def operation_finished(self):
        # 모든 버튼 다시 활성화
        self.update_buttons()

    def update_file_list(self, files=None):
        self.file_tree.clear()
        self.git_file_tree.clear()
        repo_path = self.local_path_input.text()
        # 저장소 파일 목록 (실제 폴더 내 모든 파일, 폴더 구조)
        if os.path.isdir(repo_path):
            for root, dirs, files in os.walk(repo_path):
                if '.git' in dirs:
                    dirs.remove('.git')
                rel_root = os.path.relpath(root, repo_path)
                parent = self.file_tree
                if rel_root != '.':
                    parent = self._get_or_create_tree_node(self.file_tree, rel_root)
                for file in files:
                    item = QTreeWidgetItem(parent, [file])
                    # 체크박스 추가
                    item.setCheckState(0, Qt.Unchecked)
        # Git 파일 목록 (tracked 파일, 폴더 구조)
        try:
            repo = Repo(repo_path)
            tracked_files = list(repo.git.ls_files().splitlines())
            for f in tracked_files:
                self._add_tree_path_with_checkbox(self.git_file_tree, f)
        except Exception:
            pass

    def _get_or_create_tree_node(self, tree_widget, rel_path):
        # rel_path: 'a/b/c' -> 트리에서 해당 폴더 노드 반환, 없으면 생성
        parts = rel_path.split(os.sep)
        parent = tree_widget
        for part in parts:
            found = None
            for i in range(parent.topLevelItemCount() if parent is tree_widget else parent.childCount()):
                item = parent.topLevelItem(i) if parent is tree_widget else parent.child(i)
                if item.text(0) == part:
                    found = item
                    break
            if found:
                parent = found
            else:
                new_item = QTreeWidgetItem([part])
                if parent is tree_widget:
                    parent.addTopLevelItem(new_item)
                else:
                    parent.addChild(new_item)
                parent = new_item
        return parent

    def _add_tree_path_with_checkbox(self, tree_widget, file_path):
        # file_path: 'a/b/c.txt' -> 트리 구조로 추가 + 체크박스
        parts = file_path.split('/')
        parent = tree_widget
        for part in parts[:-1]:
            found = None
            for i in range(parent.topLevelItemCount() if parent is tree_widget else parent.childCount()):
                item = parent.topLevelItem(i) if parent is tree_widget else parent.child(i)
                if item.text(0) == part:
                    found = item
                    break
            if found:
                parent = found
            else:
                new_item = QTreeWidgetItem([part])
                if parent is tree_widget:
                    parent.addTopLevelItem(new_item)
                else:
                    parent.addChild(new_item)
                parent = new_item
        item = QTreeWidgetItem(parent, [parts[-1]])
        item.setCheckState(0, Qt.Unchecked)

    # 체크박스 연동: 체크 상태 변경 시 옵션 입력란 자동 채움
    def _connect_tree_checkbox_signals(self):
        self.file_tree.itemChanged.connect(self._update_local_option_from_checkbox)
        self.git_file_tree.itemChanged.connect(self._update_git_option_from_checkbox)
    def _update_local_option_from_checkbox(self, item, col):
        checked = self._get_checked_items(self.file_tree)
        if checked:
            self.local_option_input.setText(' '.join(checked))
    def _update_git_option_from_checkbox(self, item, col):
        checked = self._get_checked_items(self.git_file_tree)
        if checked:
            self.git_option_input.setText(' '.join(checked))
    def _get_checked_items(self, tree_widget):
        result = []
        def _recurse(parent, path=''):
            for i in range(parent.topLevelItemCount() if parent is tree_widget else parent.childCount()):
                item = parent.topLevelItem(i) if parent is tree_widget else parent.child(i)
                name = item.text(0)
                full = f'{path}/{name}' if path else name
                if item.checkState(0) == Qt.Checked and item.childCount() == 0:
                    result.append(full)
                _recurse(item, full)
        _recurse(tree_widget)
        return result

    def update_local_terminal_command(self):
        action = self.local_action_combo.currentText()
        item = self.file_tree.currentItem()
        repo_path = self.local_path_input.text()
        path = ''
        rel_path = ''
        if item:
            path = self._get_full_path_from_tree(item, repo_path)
            rel_path = os.path.relpath(path, repo_path)
        if action == '새 파일 만들기':
            self.local_option_input.setText(f'type nul > "{rel_path or "새파일.txt"}"')
            self.local_option_label.setText(
                '설명: 지정한 경로에 새 파일을 생성합니다.\n'
                '옵션: type nul > "파일명" 형식으로 입력하세요.\n'
                '예시: type nul > "test.txt"\n예시(여러 개): type nul > "a.txt" & type nul > "b.txt"\n'
                '주의: 이미 존재하는 파일은 덮어쓰지 않고, 빈 파일만 생성합니다.'
            )
        elif action == '새 폴더 만들기':
            self.local_option_input.setText(f'mkdir "{rel_path or "새폴더"}"')
            self.local_option_label.setText(
                '설명: 지정한 경로에 새 폴더를 생성합니다.\n'
                '옵션: mkdir "폴더명" 형식으로 입력하세요.\n'
                '예시: mkdir "새폴더"\n예시(여러 개): mkdir "a" "b" "c"\n'
                '주의: 이미 존재하는 폴더는 무시됩니다.'
            )
        elif action == '이름 변경':
            self.local_option_input.setText(f'rename "{rel_path}" "새이름"')
            self.local_option_label.setText(
                '설명: 선택한 파일/폴더의 이름을 변경합니다.\n'
                '옵션: rename "기존이름" "새이름" 형식으로 입력하세요.\n'
                '예시: rename "old.txt" "new.txt"\n예시(폴더): rename "myfolder" "backup"\n'
                '주의: 같은 폴더 내에서만 이름 변경이 가능합니다.'
            )
        elif action == '삭제':
            self.local_option_input.setText(f'del "{rel_path}"  또는  rmdir /s /q "{rel_path}"')
            self.local_option_label.setText(
                '설명: 선택한 파일 또는 폴더를 삭제합니다.\n'
                '옵션: 파일은 del, 폴더는 rmdir 명령을 사용합니다.\n'
                '예시(파일): del "test.txt"\n예시(폴더): rmdir /s /q "myfolder"\n'
                '주의: 삭제된 파일/폴더는 복구할 수 없습니다.'
            )
        elif action == '탐색기에서 열기':
            self.local_option_input.setText(f'explorer /select,"{path}"')
            self.local_option_label.setText(
                '설명: 선택한 파일/폴더를 윈도우 탐색기에서 표시합니다.\n'
                '옵션: explorer /select,"경로" 형식으로 입력하세요.\n'
                '예시: explorer /select,"C:/Users/me/test.txt"\n예시(폴더): explorer "C:/Users/me/Documents"'
            )
        else:
            self.local_option_input.setText('')
            self.local_option_label.setText('')

    def update_git_terminal_command(self):
        action = self.git_action_combo.currentText()
        item = self.git_file_tree.currentItem()
        repo_path = self.local_path_input.text()
        path = ''
        rel_path = ''
        if item:
            path = self._get_full_path_from_tree(item, repo_path, git_tree=True)
            rel_path = os.path.relpath(path, repo_path)
        if action == 'Git add (추적 시작)':
            self.git_option_input.setText(f'git add "{rel_path}"')
            self.git_option_label.setText(
                '설명: 선택한 파일/폴더를 Git 추적 대상으로 추가합니다.\n'
                '옵션: git add "파일명" 또는 git add . 형식으로 입력하세요.\n'
                '예시: git add "test.txt"\n예시(여러 개): git add "a.txt" "b.txt"\n예시(전체): git add .\n예시(확장자): git add *.md\n'
                '주의: add 후 커밋해야 반영됩니다.'
            )
        elif action == 'Git untrack (추적 해제)':
            self.git_option_input.setText(f'git rm --cached "{rel_path}"')
            self.git_option_label.setText(
                '설명: 선택한 파일/폴더를 Git 추적에서 해제합니다(로컬 파일은 남음).\n'
                '옵션: git rm --cached "파일명" 형식으로 입력하세요.\n'
                '예시: git rm --cached "test.txt"\n예시(폴더): git rm --cached -r "data"\n'
                '주의: --cached 옵션을 빼면 실제 파일도 삭제됩니다.'
            )
        elif action == 'Git ignore에 추가':
            self.git_option_input.setText(f'echo {rel_path} >> .gitignore')
            self.git_option_label.setText(
                '설명: 선택한 파일/폴더를 .gitignore에 추가하여 Git 추적에서 제외합니다.\n'
                '옵션: echo [패턴] >> .gitignore 형식으로 입력하세요.\n'
                '예시(파일): echo test.txt >> .gitignore\n예시(확장자): echo *.log >> .gitignore\n예시(폴더): echo temp/ >> .gitignore\n예시(주석): echo # 임시파일 >> .gitignore\n'
                '주의: 이미 추적 중인 파일은 untrack해야 완전히 제외됩니다.'
            )
        elif action == 'Git 파일 삭제':
            self.git_option_input.setText(f'git rm "{rel_path}"')
            self.git_option_label.setText(
                '설명: 선택한 파일을 Git에서 삭제하고 커밋합니다.\n'
                '옵션: git rm "파일명" 형식으로 입력하세요.\n'
                '예시: git rm "test.txt"\n예시(여러 개): git rm "a.txt" "b.txt"\n'
                '주의: 실제 파일도 삭제되며, 커밋 후 push해야 원격에서도 삭제됩니다.'
            )
        elif action == 'Git 폴더 삭제':
            self.git_option_input.setText(f'git rm -r "{rel_path}"')
            self.git_option_label.setText(
                '설명: 선택한 폴더를 Git에서 삭제하고 커밋합니다.\n'
                '옵션: git rm -r "폴더명" 형식으로 입력하세요.\n'
                '예시: git rm -r "myfolder"\n'
                '주의: 실제 폴더/내용도 삭제되며, 커밋 후 push해야 원격에서도 삭제됩니다.'
            )
        elif action == 'Git 이름 바꾸기':
            self.git_option_input.setText(f'git mv "{rel_path}" "새이름"')
            self.git_option_label.setText(
                '설명: 선택한 파일/폴더의 이름을 변경(Git 추적 포함)합니다.\n'
                '옵션: git mv "기존이름" "새이름" 형식으로 입력하세요.\n'
                '예시(파일): git mv "old.txt" "new.txt"\n예시(폴더): git mv "myfolder" "backup"\n'
                '주의: mv 후 커밋해야 반영됩니다.'
            )
        elif action == 'Git 폴더 생성':
            self.git_option_input.setText(f'mkdir "{rel_path or "새폴더"}" && type nul > "{rel_path or "새폴더"}/.gitkeep" && git add "{rel_path or "새폴더"}/.gitkeep"')
            self.git_option_label.setText(
                '설명: 새 폴더를 만들고 .gitkeep 파일로 Git에 추적시킵니다.\n'
                '옵션: mkdir "폴더명" && type nul > "폴더명/.gitkeep" && git add "폴더명/.gitkeep" 형식으로 입력하세요.\n'
                '예시: mkdir "새폴더" && type nul > "새폴더/.gitkeep" && git add "새폴더/.gitkeep"\n'
                '주의: Git은 빈 폴더만은 추적하지 않으므로 더미파일(.gitkeep 등)이 필요합니다.'
            )
        else:
            self.git_option_input.setText('')
            self.git_option_label.setText('')

    def handle_local_action(self):
        self.update_local_terminal_command()
        idx = self.local_action_combo.currentIndex()
        if idx < 0:
            QMessageBox.information(self, '알림', '명령어를 선택하세요.')
            return
        cmd = self.local_option_input.text().strip()
        if not cmd:
            return
        try:
            import os
            # 현재 작업 디렉토리 저장
            current_dir = os.getcwd()
            # 명령어를 실행할 디렉토리로 이동
            repo_path = self.local_path_input.text()
            os.chdir(repo_path)
            # 명령 프롬프트 실행
            os.system(f'start cmd.exe /k "{cmd}"')
            # 원래 디렉토리로 복귀
            os.chdir(current_dir)
            self.message_label.setText('명령어가 터미널에서 실행됩니다.')
        except Exception as e:
            error_msg = f'터미널 실행 실패: {str(e)}'
            self.message_label.setText(f'<span style="color:red;">{error_msg}</span>')
            QMessageBox.critical(self, '오류', error_msg)

    def handle_git_action(self):
        self.update_git_terminal_command()
        idx = self.git_action_combo.currentIndex()
        if idx < 0:
            QMessageBox.information(self, '알림', '명령어를 선택하세요.')
            return
        cmd = self.git_option_input.text().strip()
        if not cmd:
            return
        try:
            import os
            # 현재 작업 디렉토리 저장
            current_dir = os.getcwd()
            # 명령어를 실행할 디렉토리로 이동
            repo_path = self.local_path_input.text()
            os.chdir(repo_path)
            # 명령 프롬프트 실행
            os.system(f'start cmd.exe /k "{cmd}"')
            # 원래 디렉토리로 복귀
            os.chdir(current_dir)
            self.message_label.setText('명령어가 터미널에서 실행됩니다.')
        except Exception as e:
            error_msg = f'터미널 실행 실패: {str(e)}'
            self.message_label.setText(f'<span style="color:red;">{error_msg}</span>')
            QMessageBox.critical(self, '오류', error_msg)

    def _get_full_path_from_tree(self, item, root_path, git_tree=False):
        # 트리에서 선택한 아이템의 전체 경로를 반환
        parts = []
        while item is not None:
            parts.insert(0, item.text(0))
            item = item.parent()
        return os.path.join(root_path, *parts)

    def disable_buttons_and_lists(self):
        # 더 이상 사용하지 않음 (on_path_changed로 대체)
        pass
    def confirm_path(self):
        # 더 이상 사용하지 않음 (on_path_changed로 대체)
        pass

    def show_tree_context_menu(self, tree_widget, pos, is_git=False):
        item = tree_widget.itemAt(pos)
        if not item:
            return
        # 경로 계산
        repo_path = self.local_path_input.text()
        path = self._get_full_path_from_tree(item, repo_path, git_tree=is_git)
        rel_path = os.path.relpath(path, repo_path)
        # 체크된 항목들
        checked = self._get_checked_items(tree_widget)
        # 파일/폴더 판별
        is_folder = False
        if os.path.isdir(path):
            is_folder = True
        elif not os.path.exists(path):
            # git 트리의 경우 실제 파일이 없을 수 있음(폴더 추정)
            if item.childCount() > 0:
                is_folder = True
        menu = QMenu()
        if is_folder:
            act_new = menu.addAction('폴더 생성')
            act_rename = menu.addAction('폴더 이름 바꾸기')
            act_delete = menu.addAction('폴더 삭제')
            menu.addSeparator()
            if not is_git:
                act_upload = menu.addAction('폴더 업로드')
            else:
                act_download = menu.addAction('폴더 다운로드')
        else:
            act_copy = menu.addAction('파일 복사')
            act_rename = menu.addAction('파일 이름 바꾸기')
            act_delete = menu.addAction('파일 삭제')
            menu.addSeparator()
            if not is_git:
                act_upload = menu.addAction('파일 업로드')
            else:
                act_download = menu.addAction('파일 다운로드')
        action = menu.exec_(tree_widget.viewport().mapToGlobal(pos))

        # 체크된 항목이 있으면, 체크된 것들로 명령어 생성
        targets = checked if checked else [rel_path]
        targets_str = ' '.join(f'"{t}"' for t in targets)
        # 아래 옵션 입력란에 명령어 예시 자동 입력
        if is_git:
            # 오른쪽(Git) 트리
            if is_folder:
                if action == act_new:
                    self.git_option_input.setText(f'cd "{repo_path}" && mkdir "{rel_path}/새폴더" && type nul > "{rel_path}/새폴더/.gitkeep" && git add "{rel_path}/새폴더/.gitkeep"')
                elif action == act_rename:
                    self.git_option_input.setText(f'cd "{repo_path}" && git mv {targets_str} "새이름"')
                elif action == act_delete:
                    self.git_option_input.setText(f'cd "{repo_path}" && git rm -r {targets_str}')
                elif action == act_download:
                    self.git_option_input.setText(f'cd "{repo_path}" && git checkout origin/main -- {targets_str}')
            else:
                if action == act_copy:
                    self.git_option_input.setText(f'cd "{repo_path}" && copy {targets_str} "복사본.txt" && git add "복사본.txt"')
                elif action == act_rename:
                    self.git_option_input.setText(f'cd "{repo_path}" && git mv {targets_str} "새이름"')
                elif action == act_delete:
                    self.git_option_input.setText(f'cd "{repo_path}" && git rm {targets_str}')
                elif action == act_download:
                    self.git_option_input.setText(f'cd "{repo_path}" && git checkout origin/main -- {targets_str}')
        else:
            # 왼쪽(로컬) 트리
            if is_folder:
                if action == act_new:
                    self.local_option_input.setText(f'cd "{repo_path}" && mkdir "{rel_path}/새폴더"')
                elif action == act_rename:
                    self.local_option_input.setText(f'cd "{repo_path}" && rename {targets_str} "새이름"')
                elif action == act_delete:
                    self.local_option_input.setText(f'cd "{repo_path}" && rmdir /s /q {targets_str}')
                elif action == act_upload:
                    self.local_option_input.setText(f'cd "{repo_path}" && git add {targets_str} && git commit -m "폴더 업로드" && git push')
            else:
                if action == act_copy:
                    self.local_option_input.setText(f'cd "{repo_path}" && copy {targets_str} "복사본.txt"')
                elif action == act_rename:
                    self.local_option_input.setText(f'cd "{repo_path}" && rename {targets_str} "새이름"')
                elif action == act_delete:
                    self.local_option_input.setText(f'cd "{repo_path}" && del {targets_str}')
                elif action == act_upload:
                    self.local_option_input.setText(f'cd "{repo_path}" && git add {targets_str} && git commit -m "파일 업로드" && git push')

    def update_file_list(self):
        self.file_tree.clear()
        self.git_file_tree.clear()
        repo_path = self.local_path_input.text()
        # 저장소 파일 목록 (실제 폴더 내 모든 파일, 폴더 구조)
        if os.path.isdir(repo_path):
            for root, dirs, files in os.walk(repo_path):
                if '.git' in dirs:
                    dirs.remove('.git')
                rel_root = os.path.relpath(root, repo_path)
                parent = self.file_tree
                if rel_root != '.':
                    parent = self._get_or_create_tree_node(self.file_tree, rel_root)
                for file in files:
                    item = QTreeWidgetItem(parent, [file])
                    # 체크박스 추가
                    item.setCheckState(0, Qt.Unchecked)
        # Git 파일 목록 (tracked 파일, 폴더 구조)
        try:
            repo = Repo(repo_path)
            tracked_files = list(repo.git.ls_files().splitlines())
            for f in tracked_files:
                self._add_tree_path_with_checkbox(self.git_file_tree, f)
        except Exception:
            pass

def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')  # 모던한 스타일 적용
    
    # 폰트 설정
    font = QFont('Segoe UI', 9)
    app.setFont(font)
    
    window = GitGUI()
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main() 