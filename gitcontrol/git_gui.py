import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QListWidget, QMessageBox,
    QFileDialog, QProgressBar, QTreeWidget, QTreeWidgetItem, QComboBox
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

        # 버튼 그룹 (4개)
        button_layout = QHBoxLayout()
        self.upload_button = QPushButton('내컴퓨터 → GIT 업로드')
        self.upload_button.clicked.connect(lambda: self.run_git_operation('upload'))
        self.download_button = QPushButton('GIT → 내컴퓨터 다운로드')
        self.download_button.clicked.connect(lambda: self.run_git_operation('download'))
        self.sync_local_to_git_button = QPushButton('내컴퓨터 → GIT 동기화')
        self.sync_local_to_git_button.clicked.connect(lambda: self.run_git_operation('sync_local_to_git'))
        self.sync_git_to_local_button = QPushButton('GIT → 내컴퓨터 동기화')
        self.sync_git_to_local_button.clicked.connect(lambda: self.run_git_operation('sync_git_to_local'))
        for button in [self.upload_button, self.download_button, self.sync_local_to_git_button, self.sync_git_to_local_button]:
            button_layout.addWidget(button)
        layout.addLayout(button_layout)

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
        left_layout.addWidget(self.file_tree_label)
        left_layout.addWidget(self.file_tree)
        # 내컴퓨터 관리용 명령어 UI
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
        right_layout.addWidget(self.git_file_tree_label)
        right_layout.addWidget(self.git_file_tree)
        # Git 관리용 명령어 UI
        self.git_action_combo = QComboBox()
        self.git_action_combo.addItems([
            '명령어 선택',
            'Git add (추적 시작)',
            'Git untrack (추적 해제)',
            'Git ignore에 추가'
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

        # 프로그램 시작 시 바로 갱신
        self.on_path_changed()

    def on_path_changed(self):
        self.update_buttons()
        self.update_file_list()

    def update_buttons(self):
        path = self.local_path_input.text()
        is_valid = bool(path and os.path.exists(path))
        self.upload_button.setEnabled(is_valid)
        self.download_button.setEnabled(is_valid)
        self.sync_local_to_git_button.setEnabled(is_valid)
        self.sync_git_to_local_button.setEnabled(is_valid)

    def run_git_operation(self, operation):
        repo_path = self.local_path_input.text()
        github_url = self.github_url_input.text()
        if not repo_path:
            self.message_label.setText('저장소 경로를 입력해주세요.')
            return
        # 모든 버튼 비활성화
        for button in [self.upload_button, self.download_button, self.sync_local_to_git_button, self.sync_git_to_local_button]:
            button.setEnabled(False)
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
                    QTreeWidgetItem(parent, [file])
        # Git 파일 목록 (tracked 파일, 폴더 구조)
        try:
            repo = Repo(repo_path)
            tracked_files = list(repo.git.ls_files().splitlines())
            for f in tracked_files:
                self._add_tree_path(self.git_file_tree, f)
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

    def _add_tree_path(self, tree_widget, file_path):
        # file_path: 'a/b/c.txt' -> 트리 구조로 추가
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
        QTreeWidgetItem(parent, [parts[-1]])

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
            self.local_option_input.setText(rel_path or '새파일.txt')
            self.local_option_label.setText(
                '설명: 지정한 경로에 새 파일을 생성합니다.\n'
                '옵션: 파일명(확장자 포함)을 입력하세요.\n'
                '예시: test.txt, log.txt\n'
                '주의: 이미 존재하는 파일은 덮어쓰지 않고, 빈 파일만 생성합니다.'
            )
        elif action == '새 폴더 만들기':
            self.local_option_input.setText(rel_path or '새폴더')
            self.local_option_label.setText(
                '설명: 지정한 경로에 새 폴더를 생성합니다.\n'
                '옵션: 폴더명 입력. 여러 폴더는 쉼표로 구분.\n'
                '예시: 새폴더, data/2024/06'
            )
        elif action == '이름 변경':
            self.local_option_input.setText('새이름')
            self.local_option_label.setText(
                '설명: 선택한 파일/폴더의 이름을 변경합니다.\n'
                '옵션: 새 이름 입력.\n'
                '예시: new.txt, backup\n'
                '주의: 같은 폴더 내에서만 이름 변경이 가능합니다.'
            )
        elif action == '삭제':
            self.local_option_input.setText(rel_path)
            self.local_option_label.setText(
                '설명: 선택한 파일 또는 폴더를 삭제합니다.\n'
                '옵션: 파일명 또는 폴더명 입력.\n'
                '예시: test.txt, myfolder\n'
                '주의: 삭제된 파일/폴더는 복구할 수 없습니다.'
            )
        elif action == '탐색기에서 열기':
            self.local_option_input.setText(path)
            self.local_option_label.setText(
                '설명: 선택한 파일/폴더를 윈도우 탐색기에서 표시합니다.\n'
                '옵션: 경로 입력.\n'
                '예시: C:/Users/me/test.txt, C:/Users/me/Documents'
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
            self.git_option_input.setText(rel_path)
            self.git_option_label.setText(
                '설명: 선택한 파일/폴더를 Git 추적 대상으로 추가합니다.\n'
                '옵션: 파일명/폴더명 입력. 여러 개는 쉼표로 구분.\n'
                '예시: test.txt, data/\n'
            )
        elif action == 'Git untrack (추적 해제)':
            self.git_option_input.setText(rel_path)
            self.git_option_label.setText(
                '설명: 선택한 파일/폴더를 Git 추적에서 해제합니다(로컬 파일은 남음).\n'
                '옵션: 파일명/폴더명 입력.\n'
                '예시: test.txt, data/\n'
                '주의: --cached 옵션을 빼면 실제 파일도 삭제됩니다.'
            )
        elif action == 'Git ignore에 추가':
            self.git_option_input.setText(rel_path)
            self.git_option_label.setText(
                '설명: 선택한 파일/폴더를 .gitignore에 추가하여 Git 추적에서 제외합니다.\n'
                '옵션: 파일명/폴더명/패턴 입력.\n'
                '예시: *.log, temp/\n'
                '주의: 이미 추적 중인 파일은 untrack해야 완전히 제외됩니다.'
            )
        else:
            self.git_option_input.setText('')
            self.git_option_label.setText('')

    def handle_local_action(self):
        self.update_local_terminal_command()
        action = self.local_action_combo.currentText()
        option = self.local_option_input.text().strip()
        item = self.file_tree.currentItem()
        repo_path = self.local_path_input.text()
        if not action or action == '명령어 선택':
            QMessageBox.information(self, '알림', '명령어를 선택하세요.')
            return
        if not option:
            QMessageBox.warning(self, '경고', '옵션(파일명/폴더명 등)을 입력하세요.')
            return
        path = self._get_full_path_from_tree(item, repo_path) if item else os.path.join(repo_path, option)
        if action == '새 파일 만들기':
            try:
                with open(path, 'w', encoding='utf-8') as f:
                    pass
                self.update_file_list()
            except Exception as e:
                QMessageBox.critical(self, '오류', f'파일 생성 실패: {e}')
        elif action == '새 폴더 만들기':
            try:
                os.makedirs(path, exist_ok=True)
                self.update_file_list()
            except Exception as e:
                QMessageBox.critical(self, '오류', f'폴더 생성 실패: {e}')
        elif action == '이름 변경':
            if not item:
                QMessageBox.warning(self, '경고', '트리에서 이름을 바꿀 파일/폴더를 선택하세요.')
                return
            new_name = option
            try:
                os.rename(path, os.path.join(os.path.dirname(path), new_name))
                self.update_file_list()
            except Exception as e:
                QMessageBox.critical(self, '오류', f'이름 변경 실패: {e}')
        elif action == '삭제':
            import shutil
            try:
                if os.path.isdir(path):
                    shutil.rmtree(path)
                else:
                    os.remove(path)
                self.update_file_list()
            except Exception as e:
                QMessageBox.critical(self, '오류', f'삭제 실패: {e}')
        elif action == '탐색기에서 열기':
            import subprocess
            try:
                subprocess.Popen(f'explorer /select,"{path}"')
            except Exception as e:
                QMessageBox.critical(self, '오류', f'탐색기 열기 실패: {e}')

    def handle_git_action(self):
        self.update_git_terminal_command()
        action = self.git_action_combo.currentText()
        option = self.git_option_input.text().strip()
        item = self.git_file_tree.currentItem()
        repo_path = self.local_path_input.text()
        if not action or action == '명령어 선택':
            QMessageBox.information(self, '알림', '명령어를 선택하세요.')
            return
        if not option:
            QMessageBox.warning(self, '경고', '옵션(파일명/폴더명 등)을 입력하세요.')
            return
        path = self._get_full_path_from_tree(item, repo_path, git_tree=True) if item else os.path.join(repo_path, option)
        rel_path = os.path.relpath(path, repo_path)
        if action == 'Git add (추적 시작)':
            try:
                repo = Repo(repo_path)
                repo.git.add(rel_path)
                self.update_file_list()
            except Exception as e:
                QMessageBox.critical(self, '오류', f'Git add 실패: {e}')
        elif action == 'Git untrack (추적 해제)':
            try:
                repo = Repo(repo_path)
                repo.git.rm('--cached', rel_path)
                self.update_file_list()
            except Exception as e:
                QMessageBox.critical(self, '오류', f'Git untrack 실패: {e}')
        elif action == 'Git ignore에 추가':
            try:
                gitignore_path = os.path.join(repo_path, '.gitignore')
                with open(gitignore_path, 'a', encoding='utf-8') as f:
                    f.write(option + '\n')
                self.update_file_list()
            except Exception as e:
                QMessageBox.critical(self, '오류', f'Git ignore 추가 실패: {e}')

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