import os
import time

def auto_push():
    while True:
        os.system('git add .')
        os.system('git commit -m "자동 커밋"')
        os.system('git push origin main')
        print("자동 푸시 완료! 10분 후 다시 실행됩니다.")
        time.sleep(600)  # 10분(600초)마다 반복

if __name__ == "__main__":
    auto_push() 