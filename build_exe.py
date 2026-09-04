# -*- coding: utf-8 -*-
"""
csp_order_maker.py 를 onedir(폴더형) 실행 파일로 빌드하는 스크립트.

사용법
  python build_exe.py                        (기본 이름으로 빌드)
  python build_exe.py --name CSP주문접수생성기   (실행파일/프로세스 이름 직접 지정)
  python build_exe.py --name 주문접수 --icon app.ico --clean

빌드 결과물은 dist/<이름>/ 폴더에 만들어진다. 이 폴더 전체를 그대로
복사해서 배포하면 되고, 그 안의 <이름>.exe 가 실행 파일이다 (실행하면
그 폴더 위치에 'CSP 주문접수 UPLOAD' 폴더가 자동으로 생기고 생성 파일과
로그가 그 안에 쌓인다).

필요 패키지 : pyinstaller  (pip install pyinstaller)
"""

import argparse
import os
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SOURCE = os.path.join(HERE, "csp_order_maker.py")
DEFAULT_NAME = "CSP_주문접수_생성기"


def parse_args():
    parser = argparse.ArgumentParser(
        description="csp_order_maker.py 를 onedir 실행 파일로 빌드합니다.")
    parser.add_argument("--name", default=DEFAULT_NAME,
                        help="생성될 실행파일(프로세스) 이름. 기본값: %s" % DEFAULT_NAME)
    parser.add_argument("--icon", default=None,
                        help="실행파일 아이콘으로 쓸 .ico 파일 경로 (선택)")
    parser.add_argument("--clean", action="store_true",
                        help="빌드 전 기존 build/dist/*.spec 를 먼저 지운다")
    return parser.parse_args()


def check_pyinstaller():
    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        sys.exit(
            "PyInstaller가 설치되어 있지 않습니다.\n"
            "다음 명령으로 먼저 설치하세요:  pip install pyinstaller")


def clean_previous(name):
    for path in (
        os.path.join(HERE, "build"),
        os.path.join(HERE, "dist"),
        os.path.join(HERE, "%s.spec" % name),
    ):
        if os.path.isdir(path):
            shutil.rmtree(path)
        elif os.path.exists(path):
            os.remove(path)


def main():
    args = parse_args()

    if not os.path.exists(SOURCE):
        sys.exit("csp_order_maker.py 를 찾을 수 없습니다: %s" % SOURCE)
    if args.icon and not os.path.exists(args.icon):
        sys.exit("아이콘 파일을 찾을 수 없습니다: %s" % args.icon)

    check_pyinstaller()

    if args.clean:
        clean_previous(args.name)

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onedir",            # 폴더형 배포 (실행 속도 빠르고 내부 구조 확인 쉬움)
        "--windowed",          # Tkinter GUI 이므로 콘솔창 없이 실행
        "--noconfirm",         # 기존 dist/build 를 덮어써도 되는지 매번 묻지 않음
        "--name", args.name,   # 실행파일/프로세스 이름을 직접 지정
        SOURCE,
    ]
    if args.icon:
        cmd += ["--icon", args.icon]

    print("실행 명령:", " ".join(cmd))
    subprocess.run(cmd, cwd=HERE, check=True)

    dist_dir = os.path.join(HERE, "dist", args.name)
    exe_name = "%s.exe" % args.name if sys.platform.startswith("win") else args.name
    print()
    print("빌드 완료 : %s" % dist_dir)
    print("이 폴더 전체를 배포하세요. 실행 파일 : %s" % os.path.join(dist_dir, exe_name))


if __name__ == "__main__":
    main()
