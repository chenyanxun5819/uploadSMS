import argparse
import getpass
import sys

# 範例的固定測試帳密（你可以改成想要的）
VALID_USERNAME = "admin"
VALID_PASSWORD = "secret"

def check_login(username: str, password: str) -> bool:
    return username == VALID_USERNAME and password == VALID_PASSWORD

def main() -> None:
    parser = argparse.ArgumentParser(description="測試登入腳本")
    parser.add_argument("--username", "-u", help="使用者名稱")
    parser.add_argument("--password", "-p", help="密碼 (不要在公開環境使用)")
    args = parser.parse_args()

    username = args.username if args.username is not None else input("使用者名稱: ")
    password = args.password if args.password is not None else getpass.getpass("密碼: ")

    if check_login(username, password):
        print("登入成功")
        sys.exit(0)
    else:
        print("登入失敗")
        sys.exit(1)

if __name__ == "__main__":
    main()
