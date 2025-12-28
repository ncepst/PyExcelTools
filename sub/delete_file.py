# delete_file.py
import os
import fnmatch

# 適宜変更してください(このパス以下のフォルダ内のファイルも削除対象になります)
folder = r"C:\path\to\folder"
# ---------------------------------------------------------------------

# —————————–
# 設定
# —————————–
delete_patterns = [
   "*.tmp",
   "log_*.txt",
   "backup_*.zip"
]

dry_run = True  # True: 一覧表示後に確認, False: 即削除

def is_delete_file(filename):
    """削除対象かどうか判定（ワイルドカード）"""
    return any(fnmatch.fnmatch(filename, pattern) for pattern in delete_patterns)

# -----------------------------
# 削除候補収集（深さ無制限）
# -----------------------------
delete_candidates = []

for root, dirs, files in os.walk(folder):
    for filename in files:
        if is_delete_file(filename):
            path = os.path.join(root, filename)
            delete_candidates.append(path)

# -----------------------------
# dry_run=True → 一覧表示 + 確認プロンプト
# -----------------------------
if dry_run:
    if delete_candidates:
        print("削除候補一覧:")
        for path in delete_candidates:
            print(f"- {path}")
        print(f"\n合計 {len(delete_candidates)} 件")

        # 確認プロンプト
        ans = input("\n上記のファイルを削除しますか？ [y/N]: ").strip().lower()
        if ans == "y":
            for path in delete_candidates:
                os.remove(path)
            print(f"{len(delete_candidates)} 件のファイルを削除しました。")
        else:
            print("削除をキャンセルしました。")
    else:
        print("削除対象はありません。")

# -----------------------------
# dry_run=False → 即削除
# -----------------------------
else:
    for path in delete_candidates:
        os.remove(path)

    print(f"{len(delete_candidates)} 件のファイルを削除しました。")

