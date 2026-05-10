#!/usr/bin/env python3
import urllib.request
import json
import base64
import os

BASE_URL = "https://urbproject.com/wp-json/wp/v2"
user = "caretaker_of_the_complex"
password = "v2Az rLQZ q89G EQmx WO0X SKQF"
token = base64.b64encode(f"{user}:{password}".encode()).decode()
headers = {
    "Authorization": f"Basic {token}",
    "Content-Type": "application/json",
}

def wp_get(path):
    req = urllib.request.Request(BASE_URL + path, headers=headers)
    with urllib.request.urlopen(req) as res:
        return json.load(res)

def wp_post(path, data, method="POST"):
    req = urllib.request.Request(
        BASE_URL + path,
        data=json.dumps(data).encode("utf-8"),
        method=method,
        headers=headers,
    )
    with urllib.request.urlopen(req) as res:
        return json.load(res)

# 全豪オープンカテゴリーを検索
print("カテゴリーを検索中...")
cats = wp_get("/categories?per_page=50")
aus_cat = None
for c in cats:
    print(f"  ID:{c['id']} {c['name']}")
    if "全豪" in c["name"] or "Australian" in c["name"] or "AUS" in c["name"].upper():
        aus_cat = c
        break

if aus_cat is None:
    print("\n全豪オープンカテゴリーが見つかりませんでした。")
    print("上のリストからIDを確認して、スクリプトの cat_id を手動で設定してください。")
    cat_id = int(input("カテゴリーID: "))
else:
    cat_id = aus_cat["id"]
    print(f"\n全豪オープンカテゴリー検出: ID={cat_id} / {aus_cat['name']}")

# コンテンツ読み込み
script_dir = os.path.dirname(os.path.abspath(__file__))
content_path = os.path.join(script_dir, "content", "aus_open_reference.html")
with open(content_path, encoding="utf-8") as f:
    content = f.read()

# 投稿作成
print("\n投稿を作成中...")
result = wp_post("/posts", {
    "title": "【全豪オープン】会場・気候・観客 ― 執筆資料",
    "content": content,
    "status": "draft",
    "categories": [cat_id],
})
print(f"完了! ID: {result['id']}")
print(f"Edit URL: https://urbproject.com/wp-admin/post.php?post={result['id']}&action=edit")
