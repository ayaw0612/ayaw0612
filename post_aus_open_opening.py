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

def wp_post(path, data):
    req = urllib.request.Request(
        BASE_URL + path,
        data=json.dumps(data).encode("utf-8"),
        method="POST",
        headers=headers,
    )
    with urllib.request.urlopen(req) as res:
        return json.load(res)

# 全豪オープンカテゴリーを検索
cats = wp_get("/categories?per_page=50")
cat_id = None
for c in cats:
    if "全豪" in c["name"] or "Australian" in c["name"]:
        cat_id = c["id"]
        print(f"カテゴリー: {c['name']} (ID:{cat_id})")
        break

if cat_id is None:
    print("カテゴリーが見つかりませんでした。カテゴリーIDを入力してください:")
    for c in cats:
        print(f"  ID:{c['id']} {c['name']}")
    cat_id = int(input("ID: "))

script_dir = os.path.dirname(os.path.abspath(__file__))
content_path = os.path.join(script_dir, "content", "aus_open_opening.html")
with open(content_path, encoding="utf-8") as f:
    content = f.read()

result = wp_post("/posts", {
    "title": "全豪オープン　冒頭",
    "content": content,
    "status": "draft",
    "categories": [cat_id],
})
print(f"完了! ID: {result['id']}")
print(f"Edit URL: https://urbproject.com/wp-admin/post.php?post={result['id']}&action=edit")
