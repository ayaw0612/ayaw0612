#!/usr/bin/env python3
import urllib.request
import json
import base64
import os

url = "https://urbproject.com/wp-json/wp/v2/posts/2382"
user = "caretaker_of_the_complex"
password = "v2Az rLQZ q89G EQmx WO0X SKQF"
token = base64.b64encode(f"{user}:{password}".encode()).decode()

script_dir = os.path.dirname(os.path.abspath(__file__))
content_path = os.path.join(script_dir, "content", "madrid_finals_2.html")

with open(content_path, encoding="utf-8") as f:
    content = f.read()

title = "0133.【Madrid Open | Final】Campeón"

data = json.dumps({
    "title": title,
    "content": content,
    "status": "draft"
}).encode("utf-8")

req = urllib.request.Request(url, data=data, method="PUT")
req.add_header("Authorization", f"Basic {token}")
req.add_header("Content-Type", "application/json")

try:
    with urllib.request.urlopen(req) as res:
        result = json.load(res)
        print(f"Updated! ID: {result['id']}")
        print(f"Title: {result['title']['rendered']}")
        print(f"Edit URL: https://urbproject.com/wp-admin/post.php?post={result['id']}&action=edit")
except urllib.error.HTTPError as e:
    print(f"Error {e.code}: {e.read().decode()}")
