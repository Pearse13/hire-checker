"""Builds web/index.html (the site) from src/page.html, src/engine.js and src/app.js."""
import os
root = os.path.dirname(os.path.abspath(__file__))
read = lambda p: open(os.path.join(root, p), encoding="utf-8").read()
page = read("src/page.html").replace("/*ENGINE*/", read("src/engine.js")).replace("/*APP*/", read("src/app.js"))
open(os.path.join(root, "web/index.html"), "w", encoding="utf-8").write(page)
print("built web/index.html")
