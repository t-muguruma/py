import pathlib

root = pathlib.Path(".")
notebooks = sorted(root.rglob("*.ipynb"))
for nb in notebooks:
    print(nb)