import os

for root, dirs, files in os.walk("."):
    # skip scratch and env/venv
    if "scratch" in root or ".git" in root or "__pycache__" in root:
        continue
    for file in files:
        if file.endswith(".py"):
            path = os.path.join(root, file)
            with open(path, "r", encoding="utf-8", errors="ignore") as f:
                for idx, line in enumerate(f, 1):
                    if "_cells" in line:
                        print(f"{path}:{idx}: {line.strip()}")
