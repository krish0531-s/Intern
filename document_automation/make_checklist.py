from pathlib import Path

root = Path(
    input("Enter root folder path: ").strip()
)

with open("checklist.txt", "w", encoding="utf-8") as f:

    f.write("file_name|.\n")

    for folder in sorted(root.rglob("*")):

        if folder.is_dir():

            relative_path = folder.relative_to(root)

            f.write(
                f"file_name|{relative_path.as_posix()}\n"
            )

print("checklist.txt created")