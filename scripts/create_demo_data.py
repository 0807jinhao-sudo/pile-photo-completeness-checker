"""生成可公开使用的合成照片目录；图片内容仅为 1×1 像素占位图。"""

from __future__ import annotations

import base64
from pathlib import Path


PNG_1X1 = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
)

DEMO_FILES = (
    "2025-10-20/P001/P001对中.png",
    "2025-10-20/P001/P001焊接1.png",
    "2025-10-20/P001/P001钢筋笼验收.png",
    "2025-10-20/P001/P001初灌.png",
    "2025-10-20/P001/P001二清.png",
    "2025-10-20/P001/P001入岩.png",
    "2025-10-20/P001/P001终孔.png",
    "2025-10-21/P002/P002对中.png",
    "2025-10-21/P002/P002钢筋笼焊接.png",
    "2025-10-21/P002/P002钢筋笼验收.png",
    "2025-10-21/P002/P002初灌.png",
    "2025-10-21/P002/P002二清.png",
    "2025-10-21/P002/P002终孔.png",
    "2025-10-21/P002/P002旁站记录.png",
    "2025-10-21/P002/现场照片.png",
    "2025-10-22/P003/P003对中.png",
    "2025-10-22/P003/P003焊接.png",
    "2025-10-22/P003/P003钢筋笼验收.png",
    "2025-10-22/P003/P003初灌.png",
    "2025-10-22/P003/P003二清.png",
    "2025-10-22/P003/P003入岩.png",
    "2025-10-22/P003/P003D900终孔.png",
)


def main() -> None:
    project_root = Path(__file__).resolve().parents[1]
    demo_root = project_root / "examples" / "demo_input"
    for relative_path in DEMO_FILES:
        target = demo_root / relative_path
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(PNG_1X1)
    (demo_root / "SYNTHETIC_DATA.txt").write_text(
        "该目录由 scripts/create_demo_data.py 生成，全部为合成文件名与占位图片。\n",
        encoding="utf-8",
    )
    print(f"已生成 {len(DEMO_FILES)} 张合成占位图片：{demo_root}")


if __name__ == "__main__":
    main()
