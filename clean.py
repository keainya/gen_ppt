#!/usr/bin/env python3
"""Clean script — 清理项目中的生成文件和临时内容。

功能：
1. 清除 src/ 中除了 .gitkeep, background.png, cover.png 以外的所有文件
2. 清空 content.md
3. 删除 output.pptx
"""

import os
import shutil

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
SRC_DIR = os.path.join(PROJECT_ROOT, "src")
CONTENT_MD = os.path.join(PROJECT_ROOT, "content.md")
OUTPUT_PPTX = os.path.join(PROJECT_ROOT, "output.pptx")


def clean_src():
	"""清除 src 目录中除了保留文件以外的所有文件。"""
	if not os.path.isdir(SRC_DIR):
		print(f"[跳过] src 目录不存在: {SRC_DIR}")
		return

	keep_files = {".gitkeep", "background.png", "cover.png"}
	removed = []
	for entry in os.listdir(SRC_DIR):
		if entry in keep_files:
			continue
		entry_path = os.path.join(SRC_DIR, entry)
		try:
			if os.path.isfile(entry_path) or os.path.islink(entry_path):
				os.remove(entry_path)
			elif os.path.isdir(entry_path):
				shutil.rmtree(entry_path)
			removed.append(entry)
			print(f"[删除] src/{entry}")
		except Exception as e:
			print(f"[错误] 删除 src/{entry} 失败: {e}")

	if not removed:
		print("[信息] src 目录中无额外文件需清理。")


def clean_content_md():
	"""清空 content.md。"""
	try:
		open(CONTENT_MD, "w", encoding="utf-8").close()
		print(f"[清空] content.md")
	except Exception as e:
		print(f"[错误] 清空 content.md 失败: {e}")


def clean_output_pptx():
	"""删除 output.pptx。"""
	if not os.path.isfile(OUTPUT_PPTX):
		print(f"[跳过] output.pptx 不存在")
		return
	try:
		os.remove(OUTPUT_PPTX)
		print(f"[删除] output.pptx")
	except Exception as e:
		print(f"[错误] 删除 output.pptx 失败: {e}")


def main():
	print("=" * 40)
	print("  开始清理项目...")
	print("=" * 40)
	clean_src()
	clean_content_md()
	clean_output_pptx()
	print("=" * 40)
	print("  清理完成！")
	print("=" * 40)


if __name__ == "__main__":
	main()
