#!/usr/bin/env python3
"""
钉钉表格 API 文档更新脚本

从钉钉官方 OSS 下载最新的 API 文档并转换为 Markdown 格式。

依赖：
    pip install markitdown requests

用法：
    # 更新所有文档
    python scripts/update_api_docs.py

    # 更新指定文档
    python scripts/update_api_docs.py --api Workbook Sheet Range

    # 列出所有可用的 API
    python scripts/update_api_docs.py --list
"""

import argparse
import os
import subprocess
import sys
import tempfile
from pathlib import Path

try:
    import requests
except ImportError:
    print("错误：需要安装 requests 库")
    print("运行：pip install requests")
    sys.exit(1)

# API 文档配置：API名称 -> HTML文件名
# 值为 None 表示该 API 在官方 OSS 上没有单独的文档页面
API_DOCS = {
    # 核心对象
    "Workbook": "Workbook",
    "Sheet": "Sheet",
    "Range": "Range",
    "RangeList": "RangeList",
    # 用户交互（官方 OSS 无单独页面，使用本地手写文档）
    "Input": None,
    "Output": None,
    # 筛选相关
    "Filter": "Filter-api",
    "FilterCriteria": "FilterCriteria",
    "FilterCriteriaBuilder": "FilterCriteriaBuilder",
    "FilterCondition": "FilterCondition",
    # 数据和格式选项
    "SetValueOptions": "SetValueOptions",
    "DropdownListOption": "DropdownListOption",
    "BorderType": "BorderType",
    "Color": "Color",
    "SearchOptions": "SearchOptions",
    "SortField": "SortField",
    # 其他
    "Hyperlink": "Hyperlink-script",
}

BASE_URL = "https://icms-document.oss-cn-beijing.aliyuncs.com/zh-CN/dingtalk/orgapp/topics"


def check_markitdown():
    """检查 markitdown 是否已安装"""
    try:
        subprocess.run(
            ["markitdown", "--help"],
            capture_output=True,
            check=True
        )
        return True
    except (subprocess.CalledProcessError, FileNotFoundError):
        return False


def download_html(api_name: str, html_name: str) -> str | None:
    """下载 API 文档的 HTML 文件"""
    url = f"{BASE_URL}/{html_name}.html"
    print(f"  下载: {url}")

    try:
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        return response.text
    except requests.RequestException as e:
        print(f"  错误: 下载失败 - {e}")
        return None


def convert_to_markdown(html_content: str, output_path: Path) -> bool:
    """使用 markitdown 将 HTML 转换为 Markdown"""
    with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False) as f:
        f.write(html_content)
        temp_html = f.name

    try:
        result = subprocess.run(
            ["markitdown", temp_html, "-o", str(output_path)],
            capture_output=True,
            text=True
        )

        if result.returncode != 0:
            print(f"  错误: 转换失败 - {result.stderr}")
            return False

        return True
    finally:
        os.unlink(temp_html)


def update_api_doc(api_name: str, output_dir: Path) -> bool:
    """更新单个 API 文档"""
    if api_name not in API_DOCS:
        print(f"  错误: 未知的 API '{api_name}'")
        return False

    html_name = API_DOCS[api_name]

    # 检查是否有官方文档
    if html_name is None:
        print(f"\n跳过 {api_name}... (官方 OSS 无单独页面，保留本地文档)")
        return True

    output_path = output_dir / f"{api_name}.md"

    print(f"\n更新 {api_name}...")

    # 下载 HTML
    html_content = download_html(api_name, html_name)
    if not html_content:
        return False

    # 转换为 Markdown
    if not convert_to_markdown(html_content, output_path):
        return False

    print(f"  完成: {output_path}")
    return True


def main():
    parser = argparse.ArgumentParser(
        description="更新钉钉表格 API 文档",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__
    )
    parser.add_argument(
        "--api",
        nargs="+",
        help="指定要更新的 API 名称（默认更新全部）"
    )
    parser.add_argument(
        "--list",
        action="store_true",
        help="列出所有可用的 API"
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="输出目录（默认为 docs/api）"
    )

    args = parser.parse_args()

    # 列出所有 API
    if args.list:
        print("可用的 API 文档：")
        print("\n核心对象：")
        for api in ["Workbook", "Sheet", "Range", "RangeList"]:
            print(f"  - {api}")
        print("\n用户交互（本地文档，无法从 OSS 更新）：")
        for api in ["Input", "Output"]:
            print(f"  - {api}")
        print("\n筛选相关：")
        for api in ["Filter", "FilterCriteria", "FilterCriteriaBuilder", "FilterCondition"]:
            print(f"  - {api}")
        print("\n数据和格式选项：")
        for api in ["SetValueOptions", "DropdownListOption", "BorderType", "Color", "SearchOptions", "SortField"]:
            print(f"  - {api}")
        print("\n其他：")
        for api in ["Hyperlink"]:
            print(f"  - {api}")
        return

    # 检查 markitdown
    if not check_markitdown():
        print("错误：未安装 markitdown")
        print("请运行：pip install markitdown")
        sys.exit(1)

    # 确定输出目录
    if args.output:
        output_dir = args.output
    else:
        # 默认为脚本所在目录的 ../docs/api
        script_dir = Path(__file__).parent
        output_dir = script_dir.parent / "docs" / "api"

    output_dir.mkdir(parents=True, exist_ok=True)
    print(f"输出目录: {output_dir}")

    # 确定要更新的 API
    apis_to_update = args.api if args.api else list(API_DOCS.keys())

    # 更新文档
    success_count = 0
    fail_count = 0

    for api_name in apis_to_update:
        if update_api_doc(api_name, output_dir):
            success_count += 1
        else:
            fail_count += 1

    # 汇总
    print(f"\n完成！成功: {success_count}, 失败: {fail_count}")

    if fail_count > 0:
        sys.exit(1)


if __name__ == "__main__":
    main()
