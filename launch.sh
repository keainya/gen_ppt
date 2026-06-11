#!/bin/bash
set -e

cd "$(dirname "$0")"

if [ ! -d ".venv" ]; then
    echo "🔧 创建 Python 虚拟环境..."
    python3 -m venv .venv
fi

source .venv/bin/activate
pip install -q python-pptx

echo "✅ 环境就绪"
