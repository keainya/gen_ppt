#!/bin/bash
set -e

cd "$(dirname "$0")"

if [ ! -f ".venv/bin/activate" ]; then
    echo "🔧 .venv 不完整或缺失，重新创建..."
    rm -rf .venv
    python3 -m venv .venv
fi

source .venv/bin/activate
pip install -q python-pptx

echo "✅ 环境就绪"
