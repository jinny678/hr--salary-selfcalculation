#!/bin/bash
# 薪酬计算交叉验证工具 v1.0
# 广汽五羊-本田 正式工月度薪酬核验

echo "============================================"
echo "  薪酬计算交叉验证工具 v1.0"
echo "  广汽五羊-本田 正式工月度薪酬核验"
echo "============================================"
echo ""

# 检查 Python 是否安装
if ! command -v python3 &> /dev/null; then
    echo "[错误] 未检测到 python3，请先安装 Python 3.8+"
    exit 1
fi

# 检查并安装依赖
echo "[检查] 正在检查依赖包..."
pip3 install -r requirements.txt --quiet --break-system-packages 2>/dev/null || \
pip3 install -r requirements.txt --quiet

# 运行主程序
echo ""
python3 main.py
EXIT_CODE=$?

echo ""
if [ $EXIT_CODE -ne 0 ]; then
    echo "[错误] 程序运行异常，请检查上方错误信息。"
    exit $EXIT_CODE
fi

echo "============================================"
echo "  完成！报告已生成在 output 文件夹中。"
echo "============================================"