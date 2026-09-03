#!/bin/bash
# 双击此文件启动本地演示。首次运行会自动建立 Python 环境并安装依赖。
set -euo pipefail

PROJECT_DIR="$(cd "$(dirname "$0")" && pwd)"
BACKEND_DIR="$PROJECT_DIR/ran-backend"
FRONTEND_DIR="$PROJECT_DIR/ran-page 3"
VENV_DIR="$BACKEND_DIR/.venv"
API_PORT=8003
WEB_PORT=8081

cleanup() {
  [ -n "${API_PID:-}" ] && kill "$API_PID" 2>/dev/null || true
  [ -n "${WEB_PID:-}" ] && kill "$WEB_PID" 2>/dev/null || true
}
trap cleanup EXIT INT TERM

if ! command -v python3 >/dev/null 2>&1; then
  echo "未找到 Python 3。请先安装 Python 3 后再运行。"
  read -r -p "按回车键退出…"
  exit 1
fi

if [ ! -x "$VENV_DIR/bin/python" ]; then
  echo "正在创建本项目的 Python 环境（首次运行仅需一次）…"
  python3 -m venv "$VENV_DIR"
fi

echo "正在检查依赖…"
"$VENV_DIR/bin/python" -m pip install --quiet --upgrade pip
"$VENV_DIR/bin/python" -m pip install --quiet -r "$BACKEND_DIR/requirements.txt"

echo "正在启动后端…"
(
  cd "$BACKEND_DIR"
  exec "$VENV_DIR/bin/python" -m uvicorn main:app --host 127.0.0.1 --port "$API_PORT"
) &
API_PID=$!

for _ in {1..30}; do
  if curl --silent --fail "http://127.0.0.1:$API_PORT/trial-case" >/dev/null; then
    break
  fi
  sleep 1
done

if ! curl --silent --fail "http://127.0.0.1:$API_PORT/trial-case" >/dev/null; then
  echo "后端未能启动，请检查 .env 中的模型配置与终端错误信息。"
  read -r -p "按回车键退出…"
  exit 1
fi

echo "正在打开演示页面…"
python3 -m http.server "$WEB_PORT" --directory "$FRONTEND_DIR" >/dev/null 2>&1 &
WEB_PID=$!
open "http://127.0.0.1:$WEB_PORT/index.html"

echo "研行记已启动。关闭此终端窗口即可停止服务。"
wait "$API_PID"
