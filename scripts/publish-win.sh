#!/usr/bin/env bash

set -euo pipefail

project_dir="$(cd "$(dirname "$0")/.." && pwd)"
publish_dir="$project_dir/src/ExcelTool/bin/Release/net10.0/win-x64/publish"

dotnet publish "$project_dir/src/ExcelTool/ExcelTool.csproj" \
  --configuration Release \
  --runtime win-x64 \
  --self-contained true \
  -p:PublishSingleFile=true \
  -p:EnableCompressionInSingleFile=true \
  -p:IncludeNativeLibrariesForSelfExtract=true \
  -p:DebugType=None \
  -p:DebugSymbols=false

# SkiaSharp 和 HarfBuzzSharp 的原生调试符号不影响程序运行，发布时移除。
find "$publish_dir" -maxdepth 1 -type f -name '*.pdb' -delete

echo "发布完成：$publish_dir/"
