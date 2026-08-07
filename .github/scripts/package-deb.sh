#!/usr/bin/env bash
# 将 PyInstaller 产物 dist/StockWidget 打包成 Debian 包
set -euo pipefail

if [[ "${GITHUB_REF_NAME}" == v* ]]; then
  VER="${GITHUB_REF_NAME#v}"
else
  VER="dev-$(date +%Y%m%d)"
fi

ARCH="$(dpkg --print-architecture)"
PKG="stockwidget_${VER}_${ARCH}"

rm -rf "$PKG"
mkdir -p "$PKG/DEBIAN" "$PKG/opt/StockWidget" "$PKG/usr/bin" "$PKG/usr/share/applications"

# 把 PyInstaller 产物装进 /opt/StockWidget
cp -a dist/StockWidget/. "$PKG/opt/StockWidget/"

# 应用菜单入口
cat > "$PKG/usr/share/applications/stockwidget.desktop" <<'EOF'
[Desktop Entry]
Name=StockWidget
Comment=极简透明盯盘 Widget 浮窗
Exec=/opt/StockWidget/StockWidget
Type=Application
Terminal=false
Categories=Utility;
EOF

# 命令行入口: /usr/bin/stockwidget
ln -s /opt/StockWidget/StockWidget "$PKG/usr/bin/stockwidget"

# 包元信息
cat > "$PKG/DEBIAN/control" <<EOF
Package: stockwidget
Version: $VER
Section: utils
Priority: optional
Architecture: $ARCH
Depends: libxcb-cursor0, libxkbcommon0, libegl1, libgl1, libfontconfig1, libdbus-1-3
Maintainer: sbr0574 <sbr0574@users.noreply.github.com>
Description: 极简透明盯盘 Widget 浮窗
EOF

dpkg-deb --build "$PKG"
mv "$PKG.deb" "StockWidget-linux_${VER}_${ARCH}.deb"
echo "Built StockWidget-linux_${VER}_${ARCH}.deb"
