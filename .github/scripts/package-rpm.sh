#!/usr/bin/env bash
# Package the existing PyInstaller onedir output without modifying its binaries.
set -euo pipefail

if [[ -n "${VERSION:-}" ]]; then
  VER="$VERSION"
elif [[ "${GITHUB_REF_NAME:-}" == v* ]]; then
  VER="${GITHUB_REF_NAME#v}"
else
  VER="0.0.0-$(date +%Y%m%d)"
fi

if [[ ! "$VER" =~ ^[0-9][0-9A-Za-z.+~_-]*$ ]]; then
  echo "Invalid RPM version: $VER" >&2
  exit 1
fi
# RPM versions cannot contain '-'; '~' sorts prereleases before the final version.
RPM_VER="${VER//-/~}"
ARCH="$(uname -m)"
case "$ARCH" in
  x86_64|aarch64) ;;
  *) echo "Unsupported RPM architecture: $ARCH" >&2; exit 1 ;;
esac
test -x dist/StockWidget/StockWidget

topdir="$(mktemp -d)"
trap 'rm -rf -- "$topdir"' EXIT
mkdir -p "$topdir"/{BUILD,BUILDROOT,RPMS,SOURCES,SPECS,SRPMS}
cp -a dist/StockWidget "$topdir/SOURCES/StockWidget"
cp README.md LICENSE NOTICE "$topdir/SOURCES/"

cat > "$topdir/SPECS/stockwidget.spec" <<'EOF'
# Stripping a PyInstaller bootloader can corrupt its embedded archive.
%global __os_install_post %{nil}
%global debug_package %{nil}
%global _build_id_links none

Name: stockwidget
Version: %{app_version}
Release: 1
Summary: Transparent desktop stock quote widget
License: Apache-2.0
URL: https://github.com/sbr0574/StockWidget
# Python, Qt and their private libraries are bundled under /opt. Do not expose
# them as system providers or require a system Python/Qt installation.
AutoReqProv: no
# The workflow builds on Ubuntu 22.04 (glibc 2.35). Runtime system libraries
# are expressed as SONAMEs so RPM-based distributions can resolve providers.
Requires: glibc >= 2.35
Requires: libEGL.so.1()(64bit), libGL.so.1()(64bit)
Requires: libxkbcommon.so.0()(64bit), libxkbcommon-x11.so.0()(64bit)
Requires: libdbus-1.so.3()(64bit), libfontconfig.so.1()(64bit)
Requires: libxcb.so.1()(64bit), libxcb-cursor.so.0()(64bit)
Requires: libxcb-icccm.so.4()(64bit), libxcb-image.so.0()(64bit)
Requires: libxcb-keysyms.so.1()(64bit), libxcb-render-util.so.0()(64bit)
Requires: libX11.so.6()(64bit), libXext.so.6()(64bit), libXtst.so.6()(64bit)

%description
StockWidget is a transparent desktop widget for monitoring stock quotes.

%install
mkdir -p "%{buildroot}/opt" "%{buildroot}/usr/bin"
mkdir -p "%{buildroot}/usr/share/applications" "%{buildroot}/usr/share/doc/stockwidget"
cp -a "%{_sourcedir}/StockWidget" "%{buildroot}/opt/StockWidget"
cp "%{_sourcedir}/README.md" "%{_sourcedir}/LICENSE" "%{_sourcedir}/NOTICE" "%{buildroot}/usr/share/doc/stockwidget/"
ln -s /opt/StockWidget/StockWidget "%{buildroot}/usr/bin/stockwidget"
install -Dm644 "%{_sourcedir}/StockWidget/_internal/icons/StockWidget.png" \
  "%{buildroot}/usr/share/pixmaps/stockwidget.png"
cat > "%{buildroot}/usr/share/applications/stockwidget.desktop" <<'DESKTOP'
[Desktop Entry]
Name=StockWidget
Comment=Transparent desktop stock quote widget
Exec=/opt/StockWidget/StockWidget
Icon=stockwidget
Type=Application
Terminal=false
Categories=Utility;
DESKTOP

%files
%defattr(-,root,root,-)
/opt/StockWidget
/usr/bin/stockwidget
/usr/share/applications/stockwidget.desktop
/usr/share/pixmaps/stockwidget.png
%dir /usr/share/doc/stockwidget
%doc /usr/share/doc/stockwidget/README.md
%license /usr/share/doc/stockwidget/LICENSE
%doc /usr/share/doc/stockwidget/NOTICE
EOF

rpmbuild -bb --define "_topdir $topdir" --define "app_version $RPM_VER" \
  --target "$ARCH" "$topdir/SPECS/stockwidget.spec"

package="StockWidget-linux_${VER}_${ARCH}.rpm"
cp "$topdir/RPMS/$ARCH/stockwidget-$RPM_VER-1.$ARCH.rpm" "$package"
rpm -qpi "$package"
rpm -qp --requires "$package"
echo "Built $package"
