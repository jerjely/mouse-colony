#!/bin/bash
# Packs "dist/Mouse Colony.app" into an installer package that puts it in /Applications.
#   bash install/make_pkg.sh <version> <label> <host architectures>
#   e.g. bash install/make_pkg.sh 2.3.3 AppleSilicon arm64      -> dist/MouseColony-2.3.3-Mac-AppleSilicon.pkg
#        bash install/make_pkg.sh 2.3.3 Intel x86_64,arm64      -> dist/MouseColony-2.3.3-Mac-Intel.pkg
# Run on a Mac after PyInstaller (the GitHub workflow does both).
set -euo pipefail
VER="$1"; LABEL="$2"; HOSTARCH="$3"
ID="hu.pte.volgyilab.mousecolony"
APP="dist/Mouse Colony.app"
WORK="build/pkg-$LABEL"
OUT="dist/MouseColony-$VER-Mac-$LABEL.pkg"
[ -d "$APP" ] || { echo "Missing $APP - run PyInstaller first"; exit 1; }
rm -rf "$WORK"; mkdir -p "$WORK/root"
ditto "$APP" "$WORK/root/Mouse Colony.app"
# ad-hoc signature (required to run on Apple silicon; PyInstaller adds one, this renews it after copying)
codesign --force --deep --sign - "$WORK/root/Mouse Colony.app"
# always install into /Applications (don't "relocate" onto another copy found elsewhere on the disk)
pkgbuild --analyze --root "$WORK/root" "$WORK/component.plist"
/usr/libexec/PlistBuddy -c "Set :0:BundleIsRelocatable false" "$WORK/component.plist"
pkgbuild --root "$WORK/root" --component-plist "$WORK/component.plist" --install-location /Applications \
         --identifier "$ID" --version "$VER" "$WORK/MouseColony.pkg"
cat > "$WORK/distribution.xml" <<XML
<?xml version="1.0" encoding="utf-8"?>
<installer-gui-script minSpecVersion="2">
  <title>Mouse Colony $VER</title>
  <options customize="never" require-scripts="false" hostArchitectures="$HOSTARCH"/>
  <domains enable_localSystem="true"/>
  <choices-outline><line choice="default"><line choice="$ID"/></line></choices-outline>
  <choice id="default"/>
  <choice id="$ID" visible="false"><pkg-ref id="$ID"/></choice>
  <pkg-ref id="$ID" version="$VER" onConclusion="none">MouseColony.pkg</pkg-ref>
</installer-gui-script>
XML
productbuild --distribution "$WORK/distribution.xml" --package-path "$WORK" "$OUT"
echo "Built $OUT"
