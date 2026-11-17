#!/bin/sh
# Pack src/ into a signed kintone plugin zip. The key fixes the plugin ID.
set -eu
here=$(cd "$(dirname "$0")" && pwd)
key=${KEY:-/home/ubuntu/.oxi-keys/kintone-oxi-preview.ppk}
out=${1:-$here/oxi-kintone-plugin.zip}
b=$(mktemp -d)
(cd "$here/src" && zip -qrX "$b/contents.zip" .)
openssl rsa -in "$key" -pubout -outform DER -out "$b/PUBKEY" 2>/dev/null
openssl dgst -sha1 -sign "$key" -out "$b/SIGNATURE" "$b/contents.zip"
rm -f "$out"
(cd "$b" && zip -qX "$out" contents.zip PUBKEY SIGNATURE)
rm -rf "$b"
echo "$out"
