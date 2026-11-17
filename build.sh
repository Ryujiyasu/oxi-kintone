#!/bin/sh
# Assemble the plugin scripts from lib/. kintone loads plugin files as
# classic scripts, so the ES-module fill library is inlined.
set -eu
here=$(cd "$(dirname "$0")" && pwd)
ver=${1:?engine version, e.g. 0.9.0}
part() { grep -v '^// SPDX' "$1" | sed 's/^export //'; }
{
  echo '// SPDX-License-Identifier: MIT OR Apache-2.0'
  echo '// Built by build.sh from lib/. Do not edit by hand.'
  echo '(function () {'
  echo "'use strict';"
  part "$here/lib/oxi-fill.js"
  part "$here/lib/config-store.js"
  part "$here/lib/plugin-main.js" | sed "s/__ENGINE_VERSION__/$ver/"
  echo '})();'
} > "$here/src/js/desktop.js"
{
  echo '// SPDX-License-Identifier: MIT OR Apache-2.0'
  echo '// Built by build.sh from lib/. Do not edit by hand.'
  echo '(function () {'
  echo "'use strict';"
  part "$here/lib/config-store.js"
  part "$here/lib/config-main.js"
  echo '})();'
} > "$here/src/js/config.js"
node --check "$here/src/js/desktop.js"
node --check "$here/src/js/config.js"
echo built
