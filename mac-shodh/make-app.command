#!/bin/bash
# एकदाच चालवा: Shodh.app बनवून Applications मध्ये टाकते. नंतर Launchpad/Dock मधून उघडा.
cd "$(dirname "$0")"
[ -d node_modules ] || npm install
npm run dist || exit 1
APP=$(ls -d dist/mac*/Shodh.app | head -1)
rm -rf /Applications/Shodh.app
cp -R "$APP" /Applications/ && xattr -cr /Applications/Shodh.app
echo "तयार: /Applications/Shodh.app"
open /Applications/Shodh.app
