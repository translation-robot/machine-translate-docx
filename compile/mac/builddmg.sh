#!/bin/sh
# Create a folder (named dmg) to prepare our DMG in (if it doesn't already exist).
mkdir -p dist/dmg
# Empty the dmg folder.
rm -r dist/dmg/*
# Copy the app bundle to the dmg folder.
cp -r "dist/Machine Translator.app" dist/dmg
# If the DMG already exists, delete it.
test -f "dist/Machine Translator.dmg" && rm "dist/Machine Translator.dmg"
create-dmg \
  --volname "Machine Translator" \
  --volicon "Machine Translator.icns" \
  --window-pos 200 120 \
  --window-size 600 300 \
  --icon-size 100 \
  --icon "Machine Translator.app" 175 120 \
  --hide-extension "Machine Translator.app" \
  --app-drop-link 425 120 \
  --no-mount "dist/Machine Translator.dmg" \
  "dist/dmg/"
