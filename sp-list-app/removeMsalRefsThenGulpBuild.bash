#!/usr/bin/bash
# Remove Msal* files & refrences then build
# WARN: Msal* files cause spfx-list-app bundling (gulp bundle) errors

echo "Remove Msal* files from mgwdev-m365-helpers package..."
find ./node_modules/mgwdev-m365-helpers -name Msal\* -exec rm {} \;

echo "Remove Msal* refs from files in mgwdev-m365-helpers package..."
grep -l --exclude=\*.md -r ./node_modules/mgwdev-m365-helpers/ -e Msal* | xargs sed -i '/Msal*/d'

echo "gulp build..."
gulp build