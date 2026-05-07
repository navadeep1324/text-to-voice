#!/bin/sh
set -eu

cd "$(dirname "$0")/.."
node scripts/generate-sharepoint-audio.cjs
