#!/bin/bash
# Finder मध्ये या फाइलवर डबल-क्लिक केले की ॲप सुरू होते
cd "$(dirname "$0")"
[ -d node_modules ] || npm install
npm start
