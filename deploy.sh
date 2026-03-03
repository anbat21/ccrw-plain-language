#!/bin/bash

# ----------------------
# AZURE DEPLOYMENT SCRIPT
# ----------------------

# Deployment
# ----------

echo "Running deployment for Node.js app..."

# 1. Select node version
echo "Selecting Node.js version..."

# 2. Install npm packages
if [ -e "$DEPLOYMENT_TARGET/package.json" ]; then
  cd "$DEPLOYMENT_TARGET"
  echo "Running npm install --production..."
  npm install --production
  exitWithMessageOnError "npm install failed"
  cd - > /dev/null
fi

echo "Deployment complete!"
