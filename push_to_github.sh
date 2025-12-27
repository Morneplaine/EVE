#!/bin/bash
# Quick script to push changes to GitHub
echo "Pushing to GitHub..."
git add .
git commit -m "Update: $(date)" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "No changes to commit."
else
    echo "Changes committed."
fi
git push
if [ $? -eq 0 ]; then
    echo "Successfully pushed to GitHub!"
else
    echo "Push failed. Check your connection and try again."
fi

