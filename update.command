#!/bin/bash
# Theia Updater for macOS
# Double-click to update from the latest release branch and reinstall

set -Euo pipefail

UPDATER_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

pause_before_exit() {
    echo ""
    read -r -p "Press Enter to exit..." _ || true
}

handle_error() {
    local exit_code=$?
    trap - ERR
    echo ""
    echo "ERROR: Theia could not be updated."
    echo "Review the message above, then try again."
    pause_before_exit
    exit "$exit_code"
}

die() {
    echo "ERROR: $1"
    pause_before_exit
    exit 1
}

trap handle_error ERR

echo "======================================"
echo "Theia Updater"
echo "======================================"
echo ""

command -v git >/dev/null 2>&1 || die "Git is not installed. Install Xcode Command Line Tools and try again."

cd "$UPDATER_DIR"

git rev-parse --is-inside-work-tree >/dev/null 2>&1 || die "This updater must be run from a Git checkout of Theia."
git remote get-url origin >/dev/null 2>&1 || die "The Git checkout does not have an 'origin' remote."

# Never discard or overwrite work in the checkout.
if [ -n "$(git status --porcelain)" ]; then
    echo "Your Theia checkout contains local changes:"
    git status --short
    echo ""
    die "Commit, stash, or remove these changes before updating."
fi

echo "Checking GitHub for release branches..."
git fetch --prune origin "+refs/heads/release/*:refs/remotes/origin/release/*"

LATEST_REMOTE_REF="$(git for-each-ref \
    --count=1 \
    --sort=-version:refname \
    --format='%(refname:short)' \
    'refs/remotes/origin/release/*')"

[ -n "$LATEST_REMOTE_REF" ] || die "No release branches were found on origin."

LATEST_BRANCH="${LATEST_REMOTE_REF#origin/}"
echo "Latest release: $LATEST_BRANCH"

if git show-ref --verify --quiet "refs/heads/$LATEST_BRANCH"; then
    # Confirm the local release branch can be updated without rewriting history.
    if ! git merge-base --is-ancestor "$LATEST_BRANCH" "$LATEST_REMOTE_REF"; then
        die "Local branch '$LATEST_BRANCH' has commits that are not on GitHub; update stopped to protect them."
    fi

    git checkout "$LATEST_BRANCH"
else
    git checkout --track -b "$LATEST_BRANCH" "$LATEST_REMOTE_REF"
fi

git merge --ff-only "$LATEST_REMOTE_REF"

INSTALLER="$UPDATER_DIR/install.command"
[ -f "$INSTALLER" ] || die "The updated release does not contain install.command."

echo ""
echo "Source updated. Installing $LATEST_BRANCH..."
echo ""
bash "$INSTALLER"

echo ""
echo "======================================"
echo "Update Complete!"
echo "======================================"
echo "Theia is now installed from $LATEST_BRANCH."

trap - ERR
pause_before_exit
