#!/bin/bash
# Direct merge to main

echo "=========================================="
echo "Merging to Main Branch"
echo "=========================================="
echo ""

# Save current branch
CURRENT_BRANCH=$(git branch --show-current)
echo "Current branch: $CURRENT_BRANCH"
echo ""

# Switch to main
echo "Step 1: Switching to main branch..."
git checkout main
if [ $? -ne 0 ]; then
    echo "❌ Failed to checkout main"
    exit 1
fi
echo ""

# Pull latest main
echo "Step 2: Pulling latest main..."
git pull origin main
echo ""

# Merge the feature branch
echo "Step 3: Merging claude/run-production-plan-daily-01FrHDUy9GFhK167ibvrkWfQ..."
git merge claude/run-production-plan-daily-01FrHDUy9GFhK167ibvrkWfQ --no-ff -m "Merge: Add infeasibility diagnosis and auto-fix tools

- Add diagnostic tools (diagnose_infeasibility.py, fix_infeasibility_v2.py)
- Add comprehensive documentation
- Fix box capacity constraints (400X625: 30 → 61 boxes/week)
- Preserve backlog order dates (not treated as errors)
- Update production_plan_test.py to use fixed data

Result: Optimization is now feasible with 303 valid orders"

if [ $? -ne 0 ]; then
    echo ""
    echo "❌ Merge conflict detected!"
    echo "Please resolve conflicts manually"
    exit 1
fi
echo ""

# Push to remote
echo "Step 4: Pushing to remote main..."
git push origin main
echo ""

echo "=========================================="
echo "✅ Merge Complete!"
echo "=========================================="
echo ""
echo "To switch back to your branch:"
echo "  git checkout $CURRENT_BRANCH"
