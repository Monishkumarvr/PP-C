#!/bin/bash
# Install HiGHS on WSL/Ubuntu for use with PuLP

echo "=========================================="
echo "Installing HiGHS for WSL/Ubuntu"
echo "=========================================="
echo ""

# Check OS
if [[ ! -f /etc/lsb-release ]]; then
    echo "⚠️  This script is for Ubuntu/Debian on WSL"
    echo "   For Windows native, see manual installation guide"
    exit 1
fi

# Update package list
echo "Step 1: Updating package list..."
sudo apt-get update -qq

# Install HiGHS
echo ""
echo "Step 2: Installing HiGHS executable..."
sudo apt-get install -y highs

# Verify installation
echo ""
echo "Step 3: Verifying installation..."
if command -v highs &> /dev/null; then
    echo "✅ HiGHS executable installed!"
    echo "   Location: $(which highs)"
    highs --version 2>/dev/null || echo "   Version check unavailable"
else
    echo "❌ Installation failed"
    echo ""
    echo "Alternative: Install from conda-forge"
    echo "  conda install -c conda-forge highs"
    exit 1
fi

# Test with PuLP
echo ""
echo "Step 4: Testing with PuLP..."
python3 << 'PYEOF'
from pulp import HiGHS_CMD
solver = HiGHS_CMD(msg=0)
if solver.available():
    print("✅ PuLP can use HiGHS!")
else:
    print("❌ PuLP cannot find HiGHS")
    print("   Try: sudo apt-get install highs")
PYEOF

echo ""
echo "=========================================="
echo "✅ Installation complete!"
echo "=========================================="
echo ""
echo "Next: Run test_highs_solver.py to compare with CBC"
