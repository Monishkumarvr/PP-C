#!/bin/bash
# Install HiGHS solver - Modern free MILP solver

echo "=========================================="
echo "Installing HiGHS Solver"
echo "=========================================="
echo ""

# Install HiGHS
echo "Step 1: Installing highspy package..."
pip install highspy

echo ""
echo "Step 2: Verifying installation..."
python3 << 'PYEOF'
try:
    import highspy
    print("✅ HiGHS successfully installed!")
    print(f"   Version: {highspy.__version__ if hasattr(highspy, '__version__') else 'Unknown'}")
except ImportError:
    print("❌ Installation failed")
    exit(1)
PYEOF

echo ""
echo "=========================================="
echo "✅ HiGHS is ready to use!"
echo "=========================================="
echo ""
echo "Next steps:"
echo "1. Run: python3 test_highs_solver.py"
echo "2. If it works, update production_plan_test.py to use HiGHS"
