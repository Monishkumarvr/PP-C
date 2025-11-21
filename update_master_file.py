"""
Simple script to update the master data file reference in production_plan_test.py
"""

import sys

def update_master_file_reference(new_file_name):
    """Update the master file reference in production_plan_test.py"""

    script_path = 'production_plan_test.py'

    # Read the file
    with open(script_path, 'r') as f:
        content = f.read()

    # Find and replace the file path line
    old_line = "    file_path = 'Master_Data_Updated_Nov_Dec.xlsx'"
    new_line = f"    file_path = '{new_file_name}'"

    if old_line in content:
        content = content.replace(old_line, new_line)

        # Write back
        with open(script_path, 'w') as f:
            f.write(content)

        print(f"✅ Updated production_plan_test.py")
        print(f"   Old: {old_line}")
        print(f"   New: {new_line}")
        print()
        print("You can now run:")
        print("   python3 production_plan_test.py")
    else:
        print(f"❌ Could not find the expected line in {script_path}")
        print(f"   Looking for: {old_line}")
        print()
        print("You may need to manually update line ~2879 in production_plan_test.py")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage:")
        print("   python3 update_master_file.py <new_file_name.xlsx>")
        print()
        print("Examples:")
        print("   python3 update_master_file.py Master_Data_Updated_Nov_Dec2.xlsx")
        print("   python3 update_master_file.py Master_Data_Updated_Nov_Dec_FIXED_V2.xlsx")
    else:
        new_file = sys.argv[1]
        update_master_file_reference(new_file)
