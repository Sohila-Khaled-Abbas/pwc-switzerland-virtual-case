import os
import shutil
from verify_complete_enterprise_app import run_qa_verification

def promote_build():
    base_dir = r"D:\courses\Data Analysis 26-27\7-Introducation to Data Fields (Excel)\11_Demos_and_Workbooks\10_Projects_and_Demos\PWC"
    staged_wb = os.path.join(base_dir, "scratch", "test_build_complete.xlsm")
    prod_wb = os.path.join(base_dir, "PWC_Switzerland_Virtual_Case.xlsm")

    print(f"Promoting staged build: {staged_wb} -> {prod_wb}")
    shutil.copy2(staged_wb, prod_wb)
    print("File promoted successfully! Running QA on production workbook...")

    success = run_qa_verification(prod_wb)
    if success:
        print("\nPROMOTION COMPLETE: Production workbook validated with 100% QA pass!")
    else:
        print("\nPROMOTION WARNING: Verification reported issues on production workbook.")

if __name__ == "__main__":
    promote_build()
