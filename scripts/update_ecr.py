import random
import openpyxl

def main():
    print("Starting ECR automation script with Python (Preserving Styles & Longer Names)...")

    # 1. Load the original DepEd template
    template_path = './public/templates/ecr_shs_v26.xlsx'
    wb = openpyxl.load_workbook(template_path)

    # 2. Populate the 'INPUT DATA' sheet safely via .value assignment
    input_sheet = wb['INPUT DATA']

    print("Inserting school, teacher, and subject metadata...")
    input_sheet['F10'].value = "Region V"
    input_sheet['F11'].value = "Iriga City"
    input_sheet['F13'].value = "12345"
    input_sheet['F14'].value = "Zeferino Arroyo High School"
    input_sheet['F16'].value = "2026"
    input_sheet['F22'].value = "Cruzata, Kevin Alejandro"
    input_sheet['F24'].value = 11  # Grade level dropdown value
    input_sheet['F25'].value = "Diamond - Section A"
    input_sheet['F28'].value = "SSHS - ACADEMIC"
    input_sheet['F29'].value = "OTHER ELECTIVE / SPECIAL CURRICULAR PROGRAM"
    input_sheet['F30'].value = "German Language Advanced"
    
    # Injecting Elective and Component Weight fields safely
    input_sheet['F39'].value = "German Language Advanced"  
    input_sheet['F43'].value = 0.25  # Written/Oral Works (25%)
    input_sheet['F44'].value = 0.50  # Performance Tasks (50%)
    input_sheet['F45'].value = 0.25  # Examinations (25%)
    input_sheet['F46'].value = 0.05  # Summative Test 1 (5%)
    input_sheet['F47'].value = 0.05  # Summative Test 2 (5%)
    input_sheet['F48'].value = 0.15  # Term Exam (15%)

    # Terms and Units Configuration
    input_sheet['F51'].value = 3     # Terms Taught
    input_sheet['F52'].value = 2     # Units per Term

    # 3. Define 10 Students with longer names
    print("Preparing 10 students roster with extended names...")

    male_students = [
        {"lrn": "123501", "name": "Alvarez y de la Rosa, Juan Miguel"},
        {"lrn": "123502", "name": "Bautista-Villarreal, Christopher James"},
        {"lrn": "123503", "name": "Cruzata, Kevin Alejandro"},
        {"lrn": "123504", "name": "Dizon de Guzman, Francisco Emmanuel"},
        {"lrn": "123505", "name": "Espiritu Santo, Maximilian George"}
    ]

    female_students = [
        {"lrn": "223501", "name": "Santos-Ocampo, Maria Angelica Victoria"},
        {"lrn": "223502", "name": "Reyes de Leon, Isabella Genevieve"},
        {"lrn": "223503", "name": "Garcia Fernandez, Sophia Bernadette"},
        {"lrn": "223504", "name": "Mendoza Alcantara, Katarina Alejandra"},
        {"lrn": "223505", "name": "Villanueva de la Vega, Alexandra Nicole"}
    ]

    # Insert male students safely using .value (Columns N and O, starting at row 11)
    male_start_row = 11
    for i, student in enumerate(male_students):
        row = male_start_row + i
        input_sheet[f'N{row}'].value = student["lrn"]
        input_sheet[f'O{row}'].value = student["name"]

    # Insert female students safely using .value (Columns R and S, starting at row 11)
    female_start_row = 11
    for i, student in enumerate(female_students):
        row = female_start_row + i
        input_sheet[f'R{row}'].value = student["lrn"]
        input_sheet[f'S{row}'].value = student["name"]

    # 4. Populate Grades across Term 1, Term 2, and Term 3 safely via .value
    terms = ['Term 1', 'Term 2', 'Term 3']
    
    for term_name in terms:
        if term_name not in wb.sheetnames:
            print(f"Warning: Sheet '{term_name}' not found in workbook.")
            continue

        print(f"Populating Highest Possible Scores and grades for {term_name}...")
        term_sheet = wb[term_name]

        # Standard HPS (Highest Possible Scores) configuration row
        term_sheet['D14'].value = 25   # Written Works (WW) HPS
        term_sheet['Q14'].value = 50   # Performance Tasks (PT) HPS
        term_sheet['AD14'].value = 100 # Examination (Exam) HPS

        # Populate Male Student Grades (Rows 17 onwards)
        for i in range(len(male_students)):
            row = 17 + i
            term_sheet[f'D{row}'].value = random.randint(18, 25)  # WW Score
            term_sheet[f'Q{row}'].value = random.randint(35, 50)  # PT Score
            term_sheet[f'AD{row}'].value = random.randint(75, 98) # Exam Score

        # Populate Female Student Grades (Rows 68 onwards)
        for i in range(len(female_students)):
            row = 68 + i
            term_sheet[f'D{row}'].value = random.randint(19, 25)  # WW Score
            term_sheet[f'Q{row}'].value = random.randint(38, 50)  # PT Score
            term_sheet[f'AD{row}'].value = random.randint(80, 100)# Exam Score

    # Force Excel to recalculate formulas and validate data lists on load
    wb.properties.fullCalcOnLoad = True

    # 5. Save the final multi-term workbook
    output_path = './test-output.xlsx'
    wb.save(output_path)
    print(f"File successfully saved to {output_path} with styling completely intact.")

if __name__ == '__main__':
    main()