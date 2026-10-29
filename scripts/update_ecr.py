import random
import openpyxl

def main():
    print("Starting ECR automation script with Python (Component Weights & Terms Mapping)...")

    # 1. Load the original DepEd template
    template_path = './public/templates/ecr_shs_v26.xlsx'
    wb = openpyxl.load_workbook(template_path)

    # 2. Populate the 'INPUT DATA' sheet (Metadata, Component Weights, and Roster)
    input_sheet = wb['INPUT DATA']

    print("Inserting school, teacher, and subject metadata...")
    input_sheet['F10'] = "Region V"
    input_sheet['F11'] = "Iriga City"
    input_sheet['F13'] = "12345"
    input_sheet['F14'] = "Zeferino Arroyo High School"
    input_sheet['F16'] = "2026"
    input_sheet['F22'] = "Cruzata Kevin"
    input_sheet['F24'] = 11  # Grade level dropdown value
    input_sheet['F25'] = "Diamond"
    input_sheet['F28'] = "SSHS - ACADEMIC"
    input_sheet['F29'] = "OTHER ELECTIVE / SPECIAL CURRICULAR PROGRAM"
    input_sheet['F30'] = "German Language"
    
    # --- Injecting Elective and Component Weight fields from the new layout mapping ---
    input_sheet['F39'] = "German Language"  # Other Elective title
    
    # Component Weights percentages (F43 - F48)
    input_sheet['F43'] = 0.25  # Written/Oral Works (25%)
    input_sheet['F44'] = 0.50  # Performance Tasks (50%)
    input_sheet['F45'] = 0.25  # Examinations (25%)
    input_sheet['F46'] = 0.05  # Summative Test 1 (5%)
    input_sheet['F47'] = 0.05  # Summative Test 2 (5%)
    input_sheet['F48'] = 0.15  # Term Exam (15%)

    # Terms and Units Configuration (F51 - F52)
    input_sheet['F51'] = 3     # Terms Taught
    input_sheet['F52'] = 2     # Units per Term

    # 3. Define 10 Students (5 Males, 5 Females)
    print("Preparing 10 students roster...")

    male_students = [
        {"lrn": "123501", "name": "Alvarez, Lucas"},
        {"lrn": "123502", "name": "Bautista, Gabriel"},
        {"lrn": "123503", "name": "Cruzata, Kevin"},
        {"lrn": "123504", "name": "Dizon, Mateo"},
        {"lrn": "123505", "name": "Espiritu, Nathan"}
    ]

    female_students = [
        {"lrn": "223501", "name": "Santos, Bianca"},
        {"lrn": "223502", "name": "Reyes, Chloe"},
        {"lrn": "223503", "name": "Garcia, Danielle"},
        {"lrn": "223504", "name": "Mendoza, Erika"},
        {"lrn": "223505", "name": "Villanueva, Fiona"}
    ]

    # Insert male students (Columns N and O, starting at row 11)
    male_start_row = 11
    for i, student in enumerate(male_students):
        row = male_start_row + i
        input_sheet[f'N{row}'] = student["lrn"]
        input_sheet[f'O{row}'] = student["name"]

    # Insert female students (Columns R and S, starting at row 11)
    female_start_row = 11
    for i, student in enumerate(female_students):
        row = female_start_row + i
        input_sheet[f'R{row}'] = student["lrn"]
        input_sheet[f'S{row}'] = student["name"]

    # 4. Populate Grades across Term 1, Term 2, and Term 3
    terms = ['Term 1', 'Term 2', 'Term 3']
    
    for term_name in terms:
        if term_name not in wb.sheetnames:
            print(f"Warning: Sheet '{term_name}' not found in workbook.")
            continue

        print(f"Populating Highest Possible Scores and grades for {term_name}...")
        term_sheet = wb[term_name]

        # Standard HPS (Highest Possible Scores) configuration row
        term_sheet['D14'] = 25   # Written Works (WW) HPS
        term_sheet['Q14'] = 50   # Performance Tasks (PT) HPS
        term_sheet['AD14'] = 100 # Examination (Exam) HPS

        # Populate Male Student Grades (Rows 17 onwards)
        for i in range(len(male_students)):
            row = 17 + i
            term_sheet[f'D{row}'] = random.randint(18, 25)  # WW Score
            term_sheet[f'Q{row}'] = random.randint(35, 50)  # PT Score
            term_sheet[f'AD{row}'] = random.randint(75, 98) # Exam Score

        # Populate Female Student Grades (Rows 68 onwards)
        for i in range(len(female_students)):
            row = 68 + i
            term_sheet[f'D{row}'] = random.randint(19, 25)  # WW Score
            term_sheet[f'Q{row}'] = random.randint(38, 50)  # PT Score
            term_sheet[f'AD{row}'] = random.randint(80, 100)# Exam Score

    # Force Excel to recalculate formulas and validate data lists on load
    wb.properties.fullCalcOnLoad = True

    # 5. Save the final multi-term workbook with updated component weights
    output_path = './test-output.xlsx'
    wb.save(output_path)
    print(f"File successfully saved to {output_path} with 10 students and component weights mapped.")

if __name__ == '__main__':
    main()