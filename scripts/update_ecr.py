import os
import random
import openpyxl

def generate_ecr(payload):
    
    print("Starting dynamic ECR generation...")

    # Retrieve template path from the payload
    template_path = payload.get("template_path")
    
    # Check if template_path was provided and if the file actually exists on disk
    if not template_path:
        raise ValueError("Missing required payload parameter: 'template_path'")
    
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"The specified template file does not exist: {template_path}")

    # Load the workbook safely
    wb = openpyxl.load_workbook(template_path)

    # Populate 'INPUT DATA' sheet from payload metadata
    input_sheet = wb['INPUT DATA']
    school_info = payload.get("school_info", {})

    print("Injecting metadata and teacher-defined configurations dynamically...")
    input_sheet['F10'].value = school_info.get("region")
    input_sheet['F11'].value = school_info.get("division")
    input_sheet['F13'].value = school_info.get("school_id")
    input_sheet['F14'].value = school_info.get("school_name")
    input_sheet['F16'].value = school_info.get("school_year")
    input_sheet['F22'].value = school_info.get("teacher")
    input_sheet['F24'].value = school_info.get("grade_level")
    input_sheet['F25'].value = school_info.get("section_name")
    input_sheet['F28'].value = school_info.get("subject_category")
    
    cluster_value = school_info.get("cluster", "")
    input_sheet['F29'].value = cluster_value
    input_sheet['F30'].value = school_info.get("subject")
    
    # --- CONDITIONAL CHECK FOR ELECTIVE FIELDS ---
    # Only populate specific elective details if the cluster matches the elective category
    if cluster_value == "OTHER ELECTIVE / SPECIAL CURRICULAR PROGRAM":
        print("Cluster is an Elective. Injecting elective-specific fields and weights...")
        
        # Elective Subject Name
        input_sheet['F39'].value = school_info.get("subject")
        
        # Teacher-defined Component Weights (F43 - F48)
        weights = school_info.get("component_weights", {})
        input_sheet['F43'].value = weights.get("written_works")
        input_sheet['F44'].value = weights.get("performance_tasks")
        input_sheet['F45'].value = weights.get("examinations")
        input_sheet['F46'].value = weights.get("summative_1")
        input_sheet['F47'].value = weights.get("summative_2")
        input_sheet['F48'].value = weights.get("term_exam")

        # Teacher-defined Terms and Units (F51 - F52)
        input_sheet['F51'].value = school_info.get("terms_taught")
        input_sheet['F52'].value = school_info.get("units_per_term")
    else:
        print("Cluster is NOT an elective. Skipping elective-specific cells.")

    # Separate students dynamically by gender
    students = payload.get("students", [])
    male_students = [s for s in students if s.get("sex", "").upper() == "M"]
    female_students = [s for s in students if s.get("sex", "").upper() == "F"]

    print(f"Injecting {len(male_students)} male and {len(female_students)} female students...")

    # Insert males (Columns N and O, starting at row 11)
    for i, student in enumerate(male_students):
        row = 11 + i
        input_sheet[f'N{row}'].value = student.get("lrn")
        input_sheet[f'O{row}'].value = student.get("name")

    # Insert females (Columns R and S, starting at row 11)
    for i, student in enumerate(female_students):
        row = 11 + i
        input_sheet[f'R{row}'].value = student.get("lrn")
        input_sheet[f'S{row}'].value = student.get("name")

    # Populate grades dynamically across term sheets (Term 1, Term 2, Term 3)
    terms_data = payload.get("terms_data", {})

    for term_name, term_info in terms_data.items():
        if term_name not in wb.sheetnames:
            print(f"Warning: Sheet '{term_name}' not found.")
            continue

        print(f"Populating scores for {term_name}...")
        term_sheet = wb[term_name]

        # Highest Possible Scores defined by the teacher for this term
        hps = term_info.get("hps", {})
        term_sheet['D14'].value = hps.get("ww")
        term_sheet['Q14'].value = hps.get("pt")
        term_sheet['AD14'].value = hps.get("exam")

        # Map scores for males
        for i, student in enumerate(male_students):
            row = 17 + i
            lrn = student.get("lrn")
            scores = term_info.get("scores", {}).get(lrn, {})
            
            if "ww" in scores: term_sheet[f'D{row}'].value = scores["ww"]
            if "pt" in scores: term_sheet[f'Q{row}'].value = scores["pt"]
            if "exam" in scores: term_sheet[f'AD{row}'].value = scores["exam"]

        # Map scores for females
        for i, student in enumerate(female_students):
            row = 68 + i
            lrn = student.get("lrn")
            scores = term_info.get("scores", {}).get(lrn, {})
            
            if "ww" in scores: term_sheet[f'D{row}'].value = scores["ww"]
            if "pt" in scores: term_sheet[f'Q{row}'].value = scores["pt"]
            if "exam" in scores: term_sheet[f'AD{row}'].value = scores["exam"]

    # Force Excel recalculation flag
    wb.properties.fullCalcOnLoad = True

    # Save output file with strict validation
    output_path = payload.get("output_path")
    
    if not output_path:
        raise ValueError("Missing required payload parameter: 'output_path'")

    wb.save(output_path)
    print(f"Successfully generated dynamic ECR file at: {output_path}")
    return output_path

# --- Local Testing Payload with Teacher Inputs for 3 Terms ---
if __name__ == '__main__':
    sample_payload = {
        "template_path": "./public/templates/ecr_shs_v26.xlsx",
        "school_info": {
            "region": "Region V",
            "division": "Iriga City",
            "school_id": "12345",
            "school_name": "Zeferino Arroyo High School",
            "school_year": "2026",
            "teacher": "Cruzata, Kevin Alejandro",
            "grade_level": 11,
            "section_name": "Diamond",
            "subject_category": "SSHS - CORE",
            "cluster": "CORE",
            "subject": "General Science",
            "component_weights": {
                "written_works": 0.25,
                "performance_tasks": 0.50,
                "examinations": 0.25,
                "summative_1": 0.05,
                "summative_2": 0.05,
                "term_exam": 0.15
            },
            "terms_taught": 3,
            "units_per_term": 2
        },
        "students": [
            {"lrn": "123501", "name": "Alvarez y de la Rosa, Juan Miguel", "sex": "M"},
            {"lrn": "223501", "name": "Santos-Ocampo, Maria Angelica Victoria", "sex": "F"}
        ],
        "terms_data": {
            "Term 1": {
                "hps": {"ww": 25, "pt": 50, "exam": 100},
                "scores": {
                    "123501": {"ww": 20, "pt": 45, "exam": 90},
                    "223501": {"ww": 24, "pt": 49, "exam": 98}
                }
            },
            "Term 2": {
                "hps": {"ww": 30, "pt": 60, "exam": 100},
                "scores": {
                    "123501": {"ww": 25, "pt": 52, "exam": 88},
                    "223501": {"ww": 28, "pt": 58, "exam": 95}
                }
            },
            "Term 3": {
                "hps": {"ww": 20, "pt": 40, "exam": 80},
                "scores": {
                    "123501": {"ww": 18, "pt": 36, "exam": 75},
                    "223501": {"ww": 19, "pt": 39, "exam": 79}
                }
            }
        },
        "output_path": "./test-output.xlsx"
    }

    generate_ecr(sample_payload)