import os
import sys
import subprocess

# Automatically install missing packages
def install_and_import(package):
    try:
        __import__(package)
    except ImportError:
        print(f"Installing {package}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
    finally:
        globals()[package] = __import__(package)

install_and_import("pandas")
install_and_import("docxtpl")

import pandas as pd
from docxtpl import DocxTemplate

def create_lesson_plan(template_path, output_path, context):
    doc = DocxTemplate(template_path)
    doc.render(context)
    doc.save(output_path)

def process_csv(csv_path, template_path, output_dir):
    df = pd.read_csv(csv_path)

    for index, row in df.iterrows():
        context = {}
        for col in df.columns:
            # Split semicolon-separated values into lists for bullet points
            if isinstance(row[col], str) and ';' in row[col]:
                context[col] = [item.strip() for item in row[col].split(';')]
            else:
                context[col] = row[col]

        # Sanitize filename
        competency = row['LearningCompetency']
        sanitized_competency = "".join([c for c in competency if c.isalpha() or c.isdigit() or c.isspace()]).rstrip()
        output_filename = os.path.join(output_dir, f"{sanitized_competency}.docx")

        create_lesson_plan(template_path, output_filename, context)
        print(f"Created lesson plan: {output_filename}")

if __name__ == "__main__":
    ROOT_DIR = os.path.dirname(os.path.abspath(__file__))

    csv_file = os.path.join(ROOT_DIR, "learningcompetenciesg7q1.csv")
    template_file = os.path.join(ROOT_DIR, "lesson_plan.docx")
    output_directory = os.path.join(ROOT_DIR, "GeneratedLessonPlans")

    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    process_csv(csv_file, template_file, output_directory)
