import nbconvert as nb
import os
import subprocess

# Define the path to your .ipynb file
notebook_path = "/Users/karthik/Desktop/BCI-Internship/HKMAnalysis/HKM_DataAnalysis.ipynb"

input_file = notebook_path
output_file = notebook_path.split("/")[-1].split(".")[0] + ".pdf"

title = 'Analysis of Chanting Effects of Hare Krishna Mantra with EEG Aquisition System'
authors = ["Karthik M Dani", "Namyatha N Mulbagal"]
date = "2024-06-05"
abstract = ""

keywords = "EEG, Hare Krishna Mantra, Signal Acquisition, Brain Activity, EEG Band"

subtitle = "Comparative Study of Pre, During, and Post-Chanting Emotional Changes in Brain Activity through EEG"

thanks = ""
affiliation = ["Affiliation 1", "Affiliation 2"]
email = ["karthik.ml22@bmsce.ac.in", "namyatha.ml22@bmsce.ac.in"]
institute = ["BMS College of Engineering, Bangalore", "BMS College of Engineering, Bangalore"]

metadata = [
    f'--metadata=title={title}',
    f'--metadata=author={", ".join(authors)}',
    f'--metadata=date={date}',
    #f'--metadata=abstract={abstract}',
    f'--metadata=keywords={keywords}',
    f'--metadata=subtitle={subtitle}',
    f'--metadata=thanks={thanks}',
    f'--metadata=affiliation={", ".join(affiliation)}',
    f'--metadata=email={", ".join(email)}',
    f'--metadata=institute={", ".join(institute)}'
]

pandoc_command = ["pandoc", input_file, "-o", output_file] + metadata

try:
    subprocess.run(pandoc_command, check=True)
    print(f"Successfully converted {input_file} to {output_file}")
except subprocess.CalledProcessError as e:
    print(f"Error during conversion: {e}")

#os.system(pandoc_command)

