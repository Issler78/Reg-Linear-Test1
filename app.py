import subprocess

if __name__ == "__main__":
    # se quiser criar datasets falsos
    #subprocess.run(["python", "dataset/create_excel_files.py"])

    subprocess.run(["python", "output/create_reports.py"])