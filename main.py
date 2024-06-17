import os
import comtypes.client

def pptx_to_pdf(input_file, output_file):
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
    ppt = powerpoint.Presentations.Open(input_file)
    ppt.SaveAs(output_file, FileFormat=32)  # 32 - PDF
    ppt.Close()
    powerpoint.Quit()

def main():
    input_folder = os.path.abspath("./resources/input")
    output_folder = os.path.abspath("./resources/output")
    os.makedirs(output_folder, exist_ok=True)
    for filename in os.listdir(input_folder):
        if filename.endswith(".pptx"):
            input_file = os.path.join(input_folder, filename)
            output_file = os.path.join(output_folder, f"{os.path.splitext(filename)[0]}.pdf")
            try:
                pptx_to_pdf(input_file, output_file)
                print(f"Successfully converted {filename} to PDF.")
            except Exception as e:
                print(f"Failed to convert {filename}: {e}")


if __name__ == "__main__":
    main()
