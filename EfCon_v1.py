# FILE CONVERTER TOOL (Terminal-Based Version)
# Description: Converts PPTX and DOCX files in bulk to PDF format using Microsoft Office COM automation.
# Requirements: Python, Microsoft Office (Word & PowerPoint), pywin32

import os
import win32com.client

# PowerPoint constants
ppSaveAsPDF = 32
ppFixedFormatTypePDF = 2

# Menu for user interaction
def menu():
    print("<><><> FILE CONVERTER <><><>")
    print("1. Convert PPT to PDF")
    print("2. Convert DOCX to PDF")
    print("3. EXIT APPLICATION")

# Get a valid file path from user
def get_file_path(file_type):
    while True:
        file_path = input(f"Enter the path to your {file_type} file: ").strip('"')
        if os.path.exists(file_path):
            return os.path.abspath(file_path)
        print("Error: File not found. Please check the path and try again.")

# Get a valid folder path from user
def get_folder_path(prompt):
    while True:
        folder_path = input(prompt).strip('"')
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            return os.path.abspath(folder_path)
        print("Error: Folder not found or invalid. Please check the path and try again.")

# Convert all PPT/PPTX files in a folder to PDF
def ppt_2pdf_bulk():
    powerpoint = None
    try:
        input_folder = get_folder_path("Enter the path to the folder containing PPT files: ")
        output_folder = get_folder_path("Enter the destination folder to save PDFs: ")
        print(f"Input folder path: {input_folder}")

        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        powerpoint = win32com.client.Dispatch("Powerpoint.Application")
        print("PowerPoint initialized")

        files = [f for f in os.listdir(input_folder) if f.lower().endswith(('.ppt', '.pptx'))]
        if not files:
            print("No PowerPoint files found in the folder.")
            return

        for index, file in enumerate(sorted(files), start=1):
            try:
                input_file = os.path.join(input_folder, file)
                output_file = os.path.join(output_folder, f"{index}.pdf")

                print(f"Converting: {input_file} -> {output_file}")
                deck = powerpoint.Presentations.Open(input_file)

                try:
                    deck.SaveAs(output_file, ppSaveAsPDF)
                except:
                    print("First save method failed, trying alternative method...")
                    deck.ExportAsFixedFormat(
                        output_file,
                        ppFixedFormatTypePDF,
                        Intent=1,
                        FrameSlides=True,
                        HandoutOrder=1,
                        OutputType=1
                    )

                deck.Close()
                print(f"Successfully converted: {file}")

            except Exception as e:
                print(f"Error converting {file}: {e}")

        print("All files processed.")

    except Exception as e:
        print(f"Error: {str(e)}")

    finally:
        if powerpoint:
            try:
                powerpoint.Quit()
                print("PowerPoint application closed")
            except:
                print("Error while closing PowerPoint")

# Convert all DOCX files in a folder to PDF
def docx_2pdf():
    word = None
    try:
        input_folder = get_folder_path("Enter the path to the folder containing DOCX files: ")
        output_folder = get_folder_path("Enter the destination folder to save PDFs: ")
        print(f"Input folder path: {input_folder}")

        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        word = win32com.client.Dispatch("Word.Application")
        print("Word initialized")

        files = [f for f in os.listdir(input_folder) if f.lower().endswith('.docx')]
        if not files:
            print("No Word files found in the folder.")
            return

        for index, file in enumerate(sorted(files), start=1):
            try:
                input_file = os.path.join(input_folder, file)
                output_file = os.path.join(output_folder, f"{index}.pdf")

                print(f"Converting: {input_file} -> {output_file}")
                doc = word.Documents.Open(input_file)
                doc.SaveAs(output_file, FileFormat=17)  # 17 = wdFormatPDF
                doc.Close()
                print(f"Successfully converted: {file}")

            except Exception as e:
                print(f"Error converting {file}: {e}")

        print("All files processed.")

    except Exception as e:
        print(f"Error: {str(e)}")

    finally:
        if word:
            try:
                word.Quit()
                print("Word application closed")
            except:
                print("Error while closing Word")

# Main loop
while True:
    menu()
    choice = input("Enter your choice: ")

    if choice == "1":
        ppt_2pdf_bulk()
    elif choice == "2":
        docx_2pdf()
    elif choice == "3":
        print("NOW EXITING\n. . .\n-----")
        break
    else:
        print("Invalid Choice. Please try again")
