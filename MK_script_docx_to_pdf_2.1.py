import os
import sys
import glob
import subprocess
import importlib
from datetime import datetime


def ensure_pywin32():
    """
    Ensure that pythoncom and win32com.client are available.
    If not, try to install pywin32 via pip and import again.
    Returns: (pythoncom, win32com.client)
    """
    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
        return pythoncom, win32com.client

    except ImportError:
        print("Required modules 'pythoncom' / 'win32com.client' are missing.")
        print("Attempting to install 'pywin32' via pip...\n")

        try:
            # Install pywin32 for the same Python interpreter that runs this script
            subprocess.check_call(
                [sys.executable, "-m", "pip", "install", "pywin32"]
            )
        except Exception as e:
            print("Failed to install 'pywin32' automatically.")
            print("Error details:", e)
            print("\nPlease run manually in your terminal / PowerShell:")
            print(f"  {os.path.basename(sys.executable)} -m pip install pywin32")
            input("\nPress Enter to exit...")
            sys.exit(1)

        print("\nSuccessfully installed 'pywin32'. Trying to import it...\n")

        try:
            pythoncom = importlib.import_module("pythoncom")
            win32com_client = importlib.import_module("win32com.client")
            return pythoncom, win32com_client
        except Exception as e:
            print("Installation finished, but 'pywin32' still cannot be imported.")
            print("Error details:", e)
            input("\nPress Enter to exit...")
            sys.exit(1)


def list_docx_files(base_dir: str):
    """
    Find all .docx files in the given folder, excluding temp files (~$...).
    """
    pattern = os.path.join(base_dir, "*.docx")
    docx_files = [
        f for f in glob.glob(pattern)
        if not os.path.basename(f).startswith("~$")
    ]
    return docx_files


def convert_all_docx_in_folder(base_dir: str) -> None:
    """
    Convert all .docx files in the given folder to PDF (1:1 using Microsoft Word).
    Shows basic progress information in percent based on number of files.
    """
    docx_files = list_docx_files(base_dir)

    if not docx_files:
        print("No .docx files found in this folder:")
        print(f"  {base_dir}")
        input("\nPress Enter to exit...")
        return

    total = len(docx_files)

    print("\nFiles to be converted:")
    if total == 1:
        print(f"  1 file found: {os.path.basename(docx_files[0])}")
    else:
        for f in docx_files:
            print("  -", os.path.basename(f))
        print(f"\nTotal files: {total}")

    # Ensure pywin32 is available
    pythoncom, win32com_client = ensure_pywin32()

    print("\nStarting conversion using Microsoft Word...")
    # Initialize COM and Word
    pythoncom.CoInitialize()
    word = win32com_client.DispatchEx("Word.Application")
    word.Visible = False

    try:
        for index, docx_path in enumerate(docx_files, start=1):
            try:
                pdf_path = os.path.splitext(docx_path)[0] + ".pdf"

                print(f"\n[{index}/{total}] Converting file:")
                print(f"  DOCX: {os.path.basename(docx_path)}")
                print(f"  PDF : {os.path.basename(pdf_path)}")

                doc = word.Documents.Open(docx_path)

                # 17 = wdExportFormatPDF
                wdExportFormatPDF = 17

                doc.ExportAsFixedFormat(
                    OutputFileName=pdf_path,
                    ExportFormat=wdExportFormatPDF,
                    OpenAfterExport=False,
                    OptimizeFor=0,      # 0 = wdExportOptimizeForPrint
                    CreateBookmarks=1   # 1 = wdExportCreateHeadingBookmarks
                )

                doc.Close(SaveChanges=False)

                percent = int(index / total * 100)
                print(f"  Status: OK   |   Progress: {percent}%")

            except Exception as e:
                percent = int(index / total * 100)
                print(f"  Status: ERROR (see details below) | Progress: {percent}%")
                print(f"  Error while processing {os.path.basename(docx_path)}:")
                print(f"    {e}")

    finally:
        word.Quit()
        pythoncom.CoUninitialize()

    print("\nAll conversions finished.")
    if total == 1:
        out_path = os.path.splitext(docx_files[0])[0] + ".pdf"
        print(f"Created PDF file: {os.path.basename(out_path)}")
    else:
        print("PDF files were created next to each original DOCX file.")

    input("\nPress Enter to exit...")


def print_intro(working_dir: str):
    """
    Print the initial information banner when the script/exe is started.
    """
    print("==========================================================")
    print(" MK DOCX to PDF Converter v2.1")
    print("----------------------------------------------------------")
    print(" This tool converts all DOCX files in this folder to PDF")
    print(" using Microsoft Word (Export as PDF, 1:1 formatting).")
    print("")
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f" Current date and time: {now_str}")
    print(f" Working folder: {working_dir}")
    print("==========================================================\n")


if __name__ == "__main__":
    working_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    print_intro(working_dir)
    convert_all_docx_in_folder(working_dir)
