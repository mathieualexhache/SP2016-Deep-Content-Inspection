import argparse
import subprocess
import olefile
import zipfile
import logging
import importlib
import sys
from oletools.olevba import VBA_Parser, TYPE_OLE, TYPE_OpenXML, TYPE_Word2003_XML, TYPE_MHTML

required_packages = [
    "olefile",
    "oletools",  # covers olevba and related modules
]

missing = []

for pkg in required_packages:
    if importlib.util.find_spec(pkg) is None:
        missing.append(pkg)

if missing:
    print(f"Missing required packages: {', '.join(missing)}")
    print("Please install them using: pip install " + ' '.join(missing))
    sys.exit(1)
    
def run_olevba_triage(path):
    try:
        result = subprocess.run(['olevba', '-t', path],
                                stdout=subprocess.PIPE,
                                stderr=subprocess.PIPE,
                                text=True)
        output = result.stdout.lower()
        return 'macros' in output or 'autoexec' in output or 'vba' in output
    except:
        return False

def run_oledump_check(path):
    try:
        result = subprocess.run(['oledump.py', path],
                                stdout=subprocess.PIPE,
                                stderr=subprocess.PIPE,
                                text=True)
        return any(
            line.lstrip().split()[1].lower() == 'm'
            for line in result.stdout.splitlines()
            if len(line.lstrip().split()) > 1
        )
    except Exception:
        return False

def check_excel97_macros(path):
    try:
        if olefile.isOleFile(path):
            ole = olefile.OleFileIO(path)
            return ole.exists('workbook') and ole.exists('_VBA_PROJECT_CUR')
    except:
        pass
    return False

def run_vba_parser(path):
    try:
        vbaparser = VBA_Parser(path)
        has_macros = vbaparser.detect_macros()
        vbaparser.close()
        return has_macros
    except:
        return False

def scan_zip_for_macros(path):
    if not zipfile.is_zipfile(path):
        return False

    try:
        with zipfile.ZipFile(path) as z:
            for subfile in z.namelist():
                try:
                    # Read magic bytes to identify OLE files
                    magic = z.open(subfile).read(len(olefile.MAGIC))
                    if magic == olefile.MAGIC:
                        logging.debug(f"OLE file detected: {subfile}")
                        ole_data = z.open(subfile).read()
                        vbaparser = VBA_Parser(filename=subfile, data=ole_data)
                        has_macros = vbaparser.detect_macros()
                        vbaparser.close()
                        if has_macros:
                            return True
                except Exception as e:
                    logging.debug(f"Failed to read {subfile} inside ZIP: {e}")
                    continue
    except Exception as e:
        logging.debug(f"Error opening ZIP file: {e}")

    return False

def main():
    parser = argparse.ArgumentParser(description='Check for macro presence in a single Office file')
    parser.add_argument('filepath', help='Path to Office file')
    args = parser.parse_args()
    path = args.filepath

    has_macros = (
        run_olevba_triage(path) or
        run_oledump_check(path) or
        check_excel97_macros(path) or
        run_vba_parser(path) or
        scan_zip_for_macros(path)
    )

    if has_macros:
        print(f"MACROS_PRESENT")
    else:
        print(f"NO_MACROS")

if __name__ == '__main__':
    logging.basicConfig(level=logging.DEBUG)  # Optional: enable debug output
    main()