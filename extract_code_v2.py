from pathlib import Path
from collections import Counter
from datetime import datetime
import re
import xml.etree.ElementTree as ET
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# ==========================================================
# CONFIGURATION
# ==========================================================

BASE_FOLDER = Path(
    r"C:\Users\cdourado\Downloads\Py - codes\FIDC_DASH_PROCESS"
)

# Folder containing the XML files.
#
# If the XML files are directly inside BASE_FOLDER,
# leave this exactly as it is.
XML_FOLDER = Path(r"\\SPNETAPP01\Data\DADOS\RISCO\Credito\FIDC\Montlhy XMLs")

# File containing the selected XML paths.
SELECTED_PATHS_FILE = BASE_FOLDER / "selected_paths.txt"

# Excel file that will be created.
OUTPUT_EXCEL = BASE_FOLDER / "xml_selected_wide_data.xlsx"

# Workbook used to convert XML paths to _ headers.
HEADERS_FIXED_FILE = BASE_FOLDER / "Headers_fixed.xlsx"

# True: reads XML files in BASE_FOLDER and all subfolders.
# False: reads XML files only directly inside BASE_FOLDER.
READ_SUBFOLDERS = True

# Include XML attributes only if a selected path ends with /@attribute.
INCLUDE_ATTRIBUTES = True

# Maximum number of series/class slots to include in the output.
# Increase these if new XML files contain more slots than expected.
MAX_SENIOR_SLOTS = 30   # Senior: Base + [1]..[30]
MAX_SUBORD_SLOTS = 25   # Subordinada: Base + [1]..[25]


# ==========================================================
# GENERAL FUNCTIONS
# ==========================================================

def remove_namespace(name):
    """
    Removes the namespace from an XML tag or attribute.

    Example:
    {urn:fidc}DOC_ARQ

    Becomes:
    DOC_ARQ
    """
    if name is None:
        return ""

    if "}" in name:
        return name.split("}", 1)[1]

    return name


def clean_value(value):
    """
    Removes extra spaces, tabs and line breaks.

    Returns None if the value is empty.
    """
    if value is None:
        return None

    cleaned = " ".join(str(value).split())

    if cleaned == "":
        return None

    return cleaned


# ==========================================================
# READ SELECTED PATHS
# ==========================================================

def load_selected_paths(file_path):
    """
    Reads paths from selected_paths.txt.

    This function accepts both formats:

    Format 1, one path per line:

    /DOC_ARQ/CAB_INFORM/CLASS_UNICA
    /DOC_ARQ/CAB_INFORM/DT_COMPT

    Format 2, concatenated paths:

    /DOC_ARQ/CAB_INFORM/CLASS_UNICA/DOC_ARQ/CAB_INFORM/DT_COMPT

    Every occurrence of /DOC_ARQ starts a new path.
    """
    if not file_path.exists():
        raise FileNotFoundError(
            f"The selected-path file was not found:\n"
            f"{file_path}\n\n"
            "Create selected_paths.txt and paste the selected paths into it."
        )

    text = file_path.read_text(encoding="utf-8-sig")

    # Remove accidental HTML encoding copied from Teams, Outlook or web pages.
    text = text.replace("&gt;", ">")
    text = text.replace("&lt;", "<")
    text = text.replace("&amp;", "&")
    text = text.replace("\u200b", "")
    text = text.strip()

    if not text:
        raise ValueError(
            f"The selected-path file is empty:\n{file_path}"
        )

    # Finds every path starting with /DOC_ARQ.
    #
    # The path ends before:
    # - the next /DOC_ARQ
    # - a line break
    # - the end of the file
    matches = re.findall(
        r"/DOC_ARQ.*?(?=/DOC_ARQ|\r?\n|$)",
        text
    )

    selected_paths = []
    already_added = set()

    for path in matches:
        path = path.strip()

        # Remove accidental separators.
        path = path.rstrip(" ,;")

        if not path:
            continue

        # Remove duplicate paths while maintaining the original order.
        if path not in already_added:
            selected_paths.append(path)
            already_added.add(path)

    if not selected_paths:
        raise ValueError(
            "No valid XML paths were found in selected_paths.txt.\n"
            "Every selected path must start with /DOC_ARQ."
        )

    return selected_paths


SELECTED_PATHS = load_selected_paths(SELECTED_PATHS_FILE)


# ==========================================================
# COMPLETE MISSING COTA SLOTS AND REORDER
# ==========================================================

COTA_RE = re.compile(
    r"DESC_SERIE_CLASSE_(SENIOR|SUBORD)"
    r"(?:\[(\d+)\])?"
    r"/(SERIE|TIPO|QT_COTAS|VL_COTAS)$"
)

SENIOR_PREFIX = (
    "/DOC_ARQ/LISTA_INFORM/OUTRAS_INFORM"
    "/DESC_SERIE_CLASSE/DESC_SERIE_CLASSE_SENIOR"
)
SUBORD_PREFIX = (
    "/DOC_ARQ/LISTA_INFORM/OUTRAS_INFORM"
    "/DESC_SERIE_CLASSE/DESC_SERIE_CLASSE_SUBORD"
)

SENIOR_FIELDS = ["SERIE", "QT_COTAS", "VL_COTAS"]
SUBORD_FIELDS = ["SERIE", "TIPO", "QT_COTAS", "VL_COTAS"]

# Build the full set of expected cota paths.
expected_cota_paths = set()

# Senior Base (no index).
for field in SENIOR_FIELDS:
    expected_cota_paths.add(f"{SENIOR_PREFIX}/{field}")

# Senior indexed [1] .. [MAX_SENIOR_SLOTS].
for i in range(1, MAX_SENIOR_SLOTS + 1):
    for field in SENIOR_FIELDS:
        expected_cota_paths.add(f"{SENIOR_PREFIX}[{i}]/{field}")

# Subord Base (no index) — Base has no SERIE.
for field in ["TIPO", "QT_COTAS", "VL_COTAS"]:
    expected_cota_paths.add(f"{SUBORD_PREFIX}/{field}")

# Subord indexed [1] .. [MAX_SUBORD_SLOTS].
for i in range(1, MAX_SUBORD_SLOTS + 1):
    for field in SUBORD_FIELDS:
        expected_cota_paths.add(f"{SUBORD_PREFIX}[{i}]/{field}")

# Add any missing cota paths to the list.
existing_set = set(SELECTED_PATHS)
missing_paths = sorted(expected_cota_paths - existing_set)

if missing_paths:
    # Insert after the last existing cota path.
    last_cota_pos = max(
        (i for i, p in enumerate(SELECTED_PATHS) if COTA_RE.search(p)),
        default=len(SELECTED_PATHS) - 1
    )
    for offset, path in enumerate(missing_paths, start=1):
        SELECTED_PATHS.insert(last_cota_pos + offset, path)
    print(f"Added {len(missing_paths)} missing cota paths")

# Reorder only the cota block; all other columns stay in place.
FIELD_ORDER = {"SERIE": 0, "TIPO": 1, "QT_COTAS": 2, "VL_COTAS": 3}


def _cota_sort_key(path):
    m = COTA_RE.search(path)
    kind, idx, field = m.groups()
    return (
        kind != "SENIOR",                    # Senior first
        0 if idx is None else int(idx),      # Base = 0, then 1, 2, 3 ...
        FIELD_ORDER[field],                  # SERIE, TIPO, QT, VL
    )


cota_positions = [
    i for i, p in enumerate(SELECTED_PATHS)
    if COTA_RE.search(p)
]

cota_paths_sorted = sorted(
    (SELECTED_PATHS[i] for i in cota_positions),
    key=_cota_sort_key,
)

for pos, path in zip(cota_positions, cota_paths_sorted):
    SELECTED_PATHS[pos] = path

print(f"Cota columns reordered: {len(cota_positions)} paths")


SELECTED_PATHS_SET = set(SELECTED_PATHS)

print(f"Selected columns loaded: {len(SELECTED_PATHS)}")


# ==========================================================
# COTA TOTALS CALCULATION
# ==========================================================

TOTAL_COLUMNS = ["Total_Senior", "Total_Mezanino", "Total_Subordinada"]


def _to_float(value):
    """Converts a string value to float.  Returns None on failure."""
    if value is None:
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def calc_totals(values):
    """
    Calculates Total_Senior, Total_Mezanino and Total_Subordinada
    from one XML record's extracted values.

    Total = sum(QT_COTAS * VL_COTAS) per class type.

    Senior slots are always classified as Senior (no TIPO field).

    Subordinada slots are classified by the TIPO field:
    - TIPO starts with "Mezanino"  ->  Total_Mezanino
    - Anything else (including empty)  ->  Total_Subordinada
    """
    # Group fields by (kind, index).
    slots = {}

    for path, value in values.items():
        m = COTA_RE.search(path)
        if not m:
            continue

        kind, idx, field = m.groups()
        slot_key = (kind, idx)

        if slot_key not in slots:
            slots[slot_key] = {}

        slots[slot_key][field] = value

    totals = {name: 0.0 for name in TOTAL_COLUMNS}

    for (kind, _idx), fields in slots.items():
        qt = _to_float(fields.get("QT_COTAS"))
        vl = _to_float(fields.get("VL_COTAS"))

        if qt is None or vl is None:
            continue

        product = qt * vl

        if kind == "SENIOR":
            totals["Total_Senior"] += product
        else:
            tipo = str(fields.get("TIPO", "")).strip().lower()

            if tipo.startswith("mezanino"):
                totals["Total_Mezanino"] += product
            else:
                totals["Total_Subordinada"] += product

    return totals


# ==========================================================
# XML EXTRACTION
# ==========================================================

def add_selected_value(result, xml_path, value):
    """
    Adds a value only when the XML path is selected.

    If the same exact path produces multiple values,
    all distinct values are preserved in the same cell.
    """
    if xml_path not in SELECTED_PATHS_SET:
        return

    value = clean_value(value)

    if value is None:
        return

    if xml_path not in result:
        result[xml_path] = value
        return

    existing_value = str(result[xml_path])
    existing_values = existing_value.split(" | ")

    if value not in existing_values:
        result[xml_path] = f"{existing_value} | {value}"


def extract_selected_values(xml_file):
    """
    Reads one XML file and returns only the selected paths.

    Repeated sibling elements receive indexes:

    /DESC_SERIE_CLASSE_SENIOR[1]/SERIE
    /DESC_SERIE_CLASSE_SENIOR[2]/SERIE
    /DESC_SERIE_CLASSE_SENIOR[3]/SERIE

    If an element occurs only once, it does not receive an index:

    /DESC_SERIE_CLASSE_SENIOR/SERIE
    """
    tree = ET.parse(xml_file)
    root = tree.getroot()

    extracted_values = {}

    def walk(element, element_path):
        # --------------------------------------------------
        # XML attributes
        # --------------------------------------------------

        if INCLUDE_ATTRIBUTES:
            for attribute_name, attribute_value in element.attrib.items():
                clean_attribute_name = remove_namespace(attribute_name)

                attribute_path = (
                    f"{element_path}/@{clean_attribute_name}"
                )

                add_selected_value(
                    extracted_values,
                    attribute_path,
                    attribute_value
                )

        children = list(element)

        # --------------------------------------------------
        # Leaf element
        # --------------------------------------------------

        if not children:
            add_selected_value(
                extracted_values,
                element_path,
                element.text
            )
            return

        # --------------------------------------------------
        # Optional text inside a parent element
        # --------------------------------------------------

        parent_text = clean_value(element.text)

        if parent_text is not None:
            add_selected_value(
                extracted_values,
                f"{element_path}/#text",
                parent_text
            )

        # Count siblings with the same tag name.
        child_tag_counts = Counter(
            remove_namespace(child.tag)
            for child in children
        )

        # Track the position of every repeated sibling.
        child_tag_positions = Counter()

        for child in children:
            child_tag = remove_namespace(child.tag)
            child_tag_positions[child_tag] += 1

            if child_tag_counts[child_tag] > 1:
                child_component = (
                    f"{child_tag}"
                    f"[{child_tag_positions[child_tag]}]"
                )
            else:
                child_component = child_tag

            child_path = f"{element_path}/{child_component}"

            walk(child, child_path)

    root_tag = remove_namespace(root.tag)
    root_path = f"/{root_tag}"

    walk(root, root_path)

    return extracted_values


# ==========================================================
# FIND XML FILES
# ==========================================================

if not XML_FOLDER.exists():
    raise FileNotFoundError(
        f"The XML folder was not found:\n{XML_FOLDER}"
    )

if READ_SUBFOLDERS:
    xml_files = sorted(
        XML_FOLDER.rglob("*.xml"),
        key=lambda file: str(file).lower()
    )
else:
    xml_files = sorted(
        XML_FOLDER.glob("*.xml"),
        key=lambda file: str(file).lower()
    )

if not xml_files:
    raise FileNotFoundError(
        f"No XML files were found in:\n{XML_FOLDER}"
    )

print(f"XML files found: {len(xml_files)}")


# ==========================================================
# PROCESS XML FILES
# ==========================================================

all_records = []

for file_number, xml_file in enumerate(xml_files, start=1):
    print(
        f"Reading {file_number}/{len(xml_files)}: "
        f"{xml_file.name}"
    )

    try:
        xml_values = extract_selected_values(xml_file)

        all_records.append({
            "SourceFile": xml_file.name,
            "SourceFolder": str(xml_file.parent),
            "ParseStatus": "OK",
            "ErrorMessage": None,
            "Values": xml_values
        })

    except ET.ParseError as error:
        all_records.append({
            "SourceFile": xml_file.name,
            "SourceFolder": str(xml_file.parent),
            "ParseStatus": "XML ERROR",
            "ErrorMessage": str(error),
            "Values": {}
        })

    except Exception as error:
        all_records.append({
            "SourceFile": xml_file.name,
            "SourceFolder": str(xml_file.parent),
            "ParseStatus": "ERROR",
            "ErrorMessage": str(error),
            "Values": {}
        })


# ==========================================================
# EXCEL COLUMN LIMIT
# ==========================================================

# Four control columns + selected XML columns + three total columns.
total_columns = 4 + len(SELECTED_PATHS) + len(TOTAL_COLUMNS)

if total_columns > 16384:
    raise ValueError(
        f"The output would contain {total_columns} columns.\n"
        "Excel supports a maximum of 16,384 columns."
    )


# ==========================================================
# CREATE THE EXCEL FILE
# ==========================================================

workbook = Workbook()

worksheet = workbook.active
worksheet.title = "XML Wide Data"

# Hide worksheet gridlines.
worksheet.sheet_view.showGridLines = False


# ==========================================================
# EXCEL STYLES
# ==========================================================

control_header_fill = PatternFill(
    fill_type="solid",
    fgColor="7030A0"
)

path_header_fill = PatternFill(
    fill_type="solid",
    fgColor="70AD47"
)

total_header_fill = PatternFill(
    fill_type="solid",
    fgColor="ED7D31"
)

error_fill = PatternFill(
    fill_type="solid",
    fgColor="F4CCCC"
)

header_font = Font(
    color="FFFFFF",
    bold=True
)

imported_font = Font(
    color="008000"
)


# ==========================================================
# WRITE EXCEL HEADERS
# ==========================================================

control_columns = [
    "SourceFile",
    "SourceFolder",
    "ParseStatus",
    "ErrorMessage"
]

for column_number, header in enumerate(
    control_columns,
    start=1
):
    cell = worksheet.cell(
        row=1,
        column=column_number,
        value=header
    )

    cell.fill = control_header_fill
    cell.font = header_font
    cell.alignment = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True
    )


# Selected XML paths start in Excel column E.
path_to_column = {}

for column_number, xml_path in enumerate(
    SELECTED_PATHS,
    start=5
):
    path_to_column[xml_path] = column_number

    cell = worksheet.cell(
        row=1,
        column=column_number,
        value=xml_path
    )

    cell.fill = path_header_fill
    cell.font = header_font
    cell.alignment = Alignment(
        horizontal="left",
        vertical="center",
        wrap_text=True
    )


# Total columns start after the last selected path column.
first_total_column = 5 + len(SELECTED_PATHS)

for offset, header_name in enumerate(TOTAL_COLUMNS):
    cell = worksheet.cell(
        row=1,
        column=first_total_column + offset,
        value=header_name
    )

    cell.fill = total_header_fill
    cell.font = header_font
    cell.alignment = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True
    )


# ==========================================================
# WRITE ONE EXCEL ROW PER XML
# ==========================================================

for row_number, record in enumerate(
    all_records,
    start=2
):
    worksheet.cell(
        row=row_number,
        column=1,
        value=record["SourceFile"]
    )

    worksheet.cell(
        row=row_number,
        column=2,
        value=record["SourceFolder"]
    )

    status_cell = worksheet.cell(
        row=row_number,
        column=3,
        value=record["ParseStatus"]
    )

    worksheet.cell(
        row=row_number,
        column=4,
        value=record["ErrorMessage"]
    )

    if record["ParseStatus"] != "OK":
        status_cell.fill = error_fill

    for xml_path, value in record["Values"].items():
        column_number = path_to_column.get(xml_path)

        if column_number is None:
            continue

        value_cell = worksheet.cell(
            row=row_number,
            column=column_number,
            value=value
        )

        value_cell.font = imported_font

    # Any selected path not found in this XML is not written.
    # Consequently, the corresponding Excel cell remains blank.

    # Write cota totals for successfully parsed files.
    if record["ParseStatus"] == "OK":
        totals = calc_totals(record["Values"])

        for offset, name in enumerate(TOTAL_COLUMNS):
            total_cell = worksheet.cell(
                row=row_number,
                column=first_total_column + offset,
                value=totals[name],
            )

            total_cell.font = imported_font
            total_cell.number_format = "#,##0.00"


# ==========================================================
# EXCEL FORMATTING
# ==========================================================

worksheet.auto_filter.ref = worksheet.dimensions
worksheet.row_dimensions[1].height = 90

worksheet.column_dimensions["A"].width = 34
worksheet.column_dimensions["B"].width = 45
worksheet.column_dimensions["C"].width = 15
worksheet.column_dimensions["D"].width = 45

for column_number in range(5, worksheet.max_column + 1):
    column_letter = get_column_letter(column_number)
    worksheet.column_dimensions[column_letter].width = 22



# ==========================================================
# REPLACE HEADERS DIRECTLY USING HEADERS_FIXED.XLSX
# ==========================================================
# The Python code performs the lookup directly.
#
# No Excel formulas are created.
# No auxiliary worksheet is created.
# No rows are inserted or hidden.
# The original headers are replaced directly in row 1.

if not HEADERS_FIXED_FILE.exists():
    raise FileNotFoundError(
        f"The header mapping file was not found:\n"
        f"{HEADERS_FIXED_FILE}"
    )

from openpyxl import load_workbook


# ----------------------------------------------------------
# LOAD HEADER MAPPINGS
# ----------------------------------------------------------

headers_workbook = load_workbook(
    HEADERS_FIXED_FILE,
    read_only=True,
    data_only=True
)

headers_worksheet = headers_workbook.active

# Dictionary structure:
# {
#     "original XML path": "h header",
#     ...
# }
header_mapping = {}

for source_field, header_name in headers_worksheet.iter_rows(
    min_row=2,
    min_col=1,
    max_col=2,
    values_only=True
):
    if source_field is None:
        continue

    source_field = str(source_field).strip()

    if not source_field:
        continue

    # If the header name is empty, keep the original header.
    if header_name is None:
        header_name = source_field
    else:
        header_name = str(header_name).strip()

        if not header_name:
            header_name = source_field

    header_mapping[source_field] = header_name

headers_workbook.close()


if not header_mapping:
    raise ValueError(
        "Headers_fixed.xlsx does not contain valid mappings "
        "in columns A and B."
    )


# ----------------------------------------------------------
# AUTO-GENERATE HEADERS FOR NEW COTA PATHS
# ----------------------------------------------------------
# Any cota paths added by the slot-completion step above
# may not exist in Headers_fixed.xlsx yet.
# Generate their friendly names using the same pattern.

for path in SELECTED_PATHS:
    if path in header_mapping:
        continue

    m = COTA_RE.search(path)
    if not m:
        continue

    kind, idx, field = m.groups()
    prefix = "Senior_Serie" if kind == "SENIOR" else "Subordinada_Classe"
    index_str = "Base" if idx is None else idx
    header_mapping[path] = f"{prefix}_{index_str}_{field}"


# ----------------------------------------------------------
# REPLACE THE HEADERS DIRECTLY IN ROW 1
# ----------------------------------------------------------

headers_replaced = 0
headers_not_mapped = []

for column_number in range(1, worksheet.max_column + 1):
    header_cell = worksheet.cell(
        row=1,
        column=column_number
    )

    original_header = header_cell.value

    if original_header is None:
        continue

    original_header = str(original_header).strip()

    header_header = header_mapping.get(original_header)

    if header_header is not None:
        # Replace the original header directly.
        header_cell.value = header_header
        headers_replaced += 1

    else:
        # No mapping was found, so keep the original header.
        header_cell.value = original_header
        headers_not_mapped.append(original_header)


# ----------------------------------------------------------
# REAPPLY HEADER FORMATTING
# ----------------------------------------------------------

for column_number in range(1, worksheet.max_column + 1):
    header_cell = worksheet.cell(
        row=1,
        column=column_number
    )

    header_cell.font = header_font

    if column_number <= 4:
        header_cell.fill = control_header_fill
        header_cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )
    elif column_number >= first_total_column:
        header_cell.fill = total_header_fill
        header_cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )
    else:
        header_cell.fill = path_header_fill
        header_cell.alignment = Alignment(
            horizontal="left",
            vertical="center",
            wrap_text=True
        )


# ----------------------------------------------------------
# FINAL WORKSHEET SETTINGS
# ----------------------------------------------------------
# There is only one header row, so the first data row is row 2.

worksheet.auto_filter.ref = worksheet.dimensions
worksheet.row_dimensions[1].height = 90

# ==========================================================
# SAVE EXCEL - DELETE OLD FILE AND CREATE A NEW ONE
# ==========================================================

try:
    if OUTPUT_EXCEL.exists():
        OUTPUT_EXCEL.unlink()

    workbook.save(OUTPUT_EXCEL)

except PermissionError:
    raise PermissionError(
        f"Could not replace the output file:\n"
        f"{OUTPUT_EXCEL}\n\n"
        "The file is probably open in Excel. "
        "Close it and run the script again."
    )

# ==========================================================
# SUMMARY
# ==========================================================

successful_files = sum(
    record["ParseStatus"] == "OK"
    for record in all_records
)

files_with_errors = len(all_records) - successful_files

paths_found = {
    xml_path
    for record in all_records
    for xml_path in record["Values"].keys()
}

paths_not_found = [
    xml_path
    for xml_path in SELECTED_PATHS
    if xml_path not in paths_found
]

print()
print("Processing completed.")
print(f"XML files processed: {len(all_records)}")
print(f"Successfully processed: {successful_files}")
print(f"Files with errors: {files_with_errors}")
print(f"Selected columns: {len(SELECTED_PATHS)}")
print(f"Selected paths found: {len(paths_found)}")
print(f"Selected paths not found: {len(paths_not_found)}")
print(f"Output file: {OUTPUT_EXCEL}")

if paths_not_found:
    print()
    print("Selected paths not found in any XML:")

    for xml_path in paths_not_found:
        print(f"  {xml_path}")
