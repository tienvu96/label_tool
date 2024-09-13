import pandas as pd
import json
import io
import csv
import re
from docx import Document
from docx.text.paragraph import Paragraph
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.shared import Pt
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.drawing.image import Image


def is_file_locked(file_path):
        """Check if the file is currently locked (opened by another application)."""
        try:
            # Try opening the file in append mode (or any mode) to check for a lock
            with open(file_path, 'a'):
                pass
            return False  # File is not locked
        except IOError:
            return True  # File is locked




# Function to read the JSON file and load the keywords
def load_keywords_from_json():
    """Load keywords from a JSON file."""
    json_file = '/Users/vuanhtien/Documents/CodeForMe/BIDV/tool_label20240827/keyword.json'
    with open(json_file, 'r', encoding='utf-8') as file:
        data = json.load(file)
    return data


# Function to extract content from a DOCX file
def extract_and_iterate_docx_content(file_path, table_id=None, **pandas_kwargs):
    """Extracts content from a DOCX file, including paragraphs and tables."""
    document = Document(file_path)


    def extract_single_table(table):
        memory_file = io.StringIO()
        csv_writer = csv.writer(memory_file)
        for row in table.rows:
            csv_writer.writerow(cell.text for cell in row.cells)
        memory_file.seek(0)
        return pd.read_csv(memory_file, **pandas_kwargs)


    content_list = []
    for child_element in document.element.body.iterchildren():
        if isinstance(child_element, CT_P):
            content_list.append(Paragraph(child_element, document).text)
        elif isinstance(child_element, CT_Tbl):
            table = Table(child_element, document)
            if table_id is None or document.tables.index(table) == table_id:
                content_list.append(extract_single_table(table))
    return content_list


# Function to search for keywords within the DOCX content
def find_keywords_in_docx(docx_content, keywords_dict):
    """Search for keywords in the DOCX content without returning categories."""
    found_keywords = []
    docx_content_lower = docx_content.lower()


    for keywords in keywords_dict.values():
        for keyword_list in keywords.values():
            for keyword in keyword_list:
                if keyword.lower() in docx_content_lower:
                    found_keywords.append({"Found Keyword": keyword})


    return found_keywords


# Function to search for regex patterns within the DOCX content
def find_patterns_in_docx(docx_content, patterns):
    """Search for regex patterns in the DOCX content."""
    found_patterns = []
    for pattern_name, pattern_regex in patterns.items():
        matches = pattern_regex.findall(docx_content)
        for match in matches:
            found_patterns.append({"Pattern Name": pattern_name, "Matched Text": match})
    return found_patterns


# Main function to check keywords and patterns in a DOCX file
def check_keywords_and_patterns_in_docx(docx_file, patterns):
    """Check a DOCX file for both keywords and patterns without returning categories."""
    keywords_dict = load_keywords_from_json()  # Load keywords
    docx_content_list = extract_and_iterate_docx_content(docx_file)
    docx_content = " ".join([str(content) for content in docx_content_list])


    found_keywords = find_keywords_in_docx(docx_content, keywords_dict)
    found_patterns = find_patterns_in_docx(docx_content, patterns)


    return {"Keywords": found_keywords, "Patterns": found_patterns}


# Function to extract all sheets from an Excel file
def extract_and_iterate_excel_content(file_path, **pandas_kwargs):
    """Extracts content from all sheets of an Excel file."""
    sheet_dict = pd.read_excel(file_path, sheet_name=None, **pandas_kwargs)
    content_dict = {}


    for sheet_name, df in sheet_dict.items():
        content_list = []
        for _, row in df.iterrows():
            for cell in row:
                if pd.notnull(cell):
                    content_list.append(str(cell))
        content_dict[sheet_name] = content_list


    return content_dict


# Function to search for keywords in Excel content
def find_keywords_in_excel(excel_content_dict, keywords_dict):
    """Search for keywords in Excel content without returning categories."""
    found_keywords = []


    for sheet_name, excel_content in excel_content_dict.items():
        for keywords in keywords_dict.values():
            for keyword_list in keywords.values():
                for keyword in keyword_list:
                    if any(keyword.lower() in str(content).lower() for content in excel_content):
                        found_keywords.append({"Sheet": sheet_name, "Found Keyword": keyword})


    return found_keywords


# Function to search for patterns in Excel content
def find_patterns_in_excel(excel_content_dict, patterns):
    """Search for patterns in Excel content."""
    found_patterns = []
    for sheet_name, excel_content in excel_content_dict.items():
        for pattern_name, pattern_regex in patterns.items():
            for content in excel_content:
                matches = pattern_regex.findall(str(content))
                for match in matches:
                    found_patterns.append({"Sheet": sheet_name, "Pattern Name": pattern_name, "Matched Text": match})
    return found_patterns


# Main function to check keywords and patterns in an Excel file
def check_keywords_and_patterns_in_excel(xlsx_file, patterns):
    """Check an Excel file for both keywords and patterns without returning categories."""
    keywords_dict = load_keywords_from_json()
    excel_content_dict = extract_and_iterate_excel_content(xlsx_file)


    found_keywords = find_keywords_in_excel(excel_content_dict, keywords_dict)
    found_patterns = find_patterns_in_excel(excel_content_dict, patterns)


    return {"Keywords": found_keywords, "Patterns": found_patterns}


# Function to extract content from a CSV file
def extract_and_iterate_csv_content(file_path, encoding='utf-8'):
    """Extracts content from a CSV file."""
    content_list = []
    try:
        with open(file_path, newline='', encoding=encoding) as csvfile:
            reader = csv.reader(csvfile)
            for row in reader:
                for cell in row:
                    if cell:
                        content_list.append(cell)
    except UnicodeDecodeError:
        with open(file_path, newline='', encoding='ISO-8859-1') as csvfile:
            reader = csv.reader(csvfile)
            for row in reader:
                for cell in row:
                    if cell:
                        content_list.append(cell)
    return content_list


# Function to search for keywords in CSV content
def find_keywords_in_csv(csv_content, keywords_dict):
    """Search for keywords in CSV content without returning categories."""
    found_keywords = []


    for keywords in keywords_dict.values():
        for keyword_list in keywords.values():
            for keyword in keyword_list:
                if any(keyword.lower() in content.lower() for content in csv_content):
                    found_keywords.append({"Found Keyword": keyword})


    return found_keywords


# Function to search for patterns in CSV content
def find_patterns_in_csv(csv_content, patterns):
    """Search for patterns in CSV content."""
    found_patterns = []
    for pattern_name, pattern_regex in patterns.items():
        for content in csv_content:
            matches = pattern_regex.findall(content)
            for match in matches:
                found_patterns.append({"Pattern Name": pattern_name, "Matched Text": match})
    return found_patterns


# Main function to check keywords and patterns in a CSV file
def check_keywords_and_patterns_in_csv(csv_file, patterns):
    """Check a CSV file for both keywords and patterns without returning categories."""
    keywords_dict = load_keywords_from_json()
    csv_content_list = extract_and_iterate_csv_content(csv_file)


    found_keywords = find_keywords_in_csv(csv_content_list, keywords_dict)
    found_patterns = find_patterns_in_csv(csv_content_list, patterns)


    return {"Keywords": found_keywords, "Patterns": found_patterns}


# Function to scan files and check keywords/patterns based on file extension
def scan_file(file_path):
    """Scan a file and check for both keywords and patterns depending on the file type."""
    if not file_path:
        return None, "Vui lòng chọn một file trước khi scan"


    file_extension = Path(file_path).suffix.lower()
   
    patterns = {
        "stk 10 số chứa 3 BDS": re.compile(r'\b(111|116|117|118|119|120|121|122|123|124|125|126|128|129|130|131|132|133|134|135|136|138|139|140|141|144|145|147|149|150|151|159|160|166|168|169|177|180|181|186|188|189|199|211|212|213|214|215|216|217|220|222|256|260|261|268|279|289|310|311|313|314|315|317|318|319|321|328|330|341|345|351|362|368|371|375|376|390|395|398|411|421|425|426|427|428|431|432|433|440|441|443|444|448|450|451|452|455|460|461|465|466|468|471|480|482|483|486|488|501|502|505|512|513|518|520|522|531|532|540|556|558|560|561|562|566|565|570|573|580|581|590|601|602|611|615|620|621|625|631|632|633|635|636|641|642|646|650|651|652|653|655|656|661|670|671|672|679|680|686|691|696|701|702|710|711|721|729|730|735|737|741|742|745|748|750|753|760|761|762|766|780|785|788)\d{7}\b'),
        "cvv": re.compile(r"(?i)cvv:\s?\d{3}"),
        "id number (cccd/cmnd)": re.compile(r'\b(?:cmnd|cccd|CMND|CCCD)\b.*\b\d{9}\b|\b(?:cmnd|cccd|CMND|CCCD)\b.*\b\d{12}\b'),
        "địa chỉ": re.compile(r'(\d+)\s+(đường|phố|phường|quận|thành phố)\s+(\w+),\s+(\w+),\s+(\w+)'),
        "giá trị tiền tối thiểu 6 chữ số": re.compile(r'\b\d{1,3}(,\d{3}){1,}\b'),
        "số điện thoại": re.compile(r'\b0\d{9}\b'),
        "email": re.compile(r'\S+@\S+'),
        "số thẻ": re.compile(r'\b(?:9704|476632|411153|428695|427126|402460|406220|511957|517107|517453|542726|530515|515110)(\d{12}|\d{15}|\d{3} \d{4} \d{4} \d{4} \d{4}|\d{4} \d{4} \d{4} \d{4} \d{3}|\d{4} \d{4} \d{4} \d{4} \d{4})\b'),
        "số tài khoản 14 số": re.compile(r'^(9704\d{12,15}|476632\d{10,13}|411153\d{10,13}|428695\d{10,13}|427126\d{10,13}|402460\d{10,13}|406220\d{10,13}|511957\d{10,13}|517107\d{10,13}|517453\d{10,13}|542726\d{10,13}|530515\d{10,13}|51511\d{11,14}|9704 \d{4} \d{4} \d{4}( \d{3})?|476632 \d{4} \d{4} \d{4}( \d{3})?|411153 \d{4} \d{4} \d{4}( \d{3})?|428695 \d{4} \d{4} \d{4}( \d{3})?|427126 \d{4} \d{4} \d{4}( \d{3})?|402460 \d{4} \d{4} \d{4}( \d{3})?|406220 \d{4} \d{4} \d{4}( \d{3})?|511957 \d{4} \d{4} \d{4}( \d{3})?|517107 \d{4} \d{4} \d{4}( \d{3})?|517453 \d{4} \d{4} \d{4}( \d{3})?|542726 \d{4} \d{4} \d{4}( \d{3})?|530515 \d{4} \d{4} \d{4}( \d{3})?|51511 \d{4} \d{4} \d{4}( \d{3})?)$'),
        "số thẻ 16 hoặc 19 số": re.compile(r'\b(?:9704|476632|411153|428695|427126|402460|406220|511957|517107|517453|542726|530515|515110)(\d{12}|\d{15}|\d{3} \d{4} \d{4} \d{4} \d{4}|\d{4} \d{4} \d{4} \d{4} \d{3}|\d{4} \d{4} \d{4} \d{4} \d{4})\b'),
        "số tài khoản ẩn": re.compile(r'\b(?:9704|476632|411153|428695|427126|402460|406220|511957|517107|517453|542726|530515|515110)(?:\d{2}xxxx\d{4}|\d{3}xxxx\d{4}|\d{4}xxxx\d{4})\b')
    }


    if file_extension == '.docx':
        results = check_keywords_and_patterns_in_docx(file_path, patterns)
    elif file_extension == '.xlsx':
        results = check_keywords_and_patterns_in_excel(file_path, patterns)
    elif file_extension == '.csv':
        results = check_keywords_and_patterns_in_csv(file_path, patterns)
    else:
        return None, "Không hỗ trợ định dạng tệp."


    if not results.get('Keywords') and not results.get('Patterns'):
        message = "Không tìm thấy bất kỳ thông tin nội bộ hay bí mật nào trong tài liệu."
    else:
        message = json.dumps(results, indent=4, ensure_ascii=False)


    return results, message


def define_rules():
    rules = {
            "rule_1": {
                "số thẻ": "số thẻ",
                "cvv": "cvv",
                "ngày hết hạn": "ngày hết hạn",
                "tên chủ thẻ": ["tên khách hàng", "chủ thẻ", "tên chủ thẻ", "họ và tên", "họ tên", "name", "full name"]
            },
            "rule_2": {
                "email": "email",
                "id number (cccd/cmnd)": "id number (cccd/cmnd)",
                "tên khách hàng": ["tên khách hàng", "họ tên", "họ và tên", "họ tên", "name", "full name"],
                "địa chỉ": "địa chỉ",
                "số điện thoại": "số điện thoại"
            },
            "rule_3": {
                "email": "email",
                "số điện thoại": "số điện thoại",
                "số tài khoản ẩn": "số tài khoản ẩn",
                "giá trị tiền tối thiểu 6 chữ số": "giá trị tiền tối thiểu 6 chữ số"
            }
        }
   
    return rules




def classify_document_with_multiple_rules(results, rules):
    """
    Classify the document based on the comparison between scan results and multiple rules,
    applying different classification criteria for each rule. Additionally, print the matched keys.


    Args:
        results (dict): The results from the scan_file function.
        rules (dict): Dictionary containing multiple rules with different classification criteria.


    Returns:
        str: label_text ("Confidential" if the document meets the criteria for any rule's classification,
             otherwise "Internal"), and prints out the matched keys.
    """
    found_keywords = results.get('Keywords', [])
    found_patterns = results.get('Patterns', [])
    category = "Internal"


    # Iterate over each rule in the rules dictionary
    for rule_name, rule_dict in rules.items():
        matched_keys = set()


        # Check found keywords
        for keyword_info in found_keywords:
            found_keyword = keyword_info.get("Found Keyword")
            if found_keyword in rule_dict:
                matched_keys.add(found_keyword)


        # Check found patterns
        for pattern_info in found_patterns:
            found_pattern_name = pattern_info.get("Pattern Name")
            if found_pattern_name in rule_dict:
                matched_keys.add(found_pattern_name)
        # Print matched keys for debugging
        show_rule = f"Vi phạm {rule_name}: {matched_keys}"
        print("vi phạm", show_rule)




        # Rule-specific conditions:
        if rule_name == "rule_1":
            # Rule 1: If at least one match is found, classify as "Confidential"
            if len(matched_keys) == 4:
                category = "Confidential"
                break
       
        elif rule_name == "rule_2":
            # Rule 2: If at least three distinct matches are found, classify as "Confidential"
            if len(matched_keys) >= 3:
                category = "Confidential"
                break
       
        elif rule_name == "rule_3":
            # Rule 3: If at least two distinct matches are found, classify as "Confidential"
            if len(matched_keys) >= 5:
                category = "Confidential"
                break
       
    # Return the label_text (either "Confidential" or "Internal")
    return category, show_rule




# Function to label DOCX file
def label_docx_file(file_path, category):
    """Add or overwrite a label in the footer of a DOCX file based on the scan results."""
    document = Document(file_path)


    # Determine the label text based on the category
    if category == "Confidential":
        label_text = "Tài liệu Mật của BIDV. Cấm sao chép, in ấn dưới mọi hình thức!"
    else:
        label_text = "Tài liệu Nội bộ của BIDV. Cấm sao chép, in ấn dưới mọi hình thức!"


    section = document.sections[0]
    footer = section.footer


    # Clear any existing content in the footer
    if footer.paragraphs:
        for paragraph in footer.paragraphs:
            paragraph.clear()


    # Add the new label text
    paragraph = footer.add_paragraph(label_text)
    run = paragraph.runs[0]
    run.font.name = 'Times New Roman'
    run.font.size = Pt(12)


    paragraph.alignment = 0  # Left alignment (0 for left-aligned)
    document.save(file_path)


    return label_text




# Function to label XLSX file with watermark
def label_xlsx_file_watermark(file_path, category, image_path_public, image_path_confidential):
    """Add watermark to an Excel file based on scan results."""
    workbook = load_workbook(file_path)
    sheet = workbook.active  


    if category == "Confidential":
        label_text = "Tài liệu Mật của BIDV. Cấm sao chép, in ấn dưới mọi hình thức!"
        img_path = image_path_confidential
    else:
        label_text = "Tài liệu Nội bộ của BIDV. Cấm sao chép, in ấn dưới mọi hình thức!"
        img_path = image_path_public


    img = Image(img_path)
    img.width, img.height = 400, 200
    sheet.add_image(img, 'A1')


    workbook.save(file_path)


    return label_text


# Function to label XLSX file footer
def label_xlsx_file_footer(file_path, category):
    """Add footer label to an Excel file based on scan results."""
    workbook = load_workbook(file_path)


    if category:
        label_text = "Tài liệu Mật của BIDV. Cấm sao chép, in ấn dưới mọi hình thức!"
    else:
        label_text = "Tài liệu Nội bộ của BIDV. Cấm sao chép, in ấn dưới mọi hình thức!"


    for sheet in workbook.worksheets:
        current_footer = sheet.oddFooter.center.text if sheet.oddFooter.center else ""
        if "&P" in current_footer:
            new_footer = f"{label_text} - Trang &P"
        else:
            new_footer = label_text


        sheet.oddFooter.center.text = new_footer
        sheet.oddFooter.center.size = 12
        sheet.oddFooter.center.font = "Times New Roman"


    workbook.save(file_path)


    return label_text


   


# call function
# file_path = "C:/Users/tienva/Documents/FolderOfTien/CongViecQuyBa2024/DLP/DLP-LabelApp/Code/files/tài liệu mật.docx"
file_path = "/Users/vuanhtien/Documents/CodeForMe/BIDV/test_file/file_docx.docx"  # Provide the file path
 # Provide the file path
r,sms= scan_file(file_path)
rule = define_rules()
a,b = classify_document_with_multiple_rules(r,rule)
print("category:",a)
print("show_rule:",b)