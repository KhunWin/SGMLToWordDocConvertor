import re
from bs4 import BeautifulSoup, NavigableString, Tag
from html import escape
from docx import Document
# from htmldocx import HtmlToDocx
from docx.shared import Inches, Pt
from datetime import datetime

class SGMLToHTMLConverter:
    def __init__(self):
        self.topic_counter = 1
        self.list1_counter = 0
        self.images_header_added = False

    
    def read_header(self, input_file):
        try:
            with open(input_file, 'r', encoding='utf-8') as f:
                sgml_content = f.read()
                
            # Extract content between <EFFXREF> and </EFFXREF> tags
            effxref_pattern = re.compile(r'<EFFXREF[^>]*>(.*?)</EFFXREF>', re.DOTALL)
            match = effxref_pattern.search(sgml_content)
            
            if match:
                effxref_content = f"<EFFXREF>{match.group(1)}</EFFXREF>"
                
                # Parse the content using BeautifulSoup
                soup = BeautifulSoup(effxref_content, 'html.parser')
                
                # Create HTML output
                html_content = self._create_html_header()
                html_content += '<div class="effectivity-table">\n'
                
                # Process title
                title_tag = soup.find('TITLE')
                if title_tag:
                    html_content += f'<h1>{title_tag.get_text()}</h1>\n'
                
                # Process paragraphs
                for para in soup.find_all('PARA'):
                    html_content += f'<p>{para.get_text()}</p>\n'
                
                # Process unordered list if present
                unlist = soup.find('UNLIST')
                if unlist:
                    html_content += '<ul>\n'
                    for item in unlist.find_all('UNLITEM'):
                        html_content += f'<li>{item.get_text()}</li>\n'
                    html_content += '</ul>\n'
                
                # Create table for EFFDATA
                html_content += '<table class="effectivity-data">\n'
                html_content += '<tr>\n'
                html_content += '<th>Customer</th>\n'
                html_content += '<th>Model Type</th>\n'
                html_content += '<th>CEC</th>\n'
                html_content += '<th>Line Number</th>\n'
                html_content += '<th>ECT</th>\n'
                html_content += '<th>Engine</th>\n'
                html_content += '<th>Standard</th>\n'
                html_content += '<th>MSN</th>\n'
                html_content += '<th>Registration</th>\n'
                html_content += '<th>Operator</th>\n'
                html_content += '</tr>\n'
                
                # Process EFFDATA entries - Debug the EFFDATA finding
                effdata_tags = soup.find_all('EFFDATA')
                # print(f"Found {len(effdata_tags)} EFFDATA tags")
                
                # If no EFFDATA tags found with uppercase, try lowercase
                if len(effdata_tags) == 0:
                    effdata_tags = soup.find_all('effdata')
                    # print(f"Found {len(effdata_tags)} effdata tags (lowercase)")
                
                for effdata in effdata_tags:
                    html_content += '<tr>\n'
                    
                    # Helper function to safely get text from a tag
                    def get_tag_text(tag_name):
                        tag = effdata.find(tag_name.upper()) or effdata.find(tag_name.lower())
                        return tag.get_text() if tag else ""
                    
                    # Add each cell with proper error handling
                    html_content += f'<td>{get_tag_text("CUS")}</td>\n'
                    html_content += f'<td>{get_tag_text("MODTYPE")}</td>\n'
                    html_content += f'<td>{get_tag_text("CEC")}</td>\n'
                    html_content += f'<td>{get_tag_text("LINENBR")}</td>\n'
                    html_content += f'<td>{get_tag_text("ECT")}</td>\n'
                    html_content += f'<td>{get_tag_text("ESNBR")}</td>\n'
                    html_content += f'<td>{get_tag_text("BENBR")}</td>\n'
                    html_content += f'<td>{get_tag_text("MSNBR")}</td>\n'
                    html_content += f'<td>{get_tag_text("ACN")}</td>\n'
                    html_content += f'<td>{get_tag_text("OPERATOR")}</td>\n'
                    html_content += '</tr>\n'
                
                html_content += '</table>\n'
                html_content += '</div>\n'
                html_content += '</body>\n</html>'
                
                # Debug - print the raw EFFXREF content to check structure
                # print("Raw EFFXREF content structure:")
                # print(soup.prettify()[:500] + "..." if len(soup.prettify()) > 500 else soup.prettify())
                
                # print("this is in sgml_v3:",html_content)
                return html_content
            else:
                print("No EFFXREF section found in the file.")
                return None
                
        except FileNotFoundError:
            print(f"Error: The file '{input_file}' was not found.")
            return None
        except Exception as e:
            print(f"An error occurred: {str(e)}")
            import traceback
            traceback.print_exc()
            return None
    
    def extract_amm_details(self, input_file):
        try:
            with open(input_file, 'r', encoding='utf-8') as f:
                sgml_content = f.read()
                
            # Find AMM tag with its attributes using regex
            amm_pattern = re.compile(r'<AMM\s+([^>]+)>')
            match = amm_pattern.search(sgml_content)
            
            if match:
                attributes = match.group(1)
                
                # Extract REVDATE and CUSNAME
                revdate_match = re.search(r'REVDATE="(\d+)"', attributes)
                cusname_match = re.search(r'CUSNAME="([^"]+)"', attributes)
                tsn_match = re.search(r'TSN="(\d+)"', attributes)
                # Format the date if found
                if revdate_match:
                    date_str = revdate_match.group(1)
                    # Convert YYYYMMDD to datetime object
                    date_obj = datetime.strptime(date_str, '%Y%m%d')
                    # Format to DD-MMM-YYYY
                    formatted_date = date_obj.strftime('%d-%b-%Y')
                else:
                    formatted_date = None

                result = {
                    'revdate': formatted_date,
                    'cusname': cusname_match.group(1) if cusname_match else None,
                    'tsn': tsn_match.group(1) if tsn_match else None
                }
                
                return result
                
            return None
            
        except Exception as e:
            print(f"Error extracting AMM details: {str(e)}")
            return None

    def convert_from_file(self, input_file, amm_code=None):
        try:
            with open(input_file, 'r', encoding='utf-8') as f:
                sgml_content = f.read()

            
            if amm_code:
                # Clean AMM code
                clean_amm = amm_code.replace('AMT ', '').replace('-', '')
                
                # Find all TASK blocks with their content until the next TASK
                task_blocks = re.split(r'(<TASK[^>]*>)', sgml_content)[1:]  # Split and keep the delimiters
                
                # Pair up the TASK tags with their content
                task_blocks = [''.join(task_blocks[i:i+2]) for i in range(0, len(task_blocks), 2)]
                
                for task_block in task_blocks:
                    # Extract the KEY attribute
                    key_match = re.search(r'KEY="([^"]*)"', task_block, re.IGNORECASE)
                    if key_match:
                        task_key = key_match.group(1)
                        if task_key.startswith('EN'):
                            task_key = task_key[2:-2]  # Remove 'EN' prefix and last 2 digits
                        
                        if task_key == clean_amm:
                            # Find where next TASK starts
                            next_task_pos = task_block.find('<TASK', 1)  # Start search after first character
                            if next_task_pos != -1:
                                task_block = task_block[:next_task_pos]
                                
                            # Convert content to HTML
                            task_html = self.get_amm_data(task_block)
                            output_name = f"task_card_{clean_amm}.docx"
                            
                            # print(task_html) ##printing html content here

                            return [(output_name, task_html)]
                
                # If no matching task found
                print(f"No matching task found for AMM code: {amm_code}")
                return []
                
        except FileNotFoundError:
            print(f"Error: The file '{input_file}' was not found.")
            return []
        except Exception as e:
            print(f"An error occurred: {str(e)}")
            import traceback
            traceback.print_exc()
            return []
        
    def validate_amm_code(self, input_string, task_data):
        # Remove 'AMT ' prefix and '-' from input string
        if input_string.startswith('AMT '):
            input_string = input_string[4:]
        input_string = input_string.replace('-', '')
        
        try:

            
            # Get KEY from TASK tag attributes
            key_value = task_data['key']  # Try lowercase
            if not key_value:
                key_value = task_data['KEY']  # Try uppercase
            
            # Remove 'EN' prefix and last two digits from KEY
            if key_value.startswith('EN'):
                key_value = key_value[2:-2]

            # Compare values
            if input_string == key_value:
                return True, []
            else:
                return False, [f"KEY mismatch: Input={input_string}, Expected={key_value}"]
                    
        except Exception as e:
            return False, [f"Validation error: {str(e)}"]
    
    def get_amm_data(self, sgml_content):

        # Remove SGML declarations
        content = re.sub(r'<!DOCTYPE[^>]*>', '', sgml_content)
        content = re.sub(r'<\?[^>]*\?>', '', content)
        
        # Parse the content using BeautifulSoup
        soup = BeautifulSoup(content, 'html.parser')
        

        html_content = self._create_html_header()
        
        for tag in soup.find_all(recursive=False):
            html_content += self._process_tag(tag)
        
        html_content += '\n</body>\n</html>'
        return html_content

    def _create_html_header(self):
        header = '<!DOCTYPE html>\n<html>\n<head>\n<style>\n'
        header += 'body { font-family: Arial, sans-serif; }\n'
        header += '.subtask { font-size: 14px; margin-top: 15px; margin-bottom: 5px; }\n'
        header += '.list1 { padding-left: 20px; }\n'
        header += '.list2 { padding-left: 20px; list-style-type: decimal; }\n'
        header += 'ol.list1 > li { font-weight: bold; }\n'
        header += 'ol.list2 > li { font-weight: normal; }\n'
        header += 'table { border-collapse: collapse; width: 100%; margin-bottom: 1em; }\n'
        header += 'th, td { border: 1px solid black; padding: 8px; text-align: left; }\n'
        header += 'th { background-color: #f2f2f2; font-weight: bold; }\n'
        # header += '.refblock { color: blue; }\n'
        header += '.refblock { color: blue; font-weight: bold; }\n'  # Updated styling

        header += '.topic-title { font-size: 18px; }\n'
        header += '.image-placeholder { border: 1px solid #ccc; padding: 10px; margin-bottom: 10px; }\n'
        header += '.graphic-content { color: red; }\n'
        header += '</style>\n</head>\n<body>\n'
        return header
        
    def _process_text(self, node):
        if isinstance(node, NavigableString):
            return escape(str(node))
        elif node.name in ['stdname', 'pan']:
            return node.get_text()
        elif node.name == 'refblock':
            return self._process_refblock(node)  # Call dedicated refblock processor
        else:
            return ''.join(self._process_text(child) for child in node.children)
    

    def _process_refblock(self, tag):
        # Get only the direct text node before any REFINT tags
        main_ref = None
        for child in tag.contents:
            if isinstance(child, NavigableString):
                main_ref = str(child).strip()
                break
        
        if main_ref:
            # Count the number of dashes in the reference
            dash_count = main_ref.count('-')
            
            if dash_count == 3:
                return f'<span class="refblock">Ref. AMM O/D {main_ref}</span>'
            else:
                return f'<span class="refblock">Ref. AMM TASK {main_ref}</span>'
        
        return ''

    def _process_tag(self, tag, level=0, list_counter=None, is_first_row=False, **kwargs):
        if tag.name is None:
            return escape(str(tag))
        
        attrs = ' '.join([f'{k}="{v}"' for k, v in tag.attrs.items()])
        result = ''

        # Add refblock to the processor mapping
        tag_processors = {
            'graphic': self._process_graphic,
            'pretopic': lambda t: self._process_topic(t, level),
            'topic': lambda t: self._process_topic(t, level),
            'subtask': lambda t: self._process_subtask(t, level),
            'table': lambda t: self._process_table(t, level),
            'row': lambda t: self._process_row(t, level, is_first_row),
            'list1': lambda t: self._process_list1(t, level),
            'l1item': lambda t: self._process_l1item(t, level),
            'list2': lambda t: self._process_list2(t, level),
            'l2item': lambda t: self._process_l2item(t, level),
            'para': lambda t: self._process_para(t, **kwargs),
            'refblock': self._process_refblock,
            # 'txtgrphc': lambda t: self.convert_txtgrphc_to_table(t),  # Call new function here
        }

        processor = tag_processors.get(tag.name)
        if processor:
            result = processor(tag)
        else:
            result = self._process_default(tag, level, attrs)
        
        return result

    def _process_graphic(self, tag):  # Note the underscore prefix
        chapnbr = tag.get('chapnbr', '')
        confnbr = tag.get('confnbr', '')
        subjnbr = tag.get('subjnbr', '')
        sectnbr = tag.get('sectnbr', '')
        func = tag.get('func', '')
        seq = tag.get('seq', '')
        confltr = tag.get('confltr', '')
        # key_formatted = f"{chapnbr}-{confnbr}-{subjnbr}-{func}-{seq}-{confltr}"
        key_formatted = f"{chapnbr}-{sectnbr}-{subjnbr}-{func}-{seq}-{confnbr}-{confltr}"

        sheet_tags = tag.find_all('sheet')
        effect_tag = tag.find('effect')
        sbeff_tag = tag.find('sbeff')
        sbnbr = sbeff_tag.get('sbnbr', '') if sbeff_tag else ''

        effrg = effect_tag.get('effrg', '') if effect_tag else ''
        sbeff_effrg = sbeff_tag.get('effrg', '') if sbeff_tag else ''

        formatted_effrg = ', '.join([f"{x[:3]}-{x[3:]}" for x in effrg.split()]) if effrg else ''
        formatted_sbeff_effrg = ', '.join([f"{x[:3]}-{x[3:]}" for x in sbeff_effrg.split()]) if sbeff_effrg else ''

        result = '<div class="graphic-container">\n'
        for sheet_index, sheet_tag in enumerate(sheet_tags, start=1):
            title_tag = tag.find('title')
            title_text = title_tag.get_text().strip() if title_tag else ''
            sheet_count = len(sheet_tags)

            result += (
                f'<div class="graphic-content">\n'
                f'<p><strong>Figure {key_formatted} (SHEET {sheet_index}/{sheet_count}) - {title_text}</strong></p>\n'
            )
            
            if formatted_effrg:
                result += f'<p>** ON A/C FSN {formatted_effrg}</p>\n'
                
            if sbeff_tag and formatted_sbeff_effrg:
                result += f'<p>EMB SB {sbnbr} for A/C {formatted_sbeff_effrg}</p>\n'
                
            result += '</div>\n'
        
        result += '</div>\n'
        return result

    def _process_topic(self, tag, level):
        title = tag.find('title')
        title_text = title.get_text() if title else ''
        result = f'<h{level+1} class="topic-title">{self.topic_counter}. {title_text}</h{level+1}>\n'

        self.topic_counter += 1
        self.list1_counter = 0
        for child in tag.children:
            if child.name != 'title':
                result += self._process_tag(child, level + 1)
        return result

    def _process_subtask(self, tag, level):
        chapnbr = tag.get('chapnbr', '')
        sectnbr = tag.get('sectnbr', '')
        subjnbr = tag.get('subjnbr', '')
        func = tag.get('func', '')
        seq = tag.get('seq', '')
        confltr = tag.get('confltr', '')
        subtask_text = f'SUBTASK {chapnbr}-{sectnbr}-{subjnbr}-{func}-{seq}-{confltr}'
        result = f'<p class="subtask">{subtask_text}</p>\n'
        for child in tag.children:
            result += self._process_tag(child, level)
        return result

    def _process_table(self, tag, level):

        attrs = ' '.join([f'{k}="{v}"' for k, v in tag.attrs.items()])
        result = f'<table {attrs}>\n'
        first_row = True
        for child in tag.children:
            if child.name == 'row':
                result += self._process_tag(child, level, None, first_row)
                first_row = False
            else:
                result += self._process_tag(child, level)
        result += '</table>\n'
        return result

    def _process_row(self, tag, level, is_first_row):
        result = '<tr>\n'
        for child in tag.children:
            if child.name == 'entry':
                if is_first_row:
                    result += f'<th>{self._process_tag(child, level)}</th>\n'
                else:
                    result += f'<td>{self._process_tag(child, level)}</td>\n'
            else:
                result += self._process_tag(child, level)
        result += '</tr>\n'
        return result

    def _process_list1(self, tag, level):
        result = '<ol class="list1" style="list-style-type: upper-alpha;">\n'
        for child in tag.children:
            result += self._process_tag(child, level+1)
        result += '</ol>\n'
        return result

    def _process_l1item(self, tag, level):
        self.list1_counter += 1
        result = f'<li value="{self.list1_counter}">'
        for child in tag.children:
            if child.name == 'para':
                result += self._process_tag(child, level, strip=True)
            else:
                result += self._process_tag(child, level)
        result += '</li>\n'
        return result

    def _process_list2(self, tag, level):
        result = '<ol class="list2">\n'
        for child in tag.children:
            result += self._process_tag(child, level+1)
        result += '</ol>\n'
        return result

    def _process_l2item(self, tag, level):
        result = '<li>'
        for child in tag.children:
            result += self._process_tag(child, level)
        result += '</li>\n'
        return result

    def _process_para(self, tag, **kwargs):
        content = ''
        for child in tag.children:
            if isinstance(child, str):
                content += escape(str(child))
            elif child.name == 'GRPHCREF':
                content += self._process_grphcref(child)
            else:
                content += self._process_tag(child, 0)
        
        if kwargs.get('strip'):
            return content.strip()
        return f'<p>{content}</p>\n'

    def _process_grphcref(self, tag):
        ref_text = ''
        for child in tag.children:
            if child.name == 'EFFECT':
                continue  # Skip EFFECT tag
            ref_text += self._process_tag(child, 0)
        
        return f'Ref. Fig. {ref_text}'
    
    def _process_default(self, tag, level, attrs):
        result = f'<{tag.name} {attrs}>' if attrs else f'<{tag.name}>'
        for child in tag.children:
            result += self._process_tag(child, level)
        result += f'</{tag.name}>'
        return result

def run_decoder(input_file='amm_a330.sgm', amm_code='AMT 052570-200-807'):
    try:
        converter = SGMLToHTMLConverter()
        html_output = converter.convert_from_file(input_file, amm_code)
        return html_output
    except Exception as e:
        print(f"Error in decoding: {str(e)}")
        return None

if __name__ == "__main__":
    # Run from command line
    import sys
    
    if len(sys.argv) > 2:
        input_file = sys.argv[1]
        amm_code = sys.argv[2]
    else:
        input_file = 'amm_a330.sgm'
        amm_code = 'AMT 052570-200-807'
    
    result = run_decoder(input_file, amm_code)
    if result:
        print(result)
