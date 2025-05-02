
from pydoc import doc
from xml.etree import ElementTree as ET
from lxml import etree
from bs4 import BeautifulSoup, NavigableString, Tag
from docx import Document
from docx.shared import Pt, Inches, Mm
from docx.oxml import parse_xml
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
import pandas as pd
from docx.oxml.ns import qn
from decode_sgml_v3 import SGMLToHTMLConverter
from docx.shared import RGBColor, Pt, Inches
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os, re, html
from find_warning import WarningHandler
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml.table import CT_Tbl
from docx.oxml.shared import qn
import math
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

class TableGenerator:

    def __init__(self, xml_content, amm_code, swe_path=None, sce_path=None, selected_effrg=None, csn=None, card_no=None, workpack_num=None,check=None,jobtype=None):
        self.selected_effrg = selected_effrg
        self.xml_content = xml_content
        self.doc = Document()
        self.amm_code = amm_code
        self.output_file = None
        self.warning_handler = WarningHandler(swe_path, sce_path)
        self.processed_texts = set()
        self.no_graphic = False
        
        ##newly added
        self.csn = csn
        self.card_no = card_no
        self.workpack_num = workpack_num
        self.check = check
        self.jobtype = jobtype


        # Set A4 page size at initialization
        section = self.doc.sections[0]
        
        # Set A4 dimensions (210mm x 297mm)
        section.page_width = Mm(210)    # A4 width
        section.page_height = Mm(297)   # A4 height
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)

        # Force A4 size using XML
        section._sectPr.xpath('./w:pgSz')[0].set(qn('w:h'), str(int(297 * 56.7)))  # 297mm in twips
        section._sectPr.xpath('./w:pgSz')[0].set(qn('w:w'), str(int(210 * 56.7)))  # 210mm in twips
        section._sectPr.xpath('./w:pgSz')[0].set(qn('w:orient'), 'portrait')

    def process_reason_for_job(self, element):
        print("Processing reason for job")
        reason_lines = []
        
        # Initialize the current element to the next sibling of the h1
        current = element.find_next_sibling()
        
        while current and current.name != 'h1':
            if isinstance(current, NavigableString):
                current = current.find_next_sibling()
                continue
            
            if current.name == 'p':
                text = current.get_text().strip()
                if text:
                    reason_lines.append(text)
            
            elif current.name == 'note' or (current.name == 'revst' and current.find('note')):

                # If it's a revst tag, get the note inside it
                note_elem = current if current.name == 'note' else current.find('note')

                # # Process paragraphs directly under note
                # note_paras = current.find_all('p', recursive=False)
                
                # Process paragraphs directly under note
                note_paras = note_elem.find_all('p', recursive=False)

                for i, para in enumerate(note_paras):
                    note_text = para.get_text().strip()
                    if note_text:
                        # First paragraph in note gets NOTE: prefix
                        if i == 0:
                            reason_lines.append("")
                            # Check if next sibling is unlist
                            next_sibling = para.find_next_sibling()
                            if next_sibling and next_sibling.name == 'unlist':
                                if not note_text.strip().endswith(':'):
                                    note_text = note_text.strip() + ':'
                            reason_lines.append(f"NOTE: {note_text}")
                        else:
                            # Subsequent paragraphs don't get NOTE: prefix
                            reason_lines.append(f" {note_text}")
                
                # Process unlist within the note
                unlist = current.find('unlist')
                if unlist:
                    for unlitem in unlist.find_all('unlitem'):
                        for item_para in unlitem.find_all('p'):
                            item_text = item_para.get_text().strip()
                            if item_text:
                                reason_lines.append(f"    - {item_text}")
            
            elif current.name == 'unlist':
                # Handle standalone unlist elements (not within a note)
                for unlitem in current.find_all('unlitem'):
                    for p in unlitem.find_all('p'):
                        text = p.get_text().strip()
                        if text:
                            reason_lines.append(f"- {text}")
            
            # Move to the next sibling
            current = current.find_next_sibling()
        
        # # Remove any trailing empty strings and duplicates
        while reason_lines and not reason_lines[-1].strip():
            reason_lines.pop()
        
        # Remove consecutive duplicate lines
        cleaned_lines = []
        prev_line = None
        for line in reason_lines:
            if line != prev_line:
                cleaned_lines.append(line)
            prev_line = line
        
        return '\n'.join(cleaned_lines)

    def process_graphic_container(self, container, content_elements):
        try:
            # Find all GRPHCREF elements in the document
            soup = container.find_parent()
            grphcrefs = soup.find_all('grphcref')
            # print(f"Found {len(grphcrefs)} GRPHCREF elements")
            
            # Get the figure number from this container
            figure_title = container.find('strong')
            if not figure_title:
                print("No figure title found in container")
                return
            
            # Check if graphic-content exists first
            graphic_contents = container.find_all('div', class_='graphic-content')
            if not graphic_contents:
                print("no graphic is found")
                return
            
            # Extract the figure number and clean it
            figure_num = ''.join(filter(str.isdigit, figure_title.get_text().split(' ')[1]))

            
            # Check if this figure matches any GRPHCREF
            should_process = False
            for grphcref in grphcrefs:
                refid = grphcref.get('refid', '')
                
                if refid:
                    # Clean the refid
                    refid_clean = ''.join(filter(str.isdigit, refid))[:-1]
                    if refid_clean == figure_num:
                        should_process = True
                        print("MATCH FOUND!")
                        break
            
            # print(f"Should process: {should_process}")
            
            # # Process the content
            if should_process:
                # print("Processing graphic content...")
                for graphic_content in container.find_all('div', class_='graphic-content'):
                    self.no_graphic = True
                    # print("Found graphic content div")
                    for p in graphic_content.find_all('p'):
                        text = p.get_text().strip()
                        # print(f"Processing text: {text}")
                        
                        if "001-999" in text:
                            text = text.replace("** ON A/C FSN 001-999", "** ON A/C FSN ALL")
                            # print(f"Replaced text: {text}")
                            
                        if p.find('strong'):
                            # print("Adding as RED text (strong)")
                            content_elements.append(('text', f"__RED__{text}__END_RED__"))
                        else:
                            print("Adding as RED text")
                            content_elements.append(('text', f"__RED__{text}__END_RED__"))
            else:
                print("\nNo matching GRPHCREF found - skipping this graphic")
                

            
        except Exception as e:
            print(f"Error processing graphic container: {str(e)}")

    def process_paragraph_content(self, p_element):

        result_parts = []  # Single list to maintain order

        def process_element(element):
            if isinstance(element, str):
                text = element.strip()
                if text:
                    result_parts.append(text)
            elif element.name == 'span' and 'refblock' in element.get('class', []):

                ref_text = element.get_text().strip()
                if ref_text:
                    result_parts.append(f"{ref_text}")

            elif element.name == 'expd':
                csn_tag = element.find('csn')
                if csn_tag:
                    csn_text = csn_tag.get_text(strip=True)
                    expdname = element.find('expdname').get_text(strip=True) if element.find('expdname') else ''
                    itemnbr = element.find('itemnbr').get_text(strip=True) if element.find('itemnbr') else ''
                    formatted_csn = self.format_csn(csn_text)
                    result_parts.append(f"{formatted_csn} {expdname} ({itemnbr})")
                
            elif element.name == 'tor':
                torque_text = self.format_torque(element)
                if torque_text:
                    result_parts.append(torque_text)
            elif element.name == 'li':
                self.handle_list_item(element, result_parts, process_element)
            elif element.name == 'stdname':
                stdname_text = element.get_text().strip()
                if stdname_text:
                    result_parts.append(stdname_text)
                # result_parts.append(element.get_text().strip())
            elif element.name == 'pan':
                result_parts.append(element.get_text().strip())
            elif element.name == 'refext':
                self.handle_refext(element, result_parts)
            elif element.name == 'ted':
                self.handle_tool_reference(element, result_parts)
            elif element.name == 'con':
                self.handle_consumable_reference(element, result_parts)
            elif element.name == 'grphcref':
                temp_parts = []
                self.handle_graphic_reference(element, temp_parts, process_element)
                if temp_parts:
                    result_parts.extend(temp_parts)
            elif element.name in ['revst', 'revend']:
                for child in element.children:
                    process_element(child)

            elif element.name == 'ein':
                # Remove hyphens from ein text
                ein_text = element.get_text(strip=True)
                ein_text = ''.join(ein_text.split('-'))
                result_parts.append(ein_text)
            else:
                for child in element.children:
                    process_element(child)

        # Process all elements in order
        for content in p_element:
            process_element(content)

        # Join all parts preserving the original order
        return " ".join(result_parts).strip()

    def format_csn(self, csn_text):

        if len(csn_text) >= 8:
            main_part = csn_text[:-4]  # Get all except last 4 chars
            item_part = csn_text[-3:]  # Get last 3 digits
            formatted_csn = '-'.join([main_part[i:i+2] for i in range(0, len(main_part), 2)])
            return f"IPC-CSN ({formatted_csn} ITEM {item_part})"
        return csn_text

    def format_torque(self, tor_element):
        """Format torque values with proper ranges and units"""
        torvalues = tor_element.find_all('torvalue')
        if len(torvalues) >= 2:
            # Get primary values (m.daN)
            primary_tor = torvalues[0]
            min_val1 = primary_tor.get('min', '')
            max_val1 = primary_tor.get('max', '')
            unit1 = primary_tor.get('unit', '')
            
            # Get secondary values (lbf.in)
            secondary_tor = torvalues[1]
            min_val2 = secondary_tor.get('min', '')
            max_val2 = secondary_tor.get('max', '')
            unit2 = secondary_tor.get('unit', '')
            
            return f"to between {min_val1} and {max_val1} {unit1} ({min_val2} and {max_val2} {unit2})"
        return ""

    def handle_graphic_reference(self, element, text_parts, process_element):

            # Check effectivity first before processing
        effect_tag = element.find('effect')
        if effect_tag:
            effrg = effect_tag.get('effrg', '')
            # Only process this graphic reference if its effectivity matches our filter
            if effrg and not self.is_effrg_included(effrg):
                return  # Skip this grphcref if its effectivity doesn't match
                
        grphc_text = []
        
        # Handle effect tag if present
        effect_tag = element.find('effect')
        effrg = effect_tag.get('effrg', '') if effect_tag else ''
        ##just adding this if statement
        if effect_tag:
            effrg = effect_tag.get('effrg', '')
            if effrg:
                # Add effect header before the graphic reference
                effect_text = f"**ON A/C FSN {self.format_effrg(effrg)}"
                text_parts.append(effect_text)
        ##end of if statement
        
        # Get direct text content excluding effect tag text
        direct_text = ''.join(child.string or '' 
                            for child in element.children 
                            if isinstance(child, NavigableString) or 
                            (child.name != 'effect' and child.name != 'revst')).strip()
        
        if direct_text:
            grphc_text.append(direct_text)
        else:
            # Handle nested revst structure if present
            revst_elem = element.find('revst')
            if revst_elem:
                revend_elem = revst_elem.find('revend')
                if revend_elem:
                    revend_text = revend_elem.get_text(strip=True)
                    if revend_text:
                        grphc_text = [revend_text]

        if grphc_text:
            # Check if this is a subsequent grphcref
            prev_sibling = element.previous_sibling
            if prev_sibling and prev_sibling.name == 'grphcref':
                # Use the same indentation as the text content
                text_parts.append(f"Ref. Fig. {' '.join(grphc_text)}")
            else:
                text_parts.append(f"Ref. Fig. {' '.join(grphc_text)}")

    def handle_list_item(self, element, text_parts, process_element):
        """Handle list items (`li`) and their children."""
        for child in element:
            if isinstance(child, str):
                text = child.strip()
                if text:
                    text_parts.append(text)
            else:
                process_element(child)

    def handle_refext(self, element, text_parts):
        """Handle `refext` elements."""
        ref_loc = element.get('refloc', '')
        ref_man = element.get('refman', '')
        refspl = element.get('refspl', '')  # Add this line to get refspl attribute

        if ref_man and ref_loc:
            formatted_ref_loc = f"{ref_loc[:2]}-{ref_loc[2:4]}-{ref_loc[4:]}"
            # formatted_ref = f"Ref. {ref_man} {ref_loc}"
            spl_text = f" ({refspl})" if refspl else ""
            formatted_ref = f"Ref. {ref_man} {formatted_ref_loc}{spl_text}"
            text_parts.append(formatted_ref)

    def handle_tool_reference(self, element, text_parts):
        """Handle `ted` elements for tool references."""
        toolname_elem = element.find('toolname')
        toolnbr_elem = element.find('toolnbr')
        toolname_text = toolname_elem.get_text().strip() if toolname_elem else ''
        toolnbr_text = toolnbr_elem.get_text().strip() if toolnbr_elem else ''
        if toolname_text and toolnbr_text:
            text_parts.append(f"{toolname_text} ({toolnbr_text})")
        elif toolname_text:
            text_parts.append(toolname_text)

    def handle_consumable_reference(self, element, text_parts):
        """Handle `con` elements for consumable references."""
        connbr_elem = element.find('connbr')
        connbr_text = connbr_elem.get_text().strip() if connbr_elem else ''
        conname_elem = element.find('conname')
        conname_text = conname_elem.get_text().strip() if conname_elem else ''
        if connbr_text and conname_text:
            text_parts.append(f"{conname_text} (Material Ref. {connbr_text})")

    def format_effrg(self, effrg_str):
        """Format the effectivity range string."""
        if not effrg_str:
            return ""
        parts = effrg_str.split()
        formatted_parts = []
        for part in parts:
            if len(part) == 6:
                formatted_parts.append(f"{part[:3]}-{part[3:]}")
            else:
                formatted_parts.append(part)
        return ", ".join(formatted_parts)

    def process_cblst(self, cblst_element):
        headers = ["PANEL", "DESIGNATION", "FIN", "LOCATION"]
        rows = []
        num_cols = len(headers)
        data_groups = []
        
        try:
            if cblst_element is None:
                raise ValueError("<cblst> element is missing or None")

            has_ein = bool(cblst_element.find('ein'))

            for cbsublst in cblst_element.find_all('cbsublst'):
                ein = cbsublst.find('ein')
                equname = cbsublst.find('equname')
                current_fin_info = None
                if ein and equname:
                    ein_text = ein.get_text(strip=True).replace('-', '')
                    equname_text = equname.get_text(strip=True)
                    current_fin_info = f"FOR FIN {ein_text} ({equname_text})"
                
                temp_data = []
                for cbdata in cbsublst.find_all('cbdata'):
                    effect_text = None
                    effrg = None
                    sb_text = None
                    
                    # Check for effect tag first - it can be at different levels
                    effect_tag = cbdata.find('effect')
                    if effect_tag:
                        effrg = effect_tag.get('effrg', '')
                        
                        # Add SB effect processing
                        sbeff = effect_tag.find('sbeff')
                        if sbeff:
                            sb_nbr = sbeff.get('sbnbr', '')
                            sb_effrg = sbeff.get('effrg', '')
                            sb_cond = sbeff.get('sbcond', '')
                            
                            if sb_nbr and sb_effrg and sb_cond:
                                sb_ranges = [f"{r[:3]}-{r[3:]}" for r in sb_effrg.split()]
                                sb_text = f"{sb_cond} SB {sb_nbr} for A/C {', '.join(sb_ranges)}"
                    
                    # Skip this cbdata if it doesn't match the filter
                    if effrg and not self.is_effrg_included(effrg):
                        continue
                    
                    # Set effect text if effrg exists
                    if effrg:
                        effect_text = "**ON A/C FSN ALL" if effrg == "001999" else f"**ON A/C FSN {self.format_effrg(effrg)}"
                        if sb_text:
                            effect_text = f"{effect_text}\n{sb_text}"
                    
                    # Extract data elements
                    panel = ""
                    designation = ""
                    fin = ""
                    location = ""
                    
                    # Find the CB tag at the cbdata level - it appears to be a direct child
                    cb_tag = cbdata.find('cb', recursive=False)
                    if cb_tag:
                        fin = cb_tag.get_text(strip=True).replace('-', '')
                    
                    # Check if we have a revst tag
                    revst = cbdata.find('revst')
                    if revst:
                        # Get designation from revst
                        cbname_tag = revst.find('cbname')
                        if cbname_tag:
                            designation = cbname_tag.get_text(strip=True)
                        
                        # Get panel and location from revend
                        revend = revst.find('revend')
                        if revend:
                            pan = revend.find('pan')
                            cbloc = revend.find('cbloc')
                            
                            if pan:
                                panel = pan.get_text(strip=True)
                            if cbloc:
                                location = cbloc.get_text(strip=True)
                    else:
                        # Standard extraction without revst
                        pan = cbdata.find('pan')
                        cbname = cbdata.find('cbname')
                        cbloc = cbdata.find('cbloc')
                        
                        if pan:
                            panel = pan.get_text(strip=True)
                        if cbname:
                            designation = cbname.get_text(strip=True)
                        if cbloc:
                            location = cbloc.get_text(strip=True)

                    row_data = [panel, designation, fin, location]
                    temp_data.append({
                        'effect': effect_text,
                        'data': row_data,
                        'panel': panel
                    })

                if temp_data:
                    if has_ein:
                        data_groups.append({
                            'effect': temp_data[0]['effect'],
                            'fin_info': current_fin_info,
                            'data': sorted(temp_data, key=lambda x: x['panel']),
                        })
                    else:
                        for item in temp_data:
                            data_groups.append({
                                'effect': item['effect'],
                                'fin_info': None,
                                'data': [item],
                            })

            current_effect = None
            current_fin_info = None
            
            for group in data_groups:
                effect = group['effect']
                fin_info = group['fin_info']
                
                if effect != current_effect:
                    if effect:
                        rows.append({
                            'merged': True,
                            'content': effect,
                            'span': num_cols
                        })
                    current_effect = effect

                if fin_info and fin_info != current_fin_info:
                    rows.append({
                        'merged': True,
                        'content': fin_info,
                        'span': num_cols
                    })
                    current_fin_info = fin_info
                
                for item in group['data']:
                    rows.append(item['data'])

        except Exception as e:
            print(f"An error occurred while processing <cblst>: {str(e)}")
            import traceback
            print(traceback.format_exc())

        return headers, rows

    ##this format is working for merging rows
    def format_table(self, headers, data_rows):
        return headers, data_rows

    def process_table(self, tgroup_element):
        """Process table elements (`tgroup`)."""
        headers = []
        data_rows = []
        seen_effect_texts = set()
        current_effect = None  # Track current effect
        last_added_effect = None  # Track the last added effect header
        num_cols = len(headers)
        tbody = tgroup_element.find('tbody')
       
        def format_refman(text):
            soup = BeautifulSoup(text, 'html.parser') if isinstance(text, str) else text
            for refext in soup.find_all('refext'):
                refloc = refext.get('refloc', '')
                refman = refext.get('refman', '')
                refspl = refext.get('refspl', '')  # Add this line to get refspl attribute

                
                if refloc and refman:
                    # Format the reference number with hyphens
                    if len(refloc) == 6:  # Ensure we have a 6-digit number
                        formatted_num = f"{refloc[:2]}-{refloc[2:4]}-{refloc[4:]}"
                        spl_text = f" ({refspl})" if refspl else ""
                        new_text = f"Ref. {refman} {formatted_num}{spl_text}"
                        refext.replace_with(new_text)
                    else:
                        # For non-standard refloc, just append refspl if it exists
                        spl_text = f" ({refspl})" if refspl else ""
                        new_text = f"{refman} {refloc}{spl_text}"
                        refext.replace_with(new_text)
                
            return soup

        def process_text_with_tool(text):
            soup = BeautifulSoup(text, 'html.parser') if isinstance(text, str) else text

            # Handle refman
            soup = format_refman(soup)

            # Handle stdname tags - keep the text in uppercase
            for stdname in soup.find_all('stdname'):
                text = stdname.get_text(strip=True)
                stdname.replace_with(text.upper())

            # Handle revst and revend tags
            for revst in soup.find_all('revst'):
                revst.unwrap()
            for revend in soup.find_all('revend'):
                revend.unwrap()

            for ted in soup.find_all('ted'):
                tool_nbr = ted.find('toolnbr').get_text(strip=True) if ted.find('toolnbr') else ''
                tool_name = ted.find('toolname').get_text(strip=True) if ted.find('toolname') else ''
                new_text = f" {tool_name} ({tool_nbr}) "
                ted.replace_with(new_text)
            
            # Handle expd tags with csn
            for expd in soup.find_all('expd'):
                csn_tag = expd.find('csn')
                if csn_tag:
                    csn_text = csn_tag.get_text(strip=True)
                    main_part = csn_text[:-4]
                    item_part = csn_text[-3:]
                    formatted_main = '-'.join([main_part[i:i+2] for i in range(0, len(main_part), 2)])
                    new_text = f"{formatted_main} ITEM {item_part}"
                    expd.replace_with(new_text)

            # Handle refint tags
            for refint in soup.find_all('refint'):
                ref_text = refint.get_text(strip=True)
                new_text = f"Ref. {ref_text}"
                refint.replace_with(new_text)
            
            # Handle grphcref tags
            for grphcref in soup.find_all('grphcref'):
                grph_text = grphcref.get_text(strip=True)
                new_text = f"Ref. Fig. {grph_text}"
                grphcref.replace_with(new_text)


            # Handle torque range format
            tor_tag = soup.find('tor')
            if tor_tag:
                torvalues = tor_tag.find_all('torvalue')
                if len(torvalues) >= 2:
                    min_val1 = torvalues[0].get('min', '')
                    max_val1 = torvalues[0].get('max', '')
                    unit1 = torvalues[0].get('unit', '')
                    min_val2 = torvalues[1].get('min', '')
                    max_val2 = torvalues[1].get('max', '')
                    unit2 = torvalues[1].get('unit', '')
                    new_text = f"to between {min_val1} and {max_val1} {unit1} ({min_val2} and {max_val2} {unit2})"
                    tor_tag.replace_with(new_text)

            # Process ein tags
            for ein_tag in soup.find_all('ein'):
                ein_text = ein_tag.get_text(strip=True)
                ein_text = ein_text.replace('-', '')
                ein_tag.replace_with(ein_text)
            
            # Handle notes
            for note_tag in soup.find_all('note'):
                note_text = note_tag.get_text(strip=True)
                new_text = f"NOTE: {note_text}"
                note_tag.replace_with(new_text)
            
            return ' '.join(soup.get_text().split())

        num_cols = int(tgroup_element.get('cols', 2))
        # print(f"DEBUG: Table has {num_cols} columns")

        thead = tgroup_element.find('thead')
        headers_in_tbody = False
        if thead:
            # print("DEBUG: Found thead")
            for entry in thead.find_all('entry'):
                p_tag = entry.find('p')
                headers.append(process_text_with_tool(p_tag if p_tag else entry))
            # print(f"DEBUG: Found headers in tbody: {headers}")

        # If no headers found in thead, look for headers in first row of tbody
        tbody = tgroup_element.find('tbody')
        if tbody and not headers:
            # print("DEBUG: Looking for headers in tbody first row")
            first_row = tbody.find('tr')
            if first_row:
                # Look for headers in td->entry->p structure
                for td in first_row.find_all('td'):
                    entry = td.find('entry')
                    if entry:
                        p_tag = entry.find('p')
                        if p_tag:
                            header_text = p_tag.get_text(strip=True)
                            headers.append(header_text)
                            headers_in_tbody = True  # Mark that headers were found in tbody

                            # print(f"DEBUG: Found header in tbody: {header_text}")

        # Update number of columns if headers were found
        num_cols = len(headers)
        # print(f"DEBUG: Final number of columns: {num_cols}")

        if tbody:
            # print("DEBUG: Looking for headers in tbody")
            skip_row = False
            last_ref_row = None

            # print("DEBUG: Processing tbody rows")
            rows = tbody.find_all('tr')
            # Skip first row if it contained headers
            start_idx = 1 if headers_in_tbody else 0
            
            # for tr in tbody.find_all('tr'):
            for tr in rows[start_idx:]:
                skip_row = False
                
                # Check for effect tags
                effect_tag = tr.find('effect')
                revst = tr.find('revst')
                if revst:
                    effect_tag = revst.find('effect')

                # Update current effect if new effect tag found
                if effect_tag:
                    effrg = effect_tag.get('effrg', '')
                    if effrg:
                        if not self.is_effrg_included(effrg):
                            skip_row = True
                            continue
                        # Only update current_effect if it's different
                        if current_effect != effrg:
                            current_effect = effrg
                            effect_text = "**ON A/C FSN ALL" if effrg == "001999" else f"**ON A/C FSN {self.format_effrg(effrg)}"
                            if last_added_effect != effect_text:  # Only add if different from last added
                                data_rows.append({
                                    'merged': True,
                                    'content': effect_text,
                                    'span': num_cols
                                })
                                last_added_effect = effect_text
                
                # Check for new reference section without effect tag
                first_entry = tr.find('entry')
                if first_entry and first_entry.find('refint'):
                    if not effect_tag:
                        next_tr = tr.find_next_sibling('tr')
                        if next_tr and not next_tr.find('effect'):
                            current_effect = "001999"
                            effect_text = "**ON A/C FSN ALL"
                            if last_added_effect != effect_text:  # Only add if different from last added
                                data_rows.append({
                                    'merged': True,
                                    'content': effect_text,
                                    'span': num_cols
                                })
                                last_added_effect = effect_text
                    last_ref_row = tr

                if skip_row:
                    continue

                # Check for EIN tags in the row
                ein_entry = tr.find('entry')
                if ein_entry and ein_entry.find('ein'):
                    ein_text = process_text_with_tool(ein_entry)
                    data_rows.append({
                        'merged': True,
                        'content': ein_text,
                        'span': num_cols
                    })
                    continue
                # Check for grphcref tags in the row
                grphcref_entry = tr.find('entry')
                if grphcref_entry:
                    grphcref = grphcref_entry.find('grphcref')
                    if grphcref:
                        # Check for effect within grphcref
                        grphc_effect = grphcref.find('effect')
                        if grphc_effect:
                            grphc_effrg = grphc_effect.get('effrg', '')
                            if not self.is_effrg_included(grphc_effrg):
                                continue
                                
                        grphcref_text = process_text_with_tool(grphcref_entry)
                        data_rows.append({
                            'merged': True,
                            'content': grphcref_text,
                            'span': num_cols
                        })
                        continue

                # Process span entries
                span_entry = tr.find('entry', attrs={'spanname': 'WHOLE'})
                if span_entry:
                    if current_effect and not self.is_effrg_included(current_effect):
                        continue
                    text_parts = self.process_table_entry(span_entry, process_text_with_tool)
                    merged_text = "\n".join(filter(None, text_parts))
                    row_data = [''] * num_cols
                    row_data[0] = merged_text
                    row_data[1] = None
                    data_rows.append(row_data)
                    continue

                # Process normal rows
                row_data = [''] * num_cols
                col_index = 0

                ####verion 6
                # Process td elements with new structure handling
                td_elements = tr.find_all('td')
                if td_elements:
                    for td in td_elements:
                        entry = td.find('entry')
                        if entry:
                            # Check for column name first
                            colname = entry.get('colname', '')
                            if colname:
                                col_index = int(colname.replace('COL', '')) - 1

                            # Use process_table_entry as the single source of processing
                            text_parts = self.process_table_entry(entry, process_text_with_tool)
                            
                            # Join all parts preserving line breaks
                            if text_parts:
                                row_data[col_index] = '\n'.join(filter(None, text_parts))
                            
                            if not colname:
                                col_index += 1
                else:
                    # Handle direct entry elements if no td elements found
                    for entry in tr.find_all('entry'):
                        colname = entry.get('colname', '')
                        if colname:
                            col_index = int(colname.replace('COL', '')) - 1
                        
                        # Use process_table_entry as the single source of processing
                        text_parts = self.process_table_entry(entry, process_text_with_tool)
                        
                        if text_parts:
                            row_data[col_index] = '\n'.join(filter(None, text_parts))
                        
                        if not colname:
                            col_index += 1
                    ###end of verion6###
                    ###the end of the code

                if any(row_data):
                # Only add effect header if it's different from the last added
                    if current_effect:
                        effect_text = "**ON A/C FSN ALL" if current_effect == "001999" else f"**ON A/C FSN {self.format_effrg(current_effect)}"
                        if effect_text != last_added_effect and tr != last_ref_row:
                            data_rows.append({
                                'merged': True,
                                'content': effect_text,
                                'span': num_cols
                            })
                            last_added_effect = effect_text
                    data_rows.append(row_data)
                    # print(f"DEBUG: Added row: {row_data}")


        return headers, data_rows
    
    def process_table_entry(self, entry, process_text_with_tool):
        """Process individual table entries."""
        text_parts = []
        current_panel = None
        warnings_elements = []

        # Process warnings and cautions directly under entry first
        warnings_elements = []
        self.process_warnings_and_cautions(entry, warnings_elements)
        for elem_type, content in warnings_elements:
            if elem_type == 'text':
                text_parts.append(content + "\n")

        # Process elements in order
        for element in entry.children:
            if isinstance(element, NavigableString):
                continue
            if element.name in ['warning', 'caution']:
                continue  # Skip as they're already processed above
            if element.name == 'p':
                text_content = process_text_with_tool(element)
                if text_content:
                    # If this p tag contains a panel header, set it as current_panel
                    if "panel" in text_content.lower():
                        current_panel = text_content
                        text_parts.append(text_content + '\n')
                    else:
                        text_parts.append(text_content + '\n')
            
            # elif element.name == 'unlist':
            #     for item in element.find_all('unlitem'):
            #         item_text = ""
            #         for p in item.find_all('p', recursive=False):
            #             processed_text = process_text_with_tool(p)
            #             if processed_text:
            #                 item_text += processed_text
            #         if item_text:
            #             text_parts.append(f"- {item_text}\n")
            elif element.name == 'unlist':
                for item in element.find_all('unlitem'):
                    item_text = ""
                    p_elements = item.find_all('p', recursive=False)
                    
                    # Handle the first paragraph
                    if p_elements:
                        processed_text = process_text_with_tool(p_elements[0])
                        if processed_text:
                            item_text = f"- {processed_text}\n"
                            text_parts.append(item_text)
                    
                    # Handle subsequent paragraphs - add them with proper line breaks
                    for p in p_elements[1:]:
                        processed_text = process_text_with_tool(p)
                        if processed_text:
                            # Add with indentation to align with the bullet
                            text_parts.append(f"  {processed_text}\n")
            elif element.name == 'note':
                note_text = process_text_with_tool(element)
                if note_text:
                    text_parts.append(f"NOTE: {note_text}\n")

        return text_parts
        
    ##adding effrg in title
    def extract_task_key_and_title(self, soup):
        try:
            # First, find the raw task string from the HTML
            html_str = str(soup)


            task_start = html_str.find('<task')
            task_end = html_str.find('>', task_start)
            if task_start != -1 and task_end != -1:
                
                task_str = html_str[task_start:task_end]
                # print(f"Found task string: {task_str}")
                
                # Extract key and confltr using string operations
                key_match = task_str.split('key="')[1].split('"')[0]
                confltr_match = task_str.split('confltr="')[1].split('"')[0]
                
                # print(f"Found key: {key_match}")
                # print(f"Found confltr: {confltr_match}")
                
                if key_match.startswith('EN'):
                    # Convert EN52310022081000 to TASK 52-31-00-220-810-A format
                    task = key_match[2:4]      # 52
                    section = key_match[4:6]    # 31
                    subject = key_match[6:8]    # 00
                    func = key_match[8:11]      # 220
                    seq = key_match[11:14]      # 810
                    
                    task_key = f"TASK {task}-{section}-{subject}-{func}-{seq}-{confltr_match}"
                    # print(f"Converted task key: {task_key}")
                else:
                    task_key = None
            else:
                print("No task string found")
                task_key = None

            # Extract title
            title_elem = soup.find('title')
            if title_elem:
                title = title_elem.text.strip()
                # print(f"Extracted title: {title}")
            else:
                print("No title element found")
                title = None

            # Extract EINs with their effects
            einlst = soup.find('einlst')
            ein_groups = {}  # Dictionary to store EINs grouped by effrg
            
            if einlst:
                for eindata in einlst.find_all('eindata'):
                    effect_elem = eindata.find('effect')
                    effrg = effect_elem.get('effrg', '001999') if effect_elem else '001999'
                    ein = eindata.find('ein')

                    # Skip if this effrg doesn't match our filter
                    if self.selected_effrg and not self.is_effrg_included(effrg):
                        continue

                    if ein:
                        ein_value = ein.text.replace('-', '')
                        if effrg not in ein_groups:
                            ein_groups[effrg] = []
                        ein_groups[effrg].append(ein_value)

            # Format EIN strings with effects
            content_list = []
            # Always process ALL (001999) first if it exists
            if '001999' in ein_groups:
                content_list.append(f"FIN: {' '.join(ein_groups['001999'])}")
            
            for effrg, eins in ein_groups.items():
                if effrg != '001999':
                    # Split effrg into ranges and format
                    formatted_effrg = self.format_effrg(effrg)
                    content_list.append(f"**ON A/C FSN {formatted_effrg}")
                    content_list.append(f"{' '.join(eins)}")

            ein_str = '\n'.join(content_list)

            # Extract main effect (effrg)
            effect_elem = soup.find('effect')
            effrg = effect_elem.get('effrg', '') if effect_elem else ''
            
            # Format effect
            if effrg == "001999":
                effect_str = f"__RED__**ON A/C FSN ALL__END_RED__"
            else:
                formatted_effrg = self.format_effrg(effrg)
                effect_str = f"__RED__**ON A/C FSN {formatted_effrg}__END_RED__"

            return effect_str, task_key, title, ein_str

        except Exception as e:
            print(f"Error in extract_task_key_and_title: {str(e)}")
            return None, None, None


    def extract_content_from_html(self, html_content):
        print("Extracting content from HTML")
        soup = BeautifulSoup(html_content, 'html.parser')

        for script in soup(["script", "style"]):
            script.decompose()

        content_elements = []
        processed_titles = set()

        # Extract the task's effrg value ##to skip the entire content if effrg is not in the range
        task_elem = soup.find('task')
        if task_elem:
            effect_elem = task_elem.find('effect')
            if effect_elem:
                effrg = effect_elem.get('effrg', '')
                if effrg and not self.is_effrg_included(effrg):
                    print(f"Effrg {effrg} is not within the selected range. Skipping content extraction.")
                    return content_elements  # Return empty list if effrg is not within range
        ###end of effrg extraction

        ##add Task and Title
        # Extract and add task key and title
        effect_key, task_key, title, ein_str = self.extract_task_key_and_title(soup)
        if effect_key:
            content_elements.append(('text', effect_key))
            # print(f"Added effect key: {effect_key}")
        if task_key:
            content_elements.append(('text', task_key))
            # print(f"Added task key: {task_key}")
        if title:
            content_elements.append(('text', title))
            # print(f"Added title: {title}")
        if ein_str:
                    content_elements.append(('text', ein_str))
                    # print(f"Added EINs: {ein_str}")
                    content_elements.append(('text', ''))  # Add a new line break

        self.handle_warnings_and_cautions(html_content, content_elements)

        self.process_titles(soup, content_elements, processed_titles)

        return content_elements

    def handle_warnings_and_cautions(self, html_content, content_elements):
        warnings_text = self.warning_handler.process_warnings(html_content)
        if warnings_text:
            content_elements.append(('text', f"\n__RED__ *{warnings_text}*__END_RED__"))
        cautions_text = self.warning_handler.process_cautions(html_content)
        if cautions_text:
            content_elements.append(('text', f"\n__YELLOW__*{cautions_text}*__END_YELLOW__"))

            
    def process_titles(self, soup, content_elements, processed_titles):
        """Process titles (h1) and their associated content."""
        for title in soup.find_all('h1', class_='topic-title'):
            title_text = title.get_text().strip()
            if title_text not in processed_titles:
                content_elements.append(('text', f"__BOLD__{title_text}__END_BOLD__".rstrip()))


                parent = title.find_parent('tfmatr')
                if parent and "Reason for the Job" in title_text:
                    reason_text = self.process_reason_for_job(title)
                    # reason_text = "text"
                    if reason_text:
                        content_elements.append(('text', reason_text))

                # self.process_effectivity(title, content_elements)
                self.process_subtasks_and_content(title, content_elements)

                processed_titles.add(title_text)

    ##this one has effectivity for Ref. Fig. 
    def process_subtasks_and_content(self, title, content_elements):
        # Process top-level notes first
        next_title = title.find_next('h1', class_='topic-title')
        current_elem = title.next_sibling
        seen_sb_effrg = set()

        # Skip processing top-level notes if title is "1. Reason for the Job"
        if title.get_text(strip=True) != "1. Reason for the Job":
            # Process top-level notes first
            elem = title.next_sibling
            while elem and elem != next_title:
                if elem.name == 'note':
                    para = elem.find('p')
                    if para:
                        note_text = para.get_text(strip=True)
                        if note_text:
                            content_elements.append(('text', f"NOTE: {note_text}"))
                elif elem.name == 'p' and 'subtask' in elem.get('class', []):
                    break
                elem = elem.next_sibling
            
            content_elements.append(('text', ''))

        # Process top-level warnings/cautions first
        elem = title.next_sibling
        while elem and elem != next_title:
            if elem.name in ['warning', 'caution']:
                warnings_and_cautions = []
                if elem.name == 'warning':
                    warning_text = elem.get_text().strip()
                    warning_id = warning_text.strip('&;')
                    if warning_id in self.warning_handler.warning_texts:
                        warning_content = self.warning_handler.warning_texts[warning_id]
                        warnings_and_cautions.append(('warning', warning_content))

                for warning_type, content in warnings_and_cautions:
                    prefix = "WARNING" if warning_type == 'warning' else "CAUTION"
                    color = "__RED__" if warning_type == 'warning' else "__YELLOW__"
                    content_elements.append(('text', f"{color} {prefix}: {content}__END_{color.strip('_')}__\n"))
            elem = elem.next_sibling

        # Process all grphcref elements first
        elem = title.next_sibling
        while elem and elem != next_title:
            if elem.name == 'grphcref':
                # Initialize variables
                effect_tag = None
                effrg = ""
                grph_text = ""
                
                # Find effect tag either directly or within revst
                revst = elem.find('revst')
                if revst:
                    effect_tag = revst.find('effect')
                    if effect_tag:
                        effrg = effect_tag.get('effrg', '')
                    
                    # Get text from revend if present
                    revend = revst.find('revend')
                    if revend:
                        grph_text = revend.get_text(strip=True)
                else:
                    effect_tag = elem.find('effect')
                    if effect_tag:
                        effrg = effect_tag.get('effrg', '')
                    
                    # Get direct text content
                    grph_text = ''.join(child.string or '' 
                                    for child in elem.children 
                                    if isinstance(child, NavigableString) or 
                                    (child.name != 'effect' and child.name != 'revst')).strip()
                
                # Apply filtering - only process if effrg matches our filter
                if effrg and self.is_effrg_included(effrg):
                    # Add effectivity header before graphic reference
                    content_elements.append(('text', f"**ON A/C FSN {self.format_effrg(effrg)}"))
                    
                    # Add the graphic reference
                    if grph_text:
                        content_elements.append(('text', f"Ref. Fig. {grph_text}"))
                
            elem = elem.next_sibling

        # Check if there's a top-level ordered list (not within subtasks)
        has_top_level_ol = False
        elem = title.next_sibling
        while elem and elem != next_title:
            if elem.name == 'ol' and 'list1' in elem.get('class', []) and not elem.find_previous_sibling('p', class_='subtask'):
                has_top_level_ol = True
                break
            elem = elem.next_sibling

        # If we have a top-level ordered list, process it first
        if has_top_level_ol:
            elem = title.next_sibling
            while elem and elem != next_title:
                if elem.name == 'ol' and 'list1' in elem.get('class', []) and not elem.find_previous_sibling('p', class_='subtask'):
                    self.process_ordered_list(elem, content_elements)
                elem = elem.next_sibling
            
            # Return early if we've processed a top-level ordered list
            return

        # Find and store all subtask-effect-content groups
        groups = []
        current_subtask = None
        current_effect = None
        current_content = None
        current_grphcref = None
        current_is_deleted = False

        # Reset to process all elements again
        current_elem = title.next_sibling
        
        # Process all elements again to find subtask-effect-content groups  
        while current_elem and current_elem != next_title:
            ###for special case
            if current_elem.name == 'revst':
                # Look for subtask and effect within revst
                revst_subtask = current_elem.find('p', class_='subtask')
                revst_effect = current_elem.find('effect')
                # Look for ordered list within revst
                revst_ol = current_elem.find('ol', class_='list1')

                if revst_subtask:
                    current_subtask = revst_subtask.get_text().strip()
                    current_effect = None
                    current_content = None
                    current_grphcref = None
                    current_is_deleted = False
                    
                if revst_effect:
                    effrg = revst_effect.get('effrg', '')
                    current_effect = {
                        'original': effrg,
                        'effrg': "ALL" if effrg == "001999" else (self.format_effrg(effrg) if effrg else ""),
                        'sbeff': []
                    }
                    
                    # Process nested SB effectivity
                    def process_sbeff(sbeff_elem, current_effect, seen_sb_combinations):
                        sb_effrg = sbeff_elem.get('effrg', '')
                        sb_cond = sbeff_elem.get('sbcond', '')
                        sb_nbr = sbeff_elem.get('sbnbr', '')
                        sb_combination = f"{sb_effrg}_{sb_nbr}"
                        
                        if sb_effrg and sb_nbr and sb_combination not in seen_sb_combinations:
                            current_effect['sbeff'].append({
                                'effrg': self.format_effrg(sb_effrg),
                                'cond': sb_cond,
                                'sbnbr': sb_nbr,
                                'original': sb_effrg
                            })
                            seen_sb_combinations.add(sb_combination)
                        
                        # Process nested sbeff elements
                        for nested_sbeff in sbeff_elem.find_all('sbeff', recursive=False):
                            process_sbeff(nested_sbeff, current_effect, seen_sb_combinations)
                    
                    # Find and process top-level sbeff elements
                    seen_sb_combinations = set()
                    
                    # Check if sbeff is nested inside revend
                    revend_tag = revst_effect.find('revend')
                    if revend_tag:
                        for sbeff_elem in revend_tag.find_all('sbeff', recursive=False):
                            process_sbeff(sbeff_elem, current_effect, seen_sb_combinations)
                    else:
                        # Process direct children of revst_effect
                        for sbeff_elem in revst_effect.find_all('sbeff', recursive=False):
                            process_sbeff(sbeff_elem, current_effect, seen_sb_combinations)
                    
                if revst_ol:
                    current_content = revst_ol

            ######
            if current_elem.name == 'p' and 'subtask' in current_elem.get('class', []):  
                if current_subtask:  # Save previous group
                    groups.append({
                        'subtask': current_subtask,
                        'effect': current_effect,
                        'content': current_content,
                        'grphcref': current_grphcref,
                        'is_deleted': current_is_deleted
                    })
                current_subtask = current_elem.get_text().strip()
                
                # Look ahead for the next effect tag that belongs to this subtask
                next_effect = None
                temp_elem = current_elem.next_sibling
                while temp_elem and (isinstance(temp_elem, NavigableString) or 
                                (temp_elem.name != 'p' and 'subtask' not in temp_elem.get('class', []))):
                    if temp_elem.name == 'effect':
                        next_effect = temp_elem
                        break
                    temp_elem = temp_elem.next_sibling
                    
                # If we found an effect tag, process it
                if next_effect:
                    effrg = next_effect.get('effrg', '')
                    current_effect = {
                        'original': effrg,
                        'effrg': "ALL" if effrg == "001999" else (self.format_effrg(effrg) if effrg else ""),
                        'sbeff': []
                    }
                    
                    # Process nested SB effectivity
                    def process_sbeff(sbeff_elem, current_effect, seen_sb_combinations):
                        sb_effrg = sbeff_elem.get('effrg', '')
                        sb_cond = sbeff_elem.get('sbcond', '')
                        sb_nbr = sbeff_elem.get('sbnbr', '')
                        sb_combination = f"{sb_effrg}_{sb_nbr}"
                        
                        if sb_effrg and sb_nbr and sb_combination not in seen_sb_combinations:
                            current_effect['sbeff'].append({
                                'effrg': self.format_effrg(sb_effrg),
                                'cond': sb_cond,
                                'sbnbr': sb_nbr,
                                'original': sb_effrg
                            })
                            seen_sb_combinations.add(sb_combination)
                        
                        # Process nested sbeff elements
                        for nested_sbeff in sbeff_elem.find_all('sbeff', recursive=False):
                            process_sbeff(nested_sbeff, current_effect, seen_sb_combinations)
                    
                    # Find and process top-level sbeff elements
                    seen_sb_combinations = set()
                    
                    # Check if sbeff is nested inside revend
                    revst_tag = next_effect.find('revst')
                    if revst_tag:
                        revend_tag = revst_tag.find('revend')
                        if revend_tag:
                            for sbeff_elem in revend_tag.find_all('sbeff', recursive=False):
                                process_sbeff(sbeff_elem, current_effect, seen_sb_combinations)
                        else:
                            # Process direct children of revst
                            for sbeff_elem in revst_tag.find_all('sbeff', recursive=False):
                                process_sbeff(sbeff_elem, current_effect, seen_sb_combinations)
                    else:
                        # Process direct children of effect
                        for sbeff_elem in next_effect.find_all('sbeff', recursive=False):
                            process_sbeff(sbeff_elem, current_effect, seen_sb_combinations)
                else:
                    current_effect = None
                    
                current_content = None
                current_grphcref = None
                current_is_deleted = False

                # Check for deleted tag
                next_elem = current_elem.next_sibling
                while next_elem and isinstance(next_elem, NavigableString):
                    next_elem = next_elem.next_sibling
                if next_elem and next_elem.name == 'revst':
                    deleted_tag = next_elem.find('deleted')
                    if deleted_tag:
                        current_is_deleted = True

            elif current_elem.name == 'effect':
                effrg = current_elem.get('effrg', '')
                current_effect = {
                    'original': effrg,
                    'effrg': "ALL" if effrg == "001999" else (self.format_effrg(effrg) if effrg else ""),
                    'sbeff': []
                }
                
                # Process nested SB effectivity
                def process_sbeff(sbeff_elem, current_effect, seen_sb_combinations):
                    sb_effrg = sbeff_elem.get('effrg', '')
                    sb_cond = sbeff_elem.get('sbcond', '')
                    sb_nbr = sbeff_elem.get('sbnbr', '')
                    sb_combination = f"{sb_effrg}_{sb_nbr}"
                    
                    if sb_effrg and sb_nbr and sb_combination not in seen_sb_combinations:
                        current_effect['sbeff'].append({
                            'effrg': self.format_effrg(sb_effrg),
                            'cond': sb_cond,
                            'sbnbr': sb_nbr,
                            'original': sb_effrg
                        })
                        seen_sb_combinations.add(sb_combination)
                    
                    # Process nested sbeff elements
                    for nested_sbeff in sbeff_elem.find_all('sbeff', recursive=False):
                        process_sbeff(nested_sbeff, current_effect, seen_sb_combinations)
                
                # Find and process top-level sbeff elements
                seen_sb_combinations = set()
                
                # Check if sbeff is nested inside revend
                revst_tag = current_elem.find('revst')
                if revst_tag:
                    revend_tag = revst_tag.find('revend')
                    if revend_tag:
                        for sbeff_elem in revend_tag.find_all('sbeff', recursive=False):
                            process_sbeff(sbeff_elem, current_effect, seen_sb_combinations)
                    else:
                        # Process direct children of revst
                        for sbeff_elem in revst_tag.find_all('sbeff', recursive=False):
                            process_sbeff(sbeff_elem, current_effect, seen_sb_combinations)
                else:
                    # Process direct children of effect
                    for sbeff_elem in current_elem.find_all('sbeff', recursive=False):
                        process_sbeff(sbeff_elem, current_effect, seen_sb_combinations)

            elif current_elem.name == 'ol' and 'list1' in current_elem.get('class', []):
                current_content = current_elem

            current_elem = current_elem.next_sibling

        # Add the last group
        if current_subtask:
            groups.append({
                'subtask': current_subtask,
                'effect': current_effect,
                'content': current_content,
                'grphcref': current_grphcref,
                'is_deleted': current_is_deleted
            })

        # Filter groups based on effrg or SB effectivity
        filtered_groups = []
        for group in groups:
            # Skip groups without effect
            if not group['effect']:
                filtered_groups.append(group)
                continue
                
            # Check if the effrg matches the filter
            if group['effect'].get('original') and self.is_effrg_included(group['effect']['original']):
                filtered_groups.append(group)
                continue
                
            # Check if any sbeff matches
            has_matching_sbeff = False
            if group['effect'] and group['effect']['sbeff']:
                for sb in group['effect']['sbeff']:
                    if sb.get('original') and self.is_effrg_included(sb.get('original')):
                        has_matching_sbeff = True
                        break
            
            # For groups with only sbeff (no effrg), include only if sbeff matches
            if not group['effect'].get('original') and has_matching_sbeff:
                filtered_groups.append(group)
        
        # Process filtered groups
        for idx, group in enumerate(filtered_groups):
            # Add a blank line between subtasks except for the first one
            if idx > 0:
                content_elements.append(('text', ''))
                
            if group['effect']:
                if group['effect']['effrg']:
                    content_elements.append(('text', f"**ON A/C FSN {group['effect']['effrg']}"))
                elif group['effect']['sbeff']:  # If no effrg but has sbeff
                    content_elements.append(('text', f"** ON A/C FSN"))
                    
                if group['effect']['sbeff']:
                    for sb in group['effect']['sbeff']:
                        content_elements.append(('text', 
                            f"{sb['cond']} SB {sb['sbnbr']} for A/C {sb['effrg']}"))

            if group['subtask']:
                content_elements.append(('text', group['subtask']))
                if group['is_deleted']:
                    content_elements.append(('text', 'DELETED'))
            if group['grphcref']:
                content_elements.append(('text', f"{group['grphcref']}\n"))
            if group['content']:
                self.process_ordered_list(group['content'], content_elements)

        # Process graphic containers after </task> tag only if this is the last h1 in the task
        task_elem = title.find_parent('task')
        if task_elem:
            h1s_in_task = task_elem.find_all('h1', class_='topic-title')
            if h1s_in_task and h1s_in_task[-1] == title:
                next_elem = task_elem.next_sibling
                graphic_container_found = False

                while next_elem:
                    if isinstance(next_elem, Tag) and next_elem.name == 'div' and 'graphic-container' in next_elem.get('class', []):
                        self.process_graphic_container(next_elem, content_elements)
                    next_elem = next_elem.next_sibling


    def is_effrg_included(self, effrg_str):
        if not effrg_str:
            return False  # Don't include empty effrg values

        # If no effrg is selected, include everything
        if self.selected_effrg is None:
            return True

        # Always include "ALL" (001999)
        if effrg_str == "001999":
            return True

        # Check if selected_effrg matches effrg_str directly or as a prefix
        if effrg_str.startswith(self.selected_effrg):
            return True

        # Split the effrg string into ranges
        ranges = effrg_str.split()
        for range_str in ranges:
            if len(range_str) != 6:
                continue

            # Extract start and end values
            start = int(range_str[:3])
            end = int(range_str[3:])

            # Check if selected_effrg falls within this range
            if start <= int(self.selected_effrg) <= end:
                return True

        return False
    
    # def handle_warnings(self, element, content_elements, recursive=False):
    #     warnings = element.find_all('warning', recursive=recursive)
    #     for warning in warnings:
    #         warning_text = warning.get_text().strip()
    #         warning_id = warning_text.strip('&;')
    #         if warning_id in self.warning_handler.warning_texts:
    #             warning_content = self.warning_handler.warning_texts[warning_id]
    #             content_elements.append(('text', f"\n__RED__ *WARNING:* {warning_content}__END_RED__"))

    # def handle_cautions(self, element, content_elements, recursive=False):
    #     cautions = element.find_all('caution', recursive=recursive)
    #     for caution in cautions:
    #         caution_text = caution.get_text().strip()
    #         caution_id = caution_text.strip('&;')
    #         if caution_id in self.warning_handler.caution_texts:
    #             caution_content = self.warning_handler.caution_texts[caution_id]
    #             content_elements.append(('text', f"\n__YELLOW__ *CAUTION:* {caution_content}__END_YELLOW__"))

    def process_expd_content(self, para):
        result_text = []
        
        # Get all text nodes and elements in order
        for element in para.children:
            if isinstance(element, str):
                text = element.strip()
                if text:
                    result_text.append(text)
            elif element.name == 'expd':
                csn = element.find('csn')
                expdname = element.find('expdname')
                itemnbr = element.find('itemnbr')
                
                # Format CSN with hyphens
                if csn:
                    csn_text = csn.get_text().strip()
                    formatted_csn = '-'.join([
                        csn_text[:2],
                        csn_text[2:4],
                        csn_text[4:6],
                        csn_text[6:8],
                        csn_text[8:]
                    ])
                    
                    # Build the formatted text
                    expd_parts = ["ICP-CSN", f"({formatted_csn})"]
                    if expdname:
                        expd_parts.append(expdname.get_text().strip())
                    if itemnbr:
                        expd_parts.append(itemnbr.get_text().strip())
                    
                    result_text.append(' '.join(expd_parts))
            elif element.name == 'refblock':
                ref_text = element.get_text().strip()
                result_text.append(f"Ref. AMM TASK {ref_text}")
        
        return ' '.join(result_text)
    
    def get_letter_from_number(self,number):
        # Skip I and O by adjusting the number
        if number > 8:  # After H
            number += 1  # Skip I
        if number > 14:  # After N
            number += 1  # Skip O
        return chr(64 + number)

    def get_chgdescs(self,sub):
        chgdescs = []
        current = sub
        while current and current.next_sibling:
            current = current.next_sibling
            if current.name == 'chgdesc':
                chgdescs.append(current.text)
            elif current.name == 'ol':
                break
        return chgdescs

    def get_adjusted_value(self, li_element, original_value):
        subtask = li_element.find_previous('p', class_='subtask')
        if not subtask:
            return original_value

        # Find current topic title
        current_topic = subtask.find_previous('h1', class_='topic-title')
        if not current_topic:
            return original_value

        # Get current subtask number without last suffix
        current_subtask_text = subtask.get_text(strip=True)
        if 'SUBTASK' not in current_subtask_text:
            return original_value
        current_base = current_subtask_text.split('SUBTASK ')[1].rsplit('-', 1)[0]
        
        # print(f"\nProcessing current subtask base: {current_base}")
        
        # Initialize value counter and tracking dict
        value_counter = 1
        seen_bases = {}  # Track base -> value mapping
        
        # Find all previous subtasks within same topic
        temp_subtask = current_topic.find_next('p', class_='subtask')
        while temp_subtask:
            # Break if we're in a different topic
            if temp_subtask.find_previous('h1', class_='topic-title') != current_topic:
                break
                
            temp_text = temp_subtask.get_text(strip=True)
            if 'SUBTASK' in temp_text:
                temp_base = temp_text.split('SUBTASK ')[1].rsplit('-', 1)[0]
                # print(f"Found subtask base: {temp_base}")
                
                # If this is a new base
                if temp_base not in seen_bases:
                    if seen_bases:  # If we've seen other bases before
                        value_counter += 1
                    seen_bases[temp_base] = value_counter
                    # print(f"New base {temp_base} -> value {value_counter}")
                
                if temp_subtask == subtask:
                    final_value = seen_bases[temp_base]
                    # print(f"Found target subtask, returning value: {final_value}")
                    return final_value
                    
            temp_subtask = temp_subtask.find_next('p', class_='subtask')
        
        # print(f"No matching subtask found, returning counter: {value_counter}")
        return value_counter

    
    def process_table_element(self, table, content_elements, processed_tables, indent=""):
    
        if not table:
            return
            
        table_id = hash(str(table))
        if table_id not in processed_tables:
            # Extract and process title first
            title = table.find('title')
            if title:
                title_text = title.get_text(strip=True)
                if title_text:
                    content_elements.append(('text', f"\nTable -  {title_text}"))
                    
            tgroup = table.find('tgroup')
            # print(f"DEBUG: Found tgroup: {tgroup}")
            if tgroup:
                try:
                    headers, raw_data_rows = self.process_table(tgroup)
                    if headers and raw_data_rows:
                        formatted_headers, formatted_rows = self.format_table(headers, raw_data_rows)
                        content_elements.append(('table', (formatted_headers, formatted_rows)))
                        processed_tables.add(table_id)
                except Exception as e:
                    print(f"DEBUG: Error processing table content: {str(e)}")
        else:
            print(f"DEBUG: Table {table_id} already processed")

    def process_warnings_and_cautions(self, element, content_elements):
        warnings_and_cautions = []
        for child in element.children:
            if child.name == 'warning':
                warning_text = child.get_text().strip()
                warning_id = warning_text.strip('&;')
                if warning_id in self.warning_handler.warning_texts:
                    warning_content = self.warning_handler.warning_texts[warning_id]
                    warnings_and_cautions.append(('warning', warning_content))
            elif child.name == 'caution':
                caution_text = child.get_text().strip()
                caution_id = caution_text.strip('&;')
                if caution_id in self.warning_handler.caution_texts:
                    caution_content = self.warning_handler.caution_texts[caution_id]
                    warnings_and_cautions.append(('caution', caution_content))
        
        for warning_type, content in warnings_and_cautions:
            prefix = "WARNING" if warning_type == 'warning' else "CAUTION"
            color = "__RED__" if warning_type == 'warning' else "__YELLOW__"
            content_elements.append(('text', f"\n{color} *{prefix}: {content}*__END_{color.strip('_')}__"))

    
    def extract_topic_notes_for_ol(self, current_ol, content_elements):
        notes = []
        previous = current_ol.find_previous_sibling()

        while previous:
            if previous.name == 'h1' and 'topic-title' in previous.get('class', []):
                break
            if previous.name == 'ol':
                break

            if previous.name == 'note':
                para = previous.find('p')
                if para:
                    note_text = para.get_text(strip=True)
                    if note_text:
                        notes.insert(0, f"NOTE: {note_text}")

            elif previous.name is None and isinstance(previous, str):
                stripped_text = previous.strip()
                if stripped_text.startswith("NOTE:"):
                    notes.insert(0, stripped_text)

            previous = previous.find_previous_sibling()

        for note in notes:
            content_elements.append(('text', note))
    
    def extract_topic_level_caution(self, ol_element, content_elements):

        # Find the closest preceding h1 topic-title and check for caution/warning
        topic_title = ol_element.find_previous('h1', class_='topic-title')
        if topic_title:
            # Get the next topic title if it exists
            next_topic_title = topic_title.find_next('h1', class_='topic-title')
            
            # Find the first subtask after topic_title
            first_subtask = topic_title.find_next('p', class_='subtask')
            # Find the subtask before current ol_element
            current_subtask = ol_element.find_previous('p', class_='subtask')
            
            # Only process if this is the first ordered list after topic title
            # AND there's a grphcref and caution immediately after the topic title
            if first_subtask and current_subtask and first_subtask == current_subtask:
                # Make sure we're looking at elements between current topic title and next topic title
                grphcref = topic_title.find_next_sibling('grphcref')
                if grphcref and (not next_topic_title or grphcref.sourceline < next_topic_title.sourceline):
                    caution = grphcref.find_next_sibling('caution')
                    if caution and (not next_topic_title or caution.sourceline < next_topic_title.sourceline):
                        caution_text = caution.get_text().strip()
                        caution_id = caution_text.strip('&;')
                        if caution_id in self.warning_handler.caution_texts:
                            caution_content = self.warning_handler.caution_texts[caution_id]
                            content_elements.append(('text', f"\n__YELLOW__ CAUTION: {caution_content}__END_YELLOW__"))
    
    def process_unlist(self, unlist_element, content_elements, parent_context, indent=""):
        if not unlist_element:
            return

        # Determine indentation based on context
        if parent_context == 'list2':
            indent = "                  "
        elif parent_context == 'list4':
            indent = "                      "
        
        for unlitem in unlist_element.find_all('unlitem'):
            para_text = self.process_paragraph_content(unlitem.find(['p', 'para']))
            if para_text:
                content_elements.append(('text', f"{indent}- {para_text}"))

    ###the original
    def process_txtgrphc_element(self, txtgrphc, content_elements, processed_txtgrphc, indent=""):
        if not txtgrphc:
            return
        txtgrphc_id = hash(str(txtgrphc))
        if txtgrphc_id not in processed_txtgrphc:
            for txtline in txtgrphc.find_all('txtline'):
                line_text = txtline.string if txtline.string else ""
                formatted_text = line_text.replace(' ', '\xa0')
                # Add the indent parameter to the formatted text
                content_elements.append(('text', f"{indent}{formatted_text}"))
            processed_txtgrphc.add(txtgrphc_id)

    def process_notes(self, note, content_elements, indent=""):
        if note:
            # Find all paragraphs in the note
            paras = note.find_all(['p', 'PARA', 'para'])
            # print(f"\nNumber of paragraphs found: {len(paras)}")
            
            first_para = True
            for idx, para in enumerate(paras):                
                # Check for txtgrphc
                txtgrphc = para.find('txtgrphc')
                # print(f"Found txtgrphc?: {'Yes' if txtgrphc else 'No'}")
                
                if txtgrphc:
                    # Add blank line before txtgrphc
                    content_elements.append(('text', ''))
                    self.process_txtgrphc_element(txtgrphc, content_elements, set(), indent)
                    # print("Finished processing txtgrphc")
                    continue
                
                note_text = self.process_paragraph_content(para)
                if note_text and note_text.strip():
                    if first_para:
                        # print(f"Adding first paragraph with NOTE: prefix")
                        content_elements.append(('text', f"{indent}NOTE: {note_text}"))
                        first_para = False
                    else:
                        # print(f"Adding subsequent paragraph")
                        content_elements.append(('text', f"{indent}      - {note_text}"))

    def process_ordered_list(self, ol_element, content_elements):
        # return None
        processed_tables = set()
        processed_txtgrphc = set()

        self.extract_topic_level_caution(ol_element, content_elements)

        for li in ol_element.find_all('li', recursive=False):
            self.process_warnings_and_cautions(li, content_elements)
            # Find the controlling effect for this list item
            
            effect = li.find_previous_sibling('effect')
            if effect:
                effrg = effect.get('effrg', '')
                if not self.is_effrg_included(effrg):
                    continue  # Skip this list item if its effect is not included

            value = li.get('value', '')

            if value:
                adjusted_value = self.get_adjusted_value(li, int(value))
                letter = self.get_letter_from_number(adjusted_value)
            else:
                letter = '•'

            # Process direct content
            grphcref_texts = []
            refblock_text = None
            direct_text = []
            for content in li.contents:
                if isinstance(content, str):
                    text = content.strip()
                    if text:
                        direct_text.append(text)
                elif content.name == 'revst':
                    revst_text = content.get_text().strip()
                    if revst_text:
                        direct_text.append(revst_text)
                elif content.name == 'grphcref':
                    temp_text = []
                    self.handle_graphic_reference(content, temp_text, lambda x: None)

                    if temp_text:
                        # grphcref_text = ' '.join(temp_text)
                        grphcref_texts.extend(temp_text)  # Add to list instead of overwriting

                elif content.name == 'span' and 'refblock' in content.get('class', []):
                    ref_text = content.get_text().strip()
                    if ref_text:
                        # refblock_text = f"__BLUE__{ref_text}__END_BLUE__"
                        refblock_text = f"{ref_text}"
                elif content.name == 'p' and content.parent == li:
                    ein = content.find('ein')
                    if ein:
                        ein_text = ein.get_text(strip=True)
                        ein_text = ''.join(ein_text.split('-'))
                        direct_text.append(ein_text)
                    else:
                        para_text = self.process_paragraph_content(content)
                        if para_text:
                            direct_text.append(para_text)

                    pan = content.find('pan')
                    if pan:
                        pan_text = pan.get_text(strip=True)
                        if pan_text:
                            content_elements.append(('text', f"    {pan_text}"))

            if direct_text or refblock_text:
                if direct_text:
                    content_elements.append(('text', f"{letter}. {' '.join(direct_text)}"))
                if refblock_text:
                    content_elements.append(('text', f"     {refblock_text}"))
                # if grphcref_text:
                #     content_elements.append(('text', f"    {grphcref_text}"))
                for grphcref_text in grphcref_texts:
                    content_elements.append(('text', f"        {grphcref_text}"))
                
                notes = li.find_all(['NOTE', 'note'], recursive=False)
                for note in notes:
                    self.process_notes(note, content_elements, "    ")

            txtgrphc = li.find('txtgrphc', recursive=False)
            if txtgrphc:
                self.process_txtgrphc_element(txtgrphc, content_elements, processed_txtgrphc, "    ")

            # Process nested lists
            nested_list = li.find('ol', class_='list2') or li.find('list2')
            if nested_list:
                for idx, nested_item in enumerate(nested_list.find_all(['li', 'l2item'], recursive=False), 1):
                    self.process_warnings_and_cautions(nested_item, content_elements)

                    # Process main text of nested item
                    nested_text = []
                    grphcref_content = []  # Store grphcref content temporarily

                    for p in nested_item.find_all(['p', 'para'], recursive=False):
                        if not p.find('txtgrphc'):  # Skip if contains txtgrphc
                            para_text = self.process_paragraph_content(p)
                            if para_text:
                                nested_text.append(para_text)
                            # Handle stdname and pan tags

                    if nested_text:
                        content_elements.append(('text', f"      {idx}. {' '.join(nested_text)}"))
                        # Process all direct children in order
                        for child in nested_item.children:
                            if child.name == 'unlist':
                                for unlitem in child.find_all('unlitem'):
                                    para_text = self.process_paragraph_content(unlitem.find(['p', 'para']))
                                    if para_text:
                                        content_elements.append(('text', f"                  - {para_text}"))                        
                    
                    
                    ##process txtgrphc inside note
                    for p in nested_item.find_all(['p', 'para'], recursive=False):
                        txtgrphc = p.find('txtgrphc')
                        if txtgrphc:
                            self.process_txtgrphc_element(txtgrphc, content_elements, processed_txtgrphc, "      ")


                    # Process notes
                    notes = nested_item.find_all(['NOTE', 'note'], recursive=False)
                    for note in notes:
                        self.process_notes(note, content_elements, "    ")

                    # Process list3
                    list3 = nested_item.find('list3')
                    if list3:
                        for l3_idx, l3item in enumerate(list3.find_all('l3item'), 1):
                            alpha_idx = chr(ord('a') + l3_idx - 1)
                            self.process_warnings_and_cautions(l3item, content_elements)
                            
                            # Process l3item paragraphs
                            l3_paras = l3item.find_all(['p', 'para'], recursive=False)
                            for p_idx, p in enumerate(l3_paras):
                                # Skip if it's part of table or warning
                                if p.find_parent('table') or p.find_parent('warning'):
                                    continue
                                    
                                para_text = self.process_paragraph_content(p)
                                if para_text:
                                    if p_idx == 0:
                                        content_elements.append(('text', f"          ({alpha_idx}) {para_text}"))
                                    else:
                                        content_elements.append(('text', f"              {para_text}"))

                            # Process notes after paragraphs //this cause problem in 3530
                            ##1232 note is missing if i comment this out
                            notes = l3item.find_all(['NOTE', 'note'])
                            for note in notes:

                                # Skip if this note is inside list4 to prevent duplication
                                if note.find_parent('list4'):
                                    continue
                                    
                                # Skip if this note contains unlist (specific to file 1 case)
                                if note.find('unlist'):
                                    continue

                                note_paras = note.find_all(['p', 'para'])
                                for para in note_paras:
                                    note_text = self.process_paragraph_content(para)
                                    if note_text:
                                        content_elements.append(('text', f"               NOTE:    {note_text}"))

                            # Process list4 immediately after its parent l3item text
                            list4 = l3item.find('list4')
                            if list4:
                                for l4_idx, l4item in enumerate(list4.find_all('l4item'), 1):
                                    self.process_warnings_and_cautions(l4item, content_elements)
                                    
                                    # # Process paragraphs
                                    for p in l4item.find_all(['p', 'para'], recursive=False):

                                        # Process paragraph text before handling txtgrphc
                                        para_text = self.process_paragraph_content(p)
                                        if para_text and not p.find('txtgrphc'):
                                            content_elements.append(('text', f"                {l4_idx}. {para_text}"))

                                        # Skip if paragraph contains txtgrphc
                                        if p.find('txtgrphc'):
                                            txtgrphc = p.find('txtgrphc')
                                            if txtgrphc:
                                                self.process_txtgrphc_element(txtgrphc, content_elements, processed_txtgrphc, "      ")
                                            continue
                                    

                                    # # Process table within list4 item
                                    # table = l4item.find('table', recursive=False)
                                    # if table:
                                    #     self.process_table_element(table, content_elements, processed_tables)

                                    # Process unlist in l4item
                                    unlist = l4item.find('unlist', recursive=False)
                                    if unlist:
                                        for unlitem in unlist.find_all('unlitem'):
                                            para_text = self.process_paragraph_content(unlitem.find(['p', 'para']))
                                            if para_text:
                                                content_elements.append(('text', f"                      - {para_text}"))

                                    # Process notes before table/txtgrphc
                                    notes = l4item.find_all('note', recursive=False)    
                                    for note in notes:
                                        txtgrphc_in_note = note.find('txtgrphc')
                                        if txtgrphc_in_note:
                                            # Process note text first
                                            note_text = ' '.join(p.get_text().strip() for p in note.find_all('p') 
                                                            if not p.find('txtgrphc'))
                                            if note_text:
                                                content_elements.append(('text', f"                  NOTE: {note_text}"))
                                                content_elements.append(('text', ''))  # Add blank line
                                            
                                            # Then process table title
                                            self.process_txtgrphc_element(txtgrphc_in_note, content_elements, 
                                                                        processed_txtgrphc, "\t\t")
                                        else:
                                            # Handle regular note
                                            note_text = note.get_text().strip()
                                            if note_text:
                                                content_elements.append(('text', f"                  NOTE: {note_text}"))

                                    # Process table within list4 item
                                    table = l4item.find('table', recursive=False)
                                    if table:
                                        self.process_table_element(table, content_elements, processed_tables)

                                    # Add processing for list5
                                    list5 = l4item.find('list5')
                                    if list5:
                                        for l5_idx, l5item in enumerate(list5.find_all('l5item'), 1):
                                            # Convert number to letter (1->a, 2->b, etc.)
                                            alpha_idx = chr(ord('a') + l5_idx - 1)

                                            # Process l5item paragraphs
                                            for p in l5item.find_all(['p', 'para'], recursive=False):
                                                para_text  = self.process_paragraph_content(p)  # Modified to handle tuple return
                                                if para_text:
                                                    content_elements.append(('text', f"                 ({alpha_idx})  {para_text}"))
                                            # Process all unlist elements within l5item
                                            for unlist in l5item.find_all('unlist'):
                                                for unlitem in unlist.find_all('unlitem'):
                                                    para = unlitem.find(['p', 'para'])
                                                    if para:
                                                        para_text = self.process_paragraph_content(para)  # Modified to handle tuple return
                                                        if para_text:
                                                            content_elements.append(('text', f"                      - {para_text}"))
                                                       
                                            # Process notes within l5item
                                            notes = l5item.find_all('note', recursive=False)
                                            for note in notes:
                                                for p in note.find_all(['p', 'para']):
                                                    note_text = self.process_paragraph_content(p)
                                                    if note_text:
                                                        content_elements.append(('text', f"                     NOTE: {note_text}"))
                                            # Process table within list4 item
                                            table = l5item.find('table', recursive=False)
                                            if table:
                                                self.process_table_element(table, content_elements, processed_tables)

                            # Process table within l3item after all text content
                            table = l3item.find('table', recursive=False)
                            if table:
                                self.process_table_element(table, content_elements, processed_tables)

                            # Process unlist in l3item
                            unlist = l3item.find('unlist', recursive=False)
                            if unlist:
                                for unlitem in unlist.find_all('unlitem'):
                                    para_text = self.process_paragraph_content(unlitem.find(['p', 'para']))
                                    if para_text:
                                        content_elements.append(('text', f"               -{para_text}"))

                    # Process remaining tables at nested_item level
                    table = nested_item.find('table', recursive=False)
                    if table and table not in processed_tables:
                        self.process_table_element(table, content_elements, processed_tables)

            # # Process remaining tables at main list level
            table = li.find('table', recursive=False)
            if table and table not in processed_tables:
                self.process_table_element(table, content_elements, processed_tables)

            cblst = li.find('cblst')
            if cblst:
                table_id = hash(str(cblst))
                if table_id not in processed_tables:
                    headers, data_rows = self.process_cblst(cblst)
                    if headers and data_rows:
                        formatted_headers, formatted_rows = self.format_table(headers, data_rows)
                        content_elements.append(('table', (formatted_headers, formatted_rows)))
                        processed_tables.add(table_id)
    
###########end of keeping all function out of process_ordered_list
    
##do not change this functions
    def create_nested_table(self, input_file,mpd_detail,msn_num, only_mpd=False):
        try:

            # Set document default style
            style = self.doc.styles['Normal']
            style.font.name = 'Cambria (Body)'
            style.font.size = Pt(7) ##for header font size
            style.paragraph_format.space_after = Pt(0)
            style.paragraph_format.space_before = Pt(0)
            style.paragraph_format.line_spacing = 1.0

            # print("HTML creating.")

            converter = SGMLToHTMLConverter()  # Assuming you have this class defined
            # print("Creating table formats")

            # print("header_rows:", self.extract_header_rows(input_file,msn_num))
            header_rows = self.extract_header_rows(input_file,msn_num)

            # ##if we delete this and when msn_num is empty, it will shoot "An error occurred: list index out of range" 
            if header_rows is None or not header_rows:
                # Use default values when no MSN number or header rows not found
                header_rows = [''] * 10  # Create list with empty strings for required indices
                print("No header rows found or MSN number not provided. Using default values.")
            
            # no MSN was provided, take the first row
            elif isinstance(header_rows, list) and header_rows and isinstance(header_rows[0], list):
                header_rows = header_rows[0]  # Take the first row
                # print("Multiple rows found, using first row")

            amm_details = converter.extract_amm_details(input_file)

            output_files = converter.convert_from_file(input_file, self.amm_code)
            
            if not output_files:
                print("No valid task cards to create.")
                return
                
            for output_name, html_output in output_files:
                print(f"Processing HTML content for {output_name}")
                if only_mpd:
                    content_elements = ""

                else:
                    content_elements = self.extract_content_from_html(html_output)
                
                print(f"After extracting content for {output_name}")

                ##adding footer and header
                self.create_header(self.doc, amm_details['cusname'],header_rows[0],header_rows[8],
                header_rows[1],mpd_detail['skill'],mpd_detail['zones'])

                ##end of adding footer and header
                
                # Create main table
                print("Creating main table...")

                main_table = self.doc.add_table(rows=1, cols=10) 
                # Insert the main table at the beginning of the document

                main_table.style = 'Table Grid'
                main_table.allow_autofit = False
                main_table.width = Inches(8.5)
                
                # Set up main table structure
                row = main_table.rows[0]
                main_cell = row.cells[0].merge(row.cells[1]).merge(row.cells[2]).merge(
                    row.cells[3]).merge(row.cells[4]).merge(row.cells[5]).merge(
                    row.cells[6]).merge(row.cells[7])
                
                ##############################
               
                # Set fixed row height to match page height
                tr = row._tr
                trPr = tr.get_or_add_trPr()
                trHeight = OxmlElement('w:trHeight')
                trHeight.set(qn('w:val'), str(10050))  # Approximately 6.7 inches in twips (minus header/footer) ##9900 this is for border of main cells
                trHeight.set(qn('w:hRule'), 'atLeast')  # Use 'exact' instead of 'atLeast'
                trPr.append(trHeight)

                # Create and set table properties
                tblPr = parse_xml(r'<w:tblPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
                            r'<w:tblStyle w:val="TableGrid"/>' +
                            r'<w:tblW w:w="0" w:type="auto"/>' +
                            r'<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>' +
                            r'<w:tblBorders>' +
                            r'<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
                            r'<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
                            r'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
                            r'<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
                            r'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
                            r'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
                            r'</w:tblBorders>' +
                            r'</w:tblPr>')

                # Insert table properties at the beginning of the table
                main_table._tbl.insert(0, tblPr)

                # Set vertical alignment and borders for cells
                for cell in [main_cell, row.cells[8], row.cells[9]]:
                    tcPr = cell._tc.get_or_add_tcPr()
                    
                    # Set vertical alignment to top
                    tcVAlign = OxmlElement('w:vAlign')
                    tcVAlign.set(qn('w:val'), "top")
                    tcPr.append(tcVAlign)

                    # Set cell borders
                    tcBorders = OxmlElement('w:tcBorders')
                    
                    # Add vertical borders
                    for border_type in ['left', 'right']:
                        border = OxmlElement(f'w:{border_type}')
                        border.set(qn('w:val'), 'single')
                        border.set(qn('w:sz'), '4')
                        border.set(qn('w:space'), '0')
                        border.set(qn('w:color'), 'auto')
                        tcBorders.append(border)
                    
                    tcPr.append(tcBorders)

                ##############################

                # Add flag to track if txtline is found
                has_txtline = any(
                    content_type == 'text' and '\xa0' in content 
                    for content_type, content in content_elements
                )

                # Set cell widths based on whether txtline was found
                if has_txtline:
                    main_cell.width = Inches(6.0)  # Narrower width for txtline content
                else:
                    main_cell.width = Inches(6.9)  # Original width

                ##############################
                
                # main_cell.width = Inches(6.0) ##6.9
                row.cells[8].width = Inches(1) ##1
                row.cells[9].width = Inches(1.1)

                row.cells[8].text = "MECH"
                row.cells[9].text = "INSP"

                ##############################

                # Create main table
                # Add base text at the beginning
                try:
                    # Create table with 3 rows and 3 columns
                    table = main_cell.add_table(rows=3, cols=3)
                    table.allow_autofit = False
                    
                    # First row: AMS ITEM REV and BOX
                    cell1_1 = table.cell(0, 0)
                    run1 = cell1_1.paragraphs[0].add_run("AMS ITEM REV")
                    run1.font.size = Pt(7)
                    
                    # BOX in third column
                    box_cell = table.cell(0, 2)
                    box_cell.text = "AMM Rev. " + amm_details['tsn']  +"\n" + amm_details['revdate'] 
                    
                    # Set column widths
                    for i, width in enumerate([3.5, 1.0, 0.9]):  # inches
                        for cell in table.column_cells(i):
                            tc = cell._tc
                            tcPr = tc.get_or_add_tcPr()
                            tcW = OxmlElement('w:tcW')
                            tcW.set(qn('w:w'), str(int(width * 1440)))  # convert to twips
                            tcW.set(qn('w:type'), 'dxa')
                            tcPr.append(tcW)
                    
                    # Set height only for first row
                    tr = table.rows[0]._tr
                    trPr = tr.get_or_add_trPr()
                    trHeight = OxmlElement('w:trHeight')
                    trHeight.set(qn('w:val'), str(500))
                    trHeight.set(qn('w:hRule'), "exact")
                    trPr.append(trHeight)
                    
                    # Second row: Rest of text content
                    cell2_1 = table.cell(1, 0)

                    if only_mpd: 
                        text_content_1 = "INTERVAL\nTHRESHOLD\nSTRUCTURE\nACCESS: \
                        \n\nMS:" 
            

                        text_content_2 = self.mpd_content(mpd_detail['title'], mpd_detail['smtdesc'], 
                        mpd_detail['mpd'], mpd_detail['sourceref'], mpd_detail['amm'], mpd_detail['list'], 
                        mpd_detail['description'])
                        #title, smtdesc, mpd, sourceref, amm, list_title, description
                        print("this is text_content2:" , text_content_2)

                        text_content = text_content_1 + text_content_2
                    else:
                        # text_content = "INTERVAL\nTHRESHOLD\nSTRUCTURE\nACCESS: \
                        # \n\nMS: \
                        # \nTitle:\nPLANNING NOTE\nEFFECTIVITY:\nREFERENCES\n"

                        text_content_1 = "INTERVAL\nTHRESHOLD\nSTRUCTURE\nACCESS: \
                        \n\nMS:" 
            

                        text_content_2 = self.mpd_content(mpd_detail['title'], mpd_detail['smtdesc'], 
                        mpd_detail['mpd'], mpd_detail['sourceref'], mpd_detail['amm'],'', 
                        '')
                        #title, smtdesc, mpd, sourceref, amm, list_title, description

                        print("this is text_content2:" , text_content_2)

                        text_content = text_content_1 + text_content_2

 
                    run2 = cell2_1.paragraphs[0].add_run(text_content)
                    run2.font.size = Pt(7)
                    
                    # Merge first column cells
                    cell1_1.merge(table.cell(2, 0))
                    
                    # Set borders only for BOX cell
                    tc = box_cell._tc
                    tcPr = tc.get_or_add_tcPr()
                    tcBorders = OxmlElement('w:tcBorders')
                    for border in ['top', 'left', 'bottom', 'right']:
                        element = OxmlElement(f'w:{border}')
                        element.set(qn('w:val'), 'single')
                        tcBorders.append(element)
                    tcPr.append(tcBorders)

                except Exception as e:
                    print(f"Error creating table: {str(e)}")

                ###removing the bottom border
                # Remove bottom borders
                for cell in [main_cell, row.cells[8], row.cells[9]]:
                    tcPr = cell._tc.get_or_add_tcPr()
                    tcBorders = OxmlElement('w:tcBorders')
                    bottom = OxmlElement('w:bottom')
                    bottom.set(qn('w:val'), 'nil')
                    tcBorders.append(bottom)
                    tcPr.append(tcBorders)

                # Process content elements in order


                ##adding color format and bold to texts.
                for content_type, content in content_elements:

                    if content_type == 'text':
                        paragraph = main_cell.add_paragraph() 

                        if '__RED__' in content and '__END_RED__' in content:
                            pure_text = content.replace('__RED__', '').replace('__END_RED__', '')
                            run = paragraph.add_run(pure_text)
                            run.font.color.rgb = RGBColor(255, 0, 0)
                            if '*' in pure_text:  # Check for italic formatting
                                run.italic = True
                        # elif '__BLUE__' in content and '__END_BLUE__' in content:
                        #     pure_text = content.replace('__BLUE__', '').replace('__END_BLUE__', '')
                        #     run = paragraph.add_run(pure_text)
                            # run.font.color.rgb = RGBColor(0, 0, 255)
                        elif '__BOLD__' in content and '__END_BOLD__' in content:
                            pure_text = content.replace('__BOLD__', '').replace('__END_BOLD__', '')
                            run = paragraph.add_run(pure_text)
                            run.bold = True
                        elif '__YELLOW__' in content and '__END_YELLOW__' in content:
                            pure_text = content.replace('__YELLOW__', '').replace('__END_YELLOW__', '')
                            run = paragraph.add_run(pure_text)
                            run.font.color.rgb = RGBColor(255, 92, 0)
                            if '*' in pure_text:  # Check for italic formatting
                                run.italic = True
                        
                        else:
                            # Handle normal text with potential italic formatting
                            if '*' in content:
                                # Split by asterisks and process each part
                                parts = content.split('*')
                                for i, part in enumerate(parts):
                                    if part:  # Skip empty parts
                                        run = paragraph.add_run(part)
                                        run.italic = (i % 2 == 1)  # Alternate italic formatting
                            else:
                                run = paragraph.add_run(content)
                        # else:
                            # run = paragraph.add_run(content)

                        # Set monospace font and other properties
                        run.font.size = Pt(7)
                        # Check if the text is from txtgrphc (contains non-breaking space)
                        if '\xa0' in content:
                            run.font.name = 'Courier New'
                        else:
                            run.font.name = 'Cambria (Body)'

                        # run.font.name = 'Courier New'
                        paragraph.paragraph_format.space_after = Pt(0)
                        paragraph.paragraph_format.space_before = Pt(0)
                        paragraph.paragraph_format.line_spacing = 1.0
                    
                    elif content_type == 'table':
                        headers, data_rows = content
                        nested_table = main_cell.add_table(rows=len(data_rows) + 1, cols=len(headers))
                        nested_table.style = 'Table Grid'

                        # Add headers
                        for i, header in enumerate(headers):
                            cell = nested_table.cell(0, i)
                            paragraph = cell.paragraphs[0]
                            run = paragraph.add_run(header)
                            run.font.size = Pt(8)
                            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

                        # Add data rows
                        for row_idx, row_data in enumerate(data_rows, start=1):
                            if isinstance(row_data, dict) and row_data.get('merged'):
                                # Handle merged rows
                                merged_content = row_data['content']
                                # Merge all cells in this row
                                first_cell = nested_table.cell(row_idx, 0)
                                first_cell.text = merged_content
                                first_cell.paragraphs[0].runs[0].font.size = Pt(8)
                                first_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT

                                for col_idx in range(1, len(headers)):
                                    cell_to_merge = nested_table.cell(row_idx, col_idx)
                                    first_cell.merge(cell_to_merge)
                            else:
                                # Handle normal rows
                                for col_idx, cell_data in enumerate(row_data):
                                    cell = nested_table.cell(row_idx, col_idx)
                                    paragraph = cell.paragraphs[0]
                                    # run = paragraph.add_run(cell_data)

                                    # Check if the cell contains warning text
                                    if '__RED__' in cell_data and '__END_RED__' in cell_data:
                                        # Remove the tags and create a red-colored run
                                        pure_text = cell_data.replace('__RED__', '').replace('__END_RED__', '')
                                        run = paragraph.add_run(pure_text)
                                        # run.font.color.rgb = RGBColor(255, 0, 0)
                                    else:
                                        # Normal text
                                        run = paragraph.add_run(cell_data)

                                    run.font.size = Pt(8)
                                    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

                        # Add spacing after table
                        main_cell.add_paragraph()

                        

                if not self.no_graphic: 
                            # Add a note about missing graphics
                            paragraph = main_cell.add_paragraph()
                            run = paragraph.add_run("\n\nNo Graphic Reference Found!")
                            run.font.size = Pt(12)
                            run.font.color.rgb = RGBColor(255, 0, 0)
                            self.no_graphic = False

                # Add "End of task card" text
                paragraph = main_cell.add_paragraph()
                run = paragraph.add_run("\n==================================This is the end of the task card==================================")
                run.font.size = Pt(7)
                run.font.name = 'Cambria (Body)'
                paragraph.paragraph_format.space_after = Pt(0)
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.line_spacing = 1.0

                ####
                ##adding footer

                ##to be deleted later
                section = self.doc.sections[0]  # Get the first section
                section.footer_distance = Inches(0.2)
                footer = section.footer
                footer_paragraph = footer.paragraphs[0]

                # Adjust paragraph settings to move image left
                footer_paragraph.paragraph_format.left_indent = Inches(-0.28)  # Negative value moves it left
                footer_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT  # Ensure left alignment

                run = footer_paragraph.add_run()
                run.add_picture('footer_table.png', width=Inches(7.85), height=Inches(2.5))
                ##the end of to be deleted later

                ###end of footer

                self.doc.save(output_name)
                print(f"Task card has been generated and saved as {output_name}")

        except Exception as e:
            print(f"An error occurred: {str(e)}")
            print("This is in creat_nested_table")

    ##this is only for Mpd
    def mpd_content(self,title, smtdesc, mpd, sourceref, amm, list_title, description ):
        text = "\nTitle:  " + title + "\n\nPLANNING NOTE\n\n" + "\nEFFECTIVITY: \n" + title + "\n\n" +smtdesc + "\n\nREFERENCES\n" + "MPD:" + mpd + "\n" + "MRB:" + sourceref + "\n" + "AMM:" + amm + "\n\n" + list_title + "\n" + description + "\n"
        return text

    # ##this one is counting the pages correct.
    def create_header(self, doc, customer,custID, ac_type,ac_regn,mpd_detail1,mpd_detail2):
        # Set margins to 0.5 inches
        section = doc.sections[0]
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
        # section.top_margin = Inches(0.1)

        section.header_distance = Inches(0.3)


        # Create the header tables before the main content
        header = section.header
        header_paragraph = header.paragraphs[0]

        # Add logo and text
        run = header_paragraph.add_run()
        run.add_picture('logo.jpg', width=Inches(0.8))
        run = header_paragraph.add_run("    HONG KONG AIRCRAFT ENGINEERING CO.LTD." + "\n \n")
        run.font.size = Pt(10)  # Set company name to 10pt
        run = header_paragraph.add_run("      TASK CARD")
        run.font.size = Pt(18) #TAS CARD font size

        # First header table
        header_table1 = header.add_table(rows=1, cols=30, width=Inches(8.5))
        header_table1.style = 'Table Grid'
        header_table1.rows[0].height = Inches(0.4)

        # Define merge ranges and headers for the first table
        merge_ranges = [
            (0, 1, "CUSTOMER\n" + customer),  # Merge cells 1-2
            (2, 3, "CUST ID\n" + custID),  # Merge cells 3-4 #customer
            (4, 6, "A/C REGN\n" + ac_type),  # Merge cells 5-7 #registration
            (7, 9, "A/C TYPE\n" + ac_regn),  # Merge cells 8-10 #Model
            (10, 11, "STAGE"),  # Merge cells 11-12
            (12, 14, "TRADE\n" + mpd_detail1),  # Merge cells 13-15
            (15, 17, "JOBTYPE\n" + self.jobtype),  # Merge cells 16-18
            (18, 23, "JOB BOOKING NO"),  # Merge cells 19-24
            (24, 29, "CONTROL SERIAL NO\n" + self.csn)  # Merge cells 25-30
        ]

        # Perform merges and set text for the first table
        hdr_cells = header_table1.rows[0].cells
        for start_idx, end_idx, header_text in merge_ranges:
            merged_cell = hdr_cells[start_idx]
            for i in range(start_idx + 1, end_idx + 1):
                merged_cell = merged_cell.merge(hdr_cells[i])
            merged_cell.text = header_text

        # Second header table
        header_table2 = header.add_table(rows=1, cols=30, width=Inches(8.5))
        header_table2.style = 'Table Grid'
        header_table2.rows[0].height = Inches(0.4)


        # Convert zones list to string for the second table
        zones_str = ', '.join(mpd_detail2) if isinstance(mpd_detail2, list) else str(mpd_detail2)

        # Define merge ranges and headers for the second table
        merge_ranges = [
            (0, 1, "WORKPACK NUMBER\n" + self.workpack_num),  # Merge cells 1-2
            (2, 3, "CHECK\n" + self.check),  # Merge cells 3-4
            (4, 17, "ZONE\n" + zones_str),  # Merge cells 5-18
            (18, 21, "ISSUE NO"),  # Merge cells 19-22
            (22, 23, "PAGE"),  # Merge cells 23-24
            (24, 29, "CARD.NO\n" + self.card_no)  # Merge cells 25-30
        ]

        # Perform merges and set text for the second table
        row2 = header_table2.rows[0].cells
        for start_idx, end_idx, header_text in merge_ranges:
            merged_cell = row2[start_idx]
            for i in range(start_idx + 1, end_idx + 1):
                merged_cell = merged_cell.merge(row2[i])
            merged_cell.text = header_text

            # # Add page number to the "PAGE" cell
            if header_text == "PAGE":
                paragraph = merged_cell.paragraphs[0]
                paragraph.clear()  # Clear any existing text
                run = paragraph.add_run("PAGE \n")

                # Insert { PAGE } field for current page number
                fldChar1 = OxmlElement('w:fldChar')
                fldChar1.set(qn('w:fldCharType'), 'begin')
                instrText1 = OxmlElement('w:instrText')
                instrText1.set(qn('xml:space'), 'preserve')
                instrText1.text = 'PAGE'
                fldChar2 = OxmlElement('w:fldChar')
                fldChar2.set(qn('w:fldCharType'), 'end')

                run._r.append(fldChar1)
                run._r.append(instrText1)
                run._r.append(fldChar2)

                # Add separator (e.g., "of")
                run = paragraph.add_run(" of ")

                # Insert { NUMPAGES } field for total page count
                fldChar3 = OxmlElement('w:fldChar')
                fldChar3.set(qn('w:fldCharType'), 'begin')
                instrText2 = OxmlElement('w:instrText')
                instrText2.set(qn('xml:space'), 'preserve')
                instrText2.text = 'NUMPAGES'
                fldChar4 = OxmlElement('w:fldChar')
                fldChar4.set(qn('w:fldCharType'), 'end')

                run._r.append(fldChar3)
                run._r.append(instrText2)
                run._r.append(fldChar4)

    # ##extracting effectivity crossreference table with msn_number
    def extract_header_rows(self, input_file, msn_num):
        try:
            print("msn_num in extract_header_rows is: ", msn_num)
            #msn_num is empty
            if not msn_num:
                return None

            html_content = SGMLToHTMLConverter().read_header(input_file)
            # print("header_content:",html_content)
            
            if not html_content:
                print("No header data found in the file.")
                return []
            
            # Parse the HTML content
            soup = BeautifulSoup(html_content, 'html.parser')
            
            # Find the effectivity data table
            table = soup.find('table', class_='effectivity-data')
            if not table:
                print("No effectivity data table found in the header.")
                return []
            
            # Extract all rows except the header row
            rows = table.find_all('tr')[1:]  # Skip the header row
            
            result = []
            found_msn = False

            for row in rows:
                cells = row.find_all('td')
                if not cells:
                    continue
                    
                # Extract data from each cell
                row_data = [cell.get_text().strip() for cell in cells]
                
                # # If msn_num is provided, filter for that specific MSN
                if msn_num:
                    # Check if the MSN column (index 7) matches the requested MSN
                    if len(row_data) > 7 and str(row_data[7]) == str(msn_num):
                        found_msn = True
                        return row_data
                    
                else:
                    # If no msn_num provided, collect all rows
                    result.append(row_data)

                # result.append(row_data)
            
            # If msn_num was provided but not found
            # if msn_num and not result:
            if msn_num and not found_msn:
                print(f"MSN {msn_num} not found in the effectivity data.")
                return []  # Return empty list to indicate MSN wasn't found

            # elif not msn_num:
            #     # Print all rows if no specific MSN was requested
            #     print(f"Found {len(result)} rows in total")
            #     for i, row_data in enumerate(result):
            #         if len(row_data) > 7:
                        # print(f"Row {i+1}: MSN {row_data[7]} - {row_data[9]}")
            # print("this is result in extract_header_rows:", result)
            return result
        
        except Exception as e:
            print(f"Error extracting header rows: {str(e)}")
            import traceback
            traceback.print_exc()
            return []

    def set_cell_border(self, cell, border_type, color, size):
        tc = cell._element
        tcPr = tc.get_or_add_tcPr()
        
        for border_name in ['top', 'left', 'bottom', 'right']:
            border = OxmlElement(f'w:{border_name}')
            border.set(qn('w:val'), border_type)
            border.set(qn('w:sz'), str(size * 8))  # size in eighths of a point
            border.set(qn('w:space'), '0')
            border.set(qn('w:color'), color)
            tcPr.append(border)



def process_amm_codes_from_csv(csv_file, sgml_file):
    try:
        df = pd.read_csv(csv_file)
        
        if "AMM Reference" not in df.columns:
            print("Error: CSV file must contain 'AMM Reference' column")
            return
        
        successful_count = 0
        failed_count = 0
        
        for index, row in df.iterrows():
            amm_code = row["AMM Reference"].strip()
            print(f"\nProcessing AMM code: {amm_code}")
            
            try:
                task_card_creator = TableGenerator(sgml_file, amm_code, warning_file, caution_file, effrg)
                task_card_creator.create_nested_table(sgml_file, 324100031, None)
                successful_count += 1
            except Exception as e:
                print(f"Failed to process AMM code {amm_code}: {str(e)}")
                failed_count += 1
                continue
        
        print(f"\nProcessing Summary:")
        print(f"Total AMM codes: {len(df)}")
        print(f"Successfully processed: {successful_count}")
        
    except Exception as e:
        print(f"Error processing CSV file: {str(e)}")


if __name__ == "__main__":
    csv_file = "output_amm_1.csv"  # Your CSV file path
    # sgml_file = "amm_original_txt.txt"
    sgml_file = "sgml_a318-a321/amm.sgm"
    # sgml_file = "single_amm.txt"
    warning_file = "sgml_a318-a321/amm.swe"
    caution_file = "sgml_a318-a321/amm.sce"

    # # csv_file = "a330_less_list.csv"  # Your CSV file path
    # csv_file = "single_list.csv"
    # sgml_file = "allfiles/amm.sgm"
    # # sgml_file = "single_amm.txt"
    # warning_file = "allfiles/warning.swe"
    # caution_file = "allfiles/caution.sce"


    # effrg = '201'
    # effrg = None
    process_amm_codes_from_csv(csv_file, sgml_file)
    