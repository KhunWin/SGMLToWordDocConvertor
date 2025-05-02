import re
import csv
import pandas as pd

# Define custom exception
class KeyNotFoundException(Exception):
    pass

class MPDProcessor:
    def __init__(self, sgm_file, input_csv):
        self.sgm_file = sgm_file
        self.input_csv = input_csv
        self.data = None
        self._load_sgm_file()

    def _load_sgm_file(self):
        """Load the SGM file content"""
        try:
            with open(self.sgm_file, 'r', encoding='utf-8') as file:
                self.data = file.read()
        except Exception as e:
            print(f"Error loading SGM file: {str(e)}")
            self.data = None


    def extract_task_details(self, input_key):
        if not self.data:
            return "SGM file not loaded properly"

        try:
            patterns = [
                r'<SMTASK.*?KEY="EN(.*?)-S".*?</SMTASK>',
                r'<SMTASKINV.*?<SMTASK.*?KEY="EN(.*?)-S".*?</SMTASKINV>'
            ]
            
            for pattern in patterns:
                matches = re.finditer(pattern, self.data, re.DOTALL)
                for match in matches:
                    key = match.group(1)
                    if key == input_key:
                        full_block = match.group(0)

                        # Extract MPD from key
                        mpd_number = key.split('-')
                        mpd = f"{mpd_number[0]}-{mpd_number[1]}-{mpd_number[2]}" if len(mpd_number) >= 3 else key
                        
                        # Extract individual tags
                        skill = re.search(r'<SKILL>(.*?)</SKILL>', full_block)
                        taskcode = re.search(r'<TASKCODE>(.*?)</TASKCODE>', full_block)
                        zones = re.findall(r'<ZONE>(.*?)</ZONE>', full_block)
                        title = re.search(r'<TITLE>(.*?)</TITLE>', full_block)
                        smtdesc = re.search(r'<SMTDESC><PARA>(.*?)</PARA>', full_block)
                        
                        # Extract sourceref
                        sourcerefs = re.findall(r'<REFEXT REFLOC="(.*?)" REFMAN="MRB">', full_block)
                        sourceref = ' '.join([f"MRB {ref}" for ref in sourcerefs]) if sourcerefs else 'Not found'
                        
                        # Extract AMM reference
                        amm = re.search(r'<PROCBOX><REFEXT.*?>(.*?)</REFEXT>', full_block)
                        
                        # Get taskcode value for checking inspection types
                        taskcode_value = taskcode.group(1) if taskcode else ''
                        list_info = ''
                        description = ''
                        
                        # Dictionary mapping taskcode to inspection types and patterns
                        inspection_types = {
                            'GVI': {
                                'name': 'GVI - General Visual Inspection',
                                'title_pattern': r'<PARA><EM ROLE="BOLD">Inspection - General Visual: \(GVI\)</EM>\s*</PARA>',
                                'desc_pattern': r'<PARA><EM ROLE="BOLD">Inspection - General Visual: \(GVI\)</EM>\s*</PARA>\s*<PARA>(.*?)</PARA>'
                            },
                            # 'DET': {
                            #     'name': 'DET - Detailed Inspection',
                            #     'title_pattern': r'<PARA><EM ROLE="BOLD">Inspection - Detailed: \(DET\)</EM>\s*</PARA>',
                            #     'desc_pattern': r'<PARA><EM ROLE="BOLD">Inspection - Detailed: \(DET\)</EM>\s*</PARA>\s*<PARA>(.*?)</PARA>'
                            # },
                            'DET': {
                                'name': 'DET - Detailed Inspection',
                                'title_patterns': [
                                    r'<PARA><EM ROLE="BOLD">Inspection - Detailed: \(DET\)</EM>\s*</PARA>',
                                    r'<PARA><REVST><EM ROLE="BOLD">Inspection - Detailed: \(DET\)</EM>\s*<REVEND></PARA>'
                                ],
                                'desc_patterns': [
                                    r'<PARA><EM ROLE="BOLD">Inspection - Detailed: \(DET\)</EM>\s*</PARA>\s*<PARA>(.*?)</PARA>',
                                    r'<PARA><REVST><EM ROLE="BOLD">Inspection - Detailed: \(DET\)</EM>\s*<REVEND></PARA>\s*<PARA>(.*?)</PARA>'
                                ]
                            },
                            'SDI': {
                                'name': 'SDI - Special Detailed Inspection',
                                'title_pattern': r'<PARA><EM ROLE="BOLD">Inspection - Special Detailed: \(SDI\)</EM>\s*</PARA>',
                                'desc_pattern': r'<PARA><EM ROLE="BOLD">Inspection - Special Detailed: \(SDI\)</EM>\s*</PARA>\s*<PARA>(.*?)</PARA>'
                            },
                            'VCK': {
                                'name': 'VCK - Visual Check',
                                'title_pattern': r'<PARA><EM ROLE="BOLD">Visual Check \(VCK\)</EM>\s*</PARA>',
                                'desc_pattern': r'<PARA><EM ROLE="BOLD">Visual Check \(VCK\)</EM>\s*</PARA>\s*<PARA>(.*?)</PARA>'
                            },

                        }
                        

                        if taskcode_value in inspection_types:
                            inspection = inspection_types[taskcode_value]
                            list_info = inspection['name']
                            
                            # Try each pattern pair until we find a match
                            description = ''
                            if 'title_patterns' in inspection:  # For DET with multiple patterns
                                for title_pattern, desc_pattern in zip(inspection['title_patterns'], inspection['desc_patterns']):
                                    title_match = re.search(title_pattern, self.data)
                                    if title_match:
                                        desc_match = re.search(desc_pattern, self.data)
                                        if desc_match:
                                            description = desc_match.group(1).strip()
                                            break
                            else:  # For other inspection types with single pattern
                                title_match = re.search(inspection['title_pattern'], self.data)
                                if title_match:
                                    desc_match = re.search(inspection['desc_pattern'], self.data)
                                    if desc_match:
                                        description = desc_match.group(1).strip()
                        
                        # Prepare results
                        result = {
                            'mpd': mpd,
                            'skill': skill.group(1) if skill else 'Not found',
                            'taskcode': taskcode.group(1) if taskcode else 'Not found',
                            'zones': zones if zones else ['Not found'],
                            'title': title.group(1) if title else 'Not found',
                            'smtdesc': smtdesc.group(1) if smtdesc else 'Not found',
                            'sourceref': sourceref,
                            'amm': amm.group(1) if amm else 'Not found',
                            'list': list_info,
                            'description': description
                        }
                        return result
            
            raise KeyNotFoundException(f"{input_key} is not found. Check your MPD number again!")
            
        except Exception as e:
            print(f"Error extracting task details: {str(e)}")
            return None
            
    def extract_text_from_procbox(self, input_key):
        """Extract text from PROCBOX for a given key"""
        if not self.data:
            return "SGM file not loaded properly"

        smtask_pattern = re.compile(r'<SMTASK.*?KEY="EN(.*?)-S".*?</SMTASK>', re.DOTALL)
        smtask_found = False
        
        for match in smtask_pattern.finditer(self.data):
            key = match.group(1)
            if key == input_key:
                smtask_found = True
                return self._process_match(match.group(0))
        
        if not smtask_found:
            smtaskinv_pattern = re.compile(r'<SMTASKINV.*?<SMTASK.*?KEY="EN(.*?)-S".*?</SMTASKINV>', re.DOTALL)
            
            for match in smtaskinv_pattern.finditer(self.data):
                key = match.group(1)
                if key == input_key:
                    return self._process_match(match.group(0))
        
        print(f"This key {input_key} is not found in the SGML file. Please check the key.")

        # Raise custom exception instead of returning a string
        raise KeyNotFoundException(f"{input_key} is not found. Check your MPD number again!")

    def _process_match(self, full_block):
        """Process matched block to extract content"""
        procbox_pattern = re.compile(r'<PROCBOX>(.*?)</PROCBOX>', re.DOTALL)
        procbox_match = procbox_pattern.search(full_block)
        
        if procbox_match:
            procbox_content = procbox_match.group(1).strip()
            
            refext_pattern = re.compile(r'<REFEXT.*?>(.*?)</REFEXT>', re.DOTALL)
            refext_match = refext_pattern.search(procbox_content)
            
            if refext_match:
                return refext_match.group(1).strip()
            else:
                return procbox_content
        else:
            return "PROCBOX content not found"

    def process_csv(self):
        """Process CSV file and return DataFrame"""
        try:
            keys = []
            results = []
            not_found_keys = []  # Track keys that weren't found
            task_details_map = {}  # Map keys to their task details
            csn_list = []  # New list for CSN
            cardno_list = []  # New list for CardNo


            with open(self.input_csv, 'r', newline='') as infile:
                reader = csv.reader(infile)
                header = next(reader)  # Skip the header row
                
                for row in reader:
                    if row:  # Check if the row is not empty
                        input_key = row[1] ##change the col here
                        csn = row[2]  # CSN
                        cardno = row[3]  # CardNo
                        print(f"Processing key: {input_key}")
                        # keys.append(input_key)
                        # result = self.extract_text_from_procbox(input_key)
                        # results.append(result)
                        try:
                            result = self.extract_text_from_procbox(input_key)
                            keys.append(input_key)
                            results.append(result)

                            csn_list.append(csn)  # Add CSN
                            cardno_list.append(cardno)  # Add CardNo

                            task_details = self.extract_task_details(input_key)
                            # print("this is in mpdtoammref:", task_details)
                            if task_details:
                                task_details_map[input_key] = task_details
                        except KeyNotFoundException as e:
                            not_found_keys.append(str(e))
                            # Still add the key to the DataFrame, but with an error message
                            keys.append(input_key)
                            results.append(f"NOT_FOUND: {input_key}")

                            csn_list.append(csn)  # Add CSN even for not found keys
                            cardno_list.append(cardno)  # Add CardNo even for not found keys

            if not keys:  # If no valid keys were processed
                raise KeyNotFoundException("No valid MPD references were found to process.")

            # Create a DataFrame
            df = pd.DataFrame({
                'MPD Reference': keys,
                'AMM Reference': results,
                'CSN': csn_list,
                'CardNo': cardno_list
            })
            

            print("this is df", df['CSN'])
            if not_found_keys:
                raise KeyNotFoundException("\n".join(not_found_keys))
            return df, task_details_map
            
        except Exception as e:
            
            if isinstance(e, KeyNotFoundException) and 'df' in locals():
                return df, task_details_map
            raise e


    def save_to_csv(self, output_file):
        """Save processed data to CSV file"""
        df = self.process_csv()
        if df is not None:
            # df.to_csv(output_file, index=False)
            print(f"\nResults saved to {output_file}")

def main():
    # Example usage
    sgm_file = 'TestingA330/MPD_File_A330.sgm'
    # sgm_file = 'MPD_A318-321.sgm'
    # input_csv = 'mpd_list.csv'
    input_csv ='TestingA330/A330_MPD_list.csv'

    processor = MPDProcessor(sgm_file, input_csv)
    
    # Option 1: Get DataFrame
    # df = processor.process_csv()
    processor.process_csv()
    processor.save_to_csv('signle__list_second.csv')

    
if __name__ == "__main__":
    main()
