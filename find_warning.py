from bs4 import BeautifulSoup
import os, re, html

class WarningHandler:
    def __init__(self, warning_swe_path=None, caution_swe_path=None):
        self.warning_texts = {}
        self.caution_texts = {}
        if warning_swe_path:
            self.warning_texts = self.load_warning_texts(warning_swe_path)
        if caution_swe_path:
            self.caution_texts = self.load_caution_texts(caution_swe_path)

    def load_warning_texts(self, warning_swe_path):
        """Load warning texts from warning SWE file"""
        warning_dict = {}
        if warning_swe_path and os.path.exists(warning_swe_path):
            try:
                with open(warning_swe_path, 'r', encoding='utf-8') as file:
                    content = file.read()
                    
                    warnings = re.finditer(r'<!ENTITY\s+(W\.[A-Z0-9]+)\s+"([^"]+)">', content)
                    
                    for match in warnings:
                        warning_id = match.group(1)
                        warning_text = match.group(2)
                        
                        warning_soup = BeautifulSoup(warning_text, 'html.parser')
                        formatted_text = self._format_content(warning_soup)
                        warning_dict[warning_id] = formatted_text
                        
            except Exception as e:
                print(f"Error loading warning texts from {warning_swe_path}: {str(e)}")
        else:
            print(f"Warning SWE file not found: {warning_swe_path}")
        
        
        return warning_dict

    def load_caution_texts(self, caution_swe_path):
        """Load caution texts from caution SWE file"""
        caution_dict = {}
        if caution_swe_path and os.path.exists(caution_swe_path):
            try:
                with open(caution_swe_path, 'r', encoding='utf-8') as file:
                    content = file.read()
                    
                    cautions = re.finditer(r'<!ENTITY\s+(C\.[A-Z0-9]+)\s+"([^"]+)">', content)
                    
                    for match in cautions:
                        caution_id = match.group(1)
                        caution_text = match.group(2)
                        
                        caution_soup = BeautifulSoup(caution_text, 'html.parser')
                        formatted_text = self._format_content(caution_soup)
                        caution_dict[caution_id] = formatted_text
                        
            except Exception as e:
                print(f"Error loading caution texts from {caution_swe_path}: {str(e)}")
        else:
            print(f"Caution SWE file not found: {caution_swe_path}")
        
        return caution_dict

    def _format_content(self, soup):
        """Helper method to format content consistently for both warnings and cautions"""
        formatted_text = []
        
        for element in soup.children:
            if element.name == 'para':
                formatted_text.append(element.get_text().strip())
            elif element.name == 'unlist':
                # Process unordered list
                for unlitem in element.find_all('unlitem'):
                    para = unlitem.find('para')
                    if para:
                        formatted_text.append(f"- {para.get_text().strip()}")
        
        return '\n'.join(formatted_text)


    def process_warnings(self, html_content):
        """Extract warnings from HTML content that appear before h1"""
        if isinstance(html_content, str):
            soup = BeautifulSoup(html_content, 'html.parser')
        else:
            soup = html_content

        warning_texts = []
        h1 = soup.find('h1')
        if h1:
            # Get all content before h1
            warnings = []
            # Find all warnings before h1 and reverse the list to maintain original order
            current = h1.find_previous('warning')
            while current:
                warnings.insert(0, current)  # Insert at beginning to maintain order
                current = current.find_previous('warning')
        else:
            warnings = soup.find_all('warning')
        
        for warning in warnings:
            warning_text = warning.get_text().strip()
            warning_id = warning_text.strip('&;')
            
            if warning_id in self.warning_texts:
                warning_content = self.warning_texts[warning_id]
                warning_texts.append(f"WARNING: {warning_content}")
            else:
                print(f"Warning ID {warning_id} not found in warning_texts dictionary")

        return "\n\n".join(warning_texts) if warning_texts else ""

    def process_cautions(self, html_content):
        """Extract cautions from HTML content that appear before h1"""
        if isinstance(html_content, str):
            soup = BeautifulSoup(html_content, 'html.parser')
        else:
            soup = html_content

        caution_texts = []
        
        h1 = soup.find('h1')
        if h1:
            # Get all content before h1
            cautions = []
            # Find all cautions before h1 and reverse the list to maintain original order
            current = h1.find_previous('caution')
            while current:
                cautions.insert(0, current)  # Insert at beginning to maintain order
                current = current.find_previous('caution')
        else:
            cautions = soup.find_all('caution')

        for caution in cautions:
            caution_text = caution.get_text().strip()
            caution_id = caution_text.strip('&;')
            
            if caution_id in self.caution_texts:
                caution_content = self.caution_texts[caution_id]
                caution_texts.append(f"CAUTION: {caution_content}")
            else:
                print(f"Caution ID {caution_id} not found in caution_texts dictionary")
        
        return "\n\n".join(caution_texts) if caution_texts else ""