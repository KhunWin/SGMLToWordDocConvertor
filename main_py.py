from convert_v2 import TableGenerator
from Extract_MPDToAMMRef import MPDProcessor
import pandas as pd

class AMMProcessor:
    def __init__(self, warning_file=None, caution_file=None, effectivity=None, msn_num=None, only_mpd=False,workpack_num=None, check=None):
        self.warning_file = warning_file
        self.caution_file = caution_file
        self.effectivity = effectivity
        self.msn_num = msn_num
        self.only_mpd = only_mpd
        self.workpack_num = workpack_num
        self.check = check


    def process_files(self, df, sgml_file, task_details_map):
        try:
            if "AMM Reference" not in df.columns:
                print("Error: DataFrame must contain 'AMM Reference' column")
                return
            
            successful_count = 0
            failed_count = 0
            failed_codes = []  # Track failed AMM codes

            for index, row in df.iterrows():
                mpd_ref = row["MPD Reference"].strip()
                amm_code = row["AMM Reference"].strip()
                csn_code = row["CSN"].strip()
                card_no = row["CardNo"].strip()
                print(f"\nProcessing MPD reference: {mpd_ref}, AMM code: {amm_code}")
                print(f"CSN code: {csn_code}, CardNo: {card_no}")
                # print("task_details_map:", task_details_map)
                
                # Get the specific task details for this MPD reference
                task_details = {}  # Initialize with empty dict
                taskcode = None
                if mpd_ref in task_details_map and isinstance(task_details_map[mpd_ref], dict):
                    task_details = task_details_map[mpd_ref]
                    taskcode = task_details.get('taskcode')
                else:
                    print(f"Warning: No valid task details found for MPD reference {mpd_ref}")
                
                try:
                    task_card_creator = TableGenerator(
                        sgml_file, 
                        amm_code, 
                        self.warning_file, 
                        self.caution_file,
                        self.effectivity,
                        csn_code,
                        card_no,
                        self.workpack_num,self.check, taskcode
                    )
                    # Make sure task_details is a dictionary before passing it
                    if isinstance(task_details, dict):
                        task_card_creator.create_nested_table(sgml_file, task_details, self.msn_num, self.only_mpd)
                    else:
                        print(f"Warning: task_details for {mpd_ref} is not a dictionary. Type: {type(task_details)}")
                        # Create an empty dictionary if task_details is not a dictionary
                        task_card_creator.create_nested_table(sgml_file, {}, self.msn_num, self.only_mpd)
                    
                    successful_count += 1
                    
                except Exception as e:
                    error_msg = f"Failed to process AMM code {amm_code}: {str(e)}"
                    print(error_msg)
                    failed_codes.append(error_msg)
                    failed_count += 1
                    continue
            
            print(f"\nProcessing Summary:")
            print(f"Total AMM codes: {len(df)}")
            print(f"Successfully processed: {successful_count}")
            
        except Exception as e:
            # print(f"Error processing data: {str(e)}")
            raise e
if __name__ == "__main__":

    # Initialize MPDProcessor and get DataFrame
    mpd_processor = MPDProcessor('testingA330/MPD_File_A330.sgm', 'TestingA330/A330_mpd_list.csv') ##mpd_list.csv is a list of mpd references without amm references. 
    mpd_df = mpd_processor.process_csv()

    processor = AMMProcessor(
        warning_file="testingA330/sgml_a330/warning.swe",
        caution_file="testingA330/sgml_a330/caution.sce"
    )
    
    # processor.process_files(
    #     mpd_df,
    #     sgml_file="testingA330/sgml_a330/AMM.sgm"

    # )