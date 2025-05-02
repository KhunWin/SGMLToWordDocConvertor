import tkinter as tk
from tkinter import filedialog, messagebox
from main_py import AMMProcessor
from Extract_MPDToAMMRef import MPDProcessor
from Extract_MPDToAMMRef import MPDProcessor, KeyNotFoundException  # Add KeyNotFoundException to import
import os

class UI(tk.Tk):
    def __init__(self):
        super().__init__()

        self.all_file = True
        self.title("Capability Automation Tool")
        self.geometry("500x900")  # Increased window size to accommodate results

        # self.configure(bg="light green")  # Set the background color to light green

        self.file_paths = {}  # Dictionary to store file paths
        self.file_labels = {}  # Dictionary to store labels for each file

        self.create_buttons()
        
        self.result_text = tk.Text(self, height=10, bg="black",fg="white")  # Text widget for results (print)
        self.result_text.pack(pady=10)

        
        self.create_execute_button() #execute button

        self.create_reset_button()  # Reset button


    def create_execute_button(self):
        button = tk.Label(self, text="Click Execute to create an excel document")
        button.pack(pady=5)
        button = tk.Button(self, text="Execute", command=self.create_taskcard_file, bg="light blue",activeforeground="blue")
        button.pack(pady=5)  # Add some padding for better layout
        self.result_text.insert(tk.END, "The program is running.....\n","black")

        # Bind events for hover effect
        button.bind("<Enter>", self.on_enter)
        button.bind("<Leave>", self.on_leave)

    def on_enter(self, event):
        event.widget.config(bg="yellow")

    def on_leave(self, event):
        event.widget.config(bg="blue")
        

    def create_buttons(self):
        self.create_button_with_label("Choose MPD_SGM File(.sgm)", self.choose_file1, 1)
        self.create_button_with_label("Choose MPD List(.csv)", self.choose_file2, 2)
        self.create_button_with_label("Select AMM/Warning/Caution Folder", self.choose_folder, 3)

        # Add an input text box for "Effectivity"
        self.effectivity_label = tk.Label(self, text="Enter FSN (optional):")
        self.effectivity_label.pack(pady=5)
        self.effectivity_entry = tk.Entry(self, width=40)
        self.effectivity_entry.pack(pady=5)

        # Add an input text box for "MSN Number"
        self.msn_label = tk.Label(self, text="Enter MSN Number(Optional):")
        self.msn_label.pack(pady=5)
        self.msn_entry = tk.Entry(self, width=40)
        self.msn_entry.pack(pady=5)

        self.workpack_label = tk.Label(self, text="Enter Workpack Number(Optional):")
        self.workpack_label.pack(pady=5)
        self.workpack_entry = tk.Entry(self, width=40)
        self.workpack_entry.pack(pady=5)

        # Add an input text box for "Check"
        self.check_label = tk.Label(self, text="Enter Check(Optional):")
        self.check_label.pack(pady=5)
        self.check_entry = tk.Entry(self, width=40)
        self.check_entry.pack(pady=5)

        # # Add radio buttons for MPD-only option
        self.mpd_var = tk.StringVar(value="no")
        self.mpd_label = tk.Label(self, text="Process MPD Only:")
        self.mpd_label.pack(pady=5)
        
        self.mpd_frame = tk.Frame(self)
        self.mpd_frame.pack(pady=5)
        
        tk.Radiobutton(self.mpd_frame, text="Yes", variable=self.mpd_var, value="yes").pack(side=tk.LEFT, padx=10)
        tk.Radiobutton(self.mpd_frame, text="No", variable=self.mpd_var, value="no").pack(side=tk.LEFT, padx=10)


    def create_button_with_label(self, button_text, command, index):

        button = tk.Button(self, text=button_text, command=command, 
                           bg="light blue",
                           activeforeground="blue")  ##change the color of the button here
                        
        button.pack(pady=5)  # Add some padding for better layout

        label = tk.Label(self, text="")
        label.pack(pady=5)
        self.file_labels[index] = label

        button.bind("<Enter>", self.on_enter)
        button.bind("<Leave>", self.on_leave)

    def choose_file1(self):
        file_path = filedialog.askopenfilename()
        if file_path:
            file_extension = file_path.lower()
            if not (file_extension.endswith('.sgm') or file_extension.endswith('.SGM')):
                messagebox.showerror("Invalid File", "Please choose a .sgm file for Warning.")
            else:
                self.file_paths[1] = file_path
                self.file_labels[1].config(text=file_path.split("/")[-1])

    def choose_file2(self):
        file_path = filedialog.askopenfilename()
        if file_path:
            file_extension = file_path.lower()
            if not (file_extension.endswith('.csv') or file_extension.endswith('.CSV')):
                messagebox.showerror("Invalid File", "Please choose a .csv file for Warning.")
            else:
                self.file_paths[2] = file_path
                self.file_labels[2].config(text=file_path.split("/")[-1])

    def choose_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            # Find required files in the selected folder
            sgm_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.sgm')]
            swe_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.swe')]
            sce_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.sce')]

            # Validate file counts
            valid = True
            if len(sgm_files) != 1:
                messagebox.showerror("Error", "Folder must contain exactly one .sgm file")
                valid = False
            if len(swe_files) != 1:
                messagebox.showerror("Error", "Folder must contain exactly one .swe file")
                valid = False
            if len(sce_files) != 1:
                messagebox.showerror("Error", "Folder must contain exactly one .sce file")
                valid = False

            if valid:
                self.file_paths[3] = os.path.join(folder_path, sgm_files[0])
                self.file_paths[4] = os.path.join(folder_path, swe_files[0])
                self.file_paths[5] = os.path.join(folder_path, sce_files[0])
                self.file_labels[3].config(text=folder_path)

    #reset button
    def create_reset_button(self):
        button = tk.Button(self, text="Reset", command=self.reset, bg="light blue")
        button.pack(pady=5)  # Add some padding for better layout
        # Add hover effects for consistency
        # button.bind("<Enter>", self.on_enter)
        # button.bind("<Leave>", self.on_leave)

    def reset(self):
        self.file_paths.clear()
        for label in self.file_labels.values():
            label.config(text="")  # Clear file labels
        # self.user_input.delete(0, tk.END)  # Clear user input
        self.effectivity_entry.delete(0, tk.END)  # Clear effectivity input
        self.msn_entry.delete(0, tk.END)  # Clear MSN number input
        self.workpack_entry.delete(0, tk.END)  # Clear workpack input
        self.check_entry.delete(0, tk.END)  # Clear check input

        if hasattr(self, 'result_text'):
            self.result_text.delete(1.0, tk.END)  # Clear result text
            self.result_text.tag_configure("black", foreground="white")  # Configure tag for white text on black background
            self.result_text.insert(tk.END, "Reset complete. Please reselect files and execute.\n", "black")
            self.result_text.see(tk.END)  # Scroll to the end to show the message

            print("Reset complete. Please reselect files and execute.")

    def create_taskcard_file(self):
        MPD_sgml_file = self.file_paths.get(1)
        MPD_list_file = self.file_paths.get(2)
        AMM_sgml_file = self.file_paths.get(3)
        Warning_file = self.file_paths.get(4)
        Caution_file = self.file_paths.get(5)
  

        if not all([MPD_sgml_file, MPD_list_file, AMM_sgml_file, Warning_file, Caution_file]):
            
            self.result_text.insert(tk.END, "Cannot proceed to create an excel document. Make sure to select all files.\n", "red")
            self.result_text.insert(tk.END, "Please select all files before executing.\n","black")
            all_file = False
            return 
        
         # Get the value from the Effectivity input box
        effectivity = self.effectivity_entry.get().strip() or None  # Use None if the box is empty
        # Get the value from the MSN Number input box
        msn_num = self.msn_entry.get().strip() or None  # Use None if the box is empty
        # print("MSN Number in UI:", msn_num)
        workpack_num = self.workpack_entry.get().strip() or ''  # New variable
        check = self.check_entry.get().strip() or ''  # New variable
        
        # Clear the result text widget before displaying new results
        self.result_text.delete(1.0, tk.END)
        
        # Display file paths in the result text widget
        self.result_text.insert(tk.END, f"MPD reference file: {MPD_sgml_file}\n", "black")
        self.result_text.insert(tk.END, f"MPD list file: {MPD_list_file}\n","black")
        folder_path = os.path.dirname(AMM_sgml_file)
        self.result_text.insert(tk.END, f"AMM/Warning/Caution folder: {folder_path}\n", "black")
        self.result_text.insert(tk.END, f"Effectivity: {effectivity}\n", "black")

        # Perform calculations here using the file paths
        try:
            self.result_text.insert(tk.END, "Creating Task cards...\n", "black")
            self.result_text.insert(tk.END, "Extracting AMM reference from MPD file...\n", "black")
            mpd_processor = MPDProcessor(MPD_sgml_file, MPD_list_file)

            try:
                mpd_df, task_details_map = mpd_processor.process_csv()
                self.result_text.insert(tk.END, "Extracting AMM reference finished.\n", "black")
                
                # Display any not found keys but continue processing
                if "NOT_FOUND" in str(mpd_df['AMM Reference'].values):
                    self.result_text.tag_configure("red", foreground="red")
                    not_found = mpd_df[mpd_df['AMM Reference'].str.contains("NOT_FOUND", na=False)]
                    error_msg = "\nWARNING: The following MPD references were not found:\n"
                    for idx, row in not_found.iterrows():
                        error_msg += f"{row['MPD Reference']} is not found. Check your MPD number again!\n"
                    self.result_text.insert(tk.END, error_msg, "red")
                
                # Continue with processing valid entries
                self.result_text.insert(tk.END, "Converting to word documents...\n", "black")
                # Filter out not found entries for processing
                valid_df = mpd_df[~mpd_df['AMM Reference'].str.contains("NOT_FOUND", na=False)]
                if not valid_df.empty:
                    only_mpd = self.mpd_var.get() == "yes"
                    amm_processor = AMMProcessor(Warning_file, Caution_file, effectivity,msn_num, only_mpd, workpack_num,check)
                    # Make sure task_details_map is a dictionary and not a string
                    if isinstance(task_details_map, dict):
                        amm_processor.process_files(valid_df, AMM_sgml_file, task_details_map)
                    else:
                        raise TypeError(f"task_details_map must be a dictionary, not {type(task_details_map)}")
                    # amm_processor.process_files(valid_df, AMM_sgml_file,task_details_map)
                self.result_text.insert(tk.END, "Converting to word documents finished.\n", "black")

            except Exception as e:
                self.result_text.tag_configure("red", foreground="red")
                self.result_text.insert(tk.END, f"\nError during processing: {str(e)}\n", "red")
                return

        except Exception as e:
            self.result_text.tag_configure("red", foreground="red")
            self.result_text.insert(tk.END, f"\nError: {str(e)}\n", "red")
            print(f"Error in extracting task cards: {e}")

        finally:
            try:
                self.result_text.insert(tk.END, "Creating Task Cards finished.\n", "black")
                output_file_path = os.path.join(os.getcwd(),"")
                self.result_text.insert(tk.END, f"Output file created at: {output_file_path}\n", "black")
                print("Creating Task Cards finished running.")
            except Exception as e:
                self.result_text.insert(tk.END, "Unsuccessful execution.\n", "black")
                print("Unsuccessful execution. Error: ", e)
    
if __name__ == "__main__":
    app = UI()
    app.mainloop()
