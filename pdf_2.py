import fitz
import pandas as pd
import pathlib as plib
import customtkinter
import os

base_file = plib.Path(__file__).parents[0]
input_folder = plib.Path(base_file, "input")
input_folder.mkdir(exist_ok=True)
output_folder = plib.Path(base_file, "output")
output_folder.mkdir(exist_ok=True)


class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.geometry("400x500")

        # Button to open file dialog
        self.file_button = customtkinter.CTkButton(
            self, text="Select File", command=self.select_file
        )
        self.file_button.pack(padx=20, pady=(50, 0))

        # Label to display the selected file name
        self.file_label = customtkinter.CTkLabel(self, text="No file selected")
        self.file_label.pack(padx=20, pady=10)

        # Entry for page selection
        self.page_label = customtkinter.CTkLabel(
            self, text="Enter pages (e.g., 1-5, 8, 11):"
        )
        self.page_label.pack(padx=20, pady=5)

        self.page_entry = customtkinter.CTkEntry(self)
        self.page_entry.pack(padx=20, pady=5)

        # Checkbox for merge option
        self.merge_var = customtkinter.StringVar(value="off")
        self.merge_checkbox = customtkinter.CTkCheckBox(
            self,
            text="Merge Output",
            variable=self.merge_var,
            onvalue="on",
            offvalue="off",
        )
        self.merge_checkbox.pack(padx=20, pady=(10, 50))

        # Spacer frame
        self.spacer_frame = customtkinter.CTkFrame(self, height=20, width=400)
        self.spacer_frame.pack()

        self.button = customtkinter.CTkButton(
            self, text="Process", command=self.button_callbck
        )
        self.button.pack(padx=20, pady=(50, 0))

        # Label to display the processing status
        self.process_label = customtkinter.CTkLabel(self, text="Waiting for file...")
        self.process_label.pack(padx=20, pady=10)

    def select_file(self):
        # Open the file dialog to select a file
        self.file_path = customtkinter.filedialog.askopenfilename(title="Select a file")
        if self.file_path:
            self.file_name = os.path.basename(self.file_path)
            # Update the label with the selected file name
            self.file_label.configure(text=self.file_name)
        else:
            # Update the label to show that no file was selected
            self.file_label.configure(text="No file selected")

    def extract_tables(self, page_input):
        self.process_label.configure(text="Processing...")
        file_folder = "".join(self.file_name.split(".")[:-1])
        file_folder = plib.Path(output_folder, file_folder)
        file_folder.mkdir(exist_ok=True)
        try:
            doc = fitz.open(self.file_path)
            if page_input:
                for i in page_input:
                    tabs = doc[int(i)].find_tables()

                    for j, tab in enumerate(tabs):

                        df = tab.to_pandas()
                        df = df.replace("\n", " ", regex=True)
                        # df.to_csv(f"output/page_{i}_table_{j}.csv", index=False)
                        file_output = plib.Path(file_folder, f"page_{i}_table_{j}.xlsx")
                        df.to_excel(file_output, index=False)
            else:
                for i, page in enumerate(doc):

                    tabs = page.find_tables()

                    for j, tab in enumerate(tabs):

                        df = tab.to_pandas()
                        df = df.replace("\n", " ", regex=True)
                        # df.to_csv(f"output/page_{i}_table_{j}.csv", index=False)
                        file_output = plib.Path(file_folder, f"page_{i}_table_{j}.xlsx")
                        df.to_excel(file_output, index=False)
        except Exception as e:
            self.process_label.configure(text="❌ Something went wrong...")
            print(e)

    def parse_page_input(self, page_input):
        # Split the input string on commas to handle multiple entries
        entries = page_input.split(",")
        page_numbers = set()  # Use a set to avoid duplicate pages

        for entry in entries:
            entry = entry.strip()  # Remove whitespace
            if "-" in entry:
                # Handle ranges, e.g., '2-5'
                start, end = entry.split("-")
                page_numbers.update(range(int(start), int(end) + 1))
            else:
                # Handle single page numbers
                page_numbers.add(int(entry))

        return sorted(page_numbers)  # Return a sorted list of unique page numbers

    def button_callbck(self):
        page_input = self.page_entry.get()
        try:
            pages_to_process = self.parse_page_input(page_input)
            print(f"Pages to process: {pages_to_process}")
            merge = self.merge_var.get() == "on"
            print(f"Merge output: {merge}")
            # Add logic here to process the file using the extracted pages and merge option
        except ValueError as e:
            print("Invalid input for pages. Please enter valid page numbers or ranges.")
            page_input = None

        self.extract_tables(page_input)
        self.process_label.configure(text="✅ Completed!")


app = App()
app.mainloop()
