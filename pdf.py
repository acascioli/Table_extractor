import tabula
import pandas as pd

# File path
file_path = "input/GIULIO SRL - 210.pdf"
# file_path = "input/VUS SPA - 000_00216190.pdf"

# Extract tables from the PDF
dfs = tabula.read_pdf(file_path, pages="all", multiple_tables=True)

# Save each table as a CSV file
for i, df in enumerate(dfs):
    df.to_csv(f"output/table_{i}.csv", index=False)
