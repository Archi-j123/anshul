import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import matplotlib.pyplot as plt
from tkinter.font import Font
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.pdfgen import canvas
import os

# Function to choose the file path
def select_file():
    global excel_data, file_path
    file_path = filedialog.askopenfilename(
        title="Select an Excel File",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if file_path:
        excel_data = pd.ExcelFile(file_path)
        process_excel_data()

# Function to process the Excel file
def process_excel_data():
    global sheet_results
    tree.delete(*tree.get_children())
    sheet_results = []

    for sheet_name in excel_data.sheet_names:
        sheet_data = pd.read_excel(file_path, sheet_name=sheet_name)
        
        if 'Final_result' in sheet_data.columns:
            final_result_data = sheet_data['Final_result'].dropna()
            if not final_result_data.empty:
                unique_values = final_result_data.unique()
                if len(unique_values) == 1:
                    status = unique_values[0]
                else:
                    status = "PASS and FAIL"
            else:
                status = "No 'Final_result' data"
            filtered_data = sheet_data.drop(columns=['Final_result'])
        else:
            filtered_data = sheet_data
            status = "No 'Final_result' column"

        all_values = filtered_data.values.flatten().tolist()
        pass_count = sum(1 for value in all_values if str(value).strip().upper() == "PASS")
        fail_count = sum(1 for value in all_values if str(value).strip().upper() == "FAIL")
        
        item_id = tree.insert("", "end", values=(sheet_name, pass_count, fail_count, status))
        if status == "PASS":
            tree.item(item_id, tags=("pass",))
        elif status == "FAIL":
            tree.item(item_id, tags=("fail",))
        
        sheet_results.append((sheet_name, pass_count, fail_count))

# Function to display the bar graph
def display_graph(show=True):
    if sheet_results:
        sheet_names = [result[0] for result in sheet_results]
        pass_counts = [result[1] for result in sheet_results]
        fail_counts = [result[2] for result in sheet_results]

        x = range(len(sheet_names))
        fig, ax = plt.subplots(figsize=(10, 6))

        bar_width = 0.35
        ax.bar(x, pass_counts, bar_width, label="PASS", color='g')
        ax.bar([p + bar_width for p in x], fail_counts, bar_width, label="FAIL", color='r')

        ax.set_xlabel('Sheets')
        ax.set_ylabel('Count')
        ax.set_title('PASS/FAIL Counts per Sheet')
        ax.set_xticks([p + bar_width / 2 for p in x])
        ax.set_xticklabels(sheet_names, rotation=45)
        ax.legend()

        plt.tight_layout()
        graph_path = "graph.png"
        plt.savefig(graph_path, dpi=300)
        if show:
            plt.show()
        else:
            plt.close()
        return graph_path
    return None

# Function to download the summary as a PDF
def download_summary():
    if not file_path:
        print("No file selected. Please choose an Excel file first.")
        return

    save_dir = filedialog.askdirectory(title="Select Directory to Save PDF")
    if not save_dir:
        print("No directory selected. PDF generation cancelled.")
        return

    file_name = f"{os.path.splitext(os.path.basename(file_path))[0]}_Summary.pdf"
    pdf_file_path = os.path.join(save_dir, file_name)

    pdf = canvas.Canvas(pdf_file_path, pagesize=letter)
    pdf.setFont("Helvetica-Bold", 16)
    pdf.drawString(200, 750, "SUMMARY REPORT")
    pdf.setFont("Helvetica", 12)
    
    pdf.drawString(50, 700, "File Name")
    pdf.drawString(250, 700, "PASS Count")
    pdf.drawString(350, 700, "FAIL Count")
    pdf.drawString(450, 700, "Status")
    
    y_position = 680
    for item in tree.get_children():
        values = tree.item(item, 'values')
        pdf.drawString(50, y_position, values[0])
        pdf.drawString(250, y_position, str(values[1]))
        pdf.drawString(350, y_position, str(values[2]))
        if values[3].upper() == "PASS":
            pdf.setFillColor(colors.green)
        elif values[3].upper() == "FAIL":
            pdf.setFillColor(colors.red)
        else:
            pdf.setFillColor(colors.black)
        pdf.drawString(450, y_position, values[3])
        pdf.setFillColor(colors.black)
        y_position -= 20

        if y_position < 50:
            pdf.showPage()
            y_position = 750

    graph_path = display_graph(show=False)
    if graph_path:
        pdf.showPage()
        pdf.drawString(200, 750, "PASS/FAIL Graph")
        pdf.drawImage(graph_path, 50, 400, width=500, height=300)

    pdf.save()
    # print(f"Summary saved as {pdf_file_path}")

# Create the main window for Tkinter
root = tk.Tk()
root.title("KK :- Excel Sheets Status")
root.state('zoomed')

title_font = Font(family="Helvetica", size=24, weight="bold", underline=True)
title_label = tk.Label(root, text="SUMMARY REPORT", font=title_font, fg="red")
title_label.pack(pady=20)

choose_file_button = tk.Button(root, text="Choose Excel File", command=select_file)
choose_file_button.pack(pady=10)

graph_button = tk.Button(root, text="Show PASS/FAIL Graph", command=display_graph)
graph_button.pack(pady=10)

download_button = tk.Button(root, text="Download Summary", command=download_summary)
download_button.pack(pady=10)

tree = ttk.Treeview(root, columns=("File Name", "PASS Count", "FAIL Count", "Status"), show="headings")
tree.heading("File Name", text="File Name", anchor="center")
tree.heading("PASS Count", text="PASS Count", anchor="center")
tree.heading("FAIL Count", text="FAIL Count", anchor="center")
tree.heading("Status", text="Status", anchor="center")
tree.column("File Name", anchor="w", width=200)
tree.column("PASS Count", anchor="center", width=150)
tree.column("FAIL Count", anchor="center", width=150)
tree.column("Status", anchor="center", width=150)

scrollbar = ttk.Scrollbar(root, orient="vertical", command=tree.yview)
tree.configure(yscrollcommand=scrollbar.set)
scrollbar.pack(side="right", fill="y")
tree.pack(expand=True, fill="both")

tree.tag_configure("pass", foreground="green")
tree.tag_configure("fail", foreground="red")

root.mainloop()
