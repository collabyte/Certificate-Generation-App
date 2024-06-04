import os
import cv2
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def delete_old_data(folder):
    """Delete previously generated certificate images."""
    # Check if the output folder exists; if not, create it
    if not os.path.exists(folder):
        os.makedirs(folder)
    else:
        # If the folder exists, remove all files in it
        for file in os.listdir(folder):
            os.remove(os.path.join(folder, file))

def load_names_from_excel(file_path):
    """Load names from an Excel file, skipping the first row."""
    # Read the Excel file and skip the first row
    df = pd.read_excel(file_path, header=None, skiprows=1)
    # Extract names from the first column and convert them to a list
    return df.iloc[:, 0].tolist()

def generate_certificate(name, template_path, output_folder):
    """Generate a certificate image with the given name."""
    # Load the certificate template image
    certificate_template_image = cv2.imread(template_path)

    # Load the font
    font = cv2.FONT_HERSHEY_SIMPLEX

    # Set the font scale and thickness
    font_scale = 4.0
    font_thickness = 9  # Boldness

    # Get the text size
    (text_width, text_height), _ = cv2.getTextSize(name, font, font_scale, font_thickness)

    # Calculate the position to center align the text
    text_x = 1000 - text_width // 2
    text_y = 515 + text_height // 2

    # Set the font color (teal: #006C89)
    font_color = (137, 108, 0)

    # Put the text on the image
    cv2.putText(certificate_template_image, name, (text_x, text_y), font, font_scale, font_color, font_thickness, cv2.LINE_AA)

    # Save the generated certificate image
    output_path = os.path.join(output_folder, f"{name}.png")
    cv2.imwrite(output_path, certificate_template_image)
    print(f"Certificate generated for {name}")

def generate_certificates(names, template_path, output_folder, progress_label):
    """Generate certificates for a list of names."""
    # Iterate over each name and generate a certificate for it
    total_names = len(names)
    for index, name in enumerate(names):
        generate_certificate(name, template_path, output_folder)
        progress_label.config(text=f"Processing {index + 1} / {total_names}")
        progress_label.update()

def main():
    root = tk.Tk()
    root.title("Certificate Generator")

    def select_excel_file():
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        excel_entry.delete(0, tk.END)
        excel_entry.insert(0, file_path)

    def select_template_file():
        file_path = filedialog.askopenfilename(filetypes=[("Image files", "*.png;*.jpg;*.jpeg")])
        template_entry.delete(0, tk.END)
        template_entry.insert(0, file_path)

    def select_output_folder():
        folder_path = filedialog.askdirectory()
        output_entry.delete(0, tk.END)
        output_entry.insert(0, folder_path)

    def generate_certificates_gui():
        excel_file = excel_entry.get()
        template_file = template_entry.get()
        output_folder = output_entry.get()

        if not excel_file or not template_file or not output_folder:
            messagebox.showerror("Error", "Please fill in all fields.")
            return

        if not os.path.exists(excel_file):
            messagebox.showerror("Error", "Excel file not found.")
            return

        if not os.path.exists(template_file):
            messagebox.showerror("Error", "Template file not found.")
            return

        delete_old_data(output_folder)
        names = load_names_from_excel(excel_file)
        generate_certificates(names, template_file, output_folder, progress_label)
        messagebox.showinfo("Success", "Certificates generated successfully.")

    # GUI elements lai create gareko
    excel_label = tk.Label(root, text="Excel File:")
    excel_label.grid(row=0, column=0, sticky="w")
    excel_entry = tk.Entry(root, width=50)
    excel_entry.grid(row=0, column=1)
    excel_button = tk.Button(root, text="Browse", command=select_excel_file)
    excel_button.grid(row=0, column=2)

    template_label = tk.Label(root, text="Template File:")
    template_label.grid(row=1, column=0, sticky="w")
    template_entry = tk.Entry(root, width=50)
    template_entry.grid(row=1, column=1)
    template_button = tk.Button(root, text="Browse", command=select_template_file)
    template_button.grid(row=1, column=2)

    output_label = tk.Label(root, text="Output Folder:")
    output_label.grid(row=2, column=0, sticky="w")
    output_entry = tk.Entry(root, width=50)
    output_entry.grid(row=2, column=1)
    output_button = tk.Button(root, text="Browse", command=select_output_folder)
    output_button.grid(row=2, column=2)

    generate_button = tk.Button(root, text="Generate Certificates", command=generate_certificates_gui)
    generate_button.grid(row=3, columnspan=3)

    progress_label = tk.Label(root, text="")
    progress_label.grid(row=4, columnspan=3)

    root.mainloop()

if __name__ == '__main__':
    main()
