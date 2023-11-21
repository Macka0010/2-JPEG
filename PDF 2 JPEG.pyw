import os
import tkinter as tk
from tkinter import filedialog, ttk
import fitz
from PIL import Image, ImageTk, ImageOps
import mimetypes
from docx import Document
import io

# Global variables
file_paths = []  # Store selected PDF file paths
sliced_pdf_var = None  # Variable to store the state of the checkbox
canvas = None  # Canvas widget to display images
canvas_image = None  # Currently displayed image on the canvas
progress = None  # Progress bar widget
image_counter = None  # Label to display the number of extracted images
pdf_counter = None  # Label to display the number of PDFs left

# Function to open the folder containing extracted images
def open_images_folder():
    global file_paths
    if file_paths:
        for file_path in file_paths:
            folder_path = os.path.join(os.path.dirname(file_path), "PDF Images")
            if os.path.exists(folder_path):
                os.startfile(folder_path)
    else:
        status_label.configure(text="No files selected.")

def create_slice_input(frame):
    # Label for number of slices
    slice_label = ttk.Label(frame, text="Number of slices:")
    slice_label.grid(row=0, column=0, sticky=tk.W, padx=(10,0))

    # Entry widget for number of slices
    slice_entry = ttk.Entry(frame)
    slice_entry.grid(row=1, column=0, sticky=tk.W)

    return slice_entry

# Function to extract images from DOCX files
def extract_images_from_docx(docx_path, output_dir):
    doc = Document(docx_path)
    images = []
    for i, rel in enumerate(doc.part.rels.values()):
        if "image" in rel.reltype:
            img = rel.target_part.blob
            img_format = rel.target_part.content_type.split("/")[-1]
            image_path = os.path.join(output_dir, f"docx_image_{i + 1}.{img_format}")
            with open(image_path, "wb") as f:
                f.write(img)
            images.append(image_path)
    return images

# Function to extract and stitch images from PDFs
def extract_and_stitch_images():
    global file_paths, sliced_pdf_var, canvas, canvas_image, progress, image_counter, pdf_counter, stitched_counter
    file_paths = filedialog.askopenfilenames(filetypes=[("PDF and DOCX files", "*.pdf;*.docx")])
    if file_paths:
        open_button.configure(state="disabled")
        broken_pdf_button.configure(state="disabled")
        status_label.configure(text="Extracting and stitching images...")

        total_images_extracted = 0
        total_pdfs = len(file_paths)

    for pdf_index, file_path in enumerate(file_paths):
        file_extension = os.path.splitext(file_path)[1].lower()
        pdf_dir = os.path.dirname(file_path)
        output_dir = os.path.join(pdf_dir, "PDF Images" if file_extension == '.pdf' else "DOCX Images")
        os.makedirs(output_dir, exist_ok=True)

        image_files = []
        if file_extension == '.pdf':    

            # Iterate through selected PDF files
            for pdf_index, file_path in enumerate(file_paths):
                pdf_dir = os.path.dirname(file_path)
                output_dir = os.path.join(pdf_dir, "PDF Images")
                os.makedirs(output_dir, exist_ok=True)

                pdf_file = fitz.open(file_path)
                image_files = []

                total_pages = len(pdf_file)
                # Iterate through pages in the current PDF
                for page_num in range(total_pages):
                    page = pdf_file[page_num]
                    image_list = page.get_images()
                    image_list.sort(key=lambda img: img[0])

                    # Extract and save images from the page
                    for img in image_list:
                        xref = img[0]
                        base_image = pdf_file.extract_image(xref)
                        image_data = base_image["image"]
                        image_format = base_image["ext"]
                        image_path = os.path.join(output_dir, f"image_{pdf_index + 1}_{page_num + 1}_{xref}.{image_format}")

                        with open(image_path, 'wb') as f:
                            f.write(image_data)
                        image_files.append(image_path)

                        total_images_extracted += 1
                        image_counter.configure(text=f"Images Extracted: {total_images_extracted}")
                        pdf_counter.configure(text=f"PDFs Left: {total_pdfs - pdf_index - 1}")

                        img = Image.open(image_path)
                        canvas_image = img
                        display_canvas_image()
                        root.update()

                    progress["value"] = (pdf_index * total_pages + page_num + 1) * 100 / (total_pdfs * total_pages)
                    root.update_idletasks()

                if sliced_pdf_var.get():
                    try:
                        # Retrieve the number of slices from the text box
                        num_slices = int(slice_entry.get())
                        assert num_slices > 0
                    except ValueError:
                        print("Please enter a valid integer for slices.")
                        return
                    except AssertionError:
                        print("Number of slices must be greater than 0.")
                        return

                    if len(image_files) >= num_slices:
                        # Stitch images vertically in groups of 'num_slices'
                        stitched_images = []
                        for i in range(0, len(image_files), num_slices):
                            group_files = image_files[i:i + num_slices]
                            group_files.sort(key=lambda file: int(os.path.splitext(os.path.basename(file))[0].split("_")[-1]))
                            group_images = [Image.open(file) for file in group_files]
                            stitched_image = stitch_images_vertically(group_images)

                            # Save the stitched image
                            save_path = os.path.join(output_dir, f"stitched_{pdf_index + 1}_{len(stitched_images) + 1}.jpeg")
                            stitched_image.save(save_path)

                            # Remove individual image files
                            for file in group_files:
                                os.remove(file)

                            # Update canvas with current slice
                            canvas_image = stitched_image
                            display_canvas_image()

                            stitched_images.append(stitched_image)

                        # Update stitched images tally
                        stitched_counter.configure(text=f"Images stitched: {len(stitched_images)}")

                    pdf_file.close()
        elif file_extension == '.docx':
            image_files = extract_images_from_docx(file_path, output_dir)
            total_images_extracted += len(image_files)

        
        open_button.configure(state="normal")
        broken_pdf_button.configure(state="normal")
        status_label.configure(text="Image extraction and stitching complete!")
        progress["value"] = 0


# Function to stitch images vertically
def stitch_images_vertically(images):
    total_height = sum(image.size[1] for image in images)
    max_width = max(image.size[0] for image in images)
    stitched_image = Image.new("RGB", (max_width, total_height))
    y_offset = 0
    for image in images:
        if image.mode != "RGB":
            image = image.convert("RGB")
        stitched_image.paste(image, (0, y_offset))
        y_offset += image.size[1]

    stitched_image = ImageOps.exif_transpose(stitched_image)
    stitched_image = stitched_image.convert("RGB")

    return stitched_image

# Function to display the current canvas image
def display_canvas_image():
    global canvas, canvas_image
    if canvas_image:
        canvas.delete("all")
        scaled_image = canvas_image.copy()
        scaled_image.thumbnail((300, 300))
        photo_image = ImageTk.PhotoImage(scaled_image)
        canvas.create_image(0, 0, anchor="nw", image=photo_image)
        canvas.image_reference = photo_image
        root.update()
        canvas.update()

def show_quick_guide():
    quick_guide_text = """
    PDF Image Extractor: Quick Guide

Extracting Images from PDFs:

-   Launch the PDF Image Extractor application.
-   Click the "Extract Images" button.
-   Select one or more PDF files to extract images from.
-   Check the "Sliced PDF" option to stitch images together if desired.
-   Click "Extract Images" to start the process.
-   Monitor progress through the progress bar.
-   Extracted images will be saved in a folder named "PDF Images" located in the same folder as your PDF files.
-   View extracted images in the "PDF Images" folder.


Stitching Extracted Images (Optional):
-   If "Sliced PDF" is checked, images will be stitched.
-   The number of stitched images will be displayed.
-   Stitched images will also be saved in the "PDF Images" folder.
-   Exiting the Application:

Once extraction is complete, the status will show "Image extraction and stitching complete!"
Close the application by clicking the close button (X).

    """
    popup = tk.Toplevel(root)
    popup.title("Quick Guide")
    guide_label = ttk.Label(popup, text=quick_guide_text, wraplength=900, justify="left")
    guide_label.pack(padx=20, pady=20)

# Create the main Tkinter window
root = tk.Tk()
root.title("PDF Image Extractor")
frame = ttk.Frame(root, padding="20")
frame.grid()

# Create and place GUI widgets
status_label = ttk.Label(frame, text="")
status_label.grid(row=0, column=0)

slice_entry = create_slice_input(frame)

quick_guide_button = ttk.Button(frame, text="Quick Guide", command=show_quick_guide)
quick_guide_button.grid(row=10, column=0, pady=10)

stitched_counter = ttk.Label(frame, text="Images stitched: 0")
stitched_counter.grid(row=9, column=0, pady=5)

sliced_pdf_var = tk.BooleanVar()
sliced_pdf_checkbox = ttk.Checkbutton(frame, text="Sliced PDF", variable=sliced_pdf_var)
sliced_pdf_checkbox.grid(row=2, column=0)

progress = ttk.Progressbar(frame, orient="horizontal", length=200, mode="determinate")
progress.grid(row=5, column=0, pady=10)

open_button = ttk.Button(frame, text="Open Images Folder", command=open_images_folder, state="disabled")
open_button.grid(row=3, column=0, pady=10)

broken_pdf_button = ttk.Button(frame, text="Extract Images", command=extract_and_stitch_images)
broken_pdf_button.grid(row=4, column=0, pady=10)

canvas = tk.Canvas(frame, width=300, height=300)
canvas.grid(row=6, column=0, pady=10)
canvas_image = None
canvas.image_reference = None

image_counter = ttk.Label(frame, text="Images Extracted: 0")
image_counter.grid(row=7, column=0, pady=5)

pdf_counter = ttk.Label(frame, text="PDFs Left: 0")
pdf_counter.grid(row=8, column=0, pady=5)

root.mainloop()  # Start the Tkinter event loop
