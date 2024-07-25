import tkinter as tk
from tkinter import Label, Text, ttk, messagebox
import cv2
import pytesseract
from PIL import Image, ImageTk
import time
from pyzbar import pyzbar
import webbrowser
import serial
from fpdf import FPDF
import numpy as np
import os

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

class ColorDetector:
    def __init__(self, text_output):
        self.text_output = text_output
        self.detected_colors = {}

    def detect_color(self, frame):
        hsv_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2HSV)

        color_ranges = {
            'Red': [((0, 50, 50), (10, 255, 255)), ((160, 50, 50), (180, 255, 255))],
            'Green': ((35, 50, 50), (85, 255, 255)),
            'Blue': ((100, 50, 50), (140, 255, 255)),
            'Yellow': ((20, 50, 50), (30, 255, 255)),
            'Orange': ((10, 100, 20), (25, 255, 255)),
            'Purple': ((140, 50, 50), (160, 255, 255)),
            'Brown': ((10, 100, 20), (20, 255, 200)),
            'Black': ((0, 0, 0), (180, 255, 30)),
            'White': ((0, 0, 200), (180, 30, 255))
        }

        detected_colors = {}

        for color_name, color_range in color_ranges.items():
            if color_name == 'Red':
                lower1, upper1 = color_range[0]
                lower2, upper2 = color_range[1]

                mask1 = cv2.inRange(hsv_frame, np.array(lower1), np.array(upper1))
                mask2 = cv2.inRange(hsv_frame, np.array(lower2), np.array(upper2))

                mask = cv2.bitwise_or(mask1, mask2)
            else:
                lower, upper = color_range
                mask = cv2.inRange(hsv_frame, np.array(lower), np.array(upper))

            color_detection = cv2.countNonZero(mask)
            if color_detection > 0:
                detected_colors[color_name] = color_detection

        if detected_colors:
            detected_color = max(detected_colors, key=detected_colors.get)
            self.text_output.insert(tk.END, f"Detected color: {detected_color}\n")
            self.detected_colors[detected_color] = self.detected_colors.get(detected_color, 0) + 1


class VideoTextExtractorApp:
    alreadyText = ""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Video Text Extractor and QR Scanner")
        self.root.geometry("900x600")
        self.cap = None
        self.frame_count = 0
        self.extracted_texts = []  # List to store all extracted texts
        self.extracted_qrs = []    # List to store all extracted QR codes
        self.qr_url = ""
        self.product_scan_count = 0
        self.text_scan_count = 0
        self.qr_scan_count = 0  # Initialize qr_scan_count here
        self.weight = 0.0
        self.camera_index = 0
        self.weight_port = 'COM3'
        self.serial_connection = None
        self.detected_colors = {}  # Dictionary to store detected colors and their counts
        self.create_widgets()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Initialize ColorDetector
        self.color_detector = ColorDetector(self.text_output)

    def create_widgets(self):
        # Create a canvas
        self.canvas = tk.Canvas(self.root)
        self.canvas.pack(side="left", fill="both", expand=True)

        # Create vertical scrollbar
        self.v_scrollbar = tk.Scrollbar(self.root, orient="vertical", command=self.canvas.yview)
        self.v_scrollbar.pack(side="right", fill="y")

        # Create horizontal scrollbar
        self.h_scrollbar = tk.Scrollbar(self.root, orient="horizontal", command=self.canvas.xview)
        self.h_scrollbar.pack(side="bottom", fill="x")

        # Configure the canvas
        self.canvas.configure(yscrollcommand=self.v_scrollbar.set, xscrollcommand=self.h_scrollbar.set)

        # Create a frame on the canvas to hold the widgets
        self.frame = tk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.frame, anchor="nw")

        # Bind the configure event to update the scroll region
        self.frame.bind("<Configure>", self.update_scroll_region)

        self.camera_label = tk.Label(self.frame, text="Select Camera:")
        self.camera_label.grid(row=0, column=1, pady=10, padx=10)

        self.camera_selection = ttk.Combobox(self.frame, values=["Front Camera", "Back Camera"])
        self.camera_selection.current(0)

        self.camera_selection.grid(row=0, column=2, pady=10, padx=10)

        self.start_button = tk.Button(self.frame, text="Start Scanning", command=self.start_scanning, bg='lightblue')
        self.start_button.grid(row=1, column=1, pady=10, padx=10)

        self.stop_button = tk.Button(self.frame, text="Stop Scanning", command=self.stop_scanning, bg='lightblue')
        self.stop_button.grid(row=1, column=2, pady=10, padx=10)

        self.text_output = Text(self.frame, wrap='word',width=80, height=10)
        self.text_output.grid(row=2, column=1, columnspan=2, pady=10, padx=10, sticky="nsew")

        self.video_label = tk.Label(self.frame)
        self.video_label.grid(row=3, column=0, pady=5, padx=5, columnspan=2)

        self.text_list_label = tk.Label(self.frame, text="", wraplength=780, anchor="w", justify="left")
        self.text_list_label.grid(row=4, column=0, columnspan=2, pady=10, padx=10, sticky="nsew")

        self.qr_list_label = tk.Label(self.frame, text="", wraplength=780, anchor="w", justify="left")
        self.qr_list_label.grid(row=5, column=0, columnspan=2, pady=10, padx=10, sticky="nsew")

        self.open_link_button = tk.Button(self.frame, text="Open QR Link", command=self.open_qr_link, state='disabled', bg='lightblue')
        self.open_link_button.grid(row=6, column=1, pady=10, padx=10)

        self.generate_pdf_button = tk.Button(self.frame, text="Generate PDF", command=self.generate_pdf, bg='lightblue')
        self.generate_pdf_button.grid(row=6, column=2, pady=10, padx=10)

        self.weight_label = tk.Label(self.frame, text="Weight: 0.0 kg")
        self.weight_label.grid(row=9, column=0, pady=10, padx=10)

        self.text_scan_count_label = tk.Label(self.frame, text="Texts Scanned: 0")
        self.text_scan_count_label.grid(row=7, column=0, pady=10, padx=10)

        self.qr_scan_count_label = tk.Label(self.frame, text="QR Codes Scanned: 0")
        self.qr_scan_count_label.grid(row=8, column=0, pady=10, padx=10)

    def update_scroll_region(self, event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def start_scanning(self):
        self.camera_index = 0 if self.camera_selection.get() == "Front Camera" else 1
        self.cap = cv2.VideoCapture(self.camera_index)
        if not self.cap.isOpened():
            messagebox.showerror("Error", "Failed to open the camera.")
            return
        
        self.open_serial_connection()
        self.scan_frame()

    def stop_scanning(self):
        if self.cap is not None and self.cap.isOpened():
            self.cap.release()
            self.video_label.config(image='')
            self.cap = None
            self.close_serial_connection()
            self.text_output.insert(tk.END, "Stopped scanning.\n")

    def scan_frame(self):
        if self.cap is None or not self.cap.isOpened():
            return

        ret, frame = self.cap.read()
        if not ret:
            self.stop_scanning()
            messagebox.showerror("Error", "Failed to capture video frame.")
            return

        self.detect_qr_codes(frame)
        self.extract_text(frame)
        self.color_detector.detect_color(frame)  # Call the detect_color method from ColorDetector

        # Display video frame
        self.display_frame(frame)

        self.frame_count += 1

        if self.frame_count % 30 == 0:
            self.update_weight()

        self.root.after(10, self.scan_frame)

    def extract_text(self, frame):
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        text = pytesseract.image_to_string(gray)
        if text.strip():
            if text != self.alreadyText:
                self.text_output.insert(tk.END, f"Extracted text: {text}\n")
                self.extracted_texts.append(text)
                self.text_list_label.config(text="\n".join(self.extracted_texts))
                self.text_scan_count += 1
                self.text_scan_count_label.config(text=f"Texts Scanned: {self.text_scan_count}")
                self.alreadyText = text

    def detect_qr_codes(self, frame):
        decoded_objects = pyzbar.decode(frame)
        for obj in decoded_objects:
            qr_text = obj.data.decode('utf-8')
            if qr_text not in self.extracted_qrs:
                self.text_output.insert(tk.END, f"QR Code detected: {qr_text}\n")
                self.extracted_qrs.append(qr_text)
                self.qr_list_label.config(text="\n".join(self.extracted_qrs))
                self.qr_scan_count += 1
                self.qr_scan_count_label.config(text=f"QR Codes Scanned: {self.qr_scan_count}")
                if qr_text.startswith("http://") or qr_text.startswith("https://"):
                    self.qr_url = qr_text
                    self.open_link_button.config(state='normal')

    def open_qr_link(self):
        if self.qr_url:
            webbrowser.open(self.qr_url)

    def display_frame(self, frame):
        frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        frame = cv2.resize(frame, (400, 150))
        img = Image.fromarray(frame)
        imgtk = ImageTk.PhotoImage(image=img)
        self.video_label.imgtk = imgtk
        self.video_label.config(image=imgtk)

    def open_serial_connection(self):
        try:
            self.serial_connection = serial.Serial(self.weight_port, 9600, timeout=1)
        except serial.SerialException as e:
            self.text_output.insert(tk.END, f"Failed to connect to weight sensor: {e}\n")
            self.serial_connection = None

    def close_serial_connection(self):
        if self.serial_connection is not None:
            self.serial_connection.close()
            self.serial_connection = None

    def update_weight(self):
        if self.serial_connection is not None:
            self.serial_connection.write(b'W')
            weight_data = self.serial_connection.readline().decode('utf-8').strip()
            try:
                self.weight = float(weight_data)
                self.weight_label.config(text=f"Weight: {self.weight:.2f} kg")
            except ValueError:
                self.text_output.insert(tk.END, "Failed to read weight data.\n")

    def generate_pdf(self):
        try:
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=12)
            
            pdf.cell(200, 10, txt="Scanned Texts", ln=True, align='C')
            for text in self.extracted_texts:
                pdf.multi_cell(0, 10, txt=text)
            
            pdf.add_page()
            pdf.cell(200, 10, txt="Scanned QR Codes", ln=True, align='C')
            for qr in self.extracted_qrs:
                pdf.multi_cell(0, 10, txt=qr)
            
            pdf.add_page()
            pdf.cell(200, 10, txt="Detected Colors", ln=True, align='C')
            for color, count in self.color_detector.detected_colors.items():
                pdf.multi_cell(0, 10, txt=f"{color}: {count}")
            
            pdf.add_page()
            pdf.cell(200, 10, txt="Weight Measurements", ln=True, align='C')
            pdf.cell(0, 10, txt=f"Weight: {self.weight:.2f} kg")
 
            pdf_output = "scanned_data.pdf"
            pdf.output(pdf_output)
            self.text_output.insert(tk.END, f"PDF generated: {pdf_output}\n")
            print("PDF generated")  # Debug statement

            # Add a short delay before checking if the file exists
            time.sleep(0.5)

            # Verify the file was created
            if os.path.exists(pdf_output):
                self.text_output.insert(tk.END, f"PDF successfully saved as {pdf_output}\n")
                print("PDF successfully saved")  # Debug statement
            else:
                self.text_output.insert(tk.END, "Error: PDF file not found after saving.\n")
                print("PDF file not found")  # Debug statement
                
        except Exception as e:
            self.text_output.insert(tk.END, f"Error generating PDF: {str(e)}\n")
            print(f"Error generating PDF: {str(e)}")  # Debug statement

    def on_closing(self):
        self.stop_scanning()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = VideoTextExtractorApp(root)
    root.resizable(0,0)
    root.mainloop()
