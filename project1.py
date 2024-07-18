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

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

class VideoTextExtractorApp:

    def __init__(self, root):
        self.root = root
        self.root.title("Video Text Extractor and QR Scanner")
        self.root.geometry("800x600")
        self.cap = None
        self.frame_count = 0
        self.extracted_text = ""
        self.extracted_qr = ""
        self.qr_url = ""
        self.product_scan_count = 0  # Counter for QR code product scans
        self.text_scan_count = 0  # Counter for text scans
        self.weight = 0.0  # Variable to store weight
        self.camera_index = 0
        self.weight_port = 'COM3'  # Adjust this to your serial port
        self.serial_connection = None
        self.create_widgets()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_widgets(self):
        self.camera_label = tk.Label(self.root, text="Select Camera:")
        self.camera_label.grid(row=0, column=0, pady=10, padx=10)

        self.camera_selection = ttk.Combobox(self.root, values=["Front Camera", "Back Camera"])
        self.camera_selection.current(0)
        self.camera_selection.grid(row=0, column=1, pady=10, padx=10)

        self.start_button = tk.Button(self.root, text="Start Scanning", command=self.start_scanning, bg='lightblue')
        self.start_button.grid(row=1, column=0, pady=10, padx=10)

        self.stop_button = tk.Button(self.root, text="Stop Scanning", command=self.stop_scanning, bg='lightblue')
        self.stop_button.grid(row=1, column=1, pady=10, padx=10)

        self.text_output = Text(self.root, wrap='word', height=10)
        self.text_output.grid(row=2, column=0, columnspan=2, pady=10, padx=10, sticky="nsew")

        self.video_label = tk.Label(self.root)
        self.video_label.grid(row=3, column=0, pady=5, padx=5, columnspan=2)

        self.digits_label = tk.Label(self.root, text="", wraplength=780, anchor="w", justify="left")
        self.digits_label.grid(row=4, column=0, columnspan=2, pady=10, padx=10, sticky="nsew")

        self.qr_label = tk.Label(self.root, text="", wraplength=780, anchor="w", justify="left")
        self.qr_label.grid(row=5, column=0, columnspan=2, pady=10, padx=10, sticky="nsew")

        self.open_link_button = tk.Button(self.root, text="Open QR Code Link", command=self.open_qr_link, bg='lightgreen', state='disabled')
        self.open_link_button.grid(row=6, column=0, columnspan=2, pady=10, padx=10)

        self.qr_scan_count_label = tk.Label(self.root, text="QR Codes Scanned: 0", wraplength=780, anchor="w", justify="left")
        self.qr_scan_count_label.grid(row=7, column=0, columnspan=2, pady=10, padx=10, sticky="nsew")

        self.text_scan_count_label = tk.Label(self.root, text="Texts Scanned: 0", wraplength=780, anchor="w", justify="left")
        self.text_scan_count_label.grid(row=8, column=0, columnspan=2, pady=10, padx=10, sticky="nsew")

        self.weight_label = tk.Label(self.root, text="Weight: 0.0 kg", wraplength=780, anchor="w", justify="left")
        self.weight_label.grid(row=9, column=0, columnspan=2, pady=10, padx=10, sticky="nsew")

        self.generate_pdf_button = tk.Button(self.root, text="Generate PDF", command=self.generate_pdf, bg='lightyellow')
        self.generate_pdf_button.grid(row=10, column=0, columnspan=2, pady=10, padx=10)

        self.root.grid_rowconfigure(2, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=1)

    def start_scanning(self):
        if self.cap is not None:
            self.text_output.insert(tk.END, "Already scanning...\n")
            return

        self.camera_index = 0 if self.camera_selection.get() == "Front Camera" else 1

        self.text_output.insert(tk.END, "Starting scanning...\n")
        self.cap = cv2.VideoCapture(self.camera_index)
        self.frame_count = 0
        self.extracted_text = ""
        self.extracted_qr = ""
        self.qr_url = ""
        self.product_scan_count = 0  # Reset the QR code scan counter
        self.text_scan_count = 0  # Reset the text scan counter
        self.connect_to_weight_machine()  # Connect to the weight machine
        self.scan_video()

        # Set a timer to stop scanning after 10 minutes (600,000 milliseconds)
        self.root.after(600000, self.stop_scanning)

    def stop_scanning(self):
        if self.cap is not None:
            self.cap.release()
            self.cap = None
            self.text_output.insert(tk.END, "Scanning stopped.\n")
        if self.serial_connection is not None:
            self.serial_connection.close()
            self.serial_connection = None

    def on_closing(self):
        self.stop_scanning()
        self.root.destroy()

    def connect_to_weight_machine(self):
        try:
            self.serial_connection = serial.Serial(self.weight_port, 9600, timeout=1)
            self.text_output.insert(tk.END, "Connected to weight machine.\n")
        except serial.SerialException as e:
            self.text_output.insert(tk.END, f"Failed to connect to weight machine: {e}\n")

    def get_weight(self):
        if self.serial_connection is not None and self.serial_connection.is_open:
            self.serial_connection.write(b'R')  # Assuming 'R' requests weight
            weight_data = self.serial_connection.readline().decode('utf-8').strip()
            try:
                self.weight = float(weight_data)
            except ValueError:
                self.weight = 0.0
            self.weight_label.config(text=f"Weight: {self.weight:.2f} kg")
        else:
            self.weight = 0.0

    def resize_frame(self, frame, width, height):
        resized_frame = cv2.resize(frame, (width, height), interpolation=cv2.INTER_AREA)
        return resized_frame

    def show_frame(self, frame):
        frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        frame = Image.fromarray(frame)
        frame = ImageTk.PhotoImage(frame)
        self.video_label.configure(image=frame)
        self.video_label.image = frame

    def extract_text_from_image(self, image):
        gray_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        text = pytesseract.image_to_string(gray_image)
        return text

    def extract_qr_from_image(self, image):
        decoded_objects = pyzbar.decode(image)
        qr_data = ""
        for obj in decoded_objects:
            qr_data += obj.data.decode('utf-8') + " "
        return qr_data

    def scan_video(self):
        if self.cap is None:
            return 

        ret, frame = self.cap.read()
        if not ret:
            self.stop_scanning()
            return

        self.frame_count += 1
        if self.frame_count % 50 == 0:
            start_time = time.time()
            frame = self.resize_frame(frame, 500, 200)

            # Extract text from the frame
            text = self.extract_text_from_image(frame)
            if text.strip():
                self.extracted_text += text.replace('\n', ' ') + ' '
                self.text_scan_count += 1  # Increment the text scan counter
                self.text_scan_count_label.config(text=f"Texts Scanned: {self.text_scan_count}")

            # Extract QR code from the frame
            qr_data = self.extract_qr_from_image(frame)
            if qr_data:
                self.extracted_qr += qr_data.replace('\n', ' ') + ' '
                self.product_scan_count += 1  # Increment the QR code scan counter
                self.qr_scan_count_label.config(text=f"QR Codes Scanned: {self.product_scan_count}")
                if "http://" in qr_data or "https://" in qr_data:
                    self.qr_url = qr_data.strip()
                    self.open_link_button.config(state='normal')

            end_time = time.time()
            elapsed_time = end_time - start_time
            self.text_output.insert(tk.END, f"Text: {text} QR: {qr_data} (Time: {elapsed_time:.2f} sec)\n")
            self.digits_label.config(text=self.extracted_text)
            self.qr_label.config(text=self.extracted_qr)
            self.show_frame(frame)
            self.root.update_idletasks()

        self.root.after(10, self.scan_video)

    def open_qr_link(self):
        if self.qr_url:
            webbrowser.open(self.qr_url)
        else:
            messagebox.showinfo("No QR Code", "No QR Code link to open.")

    def generate_pdf(self):
        pdf = FPDF()
        pdf.add_page()

        pdf.set_font("Arial", size=12)
        pdf.cell(200, 10, txt="Scanned Data Report", ln=True, align='C')

        pdf.ln(10)
        pdf.cell(200, 10, txt="Extracted Text:", ln=True)
        pdf.multi_cell(0, 10, txt=self.extracted_text)

        pdf.ln(10)
        pdf.cell(200, 10, txt="Extracted QR Codes:", ln=True)
        pdf.multi_cell(0, 10, txt=self.extracted_qr)

        pdf.ln(10)
        pdf.cell(200, 10, txt=f"Number of Texts Scanned: {self.text_scan_count}", ln=True)

        pdf.ln(10)
        pdf.cell(200, 10, txt=f"Number of QR Codes Scanned: {self.product_scan_count}", ln=True)

        pdf.ln(10)
        pdf.cell(200, 10, txt=f"Current Weight: {self.weight:.2f} kg", ln=True)

        pdf_output_path = "scanned_data_report.pdf"
        pdf.output(pdf_output_path)
        messagebox.showinfo("PDF Generated", f"PDF report has been generated and saved as {pdf_output_path}")

if __name__ == "__main__":
    root = tk.Tk()
    app = VideoTextExtractorApp(root)
    # root.resizable(0, 0)
    root.mainloop()