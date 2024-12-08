import tkinter as tk
from tkinter import filedialog, messagebox
import speech_recognition as sr
from pydub import AudioSegment
from openpyxl import Workbook
from datetime import timedelta
import re


class AudioTranscriberApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Audio Transcriber with Editable Transcript")
        
        # Widgets
        self.upload_button = tk.Button(root, text="Upload Audio", command=self.upload_audio)
        self.upload_button.pack(pady=10)

        self.transcribe_button = tk.Button(root, text="Transcribe Audio", command=self.transcribe_audio, state=tk.DISABLED)
        self.transcribe_button.pack(pady=10)

        self.transcription_frame = tk.Frame(root)
        self.transcription_frame.pack(pady=10, fill=tk.BOTH, expand=True)
        
        self.transcription_canvas = tk.Canvas(self.transcription_frame)
        self.scrollbar = tk.Scrollbar(self.transcription_frame, orient="vertical", command=self.transcription_canvas.yview)
        self.transcription_canvas.configure(yscrollcommand=self.scrollbar.set)

        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.transcription_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.transcription_area = tk.Frame(self.transcription_canvas)
        self.transcription_canvas.create_window((0, 0), window=self.transcription_area, anchor="nw")

        self.transcription_area.bind("<Configure>", lambda e: self.transcription_canvas.configure(scrollregion=self.transcription_canvas.bbox("all")))

        self.save_button = tk.Button(root, text="Save to Excel", command=self.save_to_excel, state=tk.DISABLED)
        self.save_button.pack(pady=10)

        self.find_replace_button = tk.Button(root, text="Find and Replace", command=self.open_find_replace)
        self.find_replace_button.pack(pady=10)

        self.audio_file = None
        self.transcription_data = []

    def upload_audio(self):
        file_path = filedialog.askopenfilename(filetypes=[("Audio Files", "*.wav *.mp3")])
        if file_path:
            self.audio_file = file_path
            self.transcribe_button.config(state=tk.NORMAL)
            messagebox.showinfo("Audio Uploaded", "Audio file uploaded successfully!")

    def transcribe_audio(self):
        if not self.audio_file:
            messagebox.showerror("No Audio File", "Please upload an audio file first.")
            return

        recognizer = sr.Recognizer()
        audio = AudioSegment.from_file(self.audio_file)
        chunk_length = 1000  # Process 1-second chunks

        self.transcription_data = []
        for widget in self.transcription_area.winfo_children():
            widget.destroy()

        for i in range(0, len(audio), chunk_length):
            chunk = audio[i:i + chunk_length]
            chunk.export("temp_chunk.wav", format="wav")
            with sr.AudioFile("temp_chunk.wav") as source:
                try:
                    audio_data = recognizer.record(source)
                    text = recognizer.recognize_google(audio_data)
                except sr.UnknownValueError:
                    text = "[Unintelligible]"

            start_time = i / 1000  # Start time in seconds
            end_time = (i + chunk_length) / 1000  # End time in seconds
            self.transcription_data.append({
                "start": f"{start_time:.3f}",
                "end": f"{end_time:.3f}",
                "text": text,
                "speaker": "Speaker 1"
            })

        self.display_transcription()
        self.save_button.config(state=tk.NORMAL)

    def display_transcription(self):
        for widget in self.transcription_area.winfo_children():
            widget.destroy()

        for segment in self.transcription_data:
            frame = tk.Frame(self.transcription_area, borderwidth=1, relief="solid", pady=5)
            frame.pack(fill=tk.X, padx=5, pady=2)

            tk.Label(frame, text=f"Start: {self.format_time(segment['start'])}", font=("Arial", 10)).pack(anchor="w")

            speaker_label = tk.Entry(frame, width=20)
            speaker_label.insert(0, segment['speaker'])
            speaker_label.pack(side=tk.LEFT, padx=5)

            text_area = tk.Text(frame, height=2, width=60)
            text_area.insert(1.0, segment['text'])
            text_area.pack(side=tk.LEFT, padx=5)

            segment["speaker_widget"] = speaker_label
            segment["text_widget"] = text_area

    def format_time(self, seconds):
        """Convert seconds to HH:MM:SS.s format."""
        td = timedelta(seconds=float(seconds))
        return str(td)

    def save_to_excel(self):
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if save_path:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Transcription"

            # Header row
            sheet.append(["Start Time", "Speaker", "Text"])

            # Data rows
            for segment in self.transcription_data:
                speaker = segment["speaker_widget"].get()
                text = segment["text_widget"].get("1.0", tk.END).strip()
                start_time = self.format_time(segment["start"])
                sheet.append([start_time, speaker, text])

            # Save the workbook
            workbook.save(save_path)
            messagebox.showinfo("Save Successful", "Transcription saved successfully!")

    def open_find_replace(self):
        """Open the Find and Replace dialog."""
        find_replace_window = tk.Toplevel(self.root)
        find_replace_window.title("Find and Replace")

        tk.Label(find_replace_window, text="Find (Regex):").grid(row=0, column=0, padx=10, pady=5)
        find_entry = tk.Entry(find_replace_window, width=30)
        find_entry.grid(row=0, column=1, padx=10, pady=5)

        tk.Label(find_replace_window, text="Replace:").grid(row=1, column=0, padx=10, pady=5)
        replace_entry = tk.Entry(find_replace_window, width=30)
        replace_entry.grid(row=1, column=1, padx=10, pady=5)

        def perform_find_replace():
            pattern = find_entry.get()
            replacement = replace_entry.get()
            if not pattern:
                messagebox.showerror("Error", "Please enter a pattern to find.")
                return

            for segment in self.transcription_data:
                text_widget = segment["text_widget"]
                current_text = text_widget.get("1.0", tk.END)
                new_text = re.sub(pattern, replacement, current_text)
                text_widget.delete("1.0", tk.END)
                text_widget.insert("1.0", new_text)
            messagebox.showinfo("Find and Replace", "Replacement complete.")

        replace_button = tk.Button(find_replace_window, text="Replace", command=perform_find_replace)
        replace_button.grid(row=2, column=0, columnspan=2, pady=10)

# Create the GUI
root = tk.Tk()
app = AudioTranscriberApp(root)
root.mainloop()
