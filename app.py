import streamlit as st
import speech_recognition as sr
from pydub import AudioSegment
from openpyxl import Workbook
from datetime import timedelta

# Title
st.title("Audio Transcriber with Editable Transcript")

# File Upload
uploaded_file = st.file_uploader("Upload an audio file", type=["mp3", "wav"])
if uploaded_file:
    st.success("File uploaded successfully!")
    
    # Convert uploaded file to AudioSegment
    audio_file_path = "uploaded_audio.wav"
    with open(audio_file_path, "wb") as f:
        f.write(uploaded_file.read())

    # Audio playback
    st.audio(audio_file_path)

    # Transcribe Button
    if st.button("Transcribe Audio"):
        recognizer = sr.Recognizer()
        audio = AudioSegment.from_file(audio_file_path)

        chunk_length = 1000  # Process 1-second chunks
        transcription_data = []

        # Split into chunks
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
            transcription_data.append({
                "start": f"{start_time:.3f}",
                "text": text,
                "speaker": "Speaker 1"
            })

        # Display Transcription
        for segment in transcription_data:
            st.write(f"**Start Time:** {str(timedelta(seconds=float(segment['start'])))}")
            speaker = st.text_input(f"Speaker for segment starting at {segment['start']}", value=segment["speaker"])
            transcript = st.text_area(f"Transcript for segment starting at {segment['start']}", value=segment["text"])
            segment["speaker"] = speaker
            segment["text"] = transcript

        # Export Button
        if st.button("Save Transcription to Excel"):
            save_path = "transcription.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Transcription"
            sheet.append(["Start Time", "Speaker", "Text"])
            for segment in transcription_data:
                sheet.append([
                    str(timedelta(seconds=float(segment['start']))),
                    segment["speaker"],
                    segment["text"]
                ])
            workbook.save(save_path)
            st.success("Transcription saved successfully!")
            st.download_button("Download Excel File", data=open(save_path, "rb"), file_name="transcription.xlsx")
