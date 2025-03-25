# Smart Attendance System with Intruder Detection

## Overview
The **Smart Attendance System with Intruder Detection** is a computer vision-based project that automates student attendance marking and detects intruders. Using facial recognition, the system identifies registered students and marks their attendance in an Excel sheet. If an unregistered person (intruder) is detected, the system captures a snapshot and generates an alert.

## Features
- **Face Recognition for Attendance:** Recognizes registered users and marks their attendance.
- **Intruder Detection:** Detects unregistered faces and captures snapshots.
- **Excel Data Storage:** Stores attendance records in an Excel file.
- **GUI for User Interaction:** Allows users to register faces, view attendance, and check intruder logs.
- **Alert System:** Generates an alert when an intruder is detected.
- **Text-to-Speech Alerts:** Provides audio feedback using `pyttsx3`.

## Technologies Used
The project utilizes various Python libraries for face recognition, data storage, and GUI implementation:

### **Python Libraries:**
- `opencv-python (cv2)` – For image and video processing.
- `numpy` – For array operations.
- `face_recognition` – For detecting and encoding faces.
- `os` – For handling file directories.
- `pyttsx3` – For text-to-speech alerts.
- `datetime` – For logging attendance timestamps.
- `openpyxl` – For handling Excel operations.
- `Tkinter` – For GUI implementation.
- 
![Screenshot 2025-03-26 020401](https://github.com/user-attachments/assets/9ad06fe9-5d80-45cc-83d4-88d554fbb80c)
