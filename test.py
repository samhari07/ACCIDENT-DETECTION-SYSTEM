from sklearn.neighbors import KNeighborsClassifier
import cv2
import pickle
import numpy as np
import os
import csv
import time
from datetime import datetime
from win32com.client import Dispatch

def speak(str1):
    speak = Dispatch(("SAPI.SpVoice"))
    speak.Speak(str1)

video = cv2.VideoCapture(0)
facedetect = cv2.CascadeClassifier('data/haarcascade_frontalface_default.xml')

with open('data/names.pkl', 'rb') as w:
    LABELS = pickle.load(w)

with open('data/faces.pkl', 'rb') as f:
    FACES = pickle.load(f)

# Ensure FACES and LABELS have the same number of samples by creating a mapping
face_label_mapping = {}  # Dictionary to store the mapping between face images and labels

for label, face in zip(LABELS, FACES):
    if label not in face_label_mapping:
        face_label_mapping[label] = []
    face_label_mapping[label].append(face)

# Extract faces and labels in a consistent order
FACES = []
LABELS = []
for label, faces in face_label_mapping.items():
    FACES.extend(faces)
    LABELS.extend([label] * len(faces))

print('Shape of Faces matrix --> ', np.array(FACES).shape)

knn = KNeighborsClassifier(n_neighbors=5)
print("Number of samples in FACES:", len(FACES))
print("Number of samples in LABELS:", len(LABELS))

# Fit the KNeighborsClassifier with the training data and labels
knn.fit(FACES, LABELS)

# Load the background image
imgBackground = cv2.imread("C:\\Users\\USER\\Desktop\\face_recognition_project-main\\background.jpeg")

if imgBackground is None:
    raise FileNotFoundError("Background image not found. Please check the path to the background image.")

COL_NAMES = ['NAME', 'TIME']

while True:
    ret, frame = video.read()
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    faces = facedetect.detectMultiScale(gray, 1.3, 5)
    for (x, y, w, h) in faces:
        crop_img = frame[y:y+h, x:x+w, :]
        resized_img = cv2.resize(crop_img, (50, 50)).flatten().reshape(1, -1)
        output = knn.predict(resized_img)
        ts = time.time()
        date = datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
        timestamp = datetime.fromtimestamp(ts).strftime("%H:%M-%S")
        exist = os.path.isfile("Attendance/Attendance_" + date + ".csv")
        cv2.rectangle(frame, (x, y), (x+w, y+h), (0, 0, 255), 1)
        cv2.rectangle(frame, (x, y), (x+w, y+h), (50, 50, 255), 2)
        cv2.rectangle(frame, (x, y-40), (x+w, y), (50, 50, 255), -1)
        cv2.putText(frame, str(output[0]), (x, y-15), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 255, 255), 1)
        cv2.rectangle(frame, (x, y), (x+w, y+h), (50, 50, 255), 1)
        attendance = [str(output[0]), str(timestamp)]
    imgBackground[162:162 + 480, 55:55 + 640] = frame
    cv2.imshow("Frame", imgBackground)
    k = cv2.waitKey(1)
    if k == ord('o'):
        speak("Attendance Taken..")
        time.sleep(5)
        if exist:
            with open("C:\\Users\\USER\\Desktop\\face_recognition_project-main\\Attendance\\Attendance_{}.csv".format(date), 'a') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(attendance)
        else:
            with open("C:\\Users\\USER\\Desktop\\face_recognition_project-main\\Attendance\\Attendance_{}.csv".format(date), 'w', newline='') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(COL_NAMES)
                writer.writerow(attendance)
    if k == ord('q'):
        break

video.release()
cv2.destroyAllWindows()
