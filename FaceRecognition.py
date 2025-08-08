import cv2
import face_recognition
import os
from win32com.client import Dispatch
import customtkinter as CTk
from screeninfo import get_monitors
import PIL


img_directory = ''


def face_capture(folder_dir):
    '''Функция отвечает за обработку изображения с камеры в реальном времени и распознанию лиц из фотосета'''
    known_faces_encoding = []
    known_names = []

    for x in get_monitors():
        width = x.width
        height = x.height

    video = cv2.VideoCapture(0)
    for images in os.listdir(folder_dir):
        image_accesed = face_recognition.api.load_image_file(f"{folder_dir}\{images}")
        acceced_encoding = face_recognition.api.face_encodings(image_accesed)[0]
        known_faces_encoding.append(acceced_encoding)
        known_names.append(f"{images.split('.')[0]}")

    while True:
        ret, frame = video.read()
        face_locations = face_recognition.api.face_locations(frame, model='MMOD')
        face_encodings = face_recognition.api.face_encodings(frame, face_locations)

        for (top, right, bottom, left), acceced_encoding in zip(face_locations, face_encodings):
            mathes = face_recognition.api.compare_faces(known_faces_encoding, acceced_encoding)
            name = "Не распознано"
            if mathes.__contains__(True):
                first_match_index = mathes.index(True)
                name = known_names[first_match_index]
            color = (0, 252, 124)
            if name == "Не распознано":
                color = (0, 0, 255)

            cv2.rectangle(frame, (left, top), (right, bottom), color, 2)
            cv2.putText(frame, name, (left, top - 10), cv2.FONT_HERSHEY_COMPLEX, 0.9, color, 2)
        frame = cv2.resize(frame, (width, height))
        cv2.imshow("Video", frame)

        if cv2.waitKey(1) == ord("q"):
            break
    video.release()
    cv2.destroyAllWindows()


def main():
    '''Основная функция программы'''
    def start():
        '''Функция открывает окно Windows с выбором папки. Также в ней обрабатыв'''
        global img_directory
        try:
            warn_text.configure(text="Нажмите 'Q' для закрытия окна распознания лиц", text_color='white')
            face_capture(str(img_directory))
        except FileNotFoundError or AttributeError:
            warn_text.configure(text='Похоже, что Вы не выбрали папку. Попробуйте еще раз', text_color='red')
        except PIL.UnidentifiedImageError:
            warn_text.configure(text='Похоже, что в выбранной папке нет изображений. Попробуйте еще раз', text_color='red')
        except PermissionError:
            warn_text.configure(text='Похоже, что доступ к папке отсутствует. Попробуйте еще раз', text_color='red')

    def choose_folder():
        global img_directory
        directory = Dispatch('Shell.Application').BrowseForFolder(0, 'Выберите папку', 1, '')
        img_directory = directory.Self.path

    app = CTk.CTk()
    CTk.set_appearance_mode("Dark")
    CTk.set_default_color_theme("green")

    app.geometry('720x510')
    app.title('Face recognition tool')
    app.resizable(False, False)

    start_text = CTk.CTkLabel(
        app,
        text='Добро пожаловать в инструмент для распознования лиц.\nДля начала работы выберите папку с фотографиями учеников',
        font=("Helvetica", 18),
        width=15
    )
    warn_text = CTk.CTkLabel(
        app,
        text='',
        text_color='red'
    )
    start_button = CTk.CTkButton(
        app,
        text="Начать",
        command=start,
        width=150
    )
    folder_button = CTk.CTkButton(
        app,
        text='Выбор папки',
        height=20,
        command=choose_folder
    )
    help_text = CTk.CTkLabel(
        app,
        text='В выбранной папке должны находиться изображения учеников. \nНазвание файла должно содержать имя и фамилию учащегося.'
    )

    start_text.pack(side='top', pady=40)
    folder_button.pack()
    help_text.pack(pady=20)
    start_button.pack(pady=15)
    warn_text.pack()
    app.mainloop()
    app.set


if __name__ == "__main__":
    main()