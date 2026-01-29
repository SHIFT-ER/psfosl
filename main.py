import sys
import os
import shutil
import re
from datetime import datetime
from pathlib import Path
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QDialog, QTableWidgetItem,
    QMessageBox, QFileDialog, QHeaderView, QLabel, QPushButton,
    QLineEdit, QHBoxLayout, QVBoxLayout, QGroupBox, QScrollArea
)
from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QPixmap, QImage
from PyQt6.uic import loadUi
from database import Database
from betta_gidroisolation import Gidroisolation_report


def validate_date(date_str):
    """Проверяет корректность даты в формате ДД.ММ.ГГ или ДД.ММ.ГГГГ"""
    if not date_str or not date_str.strip():
        return False, "Дата не может быть пустой"

    date_str = date_str.strip()
    # Проверяем формат ДД.ММ.ГГ или ДД.ММ.ГГГГ
    pattern = r'^\d{2}\.\d{2}\.\d{2,4}$'
    if not re.match(pattern, date_str):
        return False, "Неверный формат даты. Используйте ДД.ММ.ГГ или ДД.ММ.ГГГГ"

    try:
        parts = date_str.split('.')
        day = int(parts[0])
        month = int(parts[1])
        year = int(parts[2])

        if year < 100:
            year += 2000

        if month < 1 or month > 12:
            return False, "Месяц должен быть от 1 до 12"
        if day < 1 or day > 31:
            return False, "День должен быть от 1 до 31"

        datetime(year, month, day)
        return True, ""
    except ValueError as e:
        return False, f"Некорректная дата: {str(e)}"


def show_success_message(parent, message):
    """Показывает сообщение об успехе"""
    msg = QMessageBox(parent)
    msg.setIcon(QMessageBox.Icon.Information)
    msg.setWindowTitle("Успех")
    msg.setText(message)
    msg.setStandardButtons(QMessageBox.StandardButton.Ok)
    msg.button(QMessageBox.StandardButton.Ok).setText("ОК")
    msg.exec()


def show_error_message(parent, message):
    """Показывает сообщение об ошибке"""
    msg = QMessageBox(parent)
    msg.setIcon(QMessageBox.Icon.Critical)
    msg.setWindowTitle("Ошибка")
    msg.setText(message)
    msg.setStandardButtons(QMessageBox.StandardButton.Ok)
    msg.button(QMessageBox.StandardButton.Ok).setText("ОК")
    msg.exec()


def show_question_message(parent, title, message):
    """Показывает вопрос с кнопками Да/Нет на русском"""
    msg = QMessageBox(parent)
    msg.setIcon(QMessageBox.Icon.Question)
    msg.setWindowTitle(title)
    msg.setText(message)
    msg.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
    msg.button(QMessageBox.StandardButton.Yes).setText("Да")
    msg.button(QMessageBox.StandardButton.No).setText("Нет")
    return msg.exec() == QMessageBox.StandardButton.Yes


class LoginWindow(QWidget):
    def __init__(self, db, main_window):
        super().__init__()
        loadUi("login.ui", self)
        self.db = db
        self.main_window = main_window
        self.loginButton.clicked.connect(self.login)
        self.passwordEdit.returnPressed.connect(self.login)
        self.errorLabel.setText("")

    def clear_fields(self):
        """Очищает поля логина и пароля"""
        self.loginEdit.setText("")
        self.passwordEdit.setText("")
        self.errorLabel.setText("")

    def login(self):
        login = self.loginEdit.text()
        password = self.passwordEdit.text()

        if not login or not password:
            self.errorLabel.setText("Введите логин и пароль")
            return

        user = self.db.authenticate_user(login, password)
        if user:
            self.main_window.set_current_user(user)
            self.hide()
            self.main_window.show()
            self.errorLabel.setText("")
        else:
            self.errorLabel.setText("Неверный логин или пароль")


class MainWindow(QMainWindow):
    def __init__(self, db):
        super().__init__()
        loadUi("main_window.ui", self)
        self.db = db
        self.current_user = None

        self.objectsButton.clicked.connect(self.open_objects_management)
        self.devicesButton.clicked.connect(self.open_devices_management)
        self.usersButton.clicked.connect(self.open_users_management)
        self.imagesButton.clicked.connect(self.open_images_management)
        self.createReportButton.clicked.connect(self.open_create_report)
        self.logoutButton.clicked.connect(self.logout)


        # "Завершить работу"
        self.backButton.clicked.connect(self.close_application)

        # Создаём окна управления (будут показаны по требованию)
        self.objects_window = None
        self.devices_window = None
        self.users_window = None
        self.images_window = None
        self.create_report_window = None

    def set_current_user(self, user):
        self.current_user = user
        user_id, full_name, login, role, position = user
        self.welcomeLabel.setText(f"Добро пожаловать, {full_name}!")

        # Скрываем кнопки для обычных пользователей (но не для developer)
        if role not in ("admin", "developer"):
            self.usersButton.hide()
            self.imagesButton.hide()
        else:
            self.usersButton.show()
            self.imagesButton.show()



    def close_application(self):
        """Закрывает приложение (кнопка 'Завершить работу')"""
        if show_question_message(self, "Подтверждение", "Вы уверены, что хотите завершить работу (закрыть программу)?"):
            QApplication.instance().quit()

    def open_objects_management(self):
        if self.objects_window is None:
            self.objects_window = ObjectsManagementWindow(self.db, self.current_user)
        self.objects_window.show()
        self.objects_window.refresh_table()

    def open_devices_management(self):
        if self.devices_window is None:
            self.devices_window = DevicesManagementWindow(self.db, self.current_user)
        self.devices_window.show()
        self.devices_window.refresh_table()

    def open_users_management(self):
        if self.current_user and self.current_user[3] in ("admin", "developer"):
            if self.users_window is None:
                self.users_window = UsersManagementWindow(self.db, self.current_user)
            self.users_window.show()
            self.users_window.refresh_table()
        else:
            show_error_message(self, "Доступ запрещён")

    def open_images_management(self):
        if self.current_user and self.current_user[3] in ("admin", "developer"):
            if self.images_window is None:
                self.images_window = ImagesManagementWindow(self.db, self.current_user)
            self.images_window.show()
            self.images_window.load_test_types()
        else:
            show_error_message(self, "Доступ запрещён")

    def open_create_report(self):
        if self.create_report_window is None:
            self.create_report_window = CreateReportWindow(self.db, self.current_user)
        # Обновляем данные в комбобоксах при каждом открытии
        self.create_report_window.refresh_objects()
        self.create_report_window.refresh_devices()
        self.create_report_window.show()

    def logout(self):
        """Завершить сессию (выйти из аккаунта)"""
        if not show_question_message(self, "Подтверждение", "Вы уверены, что хотите завершить сессию (выйти из аккаунта)?"):
            return

        self.hide()
        # Очищаем поля логина и пароля
        login_window.clear_fields()
        login_window.show()


from PyQt6.QtWidgets import QWidget, QTableWidgetItem, QDialog
from PyQt6.uic import loadUi

class ObjectsManagementWindow(QWidget):
    def __init__(self, db, current_user):
        super().__init__()
        loadUi("objects_management.ui", self)
        self.db = db
        self.current_user = current_user

        # Подключаем кнопки к функциям
        self.addButton.clicked.connect(self.add_object)
        self.editButton.clicked.connect(self.edit_object)
        self.deleteButton.clicked.connect(self.delete_object)
        self.backButton.clicked.connect(self.close)

        # Задаём заголовки столбцов для таблицы
        self.objectsTable.setColumnCount(4)
        self.objectsTable.setHorizontalHeaderLabels(
            ["ID", "Название объекта", "Адрес", "Описание"]
        )

        # Загружаем данные
        self.refresh_table()

    def refresh_table(self):
        """Обновляет таблицу объектов"""
        objects = self.db.get_all_objects()
        self.objectsTable.setRowCount(len(objects))
        for row, obj in enumerate(objects):
            obj_id, name, address, description = obj
            self.objectsTable.setItem(row, 0, QTableWidgetItem(str(obj_id)))
            self.objectsTable.setItem(row, 1, QTableWidgetItem(name or ""))
            self.objectsTable.setItem(row, 2, QTableWidgetItem(address or ""))
            self.objectsTable.setItem(row, 3, QTableWidgetItem(description or ""))

        self.objectsTable.resizeColumnsToContents()

    def add_object(self):
        """Добавление нового объекта"""
        dialog = ObjectDialog(self, self.db)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.refresh_table()
            show_success_message(self, "Объект успешно добавлен")

    def edit_object(self):
        """Редактирование выбранного объекта"""
        current_row = self.objectsTable.currentRow()
        if current_row < 0:
            show_error_message(self, "Выберите объект для редактирования")
            return

        obj_id = int(self.objectsTable.item(current_row, 0).text())
        obj = self.db.get_object(obj_id)
        if obj:
            dialog = ObjectDialog(self, self.db, obj)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                self.refresh_table()
                show_success_message(self, "Объект успешно изменён")

    def delete_object(self):
        """Удаление выбранного объекта"""
        current_row = self.objectsTable.currentRow()
        if current_row < 0:
            show_error_message(self, "Выберите объект для удаления")
            return

        obj_id = int(self.objectsTable.item(current_row, 0).text())
        if show_question_message(self, "Подтверждение", "Вы уверены, что хотите удалить этот объект?"):
            self.db.delete_object(obj_id)
            self.refresh_table()
            show_success_message(self, "Объект успешно удалён")


from PyQt6.QtWidgets import QWidget, QTableWidgetItem, QDialog
from PyQt6.uic import loadUi

class DevicesManagementWindow(QWidget):
    def __init__(self, db, current_user):
        super().__init__()
        loadUi("devices_management.ui", self)
        self.db = db
        self.current_user = current_user

        # Подключаем кнопки к функциям
        self.addButton.clicked.connect(self.add_device)
        self.editButton.clicked.connect(self.edit_device)
        self.deleteButton.clicked.connect(self.delete_device)
        self.backButton.clicked.connect(self.close)

        # Задаём заголовки столбцов для таблицы
        self.devicesTable.setColumnCount(6)
        self.devicesTable.setHorizontalHeaderLabels(
            ["ID", "Название прибора", "Модель", "Инвентарный номер", "Действителен до", "Описание"]
        )

        # Загружаем данные
        self.refresh_table()

    def refresh_table(self):
        """Обновляет таблицу приборов"""
        devices = self.db.get_all_devices()
        self.devicesTable.setRowCount(len(devices))
        for row, device in enumerate(devices):
            device_id, name, model, inventory_number, valid_until, description = device
            self.devicesTable.setItem(row, 0, QTableWidgetItem(str(device_id)))
            self.devicesTable.setItem(row, 1, QTableWidgetItem(name or ""))
            self.devicesTable.setItem(row, 2, QTableWidgetItem(model or ""))
            self.devicesTable.setItem(row, 3, QTableWidgetItem(inventory_number or ""))
            self.devicesTable.setItem(row, 4, QTableWidgetItem(valid_until or ""))
            self.devicesTable.setItem(row, 5, QTableWidgetItem(description or ""))

        self.devicesTable.resizeColumnsToContents()

    def add_device(self):
        """Добавление нового прибора"""
        dialog = DeviceDialog(self, self.db)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.refresh_table()
            show_success_message(self, "Прибор успешно добавлен")

    def edit_device(self):
        """Редактирование выбранного прибора"""
        current_row = self.devicesTable.currentRow()
        if current_row < 0:
            show_error_message(self, "Выберите прибор для редактирования")
            return

        device_id = int(self.devicesTable.item(current_row, 0).text())
        device = self.db.get_device(device_id)
        if device:
            dialog = DeviceDialog(self, self.db, device)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                self.refresh_table()
                show_success_message(self, "Прибор успешно изменён")

    def delete_device(self):
        """Удаление выбранного прибора"""
        current_row = self.devicesTable.currentRow()
        if current_row < 0:
            show_error_message(self, "Выберите прибор для удаления")
            return

        device_id = int(self.devicesTable.item(current_row, 0).text())
        if show_question_message(self, "Подтверждение", "Вы уверены, что хотите удалить этот прибор?"):
            self.db.delete_device(device_id)
            self.refresh_table()
            show_success_message(self, "Прибор успешно удалён")


from PyQt6.QtWidgets import QWidget, QTableWidgetItem, QDialog
from PyQt6.uic import loadUi

class UsersManagementWindow(QWidget):
    def __init__(self, db, current_user):
        super().__init__()
        loadUi("users_management.ui", self)
        self.db = db
        self.current_user = current_user

        # Подключаем кнопки к функциям
        self.addButton.clicked.connect(self.add_user)
        self.editButton.clicked.connect(self.edit_user)
        self.deleteButton.clicked.connect(self.delete_user)
        self.backButton.clicked.connect(self.close)

        # Задаём заголовки столбцов для таблицы
        self.usersTable.setColumnCount(5)
        self.usersTable.setHorizontalHeaderLabels(
            ["ID", "ФИО", "Логин", "Роль", "Должность"]
        )

        # Загружаем данные
        self.refresh_table()

    def refresh_table(self):
        """Обновляет таблицу пользователей"""
        # Исключаем разработчика из списка
        users = self.db.get_all_users(exclude_developer=True)
        self.usersTable.setRowCount(len(users))
        for row, user in enumerate(users):
            user_id, full_name, login, role, position = user
            self.usersTable.setItem(row, 0, QTableWidgetItem(str(user_id)))
            self.usersTable.setItem(row, 1, QTableWidgetItem(full_name or ""))
            self.usersTable.setItem(row, 2, QTableWidgetItem(login or ""))
            self.usersTable.setItem(row, 3, QTableWidgetItem(role or ""))
            self.usersTable.setItem(row, 4, QTableWidgetItem(position or ""))

        self.usersTable.resizeColumnsToContents()

    def add_user(self):
        """Добавление нового пользователя"""
        dialog = UserDialog(self, self.db, self.current_user)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.refresh_table()
            show_success_message(self, "Пользователь успешно добавлен")

    def edit_user(self):
        """Редактирование выбранного пользователя"""
        current_row = self.usersTable.currentRow()
        if current_row < 0:
            show_error_message(self, "Выберите пользователя для редактирования")
            return

        user_id = int(self.usersTable.item(current_row, 0).text())
        user = self.db.get_user(user_id)
        if user:
            # Проверяем, является ли это разработчиком
            if user[2] == "SHIFTER":
                # Только сам разработчик может редактировать свои данные
                if not self.current_user or self.current_user[2] != "SHIFTER":
                    show_error_message(self, "Невозможно редактировать данные разработчика")
                    return
            dialog = UserDialog(self, self.db, self.current_user, user)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                self.refresh_table()
                show_success_message(self, "Пользователь успешно изменён")

    def delete_user(self):
        """Удаление выбранного пользователя"""
        current_row = self.usersTable.currentRow()
        if current_row < 0:
            show_error_message(self, "Выберите пользователя для удаления")
            return

        user_id = int(self.usersTable.item(current_row, 0).text())
        user = self.db.get_user(user_id)

        # Проверяем, является ли это разработчиком
        if user and user[2] == "SHIFTER":
            # Только сам разработчик может удалить свои данные
            if not self.current_user or self.current_user[2] != "SHIFTER":
                show_error_message(self, "Невозможно удалить разработчика")
                return

        if show_question_message(self, "Подтверждение", "Вы уверены, что хотите удалить этого пользователя?"):
            self.db.delete_user(user_id)
            self.refresh_table()
            show_success_message(self, "Пользователь успешно удалён")


class ImageItemWidget(QWidget):
    """Виджет для отображения одного изображения в списке"""
    def __init__(self, parent, db, test_type, image_name, current_user):
        super().__init__(parent)
        self.db = db
        self.test_type = test_type
        self.image_name = image_name
        self.current_user = current_user
        self.file_path = None

        # Создаём layout
        layout = QHBoxLayout(self)

        # Название изображения
        self.nameLabel = QLabel(image_name)
        self.nameLabel.setMinimumWidth(150)
        self.nameLabel.setMaximumWidth(150)
        layout.addWidget(self.nameLabel)

        # Поле пути к файлу
        self.filePathEdit = QLineEdit()
        self.filePathEdit.setReadOnly(True)
        self.filePathEdit.setPlaceholderText("Путь к файлу...")
        layout.addWidget(self.filePathEdit)

        # Кнопка выбора файла
        self.browseButton = QPushButton("Выбрать...")
        self.browseButton.clicked.connect(self.browse_file)
        layout.addWidget(self.browseButton)

        # Превью изображения
        self.previewLabel = QLabel()
        self.previewLabel.setMinimumSize(100, 100)
        self.previewLabel.setMaximumSize(100, 100)
        self.previewLabel.setStyleSheet("border: 1px solid gray;")
        self.previewLabel.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.previewLabel.setScaledContents(True)
        layout.addWidget(self.previewLabel)

        self.load_image()

    def load_image(self):
        """Загружает изображение из базы данных"""
        image = self.db.get_image_by_test_type_and_name(self.test_type, self.image_name)
        if image:
            image_id, test_type, name, file_path, description = image
            self.file_path = file_path
            self.filePathEdit.setText(file_path)
            if os.path.exists(file_path):
                pixmap = QPixmap(file_path)
                if not pixmap.isNull():
                    pixmap = pixmap.scaled(
                        100, 100,
                        Qt.AspectRatioMode.KeepAspectRatio,
                        Qt.TransformationMode.SmoothTransformation
                    )
                    self.previewLabel.setPixmap(pixmap)
                else:
                    self.previewLabel.setText("Неверный\nформат")
            else:
                self.previewLabel.setText("Файл\nне найден")
        else:
            self.filePathEdit.setText("")
            self.previewLabel.setText("Не\nвыбрано")

    def browse_file(self):
        """Открывает диалог выбора файла"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите изображение", "", "Image Files (*.png *.jpg *.jpeg *.bmp *.gif)"
        )
        if file_path:
            self.save_image(file_path)

    def save_image(self, source_path):
        """Сохраняет изображение в папку images/<test_type>/ и в базу данных"""
        try:
            # Создаём папку для изображений если её нет
            images_dir = Path("images") / self.test_type
            images_dir.mkdir(parents=True, exist_ok=True)

            # Копируем файл в папку images/<test_type>/
            file_name = os.path.basename(source_path)
            dest_path = images_dir / file_name

            # Если файл уже существует, добавляем суффикс
            counter = 1
            original_dest = dest_path
            while dest_path.exists():
                name_part = original_dest.stem
                ext = original_dest.suffix
                dest_path = images_dir / f"{name_part}_{counter}{ext}"
                counter += 1

            shutil.copy2(source_path, dest_path)
            file_path = str(dest_path)

            # Сохраняем в базу данных
            image = self.db.get_image_by_test_type_and_name(self.test_type, self.image_name)
            if image:
                # Обновляем существующее изображение
                image_id = image[0]
                old_path = image[3]
                # Удаляем старый файл если он существует и отличается от нового
                if old_path != file_path and os.path.exists(old_path):
                    try:
                        os.remove(old_path)
                    except Exception as e:
                        print(f"Ошибка при удалении старого файла: {e}")

                self.db.update_image(
                    image_id,
                    self.test_type,
                    self.image_name,
                    file_path,
                    image[4] if len(image) > 4 else ""
                )
            else:
                # Создаём новое изображение
                self.db.add_image(self.test_type, self.image_name, file_path, "")

            self.file_path = file_path
            self.filePathEdit.setText(file_path)

            # Обновляем превью
            pixmap = QPixmap(file_path)
            if not pixmap.isNull():
                pixmap = pixmap.scaled(
                    100, 100,
                    Qt.AspectRatioMode.KeepAspectRatio,
                    Qt.TransformationMode.SmoothTransformation
                )
                self.previewLabel.setPixmap(pixmap)

            show_success_message(self, f"Изображение '{self.image_name}' успешно сохранено")
        except Exception as e:
            show_error_message(self, f"Ошибка при сохранении изображения: {str(e)}")


class ImagesManagementWindow(QWidget):
    def __init__(self, db, current_user):
        super().__init__()
        loadUi("images_management.ui", self)
        self.db = db
        self.current_user = current_user

        # Стандартные названия изображений для каждого типа испытания
        self.image_names = {
            "Гидроизоляция": ["шапка организации", "формула"]
        }

        self.testTypeCombo.currentTextChanged.connect(self.on_test_type_changed)
        self.browseTemplateButton.clicked.connect(self.browse_template)
        self.clearTemplateButton.clicked.connect(self.clear_template)
        self.deleteTestTypeButton.clicked.connect(self.delete_test_type)
        self.backButton.clicked.connect(self.close)

        self.load_test_types()
        # Скрываем элементы, связанные с шаблоном Excel
        self.browseTemplateButton.hide()
        self.clearTemplateButton.hide()
        self.templatePathEdit.hide()  # Скрываем поле с путем к шаблону
    def load_test_types(self):
        """Загружает типы испытаний в комбобокс"""
        self.testTypeCombo.clear()
        test_types = self.db.get_all_test_types()
        if not test_types:
            # Если типов нет, добавляем "Гидроизоляция" по умолчанию
            self.testTypeCombo.addItem("Гидроизоляция")
        else:
            for test_type in test_types:
                self.testTypeCombo.addItem(test_type)

        # Добавляем возможность создать новый тип
        self.testTypeCombo.addItem("+ Добавить новый тип...")
        self.on_test_type_changed()

    def on_test_type_changed(self):
        """Обработчик изменения типа испытания"""
        current_type = self.testTypeCombo.currentText()
        if current_type == "+ Добавить новый тип...":
            # Добавляем новый тип
            from PyQt6.QtWidgets import QInputDialog
            new_type, ok = QInputDialog.getText(self, "Новый тип испытания", "Введите название типа испытания:")
            if ok and new_type.strip():
                new_type = new_type.strip()
                # Добавляем в комбобокс перед пунктом "+ Добавить новый тип..."
                count = self.testTypeCombo.count()
                self.testTypeCombo.insertItem(count - 1, new_type)
                self.testTypeCombo.setCurrentText(new_type)
                # Добавляем стандартные названия изображений для нового типа
                if new_type not in self.image_names:
                    self.image_names[new_type] = []
                self.load_images_for_type(new_type)
                self.load_template_for_type(new_type)
            else:
                # Возвращаемся к предыдущему типу
                if self.testTypeCombo.count() > 1:
                    self.testTypeCombo.setCurrentIndex(0)
            return

        # Включаем/выключаем кнопку удаления в зависимости от выбранного типа
        if current_type == "+ Добавить новый тип...":
            self.deleteTestTypeButton.setEnabled(False)
        else:
            self.deleteTestTypeButton.setEnabled(True)

        self.load_images_for_type(current_type)
        self.load_template_for_type(current_type)

    def load_images_for_type(self, test_type):
        """Загружает изображения для выбранного типа испытания"""
        # Очищаем предыдущие виджеты
        scroll_widget = self.scrollArea.widget()
        layout = scroll_widget.layout()

        # Удаляем все виджеты из layout
        while layout.count():
            child = layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()

        # Получаем названия изображений для этого типа
        image_names = self.image_names.get(test_type, [])
        if not image_names:
            # Если названий нет, получаем из базы данных
            images = self.db.get_images_by_test_type(test_type)
            image_names = list(set([img[2] for img in images]))
            # Сохраняем в словарь
            if image_names:
                self.image_names[test_type] = image_names

        # Если названий всё ещё нет и это "Гидроизоляция", используем стандартные
        if not image_names and test_type == "Гидроизоляция":
            image_names = ["шапка организации", "формула"]
            self.image_names[test_type] = image_names

        # Создаём виджеты для каждого изображения
        for image_name in image_names:
            widget = ImageItemWidget(self, self.db, test_type, image_name, self.current_user)
            layout.addWidget(widget)

        # Добавляем возможность добавить новое название изображения
        add_button = QPushButton("+ Добавить новое изображение")
        add_button.clicked.connect(lambda checked, tt=test_type: self.add_new_image_name(tt))
        layout.addWidget(add_button)

        layout.addStretch()

    def add_new_image_name(self, test_type):
        """Добавляет новое название изображения"""
        from PyQt6.QtWidgets import QInputDialog
        name, ok = QInputDialog.getText(self, "Новое изображение", "Введите название изображения:")
        if ok and name.strip():
            name = name.strip()
            # Добавляем в список названий
            if test_type not in self.image_names:
                self.image_names[test_type] = []
            if name not in self.image_names[test_type]:
                self.image_names[test_type].append(name)
                self.load_images_for_type(test_type)
                show_success_message(self, f"Изображение '{name}' добавлено")
            else:
                show_error_message(self, f"Изображение '{name}' уже существует")

    def load_template_for_type(self, test_type):
        """Загружает шаблон для типа испытания"""
        template = self.db.get_template(test_type)
        if template:
            template_id, test_type, template_path, description = template
            self.templatePathEdit.setText(template_path)
        else:
            self.templatePathEdit.setText("")

    def browse_template(self):
        """Открывает диалог выбора шаблона Excel"""
        test_type = self.testTypeCombo.currentText()
        if test_type == "+ Добавить новый тип...":
            show_error_message(self, "Выберите тип испытания")
            return

        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите шаблон Excel", "", "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            try:
                # Копируем шаблон в папку templates если нужно
                templates_dir = Path("templates")
                templates_dir.mkdir(parents=True, exist_ok=True)

                file_name = os.path.basename(file_path)
                dest_path = templates_dir / f"{test_type}_{file_name}"

                shutil.copy2(file_path, dest_path)

                # Сохраняем в базу данных
                self.db.add_template(test_type, str(dest_path), "")
                self.templatePathEdit.setText(str(dest_path))
                show_success_message(self, "Шаблон успешно сохранён")
            except Exception as e:
                show_error_message(self, f"Ошибка при сохранении шаблона: {str(e)}")

    def clear_template(self):
        """Очищает шаблон для типа испытания"""
        test_type = self.testTypeCombo.currentText()
        if test_type == "+ Добавить новый тип...":
            return

        if show_question_message(self, "Подтверждение", "Вы уверены, что хотите очистить шаблон?"):
            # Удаляем шаблон из базы данных
            self.db.delete_template(test_type)
            self.templatePathEdit.setText("")
            show_success_message(self, "Шаблон очищен")

    def delete_test_type(self):
        """Удаляет тип испытания и все связанные данные"""
        test_type = self.testTypeCombo.currentText()
        if test_type == "+ Добавить новый тип...":
            show_error_message(self, "Выберите тип испытания для удаления")
            return

        if not show_question_message(
            self,
            "Подтверждение удаления",
            f"Вы уверены, что хотите удалить тип испытания '{test_type}'?\n\n"
            f"Это действие удалит:\n- Все изображения этого типа\n- Шаблон отчёта для этого типа\n\n"
            f"Это действие нельзя отменить!"
        ):
            return

        try:
            # Удаляем тип испытания из базы данных
            self.db.delete_test_type(test_type)

            # Удаляем из словаря image_names
            if test_type in self.image_names:
                del self.image_names[test_type]

            # Обновляем комбобокс
            index = self.testTypeCombo.findText(test_type)
            if index >= 0:
                self.testTypeCombo.removeItem(index)

            # Выбираем первый доступный тип или "Гидроизоляция"
            if self.testTypeCombo.count() > 1:
                # Пропускаем последний элемент "+ Добавить новый тип..."
                if self.testTypeCombo.itemText(0) != "+ Добавить новый тип...":
                    self.testTypeCombo.setCurrentIndex(0)
                else:
                    self.testTypeCombo.setCurrentIndex(1)
            else:
                # Если остался только "+ Добавить новый тип...", добавляем "Гидроизоляция"
                self.testTypeCombo.insertItem(0, "Гидроизоляция")
                self.testTypeCombo.setCurrentIndex(0)

            # Обновляем интерфейс
            self.on_test_type_changed()

            show_success_message(self, f"Тип испытания '{test_type}' успешно удалён")
        except Exception as e:
            show_error_message(self, f"Ошибка при удалении типа испытания: {str(e)}")


class CreateReportWindow(QWidget):
    def __init__(self, db, current_user):
        super().__init__()
        loadUi("create_report.ui", self)
        self.db = db
        self.current_user = current_user
        self.createButton.clicked.connect(self.create_report)
        self.backButton.clicked.connect(self.close)

        # Заполняем комбобоксы
        self.refresh_objects()
        self.refresh_devices()

        # Устанавливаем значения по умолчанию
        if self.current_user:
            user_id, full_name, login, role, position = self.current_user
            user_info = f"{position or ''} {full_name}".strip()
            self.userInfoEdit.setText(user_info)

        self.statusLabel.setText("")

    def refresh_objects(self):
        objects = self.db.get_all_objects()
        self.objectCombo.clear()
        for obj in objects:
            obj_id, name, address, description = obj
            self.objectCombo.addItem(f"{name} ({address})", obj_id)

    def refresh_devices(self):
        devices = self.db.get_all_devices()
        self.deviceCombo.clear()
        for device in devices:
            device_id, name, model, inventory_number, valid_until, description = device
            display_text = f"{name}"
            if model:
                display_text += f" ({model})"
            if inventory_number:
                display_text += f" - {inventory_number}"
            self.deviceCombo.addItem(display_text, device_id)

    def create_report(self):
        # Получаем данные из формы
        test_type = self.testTypeCombo.currentText()
        object_index = self.objectCombo.currentIndex()
        device_index = self.deviceCombo.currentIndex()

        # Проверяем, что есть объекты и приборы в базе
        if self.objectCombo.count() == 0:
            show_error_message(self, "В базе данных нет объектов. Добавьте объекты в разделе 'Управление объектами'")
            return

        if self.deviceCombo.count() == 0:
            show_error_message(self, "В базе данных нет приборов. Добавьте приборы в разделе 'Управление приборами'")
            return

        if object_index < 0:
            show_error_message(self, "Выберите объект")
            return

        if device_index < 0:
            show_error_message(self, "Выберите прибор")
            return

        # Получаем данные объекта и прибора
        obj_id = self.objectCombo.currentData()
        device_id = self.deviceCombo.currentData()

        if obj_id is None or device_id is None:
            show_error_message(self, "Ошибка при получении данных объекта или прибора")
            return

        obj = self.db.get_object(obj_id)
        device = self.db.get_device(device_id)

        if not obj or not device:
            show_error_message(self, "Ошибка при получении данных объекта или прибора")
            return

        obj_id, object_name, address, description = obj
        device_id, device_name, model, inventory_number, valid_until, device_desc = device

        # Получаем остальные данные
        result_file_name = self.resultFileNameEdit.text() or "Отчёт_по_гидроизоляции.xlsx"
        page_name = self.pageNameEdit.text() or "Участок"
        client_name = self.clientNameEdit.text()
        contract_name = self.contractNameEdit.text()
        plot_location = self.plotLocationEdit.text()
        work_date = self.workDateEdit.text()
        plot_name = self.plotNameEdit.text()
        square = self.squareSpinBox.value()
        values_str = self.valuesEdit.text()
        user_info = self.userInfoEdit.text()

        # Проверяем обязательные поля
        if not client_name or not contract_name or not plot_location or not work_date or not plot_name:
            show_error_message(self, "Заполните все обязательные поля")
            return

        # Проверяем дату работ
        is_valid, error_msg = validate_date(work_date)
        if not is_valid:
            show_error_message(self, f"Ошибка в дате проведения работ: {error_msg}")
            return

        # Парсим значения силы
        try:
            values = [float(v.strip()) for v in values_str.split(",") if v.strip()]
            if not values:
                show_error_message(self, "Введите значения силы")
                return
        except ValueError:
            show_error_message(self, "Неверный формат значений силы. Используйте числа, разделённые запятыми")
            return

        try:
            # Получаем изображения для данного типа испытания
            images = self.db.get_images_by_test_type(test_type)

            # Ищем изображения для шапки организации и формулы
            org_header = None
            formula_image = None

            for image in images:
                image_id, img_test_type, name, img_path, description = image
                if not os.path.exists(img_path):
                    continue

                name_lower = name.lower()
                if "шапка" in name_lower or "организации" in name_lower or "header" in name_lower:
                    org_header = img_path
                elif "формул" in name_lower or "formula" in name_lower:
                    formula_image = img_path

            # Получаем шаблон для типа испытания
            template = self.db.get_template(test_type)
            template_path = None
            if template:
                template_id, test_type, template_path, description = template

            # Создаём отчёт
            report = Gidroisolation_report(
                r_f=result_file_name,
                p_n=page_name,
                o_n=object_name,
                client_n=client_name,
                contract_n=contract_name,
                d_n=device_name,
                z_n=inventory_number or "777",
                d_v_u=valid_until or "01.01.01",
                p_l=plot_location,
                w_d=work_date,
                p_name=plot_name,
                s=square,
                v=values
            )

            # Создаём отчёт с изображениями и шаблоном
            report.create_gidroisolation_report(page_name, org_header, formula_image, template_path)

            # Вставляем информацию о пользователе в ячейку D37
            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            file_path = os.path.join(desktop_path, result_file_name)

            if os.path.exists(file_path):
                import openpyxl
                wb = openpyxl.load_workbook(file_path)
                ws = wb.active
                ws["D37"] = user_info
                wb.save(file_path)

            self.statusLabel.setText(f"Отчёт успешно создан: {file_path}")
            self.statusLabel.setStyleSheet("color: green;")
            show_success_message(self, f"Отчёт успешно создан:\n{file_path}")

        except Exception as e:
            self.statusLabel.setText(f"Ошибка: {str(e)}")
            self.statusLabel.setStyleSheet("color: red;")
            show_error_message(self, f"Ошибка при создании отчёта:\n{str(e)}")
            import traceback
            traceback.print_exc()


# Диалоговые окна для добавления/редактирования

class ObjectDialog(QDialog):
    def __init__(self, parent, db, obj=None):
        super().__init__(parent)
        loadUi("object_dialog.ui", self)
        self.db = db
        self.obj = obj
        self.okButton.clicked.connect(self.accept)
        self.cancelButton.clicked.connect(self.reject)

        if obj:
            obj_id, name, address, description = obj
            self.nameEdit.setText(name or "")
            self.addressEdit.setText(address or "")
            self.descriptionEdit.setPlainText(description or "")

    def accept(self):
        name = self.nameEdit.text().strip()
        if not name:
            show_error_message(self, "Введите название объекта")
            return

        address = self.addressEdit.text().strip()
        description = self.descriptionEdit.toPlainText().strip()

        try:
            if self.obj:
                obj_id = self.obj[0]
                self.db.update_object(obj_id, name, address, description)
            else:
                self.db.add_object(name, address, description)

            super().accept()
        except Exception as e:
            show_error_message(self, f"Ошибка при сохранении объекта: {str(e)}")


class DeviceDialog(QDialog):
    def __init__(self, parent, db, device=None):
        super().__init__(parent)
        loadUi("device_dialog.ui", self)
        self.db = db
        self.device = device
        self.okButton.clicked.connect(self.accept)
        self.cancelButton.clicked.connect(self.reject)

        if device:
            device_id, name, model, inventory_number, valid_until, description = device
            self.nameEdit.setText(name or "")
            self.modelEdit.setText(model or "")
            self.inventoryNumberEdit.setText(inventory_number or "")
            self.validUntilEdit.setText(valid_until or "")
            self.descriptionEdit.setPlainText(description or "")

    def accept(self):
        name = self.nameEdit.text().strip()
        if not name:
            show_error_message(self, "Введите название прибора")
            return

        model = self.modelEdit.text().strip()
        inventory_number = self.inventoryNumberEdit.text().strip()
        valid_until = self.validUntilEdit.text().strip()
        description = self.descriptionEdit.toPlainText().strip()

        # Проверяем дату
        if valid_until:
            is_valid, error_msg = validate_date(valid_until)
            if not is_valid:
                show_error_message(self, f"Ошибка в дате 'Действителен до': {error_msg}")
                return

        try:
            if self.device:
                device_id = self.device[0]
                self.db.update_device(device_id, name, model, inventory_number, valid_until, description)
            else:
                self.db.add_device(name, model, inventory_number, valid_until, description)

            super().accept()
        except Exception as e:
            show_error_message(self, f"Ошибка при сохранении прибора: {str(e)}")


class UserDialog(QDialog):
    def __init__(self, parent, db, current_user, user=None):
        super().__init__(parent)
        loadUi("user_dialog.ui", self)
        self.db = db
        self.current_user = current_user
        self.user = user
        self.okButton.clicked.connect(self.accept)
        self.cancelButton.clicked.connect(self.reject)
        self.is_edit = user is not None

        # Удаляем роль developer из комбобокса
        self.roleCombo.clear()
        self.roleCombo.addItem("user")
        self.roleCombo.addItem("admin")

        if user:
            user_id, full_name, login, password_hash, role, position = user
            self.fullNameEdit.setText(full_name or "")
            self.loginEdit.setText(login or "")
            self.passwordEdit.setPlaceholderText("Оставьте пустым, чтобы не менять")
            if role == "developer":
                # Для разработчика показываем специальный текст
                self.roleCombo.addItem("developer")
                self.roleCombo.setCurrentText("developer")
                self.roleCombo.setEnabled(False)  # Запрещаем изменение роли разработчика
            else:
                self.roleCombo.setCurrentText(role or "user")
            self.positionEdit.setText(position or "")
        else:
            self.passwordEdit.setPlaceholderText("Введите пароль")

    def accept(self):
        full_name = self.fullNameEdit.text().strip()
        login = self.loginEdit.text().strip()
        password = self.passwordEdit.text().strip()
        role = self.roleCombo.currentText()
        position = self.positionEdit.text().strip()

        if not full_name or not login:
            show_error_message(self, "Введите ФИО и логин")
            return

        if not self.is_edit and not password:
            show_error_message(self, "Введите пароль")
            return

        try:
            if self.user:
                user_id = self.user[0]
                # Не позволяем изменять логин разработчика
                if self.user[2] == "SHIFTER" and login != "SHIFTER":
                    show_error_message(self, "Невозможно изменить логин разработчика")
                    return

                if password:
                    self.db.update_user(user_id, full_name, login, password, role, position)
                else:
                    self.db.update_user(user_id, full_name, login, None, role, position)
            else:
                self.db.add_user(full_name, login, password, role, position)

            super().accept()
        except Exception as e:
            show_error_message(self, f"Ошибка при сохранении пользователя: {str(e)}")


def main():
    app = QApplication(sys.argv)

    # Создаём базу данных
    db = Database()

    # Создаём главное окно
    main_window = MainWindow(db)

    # Создаём окно авторизации
    global login_window
    login_window = LoginWindow(db, main_window)
    login_window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
