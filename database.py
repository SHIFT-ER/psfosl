import sqlite3
import os
import hashlib
from pathlib import Path


class Database:
    def __init__(self, db_path="laboratory.db"):
        self.db_path = db_path
        self.init_database()
    
    def get_connection(self):
        """Создаёт и возвращает соединение с базой данных"""
        return sqlite3.connect(self.db_path)
    
    def init_database(self):
        """Инициализирует базу данных и создаёт таблицы если их нет"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS objects (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                address TEXT,
                description TEXT
            )
        ''')
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS devices (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                model TEXT,
                inventory_number TEXT UNIQUE,
                valid_until TEXT,
                description TEXT
            )
        ''')
        
        cursor.execute('''
            SELECT name FROM sqlite_master 
            WHERE type='table' AND name='users'
        ''')
        table_exists = cursor.fetchone() is not None
        
        if table_exists:
            try:
                cursor.execute('SELECT * FROM users')
                old_users = cursor.fetchall()
                
                cursor.execute('DROP TABLE IF EXISTS users_backup')
                cursor.execute('''
                    CREATE TABLE users_backup (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        full_name TEXT NOT NULL,
                        login TEXT UNIQUE NOT NULL,
                        password_hash TEXT NOT NULL,
                        role TEXT NOT NULL,
                        position TEXT
                    )
                ''')
                
                for user in old_users:
                    user_id, full_name, login, password_hash, role, position = user
                    cursor.execute('''
                        INSERT INTO users_backup (id, full_name, login, password_hash, role, position)
                        VALUES (?, ?, ?, ?, ?, ?)
                    ''', (user_id, full_name, login, password_hash, role, position))
                
                cursor.execute('DROP TABLE users')
                conn.commit()
            except Exception as e:
                conn.rollback()
                print(f"Ошибка при миграции таблицы users: {e}")
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                full_name TEXT NOT NULL,
                login TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                role TEXT NOT NULL CHECK(role IN ('admin', 'user', 'developer')),
                position TEXT
            )
        ''')
        
        if table_exists:
            try:
                cursor.execute('SELECT name FROM sqlite_master WHERE type="table" AND name="users_backup"')
                if cursor.fetchone():
                    cursor.execute('SELECT * FROM users_backup')
                    backup_users = cursor.fetchall()
                    for user in backup_users:
                        user_id, full_name, login, password_hash, role, position = user
                        if role not in ('admin', 'user', 'developer'):
                            role = 'user'
                        try:
                            cursor.execute('''
                                INSERT OR IGNORE INTO users (id, full_name, login, password_hash, role, position)
                                VALUES (?, ?, ?, ?, ?, ?)
                            ''', (user_id, full_name, login, password_hash, role, position))
                        except sqlite3.IntegrityError:
                            try:
                                cursor.execute('''
                                    UPDATE users 
                                    SET full_name=?, password_hash=?, role=?, position=?
                                    WHERE login=?
                                ''', (full_name, password_hash, role, position, login))
                            except:
                                pass
                    cursor.execute('DROP TABLE users_backup')
                    conn.commit()
            except Exception as e:
                print(f"Ошибка при восстановлении данных: {e}")
                conn.rollback()
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS images (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                test_type TEXT NOT NULL,
                name TEXT NOT NULL,
                file_path TEXT NOT NULL UNIQUE,
                description TEXT
            )
        ''')
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS report_templates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                test_type TEXT NOT NULL UNIQUE,
                template_path TEXT NOT NULL,
                description TEXT
            )
        ''')
        
        conn.commit()
        conn.close()
        
        self.create_default_admin()
        self.create_developer()
    
    def create_default_admin(self):
        """Создаёт администратора по умолчанию (логин: admin, пароль: admin)"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute("SELECT COUNT(*) FROM users WHERE login = ?", ("admin",))
        count = cursor.fetchone()[0]
        
        if count == 0:
            password_hash = self.hash_password("admin")
            cursor.execute('''
                INSERT INTO users (full_name, login, password_hash, role, position)
                VALUES (?, ?, ?, ?, ?)
            ''', ("Администратор", "admin", password_hash, "admin", "Администратор"))
            conn.commit()
        
        conn.close()
    
    def create_developer(self):
        """Создаёт разработчика (логин: SHIFTER, пароль: tiudi1029)"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute("SELECT COUNT(*) FROM users WHERE login = ?", ("SHIFTER",))
        count = cursor.fetchone()[0]
        
        if count == 0:
            password_hash = self.hash_password("tiudi1029")
            cursor.execute('''
                INSERT INTO users (full_name, login, password_hash, role, position)
                VALUES (?, ?, ?, ?, ?)
            ''', ("Разработчик", "SHIFTER", password_hash, "developer", "Разработчик"))
            conn.commit()
        
        conn.close()
    
    @staticmethod
    def hash_password(password):
        """Хэширует пароль"""
        return hashlib.sha256(password.encode()).hexdigest()
    
    # ========== Методы для работы с объектами ==========
    
    def get_all_objects(self):
        """Возвращает все объекты"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM objects ORDER BY id")
        objects = cursor.fetchall()
        conn.close()
        return objects
    
    def add_object(self, name, address="", description=""):
        """Добавляет объект"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO objects (name, address, description)
            VALUES (?, ?, ?)
        ''', (name, address, description))
        conn.commit()
        obj_id = cursor.lastrowid
        conn.close()
        return obj_id
    
    def update_object(self, obj_id, name, address="", description=""):
        """Обновляет объект"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            UPDATE objects
            SET name = ?, address = ?, description = ?
            WHERE id = ?
        ''', (name, address, description, obj_id))
        conn.commit()
        conn.close()
    
    def delete_object(self, obj_id):
        """Удаляет объект"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM objects WHERE id = ?", (obj_id,))
        conn.commit()
        conn.close()
    
    def get_object(self, obj_id):
        """Возвращает объект по ID"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM objects WHERE id = ?", (obj_id,))
        obj = cursor.fetchone()
        conn.close()
        return obj
    
    # ========== Методы для работы с приборами ==========
    
    def get_all_devices(self):
        """Возвращает все приборы"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM devices ORDER BY id")
        devices = cursor.fetchall()
        conn.close()
        return devices
    
    def add_device(self, name, model="", inventory_number="", valid_until="", description=""):
        """Добавляет прибор"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO devices (name, model, inventory_number, valid_until, description)
            VALUES (?, ?, ?, ?, ?)
        ''', (name, model, inventory_number, valid_until, description))
        conn.commit()
        device_id = cursor.lastrowid
        conn.close()
        return device_id
    
    def update_device(self, device_id, name, model="", inventory_number="", valid_until="", description=""):
        """Обновляет прибор"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            UPDATE devices
            SET name = ?, model = ?, inventory_number = ?, valid_until = ?, description = ?
            WHERE id = ?
        ''', (name, model, inventory_number, valid_until, description, device_id))
        conn.commit()
        conn.close()
    
    def delete_device(self, device_id):
        """Удаляет прибор"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM devices WHERE id = ?", (device_id,))
        conn.commit()
        conn.close()
    
    def get_device(self, device_id):
        """Возвращает прибор по ID"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM devices WHERE id = ?", (device_id,))
        device = cursor.fetchone()
        conn.close()
        return device
    
    # ========== Методы для работы с пользователями ==========
    
    def get_all_users(self, exclude_developer=True):
        """Возвращает всех пользователей, исключая разработчика по умолчанию"""
        conn = self.get_connection()
        cursor = conn.cursor()
        if exclude_developer:
            cursor.execute("SELECT id, full_name, login, role, position FROM users WHERE login != ? ORDER BY id", ("SHIFTER",))
        else:
            cursor.execute("SELECT id, full_name, login, role, position FROM users ORDER BY id")
        users = cursor.fetchall()
        conn.close()
        return users
    
    def add_user(self, full_name, login, password, role="user", position=""):
        """Добавляет пользователя"""
        conn = self.get_connection()
        cursor = conn.cursor()
        password_hash = self.hash_password(password)
        cursor.execute('''
            INSERT INTO users (full_name, login, password_hash, role, position)
            VALUES (?, ?, ?, ?, ?)
        ''', (full_name, login, password_hash, role, position))
        conn.commit()
        user_id = cursor.lastrowid
        conn.close()
        return user_id
    
    def update_user(self, user_id, full_name, login, password=None, role="user", position=""):
        """Обновляет пользователя"""
        conn = self.get_connection()
        cursor = conn.cursor()
        if password:
            password_hash = self.hash_password(password)
            cursor.execute('''
                UPDATE users
                SET full_name = ?, login = ?, password_hash = ?, role = ?, position = ?
                WHERE id = ?
            ''', (full_name, login, password_hash, role, position, user_id))
        else:
            cursor.execute('''
                UPDATE users
                SET full_name = ?, login = ?, role = ?, position = ?
                WHERE id = ?
            ''', (full_name, login, role, position, user_id))
        conn.commit()
        conn.close()
    
    def delete_user(self, user_id):
        """Удаляет пользователя"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM users WHERE id = ?", (user_id,))
        conn.commit()
        conn.close()
    
    def get_user(self, user_id):
        """Возвращает пользователя по ID"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM users WHERE id = ?", (user_id,))
        user = cursor.fetchone()
        conn.close()
        return user
    
    def authenticate_user(self, login, password):
        """Проверяет логин и пароль, возвращает пользователя или None"""
        conn = self.get_connection()
        cursor = conn.cursor()
        password_hash = self.hash_password(password)
        cursor.execute('''
            SELECT id, full_name, login, role, position
            FROM users
            WHERE login = ? AND password_hash = ?
        ''', (login, password_hash))
        user = cursor.fetchone()
        conn.close()
        return user
    
    # ========== Методы для работы с изображениями ==========
    
    def get_all_images(self):
        """Возвращает все изображения"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM images ORDER BY test_type, name")
        images = cursor.fetchall()
        conn.close()
        return images
    
    def get_images_by_test_type(self, test_type):
        """Возвращает изображения по типу испытания"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM images WHERE test_type = ? ORDER BY name", (test_type,))
        images = cursor.fetchall()
        conn.close()
        return images
    
    def add_image(self, test_type, name, file_path, description=""):
        """Добавляет изображение"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO images (test_type, name, file_path, description)
            VALUES (?, ?, ?, ?)
        ''', (test_type, name, file_path, description))
        conn.commit()
        image_id = cursor.lastrowid
        conn.close()
        return image_id
    
    def update_image(self, image_id, test_type, name, file_path, description=""):
        """Обновляет изображение"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            UPDATE images
            SET test_type = ?, name = ?, file_path = ?, description = ?
            WHERE id = ?
        ''', (test_type, name, file_path, description, image_id))
        conn.commit()
        conn.close()
    
    def delete_image(self, image_id):
        """Удаляет изображение"""
        conn = self.get_connection()
        cursor = conn.cursor()

        cursor.execute("SELECT file_path FROM images WHERE id = ?", (image_id,))
        result = cursor.fetchone()
        file_path = result[0] if result else None
        
        cursor.execute("DELETE FROM images WHERE id = ?", (image_id,))
        conn.commit()
        conn.close()
        
        if file_path and os.path.exists(file_path):
            try:
                os.remove(file_path)
            except Exception as e:
                print(f"Ошибка при удалении файла: {e}")
    
    def get_image(self, image_id):
        """Возвращает изображение по ID"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM images WHERE id = ?", (image_id,))
        image = cursor.fetchone()
        conn.close()
        return image
    
    def get_image_by_test_type_and_name(self, test_type, name):
        """Возвращает изображение по типу испытания и названию"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM images WHERE test_type = ? AND name = ?", (test_type, name))
        image = cursor.fetchone()
        conn.close()
        return image
    
    def get_all_test_types(self):
        """Возвращает все типы испытаний из изображений и шаблонов"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute("SELECT DISTINCT test_type FROM images ORDER BY test_type")
        test_types_images = [row[0] for row in cursor.fetchall()]
        
        cursor.execute("SELECT DISTINCT test_type FROM report_templates ORDER BY test_type")
        test_types_templates = [row[0] for row in cursor.fetchall()]
        
        all_types = list(set(test_types_images + test_types_templates))
        all_types.sort()
        
        conn.close()
        return all_types
    
    
    def add_template(self, test_type, template_path, description=""):
        """Добавляет шаблон для типа испытания"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            INSERT OR REPLACE INTO report_templates (test_type, template_path, description)
            VALUES (?, ?, ?)
        ''', (test_type, template_path, description))
        conn.commit()
        conn.close()
    
    def get_template(self, test_type):
        """Возвращает шаблон для типа испытания"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM report_templates WHERE test_type = ?", (test_type,))
        template = cursor.fetchone()
        conn.close()
        return template
    
    def update_template(self, test_type, template_path, description=""):
        """Обновляет шаблон для типа испытания"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            UPDATE report_templates
            SET template_path = ?, description = ?
            WHERE test_type = ?
        ''', (template_path, description, test_type))
        conn.commit()
        conn.close()
    
    def delete_template(self, test_type):
        """Удаляет шаблон для типа испытания"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM report_templates WHERE test_type = ?", (test_type,))
        conn.commit()
        conn.close()
    
    def delete_test_type(self, test_type):
        """Удаляет все данные для типа испытания (изображения и шаблоны)"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute("SELECT file_path FROM images WHERE test_type = ?", (test_type,))
        image_paths = [row[0] for row in cursor.fetchall()]
        
        cursor.execute("DELETE FROM images WHERE test_type = ?", (test_type,))
        
        cursor.execute("DELETE FROM report_templates WHERE test_type = ?", (test_type,))
        
        conn.commit()
        conn.close()
        
        for file_path in image_paths:
            if file_path and os.path.exists(file_path):
                try:
                    os.remove(file_path)
                except Exception as e:
                    print(f"Ошибка при удалении файла {file_path}: {e}")
        

        images_dir = Path("images") / test_type
        if images_dir.exists():
            try:
                if not any(images_dir.iterdir()):  
                    images_dir.rmdir()
            except Exception as e:
                print(f"Ошибка при удалении папки {images_dir}: {e}")

