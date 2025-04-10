import os
import shutil
import json
import zipfile
from datetime import datetime
import socket
import re
import traceback
import logging
import sublime
import sublime_plugin
import urllib.parse

# Глобальная переменная для хранения текущего номера задачи
CURRENT_TASK_NUMBER = None

class FtpBackupLogger:
    def __init__(self, backup_root):
        """Настройка логирования"""
        log_dir = os.path.join(backup_root, 'logs')
        os.makedirs(log_dir, exist_ok=True)
        log_file = os.path.join(log_dir, 'ftp_backup.log')
        logging.basicConfig(
            filename=log_file, 
            level=logging.DEBUG, 
            format='%(asctime)s - %(levelname)s: %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        self.logger = logging.getLogger('FtpBackup')

    def debug(self, message):
        """Отладочное сообщение"""
        print(f"[FTP Backup Debug] {message}")
        self.logger.debug(message)

    def error(self, message):
        """Сообщение об ошибке"""
        print(f"[FTP Backup ERROR] {message}")
        self.logger.error(message)
        self.logger.error(traceback.format_exc())

class FtpBackupManager:
    def __init__(self, backup_root):
        self.backup_root = backup_root
        self.server_backup_map = {}
        self.config_path = os.path.join(backup_root, 'backup_config.json')
        
        self.logger = FtpBackupLogger(backup_root)
        self.project_roots = [
            'var\\www\\',
            'www\\',
            'public_html\\',
            'local\\',
            'htdocs\\',
            'home\\'
        ]
        
        os.makedirs(backup_root, exist_ok=True)
        self._load_config()
        
        self.logger.debug(f"Инициализация FtpBackupManager. Корневая папка: {backup_root}")

    def _load_config(self):
        """Загрузка конфигурации бэкапов с расширенной отладкой"""
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    self.server_backup_map = json.load(f)
                self.logger.debug(f"Конфигурация загружена: {len(self.server_backup_map)} записей")
            else:
                self.logger.debug("Файл конфигурации не найден. Будет создан новый.")
        except Exception as e:
            self.logger.error(f"Ошибка загрузки конфигурации: {e}")

    def _extract_relative_path(self, file_path):
        """
        Извлечение относительного пути файла с максимальной отладкой
        """
        self.logger.debug(f"Извлечение пути для: {file_path}")
        normalized_path = file_path.replace('/', '\\')
        for root in self.project_roots:
            if root in normalized_path:
                relative_path = normalized_path.split(root, 1)[1]
                result = relative_path.replace('\\', '/')
                self.logger.debug(f"Извлечен путь через корневую папку {root}: {result}")
                return result
        try:
            temp_match = re.search(r'Temp\\[^\\]+\\(.+)', normalized_path)
            if temp_match:
                result = temp_match.group(1).replace('\\', '/')
                self.logger.debug(f"Извлечен путь из временной папки: {result}")
                return result
        except Exception as e:
            self.logger.error(f"Ошибка извлечения из временной папки: {e}")
        
        result = os.path.basename(file_path)
        self.logger.debug(f"Использовано имя файла: {result}")
        return result

    def extract_site_name(self, file_path):
        """
        Извлечение имени сайта из пути к файлу
        """
        try:
            normalized_path = file_path.replace('/', '\\')
            patterns = [
                r'(?:var\\www\\|www\\|public_html\\|local\\|htdocs\\|home\\)([^\\]+)',
                r'ftp://([^/]+)',
                r'\\([^\\]+)\\(?:www|public_html|httpdocs)\\',
            ]
            
            for pattern in patterns:
                match = re.search(pattern, normalized_path)
                if match:
                    site_name = match.group(1)
                    self.logger.debug(f"Извлечено имя сайта: {site_name} из пути {file_path}")
                    return site_name
            parts = normalized_path.split('\\')
            for i, part in enumerate(parts):
                if part.lower() in ['www', 'public_html', 'httpdocs', 'htdocs'] and i > 0:
                    site_name = parts[i-1]
                    self.logger.debug(f"Использовано имя сайта из структуры пути: {site_name}")
                    return site_name
            for part in parts:
                if '.' in part and not part.endswith(('.php', '.html', '.js', '.css')):
                    self.logger.debug(f"Использовано имя с точкой как домен: {part}")
                    return part
        
            hostname = socket.gethostname()
            self.logger.debug(f"Не удалось извлечь имя сайта, используется имя хоста: {hostname}")
            return hostname
            
        except Exception as e:
            self.logger.error(f"Ошибка при извлечении имени сайта: {e}")
            return "unknown_site"

    def backup_file(self, file_path, server_name=None, mode=None, task_number=None):
        """
        Создание бэкапа с возможностью явного указания режима и номера задачи
        mode: None (автоматический), 'before', 'after'
        task_number: Номер задачи (опционально)
        """
        try:
            excluded_files = [
                'default.sublime-commands', 
                '.sublime-commands', 
                '.DS_Store',  
                'Thumbs.db'  
            ]
            
            if (os.path.basename(file_path) in excluded_files or 
                any(ext in file_path for ext in excluded_files)):
                self.logger.debug(f"Файл {file_path} исключен из бэкапа")
                return None, None

            if not os.path.exists(file_path):
                self.logger.error(f"Файл не существует: {file_path}")
                return None, None
            site_name = server_name or self.extract_site_name(file_path)
            server_key = re.sub(r'[^\w\-_.]', '_', site_name)
            
            self.logger.debug(f"Сайт: {site_name}, Ключ сайта: {server_key}")

            current_month_year = datetime.now().strftime("%B %Y")
            server_folder = os.path.join(self.backup_root, server_key)
            if not os.path.exists(server_folder):
                os.makedirs(server_folder, exist_ok=True)
                self.logger.debug(f"Создана новая папка для сайта: {server_key}")
            else:
                self.logger.debug(f"Используется существующая папка для сайта: {server_key}")
            
            month_year_folder = os.path.join(server_folder, current_month_year)
            if not os.path.exists(month_year_folder):
                os.makedirs(month_year_folder, exist_ok=True)
                self.logger.debug(f"Создана новая папка для месяца: {current_month_year}")
            else:
                self.logger.debug(f"Используется существующая папка для месяца: {current_month_year}")
            
            # Добавляем папку с номером задачи, если указан
            if task_number:
                task_folder = os.path.join(month_year_folder, f"task_{task_number}")
                self.logger.debug(f"Используется папка задачи: {task_number}")
            else:
                task_folder = month_year_folder
                self.logger.debug("Используется папка без номера задачи")
            
            if not os.path.exists(task_folder):
                os.makedirs(task_folder, exist_ok=True)
            
            before_path = os.path.join(task_folder, 'before')
            after_path = os.path.join(task_folder, 'after')

            os.makedirs(before_path, exist_ok=True)
            os.makedirs(after_path, exist_ok=True)
            
            relative_path = self._extract_relative_path(file_path)
            
            self.logger.debug(f"Относительный путь: {relative_path}")

            if mode == 'before':
                backup_path = os.path.join(before_path, relative_path)
                os.makedirs(os.path.dirname(backup_path), exist_ok=True)
                shutil.copy2(file_path, backup_path)
                self.logger.debug(f"Создан принудительный 'before' бэкап в {backup_path}")
                
                if relative_path not in self.server_backup_map:
                    self.server_backup_map[relative_path] = {
                        'first_backup_time': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        'site': site_name
                    }
            
            elif mode == 'after':
                backup_path = os.path.join(after_path, relative_path)
                os.makedirs(os.path.dirname(backup_path), exist_ok=True)
                
                if os.path.exists(backup_path):
                    os.remove(backup_path)
                
                shutil.copy2(file_path, backup_path)
                self.logger.debug(f"Создан принудительный 'after' бэкап в {backup_path}")
            
            else:
                first_backup_path = os.path.join(before_path, relative_path)
                after_backup_path = os.path.join(after_path, relative_path)
                
                os.makedirs(os.path.dirname(first_backup_path), exist_ok=True)
                os.makedirs(os.path.dirname(after_backup_path), exist_ok=True)

                if relative_path not in self.server_backup_map:
                    shutil.copy2(file_path, first_backup_path)
                    
                    self.server_backup_map[relative_path] = {
                        'first_backup_time': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        'site': site_name
                    }

                if os.path.exists(after_backup_path):
                    os.remove(after_backup_path)
                
                shutil.copy2(file_path, after_backup_path)

            if relative_path in self.server_backup_map:
                self.server_backup_map[relative_path]['last_backup_time'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                self.server_backup_map[relative_path]['site'] = site_name
            
            self._save_config()
            
            return before_path, after_path

        except Exception as e:
            self.logger.error(f"Критическая ошибка бэкапа: {e}")
            sublime.status_message(f"FTP Backup ERROR: {e}")
            return None, None

def create_backup_zip(self, folder_path, folder_type=None):
    """
    Создание ZIP-архива указанной папки бэкапа
    folder_path: путь к папке с бэкапами
    folder_type: 'before', 'after' или None (полная папка)
    """
    try:
        if not os.path.exists(folder_path):
            self.logger.error(f"Папка для архивации не существует: {folder_path}")
            return None
        
        # Приводим путь к стандартному виду Windows
        folder_path = os.path.normpath(folder_path)
        
        # Извлекаем имя сайта безопасным способом
        folder_parts = folder_path.split(os.sep)
        # Удаляем пустые элементы
        folder_parts = [part for part in folder_parts if part]
        
        # Находим имя сайта 
        site_name = "backup"  # значение по умолчанию
        for part in folder_parts:
            if re.match(r'[\w\-_.]+', part) and not part.startswith('task_') and part not in ['before', 'after']:
                site_name = part
                break
        
        # Формируем безопасное имя архива без недопустимых символов
        safe_site_name = re.sub(r'[\\/:*?"<>|]', '_', site_name)
        date_str = datetime.now().strftime("%d.%m.%Y.%H%M")
        
        if folder_type:
            zip_name = f"backup_{safe_site_name}_{folder_type}_{date_str}.zip"
        else:
            zip_name = f"backup_{safe_site_name}_{date_str}.zip"
        
        # Определяем папку для создания архива
        if any(part.startswith('task_') for part in folder_parts):
            # Если находимся в папке задачи
            for i, part in enumerate(folder_parts):
                if part.startswith('task_'):
                    task_folder_index = i
                    break
            
            # Формируем путь до папки задачи (включительно)
            task_path_parts = folder_parts[:task_folder_index+1]
            # Если Windows путь, добавляем диск
            if ':' in folder_path:
                zip_dir = os.path.join(folder_path.split(':')[0] + ':', os.sep, *task_path_parts)
            else:
                zip_dir = os.path.join(os.sep, *task_path_parts)
        else:
            # Если находимся в папке месяца или корневой папке
            zip_dir = os.path.dirname(folder_path)
            
            # Если текущая папка - before/after, поднимаемся на уровень выше
            if os.path.basename(folder_path) in ['before', 'after']:
                zip_dir = os.path.dirname(zip_dir)
        
        # Формируем полный путь к ZIP
        zip_path = os.path.join(zip_dir, zip_name)
        
        # Убеждаемся, что директория существует
        os.makedirs(os.path.dirname(zip_path), exist_ok=True)
        
        self.logger.debug(f"Создание архива: {zip_path}")
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(folder_path):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, folder_path)
                    zipf.write(file_path, arcname)
                    self.logger.debug(f"Добавлен файл {arcname}")
        
        self.logger.debug(f"Архив успешно создан: {zip_path}")
        
        # Сбрасываем текущий номер задачи после создания архива
        global CURRENT_TASK_NUMBER
        CURRENT_TASK_NUMBER = None
        self.logger.debug("Номер текущей задачи сброшен после создания архива")
        
        return zip_path
        
    except Exception as e:
        self.logger.error(f"Ошибка создания архива: {e}")
        return None
        
        
    def _save_config(self):
        """Сохранение конфигурации с расширенной отладкой"""
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self.server_backup_map, f, indent=4, ensure_ascii=False)
            self.logger.debug(f"Конфигурация сохранена: {len(self.server_backup_map)} записей")
        except Exception as e:
            self.logger.error(f"Ошибка сохранения конфигурации: {e}")

class SaveCommand(sublime_plugin.TextCommand):
    """Перекрытие стандартной команды save"""
    def run(self, edit, **kwargs):
        sublime.status_message("⛔ Используйте Ctrl+Shift+R для сохранения с бэкапом")
        
class SaveAsCommand(sublime_plugin.TextCommand):
    """Перекрытие стандартной команды save_as"""
    def run(self, edit, **kwargs):
        sublime.status_message("⛔ Используйте Ctrl+Shift+R для сохранения с бэкапом")

class PromptSaveAsCommand(sublime_plugin.TextCommand):
    """Перекрытие стандартной команды prompt_save_as"""
    def run(self, edit, **kwargs):
        sublime.status_message("⛔ Используйте Ctrl+Shift+R для сохранения с бэкапом")

class FtpBackupSaveCommand(sublime_plugin.TextCommand):
    """Команда для сохранения с бэкапом по Ctrl+Shift+R"""
    def run(self, edit):
        file_path = self.view.file_name()
        
        if not file_path:
            sublime.status_message("Сначала сохраните файл с указанием имени")
            self.view.window().run_command("save_as")
            return
        
        global CURRENT_TASK_NUMBER
        
        # Проверяем, есть ли уже номер задачи
        if CURRENT_TASK_NUMBER:
            # Если уже есть номер задачи, используем его
            self.save_with_backup(file_path, CURRENT_TASK_NUMBER)
        else:
            # Иначе запрашиваем новый номер задачи
            self.view.window().show_input_panel(
                "Введите номер задачи:", 
                "", 
                lambda task_number: self.save_with_backup(file_path, task_number), 
                None, 
                None
            )
    
    def save_with_backup(self, file_path, task_number):
        try:
            # Обработка пустого ввода
            task_number = task_number.strip() if task_number else None
            
            # Сохраняем номер задачи глобально
            global CURRENT_TASK_NUMBER
            CURRENT_TASK_NUMBER = task_number
            
            # ВАЖНО: Укажите ПОЛНЫЙ путь к директории бэкапов
            backup_root = r'C:\Users\aleksandr.kulakov\Desktop\BackUp'
            backup_manager = FtpBackupManager(backup_root)
        
            backup_manager.backup_file(file_path, mode='before', task_number=task_number)
            content = self.view.substr(sublime.Region(0, self.view.size()))
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(content)
            backup_manager.backup_file(file_path, mode='after', task_number=task_number)
            self.view.set_scratch(True)
            self.view.set_scratch(False)
            
            task_info = f" (задача #{task_number})" if task_number else ""
            sublime.status_message(f"✅ Файл успешно сохранен с бэкапом{task_info}: {os.path.basename(file_path)}")
        
        except Exception as e:
            sublime.error_message(f"❌ Ошибка сохранения с бэкапом: {str(e)}")

class BlockStandardSaveListener(sublime_plugin.EventListener):
    def on_text_command(self, view, command_name, args):
        """Перехват стандартных команд сохранения"""
        blocked_commands = ["save", "save_all", "prompt_save_as", "save_as", "save_all_with_new_window"]
        
        if command_name in blocked_commands:
            sublime.status_message("⛔ Используйте Ctrl+Shift+R для сохранения с бэкапом")
            return ("noop", None) 
        
        return None

    def on_pre_save(self, view):
        """Заблокировать любые прямые сохранения"""
        pass

    def on_post_save(self, view):
        """Заблокировать любые постобработки сохранения"""
        pass

    def on_query_context(self, view, key, operator, operand, match_all):
        """Перехватываем контекстные запросы для блокировки сохранения"""
        if key == "save_available":
            return False

class FtpBackupCreateBeforeCommand(sublime_plugin.TextCommand):
    def run(self, edit):
        """Создание принудительного 'before' бэкапа"""
        file_path = self.view.file_name()
        
        if not file_path:
            sublime.status_message("FTP Backup: Нет открытого файла")
            return
        
        global CURRENT_TASK_NUMBER
        
        # Проверяем, есть ли уже номер задачи
        if CURRENT_TASK_NUMBER:
            # Если уже есть номер задачи, используем его
            self.create_before_backup(file_path, CURRENT_TASK_NUMBER)
        else:
            # Иначе запрашиваем новый номер задачи
            self.view.window().show_input_panel(
                "Введите номер задачи:", 
                "", 
                lambda task_number: self.create_before_backup(file_path, task_number), 
                None, 
                None
            )
    
    def create_before_backup(self, file_path, task_number):
        try:
            # Обработка пустого ввода
            task_number = task_number.strip() if task_number else None
            
            # Сохраняем номер задачи глобально
            global CURRENT_TASK_NUMBER
            CURRENT_TASK_NUMBER = task_number
            
            # ВАЖНО: Укажите ПОЛНЫЙ путь к директории бэкапов
            backup_root = r'C:\Users\aleksandr.kulakov\Desktop\BackUp'
            backup_manager = FtpBackupManager(backup_root)
            
            before_path, _ = backup_manager.backup_file(file_path, mode='before', task_number=task_number)
            
            task_info = f" (задача #{task_number})" if task_number else ""
            sublime.status_message(f"FTP Backup: Создан 'before' бэкап{task_info} для {os.path.basename(file_path)}")
        
        except Exception as e:
            sublime.error_message(f"Ошибка создания 'before' бэкапа: {str(e)}")

class FtpBackupCreateAfterCommand(sublime_plugin.TextCommand):
    def run(self, edit):
        """Создание принудительного 'after' бэкапа"""
        file_path = self.view.file_name()
        
        if not file_path:
            sublime.status_message("FTP Backup: Нет открытого файла")
            return
        
        global CURRENT_TASK_NUMBER
        
        # Проверяем, есть ли уже номер задачи
        if CURRENT_TASK_NUMBER:
            # Если уже есть номер задачи, используем его
            self.create_after_backup(file_path, CURRENT_TASK_NUMBER)
        else:
            # Иначе запрашиваем новый номер задачи
            self.view.window().show_input_panel(
                "Введите номер задачи:", 
                "", 
                lambda task_number: self.create_after_backup(file_path, task_number), 
                None, 
                None
            )
    
    def create_after_backup(self, file_path, task_number):
        try:
            # Обработка пустого ввода
            task_number = task_number.strip() if task_number else None
            
            # Сохраняем номер задачи глобально
            global CURRENT_TASK_NUMBER
            CURRENT_TASK_NUMBER = task_number
            
            # ВАЖНО: Укажите ПОЛНЫЙ путь к директории бэкапов
            backup_root = r'C:\Users\aleksandr.kulakov\Desktop\BackUp'
            backup_manager = FtpBackupManager(backup_root)
            
            _, after_path = backup_manager.backup_file(file_path, mode='after', task_number=task_number)
            
            task_info = f" (задача #{task_number})" if task_number else ""
            sublime.status_message(f"FTP Backup: Создан 'after' бэкап{task_info} для {os.path.basename(file_path)}")
        
        except Exception as e:
            sublime.error_message(f"Ошибка создания 'after' бэкапа: {str(e)}")

class FtpBackupCreateZipCommand(sublime_plugin.WindowCommand):
    def run(self):
        """Создание ZIP-архива с выбором папки"""
        # ВАЖНО: Укажите ПОЛНЫЙ путь к директории бэкапов
        backup_root = r'C:\Users\aleksandr.kulakov\Desktop\BackUp'
        
        try:
            # Получаем список сайтов (папок первого уровня)
            sites = [d for d in os.listdir(backup_root) 
                    if os.path.isdir(os.path.join(backup_root, d)) and not d == 'logs']
            
            if not sites:
                sublime.status_message("FTP Backup: Нет доступных папок для архивации")
                return
            
            # Показываем выбор сайта
            self.backup_root = backup_root
            self.sites = sites
            self.window.show_quick_panel(sites, self.on_site_selected)
            
        except Exception as e:
            sublime.error_message(f"Ошибка при создании ZIP-архива: {str(e)}")
    
    def on_site_selected(self, index):
        if index == -1:
            return
        
        site = self.sites[index]
        site_path = os.path.join(self.backup_root, site)
        
        # Получаем список месяцев
        months = [d for d in os.listdir(site_path) 
                if os.path.isdir(os.path.join(site_path, d))]
        
        if not months:
            sublime.status_message(f"FTP Backup: В папке {site} нет доступных месяцев")
            return
        
        self.site = site
        self.months = months
        self.site_path = site_path
        self.window.show_quick_panel(months, self.on_month_selected)
    
    def on_month_selected(self, index):
        if index == -1:
            return
        
        month = self.months[index]
        month_path = os.path.join(self.site_path, month)
        
        # Проверяем наличие папок задач
        items = [d for d in os.listdir(month_path) 
                if os.path.isdir(os.path.join(month_path, d))]
        
        # Разделяем на задачи и папки before/after на корневом уровне
        tasks = [d for d in items if d.startswith('task_')]
        root_folders = [d for d in items if d in ['before', 'after']]
        
        all_options = []
        
        # Если есть корневые папки before/after, добавляем опцию всей папки месяца
        if root_folders:
            all_options.append(f"[Весь месяц] {month}")
            if 'before' in root_folders:
                all_options.append(f"[Before] {month}")
            if 'after' in root_folders:
                all_options.append(f"[After] {month}")
        
        # Добавляем папки задач
        for task in tasks:
            task_path = os.path.join(month_path, task)
            task_items = os.listdir(task_path)
            
            all_options.append(f"[Задача] {task}")
            
            if 'before' in task_items:
                all_options.append(f"[Before] {task}")
            if 'after' in task_items:
                all_options.append(f"[After] {task}")
        
        if not all_options:
            sublime.status_message(f"FTP Backup: В папке {month} нет доступных папок для архивации")
            return
        
        self.month = month
        self.month_path = month_path
        self.all_options = all_options
        self.tasks = tasks
        self.root_folders = root_folders
        
        self.window.show_quick_panel(all_options, self.on_folder_selected)
    
    def on_folder_selected(self, index):
        if index == -1:
            return
        
        selected = self.all_options[index]
        backup_manager = FtpBackupManager(self.backup_root)
        
        try:
            zip_path = None
            
            # Парсим выбранную опцию
            if selected.startswith('[Весь месяц]'):
                # Архивируем всю папку месяца
                zip_path = self.create_zip_archive(backup_manager, self.month_path)
                
            elif selected.startswith('[Before]') or selected.startswith('[After]'):
                folder_type = 'before' if selected.startswith('[Before]') else 'after'
                
                if selected.split('] ')[1] == self.month:
                    # Архивируем корневую папку before/after
                    folder_path = os.path.join(self.month_path, folder_type)
                else:
                    # Архивируем папку before/after задачи
                    task_name = selected.split('] ')[1]
                    folder_path = os.path.join(self.month_path, task_name, folder_type)
                
                zip_path = self.create_zip_archive(backup_manager, folder_path, folder_type)
                
            elif selected.startswith('[Задача]'):
                # Архивируем всю папку задачи
                task_name = selected.split('] ')[1]
                folder_path = os.path.join(self.month_path, task_name)
                zip_path = self.create_zip_archive(backup_manager, folder_path)
            
            # После создания архива, уведомляем пользователя
            if zip_path:
                sublime.status_message(f"FTP Backup: Архив успешно создан по пути: {zip_path}")
            else:
                sublime.status_message("FTP Backup: Ошибка при создании архива")
        
        except Exception as e:
            sublime.error_message(f"Ошибка при создании архива: {str(e)}")
    
    def create_zip_archive(self, backup_manager, folder_path, folder_type=None):
        """
        Создание ZIP-архива для указанной папки
        """
        try:
            if not os.path.exists(folder_path):
                sublime.status_message(f"FTP Backup: Папка для архивации не существует: {folder_path}")
                return None
            
            # Приводим путь к стандартному виду Windows
            folder_path = os.path.normpath(folder_path)
            
            # Извлекаем имя сайта безопасным способом
            folder_parts = folder_path.split(os.sep)
            # Удаляем пустые элементы
            folder_parts = [part for part in folder_parts if part]
            
            # Находим имя сайта 
            site_name = "backup"  # значение по умолчанию
            for part in folder_parts:
                if re.match(r'[\w\-_.]+', part) and not part.startswith('task_') and part not in ['before', 'after']:
                    site_name = part
                    break
            
            # Формируем безопасное имя архива без недопустимых символов
            safe_site_name = re.sub(r'[\\/:*?"<>|]', '_', site_name)
            date_str = datetime.now().strftime("%d.%m.%Y.%H%M")
            
            if folder_type:
                zip_name = f"backup_{safe_site_name}_{folder_type}_{date_str}.zip"
            else:
                zip_name = f"backup_{safe_site_name}_{date_str}.zip"
            
            # Определяем папку для создания архива
            if any(part.startswith('task_') for part in folder_parts):
                # Если находимся в папке задачи
                for i, part in enumerate(folder_parts):
                    if part.startswith('task_'):
                        task_folder_index = i
                        break
                
                # Формируем путь до папки задачи (включительно)
                task_path_parts = folder_parts[:task_folder_index+1]
                # Если Windows путь, добавляем диск
                if ':' in folder_path:
                    drive = folder_path.split(':')[0] + ':'
                    zip_dir = os.path.join(drive, os.sep, *task_path_parts)
                else:
                    zip_dir = os.path.join(os.sep, *task_path_parts)
            else:
                # Если находимся в папке месяца или корневой папке
                zip_dir = os.path.dirname(folder_path)
                
                # Если текущая папка - before/after, поднимаемся на уровень выше
                if os.path.basename(folder_path) in ['before', 'after']:
                    zip_dir = os.path.dirname(zip_dir)
            
            # Формируем полный путь к ZIP
            zip_path = os.path.join(zip_dir, zip_name)
            
            # Убеждаемся, что директория существует
            os.makedirs(os.path.dirname(zip_path), exist_ok=True)
            
            backup_manager.logger.debug(f"Создание архива: {zip_path}")
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root, dirs, files in os.walk(folder_path):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, folder_path)
                        zipf.write(file_path, arcname)
                        backup_manager.logger.debug(f"Добавлен файл {arcname}")
            
            backup_manager.logger.debug(f"Архив успешно создан: {zip_path}")
            
            # Сбрасываем текущий номер задачи после создания архива
            global CURRENT_TASK_NUMBER
            CURRENT_TASK_NUMBER = None
            backup_manager.logger.debug("Номер текущей задачи сброшен после создания архива")
            
            return zip_path
            
        except Exception as e:
            backup_manager.logger.error(f"Ошибка создания архива: {e}")
            return None