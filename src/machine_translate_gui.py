import tkinter as tk
from tkinter import filedialog, messagebox, font
import subprocess
import os
import shutil
import platform
import webbrowser
import argparse
import traceback
import sys
if platform.system() == 'Windows':
    from win32com import client
    from comtypes.client import CreateObject
    from win32com import client
    import comtypes
import re
import pkgutil

class MachineTranslationApp:
    def __init__(self, root, docxfile_path="", bin_path="."):
        self.root = root
        self.root.title("Word Docx Document Translator")
            
        try:
            self.parent_folder_path = f"{bin_path}"
            self.parent_folder_path = self.get_parent_folder_path()
            self.icon_path = os.path.join (os.path.join(self.parent_folder_path, 'img'), 'app.ico')
            
            self.root.iconbitmap(self.icon_path )
        except:
            var = traceback.format_exc()
            print(var)
        
        self.google_translate_lang_codes = {
            'Afrikaans': 'af',
            'Albanian': 'sq',
            'Amharic': 'am',
            'Arabic': 'ar',
            'Armenian': 'hy',
            'Azerbaijani': 'az',
            'Basque': 'eu',
            'Belarusian': 'be',
            'Bengali': 'bn',
            'Bosnian': 'bs',
            'Bulgarian': 'bg',
            'Catalan': 'ca',
            'Cebuano': 'ceb',
            'Chinese (Simplified)': 'zh-CN',
            'Chinese (Traditional)': 'zh-TW',
            'Corsican': 'co',
            'Croatian': 'hr',
            'Czech': 'cs',
            'Danish': 'da',
            'Dutch': 'nl',
            'English': 'en',
            'Esperanto': 'eo',
            'Estonian': 'et',
            'Finnish': 'fi',
            'French': 'fr',
            'Frisian': 'fy',
            'Galician': 'gl',
            'Georgian': 'ka',
            'German': 'de',
            'Greek': 'el',
            'Gujarati': 'gu',
            'Haitian Creole': 'ht',
            'Hausa': 'ha',
            'Hawaiian': 'haw',
            'Hebrew': 'iw',
            'Hindi': 'hi',
            'Hmong': 'hmn',
            'Hungarian': 'hu',
            'Icelandic': 'is',
            'Igbo': 'ig',
            'Indonesian': 'id',
            'Irish': 'ga',
            'Italian': 'it',
            'Japanese': 'ja',
            'Javanese': 'jv',
            'Kannada': 'kn',
            'Kazakh': 'kk',
            'Khmer': 'km',
            'Korean': 'ko',
            'Kurdish': 'ku',
            'Kyrgyz': 'ky',
            'Lao': 'lo',
            'Latin': 'la',
            'Latvian': 'lv',
            'Lithuanian': 'lt',
            'Luxembourgish': 'lb',
            'Macedonian': 'mk',
            'Malagasy': 'mg',
            'Malay': 'ms',
            'Malayalam': 'ml',
            'Maltese': 'mt',
            'Maori': 'mi',
            'Marathi': 'mr',
            'Mongolian': 'mn',
            'Myanmar (Burmese)': 'my',
            'Nepali': 'ne',
            'Norwegian': 'no',
            'Nyanja (Chichewa)': 'ny',
            'Pashto': 'ps',
            'Persian': 'fa',
            'Polish': 'pl',
            'Portuguese': 'pt',
            'Punjabi': 'pa',
            'Romanian': 'ro',
            'Russian': 'ru',
            'Samoan': 'sm',
            'Scots Gaelic': 'gd',
            'Serbian': 'sr',
            'Sesotho': 'st',
            'Shona': 'sn',
            'Sindhi': 'sd',
            'Sinhala (Sinhalese)': 'si',
            'Slovak': 'sk',
            'Slovenian': 'sl',
            'Somali': 'so',
            'Spanish': 'es',
            'Sundanese': 'su',
            'Swahili': 'sw',
            'Swedish': 'sv',
            'Tagalog (Filipino)': 'tl',
            'Tajik': 'tg',
            'Tamil': 'ta',
            'Telugu': 'te',
            'Thai': 'th',
            'Turkish': 'tr',
            'Ukrainian': 'uk',
            'Urdu': 'ur',
            'Uzbek': 'uz',
            'Vietnamese': 'vi',
            'Welsh': 'cy',
            'Xhosa': 'xh',
            'Yiddish': 'yi',
            'Yoruba': 'yo',
            'zu': 'Zulu'
            }
        
        self.deepl_translate_lang_codes = {
            'Acehnese': 'ace',
            'Afrikaans': 'af',
            'Albanian': 'sq',
            'Arabic': 'ar',
            'Aragonese': 'an',
            'Armenian': 'hy',
            'Assamese': 'as',
            'Aymara': 'ay',
            'Azerbaijani': 'az',
            'Bashkir': 'ba',
            'Basque': 'eu',
            'Belarusian': 'be',
            'Bengali': 'bn',
            'Bhojpuri': 'bho',
            'Bosnian': 'bs',
            'Breton': 'br',
            'Bulgarian': 'bg',
            'Burmese': 'my',
            'Cantonese': 'yue',
            'Catalan': 'ca',
            'Cebuano': 'ceb',
            'Chinese (simplified)': 'zh-Hans',
            'Chinese (traditional)': 'zh-Hant',
            'Croatian': 'hr',
            'Czech': 'cs',
            'Danish': 'da',
            'Dari': 'prs',
            'Dutch': 'nl',
            'English (American)': 'en-US',
            'English (British)': 'en-GB',
            'Esperanto': 'eo',
            'Estonian': 'et',
            'Finnish': 'fi',
            'French': 'fr',
            'Galician': 'gl',
            'Georgian': 'ka',
            'German': 'de',
            'Greek': 'el',
            'Guarani': 'gn',
            'Gujarati': 'gu',
            'Haitian Creole': 'ht',
            'Hausa': 'ha',
            'Hebrew': 'he',
            'Hindi': 'hi',
            'Hungarian': 'hu',
            'Icelandic': 'is',
            'Igbo': 'ig',
            'Indonesian': 'id',
            'Irish': 'ga',
            'Italian': 'it',
            'Japanese': 'ja',
            'Javanese': 'jv',
            'Kapampangan': 'pam',
            'Kazakh': 'kk',
            'Konkani': 'gom',
            'Korean': 'ko',
            'Kurdish (Kurmanji)': 'kmr',
            'Kurdish (Sorani)': 'ckb',
            'Kyrgyz': 'ky',
            'Latin': 'la',
            'Latvian': 'lv',
            'Lingala': 'ln',
            'Lithuanian': 'lt',
            'Lombard': 'lmo',
            'Luxembourgish': 'lb',
            'Macedonian': 'mk',
            'Maithili': 'mai',
            'Malagasy': 'mg',
            'Malay': 'ms',
            'Malayalam': 'ml',
            'Maltese': 'mt',
            'Maori': 'mi',
            'Marathi': 'mr',
            'Mongolian': 'mn',
            'Nepali': 'ne',
            'Norwegian (bokmål)': 'nb',
            'Occitan': 'oc',
            'Oromo': 'om',
            'Pangasinan': 'pag',
            'Pashto': 'ps',
            'Persian': 'fa',
            'Polish': 'pl',
            'Portuguese': 'pt-PT',
            'Portuguese (Brazilian)': 'pt-BR',
            'Punjabi': 'pa',
            'Quechua': 'qu',
            'Romanian': 'ro',
            'Russian': 'ru',
            'Sanskrit': 'sa',
            'Serbian': 'sr',
            'Sesotho': 'st',
            'Sicilian': 'scn',
            'Slovak': 'sk',
            'Slovenian': 'sl',
            'Spanish': 'es',
            'Spanish (Latin American)': 'es-419',
            'Sundanese': 'su',
            'Swahili': 'sw',
            'Swedish': 'sv',
            'Tagalog': 'tl',
            'Tajik': 'tg',
            'Tamil': 'ta',
            'Tatar': 'tt',
            'Telugu': 'te',
            'Tsonga': 'ts',
            'Tswana': 'tn',
            'Turkish': 'tr',
            'Turkmen': 'tk',
            'Ukrainian': 'uk',
            'Urdu': 'ur',
            'Uzbek': 'uz',
            'Vietnamese': 'vi',
            'Welsh': 'cy',
            'Wolof': 'wo',
            'Xhosa': 'xh',
            'Yiddish': 'yi',
            'Zulu': 'zu'
            }
            
        self.deepl_languages = self.deepl_translate_lang_codes.keys()
        self.google_languages = self.google_translate_lang_codes.keys()
        
        # Load stored target language value
        
        self.conf_folder_path = self.get_app_subfolder_path('conf')

        # Construct the full path to target_language.txt within the 'conf' folder
        
        self.target_language_file = os.path.join(self.conf_folder_path, 'target_language.txt')
        
        
        self.menu_bar = tk.Menu(root)
        root.config(menu=self.menu_bar)

        # Create Help menu
        self.help_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Help", menu=self.help_menu)

        # Add About submenu to Help menu
        self.help_menu.add_command(label="About", command=self.show_about_dialog)


        # File selection
        self.docx_file_label = tk.Label(root, text="Word docx file : ")
        self.docx_file_label.grid(row=1, column=0)

        self.docx_file_entry = tk.Entry(root)
        self.docx_file_entry.grid(row=1, column=1)

        self.docx_browse_button = tk.Button(root, text="Browse", command=self.browse_docx_file)
        self.docx_browse_button.grid(row=1, column=2)
        

        # Excel file search and replace file
        self.xlsx_file_label = tk.Label(root, text="Xlsx search are replace file (optional) : ")
        self.xlsx_file_label.grid(row=2, column=0)

        self.xlsx_file_entry = tk.Entry(root)
        self.xlsx_file_entry.grid(row=2, column=1)

        self.xlsx_browse_button = tk.Button(root, text="Browse", command=self.browse_xlsx_file)
        self.xlsx_browse_button.grid(row=2, column=2)
        
        # Source language selection
        self.source_label = tk.Label(root, text="Source language")
        self.source_label.grid(row=3, column=0)

        self.languages = [
            "English", "English (British)", "English (American)", # Added "English" to the language list
            "Arabic", "Bulgarian", "Chinese", "Chinese (Simplified)", "Chinese (Traditional)", "Czech", "French",
            "German", "Hindi", "Hungarian", "Indonesian", "Italian", "Japanese", "Korean",
            "Malay", "Mongolian", "Nepali", "Persian", "Polish", "Portuguese", "Portuguese (Brazilian)",
            "Punjabi", "Romanian", "Russian", "Spanish", "Telugu", "Thai", "Ukrainian",
            "Urdu", "Vietnamese"
        ]
        
        self.exclude_from_source_languages = [
            "English (British)", "English (American)", "Portuguese (Brazilian)"
        ]
        
        self.languages_filtered = [language for language in self.languages if language not in self.exclude_from_source_languages]

        
        self.exclude_from_deepl_target_languages = [
            "English"
        ]
        
        self.source_language = tk.StringVar(root)
        self.source_language.set(self.languages[0])
        self.source_combo = tk.OptionMenu(root, self.source_language, *sorted(self.languages_filtered))
        self.source_combo.grid(row=3, column=1)

        # Target language selection
        self.target_label = tk.Label(root, text="Target language")
        self.target_label.grid(row=4, column=0)

        self.target_language = tk.StringVar(root)
        self.target_language.set(self.languages[0])
        self.target_combo = tk.OptionMenu(root, self.target_language, *sorted(self.languages))
        self.target_combo.grid(row=4, column=1)
        
        # More languages checkbox
        self.show_all_languages_var = tk.BooleanVar(value=False)  # Set default value to checked
        self.show_all_languages_checkbox = tk.Checkbutton(root, text="Display all languages", variable=self.show_all_languages_var)
        self.show_all_languages_checkbox.grid(row=5, column=1, columnspan=2)
        self.show_all_languages_checkbox.bind("<ButtonRelease-1>", self.toggle_show_all_languages)

        # Engine selection
        self.engine_label = tk.Label(root, text="Engine")
        self.engine_label.grid(row=6, column=0)

        engines = ["Deepl", "Google","Perplexity", "Chatgpt"]
        self.engine = tk.StringVar(root)
        self.engine.set(engines[0])
        self.engine_combo = tk.OptionMenu(root, self.engine, *engines)
        self.engine_combo.grid(row=6, column=1)
        
        # Target font selection
        self.font_label = tk.Label(root, text="Select a target font (optional) :")
        self.font_label.grid(row=7, column=0, sticky=tk.W)

        self.font_var = tk.StringVar()
        self.font_combo = tk.OptionMenu(root, self.font_var, "")
        self.populate_fonts()
        self.font_combo.grid(row=7, column=1)
        
        # Split translation checkbox
        self.split_var = tk.BooleanVar(value=True)  # Set default value to checked
        self.split_checkbox = tk.Checkbutton(root, text="Split translation", variable=self.split_var)
        self.split_checkbox.grid(row=8, column=0, columnspan=2, sticky=tk.W)
        self.split_checkbox.bind("<ButtonRelease-1>", self.toggle_split_message)

        # Show browser checkbox
        self.show_browser_var = tk.BooleanVar(value=True)
        self.show_browser_checkbox = tk.Checkbutton(root, text="Show browser", variable=self.show_browser_var, state="disabled")
        self.show_browser_checkbox.grid(row=8, column=1, sticky=tk.W)

        # Open document after translation
        self.open_file_after_translation_var = tk.BooleanVar(value=False)
        self.open_file_after_translation_checkbox = tk.Checkbutton(root, text="Open file after translation", variable=self.open_file_after_translation_var, state="normal")
        self.open_file_after_translation_checkbox.grid(row=9, column=0, sticky=tk.W)

        # Generate shortcut
        generate_shortcut_button_text = ""
        if platform.system() == 'Windows':
            generate_shortcut_button_text = "Create shortcut"
            self.generate_shortcut_button = tk.Button(root, text=generate_shortcut_button_text, command=self.create_windows_shortcut_wrapper)
        elif platform.system() == 'Darwin':
            generate_shortcut_button_text = "Create Mac finder service"
            self.generate_shortcut_button = tk.Button(root, text=generate_shortcut_button_text, command=self.create_mac_shortcut_wrapper)
        
        self.generate_shortcut_button.grid(row=10, column=0)

        # Translate button
        self.translate_button = tk.Button(root, text="Translate", command=self.translate)
        self.translate_button.grid(row=10, column=1)
        
        # Create a status bar
        self.status_bar = tk.Label(root, text="Status bar", bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.grid(row=11, columnspan=3, sticky=tk.W+tk.E)
        
        # Configure rows and columns to resize with window
        # self.root.rowconfigure(0, weight=1)
        # self.root.rowconfigure(1, weight=1)
        # self.root.rowconfigure(2, weight=1)
        # self.root.rowconfigure(3, weight=1)
        # self.root.rowconfigure(4, weight=1)
        # self.root.rowconfigure(5, weight=1)
        # self.root.rowconfigure(6, weight=1)
        # self.root.rowconfigure(7, weight=1)
        # self.root.rowconfigure(8, weight=1)
        # self.root.rowconfigure(9, weight=1)
        # self.root.rowconfigure(10, weight=1)
        # self.root.columnconfigure(0, weight=1)
        # self.root.columnconfigure(1, weight=1)
        # self.root.columnconfigure(2, weight=1)
        
        # Automatically select Deepl for certain languages
        self.source_language.trace("w", self.auto_select_engine)
        
        # Automatically select Deepl for certain languages
        self.target_language.trace("w", self.auto_select_engine)
        
        # Automatically select target font
        self.target_language.trace("w", self.auto_select_font)
        
        # Connect target language selection to set xlsx_file_entry
        self.target_language.trace("w", self.auto_set_xlsx_entry)        
        
        # Connect engine selection to show_browser_var
        self.engine.trace("w", self.update_show_browser)
        
        # Connect show_browser_var change to message
        self.show_browser_var.trace("w", self.show_browser_message)
        
        self.load_target_language()
        
        if docxfile_path is not None:
            self.docx_file_entry.delete(0, tk.END)
            self.docx_file_entry.insert(0, docxfile_path)
            
        if bin_path is not None:
            self.bin_path = bin_path
        else:
            self.bin_path = '.'

        # Add tooltip for some widgets
        self.create_tooltip(self.generate_shortcut_button, "Generate a SendTo shortcut for Windows Explorer")
        self.create_tooltip(self.split_checkbox, "Split the translation into multiple cells or single cell.")
        
    def get_parent_folder_path(self):
        if getattr(sys, 'frozen', False):
            # If the program is frozen, get the directory containing the executable
            program_dir = os.path.dirname(sys.executable)
        else:
            # If the program is a script, get the current script's directory
            current_dir = os.path.dirname(os.path.abspath(__file__))
            program_dir = os.path.dirname(current_dir)
        
        return program_dir
        
    def get_app_subfolder_path(self, foldername):
        # Check if the program is running as a PyInstaller executable
        if getattr(sys, 'frozen', False):
            # Get the executable's directory
            executable_dir = os.path.dirname(sys.executable)

            # Create the 'conf' folder path within the executable's directory
            app_subfolder_path = os.path.join(executable_dir, foldername)
        else:
            # Get the current directory of the script
            current_dir = os.path.dirname(os.path.abspath(__file__))

            # Go up one level to the parent directory
            parent_dir = os.path.dirname(current_dir)

            # Create the 'conf' folder path
            app_subfolder_path = os.path.join(parent_dir, foldername)

        # Ensure the 'conf' folder exists, create it if not
        try:
            if not os.path.exists(app_subfolder_path):
                os.makedirs(app_subfolder_path)
        except Exception as e:
            if getattr(sys, 'frozen', False):
                # If creation fails and the program is frozen, use the executable's directory
                app_subfolder_path = os.path.dirname(sys.executable)
            else:
                # If creation fails and the program is not frozen, use the program's directory
                app_subfolder_path = os.path.dirname(os.path.abspath(__file__))

        return app_subfolder_path
    
    def show_about_dialog(self):
        about_message = (
            "Would like to visit the home page :\n"
            "Visit https://github.com/translation-robot/machine-translate-docx ?"
        )
        result = messagebox.askokcancel("About", about_message)
        if result:
            webbrowser.open("https://github.com/translation-robot/machine-translate-docx")

    def load_target_language(self):
        try:
            with open(self.target_language_file, "r") as file:
                stored_language = file.read().strip()
                self.target_language.set(stored_language)
        except FileNotFoundError:
            pass
            
    def auto_set_xlsx_entry(self, *args):
        target_language = self.target_language.get().lower()
        
        if getattr(sys, 'frozen', False):
            # If the application is compiled by PyInstaller
            application_path = os.path.dirname(sys.executable)
        else:
            # If running in a normal Python environment
            application_path = os.path.dirname(os.path.abspath(__file__))
        
        if platform.system() == 'Windows':
            xlsx_filename = os.path.abspath(f"{application_path}\\..\\{target_language}.xlsx")
        else:
            home_folder = os.environ['HOME']
            xlsx_filename = os.path.abspath(f"{application_path}/../xlsx/{target_language}.xlsx")

        self.xlsx_file_entry.delete(0, tk.END)
        
        if os.path.exists(xlsx_filename):
            self.xlsx_file_entry.insert(0, xlsx_filename)
    
    def populate_fonts(self):
        system_fonts = font.families()
        # Add emtpy entry for no font
        self.font_combo['menu'].delete(0, 'end')
        self.font_combo['menu'].add_command(label="", command=tk._setit(self.font_var, ""))
        for font_name in system_fonts:
            self.font_combo['menu'].add_command(label=font_name, command=tk._setit(self.font_var, font_name))

    def save_target_language(self):
        with open(self.target_language_file, "w") as file:
            file.write(self.target_language.get())
            
    def browse_docx_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
        if file_path and platform.system() == 'Windows':
            # Replace forward slashes with backslashes
            file_path = file_path.replace("/", "\\")
        else:
            file_path = file_path.replace("\\", "/")
        
        self.docx_file_entry.delete(0, tk.END)
        self.docx_file_entry.insert(0, file_path)
        self.xlsx_file_entry.icursor(len(self.docx_file_entry.get()))
        self.xlsx_file_entry.index(len(self.docx_file_entry.get()))
        self.docx_file_entry.focus()
            
    def browse_xlsx_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel xlsx Document", "*.xlsx")])
        if file_path and platform.system() == 'Windows':
            # Replace forward slashes with backslashes
            file_path = file_path.replace("/", "\\")
        else:
            file_path = file_path.replace("\\", "/")

        self.xlsx_file_entry.delete(0, tk.END)
        self.xlsx_file_entry.insert(0, file_path)
        self.xlsx_file_entry.icursor(len(self.xlsx_file_entry.get()))
        self.root.update()
        
    def toggle_split_message(self, event):
        if self.split_var.get():
            message = (
                "Each phrase translation will be on a single cell in the table, "
                "and you will need to run The split translation program afterward "
                "or manually divide the translation in the cells."
            )
            messagebox.showinfo("Split Translation Notice", message)
            self.split_var.set(False)
            
    def toggle_show_all_languages(self, event):
        if self.show_all_languages_var.get():
            #messagebox.showinfo("Show languages", "Showing all languages")
            
            src_lang = self.source_language.get()
            dest_lang = self.target_language.get()
            self.source_combo = tk.OptionMenu(root, self.source_language, *sorted(self.languages))
            self.target_combo = tk.OptionMenu(root, self.target_language, *sorted(self.languages))
            self.target_language.set(dest_lang)
            self.source_language.set(src_lang)
            
            self.source_combo.grid(row=3, column=1)
            self.target_combo.grid(row=4, column=1)
            #self.show_all_languages_var.set(False)
        else:
            #messagebox.showinfo("Show languages", "Not showing all languages")
            #self.show_all_languages_var.set(True)
            
            src_lang = self.source_language.get()
            dest_lang = self.target_language.get()
            self.source_combo = tk.OptionMenu(root, self.source_language, *sorted(self.google_translate_lang_codes.keys()))
            self.target_combo = tk.OptionMenu(root, self.target_language, *sorted(self.google_translate_lang_codes.keys()))
            self.target_language.set(dest_lang)
            self.source_language.set(src_lang)
            
            self.source_combo.grid(row=3, column=1)
            self.target_combo.grid(row=4, column=1)
            #self.show_all_languages_var.set(True)
            
    def auto_select_engine(self, *args):
        # Always-available engines
        engines = ["Chatgpt", "Google", "Perplexity"]
        
        # Add Deepl only if both source & target are supported
        if (self.source_language.get() in self.deepl_languages and
            self.target_language.get() in self.deepl_languages):
            engines.append("Deepl")
            self.engine.set("Deepl")
            self.show_browser_var.set(True)
            self.show_browser_checkbox.config(state="disabled")
        else:
            # Force Google if Deepl is not available
            self.engine.set("Google")
            self.show_browser_var.set(False)
            self.show_browser_checkbox.config(state="normal")
            if self.source_language.get() not in self.deepl_languages:
                self.target_language.set(self.languages[0])
        
        # ✅ Sort the engines alphabetically (case-insensitive)
        engines = sorted(engines, key=str.lower)
        
        # Update OptionMenu with the new engine list
        menu = self.engine_combo["menu"]
        menu.delete(0, "end")
        for engine in engines:
            menu.add_command(
                label=engine,
                command=lambda value=engine: self.engine.set(value)
            )

    def auto_select_font(self, *args):
        # List of font for the selected language

        if self.target_language.get().lower() in ['hindi','punjabi']:
            self.font_var.set("Mangal")
        else:
            #self.font_var.set("Times New Roman")
            self.font_var.set("")
            
    def update_show_browser(self, *args):
        if self.engine.get() == "Deepl" or self.engine.get() == "Perplexity":
            self.show_browser_var.set(True)
            self.show_browser_checkbox.config(state="disabled")
        elif self.engine.get() == "Google":
            self.show_browser_var.set(False)
            self.show_browser_checkbox.config(state="normal")
    
    def show_browser_message(self, *args):
        if self.show_browser_var.get() and self.engine.get() == "Google":
            message = ("When translating using Google engine, do not interact with the browser " 
                "when it is visible or the translation will likely fail.")
            result = messagebox.askokcancel("Browser Interaction Notice", message)
            if not result:
                self.show_browser_var.set(False)

    # Source https://www.programcreek.com/python/?CodeExample=create+shortcut
    def create_windows_shortcut(self, lnk_out_path, target, parameters, working_dir, description, icon=None, run_as_admin=False, minimized=False):
        shell = client.Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(lnk_out_path)
        shortcut.Targetpath = target
        shortcut.Arguments = '{}'.format(parameters)
        shortcut.Description = description
        shortcut.WorkingDirectory = working_dir
        if not icon is None:
            shortcut.IconLocation = icon
        if minimized:# 7 - Minimized, 3 - Maximized, 1 - Normal
            shortcut.WindowStyle = 7
        else:
            shortcut.WindowStyle = 1
        shortcut.save()
        
        if run_as_admin:
            with open(lnk_out_path, "r+b") as f:
                with contextlib.closing(mmap.mmap(f.fileno(), 0, access=mmap.ACCESS_WRITE)) as m:
                    m[0x15] = m[0x15] | 0x20 # Enable 6th bit = Responsible for Run As Admin
                    #m.flush()
                    
    def create_windows_shortcut_wrapper(self):
        docx_file_path = self.docx_file_entry.get()
        if platform.system() == 'Windows':
            docx_file_path = docx_file_path.replace('/', '\\')
        
        engine = self.engine.get().lower()
        src_lang_name = self.source_language.get()
        dest_lang_name = self.target_language.get()
        
        xlsx_file_path = self.xlsx_file_entry.get()
        if platform.system() == 'Windows':
            xlsx_file_path = xlsx_file_path.replace('/', '\\')
        
        font_value = self.font_var.get()
        
        if self.split_var.get():
            split_param = " --split "
        else:
            split_param = " "
        
        if self.show_browser_var.get():
            show_browser_param = " --showbrowser "
        else:
            show_browser_param = " "
        
        src_lang_code = 'en'
        if engine == 'deepl':
            src_lang_code = self.deepl_translate_lang_codes[src_lang_name]
            dest_lang_code = self.deepl_translate_lang_codes[dest_lang_name]
            # Show brower is ignored when using Deepl and always true
            show_browser_param = ""
        else:
            src_lang_code = self.google_translate_lang_codes[src_lang_name]
            dest_lang_code = self.google_translate_lang_codes[dest_lang_name]
            
        xlsx_replace_param = " "
        if xlsx_file_path is not None and xlsx_file_path != "":
            if platform.system() == 'Windows':
                xlsx_replace_param = f" --xlsxreplacefile \"{xlsx_file_path}\""
            else:
                xlsx_replace_param = f" --xlsxreplacefile \\\"{xlsx_file_path}\\\""
       
        dest_font_param = " "
        if font_value is not None and font_value != "":
            if platform.system() == 'Windows':
                dest_font_param = f" --destfont \"{font_value}\""
            else:
                dest_font_param = f" --destfont \\\"{font_value}\\\""
        
        exitonsuccess_param = ""
        if self.open_file_after_translation_var.get():
            exitonsuccess_param = "  --exitonsuccess "
            
        if self.target_language.get() == self.source_language.get():
            message = (
                "Please select a target translation language different than the source language."
            )
            messagebox.showinfo("Select a target translation language", message)
            return
            
        bin_launcher_path = ""
        target = f"{self.bin_path}\\..\\WindowsTerminal\\WindowsTerminal.exe"
        
        open_word_param = ""
        viewdocx_label = ""
        if self.open_file_after_translation_var.get():
            open_word_param = "-l"
            viewdocx_label = " - open translated file"
        
        try:
            target = os.path.abspath(target)
        except:
            print("Error, could not find target path")
        console_arguments = ""
        arguments = ""
        if platform.system() == 'Windows':
            console_arguments = f" -- "
        else:
            bin_launcher_path = f"osascript -e 'tell app \"Terminal\" to do script \"{self.bin_path}/machine-translate-docx "
            
        if platform.system() == 'Windows':
            arguments = f"{console_arguments} {self.bin_path}\\machine-translate-docx.exe --srclang {src_lang_code} --destlang {dest_lang_code} --engine {engine} {dest_font_param} {split_param} {xlsx_replace_param} {show_browser_param} {open_word_param} {exitonsuccess_param} --docxfile "
        else:
            command = f"{bin_launcher_path} {self.bin_path}\\machine-translate-docx.exe --srclang {src_lang_code} --destlang {dest_lang_code} --engine {engine} {dest_font_param} {split_param} {xlsx_replace_param} {show_browser_param} {open_word_param} {exitonsuccess_param} --docxfile "
        
        # Replace multiple spaces into a single space
        arguments = re.sub('\s+',' ',arguments)
        #print("arguments:")
        #print(arguments)
        
        shortcut_icon_path = os.path.join(bin_path, "app.ico")
        if not os.path.exists(shortcut_icon_path):
            shortcut_icon_path = None
        #print("target:")
        #print(target)
        
        engine = self.engine.get()
        split_string=""
        if not self.split_var.get():
            split_string = " - no split"
        
        # Get Sendto Folder location
        appdata_folder = os.environ['APPDATA']
        shortcut_path = os.path.join(appdata_folder,
            "Microsoft\\Windows\\SendTo",
            f'machine-translate-docx.exe - {dest_lang_name} - {engine}{split_string}{viewdocx_label}.lnk')
            
        if os.path.exists(shortcut_path):
            confirm_overwrite_message = (
                f"Sendto shortcut:\n\n\"machine-translate-docx.exe - {dest_lang_name} - {engine}{split_string}{viewdocx_label}\"\n\nalready exist."
                "\n\nOvewrite this shortcut ?"
            )
            overwrite_shortcut = messagebox.askokcancel("Ovewrite shortcut ?", confirm_overwrite_message)
            if not overwrite_shortcut:
                return
        
        self.create_windows_shortcut(shortcut_path, target, arguments, bin_path, 'Run console shortcut', icon=shortcut_icon_path)
        
        if os.path.exists(shortcut_path):
            message = f"Sendto shortcut created"
            messagebox.showinfo(message, f"Sendto shortcut is now created:\n\n\"machine-translate-docx.exe - {dest_lang_name} - {engine}{split_string}{viewdocx_label}\".\n\nThis needs to be done only once unless you want to change the settings for that language.")
        else:
            messagebox.showerror("Error, shortcut was not created", f"Error : Sendto shortcut was not created :\n\n\"machine-translate-docx.exe - {dest_lang_name} - {engine}{split_string}{viewdocx_label}.lnk\".\n\nYou may report this error to smtv.bot@gmail.com")
            
    def resource_path(self, relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)
    
    def create_mac_shortcut_wrapper(self):
        info_plist_file_path = self.resource_path("mac_service_template/machine-translate-docx_template.workflow/Contents/Info.plist")
        document_wflow_file_path = self.resource_path("mac_service_template/machine-translate-docx_template.workflow/Contents/document.wflow")
        
        engine = self.engine.get()
        src_lang_name = self.source_language.get()
        dest_lang_name = self.target_language.get()
        
        split_string=""
        if not self.split_var.get():
            split_string = " - no split"

        viewdocx_label = ""
        if self.open_file_after_translation_var.get():
            viewdocx_label = " - open translated file"
            
        dest_lang_name = self.target_language.get()
        service_name = f"machine-translate-docx - {dest_lang_name} - {engine}{split_string}{viewdocx_label}"
        
        #open text file in read mode
        info_plist_file = open(info_plist_file_path, "r")
        document_wflow_file = open(document_wflow_file_path, "r")
        
        command = self.make_translate_command(for_service_template=True)
        print("command:\n{command}")
         
        #read whole file to a string
        info_plist_content = info_plist_file.read().replace('{service_name}', service_name)
        
        if engine == 'deepl':
            src_lang_code = self.deepl_translate_lang_codes[src_lang_name]
            dest_lang_code = self.deepl_translate_lang_codes[dest_lang_name]
            # Show brower is ignored when using Deepl and always true
            show_browser_param = ""
        else:
            src_lang_code = self.google_translate_lang_codes[src_lang_name]
            dest_lang_code = self.google_translate_lang_codes[dest_lang_name]
            
        #command = f"/Users/sysprobs/SMTVRobot/bin/machine-translate-docx  --srclang en --destlang pl --engine deepl    --split   --xlsxreplacefile \\\"/Users/sysprobs/SMTVRobot/polish.xlsx\\\"  --showbrowser    --exitonsuccess  --docxfile \" &amp; thisItemsPathname &amp; \"; ; exit 0;"
        document_wflow_content = document_wflow_file.read().replace('{run_command}', command)
         
        #close file
        info_plist_file.close()
         
        #print(info_plist_content)
        #print("")
        #print(document_wflow_content)

        path = 'Info.plist'  # always use slash
        message = f"Creating mac service"
        #messagebox.showinfo(message, info_plist_content)
        #messagebox.showinfo(message, document_wflow_content)
        
        # Create service folder
        # Write to TMPDIR or /tmp folder
        tmpdir = ""
        try:
            tmpdir = os.environ['TMPDIR']
            print(f"tmpdir : {tmpdir}")
            if tmpdir == "" or tmpdir is None:
                tmpdir = '/tmp/'
        except:
            tmpdir = '/tmp/'
            
        service_folder_path = os.path.join(tmpdir, f"{service_name}.workflow")
        
        # Detele the folder and all it's content if it already exists
        if os.path.exists(service_folder_path):
            try:
                shutil.rmtree(service_folder_path)
                print(f"Folder deleted at {service_folder_path}")
            except OSError as e:
                print(f"Error deleting folder: {e}")
                
        # Create the service workflow folder
        try:
            os.mkdir(service_folder_path)
            print(f"Folder created at {service_folder_path}")
        except OSError as e:
            print(f"Error creating folder: {e}")
            
        # Create the Contents folder
        contents_folder_path = os.path.join(service_folder_path, "Contents")
        try:
            os.mkdir(contents_folder_path)
            print(f"Folder created at {contents_folder_path}")
        except OSError as e:
            print(f"Error creating folder: {e}")
            
        # Create the Info.plist file
        info_plist_file_path = os.path.join(contents_folder_path, "Info.plist")
        try:
            with open(info_plist_file_path, 'w') as file:
                file.write(info_plist_content)
            print(f"Content written to '{info_plist_file_path}' successfully.")
        except Exception as e:
            print(f"Error writing to file: {e}")
            
        # Create the document.wflow file
        document_wflow_file_path = os.path.join(contents_folder_path, "document.wflow")
        try:
            with open(document_wflow_file_path, 'w') as file:
                file.write(document_wflow_content)
            print(f"Content written to '{document_wflow_file_path}' successfully.")
        except Exception as e:
            print(f"Error writing to file: {e}")
            
        subprocess.Popen(["open", rf"{service_folder_path}"])

    def make_translate_command(self, for_service_template=False):
        self.save_target_language()
        
        docx_file_path = self.docx_file_entry.get()
        if platform.system() == 'Windows':
            docx_file_path = docx_file_path.replace('/', '\\')
        
        engine = self.engine.get().lower()
        src_lang_name = self.source_language.get()
        dest_lang_name = self.target_language.get()
        
        xlsx_file_path = self.xlsx_file_entry.get()
        if platform.system() == 'Windows':
            xlsx_file_path = xlsx_file_path.replace('/', '\\')
        
        font_value = self.font_var.get()
        
        if self.split_var.get():
            split_param = " --split "
        else:
            split_param = " "
        
        if self.show_browser_var.get():
            show_browser_param = " --showbrowser "
        else:
            show_browser_param = " "
        
        src_lang_code = 'en'
        if engine == 'deepl':
            src_lang_code = self.deepl_translate_lang_codes[src_lang_name]
            dest_lang_code = self.deepl_translate_lang_codes[dest_lang_name]
        else:
            src_lang_code = self.google_translate_lang_codes[src_lang_name]
            dest_lang_code = self.google_translate_lang_codes[dest_lang_name]
            
        split_string=""
        if not self.split_var.get():
            split_string = " - no split"
        
        if not docx_file_path and not for_service_template:
            messagebox.showerror("Error", "Please select a Word document file.")
            return
            
        xlsx_replace_param = " "
        if xlsx_file_path is not None and xlsx_file_path != "":
            if platform.system() == 'Windows':
                xlsx_replace_param = f" --xlsxreplacefile \"{xlsx_file_path}\""
            else:
                xlsx_replace_param = f" --xlsxreplacefile \\\"{xlsx_file_path}\\\""
       
        dest_font_param = " "
        if font_value is not None and font_value != "":
            if platform.system() == 'Windows':
                dest_font_param = f" --destfont \"{font_value}\""
            else:
                dest_font_param = f" --destfont \\\"{font_value}\\\""
        
        exitonsuccess_param = ""
        if self.open_file_after_translation_var.get() or platform.system() == 'Darwin':
            exitonsuccess_param = "  --exitonsuccess "

        if not os.path.exists(docx_file_path) and not for_service_template:
            messagebox.showerror("Error", "Selected Word document file does not exist.")
            return
            
        if self.target_language.get() == self.source_language.get():
            message = (
                "Please select a target translation language different than the source language."
            )
            messagebox.showinfo("Select a target translation language", message)
            return
        
        print(f"{self.bin_path}\\..")
        bin_launcher_path = ""
        if platform.system() == 'Windows':
            #bin_launcher_path = f"{self.bin_path}\\..\\ConEmuPack\\ConEmu.exe -ct -font \"Courier New\" -size 16 -run {self.bin_path}\\machine-translate-docx.exe "
            file_name = os.path.basename(docx_file_path)
            bin_launcher_path = f"{self.bin_path}\\..\\WindowsTerminal\\WindowsTerminal.exe --title \"Translating {file_name} in {dest_lang_name}\" --  {self.bin_path}\\machine-translate-docx.exe "
        else:
            if for_service_template:
                bin_launcher_path = f"\\\"{self.bin_path}/machine-translate-docx\\\" "
            else:
                bin_launcher_path = f"osascript -e 'tell app \"Terminal\" to do script \"\\\"{self.bin_path}/machine-translate-docx\\\" "
        
        open_word_param = ""
        if self.open_file_after_translation_var.get():
            open_word_param = "-l"
                
        if platform.system() == 'Windows':
            command = f"{bin_launcher_path} --srclang {src_lang_code} --destlang {dest_lang_code} --engine {engine} {dest_font_param} {split_param} {xlsx_replace_param} {show_browser_param} {exitonsuccess_param} {open_word_param} --docxfile \"{docx_file_path}\""
        else:
            if for_service_template:
                command = f"{bin_launcher_path} --srclang {src_lang_code} --destlang {dest_lang_code} --engine {engine} {dest_font_param} {split_param} {xlsx_replace_param} {show_browser_param} {exitonsuccess_param} {open_word_param} --docxfile \" &amp; thisItemsPathname &amp; \"; exit 0;"
                 
            else:
                command = f"{bin_launcher_path} --srclang {src_lang_code} --destlang {dest_lang_code} --engine {engine} {dest_font_param} {split_param} {xlsx_replace_param} {show_browser_param} {exitonsuccess_param} {open_word_param} --docxfile \\\"{docx_file_path}\\\" ; exitCode=$?; exit $exitCode; \"'"
        
        if platform.system() == 'Windows':
            pass
        else:
            command = f"{command}"
            
        print("command : %s" % (command))
        
        return command

    def translate(self):
        
        command = self.make_translate_command(for_service_template=False)
        
        print("command : %s" % (command))
        
        proc_translate = subprocess.Popen(command, shell=True)
        
        # If we want to force the program to wait
        if self.open_file_after_translation_var.get():
            # Force window to redraw with status bar
            #self.status_bar.config(text="Translating %s... please wait." % (os.path.basename(docx_file_path)))
            #self.root.update()
            
            #proc_translate.communicate()
            #proc_translate.wait()
            
            # Force window to redraw with status bar
            #self.status_bar.config(text=f"Status bar")
            #self.root.update()
            
            # Open the DOCX file in Windows
            try:
                if platform.system() == 'Windows':
                    subprocess.Popen(["start", "", rf"{docx_file_path}"], shell=True)
                elif platform.system() == "Darwin":  # macOS
                    subprocess.Popen(["open", rf"{docx_file_path}"])
                elif platform.system() == "Linux":  # Linux
                    subprocess.Popen(["xdg-open", rf"{docx_file_path}"])
                else:
                    print("Unsupported operating system.")
                    
                                
            except Exception as e:
                print("Error:", e)
                print(f"Warning, unable to open file: {docx_file_path}")
                
    
    def create_tooltip(self, widget, text):
        tool_tip = tk.Label(self.root, text=text, background="#ffffe0", relief="solid")
        tool_tip.place_forget()

        def show_tooltip(event):
            tool_tip.place(x=event.x_root - self.root.winfo_rootx() - 10, y=event.y_root - self.root.winfo_rooty() - 30)
        
        def hide_tooltip(event):
            tool_tip.place_forget()

        widget.bind("<Enter>", show_tooltip)
        widget.bind("<Leave>", hide_tooltip)

if __name__ == "__main__":
    # Parse parameters
    parser = argparse.ArgumentParser()
    parser.add_argument('--docxfile', required = False, help="Input file name")
    parser.add_argument('--binpath', required = False, help="path to the machine translation program")
    
    try:
        args = parser.parse_args()
    except:
        #print("Waiting for the input_element...")
        var = traceback.format_exc()
        print(var)
    
    try:
        word_file_to_translate = args.docxfile
    except:
        word_file_to_translate = None
    
    docxfile_path = ""
    if word_file_to_translate is not None:
        splitted_filename = os.path.splitext(os.path.basename(word_file_to_translate))

        # number of segment separated by dot in the docx filename
        splitted_filename_size = len(splitted_filename)

        docx_file_name =  "%s%s" % (splitted_filename[splitted_filename_size-2], splitted_filename[splitted_filename_size-1])

        if splitted_filename_size > 1:
            word_file_to_translate_extension = splitted_filename[splitted_filename_size-1].lower()

        if word_file_to_translate_extension == ".docx":        
            if not os.path.exists(word_file_to_translate) :
                print("Warning: File not found: %s. Ignoring." % (word_file_to_translate))
            else:
                docxfile_path = os.path.abspath(word_file_to_translate)
        else:
            print("Warning: not a word docx file: %s. Ignoring." % (word_file_to_translate))
    
    try:
        bin_path = args.binpath
    except:
        bin_path = None
        
    if bin_path is None:
        # determine if application is a script file or frozen exe
        if getattr(sys, 'frozen', False):
            bin_path = os.path.dirname(sys.executable)
        elif __file__:
            bin_path = os.path.dirname(__file__)
    
    root = tk.Tk()
    app = MachineTranslationApp(root, docxfile_path=docxfile_path, bin_path=bin_path)
    root.mainloop()
