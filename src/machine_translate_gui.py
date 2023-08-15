import tkinter as tk
from tkinter import filedialog, messagebox, font
import subprocess
import os
import platform
import webbrowser
import argparse
import traceback

class MachineTranslationApp:
    def __init__(self, root, docxfile_path="", bin_path=None):
        self.root = root
        self.root.title("Word Docx Document Translator")
        
        try:
            self.root.iconbitmap("app.ico")
        except:
            pass
        
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
            'Bulgarian': 'bg',
            'Czech': 'cs',
            'Danish': 'da',
            'German': 'de',
            'Greek': 'el',
            'English': 'en',
            'Spanish': 'es',
            'Estonian': 'et',
            'Finnish': 'fi',
            'French': 'fr',
            'Hungarian': 'hu',
            'Indonesian': 'id',
            'Italian': 'it',
            'Japanese': 'ja',
            'Korean': 'ko',
            'Lithuanian': 'lt',
            'Latvian': 'lv',
            'Norwegian': 'nb',
            'Dutch': 'nl',
            'Polish': 'pl',
            'Portuguese': 'pt',
            'Romanian': 'ro',
            'Russian': 'ru',
            'Slovak': 'sk',
            'Slovenian': 'sl',
            'Swedish': 'sv',
            'Turkish': 'tr',
            'Ukrainian': 'uk',
            'Chinese (Simplified)': 'zh',
            }
            
        self.deepl_languages = self.deepl_translate_lang_codes.keys()
        self.google_languages = self.google_translate_lang_codes.keys()
        
        # Load stored target language value
        self.target_language_file = "target_language.txt"
        
        
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
            "English",  # Added "English" to the language list
            "Arabic", "Bulgarian", "Chinese (Simplified)", "Chinese (Traditional)", "Czech", "French",
            "German", "Hindi", "Hungarian", "Indonesian", "Italian", "Japanese", "Korean",
            "Malay", "Mongolian", "Nepali", "Persian", "Polish", "Portuguese", "Punjabi",
            "Romanian", "Russian", "Spanish", "Telugu", "Thai", "Ukrainian",
            "Urdu", "Vietnamese"
        ]
        self.source_language = tk.StringVar(root)
        self.source_language.set(self.languages[0])
        self.source_combo = tk.OptionMenu(root, self.source_language, *sorted(self.languages))
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

        engines = ["Deepl", "Google"]
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

        # Translate button
        self.translate_button = tk.Button(root, text="Translate", command=self.translate)
        self.translate_button.grid(row=9, column=1)
        
        # Create a status bar
        self.status_bar = tk.Label(root, text="Status bar", bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.grid(row=10, columnspan=3, sticky=tk.W+tk.E)
        
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
        xlsx_filename = f"C:\\SMTVRobot\\{target_language}.xlsx"

        self.xlsx_file_entry.delete(0, tk.END)
        
        if os.path.exists(xlsx_filename):
            self.xlsx_file_entry.insert(0, xlsx_filename)
    
    def populate_fonts(self):
        system_fonts = font.families()
        self.font_combo['menu'].delete(0, 'end')
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
        # List of languages for which Deepl should be selected

        if self.source_language.get() in self.deepl_languages and self.target_language.get() in self.deepl_languages:
            self.engine.set("Deepl")
            self.engine_combo.config(state="normal")
            self.show_browser_var.set(True)
            self.show_browser_checkbox.config(state="disabled")
        else:
            self.engine_combo.config(state="disabled")
            self.engine.set("Google")
            self.show_browser_var.set(False)
            self.show_browser_checkbox.config(state="normal")
            
    def update_show_browser(self, *args):
        if self.engine.get() == "Deepl":
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

    def translate(self):
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
        
        if not docx_file_path:
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
                xlsx_replace_param = f" --destfont \"{font_value}\""
            else:
                xlsx_replace_param = f" --destfont \\\"{font_value}\\\""
        
        exitonsuccess_param = ""
        if self.open_file_after_translation_var.get():
            exitonsuccess_param = "  --exitonsuccess "

        if not os.path.exists(docx_file_path):
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
            bin_launcher_path = f"{self.bin_path}\\..\\ConEmuPack\\ConEmu.exe -ct -font \"Lucida Console\" -size 16 -run {self.bin_path}\\machine-translate-docx.exe "
        else:
            bin_launcher_path = f"osascript -e 'tell app \"Terminal\" to do script \"{self.bin_path}/machine-translate-docx "
            
        if platform.system() == 'Windows':
            command = f"{bin_launcher_path} --srclang {src_lang_code} --destlang {dest_lang_code} --engine {engine} {dest_font_param} {split_param} {xlsx_replace_param} {show_browser_param} {exitonsuccess_param} --docxfile \"{docx_file_path}\""
        else:
            command = f"{bin_launcher_path} --srclang {src_lang_code} --destlang {dest_lang_code} --engine {engine} {dest_font_param} {split_param} {xlsx_replace_param} {show_browser_param} {exitonsuccess_param} --docxfile \\\"{docx_file_path}\\\";open \\\"{docx_file_path}\\\";exit 0;\"'"
        
        print("command : %s" % (command))
        
        proc_translate = subprocess.Popen(command, shell=True)
        
        # If we want to force the program to wait
        if self.open_file_after_translation_var.get():
            # Force window to redraw with status bar
            self.status_bar.config(text="Translating %s... please wait." % (os.path.basename(docx_file_path)))
            self.root.update()
            
            proc_translate.communicate()
            proc_translate.wait()
            
            # Force window to redraw with status bar
            self.status_bar.config(text=f"Status bar")
            self.root.update()
            
            
            # Open the DOCX file in Windows
            
            try:
                if platform.system() == 'Windows':
                    subprocess.Popen(["start", "", rf"{docx_file_path}"], shell=True)
                elif platform.system() ==  "Darwin":
                    subprocess.Popen(["open", rf"{docx_file_path}"])
            except Exception as e:
                print("Error:", e)
            


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
    
    print(f"bin_path = {bin_path}")
    
    root = tk.Tk()
    app = MachineTranslationApp(root, docxfile_path=docxfile_path, bin_path=bin_path)
    root.mainloop()
