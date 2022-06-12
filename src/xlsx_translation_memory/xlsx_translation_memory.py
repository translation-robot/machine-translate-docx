from openpyxl import load_workbook
import datetime
import timeit
import re
from typing import Any
from typing import NamedTuple
from newmm_tokenizer.tokenizer import word_tokenize
import tinysegmenter
from pprint import pprint
import traceback
import os

# - *- coding: utf- 8 - *-

class SearchAndReplace(NamedTuple):
    """This class is used as a structure to contain a search and replace enty as a translation memory
    it is used in a list of translation memory contained in an excel file for object xlsx_translation_memory"""
    Search: str
    SearchRegularExpression: str
    Replace: str
    RegularExpression: bool
    CaseSensitive: bool
    IgnoreWordBoundary: bool
    KeepSpacesAtTheEndOfSearch: bool
    RegularExpressionCompiled: Any
    FollowedByAnotherWord : bool
    number_replacement: int

class xlsx_translation_memory():
    """class xlsx_translation_memory
    is a class that read excel xlsx files and build a list of search and replace
    with the following columns:
    Search	Replace	RegularExpression	CaseSensitive	IgnoreWordBoundary	KeepSpacesAtTheEndOfSearch
    Column Search is mandatory, column replace is optional
    columns CaseSensitive, IgnoreWordBoundary and KeepSpacesAtTheEndOfSearch are optional
    and defaults to False (F), to assign to True, use T value un excel cell."""
    def string_is_yes(self, string):
        """This method return True if a string's first non blank character is letter Y (case insensitive)"""
        res = False
        first_char = ''
        if string is not None:
            string_1st_char = string.strip()
            if len(string_1st_char) > 0:
                first_char = string_1st_char[0:1].upper()
        if first_char == 'Y':
            res = True
        return res

    def read_xlsx_search_and_replace(self):
        """"
        This method reads the content of the xlsx file
        and is called by method load_xlsx_translation_memory
        create a search and replace dictionary list for all the sheets"""
        try:
            print("Reading the content of the translation memory file %s" % (self.xlsx_path))
            
            self.ws = self.wb.active
            found_after_ws = False
            sheet_name = None
            
            search_multiple_blanks = r'[ \t]+'
            replace_single_blank = r' '
            re_multiple_blanks = re.compile(search_multiple_blanks)
            
            search_start_word_boundary = r'^\b.+'
            re_start_word_boundary = re.compile(search_start_word_boundary)
            
            search_end_word_boundary = r'.+\b$'
            re_end_word_boundary = re.compile(search_end_word_boundary)
			
            for s in range(len(self.wb.sheetnames)):
                self.wb.active = s
                self.ws = self.wb.active
                sheet_name = self.wb.sheetnames[s].lower().strip()
                	
                # We create an empty list of search and replace
                # for the sheet in the dictionary
                self.worksheets_search_and_replace_dictionary[sheet_name] = []
                search_and_replace_list = []

                rowno = 0
                for row in self.ws.values:
                    #print ("len=%s" % (len(row)))
                    rowno = rowno + 1
                    if rowno == 1:
                        continue
                    colpos = 0
                    while colpos < len(row):
                        #if (row[colpos] is not None):
                        #    #print("cell[%i,%i]=%s" % (rowno, colpos, row[colpos]))
                        colpos += 1
                    
                    Search = row[0]
                    if Search is None:
                        Search = ''
                        continue
                    else:
                        subn_result = re_multiple_blanks.subn(replace_single_blank, Search)
                        subn_count = subn_result[1]
                        Search = subn_result[0]
                    
                    Replace = row[1]
                    if Replace is None:
                        Replace = ''
                    
                    KeepSpacesAtTheEndOfSearch = row[5]
                    if self.string_is_yes(KeepSpacesAtTheEndOfSearch):
                        KeepSpacesAtTheEndOfSearch = True
                    else:
                        KeepSpacesAtTheEndOfSearch = False
                        Search = Search.strip()
                        Replace = Replace.strip()

                    RegularExpression = row[2]
                    RegularExpressionCompiled = None
                    if self.string_is_yes(RegularExpression):
                        RegularExpression = True
                        try:
                            RegularExpressionCompiled = re.compile(Search)
                            SearchRegularExpression = Search
                        except:
                            RegularExpression = False
                            RegularExpressionCompiled = None
                            SearchRegularExpression = re.escape(Search)
                            var = traceback.format_exc()
                            print(var)
                    else:
                        RegularExpression = False
                        SearchRegularExpression = re.escape(Search)
                    
                    IgnoreWordBoundary = row[4]
                    if self.string_is_yes(IgnoreWordBoundary):
                        IgnoreWordBoundary = True
                    else:
                        IgnoreWordBoundary = False
                        
                        if re_start_word_boundary.match (Search) is not None:
                            SearchRegularExpression = "\\b%s" % (SearchRegularExpression)
                        
                        if re_end_word_boundary.match (Search) is not None:
                            SearchRegularExpression = "%s\\b" % (SearchRegularExpression)

                    try:
                        FollowedByAnotherWord = row[6]
                        if self.string_is_yes(FollowedByAnotherWord):
                            FollowedByAnotherWord = True
                            SearchRegularExpression = "%s[ \\t]+[^ \\t]+" % SearchRegularExpression
                        else:
                            FollowedByAnotherWord = False
                    except:
                        FollowedByAnotherWord = False

                    number_replacement = 0
                    
                    CaseSensitive = row[3]
                    if self.string_is_yes(CaseSensitive):
                        CaseSensitive = True
                    else:
                        CaseSensitive = False
                        SearchRegularExpression = "(?i)%s" % (SearchRegularExpression)
                    
                    if Search is not None and rowno > 1:
                        try:
                            RegularExpressionCompiled = re.compile(SearchRegularExpression)
                        except:
                            RegularExpression = False
                            RegularExpressionCompiled = None
                            var = traceback.format_exc()
                            print(var)
                        
                        sr = SearchAndReplace(Search, SearchRegularExpression, Replace, RegularExpression, CaseSensitive, IgnoreWordBoundary, KeepSpacesAtTheEndOfSearch, RegularExpressionCompiled, FollowedByAnotherWord, number_replacement)
                        #if sheet_name == "keep_on_same_line":
                        #    print("search=%s" % (sr.Search))
                        search_and_replace_list.append(sr)
                    
                    self.worksheets_search_and_replace_dictionary[sheet_name] = search_and_replace_list
                    
        except Exception:
            print ("Error reading excel translation memory file.")
            var = traceback.format_exc()
            print("ERROR: %s" % (var))
            self.worksheets_search_and_replace_dictionary[sheet_name] = search_and_replace_list
    

    def search_and_replace_text(self, sheet_name, text_to_replace):
        """This method loop for all the search and replace list
        and does the search each in text_to_replace string"""
        nb_replacements = 0
        sheet_name = sheet_name.lower()
        text_replaced = text_to_replace
        try:
            #print("Search and replace text '%s'" % (text_replaced))
            se_no = 1
            for search_replace_item in self.worksheets_search_and_replace_dictionary[sheet_name]:
                #print("Search '%s' and replace '%s' # %d" % (search_replace_item.Search , search_replace_item.Replace ,se_no)) 
                #print("Search and replace # %d" % (se_no)) 
                subn_result = search_replace_item.RegularExpressionCompiled.subn(search_replace_item.Replace, text_replaced)
                subn_count = subn_result[1]
                if subn_count > 0:
                    search_replace_item._replace (number_replacement = search_replace_item.number_replacement + subn_count)
                    self.total_number_of_replacements = self.total_number_of_replacements + subn_count
                    nb_replacements = nb_replacements + subn_count
                    #print ("Original text :%s" % (text_replaced)) 
                    print ("Replaced '%s' by '%s' %d times (%s)." % (search_replace_item.Search, search_replace_item.Replace, subn_count, sheet_name))
                    #print ("Replaced text :%s" % (subn_result[0])) 
                text_replaced = subn_result[0]
                
                search_replace_item = search_replace_item._replace (number_replacement = search_replace_item.number_replacement + subn_count)
                self.worksheets_search_and_replace_dictionary[sheet_name][se_no-1] = search_replace_item
                #print ("res_str=%s, count=%i" % (res_str, subn_count))
                #pprint (search_replace_item)
                se_no = se_no + 1
            if nb_replacements > 0:
                print ("\n")
        except Exception:
            print ("Error reading excel translation memory file.")
            var = traceback.format_exc()
            print("ERROR: %s" % (var))
        return text_replaced, nb_replacements

    def tokenize_phrase(selfself, text, dest_lang):
        len_text = 0
        try:
            len_text = len(text)
        except:
            print("Error getting text length.")
        do_not_split_array = [0] * (len_text)

        if dest_lang.lower() == 'ja' or dest_lang.lower() == 'zh-cn' or dest_lang.lower() == 'zh' or dest_lang.lower() == 'zh-tw' or dest_lang.lower() == 'ko':
            words = self.cjk_segmenter.tokenize(text)
        # In other languages, just use spaces
        elif dest_lang.lower() == 'th':
            # words = thai_segmenter(text)
            words = word_tokenize(text)
        # In other languages, just use spaces
        else:
            words = text.split()
            #input("Setting do not split array values")
        return words

    def pprint_translation_memory_list(self):
        """Method pprint_translation_memory_list pprint the worksheets_search_and_replace_dictionary"""
        pprint (self.worksheets_search_and_replace_dictionary['before'])
        pprint (self.worksheets_search_and_replace_dictionary['after'])

    def get_sheet_number_of_replacements(self, sheet_name):
        """This method prints the number of replaced search items
        and the number of it they were replaced"""
        nb_replacements = 0
        sheet_name = sheet_name.lower()
        try:
            #print("Search and replace text '%s'" % (text_replaced))
            se_no = 1
            for search_replace_item in self.worksheets_search_and_replace_dictionary[sheet_name]:
                nb_replacements = self.worksheets_search_and_replace_dictionary[sheet_name][se_no-1].number_replacement + nb_replacements
                se_no = se_no + 1
        except Exception:
            print ("Error in get_sheet_number_of_replacements.")
            var = traceback.format_exc()
            print("ERROR: %s" % (var))
        return nb_replacements

    def print_replaced_items_number_of_replacements(self, sheet_name):
        """This method Method print_replaced_items_number_of_replacements prints the number of replaced search items
        and the number of it they were replaced"""
        sheet_name = sheet_name.lower()
        try:
            if self.wb is None:
                print ("\nTranslation search and replace file '%s' missing : ignored." % (self.xlsx_path))
            else:
                print ("\nSearch and replace number of replacements for excel sheet name '%s' :\n" % (sheet_name))
                se_no = 1
                for search_replace_item in self.worksheets_search_and_replace_dictionary[sheet_name]:
                    if search_replace_item.number_replacement > 0:
                        if search_replace_item.number_replacement > 1:
                            str_print =  "Replaced '%s' (excel line %d), by '%s' %d times."
                        else:
                            str_print =  "Replaced '%s' (excel line %d), by '%s' %d time."
                        print (str_print % (search_replace_item.Search , se_no + 1, search_replace_item.Replace, search_replace_item.number_replacement))
                    se_no = se_no + 1
                number_of_replacements = self.get_sheet_number_of_replacements(sheet_name)
                print ("\nNumber of replacements using sheet '%s' : %d from a list of %d entries.\n" % (sheet_name, number_of_replacements, len(self.worksheets_search_and_replace_dictionary[sheet_name])))
        except Exception:
            print ("Error calling methof print_replaced_items_number_of_replacements")
            var = traceback.format_exc()
            print("ERROR: %s" % (var))

    def load_xlsx_translation_memory(self, xlsx_path):
        """This method opens the xlsx file"""
        try:
            self.xlsx_path = xlsx_path
            self.worksheets_search_and_replace_dictionary = {}
            self.wb = load_workbook(self.xlsx_path)
            self.ws = self.wb.active
            self.read_xlsx_search_and_replace()
        except Exception:
            print ("Error loading excel translation memory file.")
            var = traceback.format_exc()
            print("ERROR: %s" % (var))
            self.wb = None
            self.ws = None

    def __init__(self, xlsx_path):
        """""""The constructor of class xlsx_translation_memory"""
        self.xlsx_path = xlsx_path
        self.total_number_of_replacements = 0
        self.worksheets_search_and_replace_dictionary = {}

        self.cjk_segmenter = tinysegmenter.TinySegmenter()

        print ("Init XLSXTranslationMemory")
        if not os.path.exists(xlsx_path) :
            print ("ERROR: File not found: %s" % (xlsx_path))
            self.wb = None
            self.ws = None
            self.worksheets_search_and_replace_dictionary = {}
        else:
            self.xlsx_splitted_filename = os.path.splitext(os.path.basename(xlsx_path))
            
            # number of segment separated by dot in the docx filename
            self.xlsx_splitted_filename_size = len(self.xlsx_splitted_filename)
            
            if self.xlsx_splitted_filename_size > 1:
                self.xlsx_file_to_translate_extension = self.xlsx_splitted_filename[self.xlsx_splitted_filename_size-1].lower()
             
            if self.xlsx_file_to_translate_extension == ".xlsx":
                self.load_xlsx_translation_memory(xlsx_path)
            else:
                print ("ERROR: translation memory file is not xlsx file: %s" % (xlsx_path))
                self.wb = None
                self.ws = None


def main():
    start = timeit.timeit()

    xtm = xlsx_translation_memory('wordlist_fr_v1.xlsx')
    res1 = xtm.search_and_replace_text("Bonjour Principal, quelle belle journée, Principal ! C'est fantastique !")
    res2 = xtm.search_and_replace_text("Fantastique !")
    print ("res1 = %s" % (res1))
    print ("res2 = %s" % (res2))
    #xtm.pprint_translation_memory_list()
    print ("total_number_of_replacements = %d" % (xtm.total_number_of_replacements))
	
    end = timeit.timeit()
    elapsedtime = end - start
    print ("\nXlsx file read. Elasped time: %s (h:mm:ss.mmm)" % str(datetime.timedelta(seconds=(end - start))))

if __name__ == '__main__':
    main()
