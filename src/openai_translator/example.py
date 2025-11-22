from translator import OpenAITranslator

if __name__ == "__main__":
    # Initialize translator with GPT-5 Nano
    oai_translator = OpenAITranslator(model="gpt-5-nano")
    
    # Set filename for the document
    oai_translator.set_filename("sample.docx")
    
    # First text to translate
    text1 = """I stayed with her till her last breath.
That turned me وگان for life.
Continue watching to find out more.
"Carlie Jackson (وگان):
"Flavorful Argentinian Sauces, Part 2 of 2 – وگان Salsa Golf on وگان Kabobs and وگان Salsa Tuco with وگان Tagliatelle Pasta.\""""
    
    response1, translation1 = oai_translator.translate("English", "Persian", text1)
    print("response1")
    print(response1)
    print("=== First Translation ===")
    print(translation1)
    
    # Second text to translate (for testing)
    text2 = """Hello world.
This is a second test.
Let's see how many lines OpenAI returns.
End of test."""
    
    response2, translation2 = oai_translator.translate("English", "French", text2)
    print("=== Second Translation ===")
    print(translation2)
    
    print("End of main")
    
