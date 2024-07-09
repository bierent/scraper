import re
import fitz  
import pandas as pd

def find_next_word(text, keyword):
    pattern = re.compile(rf'{re.escape(keyword)}\s*((?:[\w-]+(?:\s*\d+)*)[\s\S]*?(?=\s+[A-Z]))')
    match = re.search(pattern, text)
    
    if match:
        next_word = match.group(1).strip()
        if next_word and not re.match(r'^\W+$', next_word) and "Volgende" not in next_word and "Testdatum" not in next_word:
            if keyword == "Type:" and next_word in ["3", "4", "5", "6", "9", "10"]:
                return f"{next_word} Voudig"
            return next_word
    return ""

def process_pdf(pdf_path, keywords):
    data = [] 
    try:
        pdf_document = fitz.open(pdf_path)

        for page_num in range(pdf_document.page_count):
            page = pdf_document.load_page(page_num)
            text = page.get_text()

            page_data = {"Page": page_num + 1} 
            
            for keyword in keywords:
                next_word = find_next_word(text, keyword)
                if next_word and next_word not in ["Volgende", "Testcodes:", "Testdatum:", "Testcodes", "Testdatum", "Volgende", "Testdatum: 5-1-2023", "Volgende testdatum:3-1-2024"]:
                    page_data[keyword] = next_word
                else:
                    page_data[keyword] = ""

            data.append(page_data)

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        if pdf_document:
            pdf_document.close()

    return data

if __name__ == "__main__":
    pdf_path = r"C:\Users\whereyourfileis"  # Replace with the path to your PDF file
    keywords = ["Objectnummer:", "Omschrijving:", "Merk:", "Type:", "Serienummer:", "Afdeling:", "Locatie:"]  # Replace with your keywords
        
    extracted_data = process_pdf(pdf_path, keywords)
    
    # Convert extracted_data into a pandas DataFrame
    df = pd.DataFrame(extracted_data)
    
    # Save the DataFrame to an Excel file
    excel_output_path = r"C:\Users\whereyouwantyouroutput"
    df.to_excel(excel_output_path, index=False)
    
    print(f"Extracted data saved to {excel_output_path}")