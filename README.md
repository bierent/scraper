# scraper
This script extracts specific information from PDF documents which I couldn't achieve with pdf editors or online tools. This is designed for processing technical documentation. 
The tool searches for predefined keywords and captures the relevant information following each keyword, organizing it into a structured Excel output. 

Key features:

- Extracts data based on specified keywords like Object Number, Description, Brand, Type, Serial Number, Department, and Location
- Handles multi-page PDF documents with robust text pattern matching
- Special handling for type numbers with "Voudig" suffix
- Exports results to an organized Excel spreadsheet
- Built using PyMuPDF (fitz), pandas, and regex for reliable PDF processing

Requires libraties: PyMuPDF, Pandas, re
