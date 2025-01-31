

Dynamic Document Generator

Overview  
The Dynamic Document Generator is a user-friendly Python and Streamlit application designed for creating professional documents. This tool supports multiple templates like SAT (VAT registration) and Service Agreements, offering dynamic input and customizable fields. Users can upload a signature image and generate polished documents effortlessly.

Key Features  
- Document Templates: Supports two types of templates:  
  - SAT Template: VAT registration and VAT filing document.  
  - Service Agreement: Customizable client agreement.  
- Dynamic Input: Replace placeholders in Word documents with input fields.  
- Signature Integration: Upload and embed a signature image into the document.  
- Custom Naming: Output files are dynamically named based on the client name and the current date.  

Technologies Used  
- Python: Core programming language.  
- Streamlit: For building an intuitive user interface.  
- python-docx: For manipulating Word documents.  

Setup Instructions  

Prerequisites  
- Python 3.8 or higher installed.  
- Install dependencies listed in `requirements.txt`.  

Installation Steps  
1. Clone the repository:  
   ```
   git clone https://github.com/Akshara-Amirtharaj/BKR-PDF-Generator
.git
   ```  
2. Navigate to the project directory:  
   ```
   cd Client-Agreement
   ```  
3. Install the required dependencies:  
   ```
   pip install -r requirements.txt
   ```  

Usage Instructions  
1. Launch the Streamlit application:  
   ```
   streamlit run app.py
   ```  
2. Select a document template:  
   - SAT Template  
   - Service Agreement  
3. Fill in the required fields.  
4. Upload a signature image (optional).  
5. Generate the document and download it directly.  

Contributing  
Contributions are welcome! Please fork the repository and submit a pull request for review.  

License  
This project is licensed under the MIT License.  

