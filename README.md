# Image to Excel

**Description**

	A python project to convert image file(s) to excel.

**Features**:

	* Zip
	* Logging
	* Custom Configuration

**Prerequisites**:

	* Python 3.6
	* xlsxwriter module
	* Pillow(PIL)
	* PyTesseract

> Note: PyTesseract works with the help of Tesseract-OCR Engine from Google. The Tesseract-OCR must be installed into the system and PATH must be configured. I haven't worked with Tesseract-OCR on Windows, therefore, I am not experienced with the installation process of  Tesseract-OCR in Windows. These are links to install it on windows:

> Link:  https://www.youtube.com/watch?v=YM8j9dzuKsk

### Install python 3.6 and add it to the PATH and check the version by:
	
	$ python --version

### Steps to install tesseract on windows:

* Open this [link](https://github.com/UB-Mannheim/tesseract/wiki)
* Download For 64-bit: tesseract-ocr-w64-setup-v4.0.0.20181030.exe or For 32-bit tesseract-ocr-w32-setup-v4.0.0.20181030.exe 
* Complete the installation in a directory like ProgramFiles\Tesseract-OCR.
* Copy the installation directory path (ex: C:\ProgramFiles\Tesseract-OCR)
* Press Win+R and type sysdm.cpl
* On system propertied click advanced tab and then Environment Variables. 
* Under the system, variables click on Path and append the path copied in step 3 with ';'(semicolon) as a delimeter. Click OK

### For checking the installation:
	* Press Win+R and type cmd.
	* In cmd/terminal

		$ tesseract

		You will get a lot of options if it is correctly installed.

### Installing Requirements :

	$ pip install -r requirements.txt

> This will install all the required modules mentioned in requirements.txt .

### To run the project:
	
	* Configure imageconfig.ini file and run:

	$ python image_to_excel.py









