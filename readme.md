### Step 1
Create a [virtual environment](https://realpython.com/python-virtual-environments-a-primer/):
```
python -m venv .
```
### Step 2
Activate the virtual environment:
```
.\Scripts\activate
```
### Step 3
```
pip install pdfkit
```
### Step 4
Download and install the appropriate executable for [wkhtmltopdf](https://wkhtmltopdf.org/)
> [!TIP]
> You can install to the folder that contains your virtual environment.<br>Wherever you install it, the executable will be within the ```bin``` folder.
### Step 5
Modify this portion of ```convert_html_to_pdf.py``` to point to the executable:<br><br>
![image](https://github.com/user-attachments/assets/73ab0b5f-206f-4624-b530-f2d53e3b34f9)
### Step 6
```
python convert_html_to_pdf.py
```
> [!IMPORTANT]
> IMPORTANT: The script will create a folder with the pdf files in the same directory that your script is located, NOT the directory containing the folder that you selected to process.
## Picture
![image](https://github.com/user-attachments/assets/89b19cc1-1e5a-4e85-bb13-c96b518c6a01)
