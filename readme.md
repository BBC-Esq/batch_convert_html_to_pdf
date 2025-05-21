### Step 1
create a [virtual environment](https://realpython.com/python-virtual-environments-a-primer/):
```
python -m venv .
```
### Step 2
Activate the virtual environment:
```
.\Scripts\activate
```
### Step 3
Download and install the appropriate executable for [wkhtmltopdf](https://wkhtmltopdf.org/)
  > You can install to the folder that contains your virtual environment.<br>
  > Wherever you install it, the executable will be within the ```bin``` folder.
### Step 4
Modify this portion of ```convert_html_to_pdf.py``` to point to the executable:<br><br>
![image](https://github.com/user-attachments/assets/73ab0b5f-206f-4624-b530-f2d53e3b34f9)
### Step 5
```
python convert_html_to_pdf.py
```
## Picture
![image](https://github.com/user-attachments/assets/89b19cc1-1e5a-4e85-bb13-c96b518c6a01)
