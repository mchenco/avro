
## Installation
- Download these files: *'Download Code'* button > Download ZIP

## Configuring your dev environment
1. [Make sure you have python installed](https://docs.python.org/3.8/faq/windows.html)
	- Open *Command Prompt* and type `python` or `py`. Does it open up a python interpreter? If yes, you have python installed. Type `exit()` to get out of the interpreter.
	- If you don't ... idk man... [try this](https://www.liquidweb.com/kb/how-to-install-python-on-windows/)
2. [Download pip](https://www.liquidweb.com/kb/install-pip-windows/)
	- Download this [pip installer](https://bootstrap.pypa.io/get-pip.py) script
	- Open Command Prompt, navigate to the folder that you downloaded the above file to
	- Type `python get-pip.py` in Command Prompt
	- It should download lol idk you can type `pip -v` into the terminal to check if it did
3. In Command Prompt, navigate to the directory of the files you just downloaded from this repo.
	- You should have `script.py`, `requirements.txt` `template.xlsx` and `README.md`in the folder.
4. Type `pip install requirements.txt` to download all the libraries that the script needs to run.
5. I think that is it for dev set up !

## Running the script
1. Make sure your folder that has the script has the following things in it:
	- 1 template file called `template.xlsx`
	- Your two xlsx data files: *2-24h* and *48h*
	- `script.py`
2. In terminal, navigate to the folder that you downloaded these repo files to
3. Type `python script.py`
	- The screen will prompt you to enter the name of the output folder and the names of your data files.
	- I put it this way so it's not too limiting in case your files are named differently each time
	- **You don't need to create the output directory, it does it automatically**
	- **Don't include the `.xlsx`ext when entering filenames**
	- The script will run and output your 3 files to your folder
3. ENJOY :)


## Requirements
- python 3.6+
- pandas
- openpyxl

## Notes
- I really want to name the files according to the formulation lol my capricorn brain is unsatisfied
- If you rename any files, the script won't work lolllll sowwy
	- I can fix this if you're okay with re-entering the filenames in the terminal each time
- Use `ctrl + c` to quit the program at any time
