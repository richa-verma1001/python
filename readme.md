python virtual environments 

We are going to create a virutal environment at parent folder level 
    C:\Users\RichaVerma\github\richa-verma1001\python>

    To do this - 
    1. Delete the old evn folder
    ````
        cd path/to/python/ExcelScripts
        rm -rf venv      # Mac/Linux
        rmdir /s /q venv # Windows
    ````
    2. Recreate new one at parent
    ````
        cd path/to/python
        python -m venv venv
    ````
    3. Activate it 
    ````
        # Windows
        .\venv\Scripts\activate
        # or Mac/Linux
        source venv/bin/activate
    ````

    4. Reinstall dependencies 
        pip install pandas openpyxl tkinter

    5. Verify
        which python        # Mac/Linux
        where python        # Windows




