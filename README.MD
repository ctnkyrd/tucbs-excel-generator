# Excel Generator Configuration
## requires python 2.7 installed on machine

[link to Python2.7!](https://www.python.org/downloads/release/python-2715/)

### create a file named 
config.env for database connections.

### config.env
``` python
[DB]
DB_HOST=localhost
DB_PORT=5433
DB_DATABASE=tucbsdata
DB_USERNAME=postgres
DB_PASSWORD=kalman
```
### install required packages

```batch
pip install -r requirements.txt
```
### create excel files
inside the main folder
```python
python app.y
```
### create abstract forms
inside the OZET\analiz-form-fill from the cmd:
```
python loader.py
```
