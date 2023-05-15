from fastapi import FastAPI, File, UploadFile, Request
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
import pandas as pd
import io
from openpyxl.utils import get_column_letter
from openpyxl import Workbook
from openpyxl.styles import Alignment

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")



# ranking ranges & location
lower_rank = 1
upper_rank = 2
country = 'US'


@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/filter_excel")
async def filter_excel(request:Request, file: UploadFile = File(...)):    
    # Load uploaded file into a pandas dataframe
    file_contents = await file.read()
    df = pd.read_excel(io.BytesIO(file_contents))
    
    # Filter the dataframe to get users above 45 years old
    df = df[df['Alexa Rank'] != '-']
    df = df[df['Country'] == country]
    df = df[~df['Alexa Rank'].str.endswith('M')]
    df = df[~df['Alexa Rank'].str.startswith('<1')]
    df['Alexa Rank'] = df['Alexa Rank'].str.strip().apply(lambda x: x[:-1]).astype(float)
    df = df[df['Alexa Rank'] >= lower_rank]
    df = df[df['Alexa Rank'] <= upper_rank]
    df['Alexa Rank'] = df['Alexa Rank'].astype(str)
    df['Alexa Rank'] = df['Alexa Rank'].apply(lambda x: x + 'K')    
    df.style.set_properties(**{'width': '80'})
    filtered_file_name = f"static/filtered_{file.filename}"
    df.to_excel(filtered_file_name, index=False)
        
    file_url = f"{filtered_file_name}"
    
    return templates.TemplateResponse("index.html", {'request': request, "file_url": file_url})


@app.get("/static/{file_path}")
async def get_static(file_path: str):
    return StaticFilesResponse(file_path)

