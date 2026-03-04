from fastapi import FastAPI
from fastapi.responses import FileResponse
import subprocess

app = FastAPI()

SCRIPT="daily_list_pc_version.py"

def run(date,mode):

    subprocess.run(
        ["python",SCRIPT,date,mode]
    )

@app.get("/full")
def full(date:str):

    run(date,"full")

    return FileResponse(
        f"{date} list.pdf",
        media_type="application/pdf"
    )

@app.get("/client")
def client(date:str):

    run(date,"client")

    return FileResponse(
        "client.pdf",
        media_type="application/pdf"
    )

@app.get("/meals")
def meals(date:str):

    run(date,"meals")

    return FileResponse(
        "meals.pdf",
        media_type="application/pdf"
    )

@app.get("/delivery")
def delivery(date:str):

    run(date,"delivery")

    return FileResponse(
        "delivery.pdf",
        media_type="application/pdf"
    )

@app.get("/tags")
def tags(date:str):

    run(date,"tags")

    return FileResponse(
        "tags.pdf",
        media_type="application/pdf"
    )