from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse
import io, json
from pathlib import Path
from gerar_excel import gerar_workbook

app = FastAPI()

@app.post("/gerar-orcamento")
async def gerar_orcamento_endpoint(payload: dict):
    produtos = payload.get("data")
    obra     = payload.get("obra")
    cliente  = payload.get("cliente")
    if not isinstance(produtos, list) or not obra or not cliente:
        raise HTTPException(400, "Payload inválido: precisa de data, obra e cliente")

    wb = gerar_workbook(produtos, obra, cliente)
    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    headers = {"Content-Disposition": "attachment; filename=orcamento_final.xlsx"}
    return StreamingResponse(
        stream,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers
    )

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("api:app", host="0.0.0.0", port=8000, reload=True)
