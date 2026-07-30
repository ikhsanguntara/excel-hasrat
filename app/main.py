from typing import Union, List, Dict, Any
from fastapi import FastAPI, HTTPException, Body, Response
from fastapi.responses import StreamingResponse
import datetime

from app.models import CheckpointItem
from app.excel_generator import generate_checkpoint_excel

app = FastAPI(
    title="Excel Checkpoint Generator API",
    description="API untuk menerima JSON checkpoint inspeksi dan menggenerate file Excel yang terformat rapi dengan pengolahan foto seragam & grouping.",
    version="1.0.0"
)

@app.get("/")
def read_root():
    return {
        "status": "online",
        "service": "Excel Checkpoint Generator API",
        "documentation": "/docs",
        "timestamp": datetime.datetime.now().isoformat()
    }

@app.post("/api/v1/generate-excel", response_class=StreamingResponse)
async def generate_excel_endpoint(
    payload: Union[List[Dict[str, Any]], Dict[str, Any]] = Body(...)
):
    """
    Endpoint POST untuk generate file Excel dari data JSON checkpoint.
    Dapat menerima payload berupa Array `[{...}, {...}]` atau Object `{"data": [{...}]}`.
    """
    try:
        # Extract list data
        if isinstance(payload, dict):
            if "data" in payload and isinstance(payload["data"], list):
                items = payload["data"]
            else:
                items = [payload]
        elif isinstance(payload, list):
            items = payload
        else:
            raise HTTPException(status_code=400, detail="Format JSON tidak valid. Gunakan Array atau Object dengan key 'data'.")

        if not items:
            raise HTTPException(status_code=400, detail="Data JSON kosong.")

        # Buat Excel stream
        excel_stream = generate_checkpoint_excel(items)

        # Ambil doc_num untuk nama file
        first_doc = items[0].get("doc_num", "EXPORT")
        safe_doc_num = str(first_doc).replace("/", "_").replace("\\", "_")
        filename = f"Laporan_Checkpoint_{safe_doc_num}.xlsx"

        headers = {
            "Content-Disposition": f'attachment; filename="{filename}"',
            "Access-Control-Expose-Headers": "Content-Disposition"
        }

        return StreamingResponse(
            excel_stream,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Gagal memproses file Excel: {str(e)}")
