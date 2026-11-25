# -*- coding: utf-8 -*-
"""
Web Quản Trị Gia Phả Họ Bạch (FastAPI)
- Thừa kế toàn bộ dữ liệu từ Danh_Sach_Nhan_Dinh_Ho_Bach.xlsx
- Hỗ trợ Thêm / Sửa / Xóa nhân đinh & Lịch Kỵ Giỗ
- Xuất bản file HTML tĩnh độc lập (index.html / Gia_Pha_Public.html) để public lên Internet
"""
import os
from typing import Optional, Union
from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import HTMLResponse, FileResponse
from pydantic import BaseModel

import genealogy_service as svc

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ADMIN_TEMPLATE_PATH = os.path.join(BASE_DIR, "templates", "admin.html")

app = FastAPI(title="Quản Trị Gia Phả Họ Bạch", version="2.0.0")


class PersonPayload(BaseModel):
    ho_ten: str
    ma: Optional[str] = ""
    can: Optional[str] = "Can 2"
    doi: Optional[str] = "Đời 8"
    huy: Optional[str] = ""
    gioi_tinh: Optional[str] = "Nam"
    nam_sinh: Optional[str] = ""
    nam_mat: Optional[str] = ""
    ngay_mat: Optional[str] = ""
    ngay_gio: Optional[str] = ""
    cha: Optional[str] = ""
    me: Optional[str] = ""
    phoi_ngau: Optional[str] = ""
    mo_phan: Optional[str] = ""
    ghi_chu: Optional[str] = ""
    id_cha: Optional[Union[int, str]] = 0
    sdt: Optional[str] = ""
    facebook: Optional[str] = ""
    dia_chi: Optional[str] = ""
    nghe_nghiep: Optional[str] = ""


class DeathPayload(BaseModel):
    ngay_str: str
    ten: str
    can: Optional[str] = ""
    mo: Optional[str] = ""
    ghi: Optional[str] = ""


class ExportPayload(BaseModel):
    hide_private: bool = False


@app.get("/", response_class=HTMLResponse)
def serve_admin():
    if not os.path.exists(ADMIN_TEMPLATE_PATH):
        raise HTTPException(status_code=404, detail="Không tìm thấy giao diện quản trị templates/admin.html")
    with open(ADMIN_TEMPLATE_PATH, "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())


@app.get("/viewer", response_class=HTMLResponse)
def serve_viewer():
    if not os.path.exists(svc.HTML_TEMPLATE_PATH):
        raise HTTPException(status_code=404, detail="Không tìm thấy Gia_Pha_Viewer.html")
    with open(svc.HTML_TEMPLATE_PATH, "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())


@app.get("/public-preview", response_class=HTMLResponse)
def serve_public_preview():
    target = svc.PUBLIC_HTML_PATH if os.path.exists(svc.PUBLIC_HTML_PATH) else svc.HTML_TEMPLATE_PATH
    with open(target, "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())


@app.get("/api/data")
def get_all_data():
    try:
        return svc.load_all_data()
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc


def _sync_viewer():
    """Tự động cập nhật Gia_Pha_Viewer.html ngay sau khi sửa dữ liệu Excel."""
    svc.export_public_html(output_path=svc.HTML_TEMPLATE_PATH, hide_private=False)


@app.post("/api/people")
def create_person(payload: PersonPayload):
    try:
        new_stt = svc.add_person(payload.model_dump())
        _sync_viewer()
        return {"status": "ok", "stt": new_stt, "message": f"Đã thêm nhân đinh '{payload.ho_ten}' (STT {new_stt})."}
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@app.put("/api/people/{stt}")
def edit_person(stt: int, payload: PersonPayload):
    try:
        svc.update_person(stt, payload.model_dump())
        _sync_viewer()
        return {"status": "ok", "stt": stt, "message": f"Đã cập nhật thông tin '{payload.ho_ten}'."}
    except ValueError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@app.delete("/api/people/{stt}")
def remove_person(stt: int, force: bool = Query(False)):
    try:
        svc.delete_person(stt, force=force)
        _sync_viewer()
        return {"status": "ok", "message": f"Đã xóa nhân đinh STT {stt}."}
    except ValueError as exc:
        msg = str(exc)
        status = 409 if "đang là Cha" in msg else 404
        raise HTTPException(status_code=status, detail=msg) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@app.post("/api/deaths")
def create_death(payload: DeathPayload):
    try:
        row_idx = svc.add_death(payload.model_dump())
        _sync_viewer()
        return {"status": "ok", "row_idx": row_idx, "message": f"Đã thêm ngày giỗ '{payload.ten}'."}
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@app.put("/api/deaths/{row_idx}")
def edit_death(row_idx: int, payload: DeathPayload):
    try:
        svc.update_death(row_idx, payload.model_dump())
        _sync_viewer()
        return {"status": "ok", "message": f"Đã cập nhật ngày giỗ '{payload.ten}'."}
    except ValueError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@app.delete("/api/deaths/{row_idx}")
def remove_death(row_idx: int):
    try:
        svc.delete_death(row_idx)
        _sync_viewer()
        return {"status": "ok", "message": "Đã xóa ngày giỗ."}
    except ValueError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@app.post("/api/export-html")
def export_html(payload: ExportPayload):
    try:
        # Cập nhật bản nội bộ đầy đủ
        svc.export_public_html(output_path=svc.HTML_TEMPLATE_PATH, hide_private=False)
        # Xuất bản Public HTML theo tùy chọn ẩn/hiện thông tin riêng tư (đồng bộ cả index.html cho GitHub Pages)
        result = svc.export_public_html(output_path=svc.PUBLIC_HTML_PATH, hide_private=payload.hide_private)
        index_path = os.path.join(BASE_DIR, "index.html")
        svc.export_public_html(output_path=index_path, hide_private=payload.hide_private)
        return {
            "status": "ok",
            "message": f"Đã xuất bản thành công! ({result['people_count']} nhân đinh, {result['deaths_count']} ngày giỗ)",
            "file_name": os.path.basename(result["output_path"]),
            "hide_private": payload.hide_private,
        }
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@app.get("/api/download-html")
def download_public_html():
    if not os.path.exists(svc.PUBLIC_HTML_PATH):
        svc.export_public_html(output_path=svc.PUBLIC_HTML_PATH, hide_private=False)
    return FileResponse(
        path=svc.PUBLIC_HTML_PATH,
        filename="index.html",
        media_type="text/html; charset=utf-8",
    )
