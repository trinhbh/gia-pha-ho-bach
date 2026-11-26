# -*- coding: utf-8 -*-
"""
Web Quản Trị Gia Phả Họ Bạch (FastAPI)
- Thừa kế dữ liệu từ Google Sheets (trên Cloud/Render) hoặc Excel (trên máy cá nhân)
- Xác thực qua Google Sign-In & phân quyền theo sheet 'nguoi_dung'
- Hỗ trợ Thêm / Sửa / Xóa nhân đinh & Lịch Kỵ Giỗ
- Xuất bản file HTML tĩnh độc lập (index.html) để public lên Internet
"""
import os
from typing import Optional, Union
from fastapi import FastAPI, HTTPException, Query, Request, Response, Depends
from fastapi.responses import HTMLResponse
from pydantic import BaseModel

import genealogy_service as svc
import auth_service as auth

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ADMIN_TEMPLATE_PATH = os.path.join(BASE_DIR, "templates", "admin.html")

app = FastAPI(title="Quản Trị Gia Phả Họ Bạch", version="3.0.0")


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
    thu_tu: Optional[Union[int, str]] = ""


class DeathPayload(BaseModel):
    ngay_str: str
    ten: str
    can: Optional[str] = ""
    mo: Optional[str] = ""
    ghi: Optional[str] = ""


class ExportPayload(BaseModel):
    hide_private: bool = False


class GoogleLoginPayload(BaseModel):
    credential: str


def _get_current_user(request: Request) -> Optional[dict]:
    token = request.cookies.get(auth.SESSION_COOKIE_NAME)
    if not token:
        return None
    session_data = auth.verify_session_token(token)
    if not session_data:
        return None
    # Đối chiếu lại với sheet nguoi_dung để đảm bảo quyền vẫn còn hiệu lực
    authorized = auth.check_authorized_email(session_data.get("email", ""))
    if not authorized:
        return None
    return {
        "email": authorized["email"],
        "name": session_data.get("name") or authorized["name"],
        "role": authorized["role"],
        "picture": session_data.get("picture", ""),
    }


def require_editor(request: Request) -> dict:
    token = request.cookies.get(auth.SESSION_COOKIE_NAME)
    if not token or not auth.verify_session_token(token):
        raise HTTPException(status_code=401, detail="Vui lòng đăng nhập bằng tài khoản Google để thực hiện thao tác này.")
    user = _get_current_user(request)
    if not user:
        raise HTTPException(status_code=403, detail="Email của bạn không có trong danh sách sheet 'nguoi_dung' hoặc đã bị khóa quyền.")
    return user


def _sync_viewer():
    """Tự động cập nhật Gia_Pha_Viewer.html ngay sau khi sửa dữ liệu (khi chạy local)."""
    if not svc.is_gsheets_enabled():
        svc.export_public_html(output_path=svc.HTML_TEMPLATE_PATH, hide_private=False, write_to_disk=True)


@app.get("/", response_class=HTMLResponse)
def serve_admin():
    if not os.path.exists(ADMIN_TEMPLATE_PATH):
        raise HTTPException(status_code=404, detail="Không tìm thấy giao diện quản trị templates/admin.html")
    with open(ADMIN_TEMPLATE_PATH, "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())


@app.get("/viewer", response_class=HTMLResponse)
def serve_viewer():
    result = svc.export_public_html(output_path=svc.HTML_TEMPLATE_PATH, hide_private=False, write_to_disk=not svc.is_gsheets_enabled())
    return HTMLResponse(content=result["html_content"])


@app.get("/public-preview", response_class=HTMLResponse)
def serve_public_preview(hide_private: bool = Query(False)):
    result = svc.export_public_html(output_path=svc.PUBLIC_HTML_PATH, hide_private=hide_private, write_to_disk=False)
    return HTMLResponse(content=result["html_content"])


@app.get("/api/auth/config")
def get_auth_config():
    return {
        "google_client_id": auth.get_google_client_id(),
        "storage_mode": "gsheets" if svc.is_gsheets_enabled() else "excel",
    }


@app.get("/api/auth/me")
def get_auth_me(request: Request):
    user = _get_current_user(request)
    return {"authenticated": bool(user), "user": user}


@app.post("/api/auth/google")
def login_with_google(payload: GoogleLoginPayload, request: Request, response: Response):
    try:
        g_user = auth.verify_google_credential(payload.credential)
    except ValueError as exc:
        raise HTTPException(status_code=401, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=401, detail=f"Token Google không hợp lệ: {exc}") from exc

    authorized = auth.check_authorized_email(g_user["email"])
    if not authorized:
        raise HTTPException(
            status_code=403,
            detail=f"Email '{g_user['email']}' chưa được cấp quyền trong sheet 'nguoi_dung'.",
        )

    user_info = {
        "email": authorized["email"],
        "name": g_user["name"] or authorized["name"],
        "role": authorized["role"],
        "picture": g_user.get("picture", ""),
    }
    session_token = auth.create_session_token(user_info)
    is_https = request.url.scheme == "https" or request.headers.get("x-forwarded-proto") == "https"
    response.set_cookie(
        key=auth.SESSION_COOKIE_NAME,
        value=session_token,
        httponly=True,
        samesite="lax",
        secure=is_https,
        max_age=auth.SESSION_TTL_SECONDS,
    )
    return {"status": "ok", "user": user_info}


@app.post("/api/auth/logout")
def logout(response: Response):
    response.delete_cookie(key=auth.SESSION_COOKIE_NAME)
    return {"status": "ok"}


@app.get("/api/data")
def get_all_data():
    try:
        return svc.load_all_data()
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"{type(exc).__name__}: {exc}") from exc


@app.post("/api/people")
def create_person(payload: PersonPayload, user: dict = Depends(require_editor)):
    try:
        new_stt = svc.add_person(payload.model_dump())
        _sync_viewer()
        return {"status": "ok", "stt": new_stt, "message": f"Đã thêm nhân đinh '{payload.ho_ten}' (STT {new_stt})."}
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@app.put("/api/people/{stt}")
def edit_person(stt: int, payload: PersonPayload, user: dict = Depends(require_editor)):
    try:
        svc.update_person(stt, payload.model_dump())
        _sync_viewer()
        return {"status": "ok", "stt": stt, "message": f"Đã cập nhật thông tin '{payload.ho_ten}'."}
    except ValueError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@app.delete("/api/people/{stt}")
def remove_person(stt: int, force: bool = Query(False), user: dict = Depends(require_editor)):
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
def create_death(payload: DeathPayload, user: dict = Depends(require_editor)):
    try:
        row_idx = svc.add_death(payload.model_dump())
        _sync_viewer()
        return {"status": "ok", "row_idx": row_idx, "message": f"Đã thêm ngày giỗ '{payload.ten}'."}
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@app.put("/api/deaths/{row_idx}")
def edit_death(row_idx: int, payload: DeathPayload, user: dict = Depends(require_editor)):
    try:
        svc.update_death(row_idx, payload.model_dump())
        _sync_viewer()
        return {"status": "ok", "message": f"Đã cập nhật ngày giỗ '{payload.ten}'."}
    except ValueError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@app.delete("/api/deaths/{row_idx}")
def remove_death(row_idx: int, user: dict = Depends(require_editor)):
    try:
        svc.delete_death(row_idx)
        _sync_viewer()
        return {"status": "ok", "message": "Đã xóa ngày giỗ."}
    except ValueError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@app.post("/api/export-html")
def export_html(payload: ExportPayload, user: dict = Depends(require_editor)):
    try:
        write_disk = not svc.is_gsheets_enabled()
        if write_disk:
            svc.export_public_html(output_path=svc.HTML_TEMPLATE_PATH, hide_private=False, write_to_disk=True)
            index_path = os.path.join(BASE_DIR, "index.html")
            svc.export_public_html(output_path=index_path, hide_private=payload.hide_private, write_to_disk=True)
        result = svc.export_public_html(output_path=svc.PUBLIC_HTML_PATH, hide_private=payload.hide_private, write_to_disk=write_disk)
        return {
            "status": "ok",
            "message": f"Đã xuất bản thành công! ({result['people_count']} nhân đinh, {result['deaths_count']} ngày giỗ)",
            "file_name": "index.html",
            "hide_private": payload.hide_private,
        }
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@app.get("/api/download-html")
def download_public_html(hide_private: bool = Query(False)):
    result = svc.export_public_html(output_path=svc.PUBLIC_HTML_PATH, hide_private=hide_private, write_to_disk=False)
    return Response(
        content=result["html_content"].encode("utf-8"),
        media_type="text/html; charset=utf-8",
        headers={"Content-Disposition": 'attachment; filename="index.html"'},
    )
