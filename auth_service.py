# -*- coding: utf-8 -*-
"""
Dịch vụ xác thực Google Sign-In (OAuth 2.0 ID Token) & phân quyền theo sheet 'nguoi_dung'.
"""
import os
import hmac
import json
import time
import base64
import hashlib
import secrets
from dotenv import load_dotenv
from google.oauth2 import id_token
from google.auth.transport import requests as google_requests

import genealogy_service as svc

load_dotenv(os.path.join(svc.BASE_DIR, ".env"))

SESSION_COOKIE_NAME = "gia_pha_session"
SESSION_TTL_SECONDS = 7 * 24 * 3600  # 7 ngày
_FALLBACK_SECRET = secrets.token_hex(32)


def _get_session_secret() -> bytes:
    secret = os.environ.get("SESSION_SECRET", "").strip()
    if not secret:
        secret = _FALLBACK_SECRET
    return secret.encode("utf-8")


def get_google_client_id() -> str:
    return os.environ.get("GOOGLE_CLIENT_ID", "").strip()


def verify_google_credential(credential: str) -> dict:
    """Xác thực Google ID Token từ Google Identity Services."""
    client_id = get_google_client_id()
    if not client_id:
        raise ValueError("Hệ thống chưa cấu hình biến môi trường GOOGLE_CLIENT_ID.")
    if not credential or not credential.strip():
        raise ValueError("Thiếu Google credential token.")

    idinfo = id_token.verify_oauth2_token(
        credential.strip(),
        google_requests.Request(),
        client_id,
    )
    if not idinfo.get("email_verified"):
        raise ValueError("Email Google chưa được xác minh.")

    email = str(idinfo.get("email") or "").strip().lower()
    if not email:
        raise ValueError("Không lấy được địa chỉ Email từ tài khoản Google.")

    return {
        "email": email,
        "name": str(idinfo.get("name") or email).strip(),
        "picture": str(idinfo.get("picture") or "").strip(),
    }


def check_authorized_email(email: str, excel_path: str = svc.EXCEL_PATH) -> dict | None:
    """Kiểm tra email có nằm trong sheet 'nguoi_dung' và đang hoạt động không."""
    if not email:
        return None
    users = svc.load_authorized_users(excel_path=excel_path)
    return users.get(email.strip().lower())


def create_session_token(user_info: dict) -> str:
    """Tạo session token ký bằng HMAC-SHA256."""
    payload = {
        "email": user_info["email"],
        "name": user_info.get("name", ""),
        "role": user_info.get("role", "Biên tập"),
        "picture": user_info.get("picture", ""),
        "exp": int(time.time()) + SESSION_TTL_SECONDS,
    }
    raw = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    b64 = base64.urlsafe_b64encode(raw).decode("ascii").rstrip("=")
    sig = hmac.new(_get_session_secret(), b64.encode("ascii"), hashlib.sha256).hexdigest()
    return f"{b64}.{sig}"


def verify_session_token(token: str) -> dict | None:
    """Giải mã và xác thực chữ ký session token."""
    if not token or "." not in token:
        return None
    try:
        b64, sig = token.rsplit(".", 1)
        expected_sig = hmac.new(_get_session_secret(), b64.encode("ascii"), hashlib.sha256).hexdigest()
        if not hmac.compare_digest(sig, expected_sig):
            return None
        padded = b64 + "=" * (-len(b64) % 4)
        payload = json.loads(base64.urlsafe_b64decode(padded).decode("utf-8"))
        if int(payload.get("exp", 0)) < int(time.time()):
            return None
        return payload
    except Exception:
        return None
