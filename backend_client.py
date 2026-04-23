from __future__ import annotations

from datetime import datetime, timedelta, timezone
import os
from pathlib import Path
from typing import Any

import httpx
import jwt
from dotenv import load_dotenv


BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")


class BackendApiError(RuntimeError):
    pass


class BackendApiClient:
    def __init__(
        self,
        base_url: str | None = None,
        token: str | None = None,
        timeout: float | None = None,
    ) -> None:
        self.base_url = (base_url or os.getenv("BACKEND_API_URL") or "http://localhost:8000/api/v1").rstrip("/")
        self._static_token = token or os.getenv("BACKEND_API_TOKEN") or os.getenv("WORKER_API_TOKEN")
        self.token = self._static_token
        self.timeout = timeout or float(os.getenv("BACKEND_API_TIMEOUT", "30"))
        if not self._static_token and not os.getenv("JWT_SECRET_KEY"):
            raise BackendApiError(
                "Token do backend ausente. Configure WORKER_API_TOKEN no database_updater/.env "
                "e o mesmo valor em backend/.env."
            )

    def list_pending_actions(self) -> list[dict[str, Any]]:
        return self._request("GET", "/sap-action-queue/pending")

    def complete_action_result(
        self,
        action_id: int,
        success: bool,
        error_message: str | None = None,
    ) -> dict[str, Any]:
        return self._request(
            "POST",
            f"/sap-action-queue/{action_id}/result",
            json={"success": success, "error_message": error_message},
        )

    def _request(self, method: str, path: str, allow_404: bool = False, **kwargs):
        headers = kwargs.pop("headers", {})
        token = self._get_token()
        if token:
            headers["Authorization"] = f"Bearer {token}"

        url = f"{self.base_url}{path}"
        try:
            with httpx.Client(timeout=self.timeout) as client:
                response = client.request(method, url, headers=headers, **kwargs)
        except httpx.HTTPError as exc:
            raise BackendApiError(f"Erro ao chamar backend {method} {url}: {exc}") from exc

        if allow_404 and response.status_code == 404:
            return None

        try:
            response.raise_for_status()
        except httpx.HTTPStatusError as exc:
            detail = self._extract_error_detail(response)
            raise BackendApiError(
                f"Backend retornou HTTP {response.status_code} em {method} {url}: {detail}"
            ) from exc

        if response.status_code == 204:
            return None
        return response.json()

    def _get_token(self) -> str | None:
        if self._static_token:
            return self._static_token
        return self._build_service_token()

    @staticmethod
    def _extract_error_detail(response: httpx.Response) -> str:
        try:
            body = response.json()
        except ValueError:
            return response.text
        return str(body.get("detail", body))

    @staticmethod
    def _build_service_token() -> str | None:
        secret = os.getenv("JWT_SECRET_KEY")
        if not secret:
            return None

        registration = os.getenv("SAP_USER") or os.getenv("WORKER_REGISTRATION") or "sap-worker"
        now = datetime.now(timezone.utc)
        payload = {
            "sub": registration,
            "registration": registration,
            "worker": True,
            "iat": now,
            "exp": now + timedelta(hours=float(os.getenv("BACKEND_TOKEN_EXPIRE_HOURS", "12"))),
        }
        return jwt.encode(payload, secret, algorithm="HS256")
