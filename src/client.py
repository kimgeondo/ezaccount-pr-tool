# src/client.py
from __future__ import annotations

import json
import os
import logging
from dataclasses import dataclass
from datetime import date, datetime
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import urljoin

import requests
from requests.auth import HTTPBasicAuth

logger = logging.getLogger(__name__)


def _yyyymmdd(d: date) -> str:
    return d.strftime("%Y%m%d")


def _parse_iso_datetime(s: str) -> datetime:
    # 예: "2026-05-13T15:19:10Z" 또는 "2026-05-13T15:19:10"
    return datetime.fromisoformat(s.replace("Z", "+00:00"))


@dataclass
class EzConfig:
    base_url: str
    list_endpoint: str
    detail_endpoint: str
    sap_client: str = "100"
    top: int = 1000

    # Credentials env var names
    username_env: str = "EZACCOUNT_USERNAME"
    password_env: str = "EZACCOUNT_PASSWORD"

    # PRListSet filter defaults
    bukrs: str = "3100"
    zqflag: str = "All"

    bukrs_field: str = "BUKRS"
    reqid_field: str = "REQID"
    bnumb_field: str = "BNUMB"
    bstat_field: str = "BSTAT"
    zqflag_field: str = "ZQFLAG"
    reqdt_from_field: str = "REQDTF"
    reqdt_to_field: str = "REQDTT"

    include_empty_bnumb: bool = True
    include_empty_bstat: bool = True

    # First run range
    initial_days: int = 365

    # PRHeaderSet expand
    # 너가 준 cURL 기준: PRItemNavi,FileAttachmentNavi,PRLinkNavi,ApprLineNavi,ApprLineAccNavi
    # 지금은 핵심만: PRItemNavi,ApprLineNavi
    header_expand: str = "PRItemNavi,ApprLineNavi"
    # optional multi-requester support
    requester_ids: Optional[List[str]] = None
    reqdp_filter: Optional[str] = None


class ODataClient:
    """
    SAP OData v2 client (HTTP Basic Auth)
    - env(.env 포함)에서 username/password 읽음
    - PRListSet 조회(증분)
    - PRHeaderSet(BUKRS,BNUMB,GJAHR) 조회 + $expand 지원
    """

    def __init__(self, config_path: str = "config/config.json", timeout: int = 60):
        self.timeout = timeout
        self.config = self._load_config(config_path)

        user, pw = self._get_credentials()
        self.session = requests.Session()
        self.session.auth = HTTPBasicAuth(user, pw)
        self.session.headers.update({
            "Accept": "application/json",
            "DataServiceVersion": "2.0",
            "MaxDataServiceVersion": "2.0",
            "X-Requested-With": "X",
        })

    # ------------------------
    # Config / credentials
    # ------------------------
    def _load_config(self, path: str) -> EzConfig:
        with open(path, "r", encoding="utf-8") as f:
            raw = json.load(f)

        cfg = EzConfig(
            base_url=raw["base_url"],
            list_endpoint=raw["list_endpoint"],
            detail_endpoint=raw["detail_endpoint"],
            sap_client=str(raw.get("sap_client", "100")),
            top=int(raw.get("top", 1000)),
            username_env=raw.get("username_env", "EZACCOUNT_USERNAME"),
            password_env=raw.get("password_env", "EZACCOUNT_PASSWORD"),
        )

        # optional overrides
        cfg.bukrs = str(raw.get("bukrs", cfg.bukrs))
        cfg.zqflag = str(raw.get("zqflag", cfg.zqflag))
        cfg.initial_days = int(raw.get("initial_days", cfg.initial_days))

        cfg.include_empty_bnumb = bool(raw.get("include_empty_bnumb", cfg.include_empty_bnumb))
        cfg.include_empty_bstat = bool(raw.get("include_empty_bstat", cfg.include_empty_bstat))

        cfg.bukrs_field = raw.get("bukrs_field", cfg.bukrs_field)
        cfg.reqid_field = raw.get("reqid_field", cfg.reqid_field)
        cfg.bnumb_field = raw.get("bnumb_field", cfg.bnumb_field)
        cfg.bstat_field = raw.get("bstat_field", cfg.bstat_field)
        cfg.zqflag_field = raw.get("zqflag_field", cfg.zqflag_field)
        cfg.reqdt_from_field = raw.get("reqdt_from_field", cfg.reqdt_from_field)
        cfg.reqdt_to_field = raw.get("reqdt_to_field", cfg.reqdt_to_field)

        cfg.header_expand = raw.get("header_expand", cfg.header_expand)

        # optional multi-requester / reqdp filter
        cfg.requester_ids = raw.get("requester_ids", cfg.requester_ids)
        cfg.reqdp_filter = raw.get("reqdp_filter", cfg.reqdp_filter)

        if not cfg.detail_endpoint:
            raise RuntimeError("config.json에 detail_endpoint(PRHeaderSet)가 없습니다.")
        return cfg

    def _get_credentials(self) -> Tuple[str, str]:
        user = os.getenv(self.config.username_env, "").strip()
        pw = os.getenv(self.config.password_env, "").strip()
        if not user or not pw:
            raise RuntimeError(
                f"Missing credentials. Set env vars in .env or OS env: "
                f"{self.config.username_env}, {self.config.password_env}"
            )
        return user, pw

    # ------------------------
    # HTTP / OData helpers
    # ------------------------
    def _build_url(self, endpoint: str) -> str:
        return urljoin(self.config.base_url, endpoint)

    def _request_json(self, url: str, params: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        params = params or {}

        # sap-client 고정
        if self.config.sap_client and "sap-client" not in params:
            params["sap-client"] = self.config.sap_client

        logger.debug("GET %s params=%s", url, params)
        resp = self.session.get(url, params=params, timeout=self.timeout)

        if resp.status_code in (401, 403):
            raise requests.HTTPError(f"{resp.status_code} Error: {resp.url}", response=resp)

        resp.raise_for_status()

        try:
            return resp.json()
        except Exception:
            snippet = resp.text[:500]
            raise RuntimeError(f"Non-JSON response from {resp.url}: {snippet}")

    @staticmethod
    def _odata_results(payload: Dict[str, Any]) -> List[Dict[str, Any]]:
        # OData v2 list: {"d":{"results":[...]}}
        d = payload.get("d")
        if d is None:
            return []
        if isinstance(d, dict) and isinstance(d.get("results"), list):
            return d["results"]
        if isinstance(d, dict):
            return [d]
        return []

    # ------------------------
    # PRListSet (list)
    # ------------------------
    def fetch_list(self, last_sync_iso: Optional[str], reqid_override: Optional[str] = None, reqdp_override: Optional[str] = None) -> List[Dict[str, Any]]:
        """
        PRListSet 조회(증분)
        - last_sync 있으면 그 날짜부터 오늘까지
        - 없으면 initial_days(기본 365일)부터 오늘까지
        """
        today = date.today()

        if last_sync_iso:
            try:
                from_date = _parse_iso_datetime(last_sync_iso).date()
            except Exception:
                # last_sync 파싱 실패 시 해당 월 1일
                from_date = date(today.year, today.month, 1)
        else:
            # initial_days 전부터
            from_date = today.fromordinal(today.toordinal() - self.config.initial_days)

        # REQID: use override if provided, otherwise current login user
        user, _ = self._get_credentials()
        effective_reqid = reqid_override if reqid_override is not None else user

        filters: List[str] = [
            f"{self.config.bukrs_field} eq '{self.config.bukrs}'",
            f"{self.config.reqid_field} eq '{effective_reqid}'",
            f"{self.config.zqflag_field} eq '{self.config.zqflag}'",
            f"{self.config.reqdt_from_field} eq '{_yyyymmdd(from_date)}'",
            f"{self.config.reqdt_to_field} eq '{_yyyymmdd(today)}'",
        ]

        # REQDP filter: prefer explicit override, otherwise config.reqdp_filter if set
        effective_reqdp = None
        if reqdp_override is not None:
            effective_reqdp = reqdp_override
        elif getattr(self.config, "reqdp_filter", None):
            effective_reqdp = self.config.reqdp_filter

        if effective_reqdp:
            filters.append(f"REQDP eq '{effective_reqdp}'")

        if self.config.include_empty_bnumb:
            filters.append(f"{self.config.bnumb_field} eq ''")
        if self.config.include_empty_bstat:
            filters.append(f"{self.config.bstat_field} eq ''")

        url = self._build_url(self.config.list_endpoint)
        params = {
            "$format": "json",
            "$top": self.config.top,
            "$filter": " and ".join(filters),
        }

        payload = self._request_json(url, params=params)
        return self._odata_results(payload)

    # ------------------------
    # PRHeaderSet (header + expand detail/approval)
    # ------------------------
    def fetch_header(self, bukrs: str, bnumb: str, gjahr: str) -> Dict[str, Any]:
        """
        PRHeaderSet(BUKRS='..',BNUMB='..',GJAHR='..')?$expand=...
        너가 준 cURL 구조와 동일.

        반환: OData v2 단건 {"d": {...}} 중 {...} dict
        """
        key_pred = f"(BUKRS='{bukrs}',BNUMB='{bnumb}',GJAHR='{gjahr}')"
        url = self._build_url(f"{self.config.detail_endpoint}{key_pred}")
        params = {
            "$format": "json",
            "$expand": self.config.header_expand,
        }

        payload = self._request_json(url, params=params)
        d = payload.get("d", {})
        return d if isinstance(d, dict) else {}