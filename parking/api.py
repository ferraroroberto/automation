"""Thin GraphQL client. Deliberately exposes no delete/cancel operation."""

from __future__ import annotations

import logging
from typing import Any, Dict, List, Optional

import requests

logger = logging.getLogger("parking.api")

SLOTS_QUERY = (
    "query Days($center: Int!, $type: PlaceType, $size: Int) "
    "{ getBookingSlots(centerId: $center, type: $type, size: $size) { days } }"
)
BOOKINGS_QUERY = (
    "{ allBookings { id day state licensePlate "
    "place { place floor size type centerId } } }"
)
CREATE_MUTATION = (
    "mutation Book($center: String!, $type: PlaceType, $size: Int, "
    "$plate: String!, $days: [Date]!) "
    "{ createBooking(centerId: $center, type: $type, size: $size, "
    "licensePlate: $plate, days: $days) "
    "{ id day state licensePlate place { place floor size type centerId } } }"
)


class ApiError(Exception):
    """Unexpected API response (GraphQL errors, malformed payload)."""


class AuthError(ApiError):
    """401/403: the token is missing, expired or rejected."""


class RateLimited(ApiError):
    """429: back off, never push through."""


class ServerError(ApiError):
    """5xx from the API gateway."""


class ParkingApi:
    def __init__(self, endpoint: str, token: str, session: Optional[requests.Session] = None,
                 timeout: float = 20.0) -> None:
        self._endpoint = endpoint
        self._session = session or requests.Session()
        self._session.headers.update(
            {"authorization": f"Bearer {token}", "content-type": "application/json"}
        )
        self._timeout = timeout

    def _post(self, query: str, variables: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        response = self._session.post(
            self._endpoint, json={"query": query, "variables": variables or {}},
            timeout=self._timeout,
        )
        status = response.status_code
        if status in (401, 403):
            raise AuthError(f"HTTP {status}")
        if status == 429:
            raise RateLimited("HTTP 429")
        if status >= 500:
            raise ServerError(f"HTTP {status}")
        if status >= 400:
            raise ApiError(f"HTTP {status}")
        payload = response.json()
        if payload.get("errors"):
            message = "; ".join(str(e.get("message", e)) for e in payload["errors"])
            if "auth" in message.lower() or "token" in message.lower() or "jwt" in message.lower():
                raise AuthError(message)
            raise ApiError(message)
        return payload["data"]

    def slot_days(self, center_id: int, size_id: int, type_: str) -> List[str]:
        data = self._post(SLOTS_QUERY, {"center": center_id, "type": type_, "size": size_id})
        return list((data.get("getBookingSlots") or {}).get("days") or [])

    def my_bookings(self) -> List[Dict[str, Any]]:
        return list(self._post(BOOKINGS_QUERY).get("allBookings") or [])

    def create_booking(self, center_id: int, size_id: int, type_: str, plate: str,
                       raw_day: str) -> List[Dict[str, Any]]:
        """Book exactly one day. `raw_day` is passed back exactly as the availability query returned it."""
        data = self._post(
            CREATE_MUTATION,
            {"center": str(center_id), "type": type_, "size": size_id,
             "plate": plate, "days": [raw_day]},
        )
        return list(data.get("createBooking") or [])
