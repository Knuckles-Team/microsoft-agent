#!/usr/bin/env python3
"""Run a bounded, privacy-safe A2A deployment validation."""

from __future__ import annotations

import argparse
import asyncio
import os
import time
import uuid
from urllib.parse import urlparse

import httpx


def _endpoint(value: str) -> str:
    parsed = urlparse(value)
    if (
        parsed.scheme not in {"http", "https"}
        or not parsed.netloc
        or parsed.username
        or parsed.password
        or parsed.fragment
    ):
        raise argparse.ArgumentTypeError(
            "A2A URL must be an absolute HTTP(S) URL without credentials or fragments"
        )
    return value


async def validate(url: str, query: str, timeout_seconds: float) -> int:
    """Submit one request and wait for a bounded terminal task state."""

    payload = {
        "jsonrpc": "2.0",
        "method": "message/send",
        "params": {
            "message": {
                "kind": "message",
                "role": "user",
                "parts": [{"kind": "text", "text": query}],
                "messageId": str(uuid.uuid4()),
            }
        },
        "id": 1,
    }
    deadline = time.monotonic() + timeout_seconds
    async with httpx.AsyncClient(timeout=min(timeout_seconds, 30.0)) as client:
        response = await client.post(
            url, json=payload, headers={"Content-Type": "application/json"}
        )
        response.raise_for_status()
        result = response.json()
        if "error" in result:
            print("A2A validation failed: JSON-RPC error returned")
            return 1
        task = result.get("result")
        if not isinstance(task, dict):
            print("A2A validation failed: response has no result object")
            return 1
        task_id = task.get("id")
        if not isinstance(task_id, str) or not task_id:
            print("A2A validation passed: synchronous result received")
            return 0

        while time.monotonic() < deadline:
            await asyncio.sleep(min(2.0, max(0.05, deadline - time.monotonic())))
            poll = await client.post(
                url,
                json={
                    "jsonrpc": "2.0",
                    "method": "tasks/get",
                    "params": {"id": task_id},
                    "id": 2,
                },
                headers={"Content-Type": "application/json"},
            )
            poll.raise_for_status()
            poll_result = poll.json()
            task_result = poll_result.get("result")
            if not isinstance(task_result, dict):
                print("A2A validation failed: polling response is invalid")
                return 1
            status = task_result.get("status")
            state = status.get("state") if isinstance(status, dict) else None
            if state not in {"submitted", "running", "working"}:
                passed = state in {"completed", "succeeded"}
                print(
                    "A2A validation passed"
                    if passed
                    else "A2A validation failed: task did not succeed"
                )
                return 0 if passed else 1

    print("A2A validation failed: task exceeded the configured timeout")
    return 1


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--url", default=os.getenv("A2A_URL"), type=_endpoint)
    parser.add_argument(
        "--query",
        default=os.getenv(
            "A2A_VALIDATION_QUERY", "Describe your available capabilities."
        ),
    )
    parser.add_argument("--timeout-seconds", type=float, default=120.0)
    args = parser.parse_args()
    if not args.url:
        parser.error("--url or A2A_URL is required")
    if not 5.0 <= args.timeout_seconds <= 600.0:
        parser.error("--timeout-seconds must be between 5 and 600")
    try:
        return asyncio.run(validate(args.url, args.query, args.timeout_seconds))
    except (httpx.HTTPError, ValueError):
        print("A2A validation failed: transport or response error")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
