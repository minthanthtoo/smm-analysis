from __future__ import annotations

import os
import socket


def pick_available_port(candidates: list[int]) -> int:
    for port in candidates:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            try:
                sock.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    raise RuntimeError("No available port found in candidate list.")


def resolve_runtime_host_port() -> tuple[str, int]:
    render_port = os.getenv("PORT")
    if render_port:
        try:
            return "0.0.0.0", int(render_port)
        except ValueError as exc:
            raise RuntimeError("PORT environment variable must be an integer.") from exc

    return "127.0.0.1", pick_available_port([5055, 8000, 8080, 5000])
