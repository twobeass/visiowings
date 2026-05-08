"""Bounded retry helper with exponential backoff.

Used for re-establishing COM connections to Visio. We deliberately avoid a
heavyweight dependency like ``tenacity`` because the surface we need is a
single decorator with a couple of knobs, and a thin in-tree implementation
keeps the install footprint minimal.
"""

from __future__ import annotations

import functools
import logging
import time
from collections.abc import Callable
from typing import TypeVar

from .exceptions import COMConnectionError

T = TypeVar("T")
logger = logging.getLogger(__name__)


def retry_with_backoff(
    *,
    max_attempts: int = 3,
    initial_delay: float = 0.5,
    multiplier: float = 2.0,
    max_delay: float = 8.0,
    exceptions: tuple[type[BaseException], ...] = (Exception,),
    raise_on_failure: type[BaseException] = COMConnectionError,
) -> Callable[[Callable[..., T]], Callable[..., T]]:
    """Return a decorator that retries the wrapped function with exponential backoff.

    Args:
        max_attempts: Total attempts including the first call.
        initial_delay: Sleep duration in seconds before the second attempt.
        multiplier: Each subsequent delay is multiplied by this factor.
        max_delay: Cap for an individual delay.
        exceptions: Tuple of exception classes that should trigger a retry.
            Anything else propagates immediately.
        raise_on_failure: Exception class to raise once ``max_attempts`` is
            exhausted. By default :class:`COMConnectionError`.

    The wrapped function may return any value; on the final failure we raise
    ``raise_on_failure`` chained onto the original exception (``__cause__``).
    """

    if max_attempts < 1:
        raise ValueError("max_attempts must be >= 1")

    def decorator(func: Callable[..., T]) -> Callable[..., T]:
        @functools.wraps(func)
        def wrapper(*args: object, **kwargs: object) -> T:
            delay = initial_delay
            last_exc: BaseException | None = None
            for attempt in range(1, max_attempts + 1):
                try:
                    return func(*args, **kwargs)
                except exceptions as exc:
                    last_exc = exc
                    if attempt >= max_attempts:
                        break
                    logger.warning(
                        "%s failed (attempt %d/%d): %s; retrying in %.1fs",
                        func.__qualname__,
                        attempt,
                        max_attempts,
                        exc,
                        delay,
                    )
                    time.sleep(delay)
                    delay = min(delay * multiplier, max_delay)
            assert last_exc is not None  # for type narrowing
            if raise_on_failure is COMConnectionError:
                raise COMConnectionError(max_attempts, last_exc) from last_exc
            raise raise_on_failure(str(last_exc)) from last_exc

        return wrapper

    return decorator


__all__ = ["retry_with_backoff"]
