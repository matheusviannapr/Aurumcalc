from __future__ import annotations

from datetime import datetime
from typing import Iterable, Sequence


def _to_hour(value: int | str) -> int:
    if isinstance(value, int):
        return value % 24
    if isinstance(value, str):
        chunk = value.strip().split(":")[0]
        return int(chunk) % 24
    raise ValueError("Horário inválido.")


def classify_peak_hours(
    hours: Sequence[int | datetime],
    peak_start: int | str,
    peak_end: int | str,
    weekdays_only: bool = False,
) -> list[str]:
    """Classifica cada hora como `ponta` ou `fora_ponta`.

    Aceita sequência de inteiros (0-23) ou datetimes.
    """
    start = _to_hour(peak_start)
    end = _to_hour(peak_end)

    def in_window(hour: int) -> bool:
        if start == end:
            return True
        if start < end:
            return start <= hour < end
        return hour >= start or hour < end

    result: list[str] = []
    for item in hours:
        if isinstance(item, datetime):
            hour = item.hour
            weekday = item.weekday()
            if weekdays_only and weekday >= 5:
                result.append("fora_ponta")
                continue
        else:
            hour = int(item) % 24
        result.append("ponta" if in_window(hour) else "fora_ponta")
    return result


def hourly_period_labels_for_month(
    month_index: Iterable[datetime],
    peak_start: int | str,
    peak_end: int | str,
    weekdays_only: bool = False,
) -> list[str]:
    return classify_peak_hours(list(month_index), peak_start, peak_end, weekdays_only)
