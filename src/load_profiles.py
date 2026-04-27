from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import numpy as np

TYPICAL_LOAD_PROFILES_PATH = Path("data/load_profiles/typical_load_profiles.json")


def load_typical_profiles(path: Path = TYPICAL_LOAD_PROFILES_PATH) -> list[dict[str, Any]]:
    payload = json.loads(path.read_text(encoding="utf-8"))
    return payload.get("profiles", [])


def get_profile_by_id(profile_id: str, path: Path = TYPICAL_LOAD_PROFILES_PATH) -> dict[str, Any]:
    profiles = load_typical_profiles(path)
    for profile in profiles:
        if profile.get("id") == profile_id:
            return profile
    raise ValueError(f"Perfil não encontrado: {profile_id}")


def validate_profile(profile: dict[str, Any], tol: float = 1e-3) -> list[str]:
    errors: list[str] = []
    curve = profile.get("curva_24h_pu", [])
    if len(curve) != 24:
        errors.append("A curva deve ter 24 pontos.")
        return errors
    total = float(np.sum(np.asarray(curve, dtype=float)))
    if abs(total - 1.0) > tol:
        errors.append(f"Soma da curva deve ser 1.0 (+/-{tol}), atual={total:.6f}")
    if any(float(v) < 0 for v in curve):
        errors.append("Curva não pode ter valores negativos.")
    return errors


def validate_all_profiles(path: Path = TYPICAL_LOAD_PROFILES_PATH, tol: float = 1e-3) -> list[str]:
    issues: list[str] = []
    for profile in load_typical_profiles(path):
        profile_issues = validate_profile(profile, tol=tol)
        for issue in profile_issues:
            issues.append(f"{profile.get('id')}: {issue}")
    return issues
