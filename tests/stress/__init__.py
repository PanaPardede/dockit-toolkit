"""
tests/stress package initialization.
Exposes the synthetic pathological input generator and stress testing fixtures.
"""

from typing import Any

__all__ = [
    "generate_extreme_pdfs",
    "generate_encrypted_pdfs",
    "generate_extreme_presentations",
    "generate_filesystem_edge_cases",
    "generate_all_fixtures",
    "hold_file_lock",
    "spawn_file_lock_process",
    "release_file_lock_process",
    "is_office_encrypted",
    "is_pdf_encrypted",
    "safe_rmtree",
    "to_extended_path",
    "write_catalog_manifest",
]


def __getattr__(name: str) -> Any:
    if name in __all__:
        from . import generate_stress_data

        return getattr(generate_stress_data, name)
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
