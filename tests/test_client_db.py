from pathlib import Path

from src.client_db import init_db, list_projects, load_project, upsert_project


def test_upsert_and_load_project(tmp_path: Path):
    db_path = tmp_path / "aurumcalc_test.db"
    init_db(db_path)

    payload = {"cliente_nome": "ACME", "current_step": 4, "historico_df": {"__type__": "dataframe", "records": []}}
    upsert_project("ACME", "Projeto 1", payload, db_path=db_path)

    loaded = load_project("ACME", "Projeto 1", db_path=db_path)
    assert loaded is not None
    assert loaded["cliente_nome"] == "ACME"
    assert loaded["current_step"] == 4


def test_upsert_updates_existing_project(tmp_path: Path):
    db_path = tmp_path / "aurumcalc_test.db"
    init_db(db_path)

    upsert_project("Beta", "Usina", {"v": 1}, db_path=db_path)
    upsert_project("Beta", "Usina", {"v": 2}, db_path=db_path)

    loaded = load_project("Beta", "Usina", db_path=db_path)
    assert loaded == {"v": 2}

    rows = list_projects("Beta", db_path=db_path)
    assert len(rows) == 1
    assert rows[0][0] == "Beta"
    assert rows[0][1] == "Usina"
