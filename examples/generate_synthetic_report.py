import sys
from pathlib import Path
sys.path.append(str(Path(__file__).resolve().parents[1]))
from reporting import generate_latex_report

report_context = {
    "project": {"client_name": "Cliente Sintético", "unit_name": "UC-001", "address": "Rua Exemplo, 100", "city": "São Paulo", "state": "SP", "report_date": "2026-04-29", "software_version": "1.0.0", "responsible_engineer": "Eng. Teste", "company": "Aurum"},
    "utility": {"company": "Concessionária X", "tariff_group": "Grupo A", "tariff_modality": "Verde", "billing_days": 30},
    "bill": {"energy_total_kwh": 12000, "energy_peak_kwh": 1800, "energy_offpeak_kwh": 10200, "demand_peak_kw": 220, "tariff_energy_peak": 0.95},
    "load_calibration": {"factor": 1.08, "tolerance_percent": 5.0, "errors": {"energy_total_percent": 3.2}},
    "pv_generation": {"annual_kwh": 98000, "specific_yield_kwh_kwp": 1480, "capacity_factor_percent": 17.2, "source": "PVWatts"},
    "hourly_balance": {"self_consumption_kwh": 76000, "grid_import_kwh": 44000, "export_kwh": 22000},
    "attachments_data": {"hourly_load": [[f"h{i}", i % 120] for i in range(8760)]},
}

print(generate_latex_report(report_context, "outputs/reports/cliente_sintetico"))
