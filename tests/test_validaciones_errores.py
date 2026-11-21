"""Datos que son errores desde la 2.7.0: tareas de 0 días sin marca de hito (FEC-005), empresas que no son las de la
configuración (EMP-005) y proyecto sin la línea base seleccionada (FEC-009); los códigos EDT que no dependen de su
tarea superior (EDT-005) y los duplicados (EDT-009) ya lo eran."""
from datetime import datetime

from evm.config import Config
from evm.messages import SEV_ERROR
from evm.validation import validate_part
from evm.wbs import build_report

from builders import P, T

S, F = datetime(2026, 1, 5, 8), datetime(2026, 3, 27, 17)


def issues(tasks, cfg=None, **kw):
    cfg = cfg or Config()
    cfg.validate()
    part = validate_part(P(tasks, **kw), cfg)
    return part.issues + part.checks


def found(found_issues, code):
    return [i for i in found_issues if i.id == code]


def test_zero_duration_without_milestone_flag_is_an_error():
    tasks = [T(1, 1, "1", "Fase"), T(2, 2, "1.1", "A", duration=10.0),
             T(3, 2, "1.2", "Entrega", bs=F, bf=F, start=F, finish=F, duration=0.0),
             T(4, 2, "1.3", "Hito", bs=F, bf=F, start=F, finish=F, duration=0.0, milestone=True)]
    [issue] = found(issues(tasks, baseline_found=True), "FEC-005")
    assert (issue.code, issue.severity) == ("1.2.0", SEV_ERROR)


def test_company_not_in_the_configuration_is_an_error():
    cfg = Config(multi_enabled=True, company_a="EmpresaA", company_b="EmpresaB", company_both="Ambas")
    tasks = [T(1, 1, "1", "Fase", company="Ambas"), T(2, 2, "1.1", "A", company="EmpresaA"),
             T(3, 2, "1.2", "B", company="Otra")]
    [issue] = found(issues(tasks, cfg, fields_found={"campo_empresa": "Empresa"}), "EMP-005")
    assert issue.severity == SEV_ERROR and "Empresa 'Otra' no reconocida" in issue.message


def test_project_without_the_baseline_is_an_error():
    tasks = [T(1, 1, "1", "Fase", bs=None, bf=None), T(2, 2, "1.1", "A", bs=None, bf=None)]
    [issue] = found(build_report(P(tasks, baseline_found=False), Config()).issues, "FEC-009")
    assert issue.severity == SEV_ERROR


def test_wbs_codes_not_under_the_parent_and_duplicates_are_errors():
    tasks = [T(1, 1, "1", "Fase"), T(2, 2, "1.1", "A"), T(3, 3, "2.1.1", "Fuera de su superior"),
             T(4, 2, "1.2", "B"), T(5, 2, "1.2", "B repetida")]
    found_issues = issues(tasks, baseline_found=True)
    assert {i.severity for i in found(found_issues, "EDT-005")} == {SEV_ERROR}
    assert {i.severity for i in found(found_issues, "EDT-009")} == {SEV_ERROR}
    assert found(found_issues, "EDT-005") and found(found_issues, "EDT-009")


def test_split_is_the_share_of_company_a_and_zero_means_the_default():
    """% Reparto = 30: 30 % para la empresa A y 70 % para la B. Vacío o 0: reparto_por_defecto, nunca 0 %. En una
    tarea que no es «ambas» se ignora."""
    from evm.companies import explicit_share
    cfg = Config(multi_enabled=True, company_a="EmpresaA", company_b="EmpresaB", company_both="Ambas",
                 split_default=50.0)
    assert explicit_share(T(1, 1, "1", "A", company="Ambas", split_a=30.0), cfg) == (0.30, "")
    assert explicit_share(T(2, 1, "2", "B", company="Ambas", split_a=0.0), cfg) == (0.50, "")
    assert explicit_share(T(3, 1, "3", "C", company="Ambas", split_a=None), cfg) == (0.50, "")
    assert explicit_share(T(4, 1, "4", "D", company="EmpresaB", split_a=30.0), cfg) == (0.0, "")
