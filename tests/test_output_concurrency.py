"""Contratos interproceso de reserva y publicación de outputs (F05)."""

from __future__ import annotations

import multiprocessing
from contextlib import contextmanager
from pathlib import Path

from remasep.services import excel_writer as w
from remasep.services.metric_value_producer import PendingWrite
from remasep.testing import FakeWorkbookModel, FakeWriterHarness


def _concurrent_generate(
    label: str,
    template_path: str,
    output_path: str,
    entered,
    release,
    writer_called,
    results,
) -> None:
    pending = [
        PendingWrite(
            "wi:0", "LEGACY::X::0", "REMASEP_OD", "B10", 1,
            "INTEGER_COUNT", "Julio 2026",
        )
    ]
    model = FakeWorkbookModel(
        sheet_names=["REMASEP_OD", "CONTROL"],
        formula_map={("CONTROL", "E16"): "=E7"},
        values={("REMASEP_OD", "B10"): 0, ("CONTROL", "E16"): 0,
                ("CONTROL", "E7"): 0},
        writable_cells={("REMASEP_OD", "B10")},
        structural_fingerprint_id="stf:concurrency",
    )
    harness = FakeWriterHarness(model)
    control_map = w.load_control_map(
        {
            "sheet": "CONTROL",
            "total_errors_cell": "E16",
            "data_status_ok_value": "OK",
            "data_status_missing_value": "SIN DATOS",
            "modules_not_yet_produced": [],
            "rows": [],
        }
    )
    request = w.GenerationRequest(
        mode=w.MODE_PRODUCTION,
        template_path=Path(template_path),
        output_path=Path(output_path),
        period_year=2026,
        period_month=7,
        expected_fingerprint_id="stf:concurrency",
        zero_write_policy="WRITE_ZERO",
        zero_write_policy_accepted=("WRITE_ZERO",),
        manifest_instruction_ids=("wi:0",),
    )

    @contextmanager
    def controlled_writer(path):
        writer_called.set()
        if label == "first":
            entered.set()
            if not release.wait(10):
                raise RuntimeError("test release timeout")
        with harness.open_writer(path) as writer:
            yield writer

    result = w.generate(
        request,
        pending,
        open_writer=controlled_writer,
        control_map=control_map,
        inspect=harness.inspect,
    )
    results.put((label, result.status, tuple(result.errors), result.run_id))


def test_two_processes_exactly_one_publishes_and_loser_never_writes(tmp_path):
    ctx = multiprocessing.get_context("spawn")
    output = tmp_path / "same-output.xlsm"
    first_template = tmp_path / "first.xlsm"
    second_template = tmp_path / "second.xlsm"
    first_template.write_bytes(b"winner:first")
    second_template.write_bytes(b"winner:second")
    entered, release = ctx.Event(), ctx.Event()
    first_writer_called, second_writer_called = ctx.Event(), ctx.Event()
    results = ctx.Queue()

    first = ctx.Process(
        target=_concurrent_generate,
        args=(
            "first", str(first_template), str(output), entered, release,
            first_writer_called, results,
        ),
    )
    second = ctx.Process(
        target=_concurrent_generate,
        args=(
            "second", str(second_template), str(output), entered, release,
            second_writer_called, results,
        ),
    )
    first.start()
    assert entered.wait(10)
    second.start()
    second.join(10)
    release.set()
    first.join(10)

    assert first.exitcode == 0
    assert second.exitcode == 0
    outcomes = {label: (status, errors, run_id) for label, status, errors, run_id in
                (results.get(timeout=2), results.get(timeout=2))}
    assert outcomes["first"][0] == w.STATUS_GENERATED_DRAFT
    assert outcomes["second"][0] == w.STATUS_OUTPUT_EXISTS
    assert "OUTPUT_RESERVED" in outcomes["second"][1]
    assert first_writer_called.is_set()
    assert not second_writer_called.is_set()
    assert output.read_bytes() == b"winner:first"


def test_reservation_cleanup_after_managed_writer_exception(tmp_path):
    target = tmp_path / "report.xlsm"
    reservation = w.acquire_output_reservation(target, "run-without-pii")
    assert reservation is not None
    lock_path = reservation.lock_path
    try:
        raise RuntimeError("managed writer failure")
    except RuntimeError:
        reservation.release()
    assert not lock_path.exists()


def test_stale_reservation_is_not_deleted_by_age_or_foreign_owner(tmp_path):
    target = tmp_path / "patient-name-must-not-leak.xlsm"
    stale = w.acquire_output_reservation(target, "old-run-token")
    assert stale is not None
    assert target.name not in stale.lock_path.name
    assert target.name not in stale.lock_path.read_text(encoding="ascii")

    contender = w.acquire_output_reservation(target, "new-run-token")
    assert contender is None
    assert stale.lock_path.exists()
    stale.release()


def test_release_never_deletes_a_lock_replaced_by_another_owner(tmp_path):
    target = tmp_path / "report.xlsm"
    original = w.acquire_output_reservation(target, "original-run")
    assert original is not None
    original.lock_path.write_text('{"run_id":"foreign-run"}', encoding="ascii")

    assert original.release() is False
    assert original.lock_path.exists()
    original.lock_path.unlink()


def test_atomic_publication_refuses_an_unrelated_existing_destination(tmp_path):
    source = tmp_path / "complete-working.xlsm"
    target = tmp_path / "final.xlsm"
    source.write_bytes(b"our complete workbook")
    target.write_bytes(b"unrelated existing workbook")

    try:
        w._atomic_promote(source, target)
    except w.OutputAlreadyExistsError:
        pass
    else:
        raise AssertionError("publication replaced an unrelated destination")

    assert target.read_bytes() == b"unrelated existing workbook"
    assert source.read_bytes() == b"our complete workbook"
