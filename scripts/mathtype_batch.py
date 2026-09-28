#!/usr/bin/env python3
"""Batched, resumable MathType rendering for large Word documents.

A single Word session that converts hundreds of equations eventually exhausts Word (RPC failures
after roughly 200 Toggle TeX calls were observed), and one long call also exceeds client time
limits. This module renders a manifest in chunks:

* every chunk runs in a fresh, isolated Word process with ``-AllowUnresolvedMarkers`` so markers of
  later chunks may remain;
* after each chunk the document is saved as a checkpoint and the job state is written, so a crashed
  or interrupted job resumes from the last checkpoint (same input, manifest and output);
* a failed chunk is retried once, split in half;
* numbered equations that are referenced, and all references, are rendered together in the last
  chunk, because MathType inserts a reference by clicking its live number field;
* equation preferences, the table layout, the inline line-spacing fix and the strict final
  validation run once, after the last chunk.

Job state lives in %APPDATA%\\MathTypeForWordAgent\\jobs\\<hash of output path>\\, never next to the
user's files. The job can run in the foreground or as a detached background process.

Usage: mathtype_batch.py render --input IN.docx --output OUT.docx --manifest M.json
                                [--batch-size 40] [--no-resume] [--overwrite]
       mathtype_batch.py status --output OUT.docx [--cleanup]
Prints one JSON result line.
"""
from __future__ import annotations

import argparse
import datetime as _dt
import hashlib
import json
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path
from typing import Any, Callable

sys.path.insert(0, str(Path(__file__).resolve().parent))

DEFAULT_BATCH_SIZE = 40
STATE_SCHEMA = "mathtype-for-word-job/1"


def _now() -> str:
    return _dt.datetime.now().isoformat(timespec="seconds")


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def jobs_root() -> Path:
    base = os.environ.get("MATHTYPE_JOBS_DIR") or os.path.join(
        os.environ.get("APPDATA") or str(Path.home()), "MathTypeForWordAgent", "jobs")
    return Path(base)


def job_dir(output_path: str) -> Path:
    key = hashlib.sha256(str(Path(output_path).resolve()).lower().encode("utf-8")).hexdigest()[:16]
    return jobs_root() / key


def load_manifest(path: str) -> dict:
    manifest = json.loads(Path(path).read_text(encoding="utf-8-sig"))
    if manifest.get("schema_version") != 1 or not isinstance(manifest.get("equations"), list):
        raise ValueError("Manifest must be schema_version 1 with an equations array.")
    ids, markers = set(), set()
    for equation in manifest["equations"]:
        if not equation.get("id") or equation["id"] in ids:
            raise ValueError(f"Missing or duplicate equation id: {equation.get('id')!r}")
        if not equation.get("marker") or equation["marker"] in markers:
            raise ValueError(f"Missing or duplicate marker: {equation.get('marker')!r}")
        ids.add(equation["id"])
        markers.add(equation["marker"])
    numbered = {e["id"] for e in manifest["equations"] if e.get("numbered")}
    for reference in manifest.get("references") or []:
        if reference.get("target") not in numbered:
            raise ValueError(f"Reference target is missing or not numbered: {reference.get('target')!r}")
        if not reference.get("marker") or reference["marker"] in markers:
            raise ValueError(f"Missing or duplicate reference marker: {reference.get('marker')!r}")
        markers.add(reference["marker"])
    return manifest


def plan_chunks(manifest: dict, batch_size: int) -> list[list[str]]:
    """Equation ids per chunk. Referenced numbered equations go with the references into the last chunk."""
    references = manifest.get("references") or []
    targets = {r["target"] for r in references}
    ordinary = [e["id"] for e in manifest["equations"] if e["id"] not in targets]
    size = batch_size if batch_size and batch_size > 0 else max(1, len(ordinary))
    chunks = [ordinary[k:k + size] for k in range(0, len(ordinary), size)]
    if references:
        final = [e["id"] for e in manifest["equations"] if e["id"] in targets]
        chunks.append(final)
    return chunks or [[]]


def chunk_manifest(manifest: dict, ids: list[str], include_references: bool) -> dict:
    wanted = set(ids)
    part = {k: v for k, v in manifest.items() if k not in ("equations", "references")}
    part["equations"] = [e for e in manifest["equations"] if e["id"] in wanted]
    part["references"] = list(manifest.get("references") or []) if include_references else []
    return part


class Job:
    def __init__(self, input_path: str, output_path: str, manifest_path: str, batch_size: int) -> None:
        self.input = Path(input_path).resolve()
        self.output = Path(output_path).resolve()
        self.manifest_path = Path(manifest_path).resolve()
        self.batch_size = batch_size
        self.dir = job_dir(str(self.output))
        self.state_path = self.dir / "state.json"
        self.checkpoint = self.dir / "checkpoint.docx"
        self.log_path = self.dir / "job.log"
        self.state: dict[str, Any] = {}

    def log(self, message: str) -> None:
        self.dir.mkdir(parents=True, exist_ok=True)
        line = f"[{_now()}] {message}"
        with self.log_path.open("a", encoding="utf-8") as handle:
            handle.write(line + "\n")
        print(line, file=sys.stderr, flush=True)

    def save(self, **changes: Any) -> None:
        self.state.update(changes, updated=_now())
        self.dir.mkdir(parents=True, exist_ok=True)
        temporary = self.state_path.with_suffix(".tmp")
        temporary.write_text(json.dumps(self.state, ensure_ascii=False, indent=1), encoding="utf-8")
        os.replace(temporary, self.state_path)

    def load_existing(self) -> dict | None:
        if not self.state_path.is_file():
            return None
        try:
            return json.loads(self.state_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return None


def _bridge_timeout(equations: int) -> int:
    base = int(os.environ.get("MATHTYPE_CHUNK_BASE_TIMEOUT_SECONDS", "120"))
    per = float(os.environ.get("MATHTYPE_SECONDS_PER_EQUATION", "8"))
    return int(base + per * max(1, equations))


def render(input_path: str, output_path: str, manifest_path: str, batch_size: int = DEFAULT_BATCH_SIZE,
           resume: bool = True, overwrite: bool = False, allow_unresolved_markers: bool = False,
           bridge: Callable[..., dict] | None = None, postprocess: Callable[..., dict] | None = None) -> dict:
    """Render ``manifest`` into ``output`` in chunks. ``bridge``/``postprocess`` are injectable for tests."""
    import mcp_server

    bridge = bridge or mcp_server._invoke_bridge
    postprocess = postprocess or mcp_server._postprocess_word
    job = Job(input_path, output_path, manifest_path, batch_size)
    if not job.input.is_file():
        raise FileNotFoundError(f"InputPath does not exist: {job.input}")
    if job.input.suffix.lower() != ".docx":
        raise ValueError("InputPath must be a .docx file.")
    if job.output == job.input:
        raise ValueError("Write the rendered DOCX to a new path; the source is preserved.")
    manifest = load_manifest(str(job.manifest_path))
    fingerprint = {
        "input_sha256": _sha256(job.input),
        "manifest_sha256": _sha256(job.manifest_path),
        "output": str(job.output),
    }
    previous = job.load_existing()
    resuming = bool(
        resume and previous and previous.get("status") in ("running", "failed", "interrupted")
        and all(previous.get(k) == v for k, v in fingerprint.items())
        and job.checkpoint.is_file()
    )
    if job.output.exists() and not overwrite and not resuming:
        raise FileExistsError(f"Output exists (pass overwrite=true): {job.output}")
    if resuming:
        job.state = previous
        job.log(f"Resuming job at chunk {job.state['completed_chunks'] + 1}/{len(job.state['chunks'])}.")
    else:
        job.dir.mkdir(parents=True, exist_ok=True)
        # a fresh start clears old state but keeps background.out, which a detached parent may hold open
        for stale in list(job.dir.glob("chunk-*")) + [job.checkpoint, job.log_path, job.state_path]:
            stale.unlink(missing_ok=True)
        shutil.copy2(job.input, job.checkpoint)
        chunks = plan_chunks(manifest, batch_size)
        job.state = {
            "schema": STATE_SCHEMA, **fingerprint, "input": str(job.input), "manifest": str(job.manifest_path),
            "status": "running", "pid": os.getpid(), "started": _now(), "batch_size": batch_size,
            "equations_total": len(manifest["equations"]), "references_total": len(manifest.get("references") or []),
            "chunks": chunks, "completed_chunks": 0, "equations_done": 0, "retries": 0,
            "log_path": str(job.log_path), "error": None, "result": None,
        }
        job.save()
        job.log(f"Started: {len(manifest['equations'])} equation(s) in {len(chunks)} chunk(s) of up to {batch_size}.")
    job.save(status="running", pid=os.getpid(), error=None)

    has_references = bool(manifest.get("references"))
    while job.state["completed_chunks"] < len(job.state["chunks"]):
        index = job.state["completed_chunks"]
        ids = job.state["chunks"][index]
        last = index == len(job.state["chunks"]) - 1
        part = chunk_manifest(manifest, ids, include_references=last and has_references)
        part_path = job.dir / f"chunk-{index + 1:03d}.json"
        part_path.write_text(json.dumps(part, ensure_ascii=False), encoding="utf-8")
        target = job.dir / "chunk-output.docx"
        if target.exists():
            target.unlink()
        started = time.monotonic()
        job.log(f"Chunk {index + 1}/{len(job.state['chunks'])}: {len(ids)} equation(s), "
                f"{len(part['references'])} reference(s).")
        try:
            result = bridge("render", {
                "input_path": str(job.checkpoint), "output_path": str(target), "manifest_path": str(part_path),
                "overwrite": True, "allow_unresolved_markers": True,
            }, timeout=_bridge_timeout(len(ids) + len(part["references"])))
        except BaseException as exc:  # Ctrl+C, a killed client: leave a resumable state behind
            job.save(status="interrupted", error=f"{type(exc).__name__}: {exc}", failed_chunk=index + 1)
            job.log(f"Chunk {index + 1} interrupted: {type(exc).__name__}: {exc}")
            raise
        if result.get("ok") and target.is_file():
            os.replace(target, job.checkpoint)
            job.save(completed_chunks=index + 1, equations_done=job.state["equations_done"] + len(ids))
            job.log(f"Chunk {index + 1} done in {time.monotonic() - started:.0f} s "
                    f"({job.state['equations_done']}/{job.state['equations_total']} equations).")
            continue
        error = result.get("error") or "; ".join(result.get("errors") or []) or "unknown bridge failure"
        job.log(f"Chunk {index + 1} failed: {error}")
        if len(ids) > 1 and not (last and has_references) and job.state["retries"] < len(job.state["chunks"]) * 2:
            half = len(ids) // 2
            job.state["chunks"][index:index + 1] = [ids[:half], ids[half:]]
            job.save(retries=job.state["retries"] + 1)
            job.log(f"Retrying chunk {index + 1} split into {half} + {len(ids) - half} equation(s).")
            continue
        job.save(status="failed", error=error, failed_chunk=index + 1, bridge_result=result)
        return {
            "ok": False, "action": "render-batched", "error": f"Chunk {index + 1} failed: {error}",
            "resume": "Call render again with the same input, output and manifest (resume=true) to continue "
                      "from the last checkpoint.", "job": _public_state(job.state), "bridge_result": result,
        }

    job.log("All chunks rendered; post-processing and final validation.")
    if job.output.exists() and not overwrite and not resuming:
        raise FileExistsError(f"Output exists (pass overwrite=true): {job.output}")
    job.output.parent.mkdir(parents=True, exist_ok=True)
    temporary = job.output.with_name(f".{job.output.stem}.{os.getpid()}.tmp.docx")
    shutil.copy2(job.checkpoint, temporary)
    os.replace(temporary, job.output)
    final = postprocess(str(job.output), manifest, str(job.manifest_path), allow_unresolved_markers)
    status = "completed" if final.get("ok") else "failed"
    summary = {
        "ok": bool(final.get("ok")), "action": "render", "input_path": str(job.input),
        "output_path": str(job.output), "equations": len(manifest["equations"]),
        "numbered_equations": sum(1 for e in manifest["equations"] if e.get("numbered")),
        "references": len(manifest.get("references") or []), "chunks": len(job.state["chunks"]),
        "retries": job.state["retries"], "number_format": "(1), (2), (3), ...",
        "reference_mechanism": "MathType-native GOTOBUTTON/REF fields",
        "reference_brackets": manifest.get("reference_brackets") or "fullwidth", **final,
    }
    summary["ok"] = bool(final.get("ok"))
    job.save(status=status, result={k: v for k, v in summary.items() if k != "post_processing"},
             error=None if summary["ok"] else final.get("error") or "final validation failed")
    job.log(f"Job {status}.")
    for leftover in job.dir.glob("chunk-*"):
        leftover.unlink(missing_ok=True)
    if summary["ok"]:
        job.checkpoint.unlink(missing_ok=True)
    return summary


def _public_state(state: dict) -> dict:
    keys = ("status", "input", "output", "started", "updated", "equations_total", "equations_done",
            "completed_chunks", "retries", "error", "failed_chunk", "log_path", "pid")
    public = {k: state.get(k) for k in keys if k in state}
    public["chunks_total"] = len(state.get("chunks") or [])
    return public


def _pid_alive(pid: int | None) -> bool:
    if not pid:
        return False
    if os.name == "nt":
        import ctypes

        handle = ctypes.windll.kernel32.OpenProcess(0x1000, False, int(pid))
        if not handle:
            return False
        code = ctypes.c_ulong()
        ctypes.windll.kernel32.GetExitCodeProcess(handle, ctypes.byref(code))
        ctypes.windll.kernel32.CloseHandle(handle)
        return code.value == 259  # STILL_ACTIVE
    try:
        os.kill(int(pid), 0)
        return True
    except OSError:
        return False


def status(output_path: str, cleanup: bool = False, log_lines: int = 15) -> dict:
    directory = job_dir(output_path)
    state_path = directory / "state.json"
    if not state_path.is_file():
        exists = Path(output_path).is_file()
        return {"ok": True, "action": "status", "status": "none", "output_path": str(Path(output_path).resolve()),
                "output_exists": exists, "message": "No job state for this output path."}
    state = json.loads(state_path.read_text(encoding="utf-8"))
    public = _public_state(state)
    if state.get("status") in ("running", "queued") and not _pid_alive(state.get("pid")):
        public["status"] = "interrupted"
        public["message"] = "The job process is no longer running; call render again with resume=true."
    log_path = directory / "job.log"
    tail = log_path.read_text(encoding="utf-8").splitlines()[-log_lines:] if log_path.is_file() else []
    if public["status"] == "interrupted":
        public.setdefault("message", "The job was interrupted; call render again with resume=true to continue.")
    result = {"ok": public["status"] not in ("failed", "interrupted"), "action": "status", **public,
              "log_tail": tail, "result": state.get("result")}
    if cleanup and state.get("status") == "completed":
        shutil.rmtree(directory, ignore_errors=True)
        result["cleaned_up"] = True
    return result


def start_background(input_path: str, output_path: str, manifest_path: str, batch_size: int,
                     resume: bool, overwrite: bool, allow_unresolved_markers: bool) -> dict:
    """Start ``render`` as a detached process and return immediately."""
    directory = job_dir(output_path)
    existing = directory / "state.json"
    if existing.is_file():
        state = json.loads(existing.read_text(encoding="utf-8"))
        if state.get("status") in ("running", "queued") and _pid_alive(state.get("pid")):
            return {"ok": False, "action": "render", "error": "A job for this output path is already running.",
                    "job": _public_state(state)}
    directory.mkdir(parents=True, exist_ok=True)
    command = [sys.executable, str(Path(__file__).resolve()), "render", "--input", input_path,
               "--output", output_path, "--manifest", manifest_path, "--batch-size", str(batch_size)]
    if not resume:
        command.append("--no-resume")
    if overwrite:
        command.append("--overwrite")
    if allow_unresolved_markers:
        command.append("--allow-unresolved-markers")
    flags = 0
    if os.name == "nt":
        flags = (getattr(subprocess, "DETACHED_PROCESS", 0x8) | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0x200)
                 | getattr(subprocess, "CREATE_NO_WINDOW", 0x8000000))
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    with (directory / "background.out").open("ab") as sink:
        process = subprocess.Popen(command, stdin=subprocess.DEVNULL, stdout=sink, stderr=sink,
                                   creationflags=flags, env=env, close_fds=True)
    # Write a queued state at once so a status call made right after this returns sees the job.
    queued = {"schema": STATE_SCHEMA, "status": "queued", "pid": process.pid, "output": str(Path(output_path).resolve()),
              "input": str(Path(input_path).resolve()), "started": _now(), "updated": _now(), "chunks": [],
              "equations_total": None, "equations_done": 0, "completed_chunks": 0, "retries": 0, "error": None}
    state_path = directory / "state.json"
    if not state_path.is_file():
        temporary = state_path.with_suffix(".tmp")
        temporary.write_text(json.dumps(queued, ensure_ascii=False, indent=1), encoding="utf-8")
        os.replace(temporary, state_path)
    return {"ok": True, "action": "render", "background": True, "pid": process.pid,
            "output_path": str(Path(output_path).resolve()), "job_dir": str(directory),
            "message": "Rendering in the background. Poll get_mathtype_render_status with the same output_path; "
                       "do not run another Word/MathType job until it finishes."}


def _utf8_stdio() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass


def main() -> int:
    _utf8_stdio()
    parser = argparse.ArgumentParser()
    sub = parser.add_subparsers(dest="cmd", required=True)
    r = sub.add_parser("render")
    r.add_argument("--input", required=True)
    r.add_argument("--output", required=True)
    r.add_argument("--manifest", required=True)
    r.add_argument("--batch-size", type=int, default=DEFAULT_BATCH_SIZE)
    r.add_argument("--no-resume", action="store_true")
    r.add_argument("--overwrite", action="store_true")
    r.add_argument("--allow-unresolved-markers", action="store_true")
    s = sub.add_parser("status")
    s.add_argument("--output", required=True)
    s.add_argument("--cleanup", action="store_true")
    args = parser.parse_args()
    try:
        if args.cmd == "render":
            result = render(args.input, args.output, args.manifest, args.batch_size, not args.no_resume,
                            args.overwrite, args.allow_unresolved_markers)
        else:
            result = status(args.output, args.cleanup)
    except Exception as exc:
        result = {"ok": False, "action": args.cmd, "error": f"{type(exc).__name__}: {exc}"}
        if args.cmd == "render":
            try:
                job = Job(args.input, args.output, args.manifest, args.batch_size)
                if job.state_path.is_file():
                    job.state = json.loads(job.state_path.read_text(encoding="utf-8"))
                    job.save(status="failed", error=result["error"])
            except Exception:
                pass
    sys.stdout.write(json.dumps(result, ensure_ascii=False) + "\n")
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(main())
