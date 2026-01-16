
# -*- coding: utf-8 -*-
"""
Orquestrador sequencial para robôs Selenium (Firefox).

Recursos:
- Aceita --manifest <caminho para orchestrator.json>
- Se não achar manifesto, pode usar --auto-discover (descobre scripts em Federal/ e Municipal/)
- Timeout por etapa, retries com backoff, logs por step, log geral
- Checkpoint (resume) e resumo em JSON/CSV
- Força --headless em todos com --headless-all
- (Novo) Controle de headless por step (campo opcional "headless": true|false)
- (Novo) Habilitar/desabilitar step via "enabled": true|false
- (Novo) Leitura da saída dos steps em UTF-8 (evita UnicodeDecodeError)

Uso:
  python 0_ORCHESTRATOR.py
  python 0_ORCHESTRATOR.py --manifest "C:\\Projeto\\orchestrator.json"
  python 0_ORCHESTRATOR.py --resume
  python 0_ORCHESTRATOR.py --only "CT-e"
  python 0_ORCHESTRATOR.py --dry-run
  python 0_ORCHESTRATOR.py --headless-all
  python 0_ORCHESTRATOR.py --auto-discover
"""

import argparse
import csv
import datetime as dt
import json
import os
import sys
import time
from pathlib import Path
from typing import Dict, Any, List, Optional
import subprocess
import re
import unicodedata

# -------------------- Locais --------------------
ROOT = Path(__file__).resolve().parent
LOG_DIR = ROOT / "logs"
STEP_LOG_DIR = LOG_DIR / "steps"
CHECKPOINT_FILE = ROOT / "checkpoint.json"

# -------------------- Util --------------------
def ts_now() -> str:
    return dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def ensure_dirs() -> None:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    STEP_LOG_DIR.mkdir(parents=True, exist_ok=True)

def python_exe() -> str:
    return sys.executable

def write_json(p: Path, data: Any) -> None:
    p.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

def read_checkpoint() -> Dict[str, Any]:
    if not CHECKPOINT_FILE.exists():
        return {"last_success_index": -1, "history": []}
    with CHECKPOINT_FILE.open("r", encoding="utf-8") as f:
        return json.load(f)

def update_checkpoint(idx: int, step_name: str, status: str) -> None:
    ck = read_checkpoint()
    if status.lower() == "success":
        ck["last_success_index"] = idx
    ck.setdefault("history", []).append(
        {"when": ts_now(), "index": idx, "step": step_name, "status": status}
    )
    write_json(CHECKPOINT_FILE, ck)

def load_env_file(path: Path) -> None:
    """
    Carrega arquivo .env simples (linhas key=value).
    Não sobrescreve variáveis já exportadas no ambiente.
    """
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, v = line.split("=", 1)
        k, v = k.strip(), v.strip().strip('"').strip("'")
        os.environ.setdefault(k, v)

def format_timedelta(seconds: float) -> str:
    return str(dt.timedelta(seconds=int(seconds)))

# --------- Safe filename para Windows (logs) ---------
_INVALID_CHARS = r'[<>:"/\\|?*]+'
_RESERVED = {
    "CON", "PRN", "AUX", "NUL",
    *[f"COM{i}" for i in range(1, 10)],
    *[f"LPT{i}" for i in range(1, 10)]
}

def safe_log_name(name: str, maxlen: int = 140) -> str:
    """
    Gera um nome seguro para arquivo de log:
    - remove acentuação (ASCII)
    - substitui caracteres inválidos por "_"
    - colapsa espaços em "_"
    - evita nomes reservados no Windows
    - limita comprimento
    """
    s = unicodedata.normalize('NFKD', name).encode('ascii', 'ignore').decode('ascii')
    s = re.sub(_INVALID_CHARS, "_", s)          # troca /, \, etc. por "_"
    s = re.sub(r"\s+", "_", s).strip("._")
    if s.upper() in _RESERVED or s == "":
        s = f"_{s}" if s else "step"
    return s[:maxlen]

# -------------------- Manifesto --------------------
def find_manifest(manifest_arg: Optional[str]) -> Optional[Path]:
    """
    Retorna o caminho do manifesto se existir.
    Ordem de busca:
      1) --manifest <arg>
      2) ROOT/orchestrator.json
      3) ROOT.parent/orchestrator.json
      4) ROOT.parent.parent/orchestrator.json
    """
    if manifest_arg:
        p = Path(manifest_arg).expanduser().resolve()
        return p if p.exists() else None

    candidates = [
        ROOT / "orchestrator.json",
        ROOT.parent / "orchestrator.json",
        ROOT.parent.parent / "orchestrator.json",
    ]
    for c in candidates:
        if c.exists():
            return c
    return None

def read_manifest(path: Path) -> Dict[str, Any]:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)

def auto_discover_manifest() -> Dict[str, Any]:
    """
    Cria um manifesto em memória descobrindo scripts .py nas pastas Federal/ e Municipal/
    (apenas nível atual, sem varrer recursivamente).
    """
    steps: List[Dict[str, Any]] = []
    default_timeout_minutes = 40
    default_retries = 1

    def add_dir(dir_name: str):
        base = ROOT / dir_name
        if not base.exists():
            return
        for py in sorted(base.glob("*.py")):
            # ignora o próprio orquestrador (nomenclaturas pt/en)
            low = py.name.lower()
            if low.startswith("0_orquestrador") or low.startswith("0_orchestrator"):
                continue
            steps.append({
                "name": f"{dir_name}: {py.stem}",
                "path": str(py.relative_to(ROOT)),
                "args": ["--headless"],  # padrão headless; ajuste se quiser diferente
                "timeout_minutes": default_timeout_minutes,
                "retries": default_retries,
                # Heurística: se contiver "IOB", permite continuar
                "continue_on_error": True if "IOB" in py.stem.upper() else False
            })

    add_dir("Federal")
    add_dir("Municipal")

    return {
        "sleep_between_steps_seconds": 8,
        "default_timeout_minutes": default_timeout_minutes,
        "default_retries": default_retries,
        "env_files": [".env", ".ENV"],
        "steps": steps
    }

def find_or_build_manifest(args) -> Dict[str, Any]:
    mf_path = find_manifest(args.manifest)
    if args.auto_discover and mf_path is None:
        print("[info] Nenhum manifesto encontrado. Usando auto-discover.")
        return auto_discover_manifest()
    if mf_path is None:
        raise FileNotFoundError(
            "Manifesto não encontrado. Informe com --manifest <caminho> "
            "ou use --auto-discover para executar sem manifesto."
        )
    print(f"[info] Manifesto carregado: {mf_path}")
    return read_manifest(mf_path)

# -------------------- Execução de step --------------------
def run_step(
    idx: int,
    step: Dict[str, Any],
    headless_all: bool = False,
    default_timeout_min: int = 40,
    default_retries: int = 1,
) -> Dict[str, Any]:
    """
    Executa um step via subprocess, com stdout/stderr redirecionados a arquivo,
    timeout por etapa e até N retries.
    """
    name = step.get("name", f"step_{idx}")
    rel_path = step["path"]
    script_path = (ROOT / rel_path).resolve()
    if not script_path.exists():
        raise FileNotFoundError(f"Script não encontrado: {script_path}")

    # ---------- Args e controle de headless por step ----------
    args: List[str] = list(step.get("args", []))
    args = [a.strip() for a in args if a and a.strip()]  # normaliza

    # Campo opcional "headless": true/false
    # - True  -> garante --headless
    # - False -> remove --headless e ignora --headless-all
    # - None  -> mantém comportamento padrão (args + --headless-all)
    step_headless = step.get("headless", None)

    def add_headless(a: List[str]) -> List[str]:
        return a if "--headless" in a else a + ["--headless"]

    def drop_headless(a: List[str]) -> List[str]:
        return [x for x in a if x != "--headless"]

    apply_headless_all = headless_all
    if step_headless is True:
        args = add_headless(args)
    elif step_headless is False:
        args = drop_headless(args)
        apply_headless_all = False  # não forçar headless neste step

    if apply_headless_all:
        args = add_headless(args)
    # ----------------------------------------------------------

    timeout_min = int(step.get("timeout_minutes", default_timeout_min))
    retries = int(step.get("retries", default_retries))
    continue_on_error = bool(step.get("continue_on_error", False))

    # Nome de arquivo de log seguro
    safe_name = safe_log_name(name)
    step_log = STEP_LOG_DIR / f"{idx:02d}_{safe_name}.log"
    step_log.parent.mkdir(parents=True, exist_ok=True)

    status = "failed"
    attempts = 0
    started_at = time.time()
    err_msg: Optional[str] = None

    while attempts <= retries:
        attempts += 1
        attempt_tag = f"(tentativa {attempts}/{retries+1})" if retries > 0 else ""
        print(f"[{ts_now()}] Iniciando step {idx}: {name} {attempt_tag}")
        with step_log.open("a", encoding="utf-8", newline="\n") as lf:
            lf.write(f"=== {ts_now()} :: {name} {attempt_tag} ===\n")
            cmd = [python_exe(), str(script_path), *args]
            lf.write(f"$ {' '.join(cmd)}\n\n")
            lf.flush()

            # Processo (UTF-8 garantido)
            env_child = os.environ.copy()
            env_child.setdefault("PYTHONIOENCODING", "utf-8")  # filho escreve UTF-8

            proc = subprocess.Popen(
                cmd,
                cwd=str(script_path.parent),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,                 # lê como texto
                encoding="utf-8",          # decodifica stdout como UTF-8
                errors="replace",          # evita UnicodeDecodeError
                bufsize=1,                 # line-buffered
                universal_newlines=True,   # \n universal
                env=env_child,
            )

            # Loop de leitura com timeout incremental
            deadline = time.time() + timeout_min * 60
            try:
                while True:
                    line = proc.stdout.readline()
                    if line:
                        lf.write(line)
                    else:
                        if proc.poll() is not None:
                            break
                        time.sleep(0.2)

                    # Timeout
                    if time.time() > deadline:
                        proc.kill()
                        raise TimeoutError(
                            f"Timeout de {timeout_min} min atingido no step '{name}'."
                        )

                rc = proc.returncode
                lf.write(f"\n[RC={rc}] Finalizado {name}\n")
                status = "success" if rc == 0 else "failed"
                err_msg = None if rc == 0 else f"Exit code {rc}"
            except TimeoutError as te:
                lf.write(f"\n[TIMEOUT] {str(te)}\n")
                status = "timeout"
                err_msg = str(te)
            except Exception as e:
                lf.write(f"\n[EXCEPTION] {repr(e)}\n")
                status = "failed"
                err_msg = repr(e)

        if status == "success":
            break
        else:
            if attempts <= retries:
                sleep_backoff = min(60, 10 * attempts)  # 10s, 20s, 30s... até 60s
                print(f"[{ts_now()}] Step '{name}' falhou ({status}). "
                      f"Nova tentativa em {sleep_backoff}s...")
                time.sleep(sleep_backoff)

    duration = time.time() - started_at
    result = {
        "index": idx,
        "name": name,
        "script": str(script_path),
        "args": args,
        "status": status,
        "attempts": attempts,
        "continue_on_error": continue_on_error,
        "duration_sec": int(duration),
        "duration_hms": format_timedelta(duration),
        "log_file": str(step_log),
        "error": err_msg,
        "finished_at": ts_now(),
    }
    return result

# -------------------- Main --------------------
def main():
    ap = argparse.ArgumentParser(description="Orquestrador de robôs Selenium (sequencial).")
    ap.add_argument("--manifest", type=str, help="Caminho para orchestrator.json.")
    ap.add_argument("--resume", action="store_true", help="Retoma a partir do próximo step após o último OK.")
    ap.add_argument("--only", type=str, default="", help="Executa apenas steps cujo nome contém este texto.")
    ap.add_argument("--dry-run", action="store_true", help="Mostra o plano de execução sem rodar.")
    ap.add_argument("--headless-all", action="store_true", help="Força --headless em todos os steps (se compatível).")
    ap.add_argument("--auto-discover", action="store_true",
                    help="Executa sem manifesto: descobre scripts em Federal/ e Municipal/.")
    args = ap.parse_args()

    ensure_dirs()

    # Manifesto
    manifest: Dict[str, Any] = find_or_build_manifest(args)

    # Carrega .env(s)
    for env_file in manifest.get("env_files", []):
        env_path = (ROOT / env_file).resolve()
        load_env_file(env_path)

    steps: List[Dict[str, Any]] = manifest.get("steps", [])
    default_timeout_min = int(manifest.get("default_timeout_minutes", 40))
    default_retries = int(manifest.get("default_retries", 1))
    sleep_between = int(manifest.get("sleep_between_steps_seconds", 8))

    # Ignora steps desativados (enabled=false)
    steps = [s for s in steps if s.get("enabled", True)]

    # Filtro --only (pelo "name")
    if args.only:
        needle = args.only.lower().strip()
        steps = [s for s in steps if needle in s.get("name", "").lower()]

    if not steps:
        print("Não há steps para executar após filtragem.")
        sys.exit(2)

    # Resume
    start_index = 0
    if args.resume:
        ck = read_checkpoint()
        start_index = int(ck.get("last_success_index", -1)) + 1
        if start_index >= len(steps):
            print("Nada a retomar: todos os steps já concluídos. Reinicie sem --resume se desejar reprocessar.")
            sys.exit(0)

    run_id = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    run_log = LOG_DIR / f"run_{run_id}.log"
    summary_json = LOG_DIR / f"summary_{run_id}.json"
    summary_csv = LOG_DIR / f"summary_{run_id}.csv"

    with run_log.open("a", encoding="utf-8") as rf:
        rf.write(f"==== RUN {run_id} | {ts_now()} ====\n")
        rf.write(f"Python: {sys.version.split()[0]} | Exec: {python_exe()}\n")
        rf.write(f"Root: {ROOT}\n\n")

    results: List[Dict[str, Any]] = []
    exit_code = 0

    # Dry-run
    if args.dry_run:
        print("Plano de execução:")
        for i, step in enumerate(steps, start=0):
            print(f"- [{i}] {step.get('name')} -> {step.get('path')} | "
                  f"args={step.get('args', [])} | headless={step.get('headless', 'default')}")
        sys.exit(0)

    # Execução
    for idx, step in enumerate(steps, start=0):
        if idx < start_index:
            continue

        res = run_step(
            idx=idx,
            step=step,
            headless_all=args.headless_all,
            default_timeout_min=default_timeout_min,
            default_retries=default_retries,
        )
        results.append(res)

        update_checkpoint(idx, res["name"], "success" if res["status"] == "success" else "failed")

        with run_log.open("a", encoding="utf-8") as rf:
            rf.write(
                f"[{ts_now()}] Step {idx} - {res['name']}: {res['status'].upper()} "
                f"({res['duration_hms']} | attempts={res['attempts']})\n"
                f"Log: {res['log_file']}\n"
            )
            if res["error"]:
                rf.write(f"Erro: {res['error']}\n")
            rf.write("\n")

        if res["status"] != "success" and not res["continue_on_error"]:
            print(f"[{ts_now()}] ORQUESTRAÇÃO INTERROMPIDA: step '{res['name']}' falhou ({res['status']}).")
            exit_code = 1
            break

        time.sleep(sleep_between)

    write_json(summary_json, {
        "run_id": run_id,
        "started_at": ts_now(),
        "root": str(ROOT),
        "results": results,
    })

    with (summary_csv).open("w", encoding="utf-8", newline="") as cf:
        w = csv.writer(cf, delimiter=";")
        w.writerow(["index", "name", "status", "attempts", "duration_hms", "log_file", "error"])
        for r in results:
            w.writerow([r["index"], r["name"], r["status"], r["attempts"],
                        r["duration_hms"], r["log_file"], r["error"] or ""])

    print(f"Resumo salvo em:\n- {summary_json}\n- {summary_csv}")
    print(f"Log geral: {run_log}")
    sys.exit(exit_code)

if __name__ == "__main__":
    main()
