#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Jul  2 14:09:49 2025

@author: pouria
"""

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Parallel HEC-1 runner — one isolated folder per .hc1 file.
"""

import shutil
import subprocess
from pathlib import Path
from joblib import Parallel, delayed
from tqdm import tqdm

# === CONFIGURATION ===
output_base_dir = Path("C:/Users/DrBZahraei/Desktop/regenalization/output")
hec_exe_source = Path("C:/Users/DrBZahraei/Desktop/regenalization/HEC1.EXE")
n_jobs = 10  # Number of parallel jobs (adjust to your CPU)

# === Collect all .hc1 files from all 'out/' folders ===
hc1_tasks = []

out_dirs = list(output_base_dir.rglob("*/out"))
print(f"🔍 Scanning {len(out_dirs)} 'out' folders...")

for out_dir in out_dirs:
    hc1_files = list(out_dir.glob("*.hc1"))
    if not hc1_files:
        print(f"⚠️ No .hc1 files found in {out_dir}")
        continue
    for hc1_file in hc1_files:
        hc1_tasks.append((out_dir, hc1_file.name))

print(f"\n📦 Total .hc1 files to process: {len(hc1_tasks)}")

# === Function: process a single .hc1 file ===
def run_hec1(out_dir: Path, hc1_filename: str):
    try:
        exe_path = hec_exe_source.resolve()
        run_folder = out_dir / Path(hc1_filename).stem
        run_folder.mkdir(exist_ok=True)

        src_hc1 = out_dir / hc1_filename
        dst_hc1 = run_folder / hc1_filename
        dst_exe = run_folder / "hec1.exe"

        # print(f"🔧 Copying {src_hc1} and {exe_path} to {run_folder}")
        shutil.copy2(src_hc1, dst_hc1)
        shutil.copy2(exe_path, dst_exe)

        subprocess.run(
                [str(dst_exe), hc1_filename],
                cwd=run_folder,
                check=True,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL
            )


        return f"✅ {hc1_filename} finished"
    except FileNotFoundError as fnf:
        return f"❌ FileNotFoundError: {fnf.filename}"
    except Exception as ex:
        return f"⚠️ Error {hc1_filename} in {out_dir}:\n{ex}"


# === Run in parallel with a progress bar ===
results = Parallel(n_jobs=n_jobs)(
    delayed(run_hec1)(folder, file) for folder, file in tqdm(hc1_tasks, desc="🚀 Running HEC-1")
)

# === Print summary ===
print("\n📋 Summary:")
for res in results:
    print(res)

print("\n✅ All HEC-1 simulations completed.")