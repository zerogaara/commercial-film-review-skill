#!/usr/bin/env python3
"""
Commercial Film Quantitative Analyzer (FFmpeg + ffprobe + numpy + Pillow)
Extracts objective editing, audio loudness, pacing, and color consistency metrics
for brand films and product launch commercials.
"""

import argparse
import json
import math
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageFont


def run_cmd(cmd):
    proc = subprocess.run(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    return proc.returncode, proc.stdout, proc.stderr


def fmt_tc(seconds: float) -> str:
    if seconds < 0:
        seconds = 0.0
    m = int(seconds // 60)
    s = seconds - m * 60
    return f"{m:02d}:{s:05.2f}"


def probe_metadata(video_path: Path) -> dict:
    cmd = [
        "ffprobe",
        "-v",
        "quiet",
        "-print_format",
        "json",
        "-show_format",
        "-show_streams",
        str(video_path),
    ]
    code, out, err = run_cmd(cmd)
    if code != 0:
        raise RuntimeError(f"ffprobe failed: {err}")
    data = json.loads(out)
    v_stream = next((s for s in data.get("streams", []) if s.get("codec_type") == "video"), None)
    a_stream = next((s for s in data.get("streams", []) if s.get("codec_type") == "audio"), None)
    fmt = data.get("format", {})

    fps = 25.0
    if v_stream and v_stream.get("avg_frame_rate") and "/" in v_stream["avg_frame_rate"]:
        num, den = v_stream["avg_frame_rate"].split("/")
        if float(den) > 0:
            fps = round(float(num) / float(den), 3)

    duration = float(fmt.get("duration") or (v_stream.get("duration") if v_stream else 0) or 0)
    return {
        "file": str(video_path),
        "duration_sec": round(duration, 3),
        "duration_tc": fmt_tc(duration),
        "width": int(v_stream.get("width", 0)) if v_stream else 0,
        "height": int(v_stream.get("height", 0)) if v_stream else 0,
        "fps": fps,
        "video_codec": v_stream.get("codec_name") if v_stream else None,
        "pix_fmt": v_stream.get("pix_fmt") if v_stream else None,
        "bitrate_kbps": round(float(fmt.get("bit_rate", 0)) / 1000.0, 1) if fmt.get("bit_rate") else None,
        "has_audio": a_stream is not None,
        "audio_codec": a_stream.get("codec_name") if a_stream else None,
        "audio_sample_rate": int(a_stream.get("sample_rate", 0)) if a_stream and a_stream.get("sample_rate") else None,
        "audio_channels": int(a_stream.get("channels", 0)) if a_stream and a_stream.get("channels") else None,
    }


def detect_cuts(video_path: Path, threshold: float, duration: float) -> list[float]:
    cmd = [
        "ffmpeg",
        "-hide_banner",
        "-i",
        str(video_path),
        "-vf",
        f"select='gt(scene,{threshold})',showinfo",
        "-f",
        "null",
        "-",
    ]
    _, _, err = run_cmd(cmd)
    cuts = []
    for line in err.splitlines():
        m = re.search(r"pts_time:([0-9]+\.?[0-9]*)", line)
        if m:
            t = float(m.group(1))
            if 0.15 < t < max(0.2, duration - 0.15):
                if not cuts or (t - cuts[-1]) >= 0.25:
                    cuts.append(round(t, 3))
    return cuts


def measure_loudness(video_path: Path) -> dict:
    cmd = [
        "ffmpeg",
        "-hide_banner",
        "-i",
        str(video_path),
        "-af",
        "loudnorm=I=-16:TP=-1.5:LRA=11:print_format=json",
        "-f",
        "null",
        "-",
    ]
    _, _, err = run_cmd(cmd)
    m = re.search(r"\{[^{}]*\"input_i\"[^{}]*\}", err, re.S)
    if not m:
        return {}
    try:
        raw = json.loads(m.group(0))
        return {
            "integrated_lufs": float(raw["input_i"]) if raw.get("input_i") != "-inf" else None,
            "true_peak_dbtp": float(raw["input_tp"]) if raw.get("input_tp") != "-inf" else None,
            "loudness_range_lu": float(raw["input_lra"]) if raw.get("input_lra") != "-inf" else None,
        }
    except Exception:
        return {}


def compute_audio_envelope(video_path: Path, duration: float) -> list[dict]:
    cmd = [
        "ffmpeg",
        "-hide_banner",
        "-i",
        str(video_path),
        "-ac",
        "1",
        "-ar",
        "8000",
        "-f",
        "s16le",
        "-",
    ]
    proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if proc.returncode != 0 or not proc.stdout:
        return []
    samples = np.frombuffer(proc.stdout, dtype=np.int16).astype(np.float32) / 32768.0
    sr = 8000
    step = sr  # 1 second bins
    env = []
    for idx in range(0, len(samples), step):
        chunk = samples[idx : idx + step]
        if len(chunk) == 0:
            continue
        rms = float(np.sqrt(np.mean(chunk * chunk) + 1e-12))
        db = round(20.0 * math.log10(max(rms, 1e-6)), 1)
        sec = idx // sr
        env.append({"sec": sec, "tc": fmt_tc(sec), "rms_db": db})
    return env


def rgb_to_hsv_stats(arr: np.ndarray) -> dict:
    # arr: H x W x 3 uint8
    rgb = arr.astype(np.float32) / 255.0
    r, g, b = rgb[:, :, 0], rgb[:, :, 1], rgb[:, :, 2]
    luma = 0.2126 * r + 0.7152 * g + 0.0722 * b
    cmax = np.max(rgb, axis=2)
    cmin = np.min(rgb, axis=2)
    delta = cmax - cmin
    sat = np.where(cmax > 1e-5, delta / (cmax + 1e-5), 0.0)
    mean_rgb = [int(round(float(np.mean(arr[:, :, i])))) for i in range(3)]
    clipped_hi = float(np.mean(luma > 0.96)) * 100.0
    crushed_lo = float(np.mean(luma < 0.04)) * 100.0
    return {
        "mean_rgb": mean_rgb,
        "mean_hex": "#{:02X}{:02X}{:02X}".format(*mean_rgb),
        "mean_luma": round(float(np.mean(luma)), 3),
        "luma_std_contrast": round(float(np.std(luma)), 3),
        "mean_saturation": round(float(np.mean(sat)), 3),
        "clipped_highlight_pct": round(clipped_hi, 2),
        "crushed_shadow_pct": round(crushed_lo, 2),
    }


def extract_shot_frames(video_path: Path, shots: list[dict], frames_dir: Path, max_shots_for_sheet: int = 36):
    frames_dir.mkdir(parents=True, exist_ok=True)
    # If too many shots, sample evenly for the contact sheet while keeping all midpoints for color stats up to 80
    sample_indices = list(range(len(shots)))
    if len(shots) > 80:
        step = len(shots) / 80.0
        sample_indices = sorted(set(int(i * step) for i in range(80)))

    extracted = {}
    for idx in sample_indices:
        shot = shots[idx]
        mid = shot["mid_sec"]
        out_jpg = frames_dir / f"shot_{idx+1:03d}.jpg"
        cmd = [
            "ffmpeg",
            "-hide_banner",
            "-y",
            "-ss",
            f"{mid:.3f}",
            "-i",
            str(video_path),
            "-frames:v",
            "1",
            "-vf",
            "scale=480:-2",
            "-q:v",
            "3",
            str(out_jpg),
        ]
        code, _, _ = run_cmd(cmd)
        if code == 0 and out_jpg.exists():
            extracted[idx] = out_jpg

    # Compute color stats per extracted shot and detect jumps
    color_jumps = []
    prev_stats = None
    prev_idx = None
    barcode_cols = []

    for idx in sample_indices:
        img_path = extracted.get(idx)
        if not img_path:
            continue
        im = Image.open(img_path).convert("RGB")
        arr = np.array(im)
        stats = rgb_to_hsv_stats(arr)
        shots[idx]["color"] = stats

        # Column for color barcode (average each row or mean color)
        col = np.mean(arr, axis=1).astype(np.uint8)  # H x 3
        col_img = Image.fromarray(col[:, None, :]).resize((12, 120), Image.Resampling.BILINEAR)
        barcode_cols.append(np.array(col_img))

        if prev_stats is not None:
            luma_diff = abs(stats["mean_luma"] - prev_stats["mean_luma"])
            rgb_diff = math.sqrt(sum((stats["mean_rgb"][c] - prev_stats["mean_rgb"][c]) ** 2 for c in range(3)))
            if luma_diff > 0.22 or rgb_diff > 65:
                color_jumps.append(
                    {
                        "between_shots": [prev_idx + 1, idx + 1],
                        "at_tc": shots[idx]["start_tc"],
                        "luma_delta": round(luma_diff, 3),
                        "rgb_delta": round(rgb_diff, 1),
                        "from_hex": prev_stats["mean_hex"],
                        "to_hex": stats["mean_hex"],
                    }
                )
        prev_stats = stats
        prev_idx = idx

    return extracted, barcode_cols, color_jumps


def build_contact_sheet(extracted: dict, shots: list[dict], out_path: Path, max_tiles: int = 36):
    keys = sorted(extracted.keys())
    if not keys:
        return
    if len(keys) > max_tiles:
        step = len(keys) / float(max_tiles)
        keys = [keys[int(i * step)] for i in range(max_tiles)]

    sample_im = Image.open(extracted[keys[0]]).convert("RGB")
    tw, th = sample_im.size
    cols = 4
    rows = math.ceil(len(keys) / cols)
    pad = 8
    header_h = 24
    sheet = Image.new("RGB", (cols * (tw + pad) + pad, rows * (th + header_h + pad) + pad), (18, 20, 24))
    draw = ImageDraw.Draw(sheet)
    font = ImageFont.load_default()

    for pos, idx in enumerate(keys):
        r = pos // cols
        c = pos % cols
        x = pad + c * (tw + pad)
        y = pad + r * (th + header_h + pad)
        shot = shots[idx]
        im = Image.open(extracted[idx]).convert("RGB").resize((tw, th), Image.Resampling.LANCZOS)
        sheet.paste(im, (x, y + header_h))
        label = f"#{idx+1:02d} {shot['start_tc']} ({shot['duration_sec']:.1f}s)"
        color_hex = shot.get("color", {}).get("mean_hex", "")
        if color_hex:
            label += f" {color_hex}"
        draw.text((x + 4, y + 5), label, fill=(230, 235, 242), font=font)

    sheet.save(out_path, quality=88)


def build_color_barcode(barcode_cols: list[np.ndarray], out_path: Path):
    if not barcode_cols:
        return
    strip = np.concatenate(barcode_cols, axis=1)
    im = Image.fromarray(strip)
    target_w = max(800, min(1920, len(barcode_cols) * 24))
    im = im.resize((target_w, 160), Image.Resampling.BILINEAR)
    im.save(out_path)


def summarize_pacing(shots: list[dict], duration: float) -> dict:
    if not shots:
        return {}
    durs = [s["duration_sec"] for s in shots]
    q1 = duration * 0.25
    q3 = duration * 0.75
    opening = [s["duration_sec"] for s in shots if s["mid_sec"] < q1]
    middle = [s["duration_sec"] for s in shots if q1 <= s["mid_sec"] <= q3]
    ending = [s["duration_sec"] for s in shots if s["mid_sec"] > q3]

    def avg(xs):
        return round(sum(xs) / len(xs), 2) if xs else None

    return {
        "total_shots": len(shots),
        "asl_sec": round(sum(durs) / len(durs), 2),
        "min_shot_sec": round(min(durs), 2),
        "max_shot_sec": round(max(durs), 2),
        "opening_25pct_asl": avg(opening),
        "middle_50pct_asl": avg(middle),
        "ending_25pct_asl": avg(ending),
        "short_shots_under_1s": sum(1 for d in durs if d < 1.0),
        "long_shots_over_5s": sum(1 for d in durs if d > 5.0),
    }


def main():
    parser = argparse.ArgumentParser(description="Analyze commercial/brand film cuts, pacing, loudness, and color.")
    parser.add_argument("video", help="Path to local video file")
    parser.add_argument("--out", help="Output directory (default: film-review-<stem> next to video)")
    parser.add_argument("--scene-threshold", type=float, default=0.30, help="FFmpeg scene cut threshold (default: 0.30)")
    args = parser.parse_args()

    video_path = Path(args.video).resolve()
    if not video_path.exists():
        print(f"Error: video not found: {video_path}", file=sys.stderr)
        sys.exit(1)

    for bin_name in ("ffmpeg", "ffprobe"):
        if not shutil.which(bin_name):
            print(f"Error: {bin_name} not found in PATH", file=sys.stderr)
            sys.exit(1)

    out_dir = Path(args.out).resolve() if args.out else video_path.parent / f"film-review-{video_path.stem}"
    out_dir.mkdir(parents=True, exist_ok=True)

    meta = probe_metadata(video_path)
    duration = meta["duration_sec"]
    cuts = detect_cuts(video_path, args.scene_threshold, duration)

    boundaries = [0.0] + cuts + [duration]
    shots = []
    for i in range(len(boundaries) - 1):
        st = round(boundaries[i], 3)
        en = round(boundaries[i + 1], 3)
        dur = round(max(0.01, en - st), 3)
        mid = round((st + en) / 2.0, 3)
        shots.append(
            {
                "shot_index": i + 1,
                "start_sec": st,
                "end_sec": en,
                "mid_sec": mid,
                "start_tc": fmt_tc(st),
                "end_tc": fmt_tc(en),
                "duration_sec": dur,
            }
        )

    loudness = measure_loudness(video_path) if meta["has_audio"] else {}
    envelope = compute_audio_envelope(video_path, duration) if meta["has_audio"] else []
    extracted, barcode_cols, color_jumps = extract_shot_frames(video_path, shots, out_dir / "frames")

    contact_sheet_path = out_dir / "contact_sheet.jpg"
    barcode_path = out_dir / "color_barcode.png"
    build_contact_sheet(extracted, shots, contact_sheet_path)
    build_color_barcode(barcode_cols, barcode_path)

    report = {
        "metadata": meta,
        "pacing": summarize_pacing(shots, duration),
        "loudness": loudness,
        "color_jump_warnings": color_jumps,
        "shots": shots,
        "audio_envelope_1s": envelope,
        "artifacts": {
            "contact_sheet": str(contact_sheet_path),
            "color_barcode": str(barcode_path),
            "frames_dir": str(out_dir / "frames"),
        },
    }

    metrics_path = out_dir / "metrics.json"
    metrics_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"OK: metrics -> {metrics_path}")
    print(f"OK: contact_sheet -> {contact_sheet_path}")
    print(f"OK: color_barcode -> {barcode_path}")
    print(
        f"Summary: {meta['width']}x{meta['height']} @ {meta['fps']}fps | "
        f"{meta['duration_tc']} | {len(shots)} shots | ASL={report['pacing'].get('asl_sec')}s | "
        f"LUFS={loudness.get('integrated_lufs')} | color_jumps={len(color_jumps)}"
    )


if __name__ == "__main__":
    main()
