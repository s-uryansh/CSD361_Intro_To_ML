#!/usr/bin/env python3
"""
strip_emdash.py — removes every em dash from report.html, replacing each with
punctuation that suits its grammatical role (comma, colon, full stop, parenthesis
or plain hyphen). Also normalises en dashes used as compound modifiers to hyphens,
while keeping en dashes where typography requires them (numeric ranges, page
ranges, section ranges, the name Runge–Kutta).

Each replacement must match exactly once; the script aborts otherwise so that no
em dash can survive silently.
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
PATH = os.path.join(HERE, "report.html")

# (old, new) — ordered, each must occur exactly once
REPLACEMENTS = [
    # ---- front matter -----------------------------------------------------
    ("<title>Struct-X &mdash; CSD361 Preliminary Project Report</title>",
     "<title>Struct-X: CSD361 Preliminary Project Report</title>"),

    ("<div class=\"course\">CSD361 &mdash; Introduction to Machine Learning",
     "<div class=\"course\">CSD361: Introduction to Machine Learning"),

    # IEEE writes "Abstract—"; use a full stop instead.
    ("<b>Abstract</b>&mdash;&mdash;Non-contact",
     "<b>Abstract.</b> Non-contact"),

    ("<b>Index Terms</b>&mdash;&mdash;event cameras",
     "<b>Index Terms.</b> event cameras"),

    # ---- abstract ---------------------------------------------------------
    ("currents &mdash; so the same physical motion",
     "currents, so the same physical motion"),

    # ---- I. Introduction --------------------------------------------------
    ("clear line of sight and a head held still while the target moves &mdash; workable on a laboratory bench, much",
     "clear line of sight and a head held still while the target moves: workable on a laboratory bench, much"),

    ("For a beam oscillating at tens of hertz to a few kilohertz &mdash; where a conventional camera either blurs the\n"
     "motion or demands lighting and bandwidth a field deployment cannot carry &mdash; this is not a marginal\n"
     "advantage.",
     "For a beam oscillating at tens of hertz to a few kilohertz, a conventional camera either blurs the motion or\n"
     "demands lighting and bandwidth a field deployment cannot carry. This is not a marginal advantage."),

    ("user rather than by the scene &mdash; this decides whether sensor-induced domain shift is real or theoretical.",
     "user rather than by the scene. This decides whether sensor-induced domain shift is real or theoretical."),

    # ---- II. Research Problem --------------------------------------------
    ("<li><b>P1 &mdash; Accuracy.</b>", "<li><b>P1. Accuracy.</b>"),
    ("<li><b>P2 &mdash; Stability.</b>", "<li><b>P2. Stability.</b>"),
    ("<li><b>P3 &mdash; Feasibility.</b>", "<li><b>P3. Feasibility.</b>"),

    # ---- III. Objectives --------------------------------------------------
    ("how the literature measures &mdash; or fails to measure &mdash; degradation under changed",
     "how the literature measures, or fails to measure, degradation under changed"),

    # ---- V. Review --------------------------------------------------------
    ("<h3>A. Approach 1 &mdash; Geometric, largely training-free event&ndash;frame deformation measurement</h3>",
     "<h3>A. Approach 1: Geometric, largely training-free event-frame deformation measurement</h3>"),

    ("100&nbsp;px displacement &mdash; 1.6&times; the previous state of the art &mdash; using 18.9% of the storage and",
     "100&nbsp;px displacement (1.6&times; the previous state of the art) using 18.9% of the storage and"),

    ("continuity and brightness constancy &mdash; with pseudo-strains under",
     "continuity and brightness constancy, with pseudo-strains under"),

    ("<h3>B. Approach 2 &mdash; Physics-informed and physics-guided learning for structural dynamics</h3>",
     "<h3>B. Approach 2: Physics-informed and physics-guided learning for structural dynamics</h3>"),

    ("generalise it and document where it fails &mdash;\n"
     "training can fail for reasons unconnected to the physics",
     "generalise it and document where it fails:\n"
     "training can fail for reasons unconnected to the physics"),

    ("a physics-based decoder &mdash;\n"
     "an updated finite element model solved by Newmark-&beta; &mdash; reconstructs the response from those states.",
     "a physics-based decoder\n"
     "(an updated finite element model solved by Newmark-&beta;) reconstructs the response from those states."),

    ("That is structurally the same problem Struct-X faces\n"
     "&mdash; fuse a non-contact visual measurement with a physical response and let the equation of motion regularise\n"
     "the fit.",
     "That is structurally the same problem Struct-X faces,\n"
     "fusing a non-contact visual measurement with a physical response and letting the equation of motion regularise\n"
     "the fit."),

    ("<h3>C. Approach 3 &mdash; Spatiotemporal Transformers and robustness-aware SHM</h3>",
     "<h3>C. Approach 3: Spatiotemporal Transformers and robustness-aware SHM</h3>"),

    ("&mdash; local connectivity and translation equivariance &mdash; as genuinely <em>beneficial on small training",
     "(local connectivity and translation equivariance) as genuinely <em>beneficial on small training"),

    # ---- VI. Comparative --------------------------------------------------
    ("and only\none needs no data at all &mdash; at the cost of a co-registered frame stream and an assumption of brightness",
     "and only\none needs no data at all, at the cost of a co-registered frame stream and an assumption of brightness"),

    ("<th style=\"width:28%\">A. Geometric event&ndash;frame VDM",
     "<th style=\"width:28%\">A. Geometric event-frame VDM"),

    ("<th style=\"width:27%\">C. Spatiotemporal Transformer + robust SHM",
     "<th style=\"width:27%\">C. Spatiotemporal Transformer + robust SHM"),

    ("<td class=\"c\">Yes, by design [8] &mdash; but categorical targets</td>",
     "<td class=\"c\">Yes, by design [8], but categorical targets</td>"),

    # ---- VII. Critical ----------------------------------------------------
    ("returns a field by per-sequence optimisation &mdash; no learned prior to amortise across 30",
     "returns a field by per-sequence optimisation: no learned prior to amortise across 30"),

    ("<span class=\"m\">k</span> identified from a single training trial &mdash; which",
     "<span class=\"m\">k</span> identified from a single training trial, which"),

    ("the equivalent notion &mdash; how far the reconstructed trace\n"
     "and recovered frequency move &mdash; has no accepted definition here.",
     "the equivalent notion, namely how far the reconstructed trace\n"
     "and recovered frequency move, has no accepted definition here."),

    # ---- VIII. Gaps -------------------------------------------------------
    ("<li><b>G1 &mdash; No event-vibration pipeline exists.</b>",
     "<li><b>G1. No event-vibration pipeline exists.</b>"),
    ("<li><b>G2 &mdash; Physics and events have never been combined.</b>",
     "<li><b>G2. Physics and events have never been combined.</b>"),
    ("<li><b>G3 &mdash; Sensing-condition shift is unmeasured.</b>",
     "<li><b>G3. Sensing-condition shift is unmeasured.</b>"),
    ("<li><b>G4 &mdash; No waveform-domain definition of robustness.</b>",
     "<li><b>G4. No waveform-domain definition of robustness.</b>"),
    ("<li><b>G5 &mdash; Physics priors are unidentified under scarce data.</b>",
     "<li><b>G5. Physics priors are unidentified under scarce data.</b>"),

    # ---- IX. Direction ----------------------------------------------------
    ("<span class=\"t2\">E = {(x, y, t, p)} &mdash; condition s</span>",
     "<span class=\"t2\">E = {(x, y, t, p)}, condition s</span>"),

    ("<span class=\"t2\">divided space&ndash;time attention; pretrained or lightweight</span>",
     "<span class=\"t2\">divided space-time attention; pretrained or lightweight</span>"),

    ("the robustness gap &Delta; of P2 &mdash; all as mean &plusmn; standard",
     "the robustness gap &Delta; of P2, all as mean &plusmn; standard"),

    # ---- X. RQs -----------------------------------------------------------
    ("Which constraint carries the benefit under scarcity &mdash; the SDOF equation\n"
     "  residual, a temporal smoothness penalty on <span class=\"m\">q&#770;</span>, or their combination &mdash; and does",
     "Which constraint carries the benefit under scarcity (the SDOF equation\n"
     "  residual, a temporal smoothness penalty on <span class=\"m\">q&#770;</span>, or their combination) and does"),
]

with open(PATH, encoding="utf-8") as f:
    text = f.read()

problems = []
for old, new in REPLACEMENTS:
    n = text.count(old)
    if n != 1:
        problems.append(f"expected 1 match, found {n}: {old[:80]!r}")
        continue
    text = text.replace(old, new)

if problems:
    print("FAILED:", file=sys.stderr)
    for p in problems:
        print("  " + p, file=sys.stderr)
    sys.exit(1)

with open(PATH, "w", encoding="utf-8") as f:
    f.write(text)

left = text.count("&mdash;") + text.count("\u2014")
print(f"Applied {len(REPLACEMENTS)} replacements.")
print(f"Remaining em dashes: {left}")
if left:
    sys.exit(1)
