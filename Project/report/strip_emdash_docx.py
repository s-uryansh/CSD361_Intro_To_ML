#!/usr/bin/env python3
"""
strip_emdash_docx.py — mirrors strip_emdash.py for the Word build: removes every
em dash from build_docx.py, replacing it with punctuation appropriate to its role.

Matching is whitespace-flexible (a run of spaces/newlines/indentation in the
pattern matches any run of whitespace in the source), and each pattern must match
exactly once or the script aborts so no em dash can survive silently.
"""
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
PATH = os.path.join(HERE, "build_docx.py")
E = "\u2014"  # em dash

PAIRS = [
    # ---- docstring / metadata --------------------------------------------
    ("build_docx.py \u2014 generates", "build_docx.py: generates"),
    ("<dc:subject>CSD361 Preliminary Project Report \u2014 Literature Survey and Research Plan</dc:subject>",
     "<dc:subject>CSD361 Preliminary Project Report: Literature Survey and Research Plan</dc:subject>"),

    # ---- front matter -----------------------------------------------------
    ("**Abstract**\u2014Non-contact", "**Abstract.** Non-contact"),
    ("**Index Terms**\u2014event cameras", "**Index Terms.** event cameras"),
    ("CSD361 — Introduction to Machine Learning", "CSD361: Introduction to Machine Learning"),

    # ---- abstract ---------------------------------------------------------
    ("threshold and bias currents \u2014 so the same physical motion",
     "threshold and bias currents, so the same physical motion"),

    # ---- I. Introduction --------------------------------------------------
    ("head held still while the target moves \u2014 workable on a laboratory",
     "head held still while the target moves: workable on a laboratory"),
    ("tens of hertz to a few kilohertz \u2014 where a conventional camera either blurs the motion or demands lighting and "
     "bandwidth a field deployment cannot carry \u2014 this is not a marginal advantage.",
     "tens of hertz to a few kilohertz, a conventional camera either blurs the motion or demands lighting and "
     "bandwidth a field deployment cannot carry. This is not a marginal advantage."),
    ("rather than by the scene \u2014 this decides whether sensor-induced domain shift is real or theoretical.",
     "rather than by the scene. This decides whether sensor-induced domain shift is real or theoretical."),

    # ---- II. Research Problem --------------------------------------------
    ("**P1 \u2014 Accuracy.**", "**P1. Accuracy.**"),
    ("**P2 \u2014 Stability.**", "**P2. Stability.**"),
    ("**P3 \u2014 Feasibility.**", "**P3. Feasibility.**"),

    # ---- III. Objectives --------------------------------------------------
    ("how the literature measures \u2014 or fails to measure \u2014 degradation under changed sensing or",
     "how the literature measures, or fails to measure, degradation under changed sensing or"),

    # ---- V. Review --------------------------------------------------------
    ("A.  Approach 1 \u2014 Geometric, largely training-free event\u2013frame deformation measurement",
     "A.  Approach 1: Geometric, largely training-free event-frame deformation measurement"),
    ("survival rate beyond 100 px displacement \u2014 1.6\u00d7 the previous state of the art \u2014 using 18.9% of the storage and",
     "survival rate beyond 100 px displacement (1.6\u00d7 the previous state of the art) using 18.9% of the storage and"),
    ("leans on spatial continuity and brightness constancy \u2014 with pseudo-strains under topological change and degraded",
     "leans on spatial continuity and brightness constancy, with pseudo-strains under topological change and degraded"),
    ("B.  Approach 2 \u2014 Physics-informed and physics-guided learning for structural dynamics",
     "B.  Approach 2: Physics-informed and physics-guided learning for structural dynamics"),
    ("document where it fails \u2014", "document where it fails:"),
    ("a physics-based decoder \u2014 an updated finite element model solved by Newmark-\u03b2 \u2014 reconstructs the response from those states.",
     "a physics-based decoder (an updated finite element model solved by Newmark-\u03b2) reconstructs the response from those states."),
    ("Struct-X faces \u2014 fuse a non-contact visual measurement with a physical response and let the equation of motion regularise",
     "Struct-X faces, fusing a non-contact visual measurement with a physical response and letting the equation of motion regularise"),
    ("C.  Approach 3 \u2014 Spatiotemporal Transformers and robustness-aware SHM",
     "C.  Approach 3: Spatiotemporal Transformers and robustness-aware SHM"),
    ("attribute the convolutional bias \u2014 local connectivity and translation equivariance \u2014 as genuinely",
     "attribute the convolutional bias (local connectivity and translation equivariance) as genuinely"),

    # ---- VI. Comparative --------------------------------------------------
    ("only one needs no data at all \u2014 at the cost of a co-registered frame stream and an assumption of brightness",
     "only one needs no data at all, at the cost of a co-registered frame stream and an assumption of brightness"),
    ("A. Geometric event\u2013frame VDM [1], [2], [4]", "A. Geometric event-frame VDM [1], [2], [4]"),

    # ---- VII. Critical ----------------------------------------------------
    ("Yes, by design [8] \u2014 but categorical targets", "Yes, by design [8], but categorical targets"),
    ("returns a field by per-sequence optimisation \u2014 no learned prior to amortise",
     "returns a field by per-sequence optimisation: no learned prior to amortise"),
    ("a single training trial \u2014 which immediately creates the failure mode of Section IX.",
     "a single training trial, which immediately creates the failure mode of Section IX."),
    ("the equivalent notion \u2014 how far the reconstructed trace and recovered frequency move \u2014 has no accepted definition here.",
     "the equivalent notion, namely how far the reconstructed trace and recovered frequency move, has no accepted definition here."),

    # ---- VIII. Gaps -------------------------------------------------------
    ("**G1 \u2014 No event-vibration pipeline exists.**", "**G1. No event-vibration pipeline exists.**"),
    ("**G2 \u2014 Physics and events have never been combined.**", "**G2. Physics and events have never been combined.**"),
    ("**G3 \u2014 Sensing-condition shift is unmeasured.**", "**G3. Sensing-condition shift is unmeasured.**"),
    ("**G4 \u2014 No waveform-domain definition of robustness.**", "**G4. No waveform-domain definition of robustness.**"),
    ("**G5 \u2014 Physics priors are unidentified under scarce data.**", "**G5. Physics priors are unidentified under scarce data.**"),

    # ---- IX. Direction ----------------------------------------------------
    ("the robustness gap \u0394 of P2 \u2014 all as mean \u00b1 standard deviation across trials.",
     "the robustness gap \u0394 of P2, all as mean \u00b1 standard deviation across trials."),

    # ---- X. RQs -----------------------------------------------------------
    ("Which constraint carries the benefit under scarcity \u2014 the SDOF equation residual, a temporal smoothness penalty on *q\u0302*, or their combination \u2014 and does the benefit survive",
     "Which constraint carries the benefit under scarcity (the SDOF equation residual, a temporal smoothness penalty on *q\u0302*, or their combination) and does the benefit survive"),
]


def to_regex(pattern):
    """Whitespace runs in the pattern match any whitespace run in the source."""
    parts = re.split(r"(\s+)", pattern)
    return re.compile("".join(re.escape(p) if i % 2 == 0 else r'["\s]*' for i, p in enumerate(parts)))


with open(PATH, encoding="utf-8") as f:
    text = f.read()

problems = []
for old, new in PAIRS:
    rx = to_regex(old)
    found = rx.findall(text)
    if len(found) != 1:
        problems.append(f"expected 1 match, found {len(found)}: {old[:90]!r}")
        continue
    text = rx.sub(lambda _m, n=new: n, text, count=1)

if problems:
    print("FAILED:", file=sys.stderr)
    for p in problems:
        print("  " + p, file=sys.stderr)
    sys.exit(1)

with open(PATH, "w", encoding="utf-8") as f:
    f.write(text)

left = text.count(E)
print(f"Applied {len(PAIRS)} replacements.")
print(f"Remaining em dashes: {left}")
if left:
    sys.exit(1)
