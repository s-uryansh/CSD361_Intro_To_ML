#!/usr/bin/env python3
"""
build_docx.py: generates the Word (.docx) version of the Struct-X preliminary report
directly as Office Open XML, so that styles, lists, tables and the embedded figure are
all first-class Word objects (LibreOffice's HTML importer silently drops images here).

Inline markup understood by `runs()`:
    **bold**      *italic*      ~subscript~      ^superscript^
"""

import os
import re
import zipfile
import struct

HERE = os.path.dirname(os.path.abspath(__file__))
FIG = os.path.join(HERE, "fig1_pipeline.png")
OUT = os.path.join(HERE, "CSD361_Preliminary_Report_StructX.docx")

# --------------------------------------------------------------------------- #
# XML helpers
# --------------------------------------------------------------------------- #

def esc(t):
    return (t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


TOKEN = re.compile(r"(\*\*.+?\*\*|\*.+?\*|~.+?~|\^.+?\^)", re.S)


def runs(text, sz=None, color=None):
    """Convert lightweight markup into a sequence of <w:r> elements."""
    out = []
    for part in TOKEN.split(text):
        if not part:
            continue
        b = i = sub = sup = False
        if part.startswith("**") and part.endswith("**"):
            b, part = True, part[2:-2]
        elif part.startswith("*") and part.endswith("*"):
            i, part = True, part[1:-1]
        elif part.startswith("~") and part.endswith("~"):
            sub, part = True, part[1:-1]
        elif part.startswith("^") and part.endswith("^"):
            sup, part = True, part[1:-1]

        rpr = ["<w:rPr>"]
        if b:
            rpr.append("<w:b/>")
        if i:
            rpr.append("<w:i/>")
        if sz:
            rpr.append(f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>')
        if color:
            rpr.append(f'<w:color w:val="{color}"/>')
        if sub:
            rpr.append('<w:vertAlign w:val="subscript"/>')
        if sup:
            rpr.append('<w:vertAlign w:val="superscript"/>')
        rpr.append("</w:rPr>")
        rpr = "".join(rpr)
        if rpr == "<w:rPr></w:rPr>":
            rpr = ""
        out.append(f'<w:r>{rpr}<w:t xml:space="preserve">{esc(part)}</w:t></w:r>')
    return "".join(out)


def para(text="", style=None, align=None, sz=None, color=None, space_before=None,
         space_after=None, ind_left=None, ind_hanging=None, keep_next=False,
         border_top=False, shade=None, page_break_before=False):
    ppr = ["<w:pPr>"]
    if style:
        ppr.append(f'<w:pStyle w:val="{style}"/>')
    if page_break_before:
        ppr.append("<w:pageBreakBefore/>")
    if keep_next:
        ppr.append("<w:keepNext/>")
    if shade:
        ppr.append(f'<w:shd w:val="clear" w:color="auto" w:fill="{shade}"/>')
    if border_top:
        ppr.append('<w:pBdr><w:top w:val="single" w:sz="6" w:space="6" w:color="808080"/></w:pBdr>')
    if ind_left is not None or ind_hanging is not None:
        a = f' w:left="{ind_left}"' if ind_left is not None else ""
        b = f' w:hanging="{ind_hanging}"' if ind_hanging is not None else ""
        ppr.append(f"<w:ind{a}{b}/>")
    sp = []
    if space_before is not None:
        sp.append(f'w:before="{space_before}"')
    if space_after is not None:
        sp.append(f'w:after="{space_after}"')
    if sp:
        ppr.append(f'<w:spacing {" ".join(sp)}/>')
    if align:
        ppr.append(f'<w:jc w:val="{align}"/>')
    ppr.append("</w:pPr>")
    ppr = "".join(ppr)
    if ppr == "<w:pPr></w:pPr>":
        ppr = ""
    return f"<w:p>{ppr}{runs(text, sz=sz, color=color)}</w:p>"


def list_para(text, num_id, level=0):
    return (
        f'<w:p><w:pPr><w:pStyle w:val="ListParagraph"/>'
        f'<w:numPr><w:ilvl w:val="{level}"/><w:numId w:val="{num_id}"/></w:numPr>'
        f'<w:spacing w:after="80"/><w:jc w:val="both"/></w:pPr>'
        f"{runs(text)}</w:p>"
    )


def cell(text, width, header=False, center=False, shade=None, valign="top"):
    shd = f'<w:shd w:val="clear" w:color="auto" w:fill="{shade}"/>' if shade else ""
    body = text if text.startswith("<w:p") else para(
        text, style="TableText", align="center" if center else None,
        space_before=20, space_after=20)
    return (
        f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{shd}'
        f'<w:vAlign w:val="{valign}"/></w:tcPr>{body}</w:tc>'
    )


def table(rows, widths, header_rows=1):
    """rows: list of list[str|xml]; widths in dxa (1/20 pt)."""
    borders = (
        "<w:tblBorders>"
        + "".join(
            f'<w:{e} w:val="single" w:sz="4" w:space="0" w:color="7F7F7F"/>'
            for e in ("top", "left", "bottom", "right", "insideH", "insideV")
        )
        + "</w:tblBorders>"
    )
    xml = [
        f'<w:tbl><w:tblPr><w:tblW w:w="{sum(widths)}" w:type="dxa"/>'
        f'<w:jc w:val="center"/>{borders}'
        '<w:tblCellMar><w:top w:w="40" w:type="dxa"/><w:left w:w="70" w:type="dxa"/>'
        '<w:bottom w:w="40" w:type="dxa"/><w:right w:w="70" w:type="dxa"/></w:tblCellMar>'
        "</w:tblPr><w:tblGrid>"
        + "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
        + "</w:tblGrid>"
    ]
    for ri, row in enumerate(rows):
        is_h = ri < header_rows
        trpr = "<w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>" if is_h else "<w:trPr><w:cantSplit/></w:trPr>"
        cells = "".join(
            cell(c, widths[ci], header=is_h,
                 shade="E8E8E8" if is_h else None,
                 center=(ci > 0 and ri == 0))
            for ci, c in enumerate(row)
        )
        xml.append(f"<w:tr>{trpr}{cells}</w:tr>")
    xml.append("</w:tbl>")
    xml.append(para("", space_after=0, sz="12"))
    return "".join(xml)


def png_size(path):
    with open(path, "rb") as f:
        head = f.read(24)
    return struct.unpack(">II", head[16:24])


# --------------------------------------------------------------------------- #
# Document content
# --------------------------------------------------------------------------- #

TITLE = "Struct-X: Physics-Aware Event Vision for Robust Non-Contact Structural Vibration Monitoring"
SUBTITLE = "A Literature Survey, Gap Analysis and Preliminary Research Plan"
AUTHORS = "Suryansh Rohil  •  Arnab Mandal  •  Shresth Tiwari"

ABSTRACT = (
    "**Abstract.** Non-contact vibration measurement with an event camera is technically attractive: "
    "microsecond timestamps, ~140 dB dynamic range and negligible motion blur suit high-frequency structural "
    "response far better than contact or frame-based sensing. In practice, however, an event stream is not a "
    "picture of the structure but the residual of a sensor whose output depends on illumination, contrast "
    "threshold and bias currents, so the same physical motion is measured differently under different settings. "
    "This report surveys three methodological families that between them cover the pieces such a system needs: "
    "event-frame geometric deformation estimation, physics-guided learning for structural dynamics, and "
    "spatiotemporal Transformers with attention factorised over space and time. We find that the three have never "
    "been joined, that no reviewed work reports sensor-shift degradation as a first-class metric, and that "
    "spatiotemporal Transformers need more data than a 30-recording beam dataset can supply. We formalise the gap, "
    "propose a physics-aware event-vision pipeline with a reduced-order single-degree-of-freedom consistency term, "
    "and pre-register a falsification criterion: if the reduction in robustness gap does not exceed trial-to-trial "
    "variance, we report the physics term as not identifiable at this scale."
)

INDEX_TERMS = (
    "**Index Terms.** event cameras, non-contact vibration measurement, structural health monitoring, "
    "physics-informed neural networks, spatiotemporal Transformers, domain shift, robustness."
)

BODY = []

BODY.append(para("CSD361: Introduction to Machine Learning  •  Preliminary Project Report",
                 align="center", sz="20", color="444444", space_after="80"))
BODY.append(para(TITLE, style="Title", align="center"))
BODY.append(para(SUBTITLE, align="center", sz="24", space_after="200"))
BODY.append(para(AUTHORS, align="center", sz="23", space_after="40"))
BODY.append(para("Department of Computer Science and Engineering", align="center",
                 sz="20", color="444444", space_after="200"))
BODY.append(para("", border_top=True, space_after="200"))
BODY.append(para(ABSTRACT, style="Abstract"))
BODY.append(para(INDEX_TERMS, style="IndexTerms", space_after="240"))

# ---- I. Introduction -------------------------------------------------------
BODY.append(para("I.  Introduction", style="Heading1"))
BODY.append(para("A.  Background", style="Heading2"))
BODY.append(para(
    "Vibration is the observable that structural health monitoring (SHM) has historically rested on: damage changes "
    "stiffness, mass and damping, and those changes leave detectable fingerprints in natural frequencies, mode "
    "shapes and acceleration responses. The instrumentation, however, is awkward. Accelerometers must be bonded to "
    "the member being watched, altering the very dynamics they measure. A laser Doppler vibrometer (LDV) avoids "
    "contact but needs a clear line of sight and a head held still while the target moves: workable on a laboratory "
    "bench, much less so on a bridge or a rotor. Cameras have therefore been pursued for years as a cheap, scalable, "
    "non-contact substitute, and event cameras have sharpened the idea considerably."))
BODY.append(para(
    "An event camera, or dynamic vision sensor, abandons fixed-rate sampling. Each pixel asynchronously reports "
    "when the log intensity at its own location changes by more than a contrast threshold, emitting an event "
    "*e* = (*x*, *y*, *t*, *p*) with polarity *p* ∈ {−1, +1} [1]. The consequences are exactly those vibration "
    "measurement wants: microsecond timestamps, roughly 140 dB of dynamic range against about 60 dB for conventional "
    "CMOS, per-pixel bandwidth on the order of kilohertz, and little motion blur [1], [2]. For a beam oscillating at "
    "tens of hertz to a few kilohertz, a conventional camera either blurs the motion or demands lighting and bandwidth a field deployment cannot carry. This is not a marginal advantage."))

BODY.append(para("B.  Motivation", style="Heading2"))
BODY.append(para(
    "There is a catch, and it is why this project exists. An event stream is not an image of the world; it is the "
    "residual of a sensor with a handful of user-set knobs. Illumination level and spectral content determine how "
    "many events a moving edge generates. The contrast threshold *C* sets how large a brightness step a pixel must "
    "see before it speaks, and threshold mismatch across pixels is itself a noise source: Gallego et al. report the "
    "effective threshold as a random variable distributed approximately normally about the nominal setting, with a "
    "spread of 2–4% temporal contrast [1]. Refractory periods, bias currents and bandwidth limits shape the rest of "
    "the noise floor. Change any of these and the sensor produces a *different measurement of the same physical "
    "motion*. The domain shift in this project lives in the sensor, not the structure, and it is present in every "
    "recording we will ever take."))
BODY.append(para(
    "Meanwhile the SHM literature has spent the last several years walking in the opposite direction. It has learned "
    "to keep models usable when the environment changes and the damage class is unknown, via open-set domain "
    "adaptation [8]; to inject governing equations into the training loss so that data-hungry networks still behave "
    "when data is scarce and noisy [6], [7]; and that video models built purely from self-attention need their "
    "attention factorised over space and time to be trainable at all [9]. What we could not find, after searching "
    "the families described below, is any work joining these threads to the event camera. That is the hole this "
    "project intends to fill."))

# ---- II. Research Problem --------------------------------------------------
BODY.append(para("II.  Research Problem", style="Heading1"))
BODY.append(para(
    "Let *E*~s~ denote the event stream recorded under sensing configuration *s* (illumination, contrast threshold, "
    "bias setting, sensor variant), and *q*(*t*) the true vibration response measured by a co-located LDV, which we "
    "treat as ground truth. We seek a mapping *f*~θ~ : *E*~s~ → *q*(*t*) trained with a physics-aware objective, "
    "satisfying three requirements:"))
for t in [
    "**P1. Accuracy.** On held-out *recordings* from the configuration seen in training, *f*~θ~(*E*~s₀~) should "
    "approximate *q*(*t*) in waveform RMSE, MAE and correlation, and recover the dominant frequency *f*~n~.",
    "**P2. Stability.** The *robustness gap* Δ = *M*(shifted) − *M*(matched) should be small, where *M* is waveform "
    "RMSE or relative frequency error. This is what makes the approach a *sensing* method rather than a "
    "one-configuration demo.",
    "**P3. Feasibility.** The mapping must be learnable from 30 independent beam recordings [10], which forces "
    "windowed training data with trial-level splits and rules out any large model trained from scratch.",
]:
    BODY.append(list_para(t, 1))
BODY.append(para(
    "Two scope decisions follow. First, this is a study of vibration *estimation* and sensing robustness, not damage "
    "classification: the XJTU-DV Beam subset provides no damage labels, and we decline to claim what the data cannot "
    "support. Second, our regression target is a continuous waveform rather than a categorical label, which changes "
    "what “robustness” has to mean (Section VIII)."))

# ---- III. Objectives -------------------------------------------------------
BODY.append(para("III.  Objectives of the Literature Survey", style="Heading1"))
BODY.append(para(
    "The survey answers five questions in sequence, following the read → categorise → compare → critique → gap → "
    "direction chain rather than summarising papers one at a time."))
for t in [
    "**O1.** Establish what an event camera physically measures, and which of its parameters are set by the user "
    "rather than by the scene. This decides whether sensor-induced domain shift is real or theoretical.",
    "**O2.** Identify pipelines that recover deformation or displacement from event data, recording what each assumes "
    "(brightness constancy, spatial continuity, co-registered frames) and what supervision it needs.",
    "**O3.** Identify how physics is injected into deep learning for structural dynamics, recording which physical "
    "form the constraint takes and where the required parameters come from.",
    "**O4.** Identify spatiotemporal architectures suited to event-derived tensors, with attention to their behaviour "
    "in low-data regimes.",
    "**O5.** Determine how the literature measures, or fails to measure, degradation under changed sensing or "
    "environmental conditions for a continuous-valued target.",
]:
    BODY.append(list_para(t, 2))  # numId 2 -> Section III list, restarts at 1

# ---- IV. Methodology -------------------------------------------------------
BODY.append(para("IV.  Literature Search Methodology", style="Heading1"))
BODY.append(para(
    "Six sources were queried: IEEE Xplore, ScienceDirect (including *Mechanical Systems and Signal Processing* and "
    "*Engineering Structures*), SpringerLink, Wiley Online Library, the ACM Digital Library and arXiv; the dataset "
    "record was retrieved separately from Zenodo. Searching ran in three passes: a broad keyword pass, a "
    "backward-citation pass from the two anchor surveys [1], [5], and a forward-citation pass from the XJTU-DV "
    "dataset paper [10]. Keyword groups and their yields are in Table 1."))
BODY.append(para("Table 1.  Search protocol. 17 candidates retrieved across three passes; 12 retained after "
                 "two-stage screening (5 core, 7 supporting).", style="Caption", keep_next=True))
BODY.append(table(
    [
        ["Concept group", "Keywords", "Retained"],
        ["Sensor", "event camera; dynamic vision sensor; contrast threshold; illumination; noise", "3"],
        ["Task", "non-contact vibration measurement; visual deformation measurement; beam", "2"],
        ["Method", "physics-informed; physics-guided; neural network; structural dynamics", "4"],
        ["Architecture", "spatiotemporal transformer; space-time attention; divided attention", "2"],
        ["Robustness", "domain adaptation; open set; out-of-distribution", "1"],
    ],
    [1900, 5600, 1100],
))
BODY.append(para(
    "Records were excluded if they addressed only frame-based sensing, if the physical system was not a vibrating "
    "structure, or if the reported task was classification where our target is a waveform. Priority followed the "
    "course brief: recent journal papers, top-tier conference papers, foundational work, systematic reviews, then "
    "work directly on non-contact vibration sensing. All twelve retained sources are listed in Section XII."))

# ---- V. Review -------------------------------------------------------------
BODY.append(para("V.  Detailed Review of Existing Literature", style="Heading1"))
BODY.append(para(
    "The retained work clusters into three methodological families, reviewed here on their own terms before "
    "comparison."))

BODY.append(para("A.  Approach 1: Geometric, largely training-free event-frame deformation measurement",
                 style="Heading2"))
BODY.append(para(
    "The most mature line of work relevant to us measures deformation directly from event data with an explicit "
    "motion model rather than a learned one. Wu et al. [4] fuse an event stream (temporally dense, spatially sparse, "
    "cheap) with a frame stream (spatially dense, low-noise) and solve for a dense displacement field. Their "
    "*Affine Invariant Simplicial* (AIS) framework partitions that field into low-parameter sub-regions and "
    "linearises each via simplicial motion parameterisation, exploiting the reproduction property *I*[*f*] = *f* for "
    "affine *f*(*X*) = *AX* + *b* to suppress the ambiguity sparse, noisy events introduce. Motion is recovered by "
    "optimising an event contrast-maximisation objective jointly with image cross-correlation, coarse-to-fine, with "
    "a neighborhood-greedy scheme letting well-converged sub-regions guide their neighbours. On a purpose-built "
    "benchmark of 120+ aligned event/frame sequences with high-speed-video ground truth, the method reaches a 65.7% "
    "survival rate beyond 100 px displacement (1.6× the previous state of the art) using 18.9% of the storage and "
    "processing budget of high-speed video."))
BODY.append(para(
    "Two features matter to us beyond those numbers. The method is entirely training-free, so there is no train/test "
    "split to leak and no data-scarcity problem. And its authors are explicit that no suitable event dataset "
    "existed, that building one is costly, that structural health monitoring is future work, and that the method "
    "leans on spatial continuity and brightness constancy, with pseudo-strains under topological change and degraded "
    "accuracy under specular reflection [4]. The wider sensor literature supplies the rest of this family: the "
    "representation taxonomy of Gallego et al. [1] (event frames and time surfaces retaining one timestamp per "
    "pixel; grid-like voxel tensors that drop into conventional CNN pipelines; learned representations) and the "
    "simulator conventions of Gehrig and Scaramuzza [2], which let a method be stress-tested against camera "
    "parameter changes without recapturing hardware."))

BODY.append(para("B.  Approach 2: Physics-informed and physics-guided learning for structural dynamics",
                 style="Heading2"))
BODY.append(para(
    "This family treats physics as a training signal rather than an inference procedure. Raissi et al. [5] define "
    "the general form: a network is trained on a data-fidelity term while simultaneously satisfying the residual of "
    "a governing nonlinear PDE, in continuous-time and discrete-time variants, the latter admitting implicit "
    "Runge–Kutta schemes of arbitrary order. The framework is data-efficient precisely because the physics term "
    "constrains the hypothesis set; Karniadakis et al. [12] later generalise it and document where it fails: "
    "training can fail for reasons unconnected to the physics, and poorly chosen residuals actively hurt."))
BODY.append(para(
    "The structural instantiation closest to Struct-X is Wang et al.’s PhyResNet [6]. A ResNet encoder maps "
    "acceleration responses to low-dimensional structural states (stiffness reductions); a physics-based decoder (an updated finite element model solved by Newmark-β) reconstructs the response from those states. The equation "
    "of motion *M ẍ* + *C ẋ* + *K x* = *F*(*t*) enters as one loss term, a second regresses the states against "
    "ground truth. The numbers establish the regime in which physics helps. On a simulated simply supported beam "
    "(10 elements) with 1–3% modelling error and 10%/20% Gaussian measurement noise, trained on 5% of the data, "
    "PhyResNet improves R-value by 13.1% over the pure ResNet; at 10% noise R rises from 0.789 to 0.870, and at 20% "
    "noise it holds 0.808; with modelling error alone, MSE falls from 2.0×10^−4^ to 1.6×10^−4^ and R from 0.805 to "
    "0.850. A seven-storey frame experiment reproduces the trend, and the authors are candid that physics-informed "
    "learning in SHM remains early-stage."))
BODY.append(para(
    "Closest of all to our setting is the work reconstructing structural displacement from *both* accelerometer and "
    "camera measurements under a physics-informed objective [7]. That is structurally the same problem Struct-X faces, fusing a non-contact visual measurement with a physical response and letting the equation of motion regularise "
    "the fit. Its weakness is that the camera modality is a conventional frame camera, inheriting frame-camera "
    "failure modes under fast, small-amplitude motion, and its physics residual is written on the accelerometer "
    "channel rather than a high-bandwidth optical one."))

BODY.append(para("C.  Approach 3: Spatiotemporal Transformers and robustness-aware SHM", style="Heading2"))
BODY.append(para(
    "This family supplies the architecture. Bertasius et al. [9] remove convolution entirely and build video "
    "classification from self-attention over frame-level patches, comparing several attention schemes. Their "
    "operationally important finding: *divided attention*, applying temporal and spatial attention separately within "
    "each block, gives the best accuracy among the choices considered, whereas joint space-time attention is not "
    "merely worse but overflows GPU memory at the resolutions they target. TimeSformer sets the best reported "
    "Kinetics-400 accuracy (80.7% top-1, large variant), trains faster than 3D convolutions, and handles clips over "
    "a minute long."))
BODY.append(para(
    "The same paper contains a warning we have taken seriously. Bertasius et al. attribute the convolutional bias (local connectivity and translation equivariance) as genuinely *beneficial on small training sets*, and argue "
    "Transformers pay off in big-data regimes where less inductive bias is needed; their spatial resolution ablation "
    "shows that accuracy gains from finer spatiotemporal partitioning require proportionally more pretraining. On "
    "robustness, Zhou and Lai [8] confront shift for classification, combining open-set domain adaptation with "
    "varying environmental conditions and unknown damage classes. The transferable idea is not their mechanism but "
    "their insistence that train and test conditions differ *by construction* rather than by random split."))

# ---- VI. Comparative -------------------------------------------------------
BODY.append(para("VI.  Comparative Analysis", style="Heading1"))
BODY.append(para(
    "Table 2 places the three families side by side on the axes that decide whether a method can be adapted to "
    "Struct-X. The last three rows matter most: none reports degradation under a changed sensing configuration, and "
    "only one needs no data at all, at the cost of a co-registered frame stream and an assumption of brightness "
    "constancy."))
BODY.append(para("Table 2.  Comparative analysis against requirements P1–P3 of Section II. Outputs are respectively "
                 "a dense per-pixel displacement field, a damage index / reconstructed response, and a class label; "
                 "none of the three emits a calibrated confidence.", style="Caption", keep_next=True))
BODY.append(table(
    [
        ["Dimension", "A. Geometric event-frame VDM [1], [2], [4]",
         "B. Physics-guided structural DL [5]–[7], [12]",
         "C. Spatiotemporal Transformer + robust SHM [8], [9]"],
        ["Input representation", "Raw events + co-registered frames; event frames, time surfaces, voxel grids",
         "1-D sensor series (acceleration, displacement), sometimes FEM-derived features",
         "Stacked voxel-grid / frame-level patches"],
        ["Supervision", "None (training-free optimisation); ground truth used only for evaluation",
         "Supervised on damage indices or response traces; low-data tolerance is the headline result",
         "Supervised, large-scale video labels; pretraining strongly recommended"],
        ["Form of physics", "Implicit: linearised solid-elastic / simplicial motion prior inside the estimator",
         "Explicit PDE residual in the loss; the decoder may be an FEM (Newmark-β)", "None"],
        ["Needs an FE model?", "No", "Yes (or equivalent identified parameters)", "No"],
        ["Data regime", "Any (per-sequence optimisation)", "Modest; gains largest at 5–10% of data",
         "Large; degrades with small samples [9]"],
        ["Handles sensing shift?", "Not evaluated", "Not evaluated (noise / modelling error only)",
         "Yes, by design [8], but categorical targets"],
        ["Meets P1 (accuracy)?", "Yes, for displacement", "Yes, for damage index", "Unproven for waveforms"],
        ["Meets P2 (stability)?", "Unknown", "Partial (noise only)", "Partial, wrong target type"],
        ["Meets P3 (30 recordings)?", "Yes", "Yes", "Only if pretrained"],
    ],
    [1750, 2900, 2900, 2600],
))

# ---- VII. Critical ---------------------------------------------------------
BODY.append(para("VII.  Critical Analysis", style="Heading1"))
for t in [
    "**Approach 1 buys robustness with assumptions we cannot accept.** The AIS framework is elegant precisely "
    "because it refuses to learn, and its authors are honest that it leans on spatial continuity and brightness "
    "constancy. Both fail exactly where we need them. Structural vibration at tens to thousands of hertz produces "
    "displacements small in amplitude relative to pixel pitch and fast relative to any achievable frame rate; "
    "brightness constancy is violated by every lighting change in the shifted conditions, and spatial continuity is "
    "violated wherever the beam edge is the only thing generating events. The method also needs a synchronised, "
    "co-registered frame stream, and returns a field by per-sequence optimisation: no learned prior to amortise "
    "across 30 trials, and no calibrated uncertainty on the output.",

    "**Approach 2 has the right inductive bias pointed the wrong way.** Every instantiation we found is formulated "
    "on *point* sensors with a known structural model. PhyResNet’s decoder literally is a converged finite element "
    "model solved by Newmark-β, and the physics term is meaningful only because *M*, *C*, *K* are known. Struct-X "
    "inverts this: we have rich imagery and *no* FE model of the beam, and our ground truth is an LDV trace, "
    "already a physical measurement of the response. Borrowing PhyResNet’s loss verbatim would require us to build "
    "and converge an FE model we do not have, importing exactly the modelling-error problems that paper measures. "
    "The honest adaptation is a *reduced-order* single-degree-of-freedom residual with *m*, *c*, *k* identified from "
    "a single training trial, which immediately creates the failure mode of Section IX.",

    "**Approach 3 is documented to be wrong at our data scale.** We read the TimeSformer results as a constraint, "
    "not an endorsement. Divided attention is right on principle, but the same ablation that selects it shows "
    "accuracy degrading toward the data-driven baseline as pretraining data thins, and 30 recordings is far below "
    "the regime where any from-scratch video Transformer is defensible. Any Struct-X encoder must therefore be "
    "pretrained or deliberately lightweight, and we should report the CNN encoder too rather than quietly presenting "
    "the Transformer as the main model.",

    "**The cross-cutting omission is the metric.** No reviewed paper reports degradation under a changed *sensing* "
    "configuration as a headline number. Approaches 1 and 2 report accuracy under measurement noise and FEM "
    "modelling error, a different perturbation: noise is additive and zero-mean, whereas a change in illumination or "
    "contrast threshold is a *multiplicative, structured* change in the event statistics that moves signal and noise "
    "together. Approach 3 confronts shift, but for categorical labels, where robustness is a confusion-matrix "
    "property; for a continuous waveform the equivalent notion, namely how far the reconstructed trace and recovered frequency move, has no accepted definition here. Establishing that definition is itself a contribution, and "
    "we propose it in Section VIII.",

    "**A methodological risk we should name.** Almost all of this literature splits at the sample or window level. "
    "With sliding windows drawn from continuous recordings, neighbouring windows overlap in time and share structure "
    "response, so a random window split leaks test data into training and inflates every number. We will split by "
    "complete recording and report mean ± standard deviation across independent trials, never a single aggregate "
    "figure.",
]:
    BODY.append(para(t))

# ---- VIII. Gaps ------------------------------------------------------------
BODY.append(para("VIII.  Research Gaps to Be Addressed", style="Heading1"))
for t in [
    "**G1. No event-vibration pipeline exists.** Event-based visual deformation measurement [4] is the closest "
    "artefact and positions itself short of SHM; no reviewed work predicts an LDV-referenced waveform from events.",
    "**G2. Physics and events have never been combined.** Physics-guided structural networks consume point-sensor "
    "series [6], [7]; event pipelines are training-free or purely data-driven [4]. The two literatures do not cite "
    "each other.",
    "**G3. Sensing-condition shift is unmeasured.** No reviewed study reports a robustness gap under changed "
    "illumination, contrast threshold or sensor configuration. The shift event cameras are *most* exposed to is the "
    "one *least* characterised.",
    "**G4. No waveform-domain definition of robustness.** Reviewed robustness metrics are classification metrics "
    "[8] or additive-noise stress tests [6]; there is no continuous-target analogue for a measured physical response.",
    "**G5. Physics priors are unidentified under scarce data.** In [6] the physical parameters are given. Where "
    "they must be estimated from the same scarce data the network learns from, the residual can be driven to zero by "
    "the degenerate solution *q̂* → 0, and we found no treatment of this in the structural literature.",
]:
    BODY.append(list_para(t, 3))  # numId 3 -> Section VIII list, restarts at 1

# ---- IX. Direction ---------------------------------------------------------
BODY.append(para("IX.  Proposed Research Direction", style="Heading1"))
BODY.append(para(
    "We propose Struct-X: a physics-aware spatiotemporal event-vision pipeline trained on the XJTU-DV Beam subset "
    "[10] with LDV ground truth, and evaluated primarily on how little performance moves when the sensing "
    "configuration changes. Figure 1 shows the pipeline; Section X states the falsification criteria."))

# Figure 1 (image inserted below via placeholder)
FIG_PLACEHOLDER = "<!--FIG1-->"
BODY.append(FIG_PLACEHOLDER)
BODY.append(para("Fig. 1.  Struct-X pipeline. Events are encoded into a spatiotemporal tensor, decoded into a "
                 "vibration waveform and a dominant-frequency estimate, and constrained by a reduced-order "
                 "single-degree-of-freedom residual. The contrastive term is an optional extension, not part of the "
                 "core model.", style="Caption"))

BODY.append(para(
    "**Physics layer.** For free vibration we adopt the damped single-mode model *m q̈*(*t*) + *c q̇*(*t*) + "
    "*k q*(*t*) = 0, with natural frequency *f*~n~ = (1/2π)√(*k*/*m*). The parameters are identified once by modal "
    "fitting on a *single* training trial and then frozen, so the physics term never sees test data. To defeat the "
    "degenerate collapse of G5, the residual is normalised by ‖*q̂*‖ so that shrinking the prediction toward zero "
    "does not reduce the loss; we also monitor the ratio of the two gradient contributions to confirm the data term "
    "still dominates early in training."))
BODY.append(para(
    "**Experimental design.** The central experiment is a leave-one-condition-out grid: train on condition *A*, then "
    "test on (i) held-out trials of *A* (the matched upper bound), (ii) low→high illumination, (iii) a different "
    "camera configuration, and (iv) a combined shift. Baselines are an FFT/feature-extraction pipeline, a CNN on the "
    "same representations, and a Transformer without the physics term. Metrics are waveform RMSE/MAE and "
    "correlation against LDV, relative frequency error |*f̂*~n~ − *f*~n~|/*f*~n~, the normalised physics residual, and "
    "the robustness gap Δ of P2, all as mean ± standard deviation across trials."))

# ---- X. RQs ----------------------------------------------------------------
BODY.append(para("X.  Preliminary Research Questions", style="Heading1"))
for t in [
    "**RQ1 (feasibility).** Within one acquisition condition, to what waveform RMSE, correlation and relative "
    "frequency error can a spatiotemporal event model reconstruct the LDV-measured response on held-out recordings, "
    "and how much of that is attributable to the physics term rather than the architecture?",
    "**RQ2 (mechanism).** Which constraint carries the benefit under scarcity (the SDOF equation residual, a temporal smoothness penalty on *q̂*, or their combination) and does the benefit survive once *m*, *c*, *k* are "
    "identified from data rather than given?",
    "**RQ3 (the actual claim).** Does the physics-aware model reduce the robustness gap Δ under illumination and "
    "contrast-threshold shift relative to each baseline, and is that reduction due to the physics term or merely to "
    "regularisation?",
    "**RQ4 (which output is stable).** Is the dominant-frequency estimate more shift-robust than waveform RMSE? We "
    "expect so, since a spectral peak is a coarser statistic than a pointwise trace, but this is an assumption to "
    "test rather than to build on.",
    "**RQ5 (falsification criterion, pre-registered).** If the reduction in Δ attributable to the physics term is "
    "smaller than the trial-to-trial standard deviation of Δ, we conclude the physics constraint is *not "
    "identifiable* at this data scale and report that as the outcome, negative result included. We will likewise "
    "report whether the physics term degrades the matched-condition result, since a regulariser that helps under "
    "shift while hurting in-distribution is a materially different finding.",
]:
    BODY.append(list_para(t, 4))  # numId 4 -> Section X list, restarts at 1

# ---- XI. Conclusion --------------------------------------------------------
BODY.append(para("XI.  Conclusion", style="Heading1"))
BODY.append(para(
    "Non-contact vibration measurement with an event camera is technically well-motivated and, on the evidence "
    "gathered here, methodologically unclaimed. The three families surveyed each solve a third of the problem: "
    "geometric event-frame estimators recover deformation but assume brightness constancy and need frames; "
    "physics-guided structural networks supply the right inductive bias but demand a finite element model our "
    "setting does not have; spatiotemporal Transformers supply the right architecture but underperform when data is "
    "scarce. What is missing, and what this project proposes to test, is the joint claim: that a physics-aware "
    "spatiotemporal model can estimate LDV-referenced vibration from events *and* degrade less when the sensing "
    "configuration changes, measurably above trial-to-trial noise on 30 recordings. This is a feasibility study on "
    "one physical system rather than a benchmark, and RQ5 commits us to reporting a negative result plainly if the "
    "evidence does not support the claim."))

# ---- XII. References -------------------------------------------------------
BODY.append(para("XII.  References", style="Heading1"))
REFS = [
    "[1]  G. Gallego, T. Delbrück, G. Orchard, C. Bartolozzi, B. Taba, A. Censi, S. Leutenegger, A. J. Davison, "
    "J. Conradt, K. Daniilidis, and D. Scaramuzza, “Event-based vision: A survey,” *IEEE Trans. Pattern Anal. "
    "Mach. Intell.*, vol. 44, no. 1, pp. 154–180, Jan. 2022.",

    "[2]  D. Gehrig and D. Scaramuzza, “Event-based vision meets computer vision,” *IEEE Trans. Pattern Anal. "
    "Mach. Intell.*, vol. 45, no. 10, pp. 15440–15468, 2023.",

    "[3]  H. Rebecq, T. Zhu, J. Moraitis, R. F. A. van den Hant, and D. Scaramuzza, “High speed and high dynamic "
    "range video with an event camera,” in *Proc. IEEE/CVF Conf. Comput. Vis. Pattern Recognit. (CVPR)*, 2017, "
    "pp. 2194–2204.",

    "[4]  Y. Wu, W. Zhai, Y. Cui, T. Zhao, Y. Cao, and Z.-J. Zha, “Event-based visual deformation measurement,” "
    "*arXiv:2602.14376*, 2026.",

    "[5]  M. Raissi, P. Perdikaris, and G. E. Karniadakis, “Physics-informed neural networks: A deep learning "
    "framework for solving forward and inverse problems involving nonlinear partial differential equations,” "
    "*J. Comput. Phys.*, vol. 378, pp. 686–707, 2019.",

    "[6]  R. Wang, J. Li, L. Li, S. An, B. Ezard, Q. Li, and H. Hao, “Structural damage identification by using "
    "physics-guided residual neural networks,” *Eng. Struct.*, vol. 318, art. 118703, 2024.",

    "[7]  “Accurate structural displacement reconstruction from acceleration and computer vision measurements "
    "using physics-informed neural networks,” *Mech. Syst. Signal Process.*, 2025.",

    "[8]  M. Zhou and Z. Lai, “Structural damage classification under varying environmental conditions and unknown "
    "classes via open set domain adaptation,” *Mech. Syst. Signal Process.*, 2024.",

    "[9]  G. Bertasius, H. Wang, and L. Torresani, “Is space-time attention all you need for video understanding?” "
    "in *Proc. Int. Conf. Mach. Learn. (ICML)*, 2021, pp. 813–824.",

    "[10]  X. Li, S. Yu, X. Chen, Y. Lei, B. Yang, and N. Li, “XJTU-DV: Open-source dynamic vision dataset for "
    "non-contact vibration measurement and fault diagnosis of mechanical systems,” *Chinese J. Mech. Eng.*, 2026.",

    "[11]  K. He, X. Zhang, S. Ren, and J. Sun, “Deep residual learning for image recognition,” in *Proc. IEEE Conf. "
    "Comput. Vis. Pattern Recognit. (CVPR)*, 2016, pp. 770–778.",

    "[12]  G. E. Karniadakis, I. G. Kevrekidis, L. Lu, A. Perdikaris, S. Wang, and Y. Yang, “Physics-informed machine "
    "learning,” *Nature Rev. Phys.*, vol. 3, pp. 422–440, 2021.",
]
for r in REFS:
    BODY.append(para(r, style="RefItem"))

BODY.append(para(
    "**Note.**  Table 1 records are our own; numbers quoted in Sections V–VII are as reported by the cited sources. "
    "All splits in Section IX are at recording level, not window level. A PDF version of this report is typeset in "
    "IEEE two-column conference format and is submitted alongside this document.",
    style="Note", border_top=True))

# --------------------------------------------------------------------------- #
# Figure XML
# --------------------------------------------------------------------------- #

EMU = 914400


def figure_xml(img_path, width_in):
    px_w, px_h = png_size(img_path)
    cx = int(width_in * EMU)
    cy = int(cx * px_h / px_w)
    did = 10
    return (
        '<w:p><w:pPr><w:jc w:val="center"/><w:keepNext/>'
        '<w:spacing w:before="200" w:after="60"/></w:pPr><w:r><w:drawing>'
        f'<wp:inline distT="0" distB="0" distL="0" distR="0" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
        f'<wp:extent cx="{cx}" cy="{cy}"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:docPr id="{did}" name="Figure 1" descr="Struct-X pipeline diagram"/>'
        '<wp:cNvGraphicFramePr>'
        '<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>'
        '</wp:cNvGraphicFramePr>'
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:nvPicPr><pic:cNvPr id="{did}" name="fig1_pipeline.png"/><pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="rIdFig1" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>'
        "<a:stretch><a:fillRect/></a:stretch></pic:blipFill>"
        '<pic:spPr><a:xfrm><a:off x="0" y="0"/>'
        f'<a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"
    )


BODY = [figure_xml(FIG, 3.45) if b == FIG_PLACEHOLDER else b for b in BODY]

# --------------------------------------------------------------------------- #
# Package parts
# --------------------------------------------------------------------------- #

SECT = (
    "<w:sectPr>"
    '<w:pgSz w:w="12240" w:h="15840"/>'
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
    'w:header="720" w:footer="720" w:gutter="0"/>'
    '<w:cols w:space="720"/><w:docGrid w:linePitch="360"/>'
    "</w:sectPr>"
)

DOCUMENT = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
    f"<w:body>{''.join(BODY)}{SECT}</w:body></w:document>"
)


def style(sid, name, based="Normal", ppr="", rpr="", quick=True, typ="paragraph"):
    return (
        f'<w:style w:type="{typ}" w:styleId="{sid}">'
        f'<w:name w:val="{name}"/>'
        + (f'<w:basedOn w:val="{based}"/>' if based else "")
        + ("<w:qFormat/>" if quick else "")
        + (f"<w:pPr>{ppr}</w:pPr>" if ppr else "")
        + (f"<w:rPr>{rpr}</w:rPr>" if rpr else "")
        + "</w:style>"
    )


STYLES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    "<w:docDefaults><w:rPrDefault><w:rPr>"
    '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
    '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-GB"/>'
    "</w:rPr></w:rPrDefault>"
    "<w:pPrDefault><w:pPr>"
    '<w:spacing w:after="140" w:line="264" w:lineRule="auto"/>'
    '<w:jc w:val="both"/>'
    "</w:pPr></w:pPrDefault></w:docDefaults>"

    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
    '<w:name w:val="Normal"/><w:qFormat/></w:style>'

    + style("Title", "Title",
            ppr='<w:spacing w:before="0" w:after="80" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/>',
            rpr='<w:b/><w:sz w:val="34"/><w:szCs w:val="34"/>')

    + style("Heading1", "heading 1",
            ppr='<w:keepNext/><w:outlineLvl w:val="0"/><w:spacing w:before="320" w:after="120"/>'
                '<w:jc w:val="left"/><w:pBdr><w:bottom w:val="single" w:sz="4" w:space="3" w:color="BFBFBF"/></w:pBdr>',
            rpr='<w:b/><w:sz w:val="26"/><w:szCs w:val="26"/><w:color w:val="1A1A1A"/>')

    + style("Heading2", "heading 2",
            ppr='<w:keepNext/><w:outlineLvl w:val="1"/><w:spacing w:before="220" w:after="80"/><w:jc w:val="left"/>',
            rpr='<w:b/><w:i/><w:sz w:val="23"/><w:szCs w:val="23"/>')

    + style("Abstract", "Abstract",
            ppr='<w:spacing w:before="80" w:after="140"/><w:ind w:left="284" w:right="284"/>'
                '<w:pBdr><w:left w:val="single" w:sz="18" w:space="8" w:color="404040"/></w:pBdr>'
                '<w:shd w:val="clear" w:color="auto" w:fill="F4F4F4"/>',
            rpr='<w:sz w:val="20"/><w:szCs w:val="20"/>')

    + style("IndexTerms", "Index Terms",
            ppr='<w:spacing w:after="200"/><w:ind w:left="284" w:right="284"/>',
            rpr='<w:i/><w:sz w:val="20"/><w:szCs w:val="20"/>')

    + style("Caption", "caption",
            ppr='<w:keepNext/><w:spacing w:before="140" w:after="80"/><w:jc w:val="center"/>',
            rpr='<w:sz w:val="19"/><w:szCs w:val="19"/><w:color w:val="333333"/>')

    + style("TableText", "Table Text",
            ppr='<w:spacing w:before="20" w:after="20" w:line="240" w:lineRule="auto"/><w:jc w:val="left"/>',
            rpr='<w:sz w:val="17"/><w:szCs w:val="17"/>')

    + style("ListParagraph", "List Paragraph",
            ppr='<w:contextualSpacing/><w:spacing w:after="80"/><w:jc w:val="both"/>')

    + style("RefItem", "Reference Item",
            ppr='<w:spacing w:after="80" w:line="240" w:lineRule="auto"/>'
                '<w:ind w:left="360" w:hanging="360"/><w:jc w:val="left"/>',
            rpr='<w:sz w:val="19"/><w:szCs w:val="19"/>')

    + style("Note", "Note",
            ppr='<w:spacing w:before="200" w:after="0" w:line="240" w:lineRule="auto"/>'
                '<w:pBdr><w:top w:val="single" w:sz="4" w:space="6" w:color="808080"/></w:pBdr>',
            rpr='<w:sz w:val="18"/><w:szCs w:val="18"/><w:color w:val="333333"/>')

    + '<w:style w:type="table" w:default="1" w:styleId="TableNormal">'
      '<w:name w:val="Normal Table"/><w:tblPr>'
      '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>'
      '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>'
      "</w:tblPr></w:style>"

    + "</w:styles>"
)


def abstract_num(aid, fmt, text, font=None):
    lv = []
    for lvl in range(3):
        t = text if lvl == 0 else (text if fmt == "bullet" else "%" + str(lvl + 1) + ".")
        f = f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:hint="default"/></w:rPr>' if font else ""
        lv.append(
            f'<w:lvl w:ilvl="{lvl}"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/>'
            f'<w:lvlText w:val="{t}"/><w:lvlJc w:val="left"/>'
            f'<w:pPr><w:ind w:left="{360 + 360 * lvl}" w:hanging="360"/></w:pPr>{f}</w:lvl>'
        )
    return (f'<w:abstractNum w:abstractNumId="{aid}">'
            f'<w:multiLevelType w:val="hybridMultilevel"/>{"".join(lv)}</w:abstractNum>')


NUMBERING = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    + abstract_num(0, "bullet", "•", font="Symbol")
    + abstract_num(1, "decimal", "%1.")
    # Each numbered list gets its own <w:num> with an explicit startOverride so that
    # Word restarts numbering at 1 for every list in the document.
    + '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
    + "".join(
        f'<w:num w:numId="{n}"><w:abstractNumId w:val="1"/>'
        f'<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>'
        f'<w:lvlOverride w:ilvl="1"><w:startOverride w:val="1"/></w:lvlOverride>'
        f'<w:lvlOverride w:ilvl="2"><w:startOverride w:val="1"/></w:lvlOverride>'
        "</w:num>"
        for n in (2, 3, 4, 5)
    )
    + "</w:numbering>"
)

CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Default Extension="png" ContentType="image/png"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
    "</Types>"
)

ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
    "</Relationships>"
)

DOC_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
    '<Relationship Id="rIdFig1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/fig1_pipeline.png"/>'
    "</Relationships>"
)

SETTINGS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/>'
    "<w:compat><w:compatSetting "
    'w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
    "</w:settings>"
)

CORE = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
    'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
    'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
    f"<dc:title>{esc(TITLE)}</dc:title>"
    "<dc:subject>CSD361 Preliminary Project Report: Literature Survey and Research Plan</dc:subject>"
    "<dc:creator>Suryansh Rohil; Arnab Mandal; Shresth Tiwari</dc:creator>"
    "<cp:keywords>event cameras; structural health monitoring; physics-informed neural networks; "
    "spatiotemporal Transformers; domain shift</cp:keywords>"
    "<cp:lastModifiedBy>Suryansh Rohil</cp:lastModifiedBy>"
    "</cp:coreProperties>"
)

APP = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
    'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
    "<Application>Microsoft Office Word</Application>"
    "<Company>Department of Computer Science and Engineering</Company>"
    "</Properties>"
)

# --------------------------------------------------------------------------- #
# Write package
# --------------------------------------------------------------------------- #

with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CONTENT_TYPES)
    z.writestr("_rels/.rels", ROOT_RELS)
    z.writestr("word/document.xml", DOCUMENT)
    z.writestr("word/_rels/document.xml.rels", DOC_RELS)
    z.writestr("word/styles.xml", STYLES)
    z.writestr("word/numbering.xml", NUMBERING)
    z.writestr("word/settings.xml", SETTINGS)
    z.writestr("docProps/core.xml", CORE)
    z.writestr("docProps/app.xml", APP)
    with open(FIG, "rb") as f:
        z.writestr("word/media/fig1_pipeline.png", f.read())

print(f"Wrote {OUT} ({os.path.getsize(OUT) / 1024:.1f} KB)")
