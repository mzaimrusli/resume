"""Generate a plain, ATS-safe .docx resume: single column, no tables,
no images, no text boxes, standard section headings, Calibri 11pt.
Run with: python scripts/generate_ats_docx.py
"""
import os
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "ats")
OUT_PATH = os.path.join(OUT_DIR, "Muhammad_Zaim_Rusli_Resume.docx")

doc = Document()

style = doc.styles["Normal"]
style.font.name = "Calibri"
style.font.size = Pt(11)

def heading(text):
    p = doc.add_paragraph()
    run = p.add_run(text.upper())
    run.bold = True
    run.font.size = Pt(12)
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(4)

def para(text, bold=False, size=11, align_center=False):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    if align_center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return p

def bullet(text):
    doc.add_paragraph(text, style="List Bullet")

def entry_header(role, org, dates):
    p = doc.add_paragraph()
    r1 = p.add_run(f"{role} — {org}")
    r1.bold = True
    p.add_run(f"\t{dates}")
    tab_stops = p.paragraph_format.tab_stops
    tab_stops.add_tab_stop(Pt(430))

para("Muhammad Zaim Bin Rusli", bold=True, size=16, align_center=True)
para("Senior System Validation & Silicon Debug Engineer", size=12, align_center=True)
para("mzaimrusli@gmail.com | +60 16-610 4339 | linkedin.com/in/muhammad-zaim-rusli-805715108 | Penang, Malaysia", align_center=True)

heading("Professional Summary")
para(
    "Senior System Validation and Silicon Debug Engineer with nearly 10 years of experience spanning "
    "semiconductor manufacturing, ATE test development, pre-silicon verification, post-silicon validation, "
    "System Level Test (SLT) and data center platform qualification across Intel Xeon and AMD EPYC server processors."
)
para(
    "Proven expertise in system-level validation, production readiness qualification, silicon bring-up, "
    "failure analysis, yield improvement, RMA debug, hardware-software-firmware integration and manufacturing "
    "test optimization. Experienced in developing automated validation infrastructures using Python and Shell "
    "scripting while driving complex cross-domain root cause analysis across silicon, firmware, BIOS, operating "
    "systems and platform hardware."
)
para(
    "Unique background spanning wafer sort, package test, ATE, SLT, pre-silicon verification and post-silicon "
    "validation enables strong correlation across manufacturing, lab validation and field deployment environments."
)

heading("System Level Test and Manufacturing Highlights")
for b in [
    "Nearly 10 years of experience across semiconductor manufacturing test, ATE, post-silicon validation and production qualification.",
    "Hands-on experience supporting High Volume Manufacturing (HVM) environments.",
    "Developed and maintained automated test infrastructure for validation and production readiness programs.",
    "Expertise correlating manufacturing failures, system-level failures and silicon root causes.",
    "Extensive experience debugging issues across hardware, firmware, BIOS and operating system boundaries.",
    "Strong background in server platform bring-up including CPU, memory, PCIe devices, NICs and BMC infrastructures.",
    "Experience supporting release and manufacturing decisions based on validation readiness criteria.",
    "Proven ability to identify validation and test coverage gaps through structured failure analysis.",
    "Experienced in System Level Test (SLT), Manufacturing Readiness, Production Qualification, Yield Improvement, "
    "Failure Correlation, Tester Correlation, Root Cause Analysis, RMA Analysis, Product Characterization, "
    "Performance Characterization, Platform Bring-up, Power Characterization and Automation Infrastructure development.",
]:
    bullet(b)

heading("Technical Skills")
para("Architecture & Platforms: AMD EPYC, Intel Xeon, x86 Architecture, Server Platforms, Datacenter Infrastructure, PCIe Gen4/Gen5, CXL, DDR4/DDR5, BIOS/UEFI, OpenBMC")
para("Programming & Automation: Python, PyTest, Shell/Bash, C/C++, SystemVerilog, UVM, Simics Automation")
para("Debug & Diagnostics: Machine Check Exceptions (MCE), WHEA, BERT, Error Injection, Status Registers, Scan Dumps, dmesg, UART, BIOS Debug, Firmware Debug")
para("Tools: PRISM, HEXA, Git, JIRA, Simics, Verdi, DediProg, Linux, Windows")

heading("Professional Experience")

entry_header("Functional Validation Engineer", "Advanced Micro Devices (AMD), Penang, Malaysia", "June 2024 - Present")
for b in [
    "Execute post-silicon functional validation and production readiness qualification activities on AMD EPYC server processors.",
    "Perform system-level validation across CPU, memory, PCIe and firmware subsystems under customer-representative workloads.",
    "Conduct failure analysis using Machine Check Exceptions (MCE), WHEA Logger, BERT records and hardware status registers to identify silicon and platform root causes.",
    "Execute error injection, reliability testing and platform characterization to evaluate robustness of production silicon.",
    "Support platform bring-up, hardware assembly, BIOS configuration, OpenBMC management, operating system provisioning and regression infrastructure maintenance.",
    "Develop Python-based automation solutions within the PRISM validation framework to improve validation throughput and execution consistency.",
    "Work cross-functionally with architecture, firmware, BIOS and platform engineering teams to resolve complex system-level issues and improve validation coverage.",
]:
    bullet(b)

entry_header("System Validation Engineer", "Intel Microelectronics, Kulim, Malaysia", "August 2019 - May 2024")
para("Post-Silicon System Validation", bold=True)
for b in [
    "Executed structured test plans, silicon qualification activities and manufacturing readiness validation on Intel Xeon server processors.",
    "Defined system-level validation methodologies supporting production release decisions for complex server SoCs.",
    "Developed and maintained Python and Shell-based automation within the HEXA validation framework to improve regression efficiency and validation scalability.",
    "Conducted extensive stress, corner-case and concurrency testing to characterize silicon quality, reliability and scalability under production-like conditions.",
    "Performed root cause analysis across hardware, BIOS, firmware and operating system layers to resolve manufacturing-blocking and customer-impacting issues.",
    "Validated security, trusted execution and platform initialization features including TDX, SGX, ACM and SEAM Loader integration.",
    "Executed large-scale concurrency testing involving hundreds of simultaneous workloads to validate platform stability and memory resource management.",
    "Collaborated with architecture teams to drive silicon defect identification, debug and closure.",
    "Supported manufacturing release criteria through objective silicon validation metrics and readiness assessments.",
]:
    bullet(b)
para("Pre-Silicon Validation and Emulation", bold=True)
for b in [
    "Executed SoC functional validation in Simics Virtual Platform and FPGA emulation environments.",
    "Developed automation to validate boot flows, feature enablement and register-level functionality.",
    "Validated security and trusted execution features prior to silicon availability, reducing downstream silicon debug effort.",
]:
    bullet(b)
para("Client and Embedded Validation", bold=True)
for b in [
    "Performed system-level validation on embedded and client platform SoCs.",
    "Developed automation supporting boot verification and platform readiness checks.",
    "Collaborated with firmware and software teams to resolve integration issues and improve platform stability.",
]:
    bullet(b)

entry_header("IP Verification Engineer", "Intel Microelectronics, Penang, Malaysia", "March 2018 - August 2019")
for b in [
    "Verified R-Link Soft IP functionality using SystemVerilog and UVM methodologies.",
    "Developed constrained-random test scenarios to improve feature coverage and RTL bug detection.",
    "Performed waveform analysis and root cause investigation using Verdi simulation tools.",
    "Developed assertions and coverage metrics supporting verification closure objectives.",
    "Collaborated closely with RTL design teams to resolve protocol and integration issues prior to tape-out.",
]:
    bullet(b)

entry_header("Test and Product Engineer", "Malaysia Microelectronic Solutions (MMS), Selangor, Malaysia", "August 2017 - March 2018")
for b in [
    "Developed and debugged production ATE test programs supporting semiconductor high-volume manufacturing.",
    "Performed yield analysis, electrical characterization and probe-level failure debug.",
    "Supported manufacturing test optimization and production release readiness.",
    "Collaborated with external organizations to resolve hardware and probe-related production issues.",
]:
    bullet(b)

entry_header("Graduate Engineer", "NXP Semiconductors, Selangor, Malaysia", "September 2016 - July 2017")
for b in [
    "Developed Teradyne J750 production test solutions.",
    "Performed device characterization, failure analysis and tester correlation activities.",
    "Supported production qualification and yield improvement initiatives.",
    "Maintained automation scripts supporting manufacturing and characterization activities.",
]:
    bullet(b)

heading("Education")
para("Bachelor of Engineering (Hons.), Electronic Engineering — Universiti Teknologi MARA (UiTM), 2016")

os.makedirs(OUT_DIR, exist_ok=True)
doc.save(OUT_PATH)
print(f"Saved: {OUT_PATH}")
