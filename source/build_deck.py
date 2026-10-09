#!/usr/bin/env python3
"""DevSecOps Academy ù The DevSecOps Programme. Syllabus deck."""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

OUT = "/Users/swilliams/Documents/workspace/training/DevSecOps_Academy_Syllabus.pptx"

W, H = Inches(13.333), Inches(7.5)
NAVY = RGBColor(0x10, 0x18, 0x20)
IVORY = RGBColor(0xF7, 0xF3, 0xEC)
INK = RGBColor(0x14, 0x1C, 0x27)
BODY = RGBColor(0x3A, 0x42, 0x4D)
MUTED = RGBColor(0x5C, 0x56, 0x4E)
GOLD = RGBColor(0x8A, 0x73, 0x50)
GOLD_LT = RGBColor(0xE4, 0xCB, 0x96)
RULE = RGBColor(0xE4, 0xD8, 0xC6)
CREAM = RGBColor(0xF3, 0xEE, 0xE6)
SOFT = RGBColor(0xC9, 0xC2, 0xB6)

SERIF = "Iowan Old Style"
SANS = "Avenir Next"
SANS_MED = "Avenir Next Medium"
SANS_DEM = "Avenir Next Demi Bold"


def rgb(value):
    return RGBColor.from_string(value)


def paint(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def rect(slide, l, t, w, h, fill, line=None, line_w=Pt(0.75)):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
        shape.line.width = line_w
    # drop the default shadow-less but also any residual
    sp_pr = shape.fill._xPr
    return shape


def hline(slide, l, t, w, color, weight=Pt(1)):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, weight)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape


def _latin(run, font):
    r_pr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        node = r_pr.find(qn(tag))
        if node is None:
            node = r_pr.makeelement(qn(tag), {})
            r_pr.append(node)
        node.set("typeface", font)
    return r_pr


def style_run(run, text, font, size, color, bold=False, italic=False, spc=None):
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    r_pr = _latin(run, font)
    if spc is not None:
        r_pr.set("spc", str(spc))


def add_text(slide, l, t, w, h, paragraphs, anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT):
    """paragraphs: list of dicts or list of run-lists.

    A paragraph dict: text, font, size, color, bold, italic, spc, space_before, space_after, leading
    Or a list of those dicts for mixed runs.
    """
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.margin_left = Emu(0)
    tf.margin_right = Emu(0)
    tf.margin_top = Emu(0)
    tf.margin_bottom = Emu(0)
    tf.paragraphs[0].alignment = align
    try:
        tf._txBody.bodyPr.set("anchor", {MSO_ANCHOR.TOP: "t", MSO_ANCHOR.MIDDLE: "ctr", MSO_ANCHOR.BOTTOM: "b"}[anchor])
    except Exception:
        pass

    for i, para in enumerate(paragraphs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        runs = para if isinstance(para, list) else [para]
        if runs:
            p.space_before = Pt(runs[0].get("before", 0))
            p.space_after = Pt(runs[0].get("after", 0))
            if runs[0].get("leading"):
                p.line_spacing = runs[0]["leading"]
        for spec in runs:
            run = p.add_run()
            style_run(
                run,
                spec.get("text", ""),
                spec.get("font", SANS),
                spec.get("size", 14),
                spec.get("color", INK),
                spec.get("bold", False),
                spec.get("italic", False),
                spec.get("spc"),
            )
    return box


def kicker(slide, l, t, w, text, color=GOLD):
    add_text(slide, l, t, w, Inches(0.28), [{
        "text": text.upper(), "font": SANS_MED, "size": 11, "color": color, "spc": 280,
    }])


def footer(slide, page, section, dark=False):
    color = SOFT if dark else GOLD
    num = CREAM if dark else INK
    add_text(slide, Inches(0.55), Inches(7.08), Inches(1.2), Inches(0.28), [{
        "text": f"{page:02d}", "font": SERIF, "size": 12, "color": num,
    }])
    add_text(slide, Inches(7.4), Inches(7.08), Inches(5.4), Inches(0.28), [{
        "text": section.upper(), "font": SANS_MED, "size": 10, "color": color, "spc": 180,
    }], align=PP_ALIGN.RIGHT)


def runhead(slide, right):
    add_text(slide, Inches(0.55), Inches(0.32), Inches(6.2), Inches(0.26), [{
        "text": "DEVSECOPS ACADEMY", "font": SANS_MED, "size": 10, "color": GOLD, "spc": 220,
    }])
    add_text(slide, Inches(6.6), Inches(0.32), Inches(6.18), Inches(0.26), [{
        "text": right.upper(), "font": SANS_MED, "size": 10, "color": GOLD, "spc": 180,
    }], align=PP_ALIGN.RIGHT)
    hline(slide, Inches(0.55), Inches(0.64), Inches(12.23), RGBColor(0xD9, 0xC7, 0xA4), Pt(1.0))


def new(prs, color=IVORY):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    paint(slide, color)
    return slide


def cover(prs):
    s = new(prs, NAVY)
    frame = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.32), Inches(0.26), Inches(12.69), Inches(6.98))
    frame.fill.background()
    frame.line.color.rgb = RGBColor(0xD4, 0xB8, 0x7A)
    frame.line.width = Pt(0.75)

    kicker(s, Inches(0.72), Inches(0.55), Inches(8), "Cloud & DevOps Academy", GOLD_LT)
    hline(s, Inches(0.72), Inches(0.98), Inches(0.48), GOLD_LT, Pt(1.0))

    add_text(s, Inches(0.70), Inches(1.85), Inches(10.5), Inches(2.15), [
        {"text": "Cloud & DevOps", "font": SERIF, "size": 54, "color": CREAM, "leading": 0.9},
        {"text": "Engineering", "font": SERIF, "size": 54, "color": CREAM, "leading": 0.9},
    ])
    add_text(s, Inches(0.72), Inches(4.05), Inches(8), Inches(0.4), [{
        "text": "The DevSecOps Programme", "font": SERIF, "size": 22, "color": GOLD_LT, "italic": True,
    }])
    add_text(s, Inches(0.72), Inches(4.52), Inches(8.2), Inches(0.7), [{
        "text": "A sixteen-week professional syllabus. Infrastructure, delivery, and operations,\ntaught as one system and proved in a capstone.",
        "font": SANS, "size": 16, "color": RGBColor(0xD5, 0xD0, 0xC6), "leading": 1.15,
    }])

    kicker(s, Inches(0.72), Inches(5.55), Inches(3), "The path", GOLD_LT)
    add_text(s, Inches(0.72), Inches(5.84), Inches(8), Inches(0.7), [{
        "text": "Linux   ù   Git   ù   GitLab CI     AWS   ù   Terraform   ù   Ansible   ù   Docker\nKubernetes   ù   Argo CD     Prometheus   ù   Alertmanager   ù   Grafana   ù   ELK",
        "font": SANS, "size": 14, "color": CREAM, "leading": 1.25,
    }])
    hline(s, Inches(0.72), Inches(6.62), Inches(11.9), RGBColor(0x8A, 0x73, 0x50), Pt(0.75))
    meta = [("2026 syllabus", 0.72), ("Practical standard", 5.15), ("Capstone assessed", 9.35)]
    for text, x in meta:
        add_text(s, Inches(x), Inches(6.70), Inches(2.6), Inches(0.28), [{
            "text": text.upper(), "font": SANS_MED, "size": 11, "color": SOFT, "spc": 140,
        }])


def programme(prs):
    s = new(prs)
    runhead(s, "The DevSecOps Programme")
    add_text(s, Inches(0.55), Inches(0.82), Inches(10), Inches(0.55), [{
        "text": "The programme", "font": SERIF, "size": 36, "color": INK,
    }])
    add_text(s, Inches(0.55), Inches(1.45), Inches(11.6), Inches(0.7), [{
        "text": "Graduates leave with a platform they can defend: infrastructure in code, a delivery pipeline,\na Kubernetes release path, and the signals required to operate it.",
        "font": SANS, "size": 16, "color": BODY, "leading": 1.15,
    }])

    stats = [
        ("16", "Weeks, one continuous build"),
        ("3", "Instructor-led sessions each week"),
        ("70%", "Labs, projects, and review"),
        ("1", "Capstone, demonstrated live"),
    ]
    hline(s, Inches(0.55), Inches(2.35), Inches(12.23), RULE, Pt(1))
    for i, (n, label) in enumerate(stats):
        x = 0.55 + i * 3.06
        if i:
            hline(s, Inches(x), Inches(2.48), Pt(1), RULE, Inches(0.85))
        add_text(s, Inches(x + 0.18), Inches(2.46), Inches(2.7), Inches(0.48), [{
            "text": n, "font": SERIF, "size": 28, "color": INK,
        }])
        add_text(s, Inches(x + 0.18), Inches(2.96), Inches(2.7), Inches(0.36), [{
            "text": label, "font": SANS, "size": 13, "color": MUTED,
        }])
    hline(s, Inches(0.55), Inches(3.48), Inches(12.23), RULE, Pt(1))

    kicker(s, Inches(0.55), Inches(3.7), Inches(4), "Who it is for")
    add_text(s, Inches(0.55), Inches(4.05), Inches(12.1), Inches(0.85), [{
        "text": "Beginners with sound computer skills, IT students, developers, and engineers moving into cloud and DevOps.\nBasic Linux is useful. A short orientation is provided before Week 1. No prior cloud account is required.",
        "font": SANS, "size": 15, "color": BODY, "leading": 1.2,
    }])
    add_text(s, Inches(0.55), Inches(5.15), Inches(12.1), Inches(1.5), [{
        "text": "Tools are taught in the order a system is built. The machine and the repository come first. Then the cloud\naccount, the container, the pipeline, the cluster, and the signals that show whether it is healthy.",
        "font": SANS, "size": 15, "color": BODY, "leading": 1.2,
    }])
    footer(s, 2, "Syllabus")


def outcomes(prs):
    s = new(prs)
    runhead(s, "The DevSecOps Programme")
    add_text(s, Inches(0.55), Inches(0.78), Inches(10), Inches(0.5), [{
        "text": "What a graduate can do", "font": SERIF, "size": 34, "color": INK,
    }])
    items = [
        ("01", "Administer Linux and troubleshoot networking, access, and services."),
        ("02", "Collaborate in Git and GitLab, and ship work through merge requests."),
        ("03", "Build, run, and harden Docker images and Compose applications."),
        ("04", "Provision AWS with least privilege and express it in Terraform."),
        ("05", "Configure systems with Ansible and keep those runs idempotent."),
        ("06", "Design GitLab CI pipelines that test, scan, publish, and roll back."),
        ("07", "Deploy on Kubernetes and release with GitOps through Argo CD."),
        ("08", "Monitor with Prometheus, Alertmanager, and Grafana, and investigate with ELK."),
    ]
    top = 1.55
    row_h = 0.62
    for i, (n, text) in enumerate(items):
        col = i // 4
        row = i % 4
        x = 0.55 + col * 6.35
        y = top + row * row_h
        hline(s, Inches(x), Inches(y), Inches(5.9), RULE, Pt(1))
        add_text(s, Inches(x), Inches(y + 0.08), Inches(0.48), Inches(0.36), [{
            "text": n, "font": SERIF, "size": 16, "color": GOLD,
        }])
        add_text(s, Inches(x + 0.5), Inches(y + 0.08), Inches(5.35), Inches(0.46), [{
            "text": text, "font": SANS, "size": 14, "color": INK,
        }])
    footer(s, 3, "Outcomes")


def arc(prs):
    s = new(prs)
    runhead(s, "The DevSecOps Programme")
    add_text(s, Inches(0.55), Inches(0.76), Inches(8), Inches(0.48), [{
        "text": "Sixteen weeks", "font": SERIF, "size": 34, "color": INK,
    }])
    months = [
        ("01", "Foundations", "Linux, the network, Git, and a first AWS account used with care.",
         [("01", "Linux and the shell"), ("02", "Networking and SSH"), ("03", "Git and GitLab"), ("04", "AWS fundamentals")]),
        ("02", "Provision", "Containers, a resilient AWS design, Terraform, and Ansible.",
         [("05", "Docker"), ("06", "AWS architecture"), ("07", "Terraform"), ("08", "Ansible")]),
        ("03", "Delivery", "GitLab CI, Kubernetes, Helm, and security inside the pipeline.",
         [("09", "GitLab CI"), ("10", "Kubernetes"), ("11", "Cluster operations"), ("12", "DevSecOps")]),
        ("04", "Operations", "Prometheus, Alertmanager, Grafana, ELK, Argo CD, and the capstone.",
         [("13", "Metrics and alerts"), ("14", "Logging with ELK"), ("15", "Argo CD"), ("16", "Capstone")]),
    ]
    y = 1.45
    for idx, (num, name, summary, weeks) in enumerate(months):
        hline(s, Inches(0.55), Inches(y), Inches(12.23), RULE, Pt(1))
        add_text(s, Inches(0.55), Inches(y + 0.08), Inches(2.3), Inches(0.22), [{
            "text": f"MONTH {num}", "font": SANS_MED, "size": 10, "color": GOLD, "spc": 160,
        }])
        add_text(s, Inches(0.55), Inches(y + 0.28), Inches(2.5), Inches(0.38), [{
            "text": name, "font": SERIF, "size": 22, "color": INK,
        }])
        add_text(s, Inches(0.55), Inches(y + 0.68), Inches(3.15), Inches(0.55), [{
            "text": summary, "font": SANS, "size": 12, "color": MUTED,
        }])
        for j, (wn, title) in enumerate(weeks):
            x = 3.9 + j * 2.2
            add_text(s, Inches(x), Inches(y + 0.22), Inches(2.05), Inches(0.24), [{
                "text": wn, "font": SERIF, "size": 13, "color": GOLD,
            }])
            add_text(s, Inches(x), Inches(y + 0.46), Inches(2.05), Inches(0.55), [{
                "text": title, "font": SANS_DEM, "size": 13, "color": INK,
            }])
        y += 1.32
    footer(s, 4, "Sixteen weeks")


def month_slide(prs, page, kicker_text, title, subtitle, weeks):
    s = new(prs)
    runhead(s, kicker_text)
    add_text(s, Inches(0.55), Inches(0.74), Inches(8), Inches(0.22), [{
        "text": subtitle.upper(), "font": SANS_MED, "size": 11, "color": GOLD, "spc": 200,
    }])
    add_text(s, Inches(0.55), Inches(0.96), Inches(8), Inches(0.5), [{
        "text": title, "font": SERIF, "size": 32, "color": INK,
    }])
    # 2 x 2
    positions = [(0.55, 1.7), (6.85, 1.7), (0.55, 4.28), (6.85, 4.28)]
    for (n, name, skills, evidence), (x, y) in zip(weeks, positions):
        hline(s, Inches(x), Inches(y), Inches(5.9), RGBColor(0xD9, 0xC7, 0xA4), Pt(1.25))
        add_text(s, Inches(x), Inches(y + 0.1), Inches(0.55), Inches(0.32), [{
            "text": n, "font": SERIF, "size": 18, "color": GOLD,
        }])
        add_text(s, Inches(x + 0.55), Inches(y + 0.1), Inches(5.3), Inches(0.36), [{
            "text": name, "font": SERIF, "size": 18, "color": INK,
        }])
        add_text(s, Inches(x), Inches(y + 0.52), Inches(5.9), Inches(0.85), [{
            "text": skills, "font": SANS, "size": 13, "color": BODY,
        }])
        add_text(s, Inches(x), Inches(y + 1.4), Inches(5.9), Inches(0.7), [
            [
                {"text": "EVIDENCE  ", "font": SANS_MED, "size": 10, "color": GOLD, "spc": 120},
                {"text": evidence, "font": SANS, "size": 13, "color": INK},
            ]
        ])
    footer(s, page, title)


def capstone(prs):
    s = new(prs)
    runhead(s, "Week 16 ù Final evidence")
    add_text(s, Inches(0.55), Inches(0.74), Inches(8), Inches(0.48), [{
        "text": "The capstone", "font": SERIF, "size": 34, "color": INK,
    }])
    add_text(s, Inches(0.55), Inches(1.28), Inches(12.2), Inches(0.55), [{
        "text": "A production-style application ù task manager, inventory service, student portal, or API ù automated, secured, observed, and demonstrated. The standard is a system another engineer could operate from the runbook.",
        "font": SANS, "size": 14, "color": BODY,
    }])
    rows = [
        ("01", "Source", "GitLab, with a branching model and merge requests."),
        ("02", "Infrastructure", "Terraform on AWS. Least privilege, tags, and a cost note."),
        ("03", "Pipeline", "GitLab CI: validate, test, build, scan, and publish."),
        ("04", "Cluster", "Kubernetes with probes, limits, and RBAC. EKS if budget allows."),
        ("05", "Release", "Argo CD from Git, with promotion and rollback."),
        ("06", "Metrics", "Prometheus, Grafana, and Alertmanager routes that have fired."),
        ("07", "Logs", "ELK, and a Kibana view used for an investigation."),
        ("08", "Record", "Diagram, README, findings, runbook, and a live demo."),
    ]
    for i, (n, name, text) in enumerate(rows):
        col = 0 if i < 4 else 1
        row = i if i < 4 else i - 4
        x = 0.55 + col * 6.4
        y = 2.05 + row * 1.12
        hline(s, Inches(x), Inches(y), Inches(6.05), RULE, Pt(1))
        add_text(s, Inches(x), Inches(y + 0.1), Inches(0.45), Inches(0.3), [{
            "text": n, "font": SERIF, "size": 16, "color": GOLD,
        }])
        add_text(s, Inches(x + 0.5), Inches(y + 0.08), Inches(5.4), Inches(0.3), [{
            "text": name, "font": SANS_DEM, "size": 15, "color": INK,
        }])
        add_text(s, Inches(x + 0.5), Inches(y + 0.4), Inches(5.4), Inches(0.5), [{
            "text": text, "font": SANS, "size": 13, "color": BODY,
        }])
    footer(s, 9, "Capstone")


def assessment(prs):
    s = new(prs)
    runhead(s, "The DevSecOps Programme")
    add_text(s, Inches(0.55), Inches(0.74), Inches(8), Inches(0.46), [{
        "text": "Assessment", "font": SERIF, "size": 34, "color": INK,
    }])
    add_text(s, Inches(0.55), Inches(1.26), Inches(12), Inches(0.32), [{
        "text": "Grades follow finished work. A claim without a repository, a diagram, or a demonstration does not count.",
        "font": SANS, "size": 14, "color": BODY,
    }])
    headers = [(0.55, "Assessment", 5.4), (6.1, "Weight", 1.3), (7.5, "What is judged", 5.2)]
    y = 1.78
    for x, label, w in headers:
        add_text(s, Inches(x), Inches(y), Inches(w), Inches(0.26), [{
            "text": label.upper(), "font": SANS_MED, "size": 11, "color": GOLD, "spc": 140,
        }])
    hline(s, Inches(0.55), Inches(2.1), Inches(12.23), RGBColor(0xCB, 0xB8, 0x92), Pt(1.15))
    rows = [
        ("Weekly practical labs", "30%", "Completion, troubleshooting, and evidence from the week."),
        ("Topic projects and reviews", "20%", "Working deliverables, documentation, and Git hygiene."),
        ("Mid-programme practical", "15%", "Docker, AWS, Terraform, Ansible, and a CI/CD path."),
        ("Security and observability", "10%", "Pipeline checks, dashboards, logs, and alerts with a reason."),
        ("Final capstone", "25%", "Architecture, automation, security, reliability, cost, and the demo."),
    ]
    y = 2.18
    for name, weight, judged in rows:
        add_text(s, Inches(0.55), Inches(y), Inches(5.4), Inches(0.36), [{
            "text": name, "font": SANS_DEM, "size": 14, "color": INK,
        }])
        add_text(s, Inches(6.1), Inches(y - 0.02), Inches(1.2), Inches(0.36), [{
            "text": weight, "font": SERIF, "size": 18, "color": INK,
        }])
        add_text(s, Inches(7.5), Inches(y), Inches(5.2), Inches(0.4), [{
            "text": judged, "font": SANS, "size": 13, "color": BODY,
        }])
        y += 0.48
        hline(s, Inches(0.55), Inches(y - 0.08), Inches(12.23), RULE, Pt(0.75))

    kicker(s, Inches(0.55), Inches(4.7), Inches(4), "The shape of a week")
    sessions = [
        ("Session 1", "Concepts", "The idea, a worked demonstration, and the decisions the lab will force."),
        ("Session 2", "Guided lab", "Instructor-led practice, with checkpoints where things are expected to break."),
        ("Session 3", "Workshop", "Project work, a review of someone elseùs change, and a short check."),
        ("Between", "Independent", "Two to four hours for practice, documentation, and the portfolio."),
    ]
    for i, (k, title, body) in enumerate(sessions):
        x = 0.55 + i * 3.15
        add_text(s, Inches(x), Inches(5.05), Inches(2.95), Inches(0.22), [{
            "text": k.upper(), "font": SANS_MED, "size": 10, "color": GOLD, "spc": 120,
        }])
        add_text(s, Inches(x), Inches(5.28), Inches(2.95), Inches(0.3), [{
            "text": title, "font": SERIF, "size": 16, "color": INK,
        }])
        add_text(s, Inches(x), Inches(5.62), Inches(2.95), Inches(0.85), [{
            "text": body, "font": SANS, "size": 12, "color": BODY,
        }])
    footer(s, 10, "Assessment")


def studio(prs):
    s = new(prs)
    runhead(s, "The DevSecOps Programme")
    add_text(s, Inches(0.55), Inches(0.74), Inches(8), Inches(0.46), [{
        "text": "Studio standard", "font": SERIF, "size": 34, "color": INK,
    }])
    rules = [
        "16 GB of memory is the recommendation once local Kubernetes begins. 8 GB is the minimum.",
        "Linux via a virtual machine, WSL2, or a dedicated machine. GitLab and a container registry are provided.",
        "AWS learner accounts use MFA, budget alerts, least privilege, and teardown at the end of every lab.",
        "Kubernetes exercises run on kind or Minikube. EKS is reserved for when the budget allows.",
        "Credentials never go into Git. Labs use temporary credentials and an approved secrets mechanism.",
    ]
    tools = [
        ("Systems", "Linux, Bash, SSH, and the usual network tools."),
        ("Collaboration", "Git and GitLab, including issues and merge requests."),
        ("Cloud", "IAM, VPC, EC2, S3, Lambda, CloudWatch, load balancing, EKS as a concept."),
        ("Automation", "Terraform and Ansible."),
        ("Delivery", "Docker, GitLab CI, Kubernetes, Helm, and Argo CD."),
        ("Observability", "Prometheus, Alertmanager, Grafana, and the ELK stack."),
    ]
    add_text(s, Inches(0.55), Inches(1.35), Inches(6), Inches(0.26), [{
        "text": "PRACTICE", "font": SANS_MED, "size": 11, "color": GOLD, "spc": 200,
    }])
    add_text(s, Inches(7.15), Inches(1.35), Inches(5.5), Inches(0.26), [{
        "text": "TOOL STACK", "font": SANS_MED, "size": 11, "color": GOLD, "spc": 200,
    }])
    y = 1.72
    for i, rule in enumerate(rules):
        add_text(s, Inches(0.55), Inches(y), Inches(0.4), Inches(0.3), [{
            "text": f"{i+1:02d}", "font": SERIF, "size": 14, "color": GOLD,
        }])
        add_text(s, Inches(1.05), Inches(y), Inches(5.7), Inches(0.7), [{
            "text": rule, "font": SANS, "size": 13, "color": INK,
        }])
        y += 0.95
    y = 1.72
    for name, text in tools:
        add_text(s, Inches(7.15), Inches(y), Inches(5.5), Inches(0.24), [{
            "text": name.upper(), "font": SANS_MED, "size": 11, "color": GOLD, "spc": 100,
        }])
        add_text(s, Inches(7.15), Inches(y + 0.24), Inches(5.5), Inches(0.4), [{
            "text": text, "font": SANS, "size": 13, "color": INK,
        }])
        y += 0.78
    footer(s, 11, "Studio standard")


def close(prs):
    s = new(prs, NAVY)
    frame = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.32), Inches(0.26), Inches(12.69), Inches(6.98))
    frame.fill.background()
    frame.line.color.rgb = RGBColor(0xD4, 0xB8, 0x7A)
    frame.line.width = Pt(0.75)
    kicker(s, Inches(0.72), Inches(1.7), Inches(8), "Cloud & DevOps Academy", GOLD_LT)
    add_text(s, Inches(0.70), Inches(2.15), Inches(11), Inches(1.6), [
        {"text": "One system.", "font": SERIF, "size": 48, "color": CREAM, "leading": 0.95},
        {"text": "Sixteen weeks.", "font": SERIF, "size": 48, "color": CREAM, "leading": 0.95},
    ])
    add_text(s, Inches(0.72), Inches(4.35), Inches(10), Inches(0.7), [{
        "text": "Linux, AWS, Git, Terraform, Docker, GitLab CI, Kubernetes,\nArgo CD, and observability ù then a platform that proves it.",
        "font": SANS, "size": 18, "color": RGBColor(0xD5, 0xD0, 0xC6), "leading": 1.2,
    }])
    add_text(s, Inches(0.72), Inches(6.15), Inches(8), Inches(0.35), [{
        "text": "The DevSecOps Programme", "font": SERIF, "size": 16, "color": GOLD_LT, "italic": True,
    }])


def build():
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H
    prs.core_properties.title = "Cloud & DevOps Academy ù The DevSecOps Programme"
    prs.core_properties.subject = "Sixteen-week professional syllabus"
    prs.core_properties.category = "Syllabus"

    cover(prs)
    programme(prs)
    outcomes(prs)
    arc(prs)
    month_slide(prs, 5, "Month 01 ù Foundations", "Foundations", "Weeks 01 ù 04", [
        ("01", "Linux and the shell",
         "Filesystem, users, permissions, systemd, logs, and Bash ù pipes, grep, find, awk, sed, and scripts.",
         "A server baseline with least privilege, and a scripted health report."),
        ("02", "Networking and SSH",
         "Addressing, DNS, firewalls, TLS, SSH keys, and the tools used to find a broken service.",
         "Key-based access between two Linux hosts, and a troubleshooting runbook."),
        ("03", "Git and GitLab",
         "Branches, merges, conflicts, merge requests, protected branches, and tagged releases.",
         "A team project with reviews, and release notes for a small application."),
        ("04", "AWS fundamentals",
         "IAM, MFA, EC2, EBS, S3, VPC, security groups, budgets, tagging, and CloudTrail.",
         "A locked-down EC2 web server and an S3 site, with security and cost notes."),
    ])
    month_slide(prs, 6, "Month 02 ù Provision", "Provision", "Weeks 05 ù 08", [
        ("05", "Docker",
         "Images, multi-stage builds, Compose, volumes, networks, health checks, and image scanning.",
         "An application and database in Compose, with an optimized multi-stage image."),
        ("06", "AWS architecture",
         "Auto Scaling, load balancing, Lambda, CloudWatch, private subnets, and secrets management.",
         "A highly available sample behind a load balancer, and a Lambda with an alarm."),
        ("07", "Terraform",
         "Workflow, state, modules, workspaces, drift, tagging, secrets, and safe teardown.",
         "VPC, compute, and storage in modules, applied for two environments, then destroyed."),
        ("08", "Ansible",
         "Playbooks, roles, templates, Vault, and idempotent configuration after Terraform.",
         "An Nginx role, then a Terraform-built host configured with a clean second run."),
    ])
    month_slide(prs, 7, "Month 03 ù Delivery", "Delivery", "Weeks 09 ù 12", [
        ("09", "GitLab CI",
         "Stages, runners, rules, image publish, environments, release strategies, and rollback.",
         "A pipeline that tests and pushes on merge, then a deploy and a demonstrated rollback."),
        ("10", "Kubernetes",
         "Pods, Deployments, Services, ConfigMaps, probes, requests, and a local kind or Minikube cluster.",
         "A multi-replica application, then a broken rollout diagnosed from events and logs."),
        ("11", "Cluster operations",
         "Ingress, persistent storage, autoscaling, Helm, security contexts, and RBAC.",
         "A Helm chart, plus a stateful workload running as a restricted service account."),
        ("12", "DevSecOps",
         "Secret, dependency, image, and Terraform checks inside the GitLab pipeline.",
         "Findings fixed, and an image that runs as non-root with versions pinned."),
    ])
    month_slide(prs, 8, "Month 04 ù Operations", "Operations", "Weeks 13 ù 16", [
        ("13", "Prometheus and alerts",
         "PromQL, Grafana dashboards, and Alertmanager routing, grouping, inhibition, and silences.",
         "A dashboard and tested routes for CPU, disk, and downtime."),
        ("14", "Logging with ELK",
         "Structured logs, Elasticsearch, Kibana, and ingestion with Filebeat, Logstash, or Elastic Agent.",
         "A Kibana dashboard and a written investigation that names the cause."),
        ("15", "Argo CD",
         "Desired state, sync, health, drift, environment promotion, and rollback from Git.",
         "A release deployed from Git, drift corrected, then promoted and rolled back."),
        ("16", "Capstone",
         "Architecture, reliability, security, cost, an incident exercise, and the handover.",
         "The full platform, submitted with its runbook and demonstrated live."),
    ])
    capstone(prs)
    assessment(prs)
    studio(prs)
    close(prs)
    prs.save(OUT)
    print(OUT)


if __name__ == "__main__":
    build()
