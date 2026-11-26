"""Vector redraws of the hand-drawn figures in the Solid State panel PDFs (Parts 1-4)."""
import random

from .docxkit import Figure

RED, GREEN, BLUE, MAG, GREY = "C00000", "2E7D32", "1F3864", "AD1457", "404040"


def classification_of_matter():
    f = Figure(440, 135)
    f.rect(175, 4, 90, 30, color=BLUE, width=1.25, text="पदार्थ", size=13, bold=True)
    f.line(220, 34, 220, 52, width=1.25)
    f.line(95, 52, 345, 52, width=1.25)
    f.arrow(95, 52, 95, 74, width=1.25)
    f.arrow(345, 52, 345, 74, width=1.25)
    f.rect(10, 76, 170, 50, color=GREEN, text="भौतिक वर्गीकरण\n(Physical Classification)", size=11)
    f.rect(260, 76, 170, 50, color=GREEN, text="रासायनिक वर्गीकरण\n(Chemical Classification)", size=11)
    return f


def crystalline_lattice():
    """Regular rhombic lattice: particles repeat at fixed distances (long-range order)."""
    f = Figure(250, 190)
    n, dx, dy, ox, oy = 5, 34, 30, 40, 20
    pts = {}
    for i in range(n):
        for j in range(n):
            pts[i, j] = (ox + i * dx + j * dx * 0.6, oy + j * dy)
    for (i, j), (x, y) in pts.items():
        if i + 1 < n:
            f.line(x, y, *pts[i + 1, j], color=MAG, width=0.75)
        if j + 1 < n:
            f.line(x, y, *pts[i, j + 1], color=MAG, width=0.75)
    for (i, j), (x, y) in pts.items():
        f.circle(x, y, 4.5, color=RED, width=1.0, fill="FFFFFF")
        if i + 1 < n and j + 1 < n:  # small particle at the centre of each cell
            cx = (x + pts[i + 1, j + 1][0]) / 2
            cy = (y + pts[i + 1, j + 1][1]) / 2
            f.circle(cx, cy, 2, color=GREEN, width=0.75, fill="FFFFFF")
    return f


def amorphous_network():
    """Irregular network: order only over short distances."""
    rnd = random.Random(7)
    f = Figure(250, 190)
    pts = [(40, 60), (75, 30), (120, 45), (160, 25), (205, 55), (60, 105), (105, 90),
           (150, 85), (195, 110), (45, 150), (95, 140), (140, 135), (180, 160), (220, 145),
           (120, 175)]
    edges = [(0, 1), (1, 2), (2, 3), (3, 4), (0, 5), (5, 6), (6, 2), (6, 7), (7, 3), (7, 8),
             (8, 4), (5, 9), (9, 10), (10, 6), (10, 11), (11, 7), (11, 12), (12, 8), (12, 13),
             (13, 8), (11, 14), (14, 10)]
    for a, b in edges:
        (x1, y1), (x2, y2) = pts[a], pts[b]
        f.line(x1, y1, x2, y2, color=RED, width=0.75)
        steps = max(2, int(((x2 - x1) ** 2 + (y2 - y1) ** 2) ** 0.5 // 12))
        for k in range(1, steps):  # beads along each chain, slightly irregular
            t = k / steps
            f.circle(x1 + (x2 - x1) * t + rnd.uniform(-1.5, 1.5),
                     y1 + (y2 - y1) * t + rnd.uniform(-1.5, 1.5), 2.6, color=GREEN, width=0.75,
                     fill="FFFFFF")
    for x, y in pts:
        f.circle(x, y, 3.5, color=GREEN, width=1.0, fill="FFFFFF")
    return f


def anisotropy_arrangement():
    """Part 3: particle arrangement differs along different directions."""
    s, ox, oy = 0.5, 1170, 160
    f = Figure(270, 240)
    T = lambda x, y: ((x - ox) * s, (y - oy) * s)
    f.rect(0, 0, 268, 238, color="C55A11", width=1.0)
    for x in (1222, 1390, 1565):
        for y in (215, 333, 445, 548):
            f.circle(*T(x, y), 9, color=RED, width=1.25)
    for x, y in ((1300, 283), (1460, 287), (1645, 288), (1300, 405), (1460, 405), (1615, 405),
                 (1300, 510), (1460, 510), (1615, 522)):
        f.circle(*T(x, y), 3.5, color=GREEN, width=1.0)
    f.line(*T(1300, 195), *T(1300, 585), color=MAG, width=1.25)
    f.line(*T(1185, 580), *T(1665, 275), color=MAG, width=1.25)
    return f


def _ion(f, x, y, plus, r=9):
    f.circle(x, y, r, color="000000", width=1.0, fill="FFFFFF", text="+" if plus else "−",
             size=10, bold=True)


def ionic_crystal():
    """Part 3: ionic crystal - 3D cube of alternating ions 'or' a 2D alternating grid."""
    f = Figure(430, 175)
    # cube (front face + back face offset)
    a, ox, oy, bx, by = 85, 20, 60, 45, -40
    corners = {}
    for i in (0, 1):
        for j in (0, 1):
            for k in (0, 1):
                corners[i, j, k] = (ox + i * a + k * bx, oy + j * a + k * by + 30)
    for (i, j, k), p in corners.items():
        for d in ((1, 0, 0), (0, 1, 0), (0, 0, 1)):
            q = (i + d[0], j + d[1], k + d[2])
            if q in corners:
                f.line(*p, *corners[q], width=1.0)
    for (i, j, k), (x, y) in sorted(corners.items(), key=lambda kv: -kv[0][2]):
        _ion(f, x, y, (i + j + k) % 2 == 0)
    f.text(165, 75, 40, 24, "या", size=13)
    # 2D grid
    gx, gy, d = 230, 25, 42
    for r in range(4):
        for c in range(4):
            x, y = gx + c * d + 10, gy + r * d + 10
            if c < 3:
                f.line(x, y, x + d, y, color=MAG, width=1.0)
            if r < 3:
                f.line(x, y, x, y + d, color=MAG, width=1.0)
    for r in range(4):
        for c in range(4):
            _ion(f, gx + c * d + 10, gy + r * d + 10, (r + c) % 2 == 0)
    return f


def polar_hcl_chain():
    """Part 4: dipole-dipole attraction between polar H-Cl molecules."""
    f = Figure(430, 105)
    x0, y, step = 10, 40, 105
    for k in range(4):
        x = x0 + k * step
        f.text(x - 4, y - 32, 26, 18, "δ⁺", size=11, color=MAG)
        f.text(x + 30, y - 32, 26, 18, "δ⁻", size=11, color=MAG)
        f.text(x, y - 12, 66, 24, "H — Cl", size=12, bold=True, color=GREEN)
        if k < 3:
            f.line(x + 66, y, x + step, y, color=GREY, width=1.25, dash="sysDot")
    f.arrow(43, y + 12, 43, y + 34, width=1.0)
    f.text(0, y + 36, 170, 24, "(ध्रुवीय सहसंयोजक बंध)", size=11)
    return f
