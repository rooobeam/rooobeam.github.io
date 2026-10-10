"""绘制正式排位前的外观示意；只显示4个姓名样式，不分配完整100人格位。"""

import json
import os
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
os.environ.setdefault("MPLCONFIGDIR", str(HERE / ".mplconfig"))

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
from matplotlib.patches import Polygon, Rectangle
import numpy as np
from scipy.optimize import Bounds, LinearConstraint, milp
from scipy.sparse import lil_matrix


CENTER = (474, 650)
X = range(463, 486)
Y = range(637, 663)
FLAGS = {(x, y) for x in (466, 471, 477, 482) for y in (641, 647, 653, 659)}
BEAR = {(x, y) for x in range(473, 476) for y in range(649, 652)}
SAMPLE_ANCHORS = [(473, 652), (471, 649), (476, 650), (474, 647)]
COLORS = ["#c96961", "#70a993", "#cdbd82", "#9892b9"]
FONT = FontProperties(fname="C:/Windows/Fonts/msyh.ttc")
BOLD = FontProperties(fname="C:/Windows/Fonts/msyhbd.ttc")


def geometry_precheck():
    cells = [(x, y) for x in X for y in Y]
    index = {cell: i for i, cell in enumerate(cells)}
    candidates = []
    for x in range(463, 485):
        for y in range(637, 662):
            footprint = {(x, y), (x + 1, y), (x, y + 1), (x + 1, y + 1)}
            if not footprint.isdisjoint(BEAR | FLAGS):
                continue
            # 保守采用整座城堡落在同一面旗子的7×7覆盖内。
            if not any(all(max(abs(px - fx), abs(py - fy)) <= 3 for px, py in footprint) for fx, fy in FLAGS):
                continue
            candidates.append(((x, y), footprint))
    matrix = lil_matrix((len(cells), len(candidates)))
    lower = np.zeros(len(candidates))
    for j, (anchor, footprint) in enumerate(candidates):
        for point in footprint:
            matrix[index[point], j] = 1
        if anchor in SAMPLE_ANCHORS:
            lower[j] = 1
    result = milp(
        -np.ones(len(candidates)),
        integrality=np.ones(len(candidates)),
        bounds=Bounds(lower, 1),
        constraints=LinearConstraint(matrix.tocsr(), 0, 1),
        options={"time_limit": 20, "mip_rel_gap": 0},
    )
    if result.x is None:
        raise RuntimeError("旗阵预检未找到可行布局。")
    occupied = set()
    chosen = []
    for selected, (anchor, footprint) in zip(result.x, candidates):
        if selected > 0.5:
            assert occupied.isdisjoint(footprint)
            occupied.update(footprint)
            chosen.append(anchor)
    assert set(SAMPLE_ANCHORS) <= set(chosen)
    assert len(chosen) >= 100
    assert {(2 * CENTER[0] - x, y) for x, y in FLAGS} == FLAGS
    assert {(x, 2 * CENTER[1] - y) for x, y in FLAGS} == FLAGS
    covered = {p for p in cells if any(max(abs(p[0] - x), abs(p[1] - y)) <= 3 for x, y in FLAGS)}
    summary = {
        "status": "仅供制作前核对；不是正式100人格位分配",
        "range_inclusive": {"x": [463, 485], "y": [637, 662]},
        "grid_size": [23, 26],
        "bear_center": list(CENTER),
        "bear_cells": sorted(BEAR),
        "flag_positions": sorted(FLAGS),
        "flag_count": len(FLAGS),
        "flag_reflection_symmetry_x_and_y": True,
        "castle_coverage_rule": "2×2全部四格均落在至少一面旗子的7×7范围内",
        "covered_cell_count": len(covered),
        "feasible_capacity_with_four_style_samples": len(chosen),
        "solver_proved_maximum_for_this_fixed_flag_array": result.status == 0,
        "capacity_satisfies_100": True,
        "final_layout_approved": False,
        "sample_names_are_final_positions": False,
    }
    (HERE / "proposal_geometry.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    return covered, summary


def iso(x, y, z=0):
    u, v = x - CENTER[0], y - CENTER[1]
    return (u - v, 0.53 * (u + v) + z)


def face(ax, x, y, width, height, color, z=0, edge="#ffffff", line=0.3, order=1):
    points = [iso(x, y, z), iso(x + width, y, z), iso(x + width, y + height, z), iso(x, y + height, z)]
    ax.add_patch(Polygon(points, closed=True, facecolor=color, edgecolor=edge, linewidth=line, zorder=order))


def block(ax, x, y, width, height, color, z, side, order):
    for p1, p2 in [((x, y), (x + width, y)), ((x, y), (x, y + height))]:
        points = [iso(*p1), iso(*p2), iso(*p2, z), iso(*p1, z)]
        ax.add_patch(Polygon(points, closed=True, facecolor=side, edgecolor=side, zorder=order))
    face(ax, x, y, width, height, color, z=z, edge=side, line=0.65, order=order + 0.1)


def write(ax, x, y, text, size=10, bold=False, **kwargs):
    ax.text(x, y, text, fontproperties=BOLD if bold else FONT, fontsize=size, **kwargs)


def main():
    roster = json.loads((HERE / "roster_review.json").read_text(encoding="utf-8"))
    assert roster["summary"]["active_member_count"] == 100
    assert not roster["needs_power_or_rank"]
    samples = [m["name"] for m in roster["active_members"][:4]]
    covered, summary = geometry_precheck()
    fig = plt.figure(figsize=(17.5, 10.5), facecolor="#f7f5f1")
    fig.text(0.035, 0.938, "熊坑围城 · 预计布局", fontproperties=BOLD, fontsize=26, color="#263743")
    fig.text(0.036, 0.898, "100 人  /  中心熊坑 3×3  /  城堡 2×2  /  对称旗阵 16 面", fontproperties=FONT, fontsize=13, color="#53606a")
    fig.text(0.948, 0.942, "制作前核对", ha="right", fontproperties=BOLD, fontsize=13, color="#9b5948", bbox={"boxstyle": "round,pad=0.6", "facecolor": "#f0e1d5", "edgecolor": "none"})

    ax = fig.add_axes([0.025, 0.18, 0.64, 0.66])
    ax.set_aspect("equal")
    ax.set_xlim(-27, 27)
    ax.set_ylim(-15.5, 15.5)
    ax.axis("off")
    for x in X:
        for y in Y:
            radius = max(abs(x - CENTER[0]), abs(y - CENTER[1]))
            zone = 0 if radius <= 4 else 1 if radius <= 7 else 2 if radius <= 10 else 3
            color = COLORS[zone] if (x, y) in covered else "#e5e3df"
            face(ax, x - 0.5, y - 0.5, 1, 1, color, edge="#f1eeeb", line=0.3)
    # 色带只表示内外分区，四座姓名块仅供检查字号与风格。
    for (x, y), name in zip(SAMPLE_ANCHORS, samples):
        block(ax, x - 0.5, y - 0.5, 2, 2, "#b34e49", 0.12, "#8c4645", 10)
        u, v = iso(x + 0.5, y + 0.5, 0.16)
        write(ax, u, v, name, size=8.2, bold=True, color="white", ha="center", va="center", zorder=50)
    block(ax, 472.5, 648.5, 3, 3, "#ead083", 0.12, "#bda35f", 11)
    write(ax, *iso(474, 650, 0.12), "熊坑\n3×3", size=12, bold=True, color="#5d512f", ha="center", va="center", zorder=40)
    for x, y in sorted(FLAGS, key=lambda p: p[0] + p[1], reverse=True):
        block(ax, x - 0.5, y - 0.5, 1, 1, "#487bac", 0.85, "#315779", 20)
    for (x, y), label, align in [((463, 637), "(463,637)", "center"), ((485, 637), "(485,637)", "left"), ((463, 662), "(463,662)", "right"), ((485, 662), "(485,662)", "center")]:
        u, v = iso(x, y)
        offset = -1.15 if (x, y) == (463, 637) else 1.0 if (x, y) == (485, 662) else 0
        u += 1.2 if align == "left" else -1.2 if align == "right" else 0
        write(ax, u, v + offset, label, size=8.6, color="#697279", ha=align, va="center")

    ax2 = fig.add_axes([0.713, 0.258, 0.252, 0.57])
    ax2.set_aspect("equal")
    ax2.set_facecolor("#fefdfb")
    for x, y in sorted(FLAGS):
        ax2.add_patch(Rectangle((x - 3.5, y - 3.5), 7, 7, facecolor="#5881a2", edgecolor="#6e91ad", alpha=0.055, linewidth=0.65))
        ax2.add_patch(Rectangle((x - 3.5, y - 3.5), 7, 7, fill=False, edgecolor="#84a1b7", linewidth=0.7, linestyle=(0, (3, 3))))
    ax2.axvline(474, color="#b4a99b", linewidth=0.9, linestyle=(0, (5, 4)), zorder=2)
    ax2.axhline(650, color="#b4a99b", linewidth=0.9, linestyle=(0, (5, 4)), zorder=2)
    for i, (x, y) in enumerate(sorted(FLAGS), 1):
        ax2.add_patch(Rectangle((x - 0.5, y - 0.5), 1, 1, facecolor="#386a98", zorder=5))
        write(ax2, x, y, str(i), size=7.2, color="white", ha="center", va="center", zorder=6)
    ax2.add_patch(Rectangle((472.5, 648.5), 3, 3, facecolor="#ead083", edgecolor="#bda35f", zorder=5))
    write(ax2, 474, 650, "熊坑", size=10, bold=True, ha="center", va="center", color="#5d512f", zorder=6)
    ax2.set_xlim(462.5, 485.5)
    ax2.set_ylim(636.5, 662.5)
    ax2.set_xticks([463, 466, 471, 474, 477, 482, 485])
    ax2.set_yticks([637, 641, 647, 650, 653, 659, 662])
    ax2.set_xticks(np.arange(462.5, 486, 1), minor=True)
    ax2.set_yticks(np.arange(636.5, 663, 1), minor=True)
    ax2.grid(which="minor", color="#d9dedf", linewidth=0.4, zorder=0)
    ax2.tick_params(which="minor", length=0)
    ax2.tick_params(which="major", labelsize=8, color="#84919a")
    for spine in ax2.spines.values():
        spine.set_color("#91a0aa")
    ax2.set_title("旗阵与覆盖范围 · 俯视", fontproperties=BOLD, fontsize=14, color="#263743", pad=14)
    ax2.set_xlabel("地图 X 坐标", fontproperties=FONT, fontsize=10, labelpad=9, color="#53606a")
    ax2.set_ylabel("地图 Y 坐标", fontproperties=FONT, fontsize=10, labelpad=7, color="#53606a")
    fig.text(0.709, 0.19, "蓝色方块：旗子本体 1×1\n蓝色虚线：每面旗子的 7×7 覆盖\n中心取 (474,650)，旗阵沿两轴对称", fontproperties=FONT, fontsize=10.4, color="#53606a", linespacing=1.65)

    for i, (color, label) in enumerate(zip(COLORS, ["内圈优先", "中圈", "外圈", "边缘补位"])):
        xpos = 0.038 + i * 0.135
        fig.patches.append(Rectangle((xpos, 0.132), 0.012, 0.017, facecolor=color, transform=fig.transFigure))
        fig.text(xpos + 0.018, 0.133, label, fontproperties=FONT, fontsize=11, color="#46545e")
    fig.text(0.036, 0.086, "色带表示预计排位分区；图中 4 个姓名仅演示城堡样式，完整 100 人格位将在确认后求解。", fontproperties=FONT, fontsize=11, color="#606970")
    fig.text(0.036, 0.053, "排位规则：截图排名优先，其余按战力降序；城堡完整受旗子覆盖，与熊坑、旗子和其他城堡均不重叠。", fontproperties=FONT, fontsize=11, color="#606970")
    output = HERE / "预计布局_待确认.png"
    fig.savefig(output, dpi=170, facecolor=fig.get_facecolor())
    plt.close(fig)
    print(json.dumps({"image": output.name, "flag_count": len(FLAGS), "feasible_capacity": summary["feasible_capacity_with_four_style_samples"], "target_count": 100}, ensure_ascii=False))


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    main()
