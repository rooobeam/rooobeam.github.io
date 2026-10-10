"""每圈严格宽2格，允许旗子打断。求解几何，预览前两圈姓名供用户核对。"""

import csv
import json
import math
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
BOARD = {(x, y) for x in range(-11, 12) for y in range(-13, 13)}
BEAR = {(x, y) for x in range(-1, 2) for y in range(-1, 2)}
FONT = FontProperties(fname="C:/Windows/Fonts/msyh.ttc")
BOLD = FontProperties(fname="C:/Windows/Fonts/msyhbd.ttc")
PALETTE = {1: "#b9524f", 2: "#588b75", 3: "#beab6e", 4: "#8a91b1", 5: "#afa3bb"}


def footprint(anchor):
    x, y = anchor
    return {(x, y), (x + 1, y), (x, y + 1), (x + 1, y + 1)}


def castle_orbit(anchor, symmetry="rotate"):
    x, y = anchor
    if symmetry == "mirror":
        return {(x, y), (-x - 1, y), (x, -y - 1), (-x - 1, -y - 1)}
    return {(x, y), (-y - 1, x), (-x - 1, -y - 1), (y, -x - 1)}


def flag_array(inner, outer):
    values = (-outer, -inner, inner, outer)
    return {(x, y) for x in values for y in values}


def ring_index(anchor):
    bands = {max(abs(x), abs(y)) // 2 for x, y in footprint(anchor)}
    return next(iter(bands)) if len(bands) == 1 else None


def enumerate_groups(flags, ring, symmetry):
    valid = {}
    for x in range(-11, 11):
        for y in range(-13, 12):
            anchor = (x, y)
            cells = footprint(anchor)
            if ring_index(anchor) != ring or not cells.isdisjoint(flags | BEAR):
                continue
            owners = [f for f in flags if all(max(abs(px - f[0]), abs(py - f[1])) <= 3 for px, py in cells)]
            if owners:
                valid[anchor] = cells
    groups = {}
    for anchor in valid:
        orbit = castle_orbit(anchor, symmetry)
        if not orbit <= valid.keys():
            continue
        cells = set().union(*(valid[a] for a in orbit))
        if len(cells) == 4 * len(orbit):
            groups[tuple(sorted(orbit))] = cells
    return sorted(groups.items())


def solve_ring(flags, ring, symmetry="rotate", target=None):
    groups = enumerate_groups(flags, ring, symmetry)
    if not groups:
        if target:
            raise ValueError("当前圈无法容纳指定人数")
        return []
    cells = sorted(set().union(*(p for group, p in groups)))
    index = {p: i for i, p in enumerate(cells)}
    rows = len(cells) + (target is not None)
    matrix = lil_matrix((rows, len(groups)))
    lower = np.full(rows, -np.inf)
    upper = np.ones(rows)
    for j, (group, occupied) in enumerate(groups):
        for p in occupied:
            matrix[index[p], j] = 1
        if target is not None:
            matrix[-1, j] = len(group)
    if target is None:
        cost = -np.array([len(group) for group, occupied in groups])
    else:
        lower[-1] = upper[-1] = target
        cost = np.array([sum((x + 0.5) ** 2 + (y + 0.5) ** 2 for x, y in group) for group, occupied in groups])
    result = milp(
        cost,
        integrality=np.ones(len(groups)),
        bounds=Bounds(0, 1),
        constraints=LinearConstraint(matrix.tocsr(), lower, upper),
        options={"time_limit": 10, "mip_rel_gap": 0},
    )
    if not result.success:
        raise RuntimeError(f"第{ring}圈求解未得到最优证明：{result.message}")
    selected = []
    for value, (group, occupied) in zip(result.x, groups):
        if value > 0.5:
            selected.extend(group)
    return sorted(selected)


def verify(anchors, flags, counts):
    assert len(anchors) == 100 and len(set(anchors)) == 100
    assert {(y, -x) for x, y in flags} == flags
    assert {(-x, y) for x, y in flags} == flags
    assert {(x, -y) for x, y in flags} == flags
    assert set().union(*(castle_orbit(a) for a in anchors)) == set(anchors)
    occupied = set()
    for anchor in anchors:
        cells = footprint(anchor)
        assert cells <= BOARD
        assert cells.isdisjoint(BEAR | flags | occupied)
        occupied.update(cells)
        k = ring_index(anchor)
        assert k in range(1, 6)
        assert all(2 * k <= max(abs(x), abs(y)) <= 2 * k + 1 for x, y in cells)
        assert any(all(max(abs(px - fx), abs(py - fy)) <= 3 for px, py in cells) for fx, fy in flags)
    assert len(occupied) == 400
    assert sum(counts.values()) == 100
    for fx, fy in flags:
        assert (fx, fy) in BOARD and (fx, fy) not in BEAR
        assert {(x, y) for x in range(fx - 3, fx + 4) for y in range(fy - 3, fy + 4)} <= BOARD


def text(ax, x, y, value, size=11, bold=False, **kwargs):
    ax.text(x, y, value, fontproperties=BOLD if bold else FONT, fontsize=size, **kwargs)


def name_lines(name):
    if len(name) <= 5 or (name.isascii() and len(name) <= 12):
        return name
    split = (len(name) + 1) // 2
    return name[:split] + "\n" + name[split:]


def title(fig, heading, subtitle):
    fig.text(0.055, 0.95, heading, fontproperties=BOLD, fontsize=24, color="#263640")
    fig.text(0.055, 0.912, subtitle, fontproperties=FONT, fontsize=11.5, color="#5a6770")


def draw_top(rows, flags):
    fig = plt.figure(figsize=(13, 11.5), facecolor="#f7f5f1")
    title(fig, "前两圈 · 真实格位预览", "每圈宽 2 格｜允许旗子打断｜城堡旋转 90° 后格位重合｜旗子同时满足上下、左右镜像")
    ax = fig.add_axes([0.11, 0.18, 0.76, 0.68])
    ax.set_aspect("equal")
    for x in range(-5, 6):
        for y in range(-5, 6):
            ax.add_patch(Rectangle((x - 0.5, y - 0.5), 1, 1, facecolor="#fefdfb", edgecolor="#d9dddc", linewidth=0.45))
    for row in rows:
        if row["ring"] > 2:
            continue
        x, y = row["relative_anchor"]
        color = PALETTE[row["ring"]]
        ax.add_patch(Rectangle((x - 0.5, y - 0.5), 2, 2, facecolor=color, edgecolor="#344c50", linewidth=1.05, zorder=3))
        ax.plot([x + 0.5, x + 0.5], [y - 0.5, y + 1.5], color="white", alpha=0.22, linewidth=0.5, zorder=4)
        ax.plot([x - 0.5, x + 1.5], [y + 0.5, y + 0.5], color="white", alpha=0.22, linewidth=0.5, zorder=4)
        text(ax, x + 0.5, y + 0.5, name_lines(row["name"]), size=11.2, bold=True, ha="center", va="center", color="white", zorder=8)
    ax.add_patch(Rectangle((-1.5, -1.5), 3, 3, facecolor="#e4c777", edgecolor="#b69a52", linewidth=1.4, zorder=5))
    text(ax, 0, 0, "熊坑\n3×3", size=17, bold=True, ha="center", va="center", color="#574827", zorder=6)
    for x, y in flags:
        if max(abs(x), abs(y)) > 5:
            continue
        ax.add_patch(Rectangle((x - 0.5, y - 0.5), 1, 1, facecolor="#38688f", edgecolor="#244861", linewidth=1.2, zorder=7))
        text(ax, x, y, "旗", size=12, bold=True, ha="center", va="center", color="white", zorder=8)
    for half in (1.5, 3.5, 5.5):
        ax.add_patch(Rectangle((-half, -half), 2 * half, 2 * half, fill=False, edgecolor="#736c62", linewidth=1.3, linestyle=(0, (4, 3)), zorder=6))
    for y1, y2, label, color in [(1.5, 3.5, "第一圈\n宽2格", PALETTE[1]), (3.5, 5.5, "第二圈\n宽2格", PALETTE[2])]:
        ax.annotate("", xy=(6.1, y1), xytext=(6.1, y2), arrowprops={"arrowstyle": "|-|", "color": color, "linewidth": 1.4})
        text(ax, 6.35, (y1 + y2) / 2, label, size=10.4, color=color, ha="left", va="center")
    ticks = list(range(-5, 6))
    ax.set_xticks(ticks, [str(CENTER[0] + n) for n in ticks])
    ax.set_yticks(ticks, [str(CENTER[1] + n) for n in ticks])
    ax.tick_params(labelsize=9, length=3, colors="#77828a")
    ax.set_xlim(-5.85, 7.45)
    ax.set_ylim(-5.85, 5.85)
    ax.set_xlabel("地图 X 坐标", fontproperties=FONT, fontsize=10, color="#586871", labelpad=10)
    ax.set_ylabel("地图 Y 坐标", fontproperties=FONT, fontsize=10, color="#586871", labelpad=10)
    for spine in ax.spines.values():
        spine.set_visible(False)
    fig.text(0.08, 0.108, "求解结果：第一圈 8 人，第二圈 16 人。人数由格位决定，未预先固定。", fontproperties=BOLD, fontsize=12, color="#3b514f")
    fig.text(0.08, 0.070, "每个姓名块都是完整 2×2 城堡；空白格是真实空隙。蓝色旗子可占住圈内位置。", fontproperties=FONT, fontsize=10.5, color="#627078")
    fig.text(0.08, 0.037, "v4 方案核对图：展示前两圈的姓名；完整100人正式图仍待确认。", fontproperties=FONT, fontsize=10.5, color="#627078")
    fig.savefig(HERE / "前两圈_俯视.png", dpi=190, facecolor=fig.get_facecolor())
    plt.close(fig)


def iso(x, y, z=0):
    return x - y, 0.57 * (x + y) + z


def iso_block(ax, x, y, width, height, color, depth=0.07, edge="#5c645e", zorder=3):
    for a, b in [((x, y), (x + width, y)), ((x, y), (x, y + height))]:
        pts = [iso(*a), iso(*b), iso(*b, depth), iso(*a, depth)]
        ax.add_patch(Polygon(pts, facecolor=edge, edgecolor=edge, linewidth=0.4, zorder=zorder))
    corners = [iso(x, y, depth), iso(x + width, y, depth), iso(x + width, y + height, depth), iso(x, y + height, depth)]
    ax.add_patch(Polygon(corners, facecolor=color, edgecolor=edge, linewidth=0.8, zorder=zorder + 0.1))


def draw_iso(rows, flags):
    fig = plt.figure(figsize=(15.5, 11.5), facecolor="#f7f5f1")
    title(fig, "两格一圈 · 前两圈姓名预览", "红色第一圈 / 绿色第二圈 / 蓝色旗子 / 金色熊坑；空白处保留实际间隙")
    ax = fig.add_axes([0.035, 0.16, 0.93, 0.68])
    ax.set_aspect("equal")
    ax.set_xlim(-12.7, 12.7)
    ax.set_ylim(-7.6, 7.6)
    ax.axis("off")
    for x in range(-5, 6):
        for y in range(-5, 6):
            iso_block(ax, x - 0.5, y - 0.5, 1, 1, "#f4f2ed", depth=0, edge="#d4d7d3", zorder=1)
    for row in rows:
        if row["ring"] > 2:
            continue
        x, y = row["relative_anchor"]
        iso_block(ax, x - 0.5, y - 0.5, 2, 2, PALETTE[row["ring"]], edge="#627068", zorder=3)
        for p1, p2 in [((x + 0.5, y - 0.5), (x + 0.5, y + 1.5)), ((x - 0.5, y + 0.5), (x + 1.5, y + 0.5))]:
            a, b = iso(*p1, 0.08), iso(*p2, 0.08)
            ax.plot([a[0], b[0]], [a[1], b[1]], color="white", alpha=0.16, linewidth=0.55, zorder=4)
        text(ax, *iso(x + 0.5, y + 0.5, 0.09), name_lines(row["name"]), size=10.2, bold=True, ha="center", va="center", color="white", zorder=12)
    iso_block(ax, -1.5, -1.5, 3, 3, "#e4c777", depth=0.1, edge="#b69a52", zorder=4)
    text(ax, *iso(0, 0, 0.12), "熊坑\n3×3", size=15, bold=True, ha="center", va="center", color="#5d4e2d", zorder=10)
    for x, y in sorted(flags, key=lambda a: -sum(a)):
        if max(abs(x), abs(y)) > 5:
            continue
        iso_block(ax, x - 0.5, y - 0.5, 1, 1, "#467da5", depth=0.6, edge="#2d526e", zorder=9)
    fig.text(0.08, 0.097, "第一圈 8 人 · 第二圈 16 人｜每圈厚度均为2格，城堡不会横跨两圈", fontproperties=BOLD, fontsize=13, color="#334d4a")
    fig.text(0.08, 0.055, "旗子保持上下、左右镜像；城堡格位采用90°旋转对称，以减少奇数净空造成的浪费。", fontproperties=FONT, fontsize=11, color="#647078")
    fig.savefig(HERE / "前两圈_斜视.png", dpi=190, facecolor=fig.get_facecolor())
    plt.close(fig)


def draw_full(rows, flags, counts):
    fig = plt.figure(figsize=(12.5, 12), facecolor="#f7f5f1")
    title(fig, "100座城堡 · 几何可行性预览", "五圈厚度全部为2格；允许旗子造成断口；外圈保留空位，尚非完整人名正式图")
    ax = fig.add_axes([0.11, 0.16, 0.75, 0.70])
    ax.set_aspect("equal")
    for x, y in sorted(BOARD):
        ax.add_patch(Rectangle((x - 0.5, y - 0.5), 1, 1, facecolor="#fdfcf9", edgecolor="#dce0dc", linewidth=0.35))
    for row in rows:
        x, y = row["relative_anchor"]
        ax.add_patch(Rectangle((x - 0.5, y - 0.5), 2, 2, facecolor=PALETTE[row["ring"]], edgecolor="#53635c", linewidth=0.55))
    for half in (1.5, 3.5, 5.5, 7.5, 9.5, 11.5):
        ax.add_patch(Rectangle((-half, -half), 2 * half, 2 * half, fill=False, edgecolor="#766f65", linewidth=1.0, linestyle=(0, (4, 3))))
    ax.add_patch(Rectangle((-1.5, -1.5), 3, 3, facecolor="#e4c777", edgecolor="#ad9251"))
    text(ax, 0, 0, "熊坑", size=11, bold=True, ha="center", va="center", color="#5d4e2d")
    for x, y in flags:
        ax.add_patch(Rectangle((x - 0.5, y - 0.5), 1, 1, facecolor="#275a83", edgecolor="#143c59", linewidth=0.8))
    ax.set_xlim(-11.6, 11.6)
    ax.set_ylim(-13.6, 12.6)
    ax.set_xticks([-11, -8, -3, 0, 3, 8, 11], [str(CENTER[0] + x) for x in [-11, -8, -3, 0, 3, 8, 11]])
    ax.set_yticks([-13, -8, -3, 0, 3, 8, 12], [str(CENTER[1] + y) for y in [-13, -8, -3, 0, 3, 8, 12]])
    ax.tick_params(labelsize=9)
    for spine in ax.spines.values():
        spine.set_color("#9aaba4")
    fig.text(0.07, 0.105, " / ".join(f"第{k}圈 {counts[k]}人" for k in range(1, 6)), fontproperties=BOLD, fontsize=11.5, color="#435a52")
    fig.text(0.07, 0.066, "验证通过：100座互不重叠、全在范围内、每座完整落在一面旗子的7×7覆盖内。", fontproperties=FONT, fontsize=10.5, color="#65706d")
    fig.text(0.07, 0.033, "外侧旗子Y坐标由上一版的641、659调整为642、658，使旗阵和城堡都满足90°旋转对称。", fontproperties=FONT, fontsize=10, color="#65706d")
    fig.savefig(HERE / "全域_几何预览.png", dpi=170, facecolor=fig.get_facecolor())
    plt.close(fig)


def main():
    roster = json.loads((HERE / "members_snapshot.json").read_text(encoding="utf-8"))
    members = roster["active_members"]
    assert len(members) == 100
    comparisons = []
    # 比较的是这一族对称16旗阵，不宣称穷尽任意数量、任意位置的旗阵。
    for inner in (2, 3, 4):
        for outer in (7, 8):
            for symmetry in ("mirror", "rotate"):
                flags = flag_array(inner, outer)
                cap = [len(solve_ring(flags, k, symmetry)) for k in range(1, 6)]
                comparisons.append({"inner_offset": inner, "outer_offset": outer, "castle_symmetry": symmetry, "ring_capacities": cap, "total_capacity": sum(cap)})
    (HERE / "flag_family_comparison.json").write_text(json.dumps(comparisons, ensure_ascii=False, indent=2), encoding="utf-8")
    eligible = [r for r in comparisons if r["castle_symmetry"] == "rotate" and r["total_capacity"] >= 100]
    selected = max(eligible, key=lambda r: tuple(r["ring_capacities"]))
    flags = flag_array(selected["inner_offset"], selected["outer_offset"])
    remaining = 100
    anchors = []
    counts = {}
    for k, capacity in enumerate(selected["ring_capacities"], 1):
        target = min(remaining, capacity)
        assert target % 4 == 0
        ring_anchors = solve_ring(flags, k, "rotate", target=target)
        anchors.extend(ring_anchors)
        counts[k] = len(ring_anchors)
        remaining -= target
    assert remaining == 0
    verify(anchors, flags, counts)
    # 成员顺序来自已核对的数据；实际分配按中心距离升序，避免较高排位被放到更远格位。
    anchors.sort(key=lambda a: ((a[0] + 0.5) ** 2 + (a[1] + 0.5) ** 2, ring_index(a), -math.atan2(a[1] + 0.5, a[0] + 0.5)))
    rows = []
    for member, anchor in zip(members, anchors):
        x, y = anchor
        rows.append({"name": member["name"], "order": member["provisional_order"], "screenshot_rank": member["screenshot_rank"], "power": member["power"], "ring": ring_index(anchor), "relative_anchor": anchor, "anchor": [CENTER[0] + x, CENTER[1] + y], "distance": math.hypot(x + 0.5, y + 0.5)})
    assert [r["distance"] for r in rows] == sorted(r["distance"] for r in rows)
    assert {r["order"] for r in rows if r["ring"] <= 2} == set(range(1, 25))
    geometry = {
        "status": "制作前核对稿；已求解100座几何，仅展示前两圈完整姓名",
        "bear_center": CENTER,
        "ring_width": 2,
        "ring_outer_side_lengths": [7, 11, 15, 19, 23],
        "ring_counts": counts,
        "ring_capacities_for_selected_flags": selected["ring_capacities"],
        "flag_positions": sorted((CENTER[0] + x, CENTER[1] + y) for x, y in flags),
        "flag_symmetry": "上下镜像、左右镜像、90度旋转",
        "castle_symmetry": "90度旋转；不要求每一行上下左右镜像",
        "coverage": "每座城堡完整落在至少一面旗子的7×7范围内",
        "validated": {"castle_count": 100, "castle_cells": 400, "flag_count": 16, "no_overlap": True, "inside_bounds": True, "all_castles_in_one_ring": True, "rank_distance_monotonic": True},
        "capacity_comparison_scope": "仅比较内偏移2/3/4、外偏移7/8的12种组合，未宣称所有旗阵全局最优",
        "placements_for_review": rows,
        "final_production_approved": False,
    }
    (HERE / "ring_preview.json").write_text(json.dumps(geometry, ensure_ascii=False, indent=2), encoding="utf-8")
    with (HERE / "前两圈_姓名坐标.csv").open("w", encoding="utf-8-sig", newline="") as output:
        writer = csv.writer(output)
        writer.writerow(["内部排位", "成员", "圈数", "城堡左下格X", "城堡左下格Y", "截图名次", "战力"])
        writer.writerows([r["order"], r["name"], r["ring"], *r["anchor"], r["screenshot_rank"], r["power"]] for r in rows if r["ring"] <= 2)
    draw_top(rows, flags)
    draw_iso(rows, flags)
    draw_full(rows, flags, counts)
    print(json.dumps({"ring_counts": counts, "capacity": sum(selected["ring_capacities"]), "chosen_inner": selected["inner_offset"], "chosen_outer": selected["outer_offset"], "verified_castles": 100, "first_two_named": sum(r["ring"] <= 2 for r in rows)}, ensure_ascii=False))


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    main()
