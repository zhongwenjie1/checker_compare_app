# -*- coding: utf-8 -*-
"""
组合票排程 + 导出
- Zone（阻塞区域）：同一 zone_id 的连续步骤视为同一区域，容量=可同时处于区域的车辆数。
- gate_zone（闸门）+ gate_buffer（缓冲）：对上游某些步骤设 gate_zone=某区域，
  表示该步骤“开工前/放行时”需要考虑该区域前的“闸门缓冲”是否已满。
  gate_buffer=缓冲允许在“闸门 → 区域入口”链路上同时存在的在制车数量（默认=2）。
  这样可实现：2 号车能继续到电检1，但 3 号车需等 2 号车进入电检2 后才开始电检准备。
"""

from __future__ import annotations
import math
import heapq
from typing import List, Dict, Any, Tuple
import pandas as pd


# ---------------- Excel 引擎选择 ---------------- #
def _choose_engine():
    try:
        import xlsxwriter  # noqa: F401
        return "xlsxwriter"
    except Exception:
        try:
            import openpyxl  # noqa: F401
            return "openpyxl"
        except Exception:
            return None


# ---------------- 解析步骤与 Zone ---------------- #
def _normalize_defs(step_defs: List[Dict[str, Any]]) -> Tuple[List[Dict[str, Any]], Dict[str, Dict[str, Any]], Dict[str, int]]:
    """
    返回 (steps, zones, gate_buffers)

    steps: 每步 {seq, display, group, duration, zone_id, gate_zone_id}
    zones: {zid: {"capacity": int, "first_seq": int, "last_seq": int}}
    gate_buffers: {zid: gate_buffer_int}  # 若未显式提供，默认=2
    """
    steps: List[Dict[str, Any]] = []
    # 收集 gate_buffer（按 zone 聚合，取出现的最大值；默认=2）
    gate_buffers: Dict[str, int] = {}

    for d in step_defs:
        display = str(d.get("display", "")).strip()
        group = str(d.get("group", "")).strip() or display
        durations = list(d.get("durations", []))
        if not display or not durations:
            continue
        dur = float(durations[0])

        zone_id = str(d.get("zone_id", "") or "").strip()
        gate_zone_id = str(d.get("gate_zone_id", "") or "").strip()

        # 聚合 gate_buffer
        if gate_zone_id:
            gb = d.get("gate_buffer", None)
            if gb is None:
                gb = 2  # 默认缓冲=2，符合“2号车可走，3号车等”的现场规则
            try:
                gb = max(1, int(float(gb)))
            except Exception:
                gb = 2
            if gate_zone_id in gate_buffers:
                gate_buffers[gate_zone_id] = max(gate_buffers[gate_zone_id], gb)
            else:
                gate_buffers[gate_zone_id] = gb

        steps.append({
            "seq": int(d.get("seq", len(steps) + 1)),
            "display": display,
            "group": group,
            "duration": dur,
            "zone_id": zone_id,
            "gate_zone_id": gate_zone_id,
        })

    steps.sort(key=lambda x: x["seq"])
    if not steps:
        raise ValueError("没有有效的步骤定义")

    # 汇总 Zone：确定起止步骤序号 + 容量
    zones: Dict[str, Dict[str, Any]] = {}
    for s in steps:
        zid = s.get("zone_id", "")
        if not zid:
            continue
        z = zones.setdefault(zid, {"capacity": 1, "first_seq": s["seq"], "last_seq": s["seq"]})
        z["first_seq"] = min(z["first_seq"], s["seq"])
        z["last_seq"] = max(z["last_seq"], s["seq"])

    # 从原始定义里补全 zone 容量
    for d in step_defs:
        zid = str(d.get("zone_id", "") or "").strip()
        if not zid or zid not in zones:
            continue
        zcap = d.get("zone_capacity", None)
        if zcap is not None:
            try:
                zones[zid]["capacity"] = max(int(zones[zid]["capacity"]), int(zcap))
            except Exception:
                pass

    # gate_buffer 若某个 gate_zone 没出现在任何行里，忽略；出现但没显式给值时已默认为 2
    return steps, zones, gate_buffers


# ---------------- 调度（含 Zone + gate_buffer） ---------------- #
def schedule(step_defs: List[Dict[str, Any]], cars: int) -> Tuple[List[Dict[str, Any]], float]:
    """
    返回：
      rows: 每车-每步记录：
        {car, step_seq, step_display, group, dur, start, svc_finish, depart, block_wait}
      max_time: 全局最后 depart
    """
    steps, zones, gate_buffers = _normalize_defs(step_defs)
    m = len(steps)

    # 工位释放时刻（考虑阻塞传递）
    server_free = [0.0 for _ in range(m)]

    # Zone 名额堆：zid -> [free_time, ...]（长度=capacity）
    zone_heaps: Dict[str, List[float]] = {}
    for zid, zinfo in zones.items():
        cap = int(zinfo.get("capacity", 1)) or 1
        zone_heaps[zid] = [0.0 for _ in range(cap)]
        heapq.heapify(zone_heaps[zid])

    # 闸门缓冲：对每个 gate_zone 维护“尚未进入该 zone 的车辆的预计进入时刻”最小堆
    # pre_heap[z] 中的元素是“已经通过闸门但尚未进入 zone 的车辆的 ‘zone 入口开始时间’ ”
    pre_heap: Dict[str, List[float]] = {}

    def is_zone_entry(idx: int) -> bool:
        s = steps[idx]
        zid = s.get("zone_id", "")
        if not zid:
            return False
        return s["seq"] == zones[zid]["first_seq"]

    def is_zone_exit(idx: int) -> bool:
        s = steps[idx]
        zid = s.get("zone_id", "")
        if not zid:
            return False
        return s["seq"] == zones[zid]["last_seq"]

    rows: List[Dict[str, Any]] = []
    max_time = 0.0

    for car in range(1, cars + 1):
        prev_depart = 0.0
        # 记录该车是否经过某个 gate_zone（用于之后把它的“进入 zone 的时刻”加入 pre_heap）
        car_gate_zones: set[str] = set()

        for j, st in enumerate(steps):
            # ---- 计算本步开始时间：受上一步 depart、本步服务器空闲约束 ----
            start = max(server_free[j], prev_depart)

            # ---- 闸门缓冲约束（在 start 阶段判断）：允许“闸门→区域入口”链路上最多 gate_buffer 辆 ----
            gz = st.get("gate_zone_id", "")
            if gz:
                car_gate_zones.add(gz)
                # 取得该 gate_zone 的缓冲与堆
                gb = max(1, int(gate_buffers.get(gz, 2)))
                heap = pre_heap.setdefault(gz, [])

                # 移除所有“进入 zone 的时刻 <= start”的条目（这些车在 start 时刻已进入 zone，不再占用缓冲）
                while heap and heap[0] <= start:
                    heapq.heappop(heap)

                # 若缓冲已满（heap 大小 >= gb），则把 start 推迟到“最早一辆进入 zone 的时刻”
                # 推迟后再次清理（可能一次就够，也可能要多次）
                while len(heap) >= gb:
                    start = max(start, heap[0])
                    while heap and heap[0] <= start:
                        heapq.heappop(heap)

            # ---- 服务结束 ----
            svc_finish = start + float(st["duration"])

            # ---- depart 受“下步可接收（服务器释放 + zone 容量）”约束 ----
            if j < m - 1:
                next_ready = server_free[j + 1]

                # 若“下步”是某 Zone 的入口，还得等该 Zone 出现名额
                if is_zone_entry(j + 1):
                    nzid = steps[j + 1]["zone_id"]
                    nheap = zone_heaps[nzid]
                    next_ready = max(next_ready, nheap[0] if nheap else 0.0)

                # 注：闸门缓冲只在 start 阶段处理，不再额外卡 depart

                depart = max(svc_finish, next_ready)
            else:
                depart = svc_finish

            block_wait = max(0.0, depart - svc_finish)

            rows.append({
                "car": car,
                "step_seq": st["seq"],
                "step_display": st["display"],
                "group": st["group"],
                "dur": float(st["duration"]),
                "start": start,
                "svc_finish": svc_finish,
                "depart": depart,
                "block_wait": block_wait,
            })

            # ---- Zone 名额占用/释放 ----
            # 进入 Zone：仅在“Zone 入口步骤”发生，占用一个名额
            if is_zone_entry(j):
                zid = st["zone_id"]
                # 如果该车之前通过过指向该 zid 的某个闸门，则把它“进入 zone 的时刻（=本步 start=上一步 depart）”加入 pre_heap
                if zid in car_gate_zones:
                    heap = pre_heap.setdefault(zid, [])
                    heapq.heappush(heap, start)  # 之后其他车的“闸门开始”会受这个时间点约束

                heap = zone_heaps[zid]
                if heap:
                    heapq.heappop(heap)  # 占用一个 zone 名额

            # 离开 Zone：仅在“Zone 最后一步”释放一个名额（名额释放时刻=本步 depart）
            if is_zone_exit(j):
                zid = st["zone_id"]
                heap = zone_heaps[zid]
                heapq.heappush(heap, depart)

            # ---- 更新状态，进入下一步 ----
            server_free[j] = depart
            prev_depart = depart
            max_time = max(max_time, depart)

    return rows, max_time


# ---------------- 等待统计 ---------------- #
def _build_car_slices(rows: List[Dict[str, Any]]) -> Dict[int, List[Dict[str, Any]]]:
    by_car: Dict[int, List[Dict[str, Any]]] = {}
    for r in rows:
        by_car.setdefault(r["car"], []).append(r)
    for k in by_car:
        by_car[k].sort(key=lambda x: x["step_seq"])
    return by_car


def _compute_entry_wait(by_car: Dict[int, List[Dict[str, Any]]]) -> Dict[int, float]:
    """入站等待：车 i 第一步 start - 车 i-1 第一步 depart（<0 计 0）"""
    wait_map: Dict[int, float] = {}
    prev_first_depart = 0.0
    for car in sorted(by_car.keys()):
        steps = by_car[car]
        first_start = steps[0]["start"]
        wait_map[car] = max(0.0, first_start - prev_first_depart)
        prev_first_depart = steps[0]["depart"]
    return wait_map


def _compute_total_wait(by_car: Dict[int, List[Dict[str, Any]]]) -> Dict[int, float]:
    """总等待 = 入站等待 + Σ block_wait（所有步）"""
    entry_map = _compute_entry_wait(by_car)
    total_map: Dict[int, float] = {}
    for car, steps in by_car.items():
        inter = sum(max(0.0, s.get("block_wait", 0.0)) for s in steps)
        total_map[car] = float(entry_map.get(car, 0.0) + inter)
    return total_map


# ---------------- 导出入口 ---------------- #
def schedule_and_export(defs: List[Dict[str, Any]],
                        cars: int,
                        grid_step: float,
                        wait_policy: str,   # "before"/"after" 仅影响是否绘入站等待条
                        project: str,
                        dst_path: str) -> None:
    grid_step = 1.0 if (not isinstance(grid_step, (int, float)) or grid_step <= 0) else float(grid_step)
    rows, max_finish = schedule(defs, cars)
    # ---- 收集用户自定义颜色 (display -> hex)
    step_color_map = {d.get("display"): d.get("color") for d in defs if d.get("color")}
    engine = _choose_engine()
    if engine is None:
        raise RuntimeError("未找到可用的 Excel 引擎，请安装 xlsxwriter 或 openpyxl")

    if engine == "xlsxwriter":
        _export_with_xlsxwriter(rows, max_finish, grid_step, wait_policy, project, dst_path, step_color_map)
    else:
        _export_with_openpyxl(rows, max_finish, grid_step, wait_policy, project, dst_path)


# ---------------- 样式与工具 ---------------- #
def _palette():
    group_colors = [
        "#4CAF50", "#2196F3", "#9C27B0", "#FF9800", "#009688",
        "#795548", "#3F51B5", "#E91E63", "#00BCD4", "#8BC34A",
    ]
    wait_color = "#FFC107"
    return group_colors, wait_color


def _fmt_num(x: float) -> str:
    if abs(x - round(x)) < 1e-9:
        return str(int(round(x)))
    return f"{x:.1f}"


# ---------------- xlsxwriter 彩色导出 ---------------- #
def _export_with_xlsxwriter(rows: List[Dict[str, Any]], max_finish: float,
                             grid_step: float, wait_policy: str,
                             project: str, dst_path: str,
                             step_color_map: Dict[str, str]) -> None:
    import xlsxwriter  # type: ignore

    by_car = _build_car_slices(rows)
    entry_wait = _compute_entry_wait(by_car)
    total_wait = _compute_total_wait(by_car)

    n_cols_grid = max(1, int(math.ceil(max_finish / grid_step)))

    with pd.ExcelWriter(dst_path, engine="xlsxwriter") as writer:
        wb = writer.book
        ws = wb.add_worksheet("作业组合票")
        writer.sheets["作业组合票"] = ws

        fmt_header = wb.add_format({"bold": True, "align": "center", "valign": "vcenter", "bg_color": "#EEEEEE", "border": 1})
        fmt_text   = wb.add_format({"align": "center", "valign": "vcenter", "border": 1})
        fmt_left   = wb.add_format({"align": "left", "valign": "vcenter", "border": 1})
        fmt_wait   = wb.add_format({"align": "left", "valign": "vcenter", "border": 1, "bg_color": "#FFF9C4"})
        fmt_bar_wait = wb.add_format({"bg_color": "#FFE082", "border": 0})
        fmt_car    = wb.add_format({"align": "center", "valign": "vcenter", "border": 1, "bg_color": "#F5F5F5"})

        group_colors, _ = _palette()
        group_fmt_cache: Dict[str, Any] = {}
        def bar_fmt(group: str, display: str):
            # 若当前步骤有自定义颜色，优先使用
            custom_hex = step_color_map.get(display)
            if custom_hex:
                if custom_hex not in group_fmt_cache:
                    group_fmt_cache[custom_hex] = wb.add_format({"bg_color": custom_hex, "border": 0})
                return group_fmt_cache[custom_hex]
            # 否则按 group 调色盘
            if group not in group_fmt_cache:
                idx = (hash(group) >> 1) % len(group_colors)
                group_fmt_cache[group] = wb.add_format({"bg_color": group_colors[idx], "border": 0})
            return group_fmt_cache[group]

        ws.set_column(0, 0, 36)
        ws.set_column(1, 1, 8)
        ws.set_column(2, 2, 18)
        ws.set_column(3, 3, 10)
        ws.set_column(4, 4 + n_cols_grid - 1, 2.8)

        ws.write(0, 0, f"连续投入{project}等待时间", fmt_header)
        ws.write(0, 1, "车号", fmt_header)
        ws.write(0, 2, "项目", fmt_header)
        ws.write(0, 3, "时间", fmt_header)
        for i in range(n_cols_grid):
            ws.write(0, 4 + i, f"{grid_step:.1f}", fmt_header)
        ws.freeze_panes(1, 0)

        row_cursor = 1
        for car in sorted(by_car.keys()):
            steps = by_car[car]
            if not steps:
                continue

            ewait = float(entry_wait.get(car, 0.0))
            twait = float(total_wait.get(car, 0.0))
            ws.write(row_cursor, 0, f"入站等待{_fmt_num(ewait)}秒；总等待{_fmt_num(twait)}秒", fmt_wait if ewait > 0 else fmt_left)
            ws.write(row_cursor, 1, car, fmt_car)
            ws.write(row_cursor, 2, "", fmt_left)
            ws.write(row_cursor, 3, _fmt_num(ewait) if ewait > 0 else "", fmt_text if ewait > 0 else fmt_left)

            if ewait > 0 and wait_policy == "before":
                first_start = steps[0]["start"]
                c0 = 4
                c1 = 4 + int(math.ceil(first_start / grid_step)) - 1
                c1 = max(c1, c0 - 1)
                for c in range(c0, c1 + 1):
                    ws.write(row_cursor, c, "", fmt_bar_wait)
            row_cursor += 1

            for idx, s in enumerate(steps):
                # 服务条
                ws.write(row_cursor, 0, "", fmt_left)
                ws.write(row_cursor, 1, "", fmt_text)
                ws.write(row_cursor, 2, s["step_display"], fmt_left)
                ws.write(row_cursor, 3, _fmt_num(s["dur"]), fmt_text)
                c_start = 4 + int(math.floor(s["start"] / grid_step))
                c_end_svc = 4 + int(math.ceil(s["svc_finish"] / grid_step)) - 1
                c_end_svc = max(c_end_svc, c_start)
                bf = bar_fmt(s["group"], s["step_display"])
                for c in range(c_start, c_end_svc + 1):
                    ws.write(row_cursor, c, "", bf)
                row_cursor += 1

                # 等待条（svc_finish → depart）
                if s["block_wait"] > 1e-9 and idx < len(steps) - 1:
                    wait_val = s["block_wait"]
                    next_name = steps[idx + 1]["step_display"]
                    ws.write(row_cursor, 0, f"等待{_fmt_num(wait_val)}秒（{s['step_display']} → {next_name}）", fmt_wait)
                    ws.write(row_cursor, 1, "", fmt_text)
                    ws.write(row_cursor, 2, "", fmt_wait)
                    ws.write(row_cursor, 3, _fmt_num(wait_val), fmt_text)
                    c_w0 = 4 + int(math.floor(s["svc_finish"] / grid_step))
                    c_w1 = 4 + int(math.ceil(s["depart"] / grid_step)) - 1
                    c_w1 = max(c_w1, c_w0)
                    for c in range(c_w0, c_w1 + 1):
                        ws.write(row_cursor, c, "", fmt_bar_wait)
                    row_cursor += 1

            row_cursor += 1  # 车与车之间空一行


# ---------------- openpyxl 回退导出（文字） ---------------- #
def _export_with_openpyxl(rows: List[Dict[str, Any]], max_finish: float,
                           grid_step: float, wait_policy: str,
                           project: str, dst_path: str) -> None:
    by_car = _build_car_slices(rows)
    entry_wait = _compute_entry_wait(by_car)
    total_wait = _compute_total_wait(by_car)

    out_rows = []
    for car in sorted(by_car.keys()):
        steps = by_car[car]
        if not steps:
            continue
        ewait = float(entry_wait.get(car, 0.0))
        twait = float(total_wait.get(car, 0.0))
        out_rows.append({
            "车号": car,
            "项目": "(入站等待/总等待)",
            "时间": ewait,
            "说明": f"入站等待{_fmt_num(ewait)}秒；总等待{_fmt_num(twait)}秒"
        })
        for idx, s in enumerate(steps):
            out_rows.append({"车号": car, "项目": s["step_display"], "时间": s["dur"], "说明": ""})
            if s["block_wait"] > 1e-9 and idx < len(steps) - 1:
                out_rows.append({
                    "车号": car,
                    "项目": f"(等待：{s['step_display']}→{steps[idx+1]['step_display']})",
                    "时间": s["block_wait"],
                    "说明": f"等待{_fmt_num(s['block_wait'])}秒"
                })
        out_rows.append({"车号": "", "项目": "", "时间": "", "说明": ""})

    df = pd.DataFrame(out_rows, columns=["车号", "项目", "时间", "说明"])
    # TODO: 未应用自定义颜色（step_color_map）到 openpyxl 导出
    with pd.ExcelWriter(dst_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="作业组合票")
        try:
            ws = writer.sheets["作业组合票"]
            ws.freeze_panes = "A2"
        except Exception:
            pass