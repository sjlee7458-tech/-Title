import io
import json
import numpy as np
import pandas as pd

WORKBOOK_CACHE = {}


def _detect_classification(sheet_name: str) -> str:
    clean = str(sheet_name).replace(" ", "")
    for key in ["기본부문", "소분류", "중분류", "대분류"]:
        if key in clean:
            return key
    return "기본부문"


def _detect_trade_scope(sheet_name: str) -> str:
    clean = str(sheet_name).replace(" ", "")
    if "국산" in clean:
        return "국산거래표"
    if "수입" in clean:
        return "수입거래표"
    return "총거래표"


def _parse_sheet_matrix(df_raw: pd.DataFrame):
    df = df_raw.dropna(how="all").dropna(axis=1, how="all").copy()
    if df.shape[0] < 4 or df.shape[1] < 4:
        raise ValueError("유효한 투입산출표 행렬이 아닙니다.")

    # Detect whether row 0 has ["부문코드", "부문명", ...] or raw BOK header around row 4~5
    first_cell = str(df.iloc[0, 0]).strip()
    if first_cell == "부문코드":
        code_col = df.iloc[1:, 0].astype(str).str.strip()
        name_col = df.iloc[1:, 1].astype(str).str.strip()
        row_offset = 1
    else:
        # Raw BOK sheet fallback: find row where col 0 == '0111'
        start_r = None
        for r in range(len(df)):
            if str(df.iloc[r, 0]).strip() == "0111":
                start_r = r
                break
        if start_r is None:
            raise ValueError("지원하지 않는 시트 헤더 형식입니다.")
        code_col = df.iloc[start_r:, 0].astype(str).str.strip()
        name_col = df.iloc[start_r:, 1].astype(str).str.strip()
        row_offset = start_r

    exclude_codes = {"9510", "9520", "9590", "9610", "9621", "9622", "9630", "9690", "9790"}
    endo_r_indices = []
    codes = []
    names = []
    special_r = {}

    for idx_rel, (c_val, n_val) in enumerate(zip(code_col, name_col)):
        abs_r = row_offset + idx_rel
        if c_val in exclude_codes:
            special_r[c_val] = abs_r
        elif c_val and c_val.lower() != "nan" and not c_val.startswith("9"):
            endo_r_indices.append(abs_r)
            codes.append(c_val)
            names.append(n_val.replace("\n", " "))

    n = len(codes)
    if n < 2:
        raise ValueError("내생 산업 부문이 2개 미만입니다.")

    # Endogenous columns start at column index 2 and span n columns
    Z = df.iloc[endo_r_indices, 2:2 + n].apply(pd.to_numeric, errors="coerce").fillna(0.0).to_numpy(dtype=float)

    if "9790" in special_r:
        total_output = pd.to_numeric(df.iloc[special_r["9790"], 2:2 + n], errors="coerce").fillna(0.0).to_numpy(dtype=float)
    else:
        total_output = np.maximum(Z.sum(axis=0) * 1.4, 1.0)

    if "9690" in special_r:
        va_row = pd.to_numeric(df.iloc[special_r["9690"], 2:2 + n], errors="coerce").fillna(0.0).to_numpy(dtype=float)
    else:
        va_row = np.maximum(total_output - Z.sum(axis=0), 0.0)

    safe_x = np.where(total_output <= 1e-9, 1.0, total_output)
    A = Z / safe_x[np.newaxis, :]
    A = np.clip(A, 0.0, 0.98)

    B = np.linalg.inv(np.eye(n) - A)

    output_mult = B.sum(axis=0)
    mean_b = float(np.mean(B))
    power_idx = B.mean(axis=0) / mean_b
    sens_idx = B.mean(axis=1) / mean_b

    va_rate = np.clip(va_row / safe_x, 0.0, 1.0)
    va_mult = (va_rate[:, np.newaxis] * B).sum(axis=0)

    display_labels = [f"[{c}] {nm}" for c, nm in zip(codes, names)]

    return {
        "codes": codes,
        "names": names,
        "labels": display_labels,
        "B": B,
        "va_rate": va_rate,
        "output_mult": output_mult,
        "va_mult": va_mult,
        "power_idx": power_idx,
        "sens_idx": sens_idx,
    }


def load_workbook_summary(excel_bytes: bytes) -> str:
    WORKBOOK_CACHE.clear()
    xl = pd.ExcelFile(io.BytesIO(excel_bytes), engine="openpyxl")
    sheets = []

    for s_name in xl.sheet_names:
        cls_type = _detect_classification(s_name)
        scope = _detect_trade_scope(s_name)
        try:
            raw = xl.parse(s_name, header=None)
            parsed = _parse_sheet_matrix(raw)
            WORKBOOK_CACHE[s_name] = {
                "classification": cls_type,
                "tradeScope": scope,
                **parsed,
            }
            sheets.append({
                "sheetName": s_name,
                "classification": cls_type,
                "tradeScope": scope,
                "sectorCount": len(parsed["labels"]),
                "sectors": parsed["labels"],
            })
        except Exception as e:
            sheets.append({
                "sheetName": s_name,
                "classification": cls_type,
                "tradeScope": scope,
                "sectorCount": 0,
                "sectors": [],
                "error": str(e),
            })

    return json.dumps({
        "sheets": sheets,
        "unit": "금액: 백만원 / 투입계수·유발계수·연쇄지수: 무차원",
    }, ensure_ascii=False)


def _linkage_type(power: float, sens: float) -> str:
    if power >= 1.0 and sens >= 1.0:
        return "중간수요적 제조업형 (전·후방 모두 높음)"
    if power >= 1.0 and sens < 1.0:
        return "최종수요적 제조업형 (후방연쇄 중심)"
    if power < 1.0 and sens >= 1.0:
        return "중간수요적 기초·인프라형 (전방연쇄 중심)"
    return "최종수요적 독립·서비스형 (전·후방 모두 낮음)"


def _match_sector_index(labels: list, target_label: str) -> int:
    if target_label in labels:
        return labels.index(target_label)
    # Keyword fallback across classification levels
    keywords = []
    for kw in ["집적회로", "반도체", "기계", "전력", "소프트웨어", "전자", "제조업"]:
        if kw in target_label:
            keywords.append(kw)
    keywords.extend(["반도체", "전자", "제조업"])
    for kw in keywords:
        for i, lbl in enumerate(labels):
            if kw in lbl:
                return i
    return 0


def run_io_analysis(sheet_name: str, target_sector: str, demand_shock: float, shock_label: str) -> str:
    item = WORKBOOK_CACHE.get(sheet_name)
    if not item:
        raise ValueError("선택한 시트가 로딩되지 않았습니다.")

    labels = item["labels"]
    n = len(labels)
    idx = _match_sector_index(labels, target_sector)

    dy = np.zeros(n, dtype=float)
    dy[idx] = demand_shock

    B = item["B"]
    va_rate = item["va_rate"]
    dx = B @ dy
    dva = va_rate * dx

    tot_out = float(np.sum(dx))
    tot_va = float(np.sum(dva))

    rows = []
    for i in range(n):
        rows.append({
            "code": item["codes"][i],
            "name": item["names"][i],
            "deltaOutput": round(float(dx[i]), 2),
            "deltaVa": round(float(dva[i]), 2),
            "outputMultiplier": round(float(item["output_mult"][i]), 4),
            "vaMultiplier": round(float(item["va_mult"][i]), 4),
            "powerIndex": round(float(item["power_idx"][i]), 4),
            "sensitivityIndex": round(float(item["sens_idx"][i]), 4),
        })

    rows.sort(key=lambda r: r["deltaOutput"], reverse=True)

    # Compare across 대·중·소분류·기본부문 under the same trade scope
    current_scope = item["tradeScope"]
    cross_level = []
    for s_name, cached in WORKBOOK_CACHE.items():
        if cached["tradeScope"] != current_scope or not cached["labels"]:
            continue
        c_idx = _match_sector_index(cached["labels"], target_sector)
        dy_c = np.zeros(len(cached["labels"]), dtype=float)
        dy_c[c_idx] = demand_shock
        dx_c = cached["B"] @ dy_c
        dva_c = cached["va_rate"] * dx_c
        sum_dx = float(np.sum(dx_c))
        cr3 = float(np.sum(np.sort(dx_c)[::-1][:3]) / sum_dx * 100.0) if sum_dx > 0 else 0.0
        cross_level.append({
            "classification": cached["classification"],
            "sheetName": s_name,
            "sectorCount": len(cached["labels"]),
            "matchedSector": cached["labels"][c_idx],
            "totalOutputEffect": round(sum_dx, 2),
            "totalVaEffect": round(float(np.sum(dva_c)), 2),
            "cr3Pct": round(cr3, 2),
        })

    return json.dumps({
        "sheetName": sheet_name,
        "classification": item["classification"],
        "tradeScope": current_scope,
        "targetSector": labels[idx],
        "shockLabel": shock_label,
        "demandShock": round(float(demand_shock), 2),
        "totalOutputEffect": round(tot_out, 2),
        "totalVaEffect": round(tot_va, 2),
        "targetOutputMultiplier": round(float(item["output_mult"][idx]), 4),
        "targetVaMultiplier": round(float(item["va_mult"][idx]), 4),
        "targetPowerIndex": round(float(item["power_idx"][idx]), 4),
        "targetSensitivityIndex": round(float(item["sens_idx"][idx]), 4),
        "targetLinkageType": _linkage_type(float(item["power_idx"][idx]), float(item["sens_idx"][idx])),
        "topRows": rows[:30],
        "crossLevel": cross_level,
    }, ensure_ascii=False)
