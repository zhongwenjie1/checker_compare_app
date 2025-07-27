import pandas as pd
import re

HEADER_ROW_STD = 16  # 0-based：第 17 行
USE_COLS_STD = 12


def _num2str(x):
    if pd.isna(x):
        return ""
    s = str(x)
    if re.fullmatch(r"\d+\.0", s):
        return s[:-2]
    return s


# 删除“看起来像表头”的数据行
def _drop_header_like_rows(df):
    """
    删除“看起来像表头”的数据行：
    条件：整行文本和列名一致的匹配数 >= 2；
    以及命中特定列等于列名（如 "BC POS" / "BC POS NAME"）。
    """
    import pandas as pd
    if df is None or df.empty:
        return df

    # 标准化：全部转字符串并去前后空格
    cols = pd.Index([str(c).strip() for c in df.columns])
    df_str = df.astype(str).apply(lambda s: s.str.strip())

    # 1) 通用规则：逐列比较，等于列名的计数≥2 视为伪表头
    eq_counts = df_str.eq(cols, axis=1).sum(axis=1)
    mask_header_like = eq_counts >= 2

    # 2) 兜底：明确针对这些列与其列名完全相等的行
    for col in ["BC POS", "BC POS NAME"]:
        if col in df_str.columns:
            mask_header_like |= df_str[col].str.fullmatch(col, case=False, na=False)

    # 过滤掉伪表头
    cleaned = df.loc[~mask_header_like].copy()
    # 去掉完全空白的行
    cleaned = cleaned.dropna(how="all")
    return cleaned

def load_std_df(path: str) -> pd.DataFrame:
    """
    读取标准文件：
      * 第 17 行作为表头（只取前 12 列）
      * 只保留 “最终判定 = Y” 的行
      * 生成 '__KEY__'（优先 BC POS NAME / Parts Name）
    """
    raw = pd.read_excel(path, header=None)
    headers = raw.iloc[HEADER_ROW_STD, :USE_COLS_STD].tolist()
    df = raw.iloc[HEADER_ROW_STD + 1:, :USE_COLS_STD].copy()
    df.columns = headers
    df = df.dropna(how="all")

    # 过滤“最终判定 = Y”
    decision_col = None
    for c in ["最终判定", "Final Decision", "Final decision", "Final"]:
        if c in df.columns:
            decision_col = c
            break
    if decision_col:
        df = df[df[decision_col].astype(str).str.upper().str.strip() == "Y"].copy()

    # 兜底必要列
    for col in ["品番", "组立番号"]:
        if col not in df.columns:
            df[col] = ""

    # KEY
    for c in ["BC POS NAME", "BC POS Name", "零件名称", "Parts Name"]:
        if c in df.columns:
            df["__KEY__"] = df[c].astype(str)
            break
    else:
        df["__KEY__"] = ""

    df.reset_index(drop=True, inplace=True)
    return df

def load_sys_df(path: str) -> pd.DataFrame:
    """
    读取系统文件（列名优先找，找不到按列号兜底）：
      H  BC POS
      I  BC POS NAME
      K  组立番号(前段)
      L  组立番号(后段, 不足 2 位补 0)
      M  是否上传
      N  品番
    """
    df_raw = pd.read_excel(path, header=0)

    def find_col(candidates, default_idx=None):
        for c in candidates:
            if c in df_raw.columns:
                return df_raw[c]
        if default_idx is not None and default_idx < df_raw.shape[1]:
            return df_raw.iloc[:, default_idx]
        return pd.Series([""] * len(df_raw))

    col_bc_pos      = find_col(["BC POS"], 7)
    col_bc_posname  = find_col(["BC POS NAME", "BC POS Name"], 8)
    col_k           = find_col(["组立番号K", "组立番号(K)"], 10)
    col_l           = find_col(["组立番号L", "组立番号(L)"], 11)
    col_upload      = find_col(["是否上传", "上传"], 12)
    col_partno      = find_col(["品番", "产品号", "Product Number"], 13)

    def _to_str(series):
        return series.map(_num2str).fillna("")

    col_k = _to_str(col_k)
    col_l = _to_str(col_l).str.zfill(2)

    df_use = pd.DataFrame({
        "BC POS":      _to_str(col_bc_pos),
        "BC POS NAME": _to_str(col_bc_posname),
        "组立番号":     (col_k + " " + col_l).str.strip(),
        "是否上传":     _to_str(col_upload),
        "品番":        _to_str(col_partno),
    })

    df_use["__KEY__"] = df_use["BC POS NAME"].astype(str)

    # 删除落入数据区的“伪表头”行（例如第一行再次出现列名）
    df_use = _drop_header_like_rows(df_use)

    # 清理空白行、去重并重建索引
    df_use = df_use.dropna(how="all")
    df_use = df_use.drop_duplicates().reset_index(drop=True)
    return df_use