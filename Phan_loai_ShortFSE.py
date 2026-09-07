"""
Financial Short/FSE Auto Filler - Hybrid Engine
Author: Mark (Do Duc Anh)

Muc tieu:
    Tu dong dien cot Short, FSE va FSE/Short vao file Target .xlsx
    dua tren Training model_Financial.xlsx.

Engine uu tien:
    1) Rule engine
    2) Exact match tren 4 cot: Description, Department Name, Account Name, Customer Name
    3) Fuzzy match bang RapidFuzz neu co cai thu vien rapidfuzz
    4) CatBoost neu co cai thu vien catboost
    5) scikit-learn fallback

Khong dung sentence-transformers.

Yeu cau cai dat khuyen nghi:
    pip install pandas openpyxl scikit-learn catboost rapidfuzz freesimplegui

Neu khong cai catboost/rapidfuzz, code van chay bang fallback scikit-learn:
    pip install pandas openpyxl scikit-learn freesimplegui

Cach chay:
    python financial_short_fse_filler_hybrid.py
"""

from __future__ import annotations

import os
import re
import shutil
import threading
from pathlib import Path
from typing import Callable, Optional, Tuple, Dict, Any

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.linear_model import LogisticRegression
from sklearn.pipeline import Pipeline

try:
    from rapidfuzz import process, fuzz
    RAPIDFUZZ_AVAILABLE = True
except ModuleNotFoundError:
    process = None
    fuzz = None
    RAPIDFUZZ_AVAILABLE = False

try:
    from catboost import CatBoostClassifier, Pool
    CATBOOST_AVAILABLE = True
except ModuleNotFoundError:
    CatBoostClassifier = None
    Pool = None
    CATBOOST_AVAILABLE = False

# ===== Thong tin hien thi tren UI =====
AUTHOR = "Mark BP Team"
PROGRAM_NAME = "Tự phân loại Short/FSE (Beta version)"
TRAINING_FILE_NAME = "Training model_Financial.xlsx"

FEATURE_COLS = ["Description", "Department Name", "Account Name", "Customer Name"]
TARGET_COLS = ["Short", "FSE"]
DERIVED_COL = "FSE/Short"
VALID_LABELS = {
    "Short": {"N", "Short"},
    "FSE": {"N", "FSE"},
}

# Threshold fuzzy. Neu muon gan hon exact-match thi tang len 97-99.
FUZZY_THRESHOLD = 96


def _progress(cb: Optional[Callable[[int, str], None]], pct: int, msg: str) -> None:
    if cb:
        cb(int(pct), msg)


def normalize_text(value) -> str:
    """Chuan hoa text de exact/fuzzy match on dinh hon."""
    if pd.isna(value):
        return ""
    text = str(value).strip().lower()
    text = text.replace("_", " ").replace("-", " ").replace("/", " ")
    text = re.sub(r"[\t\r\n]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def normalize_label(label: Any, target_col: str) -> str:
    """Dam bao label chi nam trong tap hop hop le."""
    if pd.isna(label):
        return "N"
    label = str(label).strip()
    allowed = VALID_LABELS[target_col]
    return label if label in allowed else "N"


def make_key_df(df: pd.DataFrame) -> pd.Series:
    """Tao exact-key tu 4 cot yeu cau."""
    return df[FEATURE_COLS].apply(lambda r: " ||| ".join(normalize_text(v) for v in r), axis=1)


def make_fuzzy_text_df(df: pd.DataFrame) -> pd.Series:
    """
    Tao chuoi fuzzy text co gan trong so don gian:
    Description quan trong nhat, sau do Account/Department/Customer.
    """
    desc = df["Description"].map(normalize_text)
    dept = df["Department Name"].map(normalize_text)
    acc = df["Account Name"].map(normalize_text)
    cust = df["Customer Name"].map(normalize_text)
    return (
        "desc " + desc + " desc " + desc +
        " account " + acc +
        " department " + dept +
        " customer " + cust
    )


def prepare_training_df(training_path: str | Path) -> pd.DataFrame:
    """
    Doc file training.
    File training hien tai co 1 dong trong o tren cung, nen thu doc skiprows=1.
    Neu file training sau nay khong co dong trong, function fallback ve cach doc binh thuong.
    """
    training_path = Path(training_path)
    df = pd.read_excel(training_path, skiprows=1, engine="openpyxl")

    if not set(FEATURE_COLS + TARGET_COLS).issubset(df.columns):
        df = pd.read_excel(training_path, engine="openpyxl")

    missing = [c for c in FEATURE_COLS + TARGET_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"File training thieu cot bat buoc: {missing}")

    df = df.dropna(how="all").copy()
    for col in TARGET_COLS:
        df[col] = df[col].apply(lambda x: normalize_label(x, col))

    return df.reset_index(drop=True)


def validate_target_df(target_df: pd.DataFrame) -> None:
    missing = [c for c in FEATURE_COLS + TARGET_COLS + [DERIVED_COL] if c not in target_df.columns]
    if missing:
        raise ValueError(f"File target thieu cot bat buoc: {missing}")


def combine_features(df: pd.DataFrame) -> pd.Series:
    """Ghep 4 cot thanh text input cho scikit-learn model."""
    safe = df.copy()
    for col in FEATURE_COLS:
        safe[col] = safe[col].fillna("").astype(str)

    return (
        "DESC: " + safe["Description"] +
        " | DEPT: " + safe["Department Name"] +
        " | ACC: " + safe["Account Name"] +
        " | CUST: " + safe["Customer Name"]
    )


def build_sklearn_model() -> Pipeline:
    """Model fallback bang scikit-learn."""
    return Pipeline([
        ("tfidf", TfidfVectorizer(
            analyzer="char_wb",
            ngram_range=(3, 5),
            min_df=1,
            lowercase=True,
            sublinear_tf=True,
        )),
        ("clf", LogisticRegression(
            max_iter=2000,
            class_weight="balanced",
            random_state=42,
            solver="liblinear",
        )),
    ])


def catboost_features(df: pd.DataFrame) -> pd.DataFrame:
    """Chuan bi 4 cot input dang string cho CatBoost."""
    features = df[FEATURE_COLS].copy()
    for col in FEATURE_COLS:
        features[col] = features[col].fillna("").astype(str)
    return features


def predict_with_catboost(training_df: pd.DataFrame, target_df: pd.DataFrame, label_col: str) -> np.ndarray:
    """Du doan bang CatBoostClassifier voi text_features."""
    x_train = catboost_features(training_df)
    y_train = training_df[label_col].apply(lambda x: normalize_label(x, label_col))
    x_target = catboost_features(target_df)

    train_pool = Pool(x_train, label=y_train, text_features=FEATURE_COLS)
    target_pool = Pool(x_target, text_features=FEATURE_COLS)

    model = CatBoostClassifier(
        iterations=350,
        learning_rate=0.08,
        depth=6,
        loss_function="MultiClass",
        auto_class_weights="Balanced",
        random_seed=42,
        verbose=False,
        allow_writing_files=False,
    )
    model.fit(train_pool)
    return model.predict(target_pool).reshape(-1).astype(str)


def predict_with_sklearn(training_df: pd.DataFrame, target_df: pd.DataFrame, label_col: str) -> np.ndarray:
    """Du doan bang scikit-learn TF-IDF + Logistic Regression."""
    x_train = combine_features(training_df)
    y_train = training_df[label_col].apply(lambda x: normalize_label(x, label_col))
    x_target = combine_features(target_df)

    model = build_sklearn_model()
    model.fit(x_train, y_train)
    return model.predict(x_target).astype(str)


def majority_lookup(training_df: pd.DataFrame, label_col: str) -> dict:
    """Tao lookup exact-key -> label xuat hien nhieu nhat trong training."""
    temp = pd.DataFrame({
        "key": make_key_df(training_df),
        "label": training_df[label_col].apply(lambda x: normalize_label(x, label_col)),
    })
    counts = temp.groupby(["key", "label"]).size().reset_index(name="n")
    counts = counts.sort_values(["key", "n"], ascending=[True, False])
    return counts.drop_duplicates("key").set_index("key")["label"].to_dict()


def build_fuzzy_training(training_df: pd.DataFrame, label_col: str) -> Tuple[list[str], dict[str, str]]:
    """
    Tao danh sach fuzzy choices va mapping choice -> label.
    Neu duplicate fuzzy text co nhieu label, lay label xuat hien nhieu nhat.
    """
    temp = pd.DataFrame({
        "fuzzy_text": make_fuzzy_text_df(training_df),
        "label": training_df[label_col].apply(lambda x: normalize_label(x, label_col)),
    })
    counts = temp.groupby(["fuzzy_text", "label"]).size().reset_index(name="n")
    counts = counts.sort_values(["fuzzy_text", "n"], ascending=[True, False])
    best = counts.drop_duplicates("fuzzy_text")
    mapping = dict(zip(best["fuzzy_text"], best["label"]))
    choices = list(mapping.keys())
    return choices, mapping


def fuzzy_predict_for_unmatched(
    target_df: pd.DataFrame,
    unmatched_mask: np.ndarray,
    training_df: pd.DataFrame,
    label_col: str,
    threshold: int = FUZZY_THRESHOLD,
) -> Tuple[np.ndarray, np.ndarray]:
    """
    Fuzzy match bang RapidFuzz cho cac dong chua exact match.
    Return:
        fuzzy_pred: array object cung length target, None neu khong pass threshold
        fuzzy_score: score similarity, -1 neu khong fuzzy
    """
    fuzzy_pred = np.array([None] * len(target_df), dtype=object)
    fuzzy_score = np.full(len(target_df), -1.0)

    if not RAPIDFUZZ_AVAILABLE or not unmatched_mask.any():
        return fuzzy_pred, fuzzy_score

    choices, label_map = build_fuzzy_training(training_df, label_col)
    if not choices:
        return fuzzy_pred, fuzzy_score

    target_texts = make_fuzzy_text_df(target_df)
    unmatched_indices = np.where(unmatched_mask)[0]

    for idx in unmatched_indices:
        result = process.extractOne(
            target_texts.iloc[idx],
            choices,
            scorer=fuzz.WRatio,
            score_cutoff=threshold,
        )
        if result is not None:
            match_text, score, _ = result
            fuzzy_pred[idx] = label_map.get(match_text)
            fuzzy_score[idx] = float(score)

    return fuzzy_pred, fuzzy_score


def apply_business_rules(
    target_df: pd.DataFrame,
    short_pred: np.ndarray,
    fse_pred: np.ndarray,
    method_short: np.ndarray,
    method_fse: np.ndarray,
    exact_unmatched_fse_mask: np.ndarray,
) -> Tuple[np.ndarray, np.ndarray, np.ndarray, np.ndarray]:
    """
    Rule engine.
    Luu y: Rule [FSE] chi override cac dong chua exact-match trong training
    de tranh lam sai cac dong historical/desired co ngoai le.
    """
    desc = target_df["Description"].fillna("").astype(str)

    # Rule bat buoc: Description co string [FSE] thi FSE = FSE.
    # Nhung exact-match training duoc uu tien cao hon vi training/desired co the co ngoai le.
    fse_bracket_mask = desc.str.contains(r"\[FSE\]", case=False, regex=True, na=False).to_numpy()
    rule_mask = fse_bracket_mask & exact_unmatched_fse_mask
    fse_pred[rule_mask] = "FSE"
    method_fse[rule_mask] = "Rule: [FSE]"

    return short_pred, fse_pred, method_short, method_fse


def predict_labels(
    target_df: pd.DataFrame,
    training_df: pd.DataFrame,
    progress_cb: Optional[Callable[[int, str], None]] = None,
) -> Tuple[np.ndarray, np.ndarray, pd.DataFrame]:
    """Du doan Short va FSE bang hybrid engine."""
    _progress(progress_cb, 25, "Đang tạo exact keys...")
    target_keys = make_key_df(target_df)

    predictions: Dict[str, np.ndarray] = {}
    methods: Dict[str, np.ndarray] = {}
    scores: Dict[str, np.ndarray] = {}
    exact_unmatched_masks: Dict[str, np.ndarray] = {}

    for idx, label_col in enumerate(TARGET_COLS, start=1):
        _progress(progress_cb, 30 + idx * 7, f"Exact match cho cột {label_col}...")
        lookup = majority_lookup(training_df, label_col)

        pred = target_keys.map(lookup).to_numpy(dtype=object)
        method = np.array(["Exact match" if not pd.isna(x) else "Unmatched" for x in pred], dtype=object)
        score = np.full(len(target_df), 100.0)
        unmatched_mask = pd.isna(pred)
        exact_unmatched_masks[label_col] = unmatched_mask.copy()

        _progress(progress_cb, 44 + idx * 7, f"Fuzzy match RapidFuzz cho cột {label_col}...")
        fuzzy_pred, fuzzy_score = fuzzy_predict_for_unmatched(target_df, unmatched_mask, training_df, label_col)
        fuzzy_mask = pd.notna(fuzzy_pred)
        pred[fuzzy_mask] = fuzzy_pred[fuzzy_mask]
        method[fuzzy_mask] = "Fuzzy match"
        score[fuzzy_mask] = fuzzy_score[fuzzy_mask]
        unmatched_mask = pd.isna(pred)

        if unmatched_mask.any():
            if CATBOOST_AVAILABLE:
                _progress(progress_cb, 58 + idx * 7, f"CatBoost fallback cho cột {label_col}...")
                all_pred = predict_with_catboost(training_df, target_df, label_col)
                pred[unmatched_mask] = all_pred[unmatched_mask]
                method[unmatched_mask] = "CatBoost fallback"
                score[unmatched_mask] = -1
            else:
                _progress(progress_cb, 58 + idx * 7, f"scikit-learn fallback cho cột {label_col}...")
                all_pred = predict_with_sklearn(training_df, target_df, label_col)
                pred[unmatched_mask] = all_pred[unmatched_mask]
                method[unmatched_mask] = "scikit-learn fallback"
                score[unmatched_mask] = -1

        pred = np.array([normalize_label(x, label_col) for x in pred], dtype=object)
        predictions[label_col] = pred
        methods[label_col] = method
        scores[label_col] = score

    _progress(progress_cb, 72, "Đang áp dụng business rules...")
    short_pred, fse_pred, method_short, method_fse = apply_business_rules(
        target_df=target_df,
        short_pred=predictions["Short"],
        fse_pred=predictions["FSE"],
        method_short=methods["Short"],
        method_fse=methods["FSE"],
        exact_unmatched_fse_mask=exact_unmatched_masks["FSE"],
    )

    audit_df = pd.DataFrame({
        "Row": np.arange(2, len(target_df) + 2),
        "Short_Pred": short_pred,
        "Short_Method": method_short,
        "Short_Score": scores["Short"],
        "FSE_Pred": fse_pred,
        "FSE_Method": method_fse,
        "FSE_Score": scores["FSE"],
        "Description": target_df["Description"].fillna(""),
        "Department Name": target_df["Department Name"].fillna(""),
        "Account Name": target_df["Account Name"].fillna(""),
        "Customer Name": target_df["Customer Name"].fillna(""),
    })

    return short_pred, fse_pred, audit_df


def find_training_file(target_path: str | Path) -> Path:
    """Tim file training trong folder target, folder script, va current working directory."""
    target_path = Path(target_path)
    candidates = [
        target_path.with_name(TRAINING_FILE_NAME),
        Path(__file__).resolve().with_name(TRAINING_FILE_NAME),
        Path.cwd() / TRAINING_FILE_NAME,
    ]
    for p in candidates:
        if p.exists():
            return p
    raise FileNotFoundError(
        f"Khong tim thay {TRAINING_FILE_NAME}. Hay dat file training cung folder voi file target hoac cung folder voi chuong trinh."
    )


def output_path_for(target_path: str | Path) -> Path:
    target_path = Path(target_path)
    return target_path.with_name(f"{target_path.stem}_filled{target_path.suffix}")


def fill_file(
    target_path: str | Path,
    training_path: str | Path | None = None,
    progress_cb: Optional[Callable[[int, str], None]] = None,
    create_audit_sheet: bool = True,
) -> Path:
    """Ham chinh: doc target, du doan, ghi ra file *_filled.xlsx."""
    target_path = Path(target_path)
    if training_path is None:
        training_path = find_training_file(target_path)
    training_path = Path(training_path)

    if not target_path.exists():
        raise FileNotFoundError(f"Khong tim thay file target: {target_path}")
    if not training_path.exists():
        raise FileNotFoundError(f"Khong tim thay file training: {training_path}")

    _progress(progress_cb, 5, "Đang đọc file Target...")
    target_df = pd.read_excel(target_path, engine="openpyxl")
    validate_target_df(target_df)

    _progress(progress_cb, 15, "Đang đọc file Training model...")
    training_df = prepare_training_df(training_path)

    short_pred, fse_pred, audit_df = predict_labels(target_df, training_df, progress_cb)
    fse_short = [f"{s}{f}" for s, f in zip(short_pred, fse_pred)]

    _progress(progress_cb, 82, "Đang sao chép file và ghi kết quả...")
    out_path = output_path_for(target_path)
    shutil.copy2(target_path, out_path)

    wb = load_workbook(out_path)
    ws = wb.active
    headers = [cell.value for cell in ws[1]]
    col_map = {name: i + 1 for i, name in enumerate(headers)}
    for col in TARGET_COLS + [DERIVED_COL]:
        if col not in col_map:
            raise ValueError(f"File target thieu cot bat buoc: {col}")

    for row_idx, (s, f, fs) in enumerate(zip(short_pred, fse_pred, fse_short), start=2):
        ws.cell(row=row_idx, column=col_map["Short"]).value = str(s)
        ws.cell(row=row_idx, column=col_map["FSE"]).value = str(f)
        ws.cell(row=row_idx, column=col_map[DERIVED_COL]).value = str(fs)

    if create_audit_sheet:
        _progress(progress_cb, 90, "Đang tạo sheet Audit_Log...")
        if "Audit_Log" in wb.sheetnames:
            del wb["Audit_Log"]
        audit_ws = wb.create_sheet("Audit_Log")
        audit_ws.append(list(audit_df.columns))
        for row in audit_df.itertuples(index=False, name=None):
            audit_ws.append(list(row))
        audit_ws.freeze_panes = "A2"
        audit_ws.auto_filter.ref = audit_ws.dimensions
        for col_cells in audit_ws.columns:
            max_len = 0
            col_letter = col_cells[0].column_letter
            for cell in col_cells[:200]:
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))
            audit_ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 45)

    wb.save(out_path)
    _progress(progress_cb, 100, f"Hoàn thành: {out_path.name}")
    return out_path


def engine_text() -> str:
    parts = []
    parts.append("RapidFuzz" if RAPIDFUZZ_AVAILABLE else "No RapidFuzz")
    parts.append("CatBoost" if CATBOOST_AVAILABLE else "No CatBoost")
    parts.append("scikit-learn fallback")
    return "Engine: " + " + ".join(parts)


def run_gui() -> None:
    """UI FreeSimpleGUI co hop thoai chon file va progress bar."""
    try:
        import FreeSimpleGUI as sg
    except ModuleNotFoundError as exc:
        raise ModuleNotFoundError(
            "Ban chua cai freesimplegui. Hay chay: pip install freesimplegui"
        ) from exc

    sg.theme("SystemDefault")
    layout = [
        [sg.Text(f"{PROGRAM_NAME}", font=("Arial", 14, "bold"))],
        [sg.Text(f"Author: {AUTHOR}")],
        [sg.Text(engine_text())],
        [sg.HorizontalSeparator()],
        [sg.Text("Chọn file Target (.xlsx):")],
        [sg.Input(key="-TARGET-", expand_x=True), sg.FileBrowse("Browse", file_types=(("Excel files", "*.xlsx"),))],
        [sg.Checkbox("Tạo sheet Audit_Log", default=True, key="-AUDIT-")],
        [sg.ProgressBar(100, orientation="h", size=(45, 18), key="-PROGRESS-")],
        [sg.Text("Sẵn sàng", key="-STATUS-", size=(80, 2))],
        [sg.Button("Run", key="-RUN-"), sg.Button("Exit")],
    ]
    window = sg.Window(PROGRAM_NAME, layout, finalize=True)

    def worker(target_file: str, create_audit: bool) -> None:
        try:
            def cb(pct: int, msg: str) -> None:
                window.write_event_value("-PROG-", (pct, msg))

            out = fill_file(target_file, progress_cb=cb, create_audit_sheet=create_audit)
            window.write_event_value("-DONE-", str(out))
        except Exception as e:
            window.write_event_value("-ERROR-", str(e))

    while True:
        event, values = window.read()
        if event in (sg.WIN_CLOSED, "Exit"):
            break
        if event == "-RUN-":
            target_file = values.get("-TARGET-")
            if not target_file:
                sg.popup_error("Vui lòng chọn file Target .xlsx")
                continue
            window["-RUN-"].update(disabled=True)
            window["-STATUS-"].update("Đang chạy...")
            window["-PROGRESS-"].update(0)
            threading.Thread(target=worker, args=(target_file, values.get("-AUDIT-", True)), daemon=True).start()
        elif event == "-PROG-":
            pct, msg = values[event]
            window["-PROGRESS-"].update(pct)
            window["-STATUS-"].update(msg)
        elif event == "-DONE-":
            out_file = values[event]
            window["-PROGRESS-"].update(100)
            window["-STATUS-"].update(f"Hoàn thành: {out_file}")
            window["-RUN-"].update(disabled=False)
            sg.popup_ok(f"Đã tạo file:\n{out_file}")
        elif event == "-ERROR-":
            window["-RUN-"].update(disabled=False)
            window["-STATUS-"].update("Có lỗi xảy ra!")
            sg.popup_error(values[event])

    window.close()


if __name__ == "__main__":
    run_gui()
