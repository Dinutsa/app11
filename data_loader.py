"""Модуль завантаження та попередньої підготовки даних опитувань."""

from __future__ import annotations
from dataclasses import dataclass
from typing import List, Tuple
import pandas as pd

@dataclass
class LoadedData:
    df: pd.DataFrame
    n_rows: int
    n_cols: int


def load_excels(files: List) -> LoadedData:

    if not files:
        raise ValueError("Список файлів порожній")

    frames = []
    for f in files:
        try:
            frame = pd.read_excel(f)
        except Exception as exc:
            raise ValueError(f"Не вдалося прочитати файл {getattr(f, 'name', f)}") from exc

        frames.append(frame)

    df = pd.concat(frames, ignore_index=True)

    df.columns = [str(c).strip() for c in df.columns]
    return LoadedData(df=df, n_rows=len(df), n_cols=len(df.columns))


def get_row_bounds(ld: LoadedData) -> Tuple[int, int]:
   
    if ld.n_rows == 0:
        return (0, 0)
    min_row = 2
    max_row = ld.n_rows + 1
    return min_row, max_row


def slice_range(ld: LoadedData, from_row: int, to_row: int) -> pd.DataFrame:
  
    if from_row > to_row:
        raise ValueError("Початковий номер рядка не може бути більшим за кінцевий")
    start_idx = from_row - 2
    end_idx = to_row - 2  # включно
    if start_idx < 0 or end_idx >= ld.n_rows:
        raise ValueError("Діапазон виходить за межі наявних відповідей")

    return ld.df.iloc[start_idx : end_idx + 1].copy()
