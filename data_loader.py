"""
Модуль завантаження та попередньої підготовки даних опитувань.

Використовуються підходи табличного подання соціологічних даних,
рекомендовані у вітчизняних методичних посібниках з обробки результатів
анкетування (див., наприклад, Кислова О.М., Кузіна І.І. 'Методи аналізу
та комп’ютерної обробки соціологічної інформації').
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import List, Tuple

import pandas as pd


@dataclass
class LoadedData:
    """Структура результату завантаження."""
    df: pd.DataFrame
    n_rows: int
    n_cols: int


def load_excels(files: List) -> LoadedData:
    """
    Завантажує один або кілька Excel-файлів Google Forms і об'єднує їх.

    :param files: список об’єктів UploadedFile (Streamlit) або file-like.
    :return: LoadedData з єдиним DataFrame.
    """
    if not files:
        raise ValueError("Список файлів порожній")

    frames = []
    for f in files:
        # Streamlit UploadedFile має метод getvalue()
        try:
            frame = pd.read_excel(f)
        except Exception as exc:
            raise ValueError(f"Не вдалося прочитати файл {getattr(f, 'name', f)}") from exc

        frames.append(frame)

    df = pd.concat(frames, ignore_index=True)

    # Очистимо заголовки стовпців
    df.columns = [str(c).strip() for c in df.columns]

    return LoadedData(df=df, n_rows=len(df), n_cols=len(df.columns))


def get_row_bounds(ld: LoadedData) -> Tuple[int, int]:
    """
    Повертає допустимі межі рядків для користувача (як у Excel):
    перший рядок із відповідями – 2 (1-й був заголовком у Google Forms).

    :return: (min_row, max_row) у «людській» нумерації.
    """
    if ld.n_rows == 0:
        return (0, 0)
    # У pandas перший рядок має індекс 0, але в Excel – 2 (після заголовка).
    min_row = 2
    max_row = ld.n_rows + 1
    return min_row, max_row


def slice_range(ld: LoadedData, from_row: int, to_row: int) -> pd.DataFrame:
    """
    Повертає підтаблицю для вказаного діапазону рядків (як у Excel).

    :param from_row: перший рядок (мінімум 2).
    :param to_row: останній рядок (включно).
    """
    if from_row > to_row:
        raise ValueError("Початковий номер рядка не може бути більшим за кінцевий")

    # перенесення в індекси pandas: рядок 2 → index 0
    start_idx = from_row - 2
    end_idx = to_row - 2  # включно
    if start_idx < 0 or end_idx >= ld.n_rows:
        raise ValueError("Діапазон виходить за межі наявних відповідей")

    return ld.df.iloc[start_idx : end_idx + 1].copy()
