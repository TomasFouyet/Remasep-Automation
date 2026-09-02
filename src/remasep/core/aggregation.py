from collections.abc import Iterable

import pandas as pd


def aggregate_long(
    classified_rows: Iterable[dict],
    *,
    dimensions: list[str],
    value_column: str = "valor",
) -> pd.DataFrame:
    frame = pd.DataFrame(classified_rows)

    if frame.empty:
        return pd.DataFrame(columns=[*dimensions, value_column])

    missing = [column for column in dimensions if column not in frame.columns]
    if missing:
        raise ValueError(f"Faltan dimensiones para agregar: {missing}")

    if value_column not in frame.columns:
        frame[value_column] = 1

    return (
        frame.groupby(dimensions, dropna=False, as_index=False)[value_column]
        .sum()
        .sort_values(dimensions)
        .reset_index(drop=True)
    )
