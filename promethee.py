import pandas as pd


def parse_preferences(df: pd.DataFrame) -> pd.DataFrame:
    expected = ["Critère", "Poids", "Q", "P", "V"]
    missing = [c for c in expected if c not in df.columns]
    if missing:
        raise ValueError(f"Colonnes manquantes dans les préférences : {missing}")

    pref = df.copy()
    for col in ["Poids", "Q", "P", "V"]:
        pref[col] = pd.to_numeric(pref[col], errors="coerce").fillna(0.0)
    return pref


def validate_preferences(pref_df: pd.DataFrame, criteria_names) -> None:
    pref_crits = set(pref_df["Critère"].astype(str))
    missing = [c for c in criteria_names if c not in pref_crits]
    if missing:
        raise ValueError(f"Préférences manquantes pour les critères : {missing}")

    total_weight = float(pref_df["Poids"].sum())
    if total_weight <= 0:
        raise ValueError("La somme des poids des critères doit être strictement positive.")


def preference_degree(d: float, q: float, p: float) -> float:
    if d <= q:
        return 0.0
    if p <= q:
        return 1.0 if d > q else 0.0
    if d >= p:
        return 1.0
    return (d - q) / (p - q)


def compute_promethee_ii(matrix: pd.DataFrame, pref_df: pd.DataFrame):
    pref_df = parse_preferences(pref_df)
    validate_preferences(pref_df, list(matrix.columns))
    pref_df = pref_df.set_index("Critère").loc[list(matrix.columns)].reset_index()

    alt_names = list(matrix.index)
    n = len(alt_names)
    if n < 2:
        raise ValueError("Au moins deux alternatives sont nécessaires.")

    total_weight = float(pref_df["Poids"].sum())
    norm_weights = {
        row["Critère"]: float(row["Poids"]) / total_weight
        for _, row in pref_df.iterrows()
    }
    q_map = {row["Critère"]: float(row["Q"]) for _, row in pref_df.iterrows()}
    p_map = {row["Critère"]: float(row["P"]) for _, row in pref_df.iterrows()}

    pi = pd.DataFrame(0.0, index=alt_names, columns=alt_names)

    for a in alt_names:
        for b in alt_names:
            if a == b:
                continue

            agg_pref = 0.0
            for crit in matrix.columns:
                d = float(matrix.loc[a, crit]) - float(matrix.loc[b, crit])
                agg_pref += norm_weights[crit] * preference_degree(d, q_map[crit], p_map[crit])

            pi.loc[a, b] = agg_pref

    denom = n - 1
    phi_plus = pi.sum(axis=1) / denom
    phi_minus = pi.sum(axis=0) / denom
    phi = phi_plus - phi_minus

    results = pd.DataFrame({
        "Alternative": alt_names,
        "ϕ+": phi_plus.values,
        "ϕ-": phi_minus.values,
        "ϕ": phi.values,
    })

    results = results.sort_values(by="ϕ", ascending=False).reset_index(drop=True)
    results["Rang"] = range(1, len(results) + 1)

    return pi, results


def aggregate_decision_maker_results(dm_results):
    if not dm_results:
        raise ValueError("Aucun résultat de décideur à agréger.")

    total_dm_weight = sum(max(0.0, float(w)) for _, w, _ in dm_results)
    if total_dm_weight <= 0:
        raise ValueError("La somme des poids des décideurs doit être strictement positive.")

    base = dm_results[0][2][["Alternative"]].copy()
    base["ϕ+"] = 0.0
    base["ϕ-"] = 0.0
    base["ϕ"] = 0.0

    for _, dm_weight, df in dm_results:
        coeff = float(dm_weight) / total_dm_weight
        merged = df[["Alternative", "ϕ+", "ϕ-", "ϕ"]].copy()
        base = base.merge(merged, on="Alternative", suffixes=("", "_dm"))

        base["ϕ+"] += coeff * base["ϕ+_dm"]
        base["ϕ-"] += coeff * base["ϕ-_dm"]
        base["ϕ"] += coeff * base["ϕ_dm"]

        base = base.drop(columns=["ϕ+_dm", "ϕ-_dm", "ϕ_dm"])

    base = base.sort_values(by="ϕ", ascending=False).reset_index(drop=True)
    base["Rang"] = range(1, len(base) + 1)

    return base