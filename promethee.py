import numpy as np
import pandas as pd


class Promethee:

    def __init__(self, matrix: pd.DataFrame, weights: list):
        self.matrix = matrix
        self.weights = weights
        self.alternatives = list(matrix.index)
        self.criteria = list(matrix.columns)

    # Fonction de préférence simple (type usual)
    def preference(self, d):
        return 1 if d > 0 else 0

    # PHASE 1 : matrice de préférence π(a,b)
    def build_preference_matrix(self):

        n = len(self.alternatives)

        pref_matrix = pd.DataFrame(
            np.zeros((n, n)),
            index=self.alternatives,
            columns=self.alternatives
        )

        for a in self.alternatives:
            for b in self.alternatives:
                if a != b:
                    score = 0

                    for i, crit in enumerate(self.criteria):
                        d = self.matrix.loc[a, crit] - self.matrix.loc[b, crit]
                        score += self.weights[i] * self.preference(d)

                    pref_matrix.loc[a, b] = score

        return pref_matrix

    # PHASE 2 : flux + classement final
    def compute_flows(self, pref_matrix):

        n = len(self.alternatives)

        phi_plus = pref_matrix.sum(axis=1) / (n - 1)
        phi_minus = pref_matrix.sum(axis=0) / (n - 1)

        phi_net = phi_plus - phi_minus

        result = pd.DataFrame({
            "Phi+": phi_plus,
            "Phi-": phi_minus,
            "Phi net": phi_net
        })

        result = result.sort_values("Phi net", ascending=False)

        return result