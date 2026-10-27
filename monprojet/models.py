"""Les modèles à comparer.

À vous d'écrire candidates(). Les tests de tests/test_models.py disent ce qui est attendu.
"""
from monprojet import config  # noqa: F401  (config.TASK, config.SEED)


def candidates() -> dict:
    """Renvoie un dictionnaire {nom: modèle scikit-learn non entraîné}.

    Attendu :
    - le PREMIER modèle est la baseline : le modèle le plus simple possible
      (DummyClassifier pour une classification, DummyRegressor pour une régression) ;
    - au moins trois vrais modèles ensuite, adaptés à config.TASK ;
    - random_state=config.SEED pour les modèles qui ont du hasard.

    Pistes : sklearn.dummy, sklearn.linear_model, sklearn.ensemble, sklearn.neighbors, sklearn.tree.
    """
    from sklearn.dummy import DummyClassifier, DummyRegressor
    from sklearn.ensemble import GradientBoostingClassifier, GradientBoostingRegressor, RandomForestClassifier, RandomForestRegressor
    from sklearn.linear_model import LogisticRegression, Ridge
    from sklearn.neighbors import KNeighborsClassifier, KNeighborsRegressor

    if config.TASK == "classification":
        return {
            "baseline": DummyClassifier(strategy="most_frequent"),
            "regression_logistique": LogisticRegression(max_iter=1000, random_state=config.SEED),
            "random_forest": RandomForestClassifier(n_estimators=200, random_state=config.SEED),
            "gradient_boosting": GradientBoostingClassifier(random_state=config.SEED),
            "knn": KNeighborsClassifier(),
        }
    return {
        "baseline": DummyRegressor(strategy="mean"),
        "ridge": Ridge(random_state=config.SEED),
        "random_forest": RandomForestRegressor(n_estimators=200, random_state=config.SEED),
        "gradient_boosting": GradientBoostingRegressor(random_state=config.SEED),
        "knn": KNeighborsRegressor(),
    }
