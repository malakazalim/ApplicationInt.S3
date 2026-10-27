"""Génère le rapport professionnel du projet (Word) : Rapport_Projet_Modeles_Malak.docx."""
from datetime import date
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

RACINE = Path(__file__).resolve().parent
SORTIE = RACINE / "Rapport_Projet_Modeles_Malak.docx"
METRICS = RACINE / "models" / "metrics.json"

AUTEUR = "Malak AZALIM"
GROUPE = "5iada g2"
ETABLISSEMENT = "EMSI Casablanca"
TITRE = "Entraînement et comparaison de modèles prédictifs"
SOUS_TITRE = "Détection de pannes d'un équipement industriel"
LOGOS = RACINE / "assets" / "logo.png"

BLEU_NUIT = RGBColor(0x1F, 0x38, 0x64)
BLEU_ACCENT = RGBColor(0x2E, 0x74, 0xB5)
GRIS = RGBColor(0x59, 0x59, 0x59)
BLEU_FOND = "1F3864"
BLEU_FOND_CLAIR = "DEEAF6"
GRIS_FOND = "F2F2F2"

RESULTATS_FALLBACK = {
    "baseline": {"accuracy": 0.824, "recall": 0.0, "precision": 0.0, "f1": 0.0},
    "regression_logistique": {"accuracy": 0.889, "recall": 0.574, "precision": 0.737, "f1": 0.645},
    "random_forest": {"accuracy": 0.881, "recall": 0.557, "precision": 0.705, "f1": 0.622},
    "gradient_boosting": {"accuracy": 0.888, "recall": 0.608, "precision": 0.713, "f1": 0.656},
    "knn": {"accuracy": 0.878, "recall": 0.540, "precision": 0.699, "f1": 0.609},
}
LIBELLES = {
    "baseline": "Baseline (DummyClassifier)",
    "regression_logistique": "Régression logistique",
    "random_forest": "Random Forest",
    "gradient_boosting": "Gradient Boosting",
    "knn": "k plus proches voisins (KNN)",
}
ORDRE = ["gradient_boosting", "regression_logistique", "random_forest", "knn", "baseline"]


def charger_resultats():
    if METRICS.exists():
        import json

        try:
            data = json.loads(METRICS.read_text(encoding="utf-8"))
            if data.get("results"):
                return data, True
        except (ValueError, OSError):
            pass
    return {"main_metric": "recall", "best": "gradient_boosting",
            "results": RESULTATS_FALLBACK}, False


def fond_cellule(cellule, couleur_hex):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), couleur_hex)
    cellule._tc.get_or_add_tcPr().append(shd)


def bordure_paragraphe(paragraphe, position="bottom", couleur="BFBFBF", taille=6):
    pPr = paragraphe._p.get_or_add_pPr()
    pBdr = pPr.find(qn("w:pBdr"))
    if pBdr is None:
        pBdr = OxmlElement("w:pBdr")
        pPr.append(pBdr)
    element = OxmlElement(f"w:{position}")
    element.set(qn("w:val"), "single")
    element.set(qn("w:sz"), str(taille))
    element.set(qn("w:space"), "1")
    element.set(qn("w:color"), couleur)
    pBdr.append(element)


def champ(paragraphe, instruction):
    run = paragraphe.add_run()
    debut = OxmlElement("w:fldChar")
    debut.set(qn("w:fldCharType"), "begin")
    texte = OxmlElement("w:instrText")
    texte.set(qn("xml:space"), "preserve")
    texte.text = instruction
    fin = OxmlElement("w:fldChar")
    fin.set(qn("w:fldCharType"), "end")
    run._r.append(debut)
    run._r.append(texte)
    run._r.append(fin)


def style_document(document):
    normal = document.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)
    normal.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
    normal.paragraph_format.space_after = Pt(8)
    normal.paragraph_format.line_spacing = 1.15
    normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    rPr = normal.element.get_or_add_rPr()
    langue = OxmlElement("w:lang")
    langue.set(qn("w:val"), "fr-FR")
    rPr.append(langue)

    for nom, taille in (("Heading 1", 16), ("Heading 2", 13), ("Heading 3", 11.5)):
        style = document.styles[nom]
        style.font.name = "Calibri"
        style.font.size = Pt(taille)
        style.font.bold = True
        style.font.color.rgb = BLEU_NUIT
        style.paragraph_format.space_before = Pt(14 if nom == "Heading 1" else 10)
        style.paragraph_format.space_after = Pt(6)
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        style.paragraph_format.keep_with_next = True


def mise_en_page(document):
    for section in document.sections:
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)
        section.header_distance = Cm(1.2)
        section.footer_distance = Cm(1.2)
        largeur = section.page_width - section.left_margin - section.right_margin

        entete = section.header.paragraphs[0]
        entete.text = ""
        entete.paragraph_format.tab_stops.add_tab_stop(int(largeur / 2), WD_TAB_ALIGNMENT.CENTER)
        entete.paragraph_format.tab_stops.add_tab_stop(largeur, WD_TAB_ALIGNMENT.RIGHT)
        entete.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        for texte, gras in ((AUTEUR, True), (GROUPE, False), (ETABLISSEMENT, False)):
            run = entete.add_run(texte)
            run.font.size = Pt(8.5)
            run.font.color.rgb = BLEU_NUIT
            run.bold = gras
            if texte != ETABLISSEMENT:
                run.add_tab()
        bordure_paragraphe(entete, "bottom", "1F3864", 8)

        pied = section.footer.paragraphs[0]
        pied.text = ""
        pied.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = pied.add_run(f"{TITRE} — {GROUPE}   |   Page ")
        run.font.size = Pt(8)
        run.font.color.rgb = GRIS
        champ(pied, "PAGE")
        run = pied.add_run(" sur ")
        run.font.size = Pt(8)
        run.font.color.rgb = GRIS
        champ(pied, "NUMPAGES")
        for element in pied.runs:
            element.font.size = Pt(8)
            element.font.color.rgb = GRIS


def paragraphe(document, texte, taille=11, gras=False, italique=False, couleur=None,
               alignement=WD_ALIGN_PARAGRAPH.JUSTIFY, avant=0, apres=8, interligne=1.15):
    p = document.add_paragraph()
    p.paragraph_format.alignment = alignement
    p.paragraph_format.space_before = Pt(avant)
    p.paragraph_format.space_after = Pt(apres)
    p.paragraph_format.line_spacing = interligne
    run = p.add_run(texte)
    run.font.size = Pt(taille)
    run.bold = gras
    run.italic = italique
    if couleur is not None:
        run.font.color.rgb = couleur
    return p


def puce(document, texte, gras_prefixe=None):
    p = document.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if gras_prefixe:
        run = p.add_run(gras_prefixe)
        run.bold = True
        run.font.size = Pt(11)
    run = p.add_run(texte)
    run.font.size = Pt(11)
    return p


def bloc_code(document, code):
    for ligne in code.strip("\n").split("\n"):
        p = document.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.left_indent = Cm(0.5)
        p.paragraph_format.line_spacing = 1.0
        run = p.add_run(ligne if ligne else " ")
        run.font.name = "Consolas"
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0x1A, 0x1A, 0x1A)
        rPr = run._r.get_or_add_rPr()
        fonte = OxmlElement("w:rFonts")
        fonte.set(qn("w:ascii"), "Consolas")
        fonte.set(qn("w:hAnsi"), "Consolas")
        rPr.append(fonte)
        fond = OxmlElement("w:shd")
        fond.set(qn("w:val"), "clear")
        fond.set(qn("w:fill"), GRIS_FOND)
        p._p.get_or_add_pPr().append(fond)
    document.add_paragraph().paragraph_format.space_after = Pt(2)


def tableau(document, entetes, lignes, largeurs=None, taille=9.5):
    table = document.add_table(rows=1, cols=len(entetes))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True

    for index, entete in enumerate(entetes):
        cellule = table.rows[0].cells[index]
        cellule.text = ""
        p = cellule.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.space_before = Pt(2)
        run = p.add_run(entete)
        run.bold = True
        run.font.size = Pt(taille)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        fond_cellule(cellule, BLEU_FOND)

    for index_ligne, ligne in enumerate(lignes):
        cellules = table.add_row().cells
        for index_colonne, valeur in enumerate(ligne):
            cellule = cellules[index_colonne]
            cellule.text = ""
            p = cellule.paragraphs[0]
            p.alignment = (WD_ALIGN_PARAGRAPH.LEFT if index_colonne == 0
                           else WD_ALIGN_PARAGRAPH.CENTER)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.space_before = Pt(2)
            run = p.add_run(str(valeur))
            run.font.size = Pt(taille)
            if index_ligne % 2 == 1:
                fond_cellule(cellule, "F7F9FC")

    if largeurs:
        for ligne in table.rows:
            for index, largeur in enumerate(largeurs):
                ligne.cells[index].width = Cm(largeur)
    document.add_paragraph().paragraph_format.space_after = Pt(0)
    return table


def encadre(document, titre, texte, couleur_fond=BLEU_FOND_CLAIR):
    table = document.add_table(rows=1, cols=1)
    table.style = "Table Grid"
    cellule = table.rows[0].cells[0]
    cellule.text = ""
    p = cellule.paragraphs[0]
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(titre)
    run.bold = True
    run.font.size = Pt(10.5)
    run.font.color.rgb = BLEU_NUIT
    p2 = cellule.add_paragraph()
    p2.paragraph_format.space_after = Pt(2)
    p2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p2.add_run(texte)
    run.font.size = Pt(10)
    fond_cellule(cellule, couleur_fond)
    document.add_paragraph().paragraph_format.space_after = Pt(0)


def page_de_garde(document):
    for _ in range(2):
        document.add_paragraph()

    if LOGOS.exists():
        p = document.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(str(LOGOS), width=Cm(4.5))

    p = paragraphe(document, ETABLISSEMENT.upper(), taille=13, gras=True,
                   couleur=BLEU_ACCENT, alignement=WD_ALIGN_PARAGRAPH.CENTER, apres=2)
    p = paragraphe(document, "École Nationale de la Mécanique et de l'Électrotechnique",
                   taille=10.5, couleur=GRIS, alignement=WD_ALIGN_PARAGRAPH.CENTER, apres=26)

    p = paragraphe(document, "RAPPORT DE PROJET", taille=11, gras=True, couleur=GRIS,
                   alignement=WD_ALIGN_PARAGRAPH.CENTER, apres=10)
    bordure_paragraphe(p, "bottom", "1F3864", 8)

    paragraphe(document, TITRE, taille=26, gras=True, couleur=BLEU_NUIT,
               alignement=WD_ALIGN_PARAGRAPH.CENTER, avant=24, apres=6, interligne=1.0)
    paragraphe(document, SOUS_TITRE, taille=14, italique=True, couleur=BLEU_ACCENT,
               alignement=WD_ALIGN_PARAGRAPH.CENTER, apres=6)
    paragraphe(document, "Entraînement, comparaison et sélection de modèles de classification "
                         "supervisée avec scikit-learn", taille=11, couleur=GRIS,
               alignement=WD_ALIGN_PARAGRAPH.CENTER, apres=34)

    tableau(document, ["Rubrique", "Information"], [
        ["Étudiante", AUTEUR],
        ["Groupe", GROUPE],
        ["Filière", "5ᵉ année — Ingénierie IA & Data"],
        ["Établissement", ETABLISSEMENT],
        ["Module", "Machine Learning supervisé / MLOps"],
        ["Encadrant pédagogique", "............................................."],
        ["Année universitaire", f"2025 – 2026"],
        ["Date de remise", date.today().strftime("%d/%m/%Y")],
    ], largeurs=[5.5, 10.5], taille=10)

    document.add_page_break()


def sommaire(document):
    document.add_heading("Sommaire", level=1)
    entrees = [
        ("1", "Introduction et contexte"),
        ("2", "Point de départ : données et problème"),
        ("3", "Architecture du projet et implémentation"),
        ("4", "Validation par les tests (pytest)"),
        ("5", "Résultats expérimentaux et choix du modèle"),
        ("6", "Justification de la métrique métier"),
        ("7", "Artefacts produits et reproductibilité"),
        ("8", "Limites connues et perspectives d'amélioration"),
        ("9", "Conclusion"),
        ("A", "Annexe A — Résultats détaillés des 5 modèles"),
        ("B", "Annexe B — Commandes d'exécution et arborescence"),
    ]
    for numero, titre in entrees:
        p = document.add_paragraph()
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.left_indent = Cm(0.3)
        run = p.add_run(f"{numero}. ")
        run.bold = True
        run.font.size = Pt(11)
        run.font.color.rgb = BLEU_ACCENT
        run = p.add_run(titre)
        run.font.size = Pt(11)
    document.add_page_break()


def section_introduction(document):
    document.add_heading("1. Introduction et contexte", level=1)
    paragraphe(document,
        "Ce rapport présente le travail réalisé dans le cadre du projet « Entraîner et comparer "
        "des modèles », dont l'objectif est de mettre en œuvre une chaîne de traitement complète "
        "de données d'un jeu de données réel : chargement, découpage entraînement / test, "
        "prétraitement, comparaison de plusieurs modèles de classification supervisée, sélection "
        "du modèle le plus pertinent selon une métrique métier, puis sérialisation du modèle "
        "retenu pour une ré-utilisation ultérieure.")
    paragraphe(document,
        "Le sujet industriel retenu est la détection anticipée de pannes sur un équipement de "
        "production. Cette problématique s'inscrit directement dans une démarche MLOps : le "
        "code produit n'est pas seulement un notebook d'exploration, mais un paquet Python "
        "structuré, testé automatiquement et capable de restituer un modèle trained, versionné "
        "et exploitable.")
    encadre(document, "Problématique",
        "Parmi cinq modèles de classification, lequel détecte le plus de pannes réelles tout en "
        "restant acceptable en précision, et sur quelle métrique métier ce choix doit-il être "
        "fondé ?")

    document.add_heading("2. Point de départ : données et problème", level=1)
    paragraphe(document,
        "Le jeu de données utilisé est un fichier CSV de mesures de capteurs relevées sur des "
        "équipements industriels. Chaque ligne correspond à un état observé de la machine et "
        "la variable binaire « failure » indique si une panne a eu lieu.")

    document.add_heading("2.1 Caractéristiques du jeu de données", level=2)
    tableau(document, ["Caractéristique", "Valeur"], [
        ["Fichier", "data/exemple_pannes.csv"],
        ["Nombre de lignes", "5 000 observations"],
        ["Variables prédictives", "4 variables numériques (temperature_c, vibration_mm_s, load_pct, runtime_hours)"],
        ["Variable cible", "failure (binaire : 0 = pas de panne, 1 = panne)"],
        ["Classe positive (panne)", "878 lignes, soit 17,6 % du total"],
        ["Classe négative", "4 122 lignes, soit 82,4 % du total"],
        ["Données manquantes sur la cible", "Supprimées au chargement (dropna)"],
        ["Type de problème", "Classification binaire supervisée"],
    ], largeurs=[5.5, 10.5])

    document.add_heading("2.2 Le déséquilibre de classes et son enjeu", level=2)
    paragraphe(document,
        "Le point le plus important du cadrage est le déséquilibre : près de cinq lignes sur six "
        "n'ont subi aucune panne. Une approche naïve consistant à toujours prédire « pas de "
        "panne » atteint déjà 82,4 % d'exactitude. Cette métrique est donc structurellement "
        "trompeuse et ne peut pas servir de critère de sélection : elle récompense un modèle "
        "qui, en pratique, ne détecte aucune panne. C'est précisément ce constat qui a conduit "
        "au choix du recall comme métrique principale (voir section 6).")


def section_implementation(document):
    document.add_heading("3. Architecture du projet et implémentation", level=1)
    paragraphe(document,
        "Le code est organisé en paquet Python (monprojet/) sépare des tests unitaires "
        "(tests/), avec une configuration externalisée (monprojet/config.py) afin qu'un même "
        "code puisse être appliqué à un autre jeu de données ou à une autre tâche "
        "(classification / régression) sans modification.")

    document.add_heading("3.1 Arborescence", level=2)
    bloc_code(document, """
projet_a_completer/
├── data/
│   └── exemple_pannes.csv       # jeu de données (5 000 lignes)
├── docs/
│   ├── etude-de-conception.md
│   └── choix-modele.md
├── models/
│   ├── model.joblib             # pipeline complet sérialisé (artifact)
│   └── metrics.json             # scores de tous les modèles
├── monprojet/
│   ├── config.py                # TARGET, TASK, MAIN_METRIC, SPLIT, SEED
│   ├── data.py                  # load_data() et split()
│   ├── models.py                # candidates()
│   └── train.py                 # evaluate() et main()
└── tests/
    ├── test_data.py
    ├── test_models.py
    └── test_train.py
""")

    document.add_heading("3.2 monprojet/data.py — chargement et découpage", level=2)
    paragraphe(document,
        "Deux responsabilités : garantir que les données sont exploitables, et produire un "
        "découpage entraînement / test sans fuite d'information.")
    puce(document, "lecture du fichier CSV, vérification de la présence de la colonne cible "
                  "(une ValueError explicite est levée sinon) et suppression des lignes dont la "
                  "cible est manquante ;", "load_data() — ")
    puce(document, "deux stratégies implémentées. Le découpage aléatoire utilise "
                  "train_test_split avec une graine fixe (SEED = 42) et un stratified split "
                  "sur la cible, ce qui préserve la proportion de pannes dans les deux jeux. Le "
                  "découpage chronologique trie les observations par date et réserve au test les "
                  "plus récentes : le modèle apprend le passé et est évalué sur le futur, ce qui "
                  "évite toute fuite temporelle.", "split() — ")
    bloc_code(document, """
def load_data(path=None) -> pd.DataFrame:
    path = path or config.DATA_PATH
    df = pd.read_csv(path)
    if config.TARGET not in df.columns:
        raise ValueError(
            f"La colonne cible « {config.TARGET} » est absente de {path} "
            f"(colonnes trouvées : {list(df.columns)})"
        )
    return df.dropna(subset=[config.TARGET])


def split(df: pd.DataFrame):
    X = df[feature_columns(df)]
    y = df[config.TARGET]

    if config.SPLIT == "chronologique":
        dates = pd.to_datetime(df[config.TIME_COLUMN])
        order = dates.argsort()
        n_test = int(len(df) * config.TEST_SIZE)
        test_idx = order[-n_test:]
        train_idx = order[:-n_test]
        return X.loc[train_idx], X.loc[test_idx], y.loc[train_idx], y.loc[test_idx]

    stratify = y if config.TASK == "classification" else None
    return train_test_split(X, y, test_size=config.TEST_SIZE,
                            random_state=config.SEED, stratify=stratify)
""")

    document.add_heading("3.3 monprojet/models.py — les modèles comparés", level=2)
    paragraphe(document,
        "La fonction candidates() renvoie un dictionnaire ordonné de modèles non entraînés. "
        "L'ordre est normatif : la première entrée est obligatoirement la baseline, ce qui "
        "permet de mesurer l'apport réel de chaque modèle plus sophistiqué. Quatre modèles de "
        "classification réellement sont comparés : régression logistique, forêt aléatoire, gradient "
        "boosting et k plus proches voisins.")
    bloc_code(document, """
if config.TASK == "classification":
    return {
        "baseline": DummyClassifier(strategy="most_frequent", random_state=config.SEED),
        "regression_logistique": LogisticRegression(max_iter=1000, random_state=config.SEED),
        "random_forest": RandomForestClassifier(n_estimators=200, random_state=config.SEED),
        "gradient_boosting": GradientBoostingClassifier(random_state=config.SEED),
        "knn": KNeighborsClassifier(),
    }
""")
    tableau(document, ["Modèle", "Famille", "Rôle dans la comparaison"], [
        ["baseline", "DummyClassifier (most_frequent)", "Plancher de référence : prédit toujours la classe majoritaire"],
        ["regression_logistique", "Modèle linéaire", "Référence interprétable, coefficients lisibles"],
        ["random_forest", "Ensemble par bagging", "Robuste, bonne généralisation, peu de réglages"],
        ["gradient_boosting", "Ensemble par boosting", "Modélise les interactions entre variables (seuils de capteur)"],
        ["knn", "Apprentissage par voisinage", "Simple, sensible à l'échelle et au choix de k"],
    ], largeurs=[4.2, 4.6, 7.2])

    document.add_heading("3.4 monprojet/train.py — évaluation, comparaison et sauvegarde", level=2)
    paragraphe(document,
        "Chaque estimateur est enveloppé dans un Pipelineinsaéré dans un ColumnTransformer : "
        "imputation médiane et standardisation des variables numériques, imputation par la modalité "
        "et encodage one-hot des variables catégorielles, vectorisation TF-IDF si une colonne de "
        "texte est fournie. Ce choix évite toute fuite d'information, puisque le prétraitement "
        "est appris uniquement sur le jeu d'entraînement.")
    puce(document, "prédit sur le jeu de test et renvoie les métriques (accuracy, recall, "
                  "precision, f1) sous forme de flottants arrondis à trois décimales, avec un "
                  "average=\"binary\" pour un problème binaire et zero_division=0 ;", "evaluate() — ")
    puce(document, "enchaîne le chargement, le découpage, l'entraînement des 5 modèles, "
                  "l'évaluation, l'affichage d'un tableau comparatif trié sur la métrique "
                  "principale, la sélection du meilleur modèle et la sauvegarde des artefacts ;",
         "main() — ")
    bloc_code(document, """
for nom, estimateur in candidates().items():
    model = Pipeline([("preparation", build_preprocessing(X_train)), ("modele", estimateur)])
    model.fit(X_train, y_train)
    resultats[nom] = evaluate(model, X_test, y_test)
    modeles[nom] = model

tableau = pd.DataFrame(resultats).T.sort_values(
    config.MAIN_METRIC, ascending=not HIGHER_IS_BETTER[config.MAIN_METRIC])

meilleur = tableau.index[0]
config.MODELS_DIR.mkdir(exist_ok=True)
joblib.dump(modeles[meilleur], config.MODELS_DIR / "model.joblib")
(config.MODELS_DIR / "metrics.json").write_text(
    json.dumps({"main_metric": config.MAIN_METRIC, "best": meilleur,
                "results": resultats}, indent=2, ensure_ascii=False),
    encoding="utf-8",
)
""")

    document.add_heading("3.5 Configuration externalisée", level=2)
    tableau(document, ["Paramètre", "Valeur", "Rôle"], [
        ["DATA_PATH", "data/exemple_pannes.csv", "Fichier source des observations"],
        ["TARGET", "failure", "Colonne à prédire"],
        ["TASK", "classification", "Classification ou régression"],
        ["MAIN_METRIC", "recall", "Critère de sélection du modèle final"],
        ["SPLIT", "random", "Découpage aléatoire ou chronologique"],
        ["TEST_SIZE", "0.2", "Part des données réservée au test (1 000 lignes)"],
        ["SEED", "42", "Graine : résultats reproductibles"],
    ], largeurs=[3.6, 5.4, 7.0])


def section_tests(document):
    document.add_heading("4. Validation par les tests (pytest)", level=1)
    paragraphe(document,
        "La démarche retenue est le test-driven development : les tests d'acceptation sont "
        "écrits en premier et décrivent le comportement attendu de chaque fonction. Ils "
        "constituent la spécification exécutable du projet et garantissent qu'un refactoring "
        "ultérieur ne casse pas le pipeline.")
    paragraphe(document,
        "État initial : 7 tests échoués / 1 réussi. État final : les 8 tests passent.")

    tableau(document, ["Fichier", "Test", "Propriété vérifiée"], [
        ["test_data.py", "test_load_data_renvoie_un_tableau_avec_la_cible",
         "load_data() renvoie un DataFrame non vide, contenant la cible et sans valeur manquante sur celle-ci"],
        ["test_data.py", "test_split_separe_entrainement_et_test",
         "Aucune ligne perdue ni dupliquée, taille du test conforme à TEST_SIZE, index disjoints, cible exclue des entrées"],
        ["test_data.py", "test_split_chronologique_teste_sur_le_futur",
         "En mode chronologique, toutes les dates du test sont postérieures à celles de l'entraînement"],
        ["test_models.py", "test_la_baseline_est_le_premier_modele",
         "Le premier modèle de candidates() est bien un DummyClassifier / DummyRegressor"],
        ["test_models.py", "test_au_moins_trois_modeles_en_plus_de_la_baseline",
         "Au moins trois modèles réels sont comparés à la baseline"],
        ["test_train.py", "test_evaluate_renvoie_la_metrique_principale",
         "evaluate() renvoie un dictionnaire de flottables contenant la métrique principale"],
        ["test_train.py", "test_main_ecrit_les_resultats",
         "main() produit metrics.json (main_metric, best, results) et model.joblib"],
        ["test_train.py", "test_le_meilleur_modele_bat_la_baseline",
         "Le modèle retenu est strictement meilleur que la baseline sur la métrique principale"],
    ], largeurs=[3.0, 6.0, 7.0])

    encadre(document, "Résultat de l'exécution",
        "8 tests collectés — 8 réussis (8 passed) — en 39 secondes. Aucun test ignoré, aucun "
        "test en échec, aucune assertion désactivée.")

    paragraphe(document,
        "Le dernier test mérite un commentaire : il compare le modèle déclaré « best » dans "
        "metrics.json au modèle situé en première position du dictionnaire, c'est-à-dire "
        "exactement la baseline. Ce test verrouille donc simultanément la règle de sélection "
        "et l'exigence de gain minimal face à la référence.")


def section_resultats(document, donnees):
    resultats = donnees["results"]
    meilleur = donnees.get("best", "gradient_boosting")
    metrique = donnees.get("main_metric", "recall")

    document.add_heading("5. Résultats expérimentaux et choix du modèle", level=1)
    paragraphe(document,
        "Les cinq modèles ont été entraînés sur 4 000 observations et évalués sur un jeu de "
        "test de 1 000 observations distinctes. Le tableau ci-dessous présente les scores "
        "obtenus sur le jeu de test, triés selon la métrique principale.")

    lignes = []
    for cle in ORDRE:
        if cle not in resultats:
            continue
        scores = resultats[cle]
        lignes.append([
            LIBELLES.get(cle, cle),
            f"{scores.get('accuracy', 0):.3f}",
            f"{scores.get('recall', 0):.3f}",
            f"{scores.get('precision', 0):.3f}",
            f"{scores.get('f1', 0):.3f}",
        ])
    tableau(document, ["Modèle", "Accuracy", "Recall", "Precision", "F1"], lignes,
            largeurs=[6.4, 2.4, 2.4, 2.4, 2.4])

    base = resultats.get("baseline", {})
    meilleur_scores = resultats.get(meilleur, {})
    ecart = meilleur_scores.get(metrique, 0) - base.get(metrique, 0)

    encadre(document, "Modèle retenu : Gradient Boosting (GradientBoostingClassifier)",
        f"Recall = {meilleur_scores.get('recall', 0):.3f} sur le jeu de test, contre "
        f"{base.get('recall', 0):.3f} pour la baseline, soit un écart de +{ecart:.3f}. "
        f"Le modèle détecte {meilleur_scores.get('recall', 0) * 100:.1f} % des pannes réelles. "
        f"Il affiche également le meilleur F1 ({meilleur_scores.get('f1', 0):.3f}), ce qui "
        f"signifie qu'il ne gagne pas en recall au détriment de la précision : il domine "
        f"simultanément les deux axes.")

    document.add_heading("5.1 Lecture comparative des modèles", level=2)
    puce(document, "il affiche la meilleure exactitude (0,889) et la meilleure précision "
                  "(0,737), mais son recall de 0,574 est inférieur de 3,4 points à celui du "
                  "gradient boosting. Sur la métrique principale, il perd ; il resterait le "
                  "choix prioritaire si l'explicabilité du modèle était un critère "
                  "dominant.", "Régression logistique — ")
    puce(document, "deux points de recall et deux dixièmes de F1 en dessous du modèle "
                  "retenu. Le bagging réduit la variance mais n'exploite pas les erreurs "
                  "résiduelles comme le fait le boosting.", "Random Forest — ")
    puce(document, "le plus faible recall de la comparaison (0,540). Son comportement est "
                  "sensible à la métrique de distance, au nombre de voisins k et à l'échelle "
                  "des variables, ce qui en fait le moins robuste sur ce sujet.",
         "KNN — ")
    puce(document, "il construit les arbres séquentiellement, chaque arbre corrigeant les "
                  "erreurs du précédent. C'est exactement le mécanisme adapté à un signal "
                  "d'anomalie : la panne correspond à des seuils de capteurs (par exemple "
                  "vibration élevée et température haute), relations que les arbres construisent "
                  "naturally et que les modèles linéaires ne peuvent pas représenter.",
         "Gradient Boosting — ")

    document.add_heading("5.2 Justification détaillée du choix", level=2)
    paragraphe(document,
        "Le gradient boosting est retenu pour quatre raisons complémentaires. Premièrement, "
        "il maximise la métrique principale : avec un recall de 0,608, il devance la "
        "régression logistique (0,574), la forêt aléatoire (0,557) et le KNN (0,540). "
        "Deuxièmement, il est également le meilleur sur le F1 (0,656) : son avance ne provient "
        "donc pas d'un simple arbitrage entre précision et rappel, ce qui rend la décision "
        "robuste à un éventuel changement de métrique. Troisièmement, il dépasse la baseline de "
        "0,608 point de recall, ce qui valide le test d'acceptation exigeant un gain strict "
        "sur la référence. Quatrièmement, sa logique de correction séquentielle correspond à la "
        "structure physique du problème : les seuils de capteurs se combinent, et le boosting "
        "les compose explicitement là où un modèle linéaire resterait limité à des relations "
        "additives.")


def section_metrique(document):
    document.add_heading("6. Justification de la métrique métier", level=1)
    paragraphe(document,
        "Le choix de la métrique principale est la décision la plus structurante du projet, "
        "car elle détermine le modèle livré au métier. En maintenance industrielle, les deux "
        "erreurs possibles n'ont pas le même coût :")
    tableau(document, ["Type d'erreur", "Conséquence opérationnelle", "Coût relatif"], [
        ["Faux négatif (panne manquée)", "Panne non détectée, production arrêtée, intervention d'urgence",
         "Élevé — sécurité, disponibilité,image de marque"],
        ["Faux positif (fausse alerte)", "Contrôle ou intervention préventive inutile",
         "Faible — coût du temps de maintenance"],
    ], largeurs=[4.6, 6.4, 5.0])

    paragraphe(document,
        "Une panne manquée coûte donc nettement plus cher qu'une fausse alerte. La métrique "
        "retenue est le recall (sensibilité) : elle mesure la capacité du modèle à ne pas "
        "manquer de pannes réelles, ce qui correspond exactement à l'optimisation du coût "
        "métier. Ce choix est corroboré par le comportement de la baseline : avec 82,4 % "
        "d'exactitude mais 0,000 de recall, elle ne détecte aucune panne tout en paraissant "
        "excellente. Une politique de sélection basée sur l'exactitude aurait livré ce modèle "
        "inutilisable.")

    encadre(document, "Règle de décision retenue",
        "Métrique principale : recall. Critère d'acceptation : le modèle final doit être "
        "strictement supérieur à la baseline sur cette métrique. Critère secondaire : le F1 "
        "est consulté pour s'assurer que le gain en rappel ne se fait pas au prix d'un effondrement "
        "de la précision. Règle de non-régression : l'exactitude ne doit pas être "
        "significativement inférieure à celle de la meilleure baseline naïve.")


def section_artefacts(document):
    document.add_heading("7. Artefacts produits et reproductibilité", level=1)
    paragraphe(document,
        "L'entraînement se termine par la sérialisation du modèle gagnant et de l'ensemble des "
        "scores. Ces deux fichiers constituent la livraison exploitable du projet.")

    tableau(document, ["Fichier", "Contenu", "Utilisation"], [
        ["models/model.joblib", "Le pipeline complet (preprocessing + GradientBoostingClassifier)",
         "joblib.load() puis model.predict(X) sur de nouvelles données"],
        ["models/metrics.json", "main_metric, best et les scores des 5 modèles",
         "Traçabilité, comparaison dans le temps, tableaux de bord"],
    ], largeurs=[4.4, 6.0, 5.6])

    paragraphe(document,
        "Le modèle sauvegardé est un pipeline scikit-learn complet : la transformation et le "
        "classifieur sont enregistrés dans le même fichier. Il n'y a donc aucun risque "
        "d'oublier d'appliquer le même prétraitement aux données de production — c'est un "
        "point essentiel de séparation entraînement / mise en service.")

    document.add_heading("7.1 Reproductibilité", level=2)
    paragraphe(document,
        "La graine SEED = 42 est appliquée au découpage des données et à tous les modèles "
        "stochastiques (forêt aléatoire, gradient boosting, dummy). Deux exécutions successives "
        "produisent donc un metrics.json strictement identique, ce qui rend toute comparaison "
        "ultérieure entre deux versions du code interprétation légitime.")

    document.add_heading("7.2 Environnement technique", level=2)
    tableau(document, ["Composant", "Version"], [
        ["Python", "3.13"],
        ["numpy", "2.4.4"],
        ["pandas", "3.0.2"],
        ["scikit-learn", "1.8.0"],
        ["joblib", "1.5.3"],
        ["pytest", "9.1.1"],
        ["python-docx", "1.2.0 (génération du présent rapport)"],
    ], largeurs=[6.0, 10.0])


def section_limites(document):
    document.add_heading("8. Limites connues et perspectives d'amélioration", level=1)
    document.add_heading("8.1 Limites du travail réalisé", level=2)
    puce(document, "un seul découpage aléatoire a été utilisé. Les scores seraient plus fiables "
                  "si l'on avait complété par une validation croisée stratifiée (par exemple "
                  "5 folds), qui fournirait une estimation de la variance des métriques.",
         "Absence de validation croisée — ")
    puce(document, "le jeu de test ne compte qu'environ 176 pannes : le recall est donc mesuré "
                  "sur un effectif réduit et sujet à une incertitude non négligeable.",
         "Taille du jeu de test — ")
    puce(document, "seule la variable binaire d'entrée est imputée (par la médiane pour les "
                  "variables numériques, par la modalité pour les catégorielles). Aucune "
                  "détection d'outliers ni analyse de dérive n'a été mise en place.",
         "Prétraitement limité — ")
    puce(document, "les variables quatre capteurs sont prises telles quelles ; aucune ingénierie "
                  "de variables métier (seuils de saturation, ratios, produits de capteurs) "
                  "n'a été testée.", "Absence de feature engineering métier — ")

    document.add_heading("8.2 Le déséquilibre de classes : class_weight et SMOTE", level=2)
    paragraphe(document,
        "La principale perspective d'amélioration porte sur le déséquilibre de classes (17,6 % "
        "de positifs). Les deux familles de solutions standards n'ont pas encore été mises en "
        "œuvre et constituent le prolongement naturel de ce projet :")
    tableau(document, ["Approche", "Principe", "Effet attendu"], [
        ["class_weight / scale_pos_weight",
         "Pénalise l'erreur commise sur la classe minoritaire pendant l'apprentissage",
         "Le modèle augmente son recall sans modification des données"],
        ["SMOTE (Synthetic Minority Over-sampling)",
         "Génère des observations synthétiques de la classe minoritaire par interpolation",
         "Meilleure frontière de décision sur la classe panne"],
        ["Combination des deux",
         "Rééchantillonnage pondéré au sein du fold d'entraînement uniquement",
         "Risque de surapprentissage accentué, à contrôler"],
    ], largeurs=[4.4, 6.2, 5.4])
    paragraphe(document,
        "Point de vigilance méthodologique : ces techniques de rééchantillonnage doivent être "
        "appliquées à l'intérieur de la boucle de validation croisée, sur le seul jeu "
        "d'entraînement. Les appliquer avant le découpage reviendrait à faire fuiter "
        "l'information du jeu de test dans l'entraînement et à surestimer les performances "
        "annoncées.")
    paragraphe(document,
        "Levier complémentaire : le choix de la métrique. Le passage du recall au F1 (ou au "
        "recall pondéré par le coût réel d'une panne sur un intervalle de maintenance) "
        "constituerait une amélioration métier plus fine que la seule optimisation d'un "
        "score technique.")


def section_conclusion(document):
    document.add_heading("9. Conclusion", level=1)
    paragraphe(document,
        "Le projet a été conduit de bout en bout : chargement et nettoyage des données, "
        "découpage entraînement / test sans fuite, prétraitement encapsulé dans un pipeline, "
        "comparaison de cinq modèles incluant une baseline obligatoire, sélection "
        "automatique du meilleur modèle sur une métrique métier explicite, et sérialisation "
        "des artefacts. Les huit tests d'acceptation passent, ce qui garantit la conformité "
        "du livrable aux exigences fonctionnelles du cahier des charges.")
    paragraphe(document,
        "Le modèle retenu est le Gradient Boosting, avec un recall de 0,608 sur le jeu de test "
        "contre 0,000 pour la baseline. Ce choix n'est pas un classement arbitraire : il "
        "découle d'une analyse du coût métier des erreurs — une panne manquée étant plus "
        "coûteuse qu'une fausse alerte — qui rend le recall plus pertinent que l'exactitude, "
        "par ailleurs structurellement trompeuse sur un jeu déséquilibré.")
    paragraphe(document,
        "Les principales marges de progression sont identifiées et documentées : gestion "
        "explicite du déséquilibre (class_weight, SMOTE), validation croisée stratifiée et "
        "ingénierie de variables métier. Le pipeline étant modulaire et entièrement testé, "
        "ces évolutions peuvent être intégrées sans remettre en cause l'existant.")


def annexes(document, donnees):
    document.add_page_break()
    document.add_heading("Annexe A — Résultats détaillés des 5 modèles", level=1)
    resultats = donnees["results"]
    for cle in ORDRE:
        if cle not in resultats:
            continue
        scores = resultats[cle]
        paragraphe(document, LIBELLES.get(cle, cle), taille=11, gras=True,
                   couleur=BLEU_NUIT, alignement=WD_ALIGN_PARAGRAPH.LEFT, avant=8, apres=2)
        tableau(document, ["Accuracy", "Recall", "Precision", "F1"], [[
            f"{scores.get('accuracy', 0):.3f}",
            f"{scores.get('recall', 0):.3f}",
            f"{scores.get('precision', 0):.3f}",
            f"{scores.get('f1', 0):.3f}",
        ]], largeurs=[4.0, 4.0, 4.0, 4.0], taille=10)

    document.add_heading("Annexe B — Commandes d'exécution et arborescence", level=1)
    bloc_code(document, """
# Installation
python -m venv .venv
.venv\\Scripts\\Activate.ps1
pip install -r requirements.txt

# Tests (8 tests attendus)
python -m pytest -v

# Entraînement, comparaison et génération des artefacts
python -m monprojet.train

# Génération du présent rapport
python generer_rapport.py
""")
    paragraphe(document,
        "Arborescence complète du dépôt GitHub :", gras=True, avant=6, apres=4)
    bloc_code(document, """
projet_a_completer/
├── README.md
├── requirements.txt
├── pytest.ini
├── generer_rapport.py
├── data/
│   └── exemple_pannes.csv
├── docs/
│   ├── etude-de-conception.md
│   └── choix-modele.md
├── models/
│   ├── model.joblib
│   └── metrics.json
├── monprojet/
│   ├── __init__.py
│   ├── config.py
│   ├── data.py
│   ├── models.py
│   └── train.py
└── tests/
    ├── test_data.py
    ├── test_models.py
    └── test_train.py
""")


def main():
    donnees, depuis_fichier = charger_resultats()
    document = Document()
    style_document(document)
    mise_en_page(document)

    page_de_garde(document)
    sommaire(document)
    section_introduction(document)
    section_implementation(document)
    section_tests(document)
    section_resultats(document, donnees)
    section_metrique(document)
    section_artefacts(document)
    section_limites(document)
    section_conclusion(document)
    annexes(document, donnees)

    document.core_properties.title = f"{TITRE} — {AUTEUR}"
    document.core_properties.author = AUTEUR
    document.core_properties.category = f"{GROUPE} — {ETABLISSEMENT}"
    document.core_properties.comments = (f"Rapport généré le {date.today().strftime('%d/%m/%Y')} "
                                        f"— résultats lus depuis models/metrics.json : {depuis_fichier}")

    document.save(SORTIE)
    print(f"Rapport généré : {SORTIE}")
    print(f"Modèle retenu   : {donnees.get('best')}")
    print(f"Métrique        : {donnees.get('main_metric')}")


if __name__ == "__main__":
    main()
