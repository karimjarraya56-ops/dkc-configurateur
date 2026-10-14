"""Textes et classement des produits S5 (source : catalogue général Sati Tunisia, éd. 6, p. 12-57).

Modifiable sans toucher à app.py : les clés sont les noms de produits de la base Excel.
"""

# Courte description affichée sur la carte de chaque produit
DESCRIPTIONS = {
    # Couvrir
    "Couvercle pour élément rectiligne": "Se pose par encliquetage, sans vis ni clips : il est autoporteur grâce à sa "
                                         "bordure continue. Avec couvercle, le chemin passe en IP 20 (fond foré) ou IP 40 (fond plein).",
    "Plaque de blocage du couvercle": "Maintient le couvercle verrouillé sur l'élément rectiligne.",
    "Kit pièces détachées (couvercle à charnières)": "Pour l'entretien des couvercles à charnières : 10 charnières, "
                                                     "50 rivets et 5 tresses de terre par emballage.",
    # Tourner
    "Courbe plate à 45°": "Change la direction du parcours de 45° dans le même plan. Rayon intérieur 120 mm.",
    "Courbe plate à 90°": "Change la direction du parcours de 90° dans le même plan. Rayon intérieur 120 mm.",
    # Monter / descendre
    "Courbe en montée à 45°": "Fait monter le parcours de 45° (changement de niveau vers le haut). Rayon 120 mm.",
    "Courbe en montée à 90°": "Fait passer le parcours à la verticale, vers le haut. Rayon 120 mm.",
    "Courbe en descente à 45°": "Fait descendre le parcours de 45° (changement de niveau vers le bas). Rayon 120 mm.",
    "Courbe en descente à 90°": "Fait passer le parcours à la verticale, vers le bas. Rayon 120 mm.",
    "Courbe en montée à 90° ligne oblique droite": "Montée à 90° avec sortie orientée en oblique vers la droite.",
    "Courbe en montée à 90° ligne oblique gauche": "Montée à 90° avec sortie orientée en oblique vers la gauche.",
    "Courbe en descente à 90° ligne oblique droite": "Descente à 90° avec sortie orientée en oblique vers la droite.",
    "Courbe en descente à 90° ligne oblique gauche": "Descente à 90° avec sortie orientée en oblique vers la gauche.",
    "Courbe en descente à 90° avec inversion de plan": "Descente à 90° avec changement de plan du chemin de câbles.",
    # Dériver
    "Dérivation plate en «T»": "Crée un embranchement en T à 90° dans le même plan. Rayon 120 mm.",
    "Dérivation latérale": "Ajoute une sortie courbe sur un côté du parcours principal.",
    "Dérivation plate à croisement": "Croisement à 4 voies dans le même plan.",
    "«T» en montée": "Crée un départ vertical vers le haut depuis le parcours horizontal.",
    "«T» en descente": "Crée un départ vertical vers le bas depuis le parcours horizontal.",
    "«T» en montée oblique": "Départ vers le haut avec sortie orientée en oblique.",
    "«T» en descente oblique": "Départ vers le bas avec sortie orientée en oblique.",
    "«T» en montée du couvercle": "Départ vertical vers le haut réalisé depuis le couvercle du chemin.",
    "«T» en montée oblique du couvercle": "Départ oblique vers le haut réalisé depuis le couvercle du chemin.",
    # Réduire
    "Raccord de réduction centrale": "Réduit la largeur du chemin de B à B1, de façon symétrique. "
                                     "D'autres réductions sont disponibles en option.",
    "Raccord de réduction droite": "Réduit la largeur de B à B1 avec un bord qui reste droit (réduction à droite).",
    "Raccord de réduction gauche": "Réduit la largeur de B à B1 avec un bord qui reste droit (réduction à gauche).",
    "Élément de réduction": "Crée une réduction à droite ou à gauche en pliant simplement les ailettes. "
                            "À utiliser avec le joint linéaire GTO.",
    # Raccorder
    "Joint linéaire": "Relie deux extrémités « mâles » (accessoires entre eux, ou élément coupé). "
                      "Deux joints par jonction.",
    "Joint articulé vertical": "Réalise des variations de hauteur et des courbes en montée ou en descente "
                               "sur mesure. Deux joints par jonction.",
    "Joint articulé horizontal": "Permet de dévier le parcours selon n'importe quel angle. Deux joints par jonction.",
    "Couvre-joint pour base": "Protège les câbles des arêtes vives aux jonctions coupées sur chantier, garantit l'IP 40 / "
                              "IP 44 des bases et la continuité électrique.",
    "Couvre-joint pour couvercle": "Rétablit l'IP 40 des couvercles coupés sur chantier ou mal aboutés, et assure "
                                   "la continuité électrique.",
    "Bride de raccord": "Raccorde le chemin à une autre canalisation ou à un tableau électrique, ou le fixe "
                        "au mur, au plafond ou au sol.",
    "Tête de fermeture": "Ferme proprement l'extrémité d'un parcours.",
    "Joint universel (sur demande, selon marque du canal)": "Raccorde le S5 à des goulottes d'autres marques. "
                                                            "Code sur demande selon la marque à raccorder.",
    # Installer et protéger
    "Séparateur": "Cloison qui divise le chemin en deux compartiments (par ex. courants forts / courants faibles).",
    "Étrier d'espacement": "Fixe le chemin au mur ou au sol en le gardant à 50 mm de la paroi.",
    "Plaque de fixation du câble": "Pour fixer les câbles sur les tronçons verticaux avec des colliers jusqu'à 9 mm. "
                                   "En barrettes de 5.",
    "Kit de protection IP 44 – bande IP 44": "Bande à clipser à chaque jonction. Avec le joint GA, elle fait passer une "
                                             "goulotte IP 40 (fond plein + couvercle) en IP 44.",
    "Joint auto-adhésif (kit IP 44)": "Joint mousse ignifuge (rouleau de 10 m) pour les tronçons verticaux. "
                                      "Avec les bandes IP 44, assure l'IP 44.",
    "Platine de terre pour connexions équipotentielles": "Plaquette en laiton nickelé qui garantit la continuité "
                                                         "électrique aux jonctions « mâle-mâle ».",
    # Visserie et retouche
    "Vis M6 x 10 – jonction mécanique (cadre sous-tête)": "Pour l'assemblage mécanique des éléments.",
    "Vis M6 x 08 – jonction électrique des bases": "Assure la continuité électrique entre les bases.",
    "Vis M5 x 05 – jonction électrique des couvercles": "Assure la continuité électrique entre les couvercles.",
    "Vis et écrou M6 x 15 galvanisés à chaud": "Vis + écrou anti-desserrage pour l'assemblage des éléments galvanisés à chaud.",
    "Écrou M6 anti-desserrage": "Écrou à embase crantée, ne se desserre pas avec les vibrations.",
    "Écrou M8 anti-desserrage": "Écrou à embase crantée, ne se desserre pas avec les vibrations.",
    "Écrou M5": "Pour les vis M5 des couvercles.",
    "Rondelle plate pour vis M6": "Rondelle plate pour vis M6.",
    "Peinture pour retouches (0,40 L)": "Bombe de peinture pour retoucher les éléments vernis après une coupe.",
    "Peinture au zinc – spray (0,40 L)": "Spray zinc pour les retouches à froid des éléments galvanisés à chaud.",
}

# Description du chemin de câbles selon ses options
DESCRIPTION_CHEMIN = {
    "Foré (perforé)": "Fond perforé (boutonnières 7 x 25 mm) qui ventile les câbles. IP 20 avec couvercle.",
    "Non foré (plein)": "Fond plein qui protège les câbles. IP 40 avec couvercle, IP 44 avec le kit IP 44.",
    "charnières": "Couvercle monté sur 2 charnières, dont une avec tresse en cuivre pour la continuité équipotentielle.",
    "1,5 mm": "Tôle 1,5 mm pour les fortes charges, sans emboîtement : prévoir les jonctions.",
}

# Accessoires regroupés par usage (onglets de l'étape « accessoires compatibles »)
RUBRIQUES = [
    ("Couvrir", ["Couvercle pour élément rectiligne", "Plaque de blocage du couvercle",
                 "Kit pièces détachées (couvercle à charnières)"]),
    ("Tourner", ["Courbe plate à 90°", "Courbe plate à 45°"]),
    ("Monter / descendre", ["Courbe en montée à 90°", "Courbe en montée à 45°", "Courbe en descente à 90°",
                            "Courbe en descente à 45°", "Courbe en montée à 90° ligne oblique droite",
                            "Courbe en montée à 90° ligne oblique gauche", "Courbe en descente à 90° ligne oblique droite",
                            "Courbe en descente à 90° ligne oblique gauche",
                            "Courbe en descente à 90° avec inversion de plan"]),
    ("Dériver", ["Dérivation plate en «T»", "Dérivation latérale", "Dérivation plate à croisement", "«T» en montée",
                 "«T» en descente", "«T» en montée oblique", "«T» en descente oblique", "«T» en montée du couvercle",
                 "«T» en montée oblique du couvercle"]),
    ("Réduire", ["Raccord de réduction centrale", "Raccord de réduction droite", "Raccord de réduction gauche",
                 "Élément de réduction"]),
    ("Raccorder", ["Joint linéaire", "Joint articulé vertical", "Joint articulé horizontal", "Couvre-joint pour base",
                   "Couvre-joint pour couvercle", "Bride de raccord", "Tête de fermeture",
                   "Joint universel (sur demande, selon marque du canal)"]),
    ("Installer & protéger", ["Séparateur", "Étrier d'espacement", "Plaque de fixation du câble",
                              "Kit de protection IP 44 – bande IP 44", "Joint auto-adhésif (kit IP 44)",
                              "Platine de terre pour connexions équipotentielles"]),
    ("Visserie & retouche", ["Vis M6 x 10 – jonction mécanique (cadre sous-tête)",
                             "Vis M6 x 08 – jonction électrique des bases",
                             "Vis M5 x 05 – jonction électrique des couvercles",
                             "Vis et écrou M6 x 15 galvanisés à chaud", "Écrou M6 anti-desserrage",
                             "Écrou M8 anti-desserrage", "Écrou M5", "Rondelle plate pour vis M6",
                             "Peinture pour retouches (0,40 L)", "Peinture au zinc – spray (0,40 L)"]),
]

# Choix du traitement de surface selon l'environnement (catalogue p. 8)
CONSEIL_FINITION = {
    "colonnes": ["Sendzimir (ZS)", "Vernis (VB / VG)", "Galva à chaud (ZF)", "Inox (IX)"],
    "lignes": [
        ("Intérieur", ["Excellent", "Excellent", "Excellent", "Excellent"]),
        ("Extérieur normal", ["Possible", "Bon", "Excellent", "Excellent"]),
        ("Bord de mer", ["Déconseillé", "Possible", "Bon", "Excellent"]),
        ("Industrie alimentaire", ["Déconseillé", "Bon", "Possible", "Excellent"]),
        ("Milieu acide", ["Déconseillé", "Bon", "Déconseillé", "Excellent"]),
        ("Milieu alcalin", ["Déconseillé", "Bon", "Possible", "Excellent"]),
        ("Milieu halogène", ["Déconseillé", "Bon", "Déconseillé", "Excellent"]),
    ],
}


# ---------------------------------------------------------------- Traductions
DESCRIPTIONS_EN = {
    "Couvercle pour élément rectiligne": "Snaps onto the straight element without screws or clips: self-supporting thanks to "
                                         "its continuous border. With a cover the tray is IP 20 (perforated) or IP 40 (solid bottom).",
    "Plaque de blocage du couvercle": "Keeps the cover locked on the straight element.",
    "Kit pièces détachées (couvercle à charnières)": "Spare parts for hinged covers: 10 hinges, 50 rivets and 5 bonding straps per pack.",
    "Courbe plate à 45°": "Changes the route direction by 45° in the same plane. Inner radius 120 mm.",
    "Courbe plate à 90°": "Changes the route direction by 90° in the same plane. Inner radius 120 mm.",
    "Courbe en montée à 45°": "Raises the route by 45° (level change upwards). Radius 120 mm.",
    "Courbe en montée à 90°": "Turns the route vertical, upwards. Radius 120 mm.",
    "Courbe en descente à 45°": "Lowers the route by 45° (level change downwards). Radius 120 mm.",
    "Courbe en descente à 90°": "Turns the route vertical, downwards. Radius 120 mm.",
    "Courbe en montée à 90° ligne oblique droite": "90° rising elbow with the outlet skewed to the right.",
    "Courbe en montée à 90° ligne oblique gauche": "90° rising elbow with the outlet skewed to the left.",
    "Courbe en descente à 90° ligne oblique droite": "90° falling elbow with the outlet skewed to the right.",
    "Courbe en descente à 90° ligne oblique gauche": "90° falling elbow with the outlet skewed to the left.",
    "Courbe en descente à 90° avec inversion de plan": "90° falling elbow with a change of tray plane.",
    "Dérivation plate en «T»": "Creates a 90° T branch in the same plane. Radius 120 mm.",
    "Dérivation latérale": "Adds a curved outlet on one side of the main route.",
    "Dérivation plate à croisement": "4-way crossing in the same plane.",
    "«T» en montée": "Creates a vertical branch upwards from the horizontal route.",
    "«T» en descente": "Creates a vertical branch downwards from the horizontal route.",
    "«T» en montée oblique": "Upward branch with a skewed outlet.",
    "«T» en descente oblique": "Downward branch with a skewed outlet.",
    "«T» en montée du couvercle": "Vertical upward branch taken from the tray cover.",
    "«T» en montée oblique du couvercle": "Skewed upward branch taken from the tray cover.",
    "Raccord de réduction centrale": "Reduces the tray width from B to B1 symmetrically. Other reductions available on request.",
    "Raccord de réduction droite": "Reduces the width from B to B1 with one straight edge (right-hand reduction).",
    "Raccord de réduction gauche": "Reduces the width from B to B1 with one straight edge (left-hand reduction).",
    "Élément de réduction": "Creates right or left reductions by simply bending the flaps. For use with the GTO joint plate.",
    "Joint linéaire": "Joins two “male” ends (fittings together, or a cut element). Two plates per junction.",
    "Joint articulé vertical": "Makes height differences and custom rising or falling elbows. Two joints per junction.",
    "Joint articulé horizontal": "Allows route deviations at any angle. Two joints per junction.",
    "Couvre-joint pour base": "Protects cables from sharp edges at site-cut joints, provides IP 40 / IP 44 for the bases "
                              "and electrical continuity.",
    "Couvre-joint pour couvercle": "Restores IP 40 on site-cut or poorly matched covers and ensures electrical continuity.",
    "Bride de raccord": "Connects the tray to another trunking or a distribution board, or fixes it to wall, ceiling or floor.",
    "Tête de fermeture": "Neatly closes the end of a route.",
    "Joint universel (sur demande, selon marque du canal)": "Joins the S5 to other brands' trays. Code on request depending on the brand.",
    "Séparateur": "Divider splitting the tray into two compartments (e.g. power / low-voltage cables).",
    "Étrier d'espacement": "Fixes the tray to the wall or floor, 50 mm away from the surface.",
    "Plaque de fixation du câble": "To fix cables on vertical runs with cable ties up to 9 mm. In bars of 5.",
    "Kit de protection IP 44 – bande IP 44": "Insert clipped at each joint. With the GA seal it upgrades an IP 40 trunking "
                                             "(solid bottom + cover) to IP 44.",
    "Joint auto-adhésif (kit IP 44)": "Self-extinguishing foam seal (10 m roll) for vertical runs. With IP 44 inserts it gives IP 44.",
    "Platine de terre pour connexions équipotentielles": "Nickel-plated plate ensuring electrical continuity at “male-male” joints.",
    "Vis M6 x 10 – jonction mécanique (cadre sous-tête)": "For the mechanical connection of elements.",
    "Vis M6 x 08 – jonction électrique des bases": "Ensures electrical continuity between bases.",
    "Vis M5 x 05 – jonction électrique des couvercles": "Ensures electrical continuity between covers.",
    "Vis et écrou M6 x 15 galvanisés à chaud": "Screw + anti-loosening nut for hot-dip galvanized elements.",
    "Écrou M6 anti-desserrage": "Serrated flange nut that does not loosen with vibration.",
    "Écrou M8 anti-desserrage": "Serrated flange nut that does not loosen with vibration.",
    "Écrou M5": "For the M5 cover screws.",
    "Rondelle plate pour vis M6": "Flat washer for M6 screws.",
    "Peinture pour retouches (0,40 L)": "Spray paint to touch up painted elements after cutting.",
    "Peinture au zinc – spray (0,40 L)": "Zinc spray for cold touch-up of hot-dip galvanized elements.",
}

DESCRIPTIONS_IT = {
    "Couvercle pour élément rectiligne": "Si monta a scatto, senza viti né clip: è autoportante grazie al bordo continuo. "
                                         "Con il coperchio il canale è IP 20 (forato) o IP 40 (fondo pieno).",
    "Plaque de blocage du couvercle": "Mantiene il coperchio bloccato sull'elemento rettilineo.",
    "Kit pièces détachées (couvercle à charnières)": "Ricambi per coperchi incernierati: 10 cerniere, 50 rivetti e 5 trecce di terra per confezione.",
    "Courbe plate à 45°": "Cambia la direzione del percorso di 45° sullo stesso piano. Raggio interno 120 mm.",
    "Courbe plate à 90°": "Cambia la direzione del percorso di 90° sullo stesso piano. Raggio interno 120 mm.",
    "Courbe en montée à 45°": "Fa salire il percorso di 45° (cambio di quota verso l'alto). Raggio 120 mm.",
    "Courbe en montée à 90°": "Porta il percorso in verticale, verso l'alto. Raggio 120 mm.",
    "Courbe en descente à 45°": "Fa scendere il percorso di 45° (cambio di quota verso il basso). Raggio 120 mm.",
    "Courbe en descente à 90°": "Porta il percorso in verticale, verso il basso. Raggio 120 mm.",
    "Courbe en montée à 90° ligne oblique droite": "Salita a 90° con uscita orientata obliquamente a destra.",
    "Courbe en montée à 90° ligne oblique gauche": "Salita a 90° con uscita orientata obliquamente a sinistra.",
    "Courbe en descente à 90° ligne oblique droite": "Discesa a 90° con uscita orientata obliquamente a destra.",
    "Courbe en descente à 90° ligne oblique gauche": "Discesa a 90° con uscita orientata obliquamente a sinistra.",
    "Courbe en descente à 90° avec inversion de plan": "Discesa a 90° con cambio di piano del canale.",
    "Dérivation plate en «T»": "Crea una derivazione a T a 90° sullo stesso piano. Raggio 120 mm.",
    "Dérivation latérale": "Aggiunge un'uscita curva su un lato del percorso principale.",
    "Dérivation plate à croisement": "Incrocio a 4 vie sullo stesso piano.",
    "«T» en montée": "Crea una partenza verticale verso l'alto dal percorso orizzontale.",
    "«T» en descente": "Crea una partenza verticale verso il basso dal percorso orizzontale.",
    "«T» en montée oblique": "Partenza verso l'alto con uscita obliqua.",
    "«T» en descente oblique": "Partenza verso il basso con uscita obliqua.",
    "«T» en montée du couvercle": "Partenza verticale verso l'alto realizzata dal coperchio del canale.",
    "«T» en montée oblique du couvercle": "Partenza obliqua verso l'alto realizzata dal coperchio del canale.",
    "Raccord de réduction centrale": "Riduce la larghezza del canale da B a B1 in modo simmetrico. Altre riduzioni su richiesta.",
    "Raccord de réduction droite": "Riduce la larghezza da B a B1 con un bordo dritto (riduzione a destra).",
    "Raccord de réduction gauche": "Riduce la larghezza da B a B1 con un bordo dritto (riduzione a sinistra).",
    "Élément de réduction": "Crea riduzioni a destra o a sinistra piegando semplicemente le alette. Da usare con il giunto lineare GTO.",
    "Joint linéaire": "Collega due estremità «maschio» (accessori tra loro o elemento tagliato). Due giunti per giunzione.",
    "Joint articulé vertical": "Realizza variazioni di quota e curve in salita o in discesa su misura. Due giunti per giunzione.",
    "Joint articulé horizontal": "Permette deviazioni del percorso con qualsiasi angolo. Due giunti per giunzione.",
    "Couvre-joint pour base": "Protegge i cavi dagli spigoli vivi nelle giunzioni tagliate in cantiere, garantisce IP 40 / IP 44 "
                              "dei fondi e la continuità elettrica.",
    "Couvre-joint pour couvercle": "Ripristina l'IP 40 dei coperchi tagliati in cantiere o mal accostati e assicura la continuità elettrica.",
    "Bride de raccord": "Collega il canale a un'altra canalizzazione o a un quadro elettrico, o lo fissa a parete, soffitto o pavimento.",
    "Tête de fermeture": "Chiude in modo pulito l'estremità di un percorso.",
    "Joint universel (sur demande, selon marque du canal)": "Collega l'S5 a canali di altre marche. Codice su richiesta secondo la marca.",
    "Séparateur": "Setto che divide il canale in due scomparti (es. energia / segnale).",
    "Étrier d'espacement": "Fissa il canale a parete o a pavimento, distanziato di 50 mm.",
    "Plaque de fixation du câble": "Per fissare i cavi nei tratti verticali con fascette fino a 9 mm. In stecche da 5.",
    "Kit de protection IP 44 – bande IP 44": "Banda da agganciare a ogni giunzione. Con la guarnizione GA porta un canale IP 40 "
                                             "(fondo pieno + coperchio) a IP 44.",
    "Joint auto-adhésif (kit IP 44)": "Guarnizione in espanso autoestinguente (rotolo da 10 m) per i tratti verticali. "
                                      "Con le bande IP 44 garantisce l'IP 44.",
    "Platine de terre pour connexions équipotentielles": "Piastrina in ottone nichelato che garantisce la continuità elettrica "
                                                         "nelle giunzioni «maschio-maschio».",
    "Vis M6 x 10 – jonction mécanique (cadre sous-tête)": "Per la giunzione meccanica degli elementi.",
    "Vis M6 x 08 – jonction électrique des bases": "Assicura la continuità elettrica tra i fondi.",
    "Vis M5 x 05 – jonction électrique des couvercles": "Assicura la continuità elettrica tra i coperchi.",
    "Vis et écrou M6 x 15 galvanisés à chaud": "Vite + dado antisvitamento per elementi zincati a caldo.",
    "Écrou M6 anti-desserrage": "Dado flangiato zigrinato, non si allenta con le vibrazioni.",
    "Écrou M8 anti-desserrage": "Dado flangiato zigrinato, non si allenta con le vibrazioni.",
    "Écrou M5": "Per le viti M5 dei coperchi.",
    "Rondelle plate pour vis M6": "Rondella piana per viti M6.",
    "Peinture pour retouches (0,40 L)": "Bomboletta di vernice per ritoccare gli elementi verniciati dopo il taglio.",
    "Peinture au zinc – spray (0,40 L)": "Spray allo zinco per ritocchi a freddo degli elementi zincati a caldo.",
}

DESCRIPTIONS_PAR_LANGUE = {"fr": DESCRIPTIONS, "en": DESCRIPTIONS_EN, "it": DESCRIPTIONS_IT}

DESCRIPTION_CHEMIN_PAR_LANGUE = {
    "fr": DESCRIPTION_CHEMIN,
    "en": {
        "Foré (perforé)": "Perforated bottom (7 x 25 mm slots) that ventilates the cables. IP 20 with cover.",
        "Non foré (plein)": "Solid bottom that protects the cables. IP 40 with cover, IP 44 with the IP 44 kit.",
        "charnières": "Cover mounted on 2 hinges, one with a copper braid for equipotential continuity.",
        "1,5 mm": "1.5 mm sheet for heavy loads, without interlocking: order the joints separately.",
    },
    "it": {
        "Foré (perforé)": "Fondo forato (asole 7 x 25 mm) che ventila i cavi. IP 20 con coperchio.",
        "Non foré (plein)": "Fondo pieno che protegge i cavi. IP 40 con coperchio, IP 44 con il kit IP 44.",
        "charnières": "Coperchio montato su 2 cerniere, una con treccia di rame per la continuità equipotenziale.",
        "1,5 mm": "Lamiera 1,5 mm per carichi elevati, senza innesto: prevedere le giunzioni.",
    },
}

CONSEIL_FINITION_TR = {
    "colonnes": {
        "fr": CONSEIL_FINITION["colonnes"],
        "en": ["Sendzimir (ZS)", "Painted (VB / VG)", "Hot-dip galv. (ZF)", "Stainless (IX)"],
        "it": ["Sendzimir (ZS)", "Verniciato (VB / VG)", "Zincato a caldo (ZF)", "Inox (IX)"],
    },
    "environnements": {
        "Intérieur": {"en": "Indoor", "it": "Interno"},
        "Extérieur normal": {"en": "Normal outdoor", "it": "Esterno normale"},
        "Bord de mer": {"en": "Seaside", "it": "Ambiente marino"},
        "Industrie alimentaire": {"en": "Food industry", "it": "Industria alimentare"},
        "Milieu acide": {"en": "Acid environment", "it": "Ambiente acido"},
        "Milieu alcalin": {"en": "Alkaline environment", "it": "Ambiente alcalino"},
        "Milieu halogène": {"en": "Halogen environment", "it": "Ambiente alogeno"},
    },
    "valeurs": {
        "Excellent": {"en": "Very good", "it": "Ottimo"},
        "Bon": {"en": "Good", "it": "Buono"},
        "Possible": {"en": "Possible", "it": "Possibile"},
        "Déconseillé": {"en": "Not recommended", "it": "Sconsigliato"},
    },
}
