"""Traductions de l'interface (français par défaut, italien, anglais).

Chaque texte : {"fr": ..., "it": ..., "en": ...}. Les noms anglais des produits viennent de la
colonne Product_EN de la base Excel ; les noms italiens sont dans PRODUITS_IT ci-dessous.
"""

LANGUES = {"fr": "FR", "it": "IT", "en": "EN"}

TEXTES = {
    # En-tête
    "page_titre": {"fr": "Configurateur DKC – S5", "it": "Configuratore DKC – S5", "en": "DKC Configurator – S5"},
    "tag": {"fr": "Ligne <b>S5</b> · S2 et S3 bientôt", "it": "Linea <b>S5</b> · S2 e S3 in arrivo",
            "en": "<b>S5</b> line · S2 and S3 coming soon"},
    "titre": {"fr": "Configurateur de chemins de câbles", "it": "Configuratore di canali portacavi",
              "en": "Cable tray configurator"},
    "sous_titre": {"fr": "Choisissez votre chemin de câbles : tous les accessoires compatibles s'affichent avec leur référence exacte.",
                   "it": "Scegliete il vostro canale portacavi: tutti gli accessori compatibili compaiono con il loro codice esatto.",
                   "en": "Choose your cable tray: every compatible fitting is shown with its exact part number."},
    "stat_produits": {"fr": "produits", "it": "prodotti", "en": "products"},
    "stat_refs": {"fr": "références", "it": "codici", "en": "part numbers"},
    "stat_finitions": {"fr": "finitions", "it": "finiture", "en": "finishes"},
    "mode_config": {"fr": "Configurer mon chemin de câbles", "it": "Configurare il mio canale",
                    "en": "Configure my cable tray"},
    "mode_recherche": {"fr": "Recherche directe d'un article", "it": "Ricerca diretta di un articolo",
                       "en": "Direct item search"},
    # Configurateur
    "ligne": {"fr": "Ligne de produit", "it": "Linea di prodotto", "en": "Product line"},
    "bientot": {"fr": "bientôt", "it": "in arrivo", "en": "soon"},
    "info_s2_s3": {"fr": "Les lignes S2 (passerelles à fil) et S3 (échelles à câbles) arrivent bientôt. "
                         "Seule la ligne S5 est disponible pour l'instant.",
                   "it": "Le linee S2 (passerelle a filo) e S3 (scale portacavi) arriveranno presto. "
                         "Per ora è disponibile solo la linea S5.",
                   "en": "The S2 (mesh wire trays) and S3 (cable ladders) lines are coming soon. "
                         "Only the S5 line is available for now."},
    "votre_chemin": {"fr": "Votre chemin de câbles", "it": "Il vostro canale portacavi", "en": "Your cable tray"},
    "info_chemin": {"fr": "dimensions, fond et finition", "it": "dimensioni, fondo e finitura",
                    "en": "size, bottom and finish"},
    "H_mm": {"fr": "Hauteur H (mm)", "it": "Altezza H (mm)", "en": "Height H (mm)"},
    "B_mm": {"fr": "Base B (mm)", "it": "Base B (mm)", "en": "Base B (mm)"},
    "B1_mm": {"fr": "Base de sortie B1 (mm)", "it": "Base di uscita B1 (mm)", "en": "Outlet base B1 (mm)"},
    "L_mm": {"fr": "Longueur L (mm)", "it": "Lunghezza L (mm)", "en": "Length L (mm)"},
    "Perforé": {"fr": "Fond", "it": "Fondo", "en": "Bottom"},
    "Finition": {"fr": "Matière / finition", "it": "Materiale / finitura", "en": "Material / finish"},
    "Couvercle_charnières": {"fr": "Couvercle à charnières", "it": "Coperchio incernierato", "en": "Hinged cover"},
    "Épaisseur": {"fr": "Épaisseur de tôle", "it": "Spessore lamiera", "en": "Sheet thickness"},
    "seule_option": {"fr": "Seule option", "it": "Unica opzione", "en": "Only option"},
    "options": {"fr": "options", "it": "opzioni", "en": "options"},
    "quelle_finition": {"fr": "Quelle finition choisir selon l'environnement ?",
                        "it": "Quale finitura scegliere in base all'ambiente?",
                        "en": "Which finish for which environment?"},
    "environnement": {"fr": "Environnement", "it": "Ambiente", "en": "Environment"},
    "avertissement_finition": {"fr": "Valeurs indicatives : la température, la concentration chimique, etc. peuvent influer sur la résistance.",
                               "it": "Valori indicativi: temperatura, concentrazione chimica, ecc. possono influire sulla resistenza.",
                               "en": "Indicative values: temperature, chemical concentration, etc. can affect resistance."},
    # Panneau du chemin
    "chemins_possibles": {"fr": "Chemins possibles", "it": "Canali possibili", "en": "Possible trays"},
    "faites_choix": {"fr": "Faites vos choix à gauche pour obtenir la référence exacte.",
                     "it": "Fate le vostre scelte a sinistra per ottenere il codice esatto.",
                     "en": "Make your choices on the left to get the exact part number."},
    "dimensions": {"fr": "Dimensions (mm)", "it": "Dimensioni (mm)", "en": "Dimensions (mm)"},
    "finition": {"fr": "Finition", "it": "Finitura", "en": "Finish"},
    "tua": {"fr": "Section utile (T.U.A.)", "it": "Sezione utile (T.U.A.)", "en": "Usable area (T.U.A.)"},
    "conditionnement": {"fr": "Conditionnement", "it": "Confezione", "en": "Packing"},
    "couvercle_assorti": {"fr": "Couvercle assorti", "it": "Coperchio abbinato", "en": "Matching cover"},
    "page_catalogue": {"fr": "Page catalogue", "it": "Pagina catalogo", "en": "Catalogue page"},
    "chemin_fore": {"fr": "Chemin foré", "it": "Canale forato", "en": "Perforated tray"},
    "chemin_non_fore": {"fr": "Chemin non foré", "it": "Canale non forato", "en": "Solid bottom tray"},
    "a_charnieres": {"fr": " à couvercle à charnières", "it": " con coperchio incernierato", "en": " with hinged cover"},
    # Accessoires
    "accessoires": {"fr": "Accessoires compatibles", "it": "Accessori compatibili", "en": "Compatible fittings"},
    "chemin_choisi": {"fr": "Chemin choisi", "it": "Canale scelto", "en": "Selected tray"},
    "bandeau": {"fr": "Tous les accessoires ci-dessous sont aux dimensions et dans la finition de votre chemin. "
                      "Ils se raccordent par encastrement sur l'extrémité « femelle » du chemin, ou avec le joint "
                      "linéaire GTO sur une extrémité « mâle ».",
                "it": "Tutti gli accessori qui sotto hanno le dimensioni e la finitura del vostro canale. "
                      "Si collegano a innesto sull'estremità «femmina» del canale, o con il giunto lineare GTO "
                      "su un'estremità «maschio».",
                "en": "All fittings below match the size and finish of your tray. They slide into the tray's "
                      "“female” end, or join a “male” end with the GTO joint plate."},
    "aucun_accessoire": {"fr": "Aucun accessoire de ce type pour ce chemin de câbles.",
                         "it": "Nessun accessorio di questo tipo per questo canale.",
                         "en": "No fitting of this type for this cable tray."},
    "non_dispo": {"fr": "Non disponibles pour ce chemin :", "it": "Non disponibili per questo canale:",
                  "en": "Not available for this tray:"},
    "existe_pas": {"fr": "n'existe pas en", "it": "non esiste in", "en": "not made in"},
    "sortie_b1": {"fr": "Sortie B1", "it": "Uscita B1", "en": "Outlet B1"},
    "plus_couvercle": {"fr": "+ Couvercle assorti :", "it": "+ Coperchio abbinato:", "en": "+ Matching cover:"},
    # Recherche directe
    "type_produit": {"fr": "Type de produit", "it": "Tipo di prodotto", "en": "Product type"},
    "produit": {"fr": "Produit", "it": "Prodotto", "en": "Product"},
    "produit_1": {"fr": "produit", "it": "prodotto", "en": "product"},
    "produit_n": {"fr": "produits", "it": "prodotti", "en": "products"},
    "rechercher": {"fr": "🔎  Rechercher un produit (ex. courbe, 90, couvercle…)",
                   "it": "🔎  Cercare un prodotto (es. curva, 90, coperchio…)",
                   "en": "🔎  Search a product (e.g. elbow, 90, cover…)"},
    "ref_abrege": {"fr": "réf.", "it": "cod.", "en": "ref."},
    "aucun_produit": {"fr": "Aucun produit ne correspond à cette recherche.",
                      "it": "Nessun prodotto corrisponde a questa ricerca.",
                      "en": "No product matches this search."},
    "refs_possibles": {"fr": "Références possibles", "it": "Codici possibili", "en": "Possible part numbers"},
    "precisez": {"fr": "Précisez vos choix pour obtenir une référence unique.",
                 "it": "Precisate le vostre scelte per ottenere un codice unico.",
                 "en": "Refine your choices to get a single part number."},
    "poids": {"fr": "Poids", "it": "Peso", "en": "Weight"},
    "kg_piece": {"fr": "kg/pièce", "it": "kg/pz", "en": "kg/pc"},
    "reference_dkc": {"fr": "Référence DKC", "it": "Codice DKC", "en": "DKC part number"},
    "a_commander": {"fr": "À commander avec", "it": "Da ordinare con", "en": "Order together with"},
    "couvercle_optionnel": {"fr": "Couvercle (optionnel)", "it": "Coperchio (opzionale)", "en": "Cover (optional)"},
    "joint_gto": {"fr": "Joint linéaire GTO", "it": "Giunto lineare GTO", "en": "GTO joint plate"},
    "avec_charnieres": {"fr": "Avec charnières", "it": "Con cerniere", "en": "With hinges"},
    "sans_charnieres": {"fr": "Sans charnières", "it": "Senza cerniere", "en": "Without hinges"},
    # Liste de matériel
    "ajouter": {"fr": "＋ Ajouter à ma liste", "it": "＋ Aggiungi alla lista", "en": "＋ Add to my list"},
    "ajouter_couvercle": {"fr": "＋ Couvercle", "it": "＋ Coperchio", "en": "＋ Cover"},
    "ajouter_chemin": {"fr": "＋ Ajouter ce chemin à ma liste", "it": "＋ Aggiungi questo canale alla lista",
                       "en": "＋ Add this tray to my list"},
    "ajouter_couvercle_chemin": {"fr": "＋ Couvercle assorti", "it": "＋ Coperchio abbinato", "en": "＋ Matching cover"},
    "dans_liste": {"fr": "dans ma liste", "it": "nella lista", "en": "in my list"},
    "ma_liste": {"fr": "Ma liste de matériel", "it": "La mia distinta materiali", "en": "My bill of materials"},
    "liste_vide": {"fr": "Votre liste est vide. Ajoutez le chemin de câbles et les accessoires avec les boutons « Ajouter à ma liste ».",
                   "it": "La lista è vuota. Aggiungete il canale e gli accessori con i pulsanti «Aggiungi alla lista».",
                   "en": "Your list is empty. Add the cable tray and fittings with the “Add to my list” buttons."},
    "info_liste": {"fr": "Modifiez les quantités directement dans le tableau (0 = retirer l'article).",
                   "it": "Modificate le quantità direttamente nella tabella (0 = togliere l'articolo).",
                   "en": "Edit quantities directly in the table (0 = remove the item)."},
    "telecharger_excel": {"fr": "Télécharger en Excel", "it": "Scarica in Excel", "en": "Download as Excel"},
    "telecharger_pdf": {"fr": "Télécharger en PDF", "it": "Scarica in PDF", "en": "Download as PDF"},
    "vider": {"fr": "Vider la liste", "it": "Svuota la lista", "en": "Clear the list"},
    "article_1": {"fr": "article", "it": "articolo", "en": "item"},
    "article_n": {"fr": "articles", "it": "articoli", "en": "items"},
    "col_photo": {"fr": "Photo", "it": "Foto", "en": "Photo"},
    "col_code": {"fr": "Code", "it": "Codice", "en": "Code"},
    "col_designation": {"fr": "Désignation", "it": "Descrizione", "en": "Description"},
    "col_finition": {"fr": "Finition", "it": "Finitura", "en": "Finish"},
    "col_dimensions": {"fr": "Dimensions (mm)", "it": "Dimensioni (mm)", "en": "Dimensions (mm)"},
    "col_quantite": {"fr": "Quantité", "it": "Quantità", "en": "Quantity"},
    "col_unite": {"fr": "Unité", "it": "Unità", "en": "Unit"},
    "col_page": {"fr": "Page", "it": "Pag.", "en": "Page"},
    "titre_doc": {"fr": "Liste de matériel – DKC ligne S5", "it": "Distinta materiali – DKC linea S5",
                  "en": "Bill of materials – DKC S5 line"},
    "date": {"fr": "Date", "it": "Data", "en": "Date"},
    "total_articles": {"fr": "Articles", "it": "Articoli", "en": "Items"},
    "mention_doc": {"fr": "Document généré par le configurateur DKC – à confirmer avec votre interlocuteur DKC.",
                    "it": "Documento generato dal configuratore DKC – da confermare con il vostro referente DKC.",
                    "en": "Generated by the DKC configurator – to be confirmed with your DKC contact."},
    # Calculateur de métrage
    "calcul": {"fr": "Calculer mon parcours", "it": "Calcolare il percorso", "en": "Calculate my route"},
    "calcul_info": {"fr": "métrage et accessoires → quantités", "it": "metratura e accessori → quantità",
                    "en": "length and fittings → quantities"},
    "longueur_totale": {"fr": "Longueur totale du parcours (m)", "it": "Lunghezza totale del percorso (m)",
                        "en": "Total route length (m)"},
    "avec_couvercles": {"fr": "Avec couvercles", "it": "Con coperchi", "en": "With covers"},
    "acc_parcours": {"fr": "Accessoires sur le parcours", "it": "Accessori sul percorso", "en": "Fittings on the route"},
    "vendu_par": {"fr": "Vendu uniquement en éléments de {l} m", "it": "Venduto solo in elementi da {l} m",
                  "en": "Sold only in {l} m lengths"},
    "resume_metrage": {"fr": "{m} m → {n} éléments de {l} m ({total} m livrés, chute {chute} m)",
                       "it": "{m} m → {n} elementi da {l} m ({total} m forniti, sfrido {chute} m)",
                       "en": "{m} m → {n} lengths of {l} m ({total} m supplied, {chute} m offcut)"},
    "col_motif": {"fr": "Pourquoi", "it": "Perché", "en": "Why"},
    "motif_chemin": {"fr": "{m} m ÷ {l} m, arrondi au-dessus", "it": "{m} m ÷ {l} m, arrotondato per eccesso",
                     "en": "{m} m ÷ {l} m, rounded up"},
    "motif_couvercle": {"fr": "1 couvercle par élément", "it": "1 coperchio per elemento", "en": "1 cover per length"},
    "motif_acc": {"fr": "saisi", "it": "inserito", "en": "entered"},
    "motif_cov_acc": {"fr": "couvercle des accessoires", "it": "coperchio degli accessori", "en": "fitting covers"},
    "motif_gto": {"fr": "{j} jonction(s) × 2 joints", "it": "{j} giunzione/i × 2 giunti", "en": "{j} junction(s) × 2 plates"},
    "ajouter_tout": {"fr": "＋ Ajouter tout à ma liste", "it": "＋ Aggiungi tutto alla lista", "en": "＋ Add all to my list"},
    "calcul_vide": {"fr": "Indiquez la longueur du parcours pour lancer le calcul.",
                    "it": "Indicate la lunghezza del percorso per avviare il calcolo.",
                    "en": "Enter the route length to start the calculation."},
    "hypotheses": {"fr": "Règles de calcul utilisées", "it": "Regole di calcolo utilizzate", "en": "Calculation rules used"},
    "hyp_texte": {
        "fr": "• Le chemin et son couvercle sont vendus en éléments entiers : le nombre est arrondi au-dessus.  \n"
              "• Les éléments standard s'emboîtent entre eux, sans joint.  \n"
              "• Chaque accessoire (courbe, T, croisement) crée une jonction « mâle-mâle » : 2 joints linéaires GTO.  \n"
              "• Épaisseur 1,5 mm (sans emboîtement) : 2 joints GTO à chaque jonction entre éléments.  \n"
              "• Visserie, supports et têtes de fermeture : à ajouter selon le chantier.  \n"
              "Vérifiez les quantités avec votre interlocuteur DKC.",
        "it": "• Il canale e il coperchio sono venduti in elementi interi: il numero è arrotondato per eccesso.  \n"
              "• Gli elementi standard si innestano tra loro, senza giunto.  \n"
              "• Ogni accessorio (curva, T, incrocio) crea una giunzione «maschio-maschio»: 2 giunti lineari GTO.  \n"
              "• Spessore 1,5 mm (senza innesto): 2 giunti GTO a ogni giunzione tra elementi.  \n"
              "• Viteria, supporti e testate di chiusura: da aggiungere secondo il cantiere.  \n"
              "Verificate le quantità con il vostro referente DKC.",
        "en": "• The tray and its cover are sold in whole lengths: the number is rounded up.  \n"
              "• Standard lengths interlock without a joint.  \n"
              "• Each fitting (bend, tee, cross) creates a “male-male” junction: 2 GTO joint plates.  \n"
              "• 1.5 mm thickness (no interlocking): 2 GTO plates at every junction between lengths.  \n"
              "• Screws, supports and end caps: add them according to the site.  \n"
              "Check the quantities with your DKC contact."},
}

FINITIONS = {
    "ZS": {"fr": "Acier zingué Sendzimir (ZS)", "it": "Acciaio zincato Sendzimir (ZS)", "en": "Sendzimir galvanized steel (ZS)"},
    "VB": {"fr": "Vernis bleu (VB)", "it": "Verniciato blu (VB)", "en": "Blue painted (VB)"},
    "VG": {"fr": "Vernis gris RAL 7035 (VG)", "it": "Verniciato grigio RAL 7035 (VG)", "en": "Grey painted RAL 7035 (VG)"},
    "ZF": {"fr": "Galvanisé à chaud (ZF)", "it": "Zincato a caldo (ZF)", "en": "Hot-dip galvanized (ZF)"},
    "IX": {"fr": "Inox AISI 304 (IX)", "it": "Acciaio inox AISI 304 (IX)", "en": "Stainless steel AISI 304 (IX)"},
    "BZE": {"fr": "Acier galvanisé électrolytique (BZE)", "it": "Acciaio zincato elettrolitico (BZE)",
            "en": "Electro-galvanized steel (BZE)"},
    "BZF": {"fr": "Galvanisé à chaud (BZF)", "it": "Zincato a caldo (BZF)", "en": "Hot-dip galvanized (BZF)"},
    "Unique": {"fr": "Finition unique", "it": "Finitura unica", "en": "Single finish"},
}

# Valeurs des choix du chemin (telles qu'elles sont stockées dans les données)
VALEURS = {
    "Foré (perforé)": {"fr": "Foré (perforé)", "it": "Forato", "en": "Perforated"},
    "Non foré (plein)": {"fr": "Non foré (plein)", "it": "Non forato (pieno)", "en": "Solid bottom"},
    "Standard": {"fr": "Standard", "it": "Standard", "en": "Standard"},
    "1,5 mm (sans emboîtement)": {"fr": "1,5 mm (sans emboîtement)", "it": "1,5 mm (senza innesto)",
                                  "en": "1.5 mm (without interlocking)"},
    "Oui": {"fr": "Oui", "it": "Sì", "en": "Yes"},
    "Non": {"fr": "Non", "it": "No", "en": "No"},
}

FAMILLES = {
    "Élément rectiligne": {"fr": "Élément rectiligne", "it": "Elemento rettilineo", "en": "Straight element"},
    "Couvercle": {"fr": "Couvercle", "it": "Coperchio", "en": "Cover"},
    "Accessoire": {"fr": "Accessoire", "it": "Accessorio", "en": "Fitting"},
    "Jonction": {"fr": "Jonction", "it": "Giunzione", "en": "Joint"},
    "Visserie": {"fr": "Visserie", "it": "Bulloneria", "en": "Screws"},
    "Mise à la terre": {"fr": "Mise à la terre", "it": "Messa a terra", "en": "Earthing"},
    "Retouche": {"fr": "Retouche", "it": "Ritocco", "en": "Touch-up"},
}

RUBRIQUES = {
    "Couvrir": {"fr": "Couvrir", "it": "Coprire", "en": "Cover"},
    "Tourner": {"fr": "Tourner", "it": "Curvare", "en": "Turn"},
    "Monter / descendre": {"fr": "Monter / descendre", "it": "Salire / scendere", "en": "Up / down"},
    "Dériver": {"fr": "Dériver", "it": "Derivare", "en": "Branch"},
    "Réduire": {"fr": "Réduire", "it": "Ridurre", "en": "Reduce"},
    "Raccorder": {"fr": "Raccorder", "it": "Collegare", "en": "Connect"},
    "Installer & protéger": {"fr": "Installer & protéger", "it": "Installare e proteggere", "en": "Install & protect"},
    "Visserie & retouche": {"fr": "Visserie & retouche", "it": "Bulloneria e ritocco", "en": "Screws & touch-up"},
}

# Noms italiens des produits (vocabulaire du catalogue Sati Italia)
PRODUITS_IT = {
    "Élément rectiligne (chemin de câbles)": "Elemento rettilineo (canale portacavi)",
    "Couvercle pour élément rectiligne": "Coperchio per elemento rettilineo",
    "Courbe plate à 45°": "Curva piana a 45°",
    "Courbe plate à 90°": "Curva piana a 90°",
    "Dérivation plate en «T»": "Derivazione piana a «T»",
    "Dérivation latérale": "Derivazione laterale",
    "Dérivation plate à croisement": "Incrocio piano",
    "Raccord de réduction centrale": "Riduzione centrale",
    "Raccord de réduction droite": "Riduzione destra",
    "Raccord de réduction gauche": "Riduzione sinistra",
    "Élément de réduction": "Elemento di riduzione",
    "Courbe en montée à 45°": "Curva in salita a 45°",
    "Courbe en montée à 90°": "Curva in salita a 90°",
    "Courbe en descente à 45°": "Curva in discesa a 45°",
    "Courbe en descente à 90°": "Curva in discesa a 90°",
    "Courbe en descente à 90° ligne oblique droite": "Curva in discesa a 90° obliqua destra",
    "Courbe en descente à 90° ligne oblique gauche": "Curva in discesa a 90° obliqua sinistra",
    "Courbe en montée à 90° ligne oblique droite": "Curva in salita a 90° obliqua destra",
    "Courbe en montée à 90° ligne oblique gauche": "Curva in salita a 90° obliqua sinistra",
    "«T» en descente": "«T» in discesa",
    "«T» en montée": "«T» in salita",
    "«T» en descente oblique": "«T» in discesa obliqua",
    "«T» en montée oblique": "«T» in salita obliqua",
    "«T» en montée du couvercle": "«T» in salita da coperchio",
    "«T» en montée oblique du couvercle": "«T» in salita obliqua da coperchio",
    "Courbe en descente à 90° avec inversion de plan": "Curva in discesa a 90° con inversione di piano",
    "Séparateur": "Separatore",
    "Tête de fermeture": "Testata di chiusura",
    "Bride de raccord": "Flangia di raccordo",
    "Joint linéaire": "Giunto lineare",
    "Joint articulé vertical": "Giunto snodato verticale",
    "Joint articulé horizontal": "Giunto snodato orizzontale",
    "Couvre-joint pour couvercle": "Coprigiunto per coperchio",
    "Couvre-joint pour base": "Coprigiunto per fondo",
    "Étrier d'espacement": "Staffa distanziatrice",
    "Kit de protection IP 44 – bande IP 44": "Kit di protezione IP 44 – banda IP 44",
    "Plaque de blocage du couvercle": "Piastrina di bloccaggio coperchio",
    "Kit pièces détachées (couvercle à charnières)": "Kit ricambi (coperchio incernierato)",
    "Plaque de fixation du câble": "Piastrina fissacavi",
    "Platine de terre pour connexions équipotentielles": "Piastrina di terra per collegamenti equipotenziali",
    "Peinture pour retouches (0,40 L)": "Vernice per ritocchi (0,40 L)",
    "Peinture au zinc – spray (0,40 L)": "Vernice allo zinco – spray (0,40 L)",
    "Joint auto-adhésif (kit IP 44)": "Guarnizione autoadesiva (kit IP 44)",
    "Vis M6 x 10 – jonction mécanique (cadre sous-tête)": "Vite M6 x 10 – giunzione meccanica (sottotesta quadra)",
    "Vis M6 x 08 – jonction électrique des bases": "Vite M6 x 08 – giunzione elettrica dei fondi",
    "Vis M5 x 05 – jonction électrique des couvercles": "Vite M5 x 05 – giunzione elettrica dei coperchi",
    "Vis et écrou M6 x 15 galvanisés à chaud": "Vite e dado M6 x 15 zincati a caldo",
    "Écrou M6 anti-desserrage": "Dado M6 antisvitamento",
    "Écrou M8 anti-desserrage": "Dado M8 antisvitamento",
    "Écrou M5": "Dado M5",
    "Rondelle plate pour vis M6": "Rondella piana per vite M6",
    "Joint universel (sur demande, selon marque du canal)": "Giunto universale (su richiesta, secondo la marca del canale)",
}
PRODUITS_EN_EXTRA = {"Élément rectiligne (chemin de câbles)": "Straight element (cable tray)"}
