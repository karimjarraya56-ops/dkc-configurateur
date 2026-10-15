"""Configurateur DKC – Ligne S5 (version 3, FR / IT / EN)
Lancer :  streamlit run app.py
"""
import base64
import math
import re
import unicodedata
from pathlib import Path

import pandas as pd
import streamlit as st

from catalogue_s5 import (CONSEIL_FINITION, CONSEIL_FINITION_TR, DESCRIPTION_CHEMIN_PAR_LANGUE,
                          DESCRIPTIONS_PAR_LANGUE, RUBRIQUES)
from i18n import FAMILLES as FAMILLES_TR
from i18n import FINITIONS as FINITIONS_TR
from i18n import LANGUES, PRODUITS_EN_EXTRA, PRODUITS_IT, TEXTES, VALEURS
from i18n import RUBRIQUES as RUBRIQUES_TR
import export_liste

DATA = Path(__file__).parent / "DKC_S5_base_produits.xlsx"
LOGO = Path(__file__).parent / "assets" / "dkc_logo.png"
PHOTOS = Path(__file__).parent / "assets" / "photos"

# Ordre des choix proposés au client, après le produit (libellés dans i18n.py)
ETAPES = ["Finition", "H_mm", "B_mm", "B1_mm", "Perforé", "Couvercle_charnières", "Épaisseur", "L_mm"]

# Choix du chemin de câbles, dans l'ordre présenté au client
ETAPES_CHEMIN = ["H_mm", "B_mm", "Perforé", "Finition", "Couvercle_charnières", "Épaisseur", "L_mm"]
RECT = "Élément rectiligne (chemin de câbles)"

# Finitions de visserie à proposer selon la finition du chemin
VISSERIE_POUR = {"ZS": ["BZE"], "VB": ["BZE"], "VG": ["BZE"], "ZF": ["BZF", "BZE"], "IX": ["IX"]}
# Dimensions à ne pas comparer (la valeur n'est pas celle du chemin)
DIMS_IGNOREES = {"Joint auto-adhésif (kit IP 44)": ["H_mm"]}

# Familles proposées (ordre d'affichage) et photo qui les représente
FAMILLES = {
    "Élément rectiligne": "element rectiligne fore",
    "Couvercle": "couvercle pour element rectiligne",
    "Accessoire": "courbe plate a 90",
    "Jonction": "joint lineaire",
    "Visserie": "vis et ecrou m6 x 15 galvanises a chaud",
    "Mise à la terre": "platine de terre pour connexions equipotentielles",
    "Retouche": "peinture pour retouches 0 40 l",
}
ABREV_FAMILLE = {"Visserie": "VIS", "Retouche": "VER"}

st.set_page_config(page_title=TEXTES["page_titre"][st.session_state.get("lang") or "fr"], page_icon=str(LOGO), layout="wide")

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
    :root {--dkc: #E30613; --dkc-dark: #9E0410; --dkc-soft: #FDECEE; --ink: #14171C; --muted: #6B7280;
           --line: #E6E8EC; --bg: #F4F5F7;}
    html, body, .stApp, p, label, input, textarea, button, h1, h2, h3, td, th, li {font-family: 'Inter', sans-serif;}
    .stApp {background: var(--bg);}
    #MainMenu, footer, header[data-testid="stHeader"], [data-testid="stToolbar"],
    [data-testid="stToast"], [data-testid="stToastContainer"], [data-testid="stDecoration"] {display: none !important;}
    .block-container {max-width: 1240px; padding: 1.2rem 1.5rem 4rem;}

    /* Barre du haut */
    .topbar {display: flex; align-items: center; justify-content: space-between; background: #fff;
             border-radius: 16px 16px 0 0; padding: 1rem 1.6rem; border: 1px solid var(--line); border-bottom: none;}
    .topbar img {height: 34px; display: block;}
    .topbar .tag {font-size: 0.8rem; color: var(--muted); font-weight: 500;}
    .topbar .tag b {color: var(--dkc);}
    .hero {background: linear-gradient(120deg, var(--dkc) 0%, #C00410 55%, var(--dkc-dark) 100%); color: #fff;
           border-radius: 0 0 16px 16px; padding: 1.8rem 1.6rem 1.6rem; margin-bottom: 1.6rem; position: relative; overflow: hidden;}
    .hero::after {content: ""; position: absolute; right: -60px; top: -80px; width: 320px; height: 320px; border-radius: 50%;
                  border: 40px solid rgba(255,255,255,.07);}
    .hero h1 {font-size: 1.9rem; font-weight: 800; margin: 0; padding: 0; color: #fff; letter-spacing: -0.5px;}
    .hero p {margin: 0.35rem 0 1.1rem; opacity: .9; font-size: 1rem;}
    .stats {display: flex; gap: 0.6rem; flex-wrap: wrap;}
    .stat {background: rgba(255,255,255,.14); border: 1px solid rgba(255,255,255,.22); border-radius: 10px;
           padding: 0.45rem 0.85rem; font-size: 0.82rem;}
    .stat b {font-size: 1rem; margin-right: 0.25rem;}

    /* Étapes */
    .step {display: flex; align-items: center; gap: 0.65rem; font-weight: 700; color: var(--ink);
           margin: 1.6rem 0 0.6rem 0; font-size: 1.02rem;}
    .step .num {background: var(--dkc); color: #fff; width: 28px; height: 28px; border-radius: 8px;
                display: inline-flex; align-items: center; justify-content: center; font-size: 0.85rem; font-weight: 800;}
    .step .done {background: var(--ink);}
    .step small {font-weight: 500; color: var(--muted); font-size: 0.82rem; margin-left: auto;}
    .sous-etape {display: flex; align-items: center; gap: 0.5rem; font-weight: 600; font-size: 0.9rem; color: var(--ink);
                 margin: 1rem 0 0.4rem 2.4rem;}
    .sous-etape::before {content: ""; width: 7px; height: 7px; border-radius: 50%; background: var(--dkc);}
    .sous-etape small {font-weight: 500; color: var(--muted); font-size: 0.8rem; margin-left: auto;}
    .seule {background: #fff; border: 1px dashed var(--line); border-radius: 10px; padding: 0.5rem 0.9rem;
            color: var(--muted); font-size: 0.88rem; display: inline-block;}
    .seule b {color: var(--ink);}

    /* Cartes cliquables (familles et produits) */
    [class*="st-key-card_"] {background: #fff; border: 1.5px solid var(--line); border-radius: 14px; padding: 0.7rem 0.7rem 0.9rem;
                             position: relative; transition: all .15s ease; gap: 0.3rem;}
    [class*="st-key-card_"]:hover {border-color: var(--dkc); transform: translateY(-2px);
                                   box-shadow: 0 8px 20px rgba(227,6,19,.10);}
    [class*="st-key-card_"][class*="_sel"] {border-color: var(--dkc); background: var(--dkc-soft);
                                           box-shadow: 0 0 0 3px rgba(227,6,19,.15);}
    [class*="st-key-card_"] [data-testid="stElementContainer"]:has(button) {position: absolute; inset: 0; margin: 0; z-index: 2;}
    [class*="st-key-card_"] [data-testid="stElementContainer"]:has(button) * {width: 100% !important; height: 100% !important;}
    [class*="st-key-card_"] button {opacity: 0; cursor: pointer;}
    .vign {height: 92px; display: flex; align-items: center; justify-content: center; border-radius: 10px; background: #fff;}
    .vign img {width: 100%; height: 100%; object-fit: contain;}
    .vign.vide {background: var(--bg); color: #B4BAC4; font-weight: 800; font-size: 1rem; letter-spacing: 1px;}
    .card-nom {font-size: 0.8rem; font-weight: 600; color: var(--ink); text-align: center; line-height: 1.25; min-height: 2.5em;}
    .card-sub {font-size: 0.72rem; color: var(--muted); text-align: center; padding-bottom: 0.15rem;}
    [class*="st-key-grille"] {padding-right: 0.4rem;}

    /* Pastilles d'options */
    [data-testid="stButtonGroup"] div {flex-wrap: wrap;}
    [data-testid="stButtonGroup"] button {border-radius: 10px !important; padding: 0.35rem 0.95rem !important;
                                          background: #fff; font-weight: 500;}

    /* Panneau de droite */
    .st-key-panneau, .st-key-panneau2 {position: sticky; top: 1.2rem;}
    .panel {background: #fff; border-radius: 18px; border: 1px solid var(--line); overflow: hidden;
            box-shadow: 0 1px 2px rgba(0,0,0,.04), 0 12px 32px rgba(20,23,28,.07);}
    .panel .visuel {height: 230px; display: flex; align-items: center; justify-content: center; padding: 1rem;
                    background: radial-gradient(circle at 50% 40%, #fff 0%, #F3F4F6 100%); border-bottom: 1px solid var(--line);}
    .panel .visuel img {width: 100%; height: 100%; max-height: 200px; object-fit: contain;}
    .panel .visuel.vide {color: #C3C8D0; font-size: 2rem; font-weight: 800; letter-spacing: 2px;}
    .panel .corps {padding: 1.2rem 1.4rem 1.3rem;}
    .label {color: var(--muted); font-size: 0.7rem; text-transform: uppercase; letter-spacing: 1.6px; font-weight: 700;}
    .code {font-size: 2.3rem; font-weight: 800; color: var(--dkc); letter-spacing: 2px; line-height: 1.1; margin-top: 0.15rem;}
    .nom {font-weight: 700; color: var(--ink); font-size: 1rem; margin-top: 0.3rem;}
    .chips {display: flex; flex-wrap: wrap; gap: 0.35rem; margin-top: 0.8rem;}
    .chip {background: var(--bg); border: 1px solid var(--line); border-radius: 999px; padding: 0.18rem 0.65rem;
           font-size: 0.75rem; color: var(--ink); font-weight: 500;}
    .panel table {border-collapse: collapse; width: 100%; margin-top: 1rem; border: none;}
    .panel td {padding: 0.5rem 0; border: none; border-top: 1px solid var(--line); font-size: 0.87rem; vertical-align: top;}
    .panel td:first-child {color: var(--muted); width: 46%;}
    .panel td:last-child {font-weight: 600; color: var(--ink);}
    .attente {text-align: center; padding: 0.4rem 0 0.2rem;}
    .attente .gros {font-size: 2.2rem; font-weight: 800; color: var(--ink); line-height: 1;}
    .barre {height: 6px; background: var(--bg); border-radius: 99px; overflow: hidden; margin: 0.9rem 0 0.4rem;}
    .barre div {height: 100%; background: var(--dkc); border-radius: 99px;}
    .muted {color: var(--muted); font-size: 0.84rem;}
    .extra {display: flex; align-items: center; gap: 0.8rem; padding: 0.6rem 0; border-top: 1px solid var(--line);}
    .extra .mini {width: 54px; height: 44px; flex: 0 0 54px; display: flex; align-items: center; justify-content: center;
                  background: var(--bg); border-radius: 8px; overflow: hidden;}
    .extra .mini img {max-width: 100%; max-height: 100%;}
    .extra .t {font-size: 0.82rem; color: var(--muted);}
    .extra .c {font-weight: 800; color: var(--ink); letter-spacing: 1px;}
    @media (max-width: 640px) {
      [data-testid="stHorizontalBlock"]:has([class*="st-key-card_"]):not(:has(.st-key-panneau, .st-key-panneau2)) {flex-direction: row; flex-wrap: wrap; gap: 0.5rem;}
      [data-testid="stHorizontalBlock"]:has([class*="st-key-card_"]):not(:has(.st-key-panneau, .st-key-panneau2)) > [data-testid="stColumn"] {
        flex: 1 1 calc(50% - 0.5rem) !important; min-width: calc(50% - 0.5rem) !important; width: calc(50% - 0.5rem) !important;}
      .topbar .tag {display: none;}
      .hero h1 {font-size: 1.5rem;}
      .st-key-panneau, .st-key-panneau2 {position: static;}
    }
    /* Cartes accessoires */
    [class*="st-key-acc_"] {background: #fff; border: 1px solid var(--line); border-radius: 14px; padding: 0.8rem 0.9rem 1rem;
                            gap: 0.35rem;}
    [class*="st-key-acc_"]:hover {border-color: #F3B3B7; box-shadow: 0 6px 18px rgba(20,23,28,.06);}
    [class*="st-key-acc_"] .vign {height: 110px;}
    
    .acc-nom {font-weight: 700; font-size: 0.9rem; color: var(--ink); margin-top: 0.4rem; line-height: 1.25;}
    .acc-nom span {color: var(--dkc); font-weight: 700; font-size: 0.75rem; margin-left: 0.35rem;}
    .acc-desc {font-size: 0.8rem; color: var(--muted); line-height: 1.4; margin: 0.2rem 0 0.35rem; min-height: 3.4em;}
    .acc-ref {display: flex; align-items: baseline; justify-content: space-between; border-top: 1px solid var(--line);
              padding-top: 0.5rem; margin-top: 0.2rem;}
    [class*="st-key-acc_"] .stButton {margin-top: 0.45rem;}
    .acc-ref b {font-size: 1.15rem; font-weight: 800; color: var(--dkc); letter-spacing: 1px;}
    .acc-ref span {font-size: 0.72rem; color: var(--muted);}
    .acc-cov {font-size: 0.78rem; color: var(--ink); background: var(--bg); border-radius: 8px; padding: 0.3rem 0.55rem;
              margin-top: 0.35rem;}
    .acc-cov b {letter-spacing: 0.5px;}
    .indispo {font-size: 0.8rem; color: var(--muted); margin-top: 0.8rem; line-height: 1.6;}
    .bandeau {background: #fff; border: 1px solid var(--line); border-left: 5px solid var(--dkc); border-radius: 12px;
              padding: 0.8rem 1.1rem; margin: 0.4rem 0 0.8rem; display: flex; gap: 1.2rem; align-items: center;}
    .bandeau .code {font-size: 1.4rem;}
    .bandeau .muted {font-size: 0.85rem;}
    .desc {font-size: 0.85rem; color: var(--muted); margin-top: 0.6rem; line-height: 1.45;}
    .conseil table {width: 100%; font-size: 0.8rem; border-collapse: collapse;}
    .conseil td, .conseil th {padding: 0.3rem 0.4rem; border-bottom: 1px solid var(--line); text-align: center;}
    .conseil td:first-child, .conseil th:first-child {text-align: left;}
    .conseil .ok {color: #15803D; font-weight: 600;} .conseil .moy {color: #B45309;} .conseil .non {color: #B91C1C;}
    .stTabs [data-baseweb="tab"] {font-weight: 600;}
    .st-key-langue [data-testid="stButtonGroup"] {display: flex; justify-content: flex-end;}
    .st-key-langue button {font-size: 0 !important; padding: 0.4rem 0.65rem !important; min-width: 0 !important;
        gap: 0 !important; justify-content: center !important;}
    .st-key-langue button > * {display: none !important;}
    .st-key-langue button::before {content: ""; display: inline-block; width: 26px; height: 17px; border-radius: 3px;
        background-size: 100% 100%; box-shadow: 0 0 0 1px rgba(0,0,0,.12);}
    .st-key-langue button:nth-of-type(1)::before {background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 3 2'%3E%3Cpath fill='%23002395' d='M0 0h1v2H0z'/%3E%3Cpath fill='%23fff' d='M1 0h1v2H1z'/%3E%3Cpath fill='%23ED2939' d='M2 0h1v2H2z'/%3E%3C/svg%3E");}
    .st-key-langue button:nth-of-type(2)::before {background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 3 2'%3E%3Cpath fill='%23009246' d='M0 0h1v2H0z'/%3E%3Cpath fill='%23fff' d='M1 0h1v2H1z'/%3E%3Cpath fill='%23CE2B37' d='M2 0h1v2H2z'/%3E%3C/svg%3E");}
    .st-key-langue button:nth-of-type(3)::before {background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 60 30'%3E%3CclipPath id='t'%3E%3Cpath d='M30 15h30v15zv15H0zH0V0zV0h30z'/%3E%3C/clipPath%3E%3Cpath fill='%23012169' d='M0 0v30h60V0z'/%3E%3Cpath stroke='%23fff' stroke-width='6' d='M0 0l60 30m0-30L0 30'/%3E%3Cpath stroke='%23C8102E' stroke-width='4' clip-path='url(%23t)' d='M0 0l60 30m0-30L0 30'/%3E%3Cpath stroke='%23fff' stroke-width='10' d='M30 0v30M0 15h60'/%3E%3Cpath stroke='%23C8102E' stroke-width='6' d='M30 0v30M0 15h60'/%3E%3C/svg%3E");}
    [class*="st-key-acc_"] .stButton button {padding: 0.25rem 0.6rem; min-height: 0; font-size: 0.8rem; border-radius: 9px;}
    .dans-liste {font-size: 0.75rem; color: #15803D; font-weight: 600;}
    .st-key-panneau .stButton button, .st-key-panneau2 .stButton button {border-radius: 10px;}
    .st-key-calcul {background: #fff; border: 1px solid var(--line); border-radius: 16px; padding: 1.1rem 1.2rem;}
    .resume {display: flex; align-items: center; gap: 0.8rem; background: #FFF1F2; border-left: 4px solid var(--dkc);
      border-radius: 10px; padding: 0.7rem 1rem; margin: 0.6rem 0; font-size: 0.95rem;}
    .resume b {color: var(--dkc); font-size: 1.25rem;}

    /* ===== Design v4 ===== */
    .block-container {padding-top: 0.8rem;}
    div:has(> .st-key-entete) {position: sticky; top: 0.4rem; z-index: 100;}
    .st-key-entete {background: rgba(255,255,255,.94);
      backdrop-filter: blur(10px); border: 1px solid var(--line); border-radius: 16px; padding: 0.5rem 1rem;
      box-shadow: 0 6px 20px rgba(20,23,28,.06); margin-bottom: 0.9rem;}
    .logo-entete img {height: 30px; display: block;}
    .st-key-modes [data-testid="stButtonGroup"] {justify-content: center;}
    .st-key-modes button {background: transparent !important; border: none !important; border-radius: 0 !important;
      border-bottom: 2.5px solid transparent !important; padding: 0.45rem 0.9rem !important; font-weight: 600 !important;
      color: var(--muted) !important; box-shadow: none !important;}
    .st-key-modes button:hover {color: var(--ink) !important;}
    .st-key-modes button[aria-checked="true"] {color: var(--dkc) !important; border-bottom-color: var(--dkc) !important;}
    .st-key-modes button[aria-checked="true"] * {color: var(--dkc) !important;}
    .badge-liste {display: inline-flex; align-items: center; gap: 0.5rem; background: var(--ink); color: #fff !important;
      border-radius: 999px; padding: 0.42rem 0.55rem 0.42rem 0.95rem; font-weight: 700; font-size: 0.85rem;
      text-decoration: none !important; white-space: nowrap; transition: background .15s;}
    .badge-liste:hover {background: var(--dkc);}
    .badge-liste span {background: #fff; color: var(--ink); border-radius: 999px; min-width: 24px; height: 24px;
      display: inline-flex; align-items: center; justify-content: center; font-size: 0.78rem; padding: 0 0.4rem;}
    .badge-liste.plein span {background: var(--dkc); color: #fff;}
    .hero {border-radius: 16px; padding: 1.25rem 1.6rem; margin-bottom: 1rem; display: flex; align-items: center;
      justify-content: space-between; gap: 1.5rem;}
    .hero h1 {font-size: 1.55rem;}
    .hero p {margin: 0.25rem 0 0; font-size: 0.92rem;}
    .hero-tag {font-size: 0.72rem; text-transform: uppercase; letter-spacing: 1.5px; opacity: .85; margin-bottom: 0.3rem;}
    .hero .stats {flex-wrap: nowrap; position: relative; z-index: 1;}
    .hero .stat {text-align: center; padding: 0.5rem 0.9rem; line-height: 1.2;}
    .hero .stat b {display: block; font-size: 1.25rem; margin: 0;}
    .stepper {display: flex; gap: 0.3rem; background: #fff; border: 1px solid var(--line); border-radius: 14px;
      padding: 0.4rem; margin-bottom: 0.3rem; overflow-x: auto;}
    .stepper a {flex: 1; display: flex; align-items: center; justify-content: center; gap: 0.5rem; padding: 0.45rem 0.6rem;
      border-radius: 10px; text-decoration: none !important; color: var(--muted) !important; font-size: 0.83rem;
      font-weight: 600; white-space: nowrap; transition: background .15s;}
    .stepper a:hover {background: var(--bg);}
    .stepper .n {width: 22px; height: 22px; border-radius: 50%; background: var(--bg); border: 1px solid var(--line);
      display: inline-flex; align-items: center; justify-content: center; font-size: 0.72rem; font-weight: 800;}
    .stepper a.done {color: var(--ink) !important;}
    .stepper a.done .n {background: #15803D; border-color: #15803D; color: #fff;}
    .stepper a.actif {background: var(--dkc-soft); color: var(--dkc) !important;}
    .stepper a.actif .n {background: var(--dkc); border-color: var(--dkc); color: #fff;}
    .stepper .sep {color: #CBD0D8; align-self: center;}
    .step {scroll-margin-top: 90px;}
    #ma-liste {scroll-margin-top: 90px;}
    .step .num {box-shadow: 0 3px 8px rgba(227,6,19,.25);}
    .st-key-carte_chemin {background: #fff; border: 1px solid var(--line); border-radius: 16px; padding: 0.4rem 1.2rem 1rem;}
    .st-key-carte_chemin .step {margin-top: 0.8rem;}
    .lib-choix {font-weight: 600; font-size: 0.88rem; color: var(--ink); display: flex; align-items: center; gap: 0.45rem;}
    .lib-choix::before {content: ""; width: 6px; height: 6px; border-radius: 50%; background: var(--dkc); flex: 0 0 6px;}
    .lib-choix small {display: block; font-weight: 500; color: var(--muted); font-size: 0.72rem;}
    [class*="st-key-ligne_choix_"] {border-top: 1px solid var(--line); padding-top: 0.55rem;}
    .st-key-carte_chemin [data-testid="stButtonGroup"] button[aria-checked="true"] {box-shadow: 0 0 0 2px rgba(227,6,19,.15);
      font-weight: 600;}
    @media (max-width: 640px) {
      div:has(> .st-key-entete) {position: static;}
      .st-key-entete [data-testid="stHorizontalBlock"] {flex-direction: row; flex-wrap: wrap; align-items: center; gap: 0.5rem;}
      .st-key-entete [data-testid="stColumn"] {width: auto !important; flex: 0 0 auto !important; min-width: 0 !important;}
      .st-key-entete [data-testid="stColumn"]:nth-child(2) {order: 9; flex: 1 1 100% !important;}
      .st-key-entete [data-testid="stColumn"]:nth-child(1) {flex: 1 1 auto !important;}
      .logo-entete img {height: 24px;}
      .badge-liste {padding: 0.32rem 0.45rem 0.32rem 0.75rem; font-size: 0.78rem;}
      .st-key-modes [data-testid="stButtonGroup"] {justify-content: flex-start;}
      .st-key-modes button {padding: 0.35rem 0.5rem !important; font-size: 0.8rem !important;}
      .st-key-langue button {padding: 0.3rem 0.45rem !important;}
      .hero {flex-direction: column; align-items: flex-start; padding: 1rem 1.1rem;}
      .hero p {display: none;}
      .hero h1 {font-size: 1.25rem;}
      .hero .stat {padding: 0.35rem 0.6rem;}
      .stepper a .lib {display: none;}
      .stepper a.actif .lib {display: inline;}
    }
    .vendu {display: inline-block; font-size: 0.75rem; font-weight: 700; color: var(--dkc); background: #FFF1F2;
      border: 1px solid #FECDD3; border-radius: 999px; padding: 0.15rem 0.6rem; margin-bottom: 0.4rem;}
    .titre-extra {font-weight: 700; font-size: 0.85rem; margin-top: 1.1rem; color: var(--ink);}
    </style>
    """,
    unsafe_allow_html=True,
)


def slug(texte):
    t = unicodedata.normalize("NFKD", texte).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]+", "_", t).strip("_")


@st.cache_data
def _encoder(chemin, _mtime):
    # _mtime fait partie de la clé de cache : une photo remplacée est relue automatiquement
    ext = Path(chemin).suffix[1:]
    mime = "jpeg" if ext in ("jpg", "jpeg") else ext
    return f"data:image/{mime};base64," + base64.b64encode(Path(chemin).read_bytes()).decode()


def _photo_fichier(s):
    for ext in ("jpg", "jpeg", "png", "webp"):
        f = PHOTOS / f"{s}.{ext}"
        if f.exists():
            return _encoder(str(f), f.stat().st_mtime)
    return None


def md(html):
    """Affiche du HTML (lignes collées pour que Markdown n'en fasse pas un bloc de code)."""
    st.markdown(" ".join(l.strip() for l in html.splitlines()), unsafe_allow_html=True)


def chemin_photo(noms):
    """Fichier de la première photo trouvée parmi `noms`, sinon None."""
    for nom_ in ([noms] if isinstance(noms, str) or noms is None else noms):
        for ext in ("jpg", "jpeg", "png", "webp"):
            if nom_ and (f := PHOTOS / f"{slug(nom_)}.{ext}").exists():
                return f
    return None


def photo_src(noms):
    """Première photo trouvée parmi `noms` (assets/photos/<slug>.jpg|png|webp), sinon None."""
    for nom in ([noms] if isinstance(noms, str) or noms is None else noms):
        if nom and (src := _photo_fichier(slug(nom))):
            return src
    return None


def img_html(src, classe, texte_vide):
    if src:
        return f'<div class="{classe}"><img src="{src}"></div>'
    return f'<div class="{classe} vide">{texte_vide}</div>'


def etape(num, label, info=""):
    info = f"<small>{info}</small>" if info else ""
    if num == "":  # sous-étape (sans numéro)
        st.markdown(f'<div class="sous-etape">{label}{info}</div>', unsafe_allow_html=True)
        return
    st.markdown(f'<div class="step" id="etape-{num}"><span class="num">{num}</span>{label}{info}</div>',
                unsafe_allow_html=True)


@st.cache_data
def charger():
    df = pd.read_excel(DATA, sheet_name="Produits_S5", dtype={"Code": str})
    # Codes à 7 chiffres, plus les articles « Sur demande » (ex. joint universel)
    df = df[df["Code"].str.fullmatch(r"\d{7}|Sur demande", na=False)].copy()
    # Les éléments rectilignes (foré, non foré, 1,5 mm, charnières) deviennent un seul produit
    rect = (df["Famille"] == "Élément rectiligne") & df["Produit"].str.startswith("Élément rectiligne")
    df.loc[rect, "Produit"] = RECT
    df["Perforé"] = df["Perforé"].map({"Oui": "Foré (perforé)", "Non": "Non foré (plein)"})
    df["Épaisseur"] = df["Épaisseur_mm"].apply(lambda v: "1,5 mm (sans emboîtement)" if pd.notna(v) else "Standard")
    df.loc[~rect, "Épaisseur"] = None
    df.loc[~rect, "Couvercle_charnières"] = None
    return df


# ---------- Langue ----------
LANG = st.session_state.get("lang") or "fr"


def t(cle):
    return TEXTES[cle][LANG]


def tr(dico, cle):
    """Traduction d'une valeur (dico = {valeur: {fr, it, en}}), la valeur elle-même sinon."""
    return dico.get(cle, {}).get(LANG, cle)


def nom(produit):
    """Nom du produit dans la langue choisie."""
    if LANG == "it":
        if produit.startswith("Couvercle – "):
            return "Coperchio – " + PRODUITS_IT.get(produit[12:], produit[12:])
        return PRODUITS_IT.get(produit, produit)
    if LANG == "en":
        return PRODUITS_EN.get(produit, produit)
    return produit


def fmt(col, v):
    if col == "Finition":
        return tr(FINITIONS_TR, v)
    if col in ("Perforé", "Épaisseur", "Couvercle_charnières"):
        return tr(VALEURS, v)
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def unite(u):
    return {"it": {"pcs": "pz"}, "fr": {"pcs": "pcs"}}.get(LANG, {}).get(u, u)


def pluriel(n):
    return f"{n} {t('produit_1') if n <= 1 else t('produit_n')}"


def chip(col, v):
    """Texte court d'une étiquette de résumé (ex. « H 50 », « Sans charnières »)."""
    if col in ("H_mm", "B_mm", "B1_mm", "L_mm"):
        return f"{col.split('_')[0]} {fmt(col, v)}"
    if col == "Finition":
        return v
    if col == "Couvercle_charnières":
        return t("avec_charnieres") if v == "Oui" else t("sans_charnieres")
    return fmt(col, v).split(" (")[0]


def grille(items, etat, nb_cols, prefixe):
    """Cartes cliquables avec photo. items = [(valeur, libellé, src_photo, texte_vide, sous_titre)]."""
    def choisir(v):
        st.session_state[etat] = v

    for i in range(0, len(items), nb_cols):
        cols = st.columns(nb_cols, gap="small")
        for col, (val, lib, src, vide, sous) in zip(cols, items[i:i + nb_cols]):
            sel = "_sel" if st.session_state.get(etat) == val else ""
            with col, st.container(key=f"card_{prefixe}_{slug(val)}{sel}"):
                st.markdown(img_html(src, "vign", vide) + f'<div class="card-nom">{lib}</div>'
                            + (f'<div class="card-sub">{sous}</div>' if sous else ""), unsafe_allow_html=True)
                st.button(lib, key=f"btn_{prefixe}_{slug(val)}", on_click=choisir, args=(val,))


def choix(df, col, num, prefixe="opt"):
    label = t(col)
    vals = df[col].dropna().unique().tolist()
    ordre = list(FINITIONS_TR)
    vals = sorted(vals, key=lambda x: (ordre.index(x) if x in ordre else 99, isinstance(x, str), x))
    if num == "" and prefixe == "ch":
        return choix_ligne(df, col, label, vals, prefixe)
    if len(vals) == 1:
        etape(num, label)
        st.markdown(f'<span class="seule">{t("seule_option")}{" :" if LANG == "fr" else ":"} <b>{fmt(col, vals[0])}</b></span>',
                    unsafe_allow_html=True)
        return df[df[col] == vals[0]], vals[0]
    etape(num, label, f"{len(vals)} {t('options')}")
    # La clé dépend des options : si elles changent, le choix repart sur la première
    v = st.pills(label, vals, format_func=lambda x: fmt(col, x), default=vals[0], label_visibility="collapsed",
                 key=f"{prefixe}_{col}_{abs(hash(tuple(map(str, vals))))}")
    v = vals[0] if v is None else v
    return df[df[col] == v], v


df = charger()
s5 = df[df["Ligne"] == "S5"]
NOMS = set(s5["Produit"])
PRODUITS_EN = {**s5.groupby("Produit")["Product_EN"].first().dropna().to_dict(), **PRODUITS_EN_EXTRA}


def choix_ligne(df, col, label, vals, prefixe):
    """Choix compact sur une ligne : libellé à gauche, options à droite."""
    with st.container(key=f"ligne_choix_{prefixe}_{slug(col)}"):
        g, d = st.columns([1, 3.2], vertical_alignment="center")
        with g:
            info = t("seule_option") if len(vals) == 1 else f"{len(vals)} {t('options')}"
            md(f'<div class="lib-choix"><div>{label}<small>{info}</small></div></div>')
        with d:
            if len(vals) == 1:
                md(f'<span class="seule"><b>{fmt(col, vals[0])}</b></span>')
                return df[df[col] == vals[0]], vals[0]
            v = st.pills(label, vals, format_func=lambda x: fmt(col, x), default=vals[0], label_visibility="collapsed",
                         key=f"{prefixe}_{col}_{abs(hash(tuple(map(str, vals))))}")
    v = vals[0] if v is None else v
    return df[df[col] == v], v


def finitions_pour(fin):
    """Finitions acceptables pour un accessoire, par ordre de préférence."""
    return [fin, "Unique"] + VISSERIE_POUR.get(fin, [])


def compatibles(produit, r):
    """Lignes du produit compatibles avec le chemin r, ou (None, raison)."""
    g = s5[s5["Produit"] == produit]
    for f in finitions_pour(r["Finition"]):
        if (g["Finition"] == f).any():
            g = g[g["Finition"] == f]
            break
    else:
        return None, f"{t('existe_pas')} {r['Finition']}"
    for col, lab in (("H_mm", "H"), ("B_mm", "B"), ("L_mm", "L")):
        if col in DIMS_IGNOREES.get(produit, []) or not g[col].notna().any():
            continue
        m = g[g[col] == r[col]]
        if col == "L_mm":
            g = m if len(m) else g  # longueur absente : on propose les longueurs existantes
        elif m.empty:
            return None, f"{t('existe_pas')} {lab} {fmt(col, r[col])}"
        else:
            g = m
    return g, ""


def abrev_de(g):
    return g["Abrév"].dropna().iloc[0] if g["Abrév"].notna().any() else ""


def carte_accessoire(produit, g, r):
    with st.container(key=f"acc_{slug(produit)}"):
        ab = abrev_de(g)
        md(img_html(photo_src(produit), "vign", ab or "DKC")
           + f'<div class="acc-nom">{nom(produit)}{f"<span>{ab}</span>" if ab else ""}</div>'
           + f'<div class="acc-desc">{DESCRIPTIONS_PAR_LANGUE[LANG].get(produit, "")}</div>')
        # Plusieurs versions possibles (base de sortie B1, longueur) : le client choisit
        for col, lab in (("B1_mm", t("sortie_b1")), ("L_mm", "L")):
            vals = sorted(g[col].dropna().unique())
            if len(vals) > 1:
                v = st.pills(lab, vals, format_func=lambda x, lab=lab, col=col: f"{lab} {fmt(col, x)}",
                             default=vals[0], label_visibility="collapsed",
                             key=f"var_{slug(produit)}_{col}_{abs(hash(tuple(vals)))}")
                g = g[g[col] == (vals[0] if v is None else v)]
        ligne = g.iloc[0]
        cov_html = ""
        couvercle = None
        if f"Couvercle – {produit}" in NOMS:
            cg, _ = compatibles(f"Couvercle – {produit}", r)
            if cg is not None and pd.notna(ligne["B1_mm"]):
                cg = cg[cg["B1_mm"] == ligne["B1_mm"]]
            if cg is not None and len(cg):
                cov_html = f'<div class="acc-cov">{t("plus_couvercle")} <b>{cg.iloc[0]["Code"]}</b></div>'
                couvercle = cg.iloc[0]
        cond = (f"{int(ligne['Conditionnement'])} {unite(ligne['Unité_cond'])}"
                if pd.notna(ligne["Conditionnement"]) else "")
        fin = "" if ligne["Finition"] in (r["Finition"], "Unique") else f" · {ligne['Finition']}"
        md(f'<div class="acc-ref"><b>{ligne["Code"]}</b><span>{cond}{fin}</span></div>{cov_html}')
        boutons_ajout(ligne, [produit], couvercle, [f"Couvercle – {produit}", produit])


def conseil_finition():
    classe = {"Excellent": "ok", "Bon": "ok", "Possible": "moy", "Déconseillé": "non"}

    def libelle(groupe, v):
        return v if LANG == "fr" else CONSEIL_FINITION_TR[groupe][v][LANG]

    tete = "".join(f"<th>{c}</th>" for c in CONSEIL_FINITION_TR["colonnes"][LANG])
    lignes = "".join(f"<tr><td>{libelle('environnements', env)}</td>"
                     + "".join(f'<td class="{classe[v]}">{libelle("valeurs", v)}</td>' for v in vals) + "</tr>"
                     for env, vals in CONSEIL_FINITION["lignes"])
    md(f'<div class="conseil"><table><tr><th>{t("environnement")}</th>{tete}</tr>{lignes}</table>'
       f'<div class="muted" style="margin-top:0.5rem">{t("avertissement_finition")}</div></div>')


def panneau_chemin(sel):
    """Carte de droite : le chemin de câbles choisi."""
    if len(sel) != 1:
        md(f"""<div class="panel">{img_html(photo_src("element rectiligne fore"), "visuel", "DKC")}
            <div class="corps attente"><div class="label">{t("chemins_possibles")}</div>
            <div class="gros">{len(sel)}</div><div class="nom">{t("votre_chemin")}</div>
            <div class="muted">{t("faites_choix")}</div></div></div>""")
        return
    r = sel.iloc[0]
    fore = str(r["Perforé"]).startswith("Foré")
    charn = r["Couvercle_charnières"] == "Oui"
    textes = DESCRIPTION_CHEMIN_PAR_LANGUE[LANG]
    desc = [textes[r["Perforé"]]]
    if charn:
        desc.append(textes["charnières"])
    if str(r["Épaisseur"]).startswith("1,5"):
        desc.append(textes["1,5 mm"])
    details = [(t("dimensions"), f"H {fmt('H_mm', r['H_mm'])} × B {fmt('B_mm', r['B_mm'])} × L {fmt('L_mm', r['L_mm'])}"),
               (t("finition"), fmt("Finition", r["Finition"]))]
    if pd.notna(r["TUA_cm2"]):
        details.append((t("tua"), f"{int(r['TUA_cm2'])} cm²"))
    if pd.notna(r["Conditionnement"]):
        details.append((t("conditionnement"), f"{int(r['Conditionnement'])} {unite(r['Unité_cond'])}"))
    couvercle = None
    if not charn:
        cg, _ = compatibles("Couvercle pour élément rectiligne", r)
        if cg is not None and len(cg):
            couvercle = cg.iloc[0]
            details.append((t("couvercle_assorti"), f"<b>{couvercle['Code']}</b>"))
    details.append((t("page_catalogue"), str(int(r["Page_catalogue"]))))
    rows = "".join(f"<tr><td>{k}</td><td>{v}</td></tr>" for k, v in details)
    titre = t("chemin_fore" if fore else "chemin_non_fore") + (t("a_charnieres") if charn else "")
    variante = "fore" if fore else "non fore"
    photos = [f"element rectiligne {variante}"]
    if charn:
        photos.insert(0, f"element rectiligne charnieres {variante}")
    md(f"""<div class="panel">{img_html(photo_src(photos), "visuel", "DKC")}
        <div class="corps">
          <div class="label">{t("votre_chemin")}</div>
          <div class="code">{r['Code']}</div>
          <div class="nom">{titre}</div>
          <div class="desc">{" ".join(desc)}</div>
          <table>{rows}</table>
        </div></div>""")
    boutons_ajout(r, photos, couvercle, ["couvercle pour element rectiligne"], chemin=True)


# ---------- Calculateur de métrage ----------
ACC_CALCUL = ["Courbe plate à 90°", "Courbe plate à 45°", "Dérivation plate en «T»", "Dérivation plate à croisement",
              "Courbe en montée à 90°", "Courbe en descente à 90°", "Courbe en montée à 45°", "Courbe en descente à 45°",
              "Tête de fermeture"]
SANS_JONCTION = {"Tête de fermeture"}  # se pose en bout de parcours, sans joint


def photos_chemin(r):
    variante = "fore" if str(r["Perforé"]).startswith("Foré") else "non fore"
    photos = [f"element rectiligne {variante}"]
    if r["Couvercle_charnières"] == "Oui":
        photos.insert(0, f"element rectiligne charnieres {variante}")
    return photos


def nombre(x):
    """3.0 → « 3 », 2.5 → « 2,5 » (virgule en FR/IT)."""
    txt = f"{x:g}"
    return txt if LANG == "en" else txt.replace(".", ",")


def calcul_parcours(r, metres, qtes, avec_cov):
    """Lignes (série produit, photos, quantité, motif) pour un parcours."""
    lignes = []
    long_m = r["L_mm"] / 1000
    n = math.ceil(round(metres / long_m, 6)) if metres > 0 else 0
    if n:
        lignes.append((r, photos_chemin(r), n, t("motif_chemin").format(m=nombre(metres), l=nombre(long_m))))
        if avec_cov and r["Couvercle_charnières"] != "Oui":
            cg, _ = compatibles("Couvercle pour élément rectiligne", r)
            if cg is not None and len(cg):
                lignes.append((cg.iloc[0], ["couvercle pour element rectiligne"], n, t("motif_couvercle")))
    jonctions = 0
    for produit, q in qtes.items():
        if not q:
            continue
        g, _ = compatibles(produit, r)
        lignes.append((g.iloc[0], [produit], q, t("motif_acc")))
        if produit not in SANS_JONCTION:
            jonctions += q
        if avec_cov and f"Couvercle – {produit}" in NOMS:
            cg, _ = compatibles(f"Couvercle – {produit}", r)
            if cg is not None and len(cg):
                lignes.append((cg.iloc[0], [f"Couvercle – {produit}", produit], q, t("motif_cov_acc")))
    if str(r["Épaisseur"]).startswith("1,5") and n > 1:
        jonctions += n - 1
    if jonctions:
        jg, _ = compatibles("Joint linéaire", r)
        if jg is not None and len(jg):
            lignes.append((jg.iloc[0], ["joint lineaire"], 2 * jonctions, t("motif_gto").format(j=jonctions)))
    return lignes, n, long_m


def tout_ajouter(lignes):
    for ligne, photos, q, _ in lignes:
        ajouter(ligne.to_dict(), photos, q)


def bloc_calcul(r):
    etape(3, t("calcul"), t("calcul_info"))
    with st.container(key="calcul"):
        long_m = r["L_mm"] / 1000
        md(f'<span class="vendu">{t("vendu_par").format(l=nombre(long_m))}</span>')
        c1, c2 = st.columns([1.3, 1], vertical_alignment="bottom")
        with c1:
            metres = st.number_input(t("longueur_totale"), min_value=0.0, max_value=10000.0, step=1.0,
                                     value=0.0, format="%g", key="calc_m")
        with c2:
            avec_cov = st.toggle(t("avec_couvercles"), value=True, key="calc_cov")
        etape("", t("acc_parcours"))
        dispo = [p for p in ACC_CALCUL if p in NOMS and compatibles(p, r)[0] is not None]
        qtes = {}
        for k in range(0, len(dispo), 3):
            cols = st.columns(3, gap="small")
            for col, produit in zip(cols, dispo[k:k + 3]):
                with col:
                    qtes[produit] = st.number_input(nom(produit), min_value=0, max_value=999, step=1, value=0,
                                                    key=f"calc_{slug(produit)}")
        lignes, n, _ = calcul_parcours(r, metres, qtes, avec_cov)
        if not lignes:
            st.caption(t("calcul_vide"))
        else:
            if n:
                total = n * long_m
                md(f'<div class="resume">📏 <span>{t("resume_metrage").format(m=nombre(metres), n=f"<b>{n}</b>", l=nombre(long_m), total=nombre(total), chute=nombre(round(total - metres, 3)))}</span></div>')
            st.dataframe(pd.DataFrame([{
                "photo": photo_src(ph), "code": lg["Code"],
                "designation": designation({"produit": lg["Produit"], "abrev": lg["Abrév"] if pd.notna(lg["Abrév"]) else "",
                                            "perfore": lg["Perforé"], "charn": lg["Couvercle_charnières"], "ep": lg["Épaisseur"]}),
                "quantite": q, "motif": motif} for lg, ph, q, motif in lignes]),
                hide_index=True, width="stretch", row_height=56, height=40 + 56 * len(lignes),
                column_config={
                    "photo": st.column_config.ImageColumn(t("col_photo"), width="small"),
                    "code": st.column_config.TextColumn(t("col_code"), width="small"),
                    "designation": st.column_config.TextColumn(t("col_designation"), width="large"),
                    "quantite": st.column_config.NumberColumn(t("col_quantite"), width="small"),
                    "motif": st.column_config.TextColumn(t("col_motif")),
                })
            st.button(t("ajouter_tout"), key="calc_ajouter", type="primary", on_click=tout_ajouter, args=(lignes,))
        with st.expander(t("hypotheses")):
            st.markdown(t("hyp_texte"))


def bloc_accessoires(r):
    charn = r["Couvercle_charnières"] == "Oui"
    etape(4, t("accessoires"), f"H {fmt('H_mm', r['H_mm'])} × B {fmt('B_mm', r['B_mm'])} · {r['Finition']}")
    md(f"""<div class="bandeau"><div><div class="label">{t("chemin_choisi")}</div><div class="code">{r['Code']}</div></div>
        <div class="muted">{t("bandeau")}</div></div>""")
    onglets = []
    for titre, produits in RUBRIQUES:
        dispo, indispo = [], []
        for p in produits:
            if p not in NOMS:
                continue
            if (p == "Couvercle pour élément rectiligne" and charn) or (p.startswith("Kit pièces") and not charn):
                continue  # couvercle déjà intégré / kit réservé aux chemins à charnières
            g, raison = compatibles(p, r)
            (dispo if g is not None else indispo).append((p, g, raison))
        onglets.append((titre, dispo, indispo))
    tabs = st.tabs([f"{tr(RUBRIQUES_TR, titre)} ({len(d)})" for titre, d, _ in onglets])
    for tab, (titre, dispo, indispo) in zip(tabs, onglets):
        with tab:
            if not dispo:
                st.caption(t("aucun_accessoire"))
            for i in range(0, len(dispo), 3):
                cols = st.columns(3, gap="small")
                for col, (p, g, _) in zip(cols, dispo[i:i + 3]):
                    with col:
                        carte_accessoire(p, g, r)
            if indispo:
                md(f'<div class="indispo"><b>{t("non_dispo")}</b> '
                   + " · ".join(f"{nom(p)} ({raison})" for p, _, raison in indispo) + "</div>")


# ---------- Liste de matériel ----------
def panier():
    return st.session_state.setdefault("panier", {})


def ajouter(ligne, photos, qte=1):
    p = panier()
    code = ligne["Code"]
    if code in p:
        p[code]["qte"] += qte
        return
    p[code] = {
        "produit": ligne["Produit"], "abrev": ligne["Abrév"] if pd.notna(ligne["Abrév"]) else "",
        "finition": ligne["Finition"], "perfore": ligne["Perforé"], "charn": ligne["Couvercle_charnières"],
        "ep": ligne["Épaisseur"], "unite": "pcs",
        "page": int(ligne["Page_catalogue"]) if pd.notna(ligne["Page_catalogue"]) else "",
        "dims": {lab: ligne[col] for col, lab in (("H_mm", "H"), ("B_mm", "B"), ("B1_mm", "B1"), ("L_mm", "L"))
                 if pd.notna(ligne[col])},
        "photos": photos, "qte": qte,
    }


def boutons_ajout(ligne, photos, couvercle, photos_couvercle, chemin=False):
    """Boutons « Ajouter à ma liste » (et « + couvercle ») sous une carte ou une fiche."""
    ligne = ligne.to_dict()
    cles = (ligne["Code"], couvercle["Code"] if couvercle is not None else "")
    if couvercle is not None:
        c1, c2 = st.columns([1.4, 1], gap="small")
    else:
        c1, c2 = st.container(), None
    with c1:
        st.button(t("ajouter_chemin" if chemin else "ajouter"), key=f"add_{'ch' if chemin else 'acc'}_{cles[0]}",
                  on_click=ajouter, args=(ligne, photos), width="stretch", type="primary" if chemin else "secondary")
    if c2 is not None:
        with c2:
            st.button(t("ajouter_couvercle_chemin" if chemin else "ajouter_couvercle"),
                      key=f"addc_{'ch' if chemin else 'acc'}_{cles[1]}", on_click=ajouter,
                      args=(couvercle.to_dict(), photos_couvercle), width="stretch")
    deja = [panier()[c]["qte"] for c in cles if c in panier()]
    if deja:
        md(f'<div class="dans-liste">✓ {" + ".join(map(str, deja))} {t("dans_liste")}</div>')


def designation(l):
    if l["produit"] == RECT:
        txt = f"{nom(RECT)} – {fmt('Perforé', l['perfore'])}"
        if l["charn"] == "Oui":
            txt += f", {t('avec_charnieres').lower()}"
        if str(l["ep"]).startswith("1,5"):
            txt += f", {fmt('Épaisseur', l['ep'])}"
        return txt
    return nom(l["produit"]) + (f" ({l['abrev']})" if l["abrev"] else "")


def dimensions(l):
    return " × ".join(f"{k} {fmt('H_mm', v)}" for k, v in l["dims"].items()) or "—"


def maj_liste(ver):
    """Applique les quantités modifiées dans le tableau, puis recrée le tableau."""
    etat = st.session_state.get(f"ed_{ver}", {})
    codes = list(panier())
    for idx, changes in etat.get("edited_rows", {}).items():
        if "quantite" in changes and idx < len(codes):
            q = int(changes["quantite"] or 0)
            if q <= 0:
                panier().pop(codes[idx], None)
            else:
                panier()[codes[idx]]["qte"] = q
    st.session_state["ed_ver"] = ver + 1


def vider():
    st.session_state["panier"] = {}


def bloc_liste(num):
    p = panier()
    n = len(p)
    md('<div id="ma-liste"></div>')
    etape(num, t("ma_liste"), f"{n} {t('article_1') if n <= 1 else t('article_n')}")
    if not p:
        md(f'<div class="seule">{t("liste_vide")}</div>')
        return
    lignes = [{"photo": photo_src(l["photos"]), "code": c, "designation": designation(l),
               "finition": fmt("Finition", l["finition"]), "dimensions": dimensions(l),
               "quantite": l["qte"], "unite": unite(l["unite"]), "page": l["page"]} for c, l in p.items()]
    ver = st.session_state.setdefault("ed_ver", 0)
    st.caption(t("info_liste"))
    st.data_editor(
        pd.DataFrame(lignes), key=f"ed_{ver}", on_change=maj_liste, args=(ver,), hide_index=True,
        width="stretch", row_height=64,
        disabled=["photo", "code", "designation", "finition", "dimensions", "unite", "page"],
        column_config={
            "photo": st.column_config.ImageColumn(t("col_photo"), width="small"),
            "code": st.column_config.TextColumn(t("col_code"), width="small"),
            "designation": st.column_config.TextColumn(t("col_designation"), width="large"),
            "finition": st.column_config.TextColumn(t("col_finition")),
            "dimensions": st.column_config.TextColumn(t("col_dimensions")),
            "quantite": st.column_config.NumberColumn(t("col_quantite"), min_value=0, step=1, width="small"),
            "unite": st.column_config.TextColumn(t("col_unite"), width="small"),
            "page": st.column_config.TextColumn(t("col_page"), width="small"),
        })
    export = [{**lg, "photo": chemin_photo(l["photos"])} for lg, l in zip(lignes, p.values())]
    c1, c2, c3 = st.columns([1, 1, 1])
    with c1:
        st.download_button(t("telecharger_excel"), data=export_liste.excel(export, t, t("titre_doc"), LOGO),
                           file_name="DKC_liste_materiel.xlsx", width="stretch", type="primary",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with c2:
        st.download_button(t("telecharger_pdf"), data=export_liste.pdf(export, t, t("titre_doc"), LOGO),
                           file_name="DKC_liste_materiel.pdf", mime="application/pdf",
                           width="stretch", type="primary")
    with c3:
        st.button(t("vider"), on_click=vider, width="stretch")


def stepper(barre, chemin_ok):
    """Barre de progression cliquable en haut du configurateur."""
    n = len(panier())
    actif = 2 if not chemin_ok else (5 if n else 3)
    faits = {1, 2} if chemin_ok else {1}
    etapes = [(1, t("et_ligne")), (2, t("et_chemin")), (3, t("et_calcul")), (4, t("et_accessoires")),
              (5, f"{t('et_liste')} ({n})")]
    liens = []
    for num, lib in etapes:
        classe = "actif" if num == actif else ("done" if num in faits else "")
        puce = "✓" if num in faits and num != actif else num
        liens.append(f'<a class="{classe}" href="#etape-{num}"><span class="n">{puce}</span><span class="lib">{lib}</span></a>')
    barre.markdown(f'<div class="stepper">{"".join(liens)}</div>', unsafe_allow_html=True)


def configurateur():
    barre = st.empty()
    gauche, droite = st.columns([1.55, 1], gap="large")
    with gauche:
        etape(1, t("ligne"))
        ligne = st.pills("Ligne", ["S5", "S2", "S3"], default="S5", label_visibility="collapsed", key="ligne",
                         format_func=lambda x: x if x == "S5" else f"{x} · {t('bientot')}")
        if ligne not in (None, "S5"):
            stepper(barre, False)
            st.info(t("info_s2_s3"))
            return
        with st.container(key="carte_chemin"):
            etape(2, t("votre_chemin"), t("info_chemin"))
            sel = s5[s5["Produit"] == RECT]
            for col in ETAPES_CHEMIN:
                if sel[col].dropna().nunique() == 0:
                    continue
                sel, _ = choix(sel, col, "", prefixe="ch")
                if col == "Finition":
                    with st.expander(t("quelle_finition")):
                        conseil_finition()
    with droite, st.container(key="panneau"):
        panneau_chemin(sel)
    if len(sel) == 1:
        bloc_calcul(sel.iloc[0])
        bloc_accessoires(sel.iloc[0])
    bloc_liste(5)
    stepper(barre, len(sel) == 1)


def recherche_directe():
    gauche, droite = st.columns([1.55, 1], gap="large")

    # ---------- Choix (colonne de gauche) ----------
    with gauche:
        familles = [f for f in FAMILLES if f in s5["Famille"].unique()]
        if st.session_state.get("famille") not in familles:
            st.session_state["famille"] = familles[0]
        etape(1, t("type_produit"))
        grille([(f, tr(FAMILLES_TR, f), photo_src(FAMILLES[f]), ABREV_FAMILLE.get(f, "DKC"),
                 pluriel(s5[s5["Famille"] == f]["Produit"].nunique())) for f in familles],
               "famille", 4, "fam")
        famille = st.session_state["famille"]
        sel = s5[s5["Famille"] == famille]

        produits = list(dict.fromkeys(sel["Produit"]))
        if st.session_state.get("produit") not in produits:
            st.session_state["produit"] = produits[0]
        etape(2, t("produit"), pluriel(len(produits)))
        if len(produits) > 8:
            q = st.text_input("Rechercher", placeholder=t("rechercher"), label_visibility="collapsed")
            visibles = [p for p in produits if slug(q) in slug(p) or slug(q) in slug(nom(p))] if q else produits
        else:
            visibles = produits
        items = []
        for p in visibles:
            lignes = sel[sel["Produit"] == p]
            ab = abrev_de(lignes)
            items.append((p, nom(p), photo_src([p, f"famille {famille}"]), ab or ABREV_FAMILLE.get(famille, "DKC"),
                          f"{ab} · {len(lignes)} {t('ref_abrege')}" if ab else f"{len(lignes)} {t('ref_abrege')}"))
        if not items:
            st.caption(t("aucun_produit"))
        conteneur = st.container(height=520, key="grille") if len(items) > 8 else st.container()
        with conteneur:
            grille(items, "produit", 4, "prod")
        produit = st.session_state["produit"]
        sel = sel[sel["Produit"] == produit]
        abrev = abrev_de(sel)

        # 3+. Choix en cascade : chaque étape ne propose que ce qui existe vraiment
        num = 3
        choisis = {}
        for col in ETAPES:
            if sel[col].dropna().nunique() == 0:
                continue
            sel, v = choix(sel, col, num)
            choisis[col] = v
            num += 1

    # ---------- Résultat (colonne de droite, reste visible au défilement) ----------
    with droite, st.container(key="panneau"):
        if len(sel) != 1:
            md(f"""<div class="panel">{img_html(photo_src([produit, f"famille {famille}"]), "visuel", abrev or "DKC")}
                <div class="corps attente"><div class="label">{t("refs_possibles")}</div>
                <div class="gros">{len(sel)}</div><div class="nom">{nom(produit)}</div>
                <div class="muted">{t("precisez")}</div></div></div>""")
            st.dataframe(sel[["Code", "H_mm", "B_mm", "L_mm", "Finition"]], hide_index=True, height=240)
    if len(sel) != 1:
        bloc_liste(num)
        return
    with droite, st.container(key="panneau2"):

        r = sel.iloc[0]
        dims = []
        for col, lab in (("H_mm", "H"), ("B_mm", "B"), ("B1_mm", "B1"), ("L_mm", "L")):
            if pd.notna(r[col]):
                dims.append(f"{lab} {fmt(col, r[col])}")
        details = [(t("dimensions"), " × ".join(dims) if dims else "—"),
                   (t("finition"), fmt("Finition", r["Finition"]))]
        if pd.notna(r["Poids_kg_pc"]):
            details.append((t("poids"), f"{r['Poids_kg_pc']} {t('kg_piece')}"))
        if pd.notna(r["Poids_kg_m"]):
            details.append((t("poids"), f"{r['Poids_kg_m']} kg/m"))
        if pd.notna(r["TUA_cm2"]):
            details.append((t("tua"), f"{int(r['TUA_cm2'])} cm²"))
        if pd.notna(r["Conditionnement"]):
            details.append((t("conditionnement"), f"{int(r['Conditionnement'])} {unite(r['Unité_cond'])}"))
        details.append((t("page_catalogue"), str(int(r["Page_catalogue"]))))

        photos_resultat = [r["Produit"], f"famille {famille}"]  # photo de la famille en secours
        if r["Produit"] == RECT and pd.notna(r["Perforé"]):
            # Élément rectiligne : photo différente selon le fond foré ou plein, et avec charnières
            variante = "fore" if r["Perforé"].startswith("Foré") else "non fore"
            photos_resultat.insert(0, f"element rectiligne {variante}")
            if r["Couvercle_charnières"] == "Oui":
                photos_resultat.insert(0, f"element rectiligne charnieres {variante}")

        # Compléments conseillés pour un élément rectiligne sans charnières
        extras_html = ""
        if famille == "Élément rectiligne" and r.get("Couvercle_charnières") == "Non":
            cov = df[(df["Produit"] == "Couvercle pour élément rectiligne") & (df["B_mm"] == r["B_mm"])
                     & (df["L_mm"] == r["L_mm"]) & (df["Finition"] == r["Finition"])]
            jt = df[(df["Abrév"] == "GTO") & (df["H_mm"] == r["H_mm"]) & (df["Finition"] == r["Finition"])]
            extras = []
            if len(cov):
                extras.append((t("couvercle_optionnel"), cov.iloc[0]["Code"], "couvercle pour element rectiligne"))
            if len(jt):
                extras.append((t("joint_gto"), jt.iloc[0]["Code"], "joint lineaire"))
            if extras:
                extras_html = f'<div class="titre-extra">{t("a_commander")}</div>' + "".join(
                    f'<div class="extra">{img_html(photo_src(ph), "mini", "")}'
                    f'<div><div class="t">{lib}</div><div class="c">{c}</div></div></div>' for lib, c, ph in extras)

        designation = nom(r["Produit"]) + (f" ({r['Abrév']})" if pd.notna(r["Abrév"]) else "")
        chips = "".join(f'<span class="chip">{chip(c, v)}</span>' for c, v in choisis.items() if v is not None)
        rows = "".join(f"<tr><td>{k}</td><td>{v}</td></tr>" for k, v in details)
        md(f"""<div class="panel">{img_html(photo_src(photos_resultat), "visuel", abrev or "DKC")}
            <div class="corps">
              <div class="label">{t("reference_dkc")}</div>
              <div class="code">{r['Code']}</div>
              <div class="nom">{designation}</div>
              <div class="chips">{chips}</div>
              <table>{rows}</table>
              {extras_html}
            </div></div>""")
        boutons_ajout(r, photos_resultat, None, [], chemin=r["Produit"] == RECT)
    bloc_liste(num)


# ---------- En-tête ----------
logo_b64 = base64.b64encode(LOGO.read_bytes()).decode()
with st.container(key="entete"):
    c_logo, c_mode, c_liste, c_lang = st.columns([1, 2.7, 1, 1], vertical_alignment="center")
    with c_logo:
        md(f'<div class="logo-entete"><img src="data:image/png;base64,{logo_b64}" alt="DKC"></div>')
    with c_mode, st.container(key="modes"):
        mode = st.segmented_control("Mode", ["config", "recherche"], default="config", key="mode",
                                    format_func=lambda m: t("mode_config" if m == "config" else "mode_recherche"),
                                    label_visibility="collapsed")
    with c_liste:
        nb = len(panier())
        md(f'<a class="badge-liste{" plein" if nb else ""}" href="#ma-liste">🧾 {t("ma_liste_court")} <span>{nb}</span></a>')
    with c_lang, st.container(key="langue"):
        st.segmented_control("Langue", list(LANGUES), default="fr", key="lang", format_func=LANGUES.get,
                             label_visibility="collapsed")

md(f"""<div class="hero">
      <div><div class="hero-tag">{t("tag")}</div><h1>{t("titre")}</h1><p>{t("sous_titre")}</p></div>
      <div class="stats">
        <span class="stat"><b>{s5['Produit'].nunique()}</b>{t("stat_produits")}</span>
        <span class="stat"><b>{f'{len(s5):,}'.replace(',', ' ')}</b>{t("stat_refs")}</span>
        <span class="stat"><b>{s5['Finition'].nunique()}</b>{t("stat_finitions")}</span>
      </div>
    </div>""")

if mode == "recherche":
    recherche_directe()
else:
    configurateur()
