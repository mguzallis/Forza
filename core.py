#!/usr/bin/env python3
"""
Forza SDS formatter / rebrander - core engine.

extract()  : any SDS .docx  -> canonical model
render()   : canonical model + intake -> branded .docx built on a vertical shell

Design notes
------------
Nothing is copied blindly from the source document. Every paragraph and table is
re-emitted from parsed structure with Forza house formatting applied, which is
what prevents the subtle drift you get from carrying foreign XML across.

Preserved from source : text, bold/italic/underline/super/subscript, bullet and
                        numbered lists, tables (cell text + spans), inline images
Normalised            : font (Times New Roman), sizes, table style (Table Grid),
                        section heading form, spacing, hard page breaks (removed)
Regenerated           : title, top table, footer, DCN, headers/branding
"""

from __future__ import annotations

import io
import os
import re
import shutil
import urllib.error
import urllib.request
import zipfile
import json
from dataclasses import dataclass, field

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

VERTICALS = ["Composites", "Construction", "Industrial",
             "Insulation", "Marine", "Transportation"]

SECTION_TITLES = {
    1: "Identification",
    2: "Hazard Identification",
    3: "Composition/Information on Ingredients",
    4: "First-Aid Measures",
    5: "Fire-Fighting Measures",
    6: "Accidental Release Measures",
    7: "Handling and Storage",
    8: "Exposure Controls / Personal Protection",
    9: "Physical and Chemical Properties",
    10: "Stability and Reactivity",
    11: "Toxicological Information",
    12: "Ecological Information",
    13: "Disposal Considerations",
    14: "Transport Information",
    15: "Regulatory Information",
    16: "Other Information",
}

# Standard GHS/WHMIS 16-section titles. These are internationally
# standardized (the 16-section SDS format itself is harmonized under GHS),
# not a Forza-specific translation choice.
SECTION_TITLES_FR = {
    1: "Identification",
    2: "Identification des dangers",
    3: "Composition/informations sur les composants",
    4: "Premiers secours",
    5: "Mesures de lutte contre l'incendie",
    6: "Mesures \u00e0 prendre en cas de dispersion accidentelle",
    7: "Manutention et stockage",
    8: "Contr\u00f4le de l'exposition/protection individuelle",
    9: "Propri\u00e9t\u00e9s physiques et chimiques",
    10: "Stabilit\u00e9 et r\u00e9activit\u00e9",
    11: "Donn\u00e9es toxicologiques",
    12: "Donn\u00e9es \u00e9cologiques",
    13: "Consid\u00e9rations relatives \u00e0 l'\u00e9limination",
    14: "Informations relatives au transport",
    15: "Informations r\u00e9glementaires",
    16: "Autres informations",
}

SECTION_TITLES_ES = {
    1: "Identificaci\u00f3n",
    2: "Identificaci\u00f3n de peligros",
    3: "Composici\u00f3n/informaci\u00f3n sobre los componentes",
    4: "Primeros auxilios",
    5: "Medidas de lucha contra incendios",
    6: "Medidas en caso de vertido accidental",
    7: "Manipulaci\u00f3n y almacenamiento",
    8: "Controles de exposici\u00f3n/protecci\u00f3n individual",
    9: "Propiedades f\u00edsicas y qu\u00edmicas",
    10: "Estabilidad y reactividad",
    11: "Informaci\u00f3n toxicol\u00f3gica",
    12: "Informaci\u00f3n ecol\u00f3gica",
    13: "Consideraciones relativas a la eliminaci\u00f3n",
    14: "Informaci\u00f3n relativa al transporte",
    15: "Informaci\u00f3n reglamentaria",
    16: "Otra informaci\u00f3n",
}

SECTION_TITLES_BY_LANG = {"en": SECTION_TITLES, "fr": SECTION_TITLES_FR, "es": SECTION_TITLES_ES}

# Fixed template labels (doc title, top table, Section 1 field names, footer).
# fr = Canadian French. Canada's Hazardous Products Regulations (SOR/2015-17)
# incorporates the UN GHS Annex 3 standard phrasing by reference for hazard/
# precautionary statements, rather than defining separate Canada-only
# wording - so "Canadian French" and "standard GHS French" are the same text
# for that content (see HAZARD_PHRASES_FR / PRECAUTION_PHRASES_FR below).
LANG = {
    "en": {
        "doc_title": "SAFETY DATA SHEET",
        "trade_name": "Trade Name",
        "sds_num": "SDS #",
        "date_issue": "Date of Issue",
        "replaces": "Replaces",
        "effective": "Effective Date",
        "page": "Page",
        "of": "of",
        "dcn": "DCN:",
        "product_name": "Product Name:",
        "other_means": "Other Means of Identification:",
        "product_code": "Product Code Number:",
        "recommended_use": "Recommended Use:",
        "recommended_restrictions": "Recommended Restrictions:",
        "suppliers_details": "Suppliers Details",
        "company": "Company:",
        "company_phone": "Company Phone Number:",
        "emergency_phone": "Emergency Phone Number:",
        "hazard_pictograms": "Hazard Pictograms:",
        "unverified": "",
    },
    "fr": {
        "doc_title": "FICHE DE DONN\u00c9ES DE S\u00c9CURIT\u00c9",
        "trade_name": "Nom commercial",
        "sds_num": "N\u00b0 de FDS",
        "date_issue": "Date d'\u00e9mission",
        "replaces": "Remplace",
        "effective": "Date d'entr\u00e9e en vigueur",
        "page": "Page",
        "of": "de",
        "dcn": "NCD :",
        "product_name": "Nom du produit :",
        "other_means": "Autres moyens d'identification :",
        "product_code": "Num\u00e9ro de code du produit :",
        "recommended_use": "Usage recommand\u00e9 :",
        "recommended_restrictions": "Restrictions d'usage recommand\u00e9es :",
        "suppliers_details": "Coordonn\u00e9es du fournisseur",
        "company": "Entreprise :",
        "company_phone": "T\u00e9l\u00e9phone de l'entreprise :",
        "emergency_phone": "T\u00e9l\u00e9phone d'urgence :",
        "hazard_pictograms": "Pictogrammes de danger :",
        "unverified": " [FR: \u00e0 v\u00e9rifier]",
    },
    "es": {
        "doc_title": "HOJA DE DATOS DE SEGURIDAD",
        "trade_name": "Nombre comercial",
        "sds_num": "N.\u00ba de HDS",
        "date_issue": "Fecha de emisi\u00f3n",
        "replaces": "Reemplaza a",
        "effective": "Fecha de vigencia",
        "page": "P\u00e1gina",
        "of": "de",
        "dcn": "NCD:",
        "product_name": "Nombre del producto:",
        "other_means": "Otros medios de identificaci\u00f3n:",
        "product_code": "N\u00famero de c\u00f3digo del producto:",
        "recommended_use": "Uso recomendado:",
        "recommended_restrictions": "Restricciones de uso recomendadas:",
        "suppliers_details": "Datos del proveedor",
        "company": "Empresa:",
        "company_phone": "Tel\u00e9fono de la empresa:",
        "emergency_phone": "Tel\u00e9fono de emergencia:",
        "hazard_pictograms": "Pictogramas de peligro:",
        "unverified": " [ES: por verificar]",
    },
}

# Curated H-code / P-code -> official GHS phrase lookup. This is NOT a
# complete transcription of every possible GHS code/combination - it covers
# the codes that actually appear across Forza's real product SDSs. A code
# not in this table is left in English with the "unverified" marker above
# rather than guessed, so nothing looks officially translated when it
# hasn't been checked. Expand by adding entries; never invent a phrase here.
HAZARD_PHRASES_FR = {
    "H224": "Liquide et vapeurs extr\u00eamement inflammables.",
    "H225": "Liquide et vapeurs tr\u00e8s inflammables.",
    "H226": "Liquide et vapeurs inflammables.",
    "H227": "Liquide combustible.",
    "H228": "Solide inflammable.",
    "H280": "Contient un gaz sous pression; peut exploser sous l'effet de la chaleur.",
    "H281": "Contient un gaz r\u00e9frig\u00e9r\u00e9; peut causer des br\u00fblures ou blessures cryog\u00e9niques.",
    "H282": "Produit chimique sous pression : peut exploser sous l'effet de la chaleur.",
    "H283": "Produit chimique sous pression : peut exploser sous l'effet de la chaleur.",
    "H284": "Produit chimique sous pression : peut exploser sous l'effet de la chaleur.",
    "H290": "Peut \u00eatre corrosif pour les m\u00e9taux.",
    "H302": "Nocif en cas d'ingestion.",
    "H303": "Peut \u00eatre nocif en cas d'ingestion.",
    "H304": "Peut \u00eatre mortel en cas d'ingestion et de p\u00e9n\u00e9tration dans les voies respiratoires.",
    "H305": "Peut \u00eatre nocif en cas d'ingestion et de p\u00e9n\u00e9tration dans les voies respiratoires.",
    "H312": "Nocif par contact cutan\u00e9.",
    "H314": "Provoque des br\u00fblures de la peau et des l\u00e9sions oculaires graves.",
    "H315": "Provoque une irritation cutan\u00e9e.",
    "H316": "Provoque une l\u00e9g\u00e8re irritation cutan\u00e9e.",
    "H317": "Peut provoquer une allergie cutan\u00e9e.",
    "H318": "Provoque des l\u00e9sions oculaires graves.",
    "H319": "Provoque une s\u00e9v\u00e8re irritation des yeux.",
    "H320": "Provoque une irritation oculaire.",
    "H332": "Nocif par inhalation.",
    "H334": "Peut provoquer des sympt\u00f4mes allergiques ou d'asthme ou des difficult\u00e9s respiratoires par inhalation.",
    "H335": "Peut irriter les voies respiratoires.",
    "H336": "Peut provoquer somnolence ou vertiges.",
    "H351": "Susceptible de provoquer le cancer.",
    "H360": "Peut nuire \u00e0 la fertilit\u00e9 ou au f\u0153tus.",
    "H361": "Susceptible de nuire \u00e0 la fertilit\u00e9 ou au f\u0153tus.",
    "H362": "Peut \u00eatre nocif pour les b\u00e9b\u00e9s nourris au lait maternel.",
    "H371": "Risque pr\u00e9sum\u00e9 d'effets graves pour les organes.",
    "H373": "Risque pr\u00e9sum\u00e9 d'effets graves pour les organes \u00e0 la suite d'expositions r\u00e9p\u00e9t\u00e9es ou d'une exposition prolong\u00e9e.",
    "H400": "Tr\u00e8s toxique pour les organismes aquatiques.",
    "H410": "Tr\u00e8s toxique pour les organismes aquatiques, entra\u00eene des effets n\u00e9fastes \u00e0 long terme.",
    "H411": "Toxique pour les organismes aquatiques, entra\u00eene des effets n\u00e9fastes \u00e0 long terme.",
    "H412": "Nocif pour les organismes aquatiques, entra\u00eene des effets n\u00e9fastes \u00e0 long terme.",
    "H413": "Peut \u00eatre nocif \u00e0 long terme pour les organismes aquatiques.",
    "H301+H311+H331": "Toxique en cas d'ingestion, par contact cutan\u00e9 ou par inhalation.",
    "H302+H312": "Nocif en cas d'ingestion ou par contact cutan\u00e9.",
    "H302+H332": "Nocif en cas d'ingestion ou par inhalation.",
    "H312+H332": "Nocif par contact cutan\u00e9 ou par inhalation.",
    "H302+H312+H332": "Nocif en cas d'ingestion, par contact cutan\u00e9 ou par inhalation.",
    "H315+H319": "Provoque une irritation cutan\u00e9e et une s\u00e9v\u00e8re irritation des yeux.",
    "H315+H320": "Provoque une irritation cutan\u00e9e et oculaire.",
}

PRECAUTION_PHRASES_FR = {
    "P201": "Se procurer les instructions avant utilisation.",
    "P202": "Ne pas manipuler avant d'avoir lu et compris toutes les pr\u00e9cautions de s\u00e9curit\u00e9.",
    "P210": "Tenir \u00e0 l'\u00e9cart de la chaleur, des surfaces chaudes, des \u00e9tincelles, des flammes nues et de toute autre source d'inflammation. Ne pas fumer.",
    "P211": "Ne pas vaporiser sur une flamme nue ou sur toute autre source d'inflammation.",
    "P233": "Maintenir le r\u00e9cipient ferm\u00e9 de mani\u00e8re \u00e9tanche.",
    "P240": "Mise \u00e0 la terre/liaison \u00e9quipotentielle du r\u00e9cipient et du mat\u00e9riel de r\u00e9ception.",
    "P241": "Utiliser du mat\u00e9riel \u00e9lectrique/de ventilation/d'\u00e9clairage antid\u00e9flagrant.",
    "P242": "Ne pas utiliser d'outils produisant des \u00e9tincelles.",
    "P243": "Prendre des mesures de pr\u00e9caution contre les d\u00e9charges \u00e9lectrostatiques.",
    "P251": "Ne pas perforer, ni br\u00fbler, m\u00eame apr\u00e8s usage.",
    "P260": "Ne pas respirer les poussi\u00e8res/fum\u00e9es/gaz/brouillards/vapeurs/a\u00e9rosols.",
    "P261": "\u00c9viter de respirer les poussi\u00e8res/fum\u00e9es/gaz/brouillards/vapeurs/a\u00e9rosols.",
    "P264": "Se laver soigneusement apr\u00e8s manipulation.",
    "P271": "Utiliser seulement en plein air ou dans un endroit bien ventil\u00e9.",
    "P272": "Les v\u00eatements de travail contamin\u00e9s ne devraient pas sortir du lieu de travail.",
    "P273": "\u00c9viter le rejet dans l'environnement.",
    "P280": "Porter des gants de protection/des v\u00eatements de protection/un \u00e9quipement de protection des yeux/du visage.",
    "P301+P310": "EN CAS D'INGESTION : Appeler imm\u00e9diatement un CENTRE ANTIPOISON ou un m\u00e9decin.",
    "P301+P312": "EN CAS D'INGESTION : Appeler un CENTRE ANTIPOISON ou un m\u00e9decin en cas de malaise.",
    "P302+P352": "EN CAS DE CONTACT AVEC LA PEAU : Laver abondamment \u00e0 l'eau et au savon.",
    "P303+P361+P353": "EN CAS DE CONTACT AVEC LA PEAU (ou les cheveux) : enlever imm\u00e9diatement tous les v\u00eatements contamin\u00e9s. Rincer la peau \u00e0 l'eau/se doucher.",
    "P304+P340": "EN CAS D'INHALATION : transporter la personne \u00e0 l'ext\u00e9rieur et la maintenir au repos dans une position o\u00f9 elle peut confortablement respirer.",
    "P304+P312": "EN CAS D'INHALATION : Appeler un CENTRE ANTIPOISON ou un m\u00e9decin en cas de malaise.",
    "P305+P351+P338": "EN CAS DE CONTACT AVEC LES YEUX : rincer avec pr\u00e9caution \u00e0 l'eau pendant plusieurs minutes. Enlever les lentilles de contact si la victime en porte et si elles peuvent \u00eatre facilement enlev\u00e9es. Continuer \u00e0 rincer.",
    "P306+P360": "EN CAS DE CONTACT AVEC LES V\u00caTEMENTS : rincer imm\u00e9diatement les v\u00eatements et la peau contamin\u00e9s abondamment \u00e0 l'eau avant de les enlever.",
    "P310": "Appeler imm\u00e9diatement un CENTRE ANTIPOISON ou un m\u00e9decin.",
    "P311": "Appeler un CENTRE ANTIPOISON ou un m\u00e9decin.",
    "P312": "Appeler un CENTRE ANTIPOISON ou un m\u00e9decin en cas de malaise.",
    "P313": "Consulter un m\u00e9decin.",
    "P330": "Rincer la bouche.",
    "P331": "NE PAS faire vomir.",
    "P332+P313": "En cas d'irritation cutan\u00e9e : consulter un m\u00e9decin.",
    "P333+P313": "En cas d'irritation ou d'\u00e9ruption cutan\u00e9e : consulter un m\u00e9decin.",
    "P337+P313": "Si l'irritation oculaire persiste : consulter un m\u00e9decin.",
    "P340": "Transporter la personne \u00e0 l'ext\u00e9rieur et la maintenir au repos dans une position o\u00f9 elle peut confortablement respirer.",
    "P362+P364": "Enlever les v\u00eatements contamin\u00e9s et les laver avant r\u00e9utilisation.",
    "P370+P378": "En cas d'incendie : utiliser un moyen d'extinction appropri\u00e9.",
    "P371+P380+P375": "En cas d'incendie important et de grandes quantit\u00e9s : \u00e9vacuer la zone. En raison d'un risque d'explosion, combattre l'incendie \u00e0 distance.",
    "P403+P233": "Stocker dans un endroit bien ventil\u00e9. Maintenir le r\u00e9cipient ferm\u00e9 de mani\u00e8re \u00e9tanche.",
    "P403+P235": "Stocker dans un endroit bien ventil\u00e9. Tenir au frais.",
    "P405": "Garder sous clef.",
    "P501": "\u00c9liminer le contenu/r\u00e9cipient conform\u00e9ment \u00e0 la r\u00e9glementation locale.",
}

HAZARD_PHRASES_ES = {
    "H224": "L\u00edquido y vapores extremadamente inflamables.",
    "H225": "L\u00edquido y vapores muy inflamables.",
    "H226": "L\u00edquidos y vapores inflamables.",
    "H227": "L\u00edquido combustible.",
    "H280": "Contiene gas a presi\u00f3n; peligro de explosi\u00f3n si se calienta.",
    "H281": "Contiene gas refrigerado; puede provocar quemaduras o lesiones criog\u00e9nicas.",
    "H282": "Producto qu\u00edmico a presi\u00f3n: peligro de explosi\u00f3n si se calienta.",
    "H283": "Producto qu\u00edmico a presi\u00f3n: peligro de explosi\u00f3n si se calienta.",
    "H284": "Producto qu\u00edmico a presi\u00f3n: peligro de explosi\u00f3n si se calienta.",
    "H290": "Puede ser corrosivo para los metales.",
    "H302": "Nocivo en caso de ingesti\u00f3n.",
    "H303": "Puede ser nocivo en caso de ingesti\u00f3n.",
    "H304": "Puede ser mortal en caso de ingesti\u00f3n y penetraci\u00f3n en las v\u00edas respiratorias.",
    "H305": "Puede ser nocivo en caso de ingesti\u00f3n y penetraci\u00f3n en las v\u00edas respiratorias.",
    "H312": "Nocivo en contacto con la piel.",
    "H314": "Provoca quemaduras graves en la piel y da\u00f1os oculares graves.",
    "H315": "Provoca irritaci\u00f3n cut\u00e1nea.",
    "H316": "Provoca una leve irritaci\u00f3n cut\u00e1nea.",
    "H317": "Puede provocar una reacci\u00f3n al\u00e9rgica en la piel.",
    "H318": "Provoca da\u00f1os oculares graves.",
    "H319": "Provoca irritaci\u00f3n ocular grave.",
    "H320": "Provoca irritaci\u00f3n ocular.",
    "H332": "Nocivo en caso de inhalaci\u00f3n.",
    "H334": "Puede provocar s\u00edntomas de alergia o asma o dificultades respiratorias en caso de inhalaci\u00f3n.",
    "H335": "Puede irritar las v\u00edas respiratorias.",
    "H336": "Puede provocar somnolencia o v\u00e9rtigo.",
    "H351": "Se sospecha que provoca c\u00e1ncer.",
    "H360": "Puede perjudicar la fertilidad o al feto.",
    "H361": "Se sospecha que perjudica la fertilidad o al feto.",
    "H362": "Puede ser nocivo para los ni\u00f1os alimentados con leche materna.",
    "H371": "Puede provocar da\u00f1os en los \u00f3rganos.",
    "H373": "Puede provocar da\u00f1os en los \u00f3rganos tras exposiciones prolongadas o repetidas.",
    "H400": "Muy t\u00f3xico para los organismos acu\u00e1ticos.",
    "H410": "Muy t\u00f3xico para los organismos acu\u00e1ticos, con efectos nocivos duraderos.",
    "H411": "T\u00f3xico para los organismos acu\u00e1ticos, con efectos nocivos duraderos.",
    "H412": "Nocivo para los organismos acu\u00e1ticos, con efectos nocivos duraderos.",
    "H413": "Puede ser nocivo para los organismos acu\u00e1ticos, con efectos nocivos duraderos.",
    "H301+H311+H331": "T\u00f3xico en caso de ingesti\u00f3n, contacto con la piel o inhalaci\u00f3n.",
    "H302+H312": "Nocivo en caso de ingesti\u00f3n o contacto con la piel.",
    "H302+H332": "Nocivo en caso de ingesti\u00f3n o inhalaci\u00f3n.",
    "H312+H332": "Nocivo en contacto con la piel o inhalaci\u00f3n.",
    "H302+H312+H332": "Nocivo en caso de ingesti\u00f3n, contacto con la piel o inhalaci\u00f3n.",
    "H315+H319": "Provoca irritaci\u00f3n cut\u00e1nea e irritaci\u00f3n ocular grave.",
}

PRECAUTION_PHRASES_ES = {
    "P201": "Solicitar instrucciones especiales antes del uso.",
    "P202": "No manipular la sustancia antes de haber le\u00eddo y comprendido todas las instrucciones de seguridad.",
    "P210": "Mantener alejado de fuentes de calor, superficies calientes, chispas, llamas abiertas y otras fuentes de ignici\u00f3n. No fumar.",
    "P211": "No pulverizar sobre una llama abierta u otra fuente de ignici\u00f3n.",
    "P233": "Mantener el recipiente cerrado herm\u00e9ticamente.",
    "P240": "Conexi\u00f3n a tierra/enlace equipotencial del recipiente y del equipo de recepci\u00f3n.",
    "P241": "Utilizar un equipo el\u00e9ctrico/de ventilaci\u00f3n/de iluminaci\u00f3n antideflagrante.",
    "P242": "Utilizar \u00fanicamente herramientas que no produzcan chispas.",
    "P243": "Tomar medidas de precauci\u00f3n contra descargas electrost\u00e1ticas.",
    "P251": "No perforar ni quemar, incluso despu\u00e9s de su uso.",
    "P260": "No respirar el polvo/el humo/el gas/la niebla/los vapores/el aerosol.",
    "P261": "Evitar respirar el polvo/el humo/el gas/la niebla/los vapores/el aerosol.",
    "P264": "Lavarse concienzudamente despu\u00e9s de la manipulaci\u00f3n.",
    "P271": "Usar \u00fanicamente en exteriores o en un lugar bien ventilado.",
    "P273": "Evitar su liberaci\u00f3n al medio ambiente.",
    "P280": "Llevar guantes/prendas/gafas/m\u00e1scara de protecci\u00f3n.",
    "P301+P310": "EN CASO DE INGESTI\u00d3N: Llamar inmediatamente a un CENTRO DE TOXICOLOG\u00cdA/m\u00e9dico.",
    "P301+P312": "EN CASO DE INGESTI\u00d3N: Llamar a un CENTRO DE TOXICOLOG\u00cdA/m\u00e9dico si la persona se encuentra mal.",
    "P302+P352": "EN CASO DE CONTACTO CON LA PIEL: Lavar con abundante agua y jab\u00f3n.",
    "P304+P340": "EN CASO DE INHALACI\u00d3N: Transportar a la persona al aire libre y mantenerla en una posici\u00f3n que le facilite la respiraci\u00f3n.",
    "P304+P312": "EN CASO DE INHALACI\u00d3N: Llamar a un CENTRO DE TOXICOLOG\u00cdA/m\u00e9dico si la persona se encuentra mal.",
    "P305+P351+P338": "EN CASO DE CONTACTO CON LOS OJOS: Aclarar cuidadosamente con agua durante varios minutos. Quitar las lentes de contacto, si lleva y resulta f\u00e1cil. Seguir aclarando.",
    "P310": "Llamar inmediatamente a un CENTRO DE TOXICOLOG\u00cdA/m\u00e9dico.",
    "P311": "Llamar a un CENTRO DE TOXICOLOG\u00cdA/m\u00e9dico.",
    "P312": "Llamar a un CENTRO DE TOXICOLOG\u00cdA/m\u00e9dico si la persona se encuentra mal.",
    "P313": "Consultar a un m\u00e9dico.",
    "P330": "Enjuagarse la boca.",
    "P331": "NO provocar el v\u00f3mito.",
    "P332+P313": "En caso de irritaci\u00f3n cut\u00e1nea: Consultar a un m\u00e9dico.",
    "P333+P313": "En caso de irritaci\u00f3n o erupci\u00f3n cut\u00e1nea: Consultar a un m\u00e9dico.",
    "P337+P313": "Si persiste la irritaci\u00f3n ocular: Consultar a un m\u00e9dico.",
    "P340": "Transportar a la persona al aire libre y mantenerla en una posici\u00f3n que le facilite la respiraci\u00f3n.",
    "P362+P364": "Quitar la ropa contaminada y lavarla antes de volver a usarla.",
    "P370+P378": "En caso de incendio: Utilizar los medios de extinci\u00f3n adecuados.",
    "P403+P233": "Almacenar en un lugar bien ventilado. Mantener el recipiente cerrado herm\u00e9ticamente.",
    "P405": "Guardar bajo llave.",
    "P501": "Eliminar el contenido/el recipiente conforme a la normativa local.",
}

CODE_PHRASES_BY_LANG = {
    "fr": {**HAZARD_PHRASES_FR, **PRECAUTION_PHRASES_FR},
    "es": {**HAZARD_PHRASES_ES, **PRECAUTION_PHRASES_ES},
}

# Matches a leading H-code or P-code (with optional "+combo") at the start
# of a hazard/precautionary statement line, e.g. "H282 Extremely..." or
# "P301+P310 IF SWALLOWED...".
_CODE_LINE_RE = re.compile(r"^\s*((?:H|P)\d{3}(?:\s*\+\s*(?:H|P)\d{3})*)\b\s*(.*)$")


def _split_lines(runs):
    """Split a run list into lines wherever a run's text contains an
    embedded \\x00 (a manual line break within one Word paragraph). Returns
    a list of run-lists, one per line."""
    lines = [[]]
    for r in runs:
        if r.image or "\x00" not in r.text:
            lines[-1].append(r)
            continue
        parts = r.text.split("\x00")
        for i, part in enumerate(parts):
            if i:
                lines.append([])
            if part or i == 0:
                lines[-1].append(Run(part, r.bold, r.italic, r.underline, r.vert))
    return lines


def _join_lines(lines):
    """Inverse of _split_lines: rejoin line-groups into one run list."""
    combined = []
    for idx, line_runs in enumerate(lines):
        if not line_runs:
            line_runs = [Run("")]
        if idx > 0:
            first = line_runs[0]
            line_runs = [Run("\x00" + first.text, first.bold, first.italic,
                             first.underline, first.vert)] + line_runs[1:]
        combined.extend(line_runs)
    return combined


def _is_fixed_s16_line(line_text, ver, date_of_issue):
    """True for a Section 16 line already fully handled by apply_intake()
    (a 'Prepared by' line, now a fixed company name, or an SDS date/revision
    line, now a formatted date). Neither is prose, and the label itself is
    translated separately via LANG - so a line like this must never be sent
    to DeepL at all.
    """
    m = re.match(r"([A-Za-z][A-Za-z /]+?)\s*:\s*", line_text.strip())
    if not m:
        return False
    label = m.group(1).strip()
    if PREPARED_BY_RE.search(label):
        return True
    return _s16_route(label, ver, date_of_issue) is not None


def translate_code_line(text: str, lang: str):
    """If a line starts with a recognized H-code/P-code, return the verified
    GHS phrase for that exact code (keeping the code prefix). Returns None
    when the code isn't recognized, or the line has no code prefix at all -
    in both cases the caller falls through to full-text translation (DeepL)
    instead of guessing at regulated phrasing itself.
    """
    if lang == "en":
        return None
    m = _CODE_LINE_RE.match(text)
    if not m:
        return None
    code = m.group(1).replace(" ", "")
    phrase = CODE_PHRASES_BY_LANG.get(lang, {}).get(code)
    if phrase:
        return f"{code} {phrase}"
    return None


class DeepLError(Exception):
    """Raised when a DeepL API call fails - bad/missing key, quota exceeded,
    or a network problem. Callers should surface this clearly rather than
    silently falling back, since a silent fallback would hide a real problem
    from whoever is generating the document.
    """


_DEEPL_TARGET = {"fr": "FR", "es": "ES"}

# A handful of very short conversational patterns a translation backend
# should never produce for ordinary source text (one showed up for a bare
# product code sent with no surrounding context). A response matching one
# of these is treated as a failed translation for that string, and the
# original English is kept instead - failing safe rather than shipping a
# visibly wrong sentence into a real business document.
_SUSPICIOUS_RESPONSE_RE = re.compile(
    r"^(lo siento|i'?m sorry|je ne peux pas|desculpe|es tut mir leid)\b", re.I)


def _looks_like_bare_code(value: str) -> bool:
    """True for a short, space-free, code-like identifier (e.g. 'IC947',
    'IC947-22L') as opposed to genuine descriptive prose (e.g. 'TAC850 Web
    Spray Tackifier'). Bare codes are never sent to DeepL - there's nothing
    to translate, and a very short context-free string is exactly what
    triggered a nonsense response in testing.
    """
    v = (value or "").strip()
    return bool(v) and " " not in v and len(v) <= 20


def deepl_translate_batch(texts: list, target_lang: str, api_key: str) -> list:
    """Translate a batch of strings via the DeepL API (free or pro tier -
    detected from the ':fx' suffix on the key). Returns translations in the
    same order as the input. Raises DeepLError with a clear message on any
    failure; never returns a partial/guessed result. Any individual response
    that looks like a conversational non-answer rather than a translation
    (see _SUSPICIOUS_RESPONSE_RE) is replaced with the original source text
    for that entry, so a backend glitch can't silently insert nonsense.
    """
    if not texts:
        return []
    if not api_key:
        raise DeepLError("No DeepL API key configured.")
    host = "api-free.deepl.com" if api_key.strip().endswith(":fx") else "api.deepl.com"
    target = _DEEPL_TARGET.get(target_lang.lower(), target_lang.upper())
    out = []
    # DeepL accepts many strings per call; chunk conservatively so one very
    # long document never risks hitting a request-size limit.
    for i in range(0, len(texts), 50):
        chunk = texts[i:i + 50]
        body = json.dumps({
            "text": chunk,
            "target_lang": target,
            # The classic, deterministic engine - less prone to the kind of
            # conversational hallucination an LLM-based backend produced for
            # a bare, context-free product code during testing.
            "model_type": "latency_optimized",
        }).encode("utf-8")
        req = urllib.request.Request(
            f"https://{host}/v2/translate",
            data=body,
            method="POST",
            headers={
                "Authorization": f"DeepL-Auth-Key {api_key.strip()}",
                "Content-Type": "application/json",
                "User-Agent": "ForzaSDSFormatter/1.0",
            },
        )
        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                data = json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            detail = e.read().decode("utf-8", errors="replace")
            if e.code == 403:
                raise DeepLError("DeepL rejected the API key (403 Forbidden). "
                                 "Check that the key was copied correctly.") from e
            if e.code == 456:
                raise DeepLError("DeepL quota exceeded for this API key "
                                 "(free tier: 500,000 characters/month).") from e
            raise DeepLError(f"DeepL API error {e.code}: {detail[:200]}") from e
        except urllib.error.URLError as e:
            raise DeepLError(f"Could not reach DeepL: {e.reason}") from e
        translations = [t["text"] for t in data.get("translations", [])]
        for src, tr in zip(chunk, translations):
            out.append(src if _SUSPICIOUS_RESPONSE_RE.match(tr.strip()) else tr)
    return out


_SECTION1_FREE_FIELDS = ("product_name", "other_means", "recommended_use",
                        "recommended_restrictions")


def build_translation_map(sds: "SDS", intake: dict, lang: str, api_key: str) -> dict:
    """Collect every piece of free (non-code, non-label) English text in the
    document - the Section 1 identity fields plus every source-derived
    paragraph/cell in Sections 2-16 that ISN'T a recognized hazard/
    precautionary code line, a bare product code, or an already-fixed
    Section 16 label line - and translate it in one batched DeepL call.
    Returns {original_text: translated_text}. Returns an empty dict (meaning:
    leave that text in English) if lang is 'en' or no api_key is configured,
    so the tool degrades gracefully rather than failing when DeepL isn't set up.
    """
    if lang == "en" or not api_key:
        return {}

    ver = re.sub(r"\D", "", str(intake.get("version", "1"))) or "1"
    date_of_issue = normalize_date_string(intake.get("date_of_issue", ""))

    seen = []
    seen_set = set()

    def add(t):
        t = (t or "").strip()
        if t and t not in seen_set:
            seen_set.add(t)
            seen.append(t)

    for field in _SECTION1_FREE_FIELDS:
        val = intake.get(field, "")
        if not _looks_like_bare_code(val):
            add(val)

    def walk_para_text(p):
        # A Section 16 paragraph may hold several lines glued together via
        # a manual line break (e.g. a date line, or "Prepared by" plus a
        # trailing disclaimer sharing one paragraph) - each line is checked
        # independently rather than sending the whole blob as one string.
        for line in p.text.split("\x00"):
            if not line.strip():
                continue
            if translate_code_line(line, lang) is not None:
                continue
            if _is_fixed_s16_line(line, ver, date_of_issue):
                continue
            m = _CODE_LINE_RE.match(line)
            add(m.group(2) if m else line)

    for sec in sds.sections:
        if sec.number == 1:
            continue
        for blk in sec.blocks:
            if isinstance(blk, Table):
                for row in blk.rows:
                    for cell in row:
                        for p in cell.paras:
                            walk_para_text(p)
            elif blk.runs:
                walk_para_text(blk)

    if not seen:

        return {}

    translated = deepl_translate_batch(seen, lang, api_key)
    return dict(zip(seen, translated))


def _translate_line(line_runs, lang, text_map):
    """Translate one line (a run-list with no embedded \\x00) from a
    paragraph or table cell. Returns a new run-list, or the same object
    unchanged if there's nothing to translate for this line.
    """
    line_text = "".join(r.text for r in line_runs if not r.image)
    if not line_text.strip():
        return line_runs
    verified = translate_code_line(line_text, lang)
    if verified is not None:
        m = _CODE_LINE_RE.match(line_text)
        if line_runs and line_runs[0].bold and line_runs[0].text.strip() == m.group(1).strip():
            rest = verified[len(m.group(1).replace(" ", "")):]
            return [Run(line_runs[0].text, bold=True), Run(rest)]
        return [Run(verified)]

    m = _CODE_LINE_RE.match(line_text)
    lookup_key = (m.group(2) if m else line_text).strip()
    translated = text_map.get(lookup_key)
    if translated and translated != lookup_key:
        if m and line_runs and line_runs[0].bold and line_runs[0].text.strip() == m.group(1).strip():
            # Unrecognized code: keep the code bolded exactly as the source
            # had it, translate only the remainder.
            return [Run(line_runs[0].text, bold=True), Run(" " + translated)]
        if m:
            return [Run(f"{m.group(1)} {translated}")]
        return [Run(translated)]
    return line_runs


def _translate_block(blk, lang, text_map=None):
    """Translate a source-derived block for a non-English output, line by
    line - a single Word paragraph can hold several lines glued together
    via a manual line break (e.g. a Section 16 date line, or 'Prepared by'
    sharing one paragraph with a trailing disclaimer), and each needs its
    own independent decision rather than being treated as one blob. A line
    starting with a recognized H-code/P-code gets the verified GHS phrase,
    bold code prefix preserved. Everything else is looked up in text_map
    (the DeepL-translated batch built by build_translation_map) - if it's
    not there (no API key configured, already fixed elsewhere - a Section
    16 date or 'Prepared by' line - or nothing came back for this exact
    string), the line passes through completely unchanged. Nothing here
    ever invents a translation on its own.
    """
    if lang == "en":
        return blk
    text_map = text_map or {}
    if isinstance(blk, Table):
        changed = False
        new_rows = []
        for row in blk.rows:
            new_row = []
            for cell in row:
                new_paras = []
                for p in cell.paras:
                    new_p = _translate_block(p, lang, text_map)
                    if new_p is not p:
                        changed = True
                    new_paras.append(new_p)
                new_row.append(Cell(new_paras, span=cell.span, width=cell.width))
            new_rows.append(new_row)
        return Table(grid=blk.grid, rows=new_rows) if changed else blk

    if not blk.text.strip():
        return blk

    lines = _split_lines(blk.runs)
    changed = False
    new_lines = []
    for line_runs in lines:
        new_line = _translate_line(line_runs, lang, text_map)
        if new_line is not line_runs:
            changed = True
        new_lines.append(new_line)

    if not changed:
        return blk
    return Para(_join_lines(new_lines), list_kind=blk.list_kind, level=blk.level, spacing_after=blk.spacing_after)


# Any dash, any spacing, optional colon: "SECTION 1 - Identification", "Section 1:", "SECTION 1"
SECTION_RE = re.compile(r"^\s*SECTION\s+(\d{1,2})\s*[-\u2010-\u2015:.]?\s*(.*)$", re.I)

TNR = ('<w:rFonts w:ascii="Times New Roman" w:eastAsia="Times New Roman" '
       'w:hAnsi="Times New Roman" w:cs="Times New Roman"/>')


# ----------------------------------------------------------------- model ---
@dataclass
class Run:
    text: str = ""
    bold: bool = False
    italic: bool = False
    underline: bool = False
    vert: str = ""            # "superscript" | "subscript" | ""
    image: str | None = None  # media filename once staged
    cx: int = 0
    cy: int = 0


@dataclass
class Para:
    runs: list = field(default_factory=list)
    list_kind: str = ""       # "bullet" | "number" | ""
    level: int = 0
    spacing_after: int | None = None   # dxa/20 units; None -> caller default

    @property
    def text(self):
        return "".join(r.text for r in self.runs)


@dataclass
class Table:
    grid: list = field(default_factory=list)   # column widths, dxa
    rows: list = field(default_factory=list)   # list[list[Cell]]


@dataclass
class Cell:
    paras: list = field(default_factory=list)
    span: int = 1
    width: int = 0


@dataclass
class Section:
    number: int
    title: str
    blocks: list = field(default_factory=list)  # Para | Table


@dataclass
class SDS:
    sections: list = field(default_factory=list)
    preamble: list = field(default_factory=list)
    media: dict = field(default_factory=dict)   # filename -> bytes
    detected: dict = field(default_factory=dict)


# ------------------------------------------------------------- extraction ---
def _xml_unescape(s):
    return (s.replace("&lt;", "<").replace("&gt;", ">")
             .replace("&quot;", '"').replace("&apos;", "'")
             .replace("&amp;", "&"))


def _xml_escape(s):
    return (s.replace("&", "&amp;").replace("<", "&lt;")
             .replace(">", "&gt;"))


def _blocks(body_xml):
    """Yield ('p'|'tbl', xml) for top-level children in document order."""
    i, n = 0, len(body_xml)
    while i < n:
        mp = body_xml.find("<w:p", i)
        mt = body_xml.find("<w:tbl>", i)
        if mp == -1 and mt == -1:
            return
        if mt == -1 or (mp != -1 and mp < mt):
            if body_xml[mp:mp + 5] not in ("<w:p>", "<w:p ") and body_xml[mp:mp + 6] != "<w:p/>":
                i = mp + 4
                continue
            if body_xml[mp:mp + 6] == "<w:p/>":
                yield "p", "<w:p/>"
                i = mp + 6
                continue
            end = body_xml.find("</w:p>", mp)
            if end == -1:
                return
            yield "p", body_xml[mp:end + 6]
            i = end + 6
        else:
            depth, j = 0, mt
            while True:
                nxt_o = body_xml.find("<w:tbl>", j + 1)
                nxt_c = body_xml.find("</w:tbl>", j + 1)
                if nxt_c == -1:
                    return
                if nxt_o != -1 and nxt_o < nxt_c:
                    depth += 1
                    j = nxt_o
                else:
                    if depth == 0:
                        yield "tbl", body_xml[mt:nxt_c + 8]
                        i = nxt_c + 8
                        break
                    depth -= 1
                    j = nxt_c


def _parse_runs(p_xml, rels, media_map):
    runs = []
    for m in re.finditer(r"<w:r(?:\s[^>]*)?>(.*?)</w:r>", p_xml, re.S):
        body = m.group(1)
        rpr = re.search(r"<w:rPr>(.*?)</w:rPr>", body, re.S)
        props = rpr.group(1) if rpr else ""
        bold = "<w:b/>" in props or '<w:b ' in props
        ital = "<w:i/>" in props or '<w:i ' in props
        und = "<w:u " in props
        vm = re.search(r'<w:vertAlign w:val="(\w+)"/>', props)
        vert = vm.group(1) if vm else ""

        emb = re.search(r'r:embed="([^"]+)"', body) or re.search(r'r:id="([^"]+)"', body)
        if emb and emb.group(1) in rels:
            tgt = rels[emb.group(1)]
            ext = re.search(r'<wp:extent cx="(\d+)" cy="(\d+)"/>', body)
            cx = int(ext.group(1)) if ext else 914400
            cy = int(ext.group(2)) if ext else 914400
            runs.append(Run(image=media_map.get(tgt, tgt), cx=cx, cy=cy))
            continue

        txt = "".join(_xml_unescape(t) for t in
                      re.findall(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", body, re.S))
        if not txt and re.search(r"<w:br\s*/>", body):
            runs.append(Run("\x00", bold, ital, und, vert))
            continue
        # <w:br/> inside a run becomes a paragraph split later; mark with \x00
        body_nobr = re.sub(r"<w:br\s*/>", "\x00", body)
        if "\x00" in body_nobr and txt:
            pieces = []
            for seg in re.split(r"\x00", body_nobr):
                pieces.append("".join(_xml_unescape(t) for t in
                                      re.findall(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", seg, re.S)))
            txt = "\x00".join(pieces)
        if txt:
            runs.append(Run(txt, bold, ital, und, vert))
    return runs


def _parse_para(p_xml, rels, media_map):
    ppr = re.search(r"<w:pPr>(.*?)</w:pPr>", p_xml, re.S)
    props = ppr.group(1) if ppr else ""
    kind, lvl = "", 0
    if "<w:numPr>" in props:
        kind = "bullet"
        lm = re.search(r'<w:ilvl w:val="(\d+)"/>', props)
        lvl = int(lm.group(1)) if lm else 0
    runs = _parse_runs(p_xml, rels, media_map)
    return Para(runs, kind, lvl)


def _parse_table(t_xml, rels, media_map):
    grid = [int(x) for x in re.findall(r'<w:gridCol w:w="(\d+)"', t_xml)]
    tbl = Table(grid=grid)
    depth = 0
    for m in re.finditer(r"<w:tr(?:\s[^>]*)?>(.*?)</w:tr>", t_xml, re.S):
        row_xml = m.group(1)
        cells = []
        for cm in re.finditer(r"<w:tc>(.*?)</w:tc>", row_xml, re.S):
            c_xml = cm.group(1)
            sm = re.search(r'<w:gridSpan w:val="(\d+)"/>', c_xml)
            wm = re.search(r'<w:tcW w:w="(\d+)"', c_xml)
            paras = [_parse_para(px, rels, media_map)
                     for kind, px in _blocks(c_xml) if kind == "p"]
            cells.append(Cell(paras,
                              int(sm.group(1)) if sm else 1,
                              int(wm.group(1)) if wm else 0))
        if cells:
            tbl.rows.append(cells)
    return tbl


def extract(path_or_bytes) -> SDS:
    """Parse any SDS .docx into the canonical model."""
    src = (zipfile.ZipFile(io.BytesIO(path_or_bytes))
           if isinstance(path_or_bytes, (bytes, bytearray))
           else zipfile.ZipFile(path_or_bytes))
    with src as z:
        doc = z.read("word/document.xml").decode("utf-8")
        try:
            rel_xml = z.read("word/_rels/document.xml.rels").decode("utf-8")
        except KeyError:
            rel_xml = ""
        rels = {m.group(1): m.group(2) for m in
                re.finditer(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rel_xml)}
        media = {}
        media_map = {}
        for name in z.namelist():
            if name.startswith("word/media/"):
                base = os.path.basename(name)
                media[base] = z.read(name)
                media_map["media/" + base] = base
                media_map[base] = base

    body = doc[doc.find("<w:body>"):]
    sds = SDS(media=media)
    current = None

    for kind, xml in _blocks(body):
        if kind == "tbl":
            tbl = _parse_table(xml, rels, media_map)
            (current.blocks if current else sds.preamble).append(tbl)
            continue

        para = _parse_para(xml, rels, media_map)
        raw = para.text.strip()
        m = SECTION_RE.match(raw) if raw else None
        # A heading is short and has no trailing sentence content
        if m and len(raw) < 90:
            num = int(m.group(1))
            if 1 <= num <= 16:
                title = m.group(2).strip(" -\u2013\u2014:") or SECTION_TITLES[num]
                current = Section(num, title)
                sds.sections.append(current)
                continue
        (current.blocks if current else sds.preamble).append(para)

    sds.detected = _detect(sds)
    inject_pictogram_images(sds)
    return sds


PICTOGRAM_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pictograms")

PICTOGRAM_FILES = {
    "1": "GHS01_exploding_bomb.png",
    "2": "GHS02_flame.png",
    "3": "GHS03_flame_over_circle.png",
    "4": "GHS04_gas_cylinder.png",
    "5": "GHS05_corrosion.png",
    "6": "GHS06_skull_and_crossbones.png",
    "7": "GHS07_exclamation_mark.png",
    "8": "GHS08_health_hazard.png",
    "9": "GHS09_environment.png",
}

_PICTO_CODE_RE = re.compile(r"GHS-?0*([1-9])\b", re.I)
_PICTO_KEYWORDS = [
    ("1", re.compile(r"exploding\s*bomb", re.I)),
    ("2", re.compile(r"\bflame\b(?!\s*over)", re.I)),
    ("3", re.compile(r"flame\s*over\s*circle|\boxidi[sz]", re.I)),
    ("4", re.compile(r"gas\s*cylinder", re.I)),
    ("5", re.compile(r"\bcorrosion\b", re.I)),
    ("6", re.compile(r"skull\s*(?:and|&)\s*crossbones", re.I)),
    ("7", re.compile(r"exclamation\s*mark", re.I)),
    ("8", re.compile(r"health\s*hazard", re.I)),
    ("9", re.compile(r"\benvironment\b", re.I)),
]
PICTOGRAM_LABEL_RE = re.compile(r"^(?:hazard\s+)?pictograms?\s*:?\s*(.*)$", re.I)
PICTOGRAM_SIDE_EMU = 902429  # ~0.987in square, matches the Forza reference convention


def _picto_codes(text: str) -> list:
    """Ordered, deduplicated GHS codes found in a short, already-scoped string.
    GHS-code matches ('GHS08') are tried first since they're unambiguous;
    keyword fallback ('Health Hazard') only runs if no code was found, and is
    only ever applied to text already scoped to a 'Pictograms' label - the
    same words appear constantly elsewhere in Section 2 hazard prose.
    """
    codes = []
    for m in _PICTO_CODE_RE.finditer(text):
        c = m.group(1)
        if c not in codes:
            codes.append(c)
    if not codes:
        hits = []
        for code, rx in _PICTO_KEYWORDS:
            m = rx.search(text)
            if m:
                hits.append((m.start(), code))
        hits.sort()
        for _, code in hits:
            if code not in codes:
                codes.append(code)
    return codes


def _picto_run(code: str, sds: SDS) -> Run:
    fname = PICTOGRAM_FILES[code]
    if fname not in sds.media:
        with open(os.path.join(PICTOGRAM_DIR, fname), "rb") as fh:
            sds.media[fname] = fh.read()
    return Run(image=fname, cx=PICTOGRAM_SIDE_EMU, cy=PICTOGRAM_SIDE_EMU)


def inject_pictogram_images(sds: SDS) -> SDS:
    """Where Section 2 lists hazard pictograms as text ('GHS08 (Health
    Hazard)') rather than actual images, replace that text with the real
    pictogram graphics. Detection is scoped strictly to whatever block is
    labeled 'Pictograms' - a block that already contains a real image is
    left untouched, since some sources (the Forza reference docs) already
    embed the actual artwork.
    """
    sec2 = next((s for s in sds.sections if s.number == 2), None)
    if not sec2:
        return sds

    i = 0
    while i < len(sec2.blocks):
        blk = sec2.blocks[i]
        if isinstance(blk, Para):
            m = PICTOGRAM_LABEL_RE.match(blk.text.strip())
            if m:
                has_image = any(r.image for r in blk.runs)
                trailing = m.group(1).strip()
                if not has_image and trailing:
                    codes = _picto_codes(trailing)
                    if codes:
                        # Label and images go on separate lines - a shared
                        # paragraph would let the images wrap inline after
                        # the label instead of dropping to their own row.
                        blk.runs = [Run("Hazard Pictograms:", bold=True)]
                        image_para = Para([_picto_run(c, sds) for c in codes])
                        sec2.blocks.insert(i + 1, image_para)
                        i += 1
                elif not trailing and i + 1 < len(sec2.blocks):
                    nxt = sec2.blocks[i + 1]
                    if isinstance(nxt, Para) and not any(r.image for r in nxt.runs):
                        codes = _picto_codes(nxt.text)
                        if codes:
                            nxt.runs = [_picto_run(c, sds) for c in codes]
        elif isinstance(blk, Table):
            for row in blk.rows:
                if len(row) < 2:
                    continue
                label = " ".join(p.text for p in row[0].paras).strip()
                if re.match(r"^(?:hazard\s+)?pictograms?$", label, re.I):
                    has_image = any(r.image for p in row[1].paras for r in p.runs)
                    if not has_image:
                        value_text = " ".join(p.text for p in row[1].paras)
                        codes = _picto_codes(value_text)
                        if codes:
                            row[1].paras = [Para([_picto_run(c, sds) for c in codes])]
        i += 1
    return sds


def _detect(sds: SDS) -> dict:
    """Pull likely intake values out of the source so the UI can prefill.

    Section 1 identity fields are as likely to live in a table (OA4, OS2BT)
    as in bold-label paragraphs (the Forza reference doc, TAC850's Section 16).
    Both shapes are checked, with tolerant label matching, since real
    engineering drafts are not consistent about which they use.
    """
    out = {}
    label_pats = {
        "product_name": r"^product(\s*(name|identifier))?$",
        "other_means": r"other\s*means",
        "product_code": r"product\s*code",
        "recommended_use": r"recommended\s*use",
        "recommended_restrictions": r"recommended\s*restrictions|uses?\s*advised\s*against",
    }

    def check_label(label, value):
        for key, pat in label_pats.items():
            if key not in out and re.search(pat, label, re.I):
                out[key] = value
                return

    for sec in sds.sections:
        if sec.number != 1:
            continue
        i = 0
        while i < len(sec.blocks):
            blk = sec.blocks[i]
            if isinstance(blk, Para):
                t = blk.text.strip()
                m = re.match(r"([A-Za-z][A-Za-z /]+?)\s*:\s*(.+)$", t)
                if m:
                    check_label(m.group(1).strip(), m.group(2).strip())
                else:
                    # Label alone on its own line/paragraph ("Recommended
                    # Restrictions:" with nothing after the colon) - the
                    # actual value is often on the following paragraph.
                    label_only = re.match(r"^([A-Za-z][A-Za-z /]+?)\s*:?\s*$", t)
                    if label_only and i + 1 < len(sec.blocks):
                        nxt = sec.blocks[i + 1]
                        if isinstance(nxt, Para):
                            nxt_text = nxt.text.strip()
                            if nxt_text and not re.match(r"^[A-Za-z][A-Za-z /]+?\s*:", nxt_text):
                                check_label(label_only.group(1).strip(), nxt_text)
            elif isinstance(blk, Table):
                for row in blk.rows:
                    if len(row) >= 2:
                        label = " ".join(p.text for p in row[0].paras).strip()
                        value = " ".join(p.text for p in row[1].paras).strip()
                        if label and value:
                            check_label(label, value)
            i += 1

    for blk in sds.preamble:
        if isinstance(blk, Table):
            for row in blk.rows:
                if len(row) >= 2:
                    k = "".join(p.text for p in row[0].paras).strip().lower()
                    v = "".join(p.text for p in row[1].paras).strip()
                    if not v:
                        continue
                    if k.startswith("trade name"):
                        out.setdefault("trade_name", v)
                    elif k.startswith("sds"):
                        sm = re.match(r"(S-?\d+)\s*V?(\d+)?", v, re.I)
                        if sm:
                            out.setdefault("sds_number", sm.group(1))
                            if sm.group(2):
                                out.setdefault("version", sm.group(2))
                    elif k.startswith("date of issue") or k == "issue date":
                        out.setdefault("date_of_issue", normalize_date_string(v))
                    elif k.startswith("effective"):
                        out.setdefault("effective_date", normalize_date_string(v))
                    elif k.startswith("replaces"):
                        out.setdefault("replaces", v)
    return out


# ------------------------------------------------------------- rendering ---
def _rpr(run: Run, size=None, extra=""):
    p = TNR
    if run.bold:
        p += "<w:b/><w:bCs/>"
    if run.italic:
        p += "<w:i/><w:iCs/>"
    if run.underline:
        p += '<w:u w:val="single"/>'
    if run.vert:
        p += f'<w:vertAlign w:val="{run.vert}"/>'
    if size:
        p += f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
    return f"<w:rPr>{p}{extra}</w:rPr>"


def _emit_run(run: Run, rid_for):
    if run.image:
        rid = rid_for(run.image)
        if not rid:
            return ""
        cx, cy = run.cx or 914400, run.cy or 914400
        return (f'<w:r><w:rPr><w:noProof/></w:rPr><w:drawing><wp:inline distT="0" distB="0" '
                f'distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/>'
                f'<wp:effectExtent l="0" t="0" r="0" b="0"/>'
                f'<wp:docPr id="{abs(hash(run.image)) % 90000 + 1000}" name="{run.image}"/>'
                f'<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/>'
                f'</wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.'
                f'openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr>'
                f'<pic:cNvPr id="{abs(hash(run.image)) % 90000 + 1000}" name="{run.image}"/>'
                f'<pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="{rid}"/>'
                f'<a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr>'
                f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
                f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
                f'</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')
    if not run.text:
        return ""
    parts = run.text.split("\x00")
    out = ""
    for i, seg in enumerate(parts):
        if i:
            out += f"<w:r>{_rpr(run)}<w:br/></w:r>"
        if seg:
            out += (f"<w:r>{_rpr(run)}"
                    f'<w:t xml:space="preserve">{_xml_escape(seg)}</w:t></w:r>')
    return out


BULLET_NUMID = 991

BULLET_XML = (
    '<w:abstractNum w:abstractNumId="991"><w:multiLevelType w:val="hybridMultilevel"/>'
    + "".join(
        f'<w:lvl w:ilvl="{i}"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
        f'<w:lvlText w:val="{ch}"/><w:lvlJc w:val="left"/><w:pPr>'
        f'<w:ind w:left="{720 + 360 * i}" w:hanging="360"/></w:pPr><w:rPr>'
        f'<w:rFonts w:ascii="{fnt}" w:hAnsi="{fnt}" w:hint="default"/></w:rPr></w:lvl>'
        for i, (ch, fnt) in enumerate(
            [("\uf0b7", "Symbol"), ("o", "Courier New"), ("\uf0a7", "Wingdings")] * 3)
    )
    + '</w:abstractNum>'
)

BULLET_NUM = '<w:num w:numId="991"><w:abstractNumId w:val="991"/></w:num>'


def _emit_para(p: Para, rid_for, spacing_after=120, keep_next=False):
    kn = "<w:keepNext/>" if keep_next else ""
    after = p.spacing_after if p.spacing_after is not None else spacing_after
    if p.list_kind:
        ppr = (f'<w:pPr><w:pStyle w:val="ListParagraph"/>{kn}<w:numPr>'
               f'<w:ilvl w:val="{min(p.level, 8)}"/><w:numId w:val="991"/></w:numPr>'
               f'<w:spacing w:after="40" w:line="240" w:lineRule="auto"/>'
               f'<w:contextualSpacing/><w:rPr>{TNR}</w:rPr></w:pPr>')
    else:
        ppr = (f'<w:pPr>{kn}<w:spacing w:after="{after}" w:line="240" '
               f'w:lineRule="auto"/><w:rPr>{TNR}</w:rPr></w:pPr>')
    runs = "".join(_emit_run(r, rid_for) for r in p.runs)
    return f"<w:p>{ppr}{runs}</w:p>"


def _emit_table(t: Table, rid_for):
    grid = t.grid or []
    if not grid and t.rows:
        cols = max(len(r) for r in t.rows)
        grid = [int(8996 / cols)] * cols
    total = sum(grid) or 8996
    if total > 9360:                       # keep inside Letter text column
        scale = 8996 / total
        grid = [max(400, int(g * scale)) for g in grid]

    tblpr = ('<w:tblPr><w:tblW w:w="0" w:type="auto"/>'
             '<w:tblBorders>'
             '<w:top w:val="double" w:sz="4" w:space="0" w:color="auto"/>'
             '<w:left w:val="double" w:sz="4" w:space="0" w:color="auto"/>'
             '<w:bottom w:val="double" w:sz="4" w:space="0" w:color="auto"/>'
             '<w:right w:val="double" w:sz="4" w:space="0" w:color="auto"/>'
             '<w:insideH w:val="double" w:sz="4" w:space="0" w:color="auto"/>'
             '<w:insideV w:val="double" w:sz="4" w:space="0" w:color="auto"/>'
             '</w:tblBorders>'
             '<w:tblCellMar><w:top w:w="60" w:type="dxa"/><w:left w:w="100" w:type="dxa"/>'
             '<w:bottom w:w="60" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar>'
             '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" '
             'w:lastColumn="0" w:noHBand="0" w:noVBand="0"/></w:tblPr>')
    gridxml = "<w:tblGrid>" + "".join(f'<w:gridCol w:w="{g}"/>' for g in grid) + "</w:tblGrid>"

    rows = ""
    for ri, row in enumerate(t.rows):
        cells = ""
        ci = 0
        for cell in row:
            span = max(1, cell.span)
            wide = sum(grid[ci:ci + span]) or (cell.width or 1500)
            ci += span
            spanxml = f'<w:gridSpan w:val="{span}"/>' if span > 1 else ""
            body = "".join(
                _emit_para(p, rid_for, spacing_after=0) for p in cell.paras
            ) or f'<w:p><w:pPr><w:rPr>{TNR}</w:rPr></w:pPr></w:p>'
            cells += (f'<w:tc><w:tcPr><w:tcW w:w="{wide}" w:type="dxa"/>{spanxml}'
                      f'<w:vAlign w:val="center"/></w:tcPr>{body}</w:tc>')
        hdr = '<w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>' if ri == 0 else \
              '<w:trPr><w:cantSplit/></w:trPr>'
        rows += f"<w:tr>{hdr}{cells}</w:tr>"
    spacer = f'<w:p><w:pPr><w:spacing w:after="120" w:line="240" w:lineRule="auto"/>' \
             f'<w:rPr>{TNR}</w:rPr></w:pPr></w:p>'
    return f"<w:tbl>{tblpr}{gridxml}{rows}</w:tbl>{spacer}"


LABEL_RE = re.compile(r".*[:\u2013-]\s*$")


def _has_image(blk):
    return isinstance(blk, Para) and any(r.image for r in blk.runs)


def _is_label(blk):
    """A short line ending in a colon, i.e. a heading for what follows."""
    if not isinstance(blk, Para) or _has_image(blk):
        return False
    t = blk.text.strip()
    return bool(t) and len(t) < 120 and LABEL_RE.match(t) is not None


def _cohesion(blocks):
    """Decide which paragraphs must not be orphaned from what follows them.

    Hard page breaks are never inserted. Instead a paragraph is bound to the
    next block when splitting there would strand a label, so Word reflows the
    pair onto the following page by itself.
    """
    keep = [False] * len(blocks)
    for i, blk in enumerate(blocks):
        if i + 1 >= len(blocks):
            continue
        nxt = blocks[i + 1]
        if isinstance(blk, Table):
            continue
        # an image never separates from the image beside or above it
        if _has_image(blk) and (_has_image(nxt) or isinstance(nxt, Table)):
            keep[i] = True
            continue
        # a label stays with its table, its image, or its first value line
        if _is_label(blk):
            keep[i] = True
            continue
        # the line before a table stays with the table
        if isinstance(nxt, Table):
            keep[i] = True
            continue
    # a short list is a single visual unit: never split two or three items
    i = 0
    while i < len(blocks):
        blk = blocks[i]
        if isinstance(blk, Para) and blk.list_kind:
            j = i
            while (j < len(blocks) and isinstance(blocks[j], Para)
                   and blocks[j].list_kind):
                j += 1
            if 2 <= (j - i) <= 3:
                for k in range(i, j - 1):
                    keep[k] = True
            i = j
        else:
            i += 1
    # never let a bound run grow past five paragraphs, or a whole page shifts
    run = 0
    for i, k in enumerate(keep):
        if k:
            run += 1
            if run > 5:
                keep[i] = False
                run = 0
        else:
            run = 0
    return keep


_SECTION_WORD = {"en": "SECTION", "fr": "SECTION", "es": "SECCI\u00d3N"}


def _emit_heading(num, title, lang="en"):
    txt = _xml_escape(f"{_SECTION_WORD.get(lang, 'SECTION')} {num} \u2013 {title}")
    rpr = (f'<w:rPr>{TNR}<w:b/><w:bCs/><w:kern w:val="0"/><w:sz w:val="36"/>'
           f'<w:szCs w:val="36"/><w:u w:val="thick"/></w:rPr>')
    return (f'<w:p><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120" '
            f'w:line="240" w:lineRule="auto"/>{rpr}</w:pPr>'
            f"<w:r>{rpr}<w:t>{txt}</w:t></w:r></w:p>")


# --------------------------------------------------------------- helpers ---
DATE_DOT_RE = re.compile(r"(?<!\d)(\d{1,2})\.(\d{1,2})\.(\d{2,4})(?!\d)")


def normalize_date_string(s: str) -> str:
    """Convert dot-separated dates (08.18.2026) to slash-separated
    (08/18/2026). Everything else - including non-date text like 'N/A' or
    'Revision 6' in the Replaces field - passes through untouched.
    """
    if not s:
        return s
    return DATE_DOT_RE.sub(lambda m: f"{m.group(1)}/{m.group(2)}/{m.group(3)}", s)


def build_dcn(sds_number: str, version: str, trade_name: str) -> str:
    """S-0196 + 1 + R-OS86  ->  'S0196_V1-R-OS86'  (en dash, no surrounding spaces)."""
    core = re.sub(r"[^0-9A-Za-z]", "", str(sds_number))
    ver = re.sub(r"\D", "", str(version)) or "1"
    return f"{core}_V{ver}\u2013{trade_name}".strip()


def build_sds_field(sds_number: str, version: str) -> str:
    """Top-table form keeps the hyphen: 'S-0196 V1'."""
    num = str(sds_number).strip()
    ver = re.sub(r"\D", "", str(version)) or "1"
    return f"{num} V{ver}"


S1_LABEL_PATTERNS = {
    "product_name": r"^product(\s*(name|identifier))?$",
    "other_means": r"other\s*means",
    "product_code": r"product\s*code",
    "recommended_use": r"recommended\s*use",
    "recommended_restrictions": r"recommended\s*restrictions|uses?\s*advised\s*against",
}


_PHONE_HOURS = {
    "en": "402-731-9300 (Available 8:00 am \u2013 4:30 pm CST)",
    "fr": "402-731-9300 (Disponible de 8 h \u00e0 16 h 30, heure du Centre)",
    "es": "402-731-9300 (Disponible de 8:00 a. m. a 4:30 p. m., hora del Centro)",
}
_EMERGENCY_LINE = {
    "en": "Chemtrec 1(800)-424-9300",
    "fr": "Chemtrec 1 (800) 424-9300",
    "es": "Chemtrec 1 (800) 424-9300",
}
_ADDRESS_LINES = {
    "en": ["Forza, Inc.", "3211 Nebraska Ave, Suite #300", "Council Bluffs, IA 51501, USA"],
    "fr": ["Forza, Inc.", "3211 Nebraska Ave, bureau 300", "Council Bluffs, IA 51501, \u00c9tats-Unis"],
    "es": ["Forza, Inc.", "3211 Nebraska Ave, Suite #300", "Council Bluffs, IA 51501, EE. UU."],
}


def canonical_section1(intake: dict, lang: str = "en", text_map: dict = None) -> list:
    """Section 1 is always rendered in Forza's fixed house format, regardless
    of how the source document structured it. Only the five identity/use
    fields vary by product; the supplier block is constant company boilerplate
    and is never taken from the source, since engineering drafts often carry
    placeholder brackets there (e.g. '[Company name and address - insert]').
    text_map (from build_translation_map) supplies DeepL translations for the
    free-text fields (product name, other means, recommended use/restrictions)
    when lang != 'en'; a field with no entry is left in English.
    """
    L = LANG.get(lang, LANG["en"])
    text_map = text_map or {}
    tr = lambda v: text_map.get((v or "").strip(), v) if lang != "en" else v
    B = lambda t: Run(t, bold=True)
    addr = _ADDRESS_LINES.get(lang, _ADDRESS_LINES["en"])
    return [
        Para([B(f"{L['product_name']} "), Run(tr(intake.get("product_name", "")))], spacing_after=120),
        Para([B(f"{L['other_means']} "), Run(tr(intake.get("other_means", "")))],
             spacing_after=120),
        Para([B(f"{L['product_code']} "), Run(intake.get("product_code", ""))],
             spacing_after=120),
        Para([B(f"{L['recommended_use']} "), Run(tr(intake.get("recommended_use", "")))],
             spacing_after=120),
        Para([B(f"{L['recommended_restrictions']} "), Run(tr(intake.get("recommended_restrictions", "")))],
             spacing_after=200),
        Para([Run(L["suppliers_details"])], spacing_after=120),
        Para([B(L["company"] + " ")], spacing_after=20),
        Para([Run(addr[0])], spacing_after=0),
        Para([Run(addr[1])], spacing_after=0),
        Para([Run(addr[2])], spacing_after=140),
        Para([B(L["company_phone"] + " ")], spacing_after=20),
        Para([Run(_PHONE_HOURS.get(lang, _PHONE_HOURS["en"]))], spacing_after=140),
        Para([B(L["emergency_phone"] + " ")], spacing_after=20),
        Para([Run(_EMERGENCY_LINE.get(lang, _EMERGENCY_LINE["en"]))], spacing_after=120),
    ]


PREPARED_BY_RE = re.compile(r"prepared\s*by", re.I)
FORZA_PREPARED_BY = "Forza, Inc."


def _s16_route(label: str, ver: str, date: str) -> str | None:
    """Decide what a Section 16 label wants: a version, a date, or both.

    'Revision' -> version. 'Issue Date' -> date. 'Revision date' contains
    both words but means a date, not a version, so a combined field is only
    assumed when the label has an explicit joiner ('/', '&', 'and') between
    the two concepts, e.g. 'Revision / Issue Date'.
    """
    has_date = bool(re.search(r"date", label, re.I))
    has_ver = bool(re.search(r"revision|version", label, re.I))
    combined = bool(re.search(r"(revision|version)\s*[/&]\s*\w*\s*date", label, re.I)
                    or re.search(r"(revision|version)\s+and\s+\w*\s*date", label, re.I))
    if combined:
        return f"V{ver} \u2013 {date}" if date else f"V{ver}"
    if has_date:
        return date or None
    if has_ver:
        return f"V{ver}"
    return None


PLACEHOLDER_BRACKET_RE = re.compile(r"\[[^\[\]]{1,80}\]")
PLACEHOLDER_KEYWORD_RE = re.compile(r"\b(TBD|PLACEHOLDER|FIXME|XXXX+)\b", re.I)

# Section 1 boilerplate is generated fresh every time and never sourced from
# the document, so it can never itself be short or contain a placeholder.
# Only sections 2-16 are worth scanning.
_EMPTY_SECTION_CHARS = 15   # below this with no table present -> "empty"
_SHORT_SECTION_CHARS = 60   # below this -> "unusually short", worth a glance


def _section_text_len(sec: Section) -> tuple:
    """(char_count, has_table) for a section's blocks."""
    chars, has_table = 0, False
    for blk in sec.blocks:
        if isinstance(blk, Table):
            has_table = True
            for row in blk.rows:
                for cell in row:
                    chars += sum(len(p.text) for p in cell.paras)
        else:
            chars += len(blk.text)
    return chars, has_table


def preflight(sds: SDS, intake: dict) -> dict:
    """Surface what might be wrong before the person downloads, rather than
    after. Nothing here blocks generation - it's a pre-flight check, not a
    validator - but a silent gap (a missing section, a leftover
    '[insert here]', a Section 16 with no date field to stamp) is much
    cheaper to catch here than after the document has gone out.
    """
    found = sorted(s.number for s in sds.sections)
    missing = [n for n in range(1, 17) if n not in found]

    empty, short, placeholders = [], [], []
    for sec in sds.sections:
        if sec.number == 1:
            continue  # always regenerated fresh; nothing to check
        chars, has_table = _section_text_len(sec)
        title = sec.title or SECTION_TITLES.get(sec.number, "")
        if chars < _EMPTY_SECTION_CHARS and not has_table:
            empty.append((sec.number, title))
        elif chars < _SHORT_SECTION_CHARS:
            short.append((sec.number, title, chars))

        for blk in sec.blocks:
            texts = ([p.text for row in blk.rows for c in row for p in c.paras]
                     if isinstance(blk, Table) else [blk.text])
            for t in texts:
                for rx in (PLACEHOLDER_BRACKET_RE, PLACEHOLDER_KEYWORD_RE):
                    for m in rx.finditer(t):
                        snippet = t[max(0, m.start() - 20):m.end() + 20].strip()
                        placeholders.append((sec.number, title, snippet))

    sec16 = next((s for s in sds.sections if s.number == 16), None)
    ver_found, date_found = False, False
    if sec16:
        for blk in sec16.blocks:
            rows = ([" ".join(p.text for p in row[0].paras).strip()
                     for row in blk.rows if row] if isinstance(blk, Table)
                    else [re.match(r"([A-Za-z][A-Za-z /]+?)\s*:\s*", blk.text.strip())])
            labels = rows if isinstance(blk, Table) else \
                     [m.group(1) for m in rows if m]
            for label in labels:
                has_date = bool(re.search(r"date", label, re.I))
                has_ver = bool(re.search(r"revision|version", label, re.I))
                date_found = date_found or has_date
                ver_found = ver_found or has_ver

    blank_optional = [f for f in ("recommended_use", "recommended_restrictions",
                                  "other_means", "product_code")
                      if not str(intake.get(f, "")).strip()]

    warnings = []
    if missing:
        warnings.append(f"Missing sections: {', '.join(map(str, missing))}.")
    if empty:
        names = ", ".join(f"{n} ({t})" for n, t, in empty)
        warnings.append(f"These sections parsed with no real content: {names}.")
    if short:
        names = ", ".join(f"{n} ({t}, {c} chars)" for n, t, c in short)
        warnings.append(f"Unusually short, worth a glance: {names}.")
    if placeholders:
        warnings.append(f"{len(placeholders)} placeholder-looking snippet(s) "
                        f"found (e.g. brackets or TBD/FIXME) - see detail below.")
    if not ver_found:
        warnings.append("Section 16 has no field that looks like a version/"
                        "revision - the version number will not appear there.")
    if not date_found:
        warnings.append("Section 16 has no field that looks like a date - "
                        "the date of issue will not appear there.")
    if blank_optional:
        warnings.append(f"Not filled in: {', '.join(blank_optional)}.")

    return dict(sections_found=found, sections_missing=missing,
                empty_sections=empty, short_sections=short,
                placeholders=placeholders,
                section16=dict(version_field_present=ver_found,
                               date_field_present=date_found),
                blank_optional_fields=blank_optional,
                images_total=len(sds.media), warnings=warnings)


def _patch_s16_para(blk, ver, date_of_issue):
    """Patch a Section 16 paragraph line by line, since engineering drafts
    often put several lines in one paragraph via a manual line break (e.g.
    'Prepared by: ...' and a trailing disclaimer sharing one paragraph).
    Replacing blk.runs wholesale for a matched line used to silently delete
    every other line sharing that paragraph. Each line is now checked and
    patched independently; anything that isn't a recognized label - the
    disclaimer included - is carried through untouched.
    Returns (prepared_found, dated) for this block.
    """
    lines = blk.text.split("\x00")
    new_lines = []
    prepared_found = False
    dated = False
    for line in lines:
        t = line.strip()
        m = re.match(r"([A-Za-z][A-Za-z /]+?)\s*:\s*", t)
        label = m.group(1).strip() if m else None
        if label and PREPARED_BY_RE.search(label):
            new_lines.append([Run(label + ": ", bold=True), Run(FORZA_PREPARED_BY)])
            prepared_found = True
            continue
        val = _s16_route(label, ver, date_of_issue) if label else None
        if val is not None:
            new_lines.append([Run(label + ": ", bold=True), Run(val)])
            dated = True
            continue
        new_lines.append([Run(line)])

    combined = []
    for idx, line_runs in enumerate(new_lines):
        if idx > 0:
            first = line_runs[0]
            line_runs = [Run("\x00" + first.text, first.bold, first.italic,
                             first.underline, first.vert)] + line_runs[1:]
        combined.extend(line_runs)
    blk.runs = combined
    return prepared_found, dated


def apply_intake(sds: SDS, intake: dict) -> SDS:
    """Stamp Section 16 date / version onto whatever shape the source used.

    Section 1 is not touched here: it is fully replaced at render time by
    canonical_section1(), regardless of what the source contained, so no
    in-place patching of it is needed or attempted.
    """
    ver = re.sub(r"\D", "", str(intake.get("version", "1"))) or "1"
    intake = dict(intake)
    intake["date_of_issue"] = normalize_date_string(intake.get("date_of_issue", ""))

    def set_cell(cell, text):
        cell.paras = [Para([Run(text)])]

    for sec in sds.sections:
        if sec.number == 16:
            prepared_found = False
            last_dated_idx = -1
            for i, blk in enumerate(sec.blocks):
                if isinstance(blk, Para):
                    found, dated = _patch_s16_para(blk, ver, intake.get("date_of_issue", ""))
                    prepared_found = prepared_found or found
                    if dated:
                        last_dated_idx = i
                elif isinstance(blk, Table):
                    for row in blk.rows:
                        if len(row) < 2:
                            continue
                        label = " ".join(p.text for p in row[0].paras).strip()
                        if PREPARED_BY_RE.search(label):
                            set_cell(row[1], FORZA_PREPARED_BY)
                            prepared_found = True
                            continue
                        val = _s16_route(label, ver, intake.get("date_of_issue", ""))
                        if val is not None:
                            set_cell(row[1], val)
                            last_dated_idx = i

            if not prepared_found:
                # Some engineering drafts omit "Prepared by" entirely (TAC850
                # does). It must always appear, so insert it - right after the
                # revision/date line(s) if any were found, otherwise at the
                # top of the section.
                new_para = Para([Run("Prepared by: ", bold=True), Run(FORZA_PREPARED_BY)])
                sec.blocks.insert(last_dated_idx + 1 if last_dated_idx >= 0 else 0, new_para)
    return sds


TOP_ROWS = ["trade_name", "sds_field", "date_of_issue", "replaces", "effective_date"]
TOP_LABELS = ["Trade Name", "SDS #", "Date of Issue", "Replaces", "Effective Date"]


def render(sds: SDS, intake: dict, shell_path: str, out_path: str, lang: str = "en",
          deepl_api_key: str = None) -> str:
    """Inject the model into a vertical shell and write the finished .docx.
    lang: 'en' | 'fr' | 'es'. Fixed template text (title, table labels,
    section titles, Section 1 field names, footer) is translated for the
    requested language. Hazard/precautionary statement lines that start with
    a recognized GHS H-code/P-code use the verified official phrase for that
    code (see translate_code_line). Everything else - free-text Section 1
    fields and all other source-derived content in Sections 2-16 - is
    translated via the DeepL API when deepl_api_key is supplied; with no key,
    that text is left in English (see build_translation_map).
    The top-table SDS # field and the footer DCN are both derived from
    intake["sds_number"] - callers generating multiple languages of the same
    document pass a per-language sds_number in intake, since each language
    is a distinct controlled copy and the DCN follows directly from it.
    """
    L = LANG.get(lang, LANG["en"])
    titles = SECTION_TITLES_BY_LANG.get(lang, SECTION_TITLES)
    text_map = build_translation_map(sds, intake, lang, deepl_api_key)
    ver = re.sub(r"\D", "", str(intake.get("version", "1"))) or "1"
    values = {
        "trade_name": intake.get("trade_name", ""),
        "sds_field": build_sds_field(intake.get("sds_number", ""), ver),
        "replaces": normalize_date_string(intake.get("replaces", "")),
        "date_of_issue": normalize_date_string(intake.get("date_of_issue", "")),
        "effective_date": normalize_date_string(intake.get("effective_date", "")),
        "lbl_trade_name": L["trade_name"],
        "lbl_sds_num": L["sds_num"],
        "lbl_date_issue": L["date_issue"],
        "lbl_replaces": L["replaces"],
        "lbl_effective": L["effective"],
        "lbl_doc_title": L["doc_title"],
        "lbl_page": L["page"],
        "lbl_of": L["of"],
        "lbl_dcn": L["dcn"],
    }
    dcn = build_dcn(intake.get("sds_number", ""), ver, intake.get("trade_name", ""))

    with zipfile.ZipFile(shell_path) as z:
        parts = {n: z.read(n) for n in z.namelist()}

    # stage images, allocating fresh relationship ids
    rel_extra, ct_extra, used = "", "", {}
    next_id = [1]

    def rid_for(fname):
        if fname in used:
            return used[fname]
        if fname not in sds.media:
            return None
        rid = f"rIdImg{next_id[0]}"
        next_id[0] += 1
        used[fname] = rid
        parts[f"word/media/{fname}"] = sds.media[fname]
        return rid

    # ---- body
    doc = parts["word/document.xml"].decode("utf-8")
    chunks = []
    for sec in sorted(sds.sections, key=lambda s: s.number):
        # For English, respect the source's own section title wording when it
        # has one (falls back to the canonical name only if it didn't). For
        # French/Spanish output, always use the canonical translated title -
        # the source's title text is English and should never leak through.
        heading_title = titles.get(sec.number, "") if lang != "en" else (sec.title or titles.get(sec.number, ""))
        chunks.append(_emit_heading(sec.number, heading_title, lang=lang))
        if sec.number == 1:
            # Always the fixed Forza house format, regardless of source shape.
            for p in canonical_section1(intake, lang, text_map):
                chunks.append(_emit_para(p, rid_for))
            continue
        blocks = [b for b in sec.blocks if isinstance(b, Table) or b.runs]
        blocks = [_translate_block(b, lang, text_map) for b in blocks]
        keep = _cohesion(blocks)
        for blk, kn in zip(blocks, keep):
            if isinstance(blk, Table):
                chunks.append(_emit_table(blk, rid_for))
            else:
                chunks.append(_emit_para(blk, rid_for, keep_next=kn))
    body_xml = "".join(chunks)

    marker = re.search(r"<w:p>(?:(?!</w:p>).)*\{\{BODY\}\}.*?</w:p>", doc, re.S)
    doc = doc[:marker.start()] + body_xml + doc[marker.end():]
    for key, val in values.items():
        doc = doc.replace("{{" + key.upper() + "}}", _xml_escape(val))
    doc = doc.replace("{{SDS_NUMBER}} V{{VERSION}}", _xml_escape(values["sds_field"]))
    doc = re.sub(r"\{\{[A-Z_]+\}\}", "", doc)
    parts["word/document.xml"] = doc.encode("utf-8")

    # ---- footer DCN
    ftr = parts["word/footer1.xml"].decode("utf-8")
    ftr = ftr.replace("{{DCN}}", _xml_escape(dcn))
    ftr = ftr.replace("{{LBL_PAGE}}", _xml_escape(L["page"]))
    ftr = ftr.replace("{{LBL_OF}}", _xml_escape(L["of"]))
    ftr = ftr.replace("{{LBL_DCN}}", _xml_escape(L["dcn"]))
    parts["word/footer1.xml"] = ftr.encode("utf-8")

    # ---- relationships + content types for staged images
    rl = parts["word/_rels/document.xml.rels"].decode("utf-8")
    add = "".join(f'<Relationship Id="{rid}" Type="{R}/image" Target="media/{f}"/>'
                  for f, rid in used.items())
    parts["word/_rels/document.xml.rels"] = rl.replace("</Relationships>", add + "</Relationships>").encode("utf-8")

    num = parts["word/numbering.xml"].decode("utf-8")
    if 'w:numId="991"' not in num:
        # schema order: every abstractNum must precede every num
        first_num = num.find("<w:num ")
        if first_num == -1:
            num = num.replace("</w:numbering>", BULLET_XML + BULLET_NUM + "</w:numbering>")
        else:
            num = num[:first_num] + BULLET_XML + num[first_num:]
            num = num.replace("</w:numbering>", BULLET_NUM + "</w:numbering>")
    parts["word/numbering.xml"] = num.encode("utf-8")

    ct = parts["[Content_Types].xml"].decode("utf-8")
    for ext, mime in (("png", "image/png"), ("jpeg", "image/jpeg"), ("jpg", "image/jpeg"),
                      ("gif", "image/gif"), ("emf", "image/x-emf"), ("wmf", "image/x-wmf")):
        if f'Extension="{ext}"' not in ct:
            ct = ct.replace("</Types>", f'<Default Extension="{ext}" ContentType="{mime}"/></Types>')
    parts["[Content_Types].xml"] = ct.encode("utf-8")

    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    if os.path.exists(out_path):
        os.remove(out_path)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", parts.pop("[Content_Types].xml"))
        for name, data in parts.items():
            z.writestr(name, data)
    return out_path
