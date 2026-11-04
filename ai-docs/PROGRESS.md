# Change Log

_Populated as we work. Each entry = problem + solution + files changed._

---

## 2026-11-04 — Préapurement DS (cas d'exception): lot du grand partiel = réf. LTA, pas réf. aéroport

**Problème:** LTA 235-99381402 (partiel 1 = 1329.1 kg, partiel 2 = 24 kg exception, DUM 1 split). Phase 2 saisissait `24/1/1` comme référence du lot Dépotage de DUM 1 (morceau 1/1 du grand partiel) au lieu de `235-99381402/1/1`; BADR n'a trouvé aucun lot → "Erreur validation données" (retry 1/3). Cause: `get_dum_preapurement_lots` (branche exception) prenait comme base la référence aéroport du petit partiel (ici "24") alors que les lots de l'Etat de Dépotage sont `<réf. LTA>/<split_id>` (`create_etat_depotage_partial` et la modification blocage + partiel). Cela ne marchait que si la référence aéroport saisie était égale à la référence LTA.

**Solution (`badr_login_test.py`, `get_dum_preapurement_lots`):** lot Dépotage du grand partiel = `{validated_lta_reference}/{split_id}` (ou `/{dum}`), y compris dans la branche "legacy". La référence aéroport ne sert plus que pour le lot DS MEAD du petit partiel (inchangé: `24`).

**Vérifié** sur la vraie config: DUM 1 → `235-99381402/1/1` (DS 5278 T) + `24` (DS MEAD 8478 B); DUM 2 → `235-99381402/2`; DUM 14 → `235-99381402/14`. Non testé en live BADR.

---

## 2026-11-04 — `git pull` bloqué sur un autre PC + known_companies.json ne crée plus de conflits

**Problème:** sur un autre PC (`~/Desktop/medafrica-app`), le dépôt était bloqué dans un conflit non résolu (`UU badr_login_test.py`, commit local `b6f05b2` du 20/08). À chaque lancement, le `git pull --autostash origin main` de l'exe échouait aussitôt (exit 128: "Pulling is not possible because you have unmerged files"), sans aucun message: la sortie de git est jetée (`except: pass`). Un `stash@{0}: autostash` restait en attente (réapplication échouée). Le PC avait donc raté les mises à jour du 21/09 et du 04/11. Le dépôt était public (accès anonyme OK), git était bien dans le PATH, `origin` était correct; l'exe est dans `medafrica-app/dist/`, donc git trouve bien le dépôt.

Cause racine côté fichiers: `known_companies.json` est suivi par git ET écrit par l'app (`script_all_fuzy_match.py` y ajoutait chaque société apprise). Il diffère donc d'un PC à l'autre, et dès que le dépôt le modifie à son tour (mon commit `4f0b889` du jour, 39 → 69 noms), l'`--autostash` entre en conflit.

**Solution (`script_all_fuzy_match.py`, `.gitignore`, fichiers JSON):**
- 3 fichiers, un seul écrit par l'app:
  - `known_companies.json`: suivi, FIGÉ à sa version d'origine (39 noms, commit `ed322d1`). L'app ne l'écrit plus. Tous les PCs ont déjà exactement cette version: le pull ne la modifie donc plus.
  - `known_companies_shared.json`: suivi, géré par le mainteneur (29 noms appris sur le PC de dev, anciennement dans le commit `4f0b889`). L'app ne l'écrit jamais.
  - `known_companies_local.json`: NON suivi (jamais dans `.gitignore` en liste blanche). Noms appris sur ce PC, seul fichier écrit par l'app.
- Lecture = union des 3 fichiers (sans doublon, casse/espaces ignorés). `save_companies_database` n'écrit que les noms qui ne sont pas déjà dans les deux fichiers suivis. `add_company_to_database` et la sauvegarde finale relisent les 3 fichiers (un nom ajouté par une autre instance n'est pas perdu).
- Nettoyage automatique une seule fois par PC (`_migrate_legacy_local_changes`): si `known_companies.json` diffère de git HEAD, les noms locaux passent dans `known_companies_local.json`, puis le fichier suivi est restauré depuis HEAD (arbre de travail propre). Sans effet hors dépôt git, ou si le fichier est illisible (marqueurs de conflit): il n'est alors pas touché.
- Règle pour la suite: ne plus modifier `known_companies.json` dans git. Les noms partagés vont dans `known_companies_shared.json`.

**Vérifié:** (1) tests sur dépôts jetables: chargement, migration, ajout/doublons, deux instances, PC neuf, dossier hors git, marqueurs de conflit; (2) simulation de vrais `git pull --autostash` sur des clones du dépôt réel (voir détails ci-dessous / message de fin).

**À faire sur le PC bloqué:** réparé à la main (`git reset --hard origin/main` après sauvegarde dans `~/Desktop/pull_backup/`); le `stash@{0}` est conservé, ne pas le réappliquer.

**Files:** `script_all_fuzy_match.py`, `.gitignore`, `known_companies.json` (restauré à l'identique d'origine), `known_companies_shared.json` (nouveau).

---

## 2026-11-04 — Champ "Nb Articles Bloqués" (onglet Préparation) → ligne 14 du fichier LTA

**Demande:** l'utilisateur saisit lui-même le nombre d'articles bloqués dans la section blocage de la carte LTA (1 article = 1 contenant). Ce nombre sert à corriger "Nombre total Contenant(s)" de l'ED blocage modifié (ex: 19 + 1829 = 1848). Auparavant, l'app ne le comptait que dans le tableau des lots de l'ED BADR.

**Solution:**
- `gui/screens/preparation.py`:
  - Section blocage: 2e ligne "Nb Articles Bloqués:" + aide "1 article = 1 contenant (requis si LTA partiel)".
  - Sauvegarder: entier > 0 obligatoire quand blocage + partiel sont cochés (facultatif pour un blocage seul).
  - La valeur est conservée au rafraîchissement de la table (`keep_inputs`) et écrite aussi après "Configurer".
- `gui/utils/file_utils.py`:
  - `update_lta_blocage(..., blocked_articles)` écrit la ligne 14 du fichier `[X]eme LTA.txt`: 3e ligne vide avant le séparateur `-----` du format du script de préparation; aucun parseur ne la lit.
  - Ne réécrit jamais une ligne 14 occupée par autre chose (ancien format où le séparateur tombe en ligne 14).
  - `get_lta_blocage_info` renvoie `blocked_articles`.
- `badr_login_test.py`:
  - `read_blocked_articles_count` lit la ligne 14. `process_blocage_partial_ed` passe la valeur à `modify_etat_depotage_for_blocage(..., blocked_articles)`.
  - MED.5c: contenants des articles bloqués = valeur saisie. À défaut, le décompte du tableau des lots de l'ED. Une différence entre les deux est signalée.
  - Le filet E2800435 reste en place, avec son contrôle de cohérence basé sur le tableau des lots.

**Vérifié:**
- Sur une copie du vrai `1er LTA.txt`: ligne 14 = 19, séparateur intact; Phase 1 lit 19; `parse_lta_file` trouve toujours les 14 DUMs; un ancien format (séparateur en ligne 14) n'est pas écrasé.
- Écran Préparation réel sur une copie du dossier: vide → erreur (blocage + partiel); "abc" → erreur; 19 → sauvegardé (ligne 14 = 19); une saisie non sauvegardée (21) survit au rafraîchissement.

**Files:** `gui/screens/preparation.py`, `gui/utils/file_utils.py`, `badr_login_test.py`.

---

## 2026-11-04 — Blocage + partiel: onglet Quantités de l'ED modifié (contenants / poids total) — E2800435

**Problème:** à la modification de l'ED blocage (partiel 1, DS 5986 V), l'onglet Quantités gardait les totaux d'origine (Nombre total Contenant(s) = 1830). À la validation, BADR calcule la somme des lots: 19 articles bloqués (MA…, 1 contenant chacun) + 1829 positions des lots DUM du partiel = 1848. Résultat: "E2800435 - Le nombre total de contenants calculé par le Système (1848) est différent du nombre total de contenants déclaré (1830)". Même risque pour le Poids brut total, qui doit être le poids du partiel (1329.1).

**Solution (`badr_login_test.py`, mode blocage + partiel uniquement):**
- Nouvelle étape MED.5c dans `modify_etat_depotage_for_blocage`, après suppression des lots à la référence LTA et avant création des lots DUM:
  - `read_ed_lots_table` lit tous les lots restants (= articles bloqués), toutes les pages. Les colonnes "contenant" / "poids" sont repérées par l'en-tête du tableau, sans mélanger les sous-tableaux.
  - Contenants attendus = contenants des articles bloqués (1 par article si la colonne est absente) + somme des P des lots DUM du partiel.
  - Poids attendu = poids du partiel. Avertissement si le poids des articles bloqués lu ≠ poids bloqué saisi.
  - `fix_ed_quantites` ouvre l'onglet Quantités, lit les valeurs (poids via le champ caché `_hinput`), et ne corrige que ce qui est faux (vérifié après saisie). Puis retour à l'onglet LTA (`open_ed_tab`). Même ordre que la création d'ED: Quantités puis lots, puis SAUVEGARDER.
- Filet de sécurité MED.9: si VALIDER renvoie E2800435, `retry_validation_with_system_contenants` reprend la valeur calculée par BADR. Elle n'est acceptée que si elle est cohérente: la part restante, une fois retirés les lots DUM ajoutés, doit être ≥ nombre d'articles bloqués, et égale à leur somme si elle a été lue. Ensuite: correction de Quantités → SAUVEGARDER → VALIDER (une seule fois).
- Nouveaux helpers: `open_ed_tab`, `_parse_badr_number`, `read_ed_lots_table`, `fix_ed_quantites`, `click_ed_menu_action`, `retry_validation_with_system_contenants`.

**Vérifié** sur une maquette locale (Edge headless) reprenant le HTML Quantités fourni et un tableau de 19 lots MA:
- avec colonne contenants/poids: 19 contenants / 10.22 kg lus, poids déjà juste non modifié, contenants 1830 → 1848;
- sans colonne contenants: 1 par article; poids 1330 → 1329.1; contenants → 1848;
- le sous-tableau de détail du lot (999) est ignoré;
- le message E2800435 réel est analysé → 1848 déclaré + SAUVEGARDER + VALIDER; une valeur incohérente (1835) est refusée.

Non testé en live BADR: les colonnes réelles du tableau des lots sont inconnues (repli: 1 contenant par article bloqué).

**Files:** `badr_login_test.py`, `ai-docs/PROJECT.md`.

---

## 2026-11-04 — Blocage + partiel: Phase 1 CRÉAIT l'ED du plus grand partiel au lieu de MODIFIER l'ED blocage; DS du blocage = DS du plus grand partiel

**Problème:** LTA 235-99381402 (1er LTA, partiel 1 = 1329.1 kg DS 5986 V, partiel 2 = 24 kg exception, bloqué 10.22 kg). Phase 1 a affiché "CRÉATION ED - PARTIEL 1" au lieu de "modifier etat de dépotage". Cause: Phase 1 ne reconnaissait le blocage QUE par la ligne 5 du fichier `1er LTA.txt`; elle n'était pas encore à "blocage" au moment du lancement (onglet Préparation sauvegardé à 12:39, après la config partielle de 12:33), alors que la config partielle contenait déjà `blocage_partial_number: 1` / `blocage_weight: 10.22`. Aggravant: après "Configurer", la table Préparation se rechargeait depuis les fichiers et effaçait les saisies non sauvegardées (blocage, DS, lieu). L'utilisateur précise: la DS du blocage EST la DS du plus grand partiel; la saisir dans l'onglet Préparation est facultatif.

**Solution:**
- `badr_login_test.py`: nouveau `resolve_blocage_for_partial(blocage_info, partial_config)`, appelé dès que la config partielle est lue dans `process_lta_folder_ed_only`. Si la config a `blocage_weight` > 0, le LTA est traité comme blocage même si la ligne 5 n'est pas marquée. Poids total pris dans l'ordre: ligne 12, puis `blocage_lta_total` de la config, puis somme des poids des partiels. Si la ligne 13 diffère du poids bloqué de la config, la valeur du fichier est gardée et le garde-fou exact refuse la config périmée.
- `process_blocage_partial_ed`: DS de l'ED modifié = DS du plus grand partiel (config). La DS de l'onglet Préparation (ligne 2 du fichier shipper, lue par `read_preparation_ds`) est facultative. Si elle est saisie ET différente (`_ds_key` ignore espaces et "/"), Phase 1 refuse AVANT toute correction Excel / BADR. Le bandeau Phase 1 affiche "DS du blocage = DS du partiel N".
- `gui/screens/preparation.py`:
  - Après sauvegarde de la config partielle avec blocage coché, le blocage (lignes 5/12/13) est écrit tout de suite dans le fichier LTA (`update_lta_blocage`).
  - DS de la carte vide → remplie avec la DS du plus grand partiel; DS différente → avertissement.
  - `populate_lta_table(keep_inputs=True)` conserve les saisies non sauvegardées lors du rafraîchissement.
  - "Sauvegarder" refuse un LTA partiel dont la carte ne concorde pas avec la config: blocage coché/décoché, poids bloqué, ou DS ≠ DS du plus grand partiel (`_blocage_partial_mismatch`).
- `gui/screens/partial_config_dialog.py`: la config enregistre aussi `blocage_lta_total`. L'aperçu du partiel blocage indique "DS Série de ce partiel = DS du blocage".

**Vérifié:** Phase 1 en dry-run sur une COPIE des fichiers réels (BADR et fermeture Excel simulés):
- (A) ligne 5 vide (état du run de l'utilisateur): blocage détecté via la config. Partiel 1 → MODIFY avec DS 5986 V; DUMs 1318.88 kg exacts; 14 DUMs = generated_excel corrigé; partiel 2 ignoré (exception); aucun CREATE.
- (B) ligne 5 = blocage + DS Préparation identique: même résultat.
- (C) DS Préparation "5986 X": refusé, aucun appel BADR.
- Helpers GUI testés sur la config réelle (6 cas + remplissage automatique de la DS). Non testé en live BADR.

**Files:** `badr_login_test.py`, `gui/screens/preparation.py`, `gui/screens/partial_config_dialog.py`.

---

## 2026-11-04 — Blocage + partiel: répartition des DUMs EXACTE (centième) + correction blocage dans le dialogue

**Problème (retour utilisateur):** l'aperçu du partiel blocage totalisait 1318.9 au lieu de 1318.88, avec un faux `DUM 14 14/1: 48.9kg ⚠️ PARTIEL`.

**Causes (deux):**
1. Toute la répartition (`_calculate_dum_distribution`) et les lots ED partiels (`get_dum_lots_for_partial`, `p_brut`) arrondissaient au **dixième** (48.88 → 48.9).
2. Pour ce LTA, `generated_excel` contenait encore le poids bloqué dans ses DUMs (total 1353.1 = poids LTA complet, DUM 14 = 59.1 — vérifié: la liste de l'utilisateur totalise exactement 1353.1). La Phase 1 retire cet excédent (`correct_blocage_weights`, en partant du dernier DUM: 59.1 → 48.88) mais le dialogue répartissait AVANT cette correction → DUM 14 coupé et ~10.2 kg non attribués. (Sur d'autres LTA blocage, generated_excel vaut déjà LTA − bloqué: ça varie.)

**Solution:**
- `partial_config_dialog.py`: `WEIGHT_DECIMALS = 2` — 23 arrondis de poids passés au centième, seuils de split 0.1 → 0.005, cumuls/cibles du cas normal arrondis (comparaisons exactes, pas de micro-morceaux); affichage `.2f`.
- `_apply_blocage_correction()` (module): copie pure de la règle de `correct_blocage_weights` (excédent soustrait depuis le dernier DUM en gardant P,BRUT > P,NET, déficit ajouté au dernier); `_load_lta_data` lit aussi P,NET; le dialogue corrige les DUMs avec (total LTA − bloqué) avant de répartir. Copie nécessaire: `MedAfrica.spec` embarque `gui/` mais pas `badr_login_test.py`, qui importe Selenium au chargement.
- Le dialogue reçoit aussi le **poids total** live de la carte (`preparation.py`), affiche "DUMs generated_excel X → corrigés Y", signale les kg non attribués dans l'aperçu et **refuse de sauvegarder** une config blocage dont la répartition n'est pas exacte.
- `badr_login_test.py`: `get_dum_lots_for_partial` → `p_brut` au centième (vaut aussi pour les partiels normaux: lots ED = poids exacts). `process_blocage_partial_ed`: garde-fou au centième (plus de tolérance), puis, APRÈS la vraie correction, contrôle que chaque DUM de generated_excel (`read_generated_excel_dum_weights`, NOUVEAU) est entièrement réparti entre les partiels au centième — sinon refus avec la liste des écarts (garantit que les deux implémentations de la correction concordent).

**Vérifié (vrais fichiers Excel en scratchpad, vraie `correct_blocage_weights`):** liste de l'utilisateur → partiel blocage = 1318.88 exactement, DUM 14 = 48.88 entier (plus de split), 0 kg non attribué; dialogue et Phase 1 identiques sur les 14 DUMs; lots ED = 1318.88. Régression partiels normaux: sommes exactes, morceaux de split = poids du DUM, aucun micro-morceau.

**Files changed:** `gui/screens/partial_config_dialog.py`, `gui/screens/preparation.py`, `badr_login_test.py`

---

## 2026-11-04 — LTA à la fois BLOCAGE et PARTIEL

**Besoin:** un LTA peut être blocage + partiel (ex: total 1353.1 kg, bloqué 10.22 kg, partiel 1 = 24 kg/1P, partiel 2 = 1329.1 kg/70P). Règle métier: le **plus grand partiel** est toujours fusionné avec les waybills bloqués — son ED existe déjà (créé par l'agent avec les items bloqués), il faut donc le **MODIFIER** (pas le créer) avec la DS série de ce partiel, en n'y ajoutant que les DUMs de ce partiel. Les autres partiels gardent leur traitement normal: exception (le plus petit, ici 24 kg) → DS MEAD sans ED; sinon → création d'ED.

**Problèmes constatés:**
1. `process_lta_folder_ed_only` testait le blocage EN PREMIER et faisait `return` — la config partielle n'était jamais lue: tous les DUMs du LTA étaient ajoutés à l'ED blocage, avec la DS du fichier shipper.
2. Le dialogue de config partielle répartissait les DUMs pour remplir chaque poids de vol. Or le vol fusionné inclut le poids bloqué, absent des DUMs (`generated_excel` B6 ≈ LTA − bloqué): si ce partiel est rempli en premier, il prend 10.22 kg de DUMs appartenant à un autre partiel (vérifié: 1329.1/170.0 au lieu de 1318.9/180.2). Dans le cas exact de l'utilisateur (partiel blocage rempli en dernier) le résultat était correct par hasard.

**Solution — Phase 1 (`badr_login_test.py`):**
- `get_blocage_partial_number(config)`: `blocage_partial_number` enregistré par la GUI, sinon le plus grand partiel en poids.
- `process_lta_folder_ed_only`: config partielle lue avant le choix du workflow. Blocage sans partiel → chemin inchangé. Blocage + partiel → boucle partielle existante (skip exception + checkpoint inchangés), le partiel fusionné passe par `process_blocage_partial_ed`, les autres par `create_etat_depotage_partial`. Garde: le partiel blocage ne peut pas être un partiel d'exception.
- `process_blocage_partial_ed` (NOUVEAU): garde-fou (DUMs du partiel ≈ poids vol − poids bloqué, sinon refus avec message "re-sauvegarder la config partielle"), `correct_blocage_weights`, puis modification avec la DS du partiel.
- `modify_etat_depotage_for_blocage(..., partial_config=None, partial_number=None)`: en mode partiel, lots = `get_dum_lots_for_partial` (DUMs du partiel, morceaux split, poids/positions du config) et références au format partiel (`ref/N` ou `ref/N/k`, même `format1` que `create_etat_depotage_partial`) pour que la Phase 2 les retrouve comme pour tout partiel; DS validée enregistrée dans le partial config (`ds_validated` = checkpoint) au lieu de la ligne 4 du shipper. Paramètres par défaut → comportement blocage simple identique.

**Solution — GUI:**
- `preparation.py`: le bouton "Configurer" passe au dialogue l'état LIVE de la case Blocage et du poids bloqué de la carte (ils ne sont écrits sur disque qu'à "Enregistrer"); poids bloqué invalide → message bloquant.
- `partial_config_dialog.py`: `blocked_weight` (live, sinon lu dans le fichier LTA); bandeau 🔒 dans "Totaux LTA"; helper `_blocage_dum_targets` utilisé par l'aperçu ET la sauvegarde → cible DUMs du plus grand partiel = poids − bloqué; contrôle de la somme des poids hors blocage; le partiel fusionné est étiqueté dans l'aperçu; config enregistre `blocage_partial_number` et `blocage_weight`.

**Phase 2:** aucun changement — le partiel blocage a des lots au format partiel et sa DS dans le partial config, comme tout partiel.

**Vérifié (scripts jetables, vrai code, Selenium stubbé):** répartition avec les chiffres de l'utilisateur (partiel blocage ≈ 1318.9 kg = 1329.1 − 10.22) et cas inversé (avant/après correctif); routage des 4 types de LTA (blocage+partiel → exception ignorée / plus grand MODIFIÉ / autre CRÉÉ; relance avec checkpoint; blocage seul inchangé; partiel seul inchangé); garde-fou qui refuse une config non déduite. Non testé en live sur BADR.

**Files changed:** `badr_login_test.py`, `gui/screens/preparation.py`, `gui/screens/partial_config_dialog.py`, `ai-docs/PROJECT.md`

---

## 2026-11-03 — "Demandes diverses" (case 38): écrire "CARTON" au lieu de "COLIS"

**Demande:** À partir de maintenant, le texte écrit dans le textarea "Demandes diverses" (case 38) doit dire `SOIT {N} CARTON` au lieu de `SOIT {N} COLIS`.

**Fix (`badr_login_test.py`, ÉTAPE 5.4, `fill_declaration_form`):** les 2 occurrences du texte par défaut (`current_text = f"SOIT {carton_value} COLIS ..."`) changées en `CARTON`. Les 2 occurrences de la regex de mise à jour (`re.sub(r'SOIT\s+\d+\s+COLIS', ...)`, utilisée quand une déclaration déjà remplie est rouverte) changées pour reconnaître À LA FOIS `COLIS` et `CARTON` (`r'SOIT\s+\d+\s+(?:COLIS|CARTON)'`) et toujours réécrire `CARTON` — ainsi une déclaration déjà écrite avec l'ancien texte "COLIS" (versions précédentes de l'app) est correctement convertie si elle est rouverte/modifiée, au lieu de rester bloquée sur l'ancien mot.

**Vérifié:** regex testée sur les deux cas (texte existant "...COLIS..." et texte existant déjà "...CARTON...") → les deux produisent bien `SOIT {N} CARTON` avec le compte à jour.

**Files changed:** `badr_login_test.py` (`fill_declaration_form`, ÉTAPE 5.4 — bloc principal + bloc de secours XPath)

---

## 2026-10-29 — Phase 1 ED standard (non-partiel): checkpoint manquant (redoublonnait le travail déjà validé)

**Problem:** User a 4 LTA, dont 2 ont déjà un DS validé (visible en vert dans la colonne "DS Validé" de la GUI, ex: `5888C`) et 2 encore "En attente". En lançant Phase 1 (mode par défaut "Tous les LTAs avec DS"), le log montre `2eme LTA` — déjà validé — repartir depuis zéro dans `📦 CRÉATION ETAT DE DÉPOTAGE`, dupliquant un ED déjà créé sur BADR.

**Root cause:** Le checkpoint ajouté précédemment (voir entrée du 2026-08-20) ne couvre QUE la boucle des **partiels** dans `process_lta_folder_ed_only`. La branche "Standard LTA processing (no partial)" — celle utilisée ici — appelait `create_etat_depotage` sans aucune vérification préalable. Séparément, les cases à cocher de la GUI (`phase1_ed.py`) sont **désactivées** par défaut (mode radio "Tous les LTAs avec DS", pas "Sélection manuelle") — donc même les LTA explicitement décochés dans la table sont ignorés et TOUS sont traités.

**Fix (`badr_login_test.py`, `process_lta_folder_ed_only`, branche standard):** après lecture de `shipper_data`, relit directement la ligne 4 (index 3) du fichier `[LTA]_shipper_name.txt` — exactement le champ que `create_etat_depotage` écrit après un ED réussi (`ds_reference`, ligne ~6108-6123) et que la GUI lit pour peupler "DS Validé" (`gui/utils/file_utils.py:160`, `shipper_data.get('ds_reference')` / ligne 4). Si non vide → `⏭️ LTA DÉJÀ TRAITÉ (checkpoint)`, `return False` (compté "skipped", même convention que le cas "pas de DS MEAD" juste au-dessus) sans recréer l'ED.

Vérifié: confirmé que les LTA "blocage" ne passent jamais par cette branche (branchement séparé + `return` précoce ligne ~11505) donc non affectés. Testé la lecture ligne-4 avec un fichier 4-lignes (validé, ex: `5888C`) → détecté correctement, et un fichier 3-lignes ("En attente") → `None`, non skippé.

**Not addressed (hors scope, à signaler à l'utilisateur):** le mode de sélection GUI par défaut ("Tous les LTAs avec DS") ignore les cases à cocher individuelles — ce n'est pas un bug du script mais un comportement GUI qui, combiné à l'ancien manque de checkpoint, explique le "DÉDOUBLE le travail" rapporté. Le nouveau checkpoint règle le problème côté script indépendamment du mode GUI choisi.

**Files changed:** `badr_login_test.py` (`process_lta_folder_ed_only` — checkpoint branche standard, non-partielle)

---

## 2026-10-14 — Fix: false "MAWB MISMATCH" when a duplicate-download MAWB PDF exists (`script_all_fuzy_match.py`)

**Problem:** A folder with both `MAWB 235-96599543.pdf` and `MAWB 235-96599543[1].pdf` (a second download of the same attachment, auto-suffixed by the OS/browser) was skipped entirely with `⚠️ ALERTE: INCOHÉRENCE MAWB`, even though both PDFs are byte-identical (the existing duplicate-content detector confirms this later in the same run, but only *after* the folder had already been skipped).

**Root causes (two, compounding):**
1. `rename_mawb_pdfs_and_create_bloc_note` extracts the MAWB number straight from the filename via `re.search(r'MAWB\s*(.+)\.pdf', filename)`, which — correctly, since the regex is unbounded — captured `"235-96599543[1]"` *including* the bracket suffix. `validate_mawb_match` only strips dashes/spaces, not bracket/paren markers, so `"235-96599543[1]"` never equaled the Excel-derived `"235-96599543"`.
2. `glob.glob("MAWB*.pdf")` has no guaranteed enumeration order. The pasted log showed Phase 1 and Phase 2 of the *same run* calling this function independently and getting the two files in a *different order* each time — Phase 1 happened to process the clean file first and passed; Phase 2 processed the `[1]` file first and failed. So the bug was flaky on top of being wrong: the same folder could pass or fail depending on glob's non-deterministic directory iteration.

**Fix (`script_all_fuzy_match.py`):**
- New helpers `_has_duplicate_download_marker()` / `_strip_duplicate_download_marker()` (next to `validate_mawb_match`) detect/strip a trailing `" [N]"` or `" (N)"` suffix.
- `rename_mawb_pdfs_and_create_bloc_note`: before the main loop, when more than one `MAWB*.pdf` exists, **deduplicate by the stripped MAWB number** — keep the copy *without* the marker (whichever glob order produced them), print `ℹ️  Fichier MAWB en double ignoré (même référence): '...'` for the one dropped. The extracted `mawb_number` itself is also now passed through `_strip_duplicate_download_marker()` as defense-in-depth, so a lone surviving `[1]`-suffixed copy (no clean counterpart) still compares/renames correctly.
- Verified: unit-tested both helpers, confirmed `validate_mawb_match` now returns `True` for the exact reported pair, and simulated the dedup pass against **both** glob orderings seen in the log (clean-first and `[1]`-first) — both now deterministically keep the clean file and ignore the duplicate.

**Files changed:** `script_all_fuzy_match.py` (`_has_duplicate_download_marker`, `_strip_duplicate_download_marker` NEW; `rename_mawb_pdfs_and_create_bloc_note` updated)

---

## 2026-10-02 — Fix: `send_keys()` could silently fail on Année/Série/Clé fields (ED + Préapurement DS)

**Problem:** User reported the log would print `✓ Année: 2026` but the field in BADR sometimes stayed empty/wrong — `send_keys()` can silently not register (lost focus, PrimeFaces overlay repaint, timing) and the old code never read the field back to confirm, it just printed success unconditionally.

**Fix:** New helper `type_and_verify_field(driver, wait, locator_value, value, field_label, max_attempts=3, by=By.ID)` (placed before `create_etat_depotage`): clears + types the value, then **reads back** `field.get_attribute('value')` and compares; retries (re-clear + re-type) up to 3 times on mismatch before giving up and returning `False`. Supports `By.ID` (default) or `By.XPATH` locators.

Applied to the Année / Série / Clé fields (the exact fields the user saw fail) in:
- `create_etat_depotage` (main ED creation form, `mainTab:form1:*`)
- `create_etat_depotage_partial` (same form, partial LTA flow — identical copy-pasted block, same bug)
- The Préapurement DS form (`mainTab:form3`-area, XPath-located `anneeId`/`serieId`/`cleId`) — same vulnerable pattern, runs once per DUM, so higher exposure than the ED form alone

**Not touched (different screen, out of scope for this report):** the "Modifier une déclaration" search form (`rootForm:_anneeId` etc., ~line 1183) has the same unverified-`send_keys()` pattern but wasn't implicated by the reported log — flagged here for awareness if the same symptom shows up there.

**Files changed:** `badr_login_test.py` (`type_and_verify_field` NEW; `create_etat_depotage`, `create_etat_depotage_partial`, Préapurement DS field-fill blocks updated to use it)

---

## 2026-06-09 — Removed OCR fallback, all paths now use Gemini Vision only

**Change:** Commented out all OCR fallback methods (OCRmyPDF, Tesseract, pdftoppm) and made Gemini Vision the only extraction method for both text-based and image-based PDFs.

**Reason:** Simplified extraction flow by relying exclusively on Gemini Vision AI for better accuracy and consistency.

**Modified flow:**

1. **Text-based PDF** → Try text extraction → If fails → **Gemini Vision** (was: OCR)
2. **Image-based PDF** → **Gemini Vision** only (was: Gemini Vision → OCR fallback)
3. **Bloc PDFs** → Same Gemini Vision path (was: special OCR handling)

**Changes made:**

- Line ~1363: Text extraction failure now calls `extract_vision_meta()` instead of `extract_shipper_name_ocr()`
- Lines ~1378-1397: Commented out entire OCR fallback section (OCRmyPDF, Tesseract, bloc PDF special handling)

**Impact:** All shipper name extraction now requires Gemini API to be configured. OCR tools are no longer used.

**Files changed:** `script_all_fuzy_match.py`

---

## 2026-06-09 — Clarified Gemini Vision call flow in shipper extraction

**Question:** Where exactly is Gemini Vision called in `script_all_fuzy_match.py`?

**Answer:** Gemini Vision is called at **line 1370** in the `extract_shipper_name()` function:

```python
vision_shipper = extract_vision_meta(pdf_path)
```

**When it's called:**

- ONLY for **image-based PDFs** (scanned documents without extractable text)
- ONLY if Gemini API setup succeeds
- It's the **first method** tried for image-based PDFs, before falling back to OCR tools

**Processing flow:**

1. If text-based PDF → Try text extraction → If fails, use OCR (no Gemini Vision)
2. If image-based PDF → **Try Gemini Vision first** → If fails, use OCR fallback

**Why it wasn't called in 7eme LTA:**

- PDF was identified as text-based
- Text extraction failed → fell back to Tesseract OCR
- Tesseract found high-confidence fuzzy match (score 1.000) → used directly
- No AI verification needed

**Files referenced:** `script_all_fuzy_match.py` (Lines 1346-1438, specifically line 1370)

---

## 2026-05-06 — Fix split DUM positions inconsistency in multi-smallest exception

**Problem:** In multi-smallest exception case, `partial['positions']` for the big partial was correctly computed as `total - sum(user_exception_positions)` (e.g. 1618-2-2=1614), but the individual DUM entries inside that partial still used proportionally-derived positions from Excel data (summing to only 1561). Root cause: the split DUM pieces for exception partials kept their proportional positions (24 and 33 from weight-based split), not the user-declared values (2 each). So big-partial DUM pieces also had the wrong remainder.

**Fix:** In `_save_config`, after `exception_positions_map` is finalized, added a correction loop:

1. For each split DUM that has exception pieces: `total_orig = sum of all piece positions` (proportional)
2. Exception pieces → override to user-declared value from `exception_positions_map`
3. Big-partial pieces → `total_orig - sum(user_exception_positions)` (distributed by weight if multiple big pieces)
4. Applied to both `partials[i]['dums']` entries AND `split_dums[dum_str]['splits']` entries

**Result:** DUM-level positions now consistent with partial-level positions (both layers sum correctly).

**Files changed:** `gui/screens/partial_config_dialog.py` (`_save_config`)

---

## 2026-05-06 — Multi-smallest-partial exception support

**Task:** Support exception cases where more than one partial is airport-cleared (DS MEAD), e.g. partial 1 = 27kg/2p, partial 2 = 1222kg/72p, partial 3 = 26kg/2p → partials 1 and 3 are both exception (no ED), partial 2 gets the ED; DS préapurement generates one DS MEAD lot per small partial.

**Root cause:** Config stored a single `smallest_partial_number`; detection was limited to `num_partials in (2, 3)` with one small partial; ED-skip and DS MEAD generation all assumed only one exception partial.

**Changes:**

- `gui/screens/partial_config_dialog.py`:
  - `_setup_ui`: Replaced hardcoded single airport-reference/positions fields with a dynamic `exception_inputs_frame` container; added `self.exception_partial_vars = {}` dict.
  - New `_update_exception_frame(small_partial_nums)`: Rebuilds per-partial ref/pos Entry rows in the dynamic container; restores saved values from existing config on reload.
  - `_update_distribution_preview`: Detection now `any(w < max_dum_weight)` (no `num_partials in (2,3)` restriction); calls `_update_exception_frame`; passes single-smallest to legacy exception distribution, or None for multi-smallest (uses normal sequential distribution, split_dums captures boundaries).
  - `_save_config`: Same detection update; validates per-partial airport refs/positions; saves new keys `smallest_partial_numbers` (list), `smallest_partial_airport_references` (dict), `smallest_partial_positions_map` (dict); keeps old single-value keys for backward compat.

- `badr_login_test.py`:
  - `get_dum_preapurement_lots`: Reads `smallest_partial_numbers` list + dicts (falls back to legacy single keys); for each split DUM, iterates all pieces — pieces belonging to small partials get DS MEAD lots with that partial's airport reference; big-partial pieces get Depotage lots using the co-split small partial's airport reference as base.
  - `process_lta_folder_ed_only`: Reads `smallest_partial_numbers` list (fallback to single); removes old `partials_count in (2,3)` guardrail; skips ED for ALL partials in the list; per-partial print of ref/positions from new dict fields.
  - `fill_declaration_form` PDS block (lot_idx > 0 section): Replaced hardcoded `Depotage(05)` with per-lot `is_exception_smallest` check → DS MEAD(03) or Depotage(05) accordingly.

**Backward compat:** Old single-partial exception configs continue to work via fallback reading of legacy keys.

**Files changed:** `gui/screens/partial_config_dialog.py`, `badr_login_test.py`

---

## 2026-04-23 — Mapped full MAWB→Shipper extraction chain for Playwright/JS port

**Task:** Produce a complete function map of `script_all_fuzy_match.py` covering the pipeline from MAWB extraction to shipper name output, for porting to a Playwright JS app. Exports needed: `known_companies.json` + Gemini API key.

**Call chain documented:**

- Entry: `rename_mawb_pdfs_and_create_bloc_note` → calls `extract_mawb_from_generated_excel`, `validate_mawb_match`, `compress_pdf_if_needed`, `extract_shipper_name`, `create_bloc_note`
- `extract_shipper_name` → `is_pdf_text_based`, then text path (`extract_shipper_name_text_based`) or image path (`process_multi_page_pdf_with_detection` / `extract_shipper_name_ocr`)
- Multi-page: `find_shipper_page_text_based` → `extract_from_specific_page_text` → `find_best_company_match`
- OCR methods: `try_ocrmypdf`, `try_tesseract_pdftoppm`, `try_pdftotext` → all feed `process_extracted_text`
- Text helpers: `clean_extracted_text`, `clean_company_name`, `might_be_company`, `is_airline_or_system_text`, `clean_for_matching`, `extract_key_terms`
- AI path: `setup_gemini_api` → `verify_shipper_with_gemini` → `_gemini_response_text`
- DB: `load_companies_database`, `save_companies_database`, `add_company_to_database`
- Output: `create_bloc_note` → writes `<lta>.txt` + `<lta>_shipper_name.txt`

**Files read:** `script_all_fuzy_match.py`

---

## 2026-04-23 — Lot retry: stay on lot page instead of going to accueil

**Problem:** When a lot creation step failed (e.g. type contenant autocomplete error), the inner error handler called `return_to_home_after_error(driver)` before raising `"Lot creation error"`. The outer retry loop then tried to click `btn_new_lot` on the accueil page — which doesn't exist — causing all 3 retry attempts to fail immediately after the first real error.

**Root cause:** Inner `except` handlers (lines ~4823–5129 in `fill_declaration_form`) went to accueil proactively, but the outer `while retry_attempt < MAX_LOT_RETRIES` catch expected the driver to still be on the ED lot list page.

**Solution:**

1. Removed `return_to_home_after_error(driver)` from all inner lot creation error handlers. They now only do `driver.switch_to.default_content()` (best-effort) then raise. The outer catch decides when to go home.
2. Updated outer Selenium/unknown error recovery to do `driver.refresh()` + re-enter iframe (instead of just switching frames), so `btn_new_lot` is visible on the next attempt.
3. `return_to_home_after_error` is still called only in the outer catch's `else` branch (after max retries exhausted), which then `return False` to trigger full ED recreation.

**Files changed:** `badr_login_test.py` (`fill_declaration_form` lot creation loop)

---

## 2026-04-21 — Preparation: write shipper name to generated_excel H1

**Problem:** Shipper name entered in the Preparation screen was saved to txt files only; the `generated_excel*.xlsx` for each LTA did not carry the shipper name, making it harder to identify shipments from the Excel alone.

**Solution:** Appended logic at the end of `update_lta_shipper_name` in `gui/utils/file_utils.py`. After updating the txt files it:

1. Resolves the LTA subfolder: `os.path.join(lta_folder_path, folder_name)`
2. Globs for `generated_excel*.xlsx` in that subfolder
3. Opens each match with openpyxl, writes the shipper name to cell `H1`, saves, and closes
4. All errors are caught and logged as warnings — the main save flow is never blocked

**Files changed:** `gui/utils/file_utils.py` (`update_lta_shipper_name`)

---

## 2026-04-17 — Email: add Shipper, DS MEAD, and Blocage info

**Problem:** LTA completion emails contained no shipper name, DS MEAD signé série, or blocage flag — recipients couldn't identify the shipment at a glance.

**File structure used:**

- `<lta_name>_shipper_name.txt` (e.g. `6eme_LTA_shipper_name.txt`) — Line 1 = Expéditeur
- `<lta_name>.txt` (e.g. `6eme LTA.txt`) — Line 5 = "blocage" if blocage LTA; Line 8 = DS MEAD série

**Solution:** Added `_read_lta_info(lta_folder_path, lta_name)` helper in `email_sender.py` that reads both files (0-indexed: lines[4] and lines[7]) and builds a dict `{shipper, ds_mead, is_blocage}`.

Changes to `send_excel_via_email`:

- Subject: `[BLOCAGE] LTA Complet - ...` / `[BLOCAGE][ERREUR DUM] LTA Complet - ...` (blocage prefix always first)
- Body: `Expéditeur`, `DS MEAD Signé`, and blocage warning line injected after the DUM count line, before the error list
- No signature changes — `lta_folder_path` already present in all call sites

**Files changed:**

- `gui/utils/email_sender.py` (`_read_lta_info` NEW, `send_excel_via_email` subject + body updated)

---

## 2026-04-10 — Auto-resume: skip completed DUMs + LTA-level retry for Selenium errors

**Problem:** When an LTA finished with some DUMs failed (Selenium errors), the user had to manually delete their rows from generated_excel and restart the app. Business logic errors (BADR validation refusals) were treated the same as Selenium errors — both could only be addressed manually.

**Sub-fix (same session):** Skip logic was silent on pass 0 — impossible to know if skipping was occurring. Also, `generated_excel_path_skip` was resolved once before the pass loop, so if the file was created mid-run (first successful DUM creates the file), later DUMs in the same pass still saw `None`. Fixed by:

- Moving the glob resolution to inside the pass loop (fresh each pass)
- Adding a pre-scan before processing: prints `⏭️ DUMs déjà traités`, `🚫 DUMs erreur métier`, `🔄 DUMs à traiter` — visible on every pass including pass 0
- Added `ℹ️ Aucun generated_excel trouvé` message when file is missing

**Solution:**

### Skip logic (resume from checkpoint)

At the start of each DUM iteration in `process_lta_folder_dum_only`, the app reads the DUM's cell in generated_excel:

- **Valid series** — DUM already processed successfully → skip (count as success)
- **"error_biz"** — business logic error, not retryable → skip (count as failed)
- **"error" or empty** — process / retry this DUM

This replaces the manual "delete row + restart" workflow.

### LTA-level retry (up to 3 passes)

The DUM loop is now wrapped in `for lta_pass in range(MAX_LTA_PASSES=3)`. After each pass:

- If no Selenium failures remain → break early (skip remaining passes)
- If Selenium failures remain and passes left → print warning, wait 5s, re-run the DUM loop (skip logic automatically skips already-done DUMs)

### Error type distinction

- **`mark_dum_as_error_in_excel`**: new `is_business_error=False` parameter. Writes `"error"` (retryable) or `"error_biz"` (not retryable). Guards against downgrading `"error_biz"` → `"error"`.
- **`_read_dum_excel_status(excel_path, dum_number)`**: new helper that returns `"done"`, `"error"`, `"error_biz"`, or `None`.
- **BADR validation failure** (line ~10448 inside `fill_declaration_form`): now calls `mark_dum_as_error_in_excel(..., is_business_error=True)` → writes `"error_biz"`.
- **Outer catch in `process_lta_folder_dum_only`**: passes `is_business_error=not is_selenium` so Selenium exhausted errors stay retryable and logical errors are permanently skipped.

**Files changed:**

- `badr_login_test.py` (`_read_dum_excel_status` NEW, `mark_dum_as_error_in_excel` updated, `fill_declaration_form` BADR validation call updated, `process_lta_folder_dum_only` DUM loop rewritten with LTA-pass retry + skip logic)

---

## 2026-04-09 — Email: attach MAWB PDF + flag ERROR DUMs in subject

**Problem:** Post-Phase-2 email only attached the generated_excel. No PDF, no indication of failed DUMs.

**Solution:**

- Added `_find_mawb_pdf(lta_folder_path, lta_name)` — locates `<lta_name>/<lta_name> - <MAWB>.pdf` (fallback: any .pdf in subfolder).
- Added `_get_error_series_from_excel(excel_file_path)` — reads column C of the Summary sheet at rows 12, 19, 26, 33… and returns values containing `"error"`.
- `send_excel_via_email` now accepts `lta_folder_path`, attaches the PDF when found, and prefixes subject with `[ERREUR DUM]` when errors exist. Body lists each errored serie.
- `send_excel_after_lta_completion` signature extended with `lta_folder_path=None`.
- Both call sites in `badr_login_test.py` updated to pass `lta_folder_path=os.path.dirname(lta_folder_path)`.

**Files changed:**

- `gui/utils/email_sender.py`
- `badr_login_test.py` (lines ~10853, ~11694)

---

## 2026-04-09 — Excel corruption fix: atomic write pattern

**Problem:** `close_excel_file()` kills Excel via `proc.terminate()` which can interrupt an in-progress autosave, leaving a partially-written `.xlsx` (invalid ZIP). Subsequent calls to `save_dum_series_to_excel` detected the corruption with a ZIP pre-check and failed all 5 retries without ever fixing the file.

**Root cause:** the retry loop validated the original before writing — but never repaired it, so every attempt raised `BadZipFile`. The original file stayed corrupt indefinitely.

**Solution:** Extracted `_write_excel_cell_atomic(excel_path, cell_position, value)` helper that:

1. On entry, checks if the original is a valid ZIP. If corrupt and a `.bak` exists → auto-restores from backup.
2. Creates `.bak` of the (now valid) original.
3. Copies original → `.tmp`, writes the cell, saves `.tmp`.
4. `os.replace(.tmp, original)` — atomic swap; original untouched if step 3 crashes.
5. Deletes `.bak` on clean success (`.bak` is preserved if replace fails for recovery on next call).

`save_dum_series_to_excel` reduced to 3 retries, calls `close_excel_file` once before the loop, then delegates to the helper.  
`mark_dum_as_error_in_excel` rewritten to call the same helper (removed legacy direct wb open/save).

**Files changed:**

- `badr_login_test.py` (functions: `_write_excel_cell_atomic` NEW, `save_dum_series_to_excel`, `mark_dum_as_error_in_excel`)

---

## 2026-04-10 — Fix: openpyxl rejects `.tmp` extension in atomic write helper

**Problem:** `_write_excel_cell_atomic` named the temp file `<file>.xlsx.tmp`. openpyxl validates file extensions and raises `"does not support .tmp file format"` when `load_workbook` is called on it, causing all 3 retries to fail with the same error and the série never being written to generated_excel.

**Root cause:** `tmp_path = excel_path + ".tmp"` appends `.tmp` to the full filename, so openpyxl sees an unsupported extension.

**Fix:** Changed to `tmp_path = base + "_tmp" + ext` (e.g. `file_tmp.xlsx`) so the temp file always has a valid xlsx extension that openpyxl accepts.

**Files changed:**

- `badr_login_test.py` (`_write_excel_cell_atomic`: one-line temp path fix)

---

## 2026-04-10 — Fix: error email not triggered for large LTAs (>20 DUMs)

**Problem:** When an LTA with 36 DUMs had errors, the completion email was sent as a normal email with no `[ERREUR DUM]` prefix or error body — even though DUMs were correctly marked "error" in the Excel.

**Root cause (two bugs in `_get_error_series_from_excel`):**

1. **Safety cap of 20** — `if dum_number > 20: break` meant DUMs 21–36 were never inspected. Any errors in those DUMs were invisible to the email logic.
2. **Aggressive `None` break** — `if value is None: break` stopped the loop at the first empty cell, so a single DUM with a missing series (e.g. due to a prior write failure) silently hid all subsequent errors.

**Fix:** Replaced the loop in `_get_error_series_from_excel`:

- Cap raised to 100 (supports LTAs up to 100 DUMs).
- `None` no longer immediately breaks; instead counts consecutive empty cells and breaks only after 3 in a row.

**Files changed:**

- `gui/utils/email_sender.py` (`_get_error_series_from_excel`)
