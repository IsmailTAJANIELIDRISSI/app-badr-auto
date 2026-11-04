# TASKS.md

Last updated: 2026-11-04

## Recently Fixed

- **2026-11-04 — Champ "Nb Articles Bloqués" dans la section blocage (onglet Préparation).** Obligatoire si blocage + partiel. Enregistré en ligne 14 du fichier LTA, et utilisé en Phase 1 pour "Nombre total Contenant(s)" = articles bloqués + positions du plus grand partiel. Le décompte des lots de l'ED BADR sert de contrôle et de repli. Files: `gui/screens/preparation.py`, `gui/utils/file_utils.py`, `badr_login_test.py`. See PROGRESS.md.
- **2026-11-04 — Blocage + partiel: onglet Quantités de l'ED modifié (E2800435).** Avant de créer les lots DUM, l'app lit les lots bloqués restants et corrige si besoin "Nombre total Contenant(s)" = contenants bloqués + positions du partiel (ex: 19 + 1829 = 1848), ainsi que "Poids brut total" = poids du partiel (1329.1). Si E2800435 apparaît quand même à VALIDER, l'app reprend la valeur calculée par BADR (si cohérente) et revalide une fois. File: `badr_login_test.py`. Testé sur une maquette, non testé en live BADR. See PROGRESS.md.
- **2026-11-04 — Blocage + partiel: Phase 1 créait l'ED du plus grand partiel au lieu de MODIFIER l'ED blocage.** Phase 1 ne lisait le blocage que dans la ligne 5 du fichier LTA (pas encore sauvegardée). Elle le lit maintenant aussi dans la config partielle (`resolve_blocage_for_partial`). La DS du blocage = DS du plus grand partiel; la DS de l'onglet Préparation est facultative, et refusée si elle est différente. La sauvegarde de la config écrit le blocage dans le fichier LTA, remplit la DS de la carte et ne perd plus les saisies non sauvegardées. "Sauvegarder" vérifie que la carte concorde avec la config. Files: `badr_login_test.py`, `gui/screens/preparation.py`, `gui/screens/partial_config_dialog.py`. Non testé en live BADR. See PROGRESS.md.
- **2026-11-04 — Blocage + partiel: répartition des DUMs exacte au centième.** Le dialogue arrondissait les poids à 0.1 kg et répartissait les DUMs AVANT la correction blocage (DUM 14 = 59.1 contenait encore les 10.22 kg bloqués → somme affichée 1318.9 au lieu de 1318.88, faux split "14/1"). Corrigé: calculs à 2 décimales partout, le dialogue applique la même correction blocage que Phase 1 (DUM 14 → 48.88, entier, pas de split), la sauvegarde refuse une répartition inexacte, Phase 1 vérifie DUMs partiel = poids partiel − bloqué (au centième) et contrôle chaque DUM contre le generated_excel corrigé, lots ED p_brut à 2 décimales. Files: `gui/screens/partial_config_dialog.py`, `badr_login_test.py`. ⚠️ Ré-ouvrir et re-sauvegarder la config partielle de ce LTA. Non testé en live BADR. See PROGRESS.md.
- **2026-11-04 — LTA blocage + partiel.** Le plus grand partiel est fusionné avec le blocage: son ED existant est MODIFIÉ (DS du partiel, seuls ses DUMs, lots au format partiel), les autres partiels suivent le traitement partiel (DS MEAD si exception, sinon ED créé). Le dialogue de config partielle déduit le poids bloqué des DUMs du plus grand partiel (bug: il volait ce poids à un autre partiel quand le partiel blocage était rempli en premier). Phase 1 refuse une config non déduite. Files: `badr_login_test.py`, `gui/screens/preparation.py`, `gui/screens/partial_config_dialog.py`. ⚠️ Une config partielle déjà sauvegardée pour un LTA blocage doit être ré-ouverte et re-sauvegardée. Non testé en live BADR. See PROGRESS.md.

- **2026-11-03 — "Demandes diverses" (case 38): CARTON au lieu de COLIS.** Le textarea écrit désormais `SOIT {N} CARTON ...` (au lieu de `COLIS`). La regex de mise à jour reconnaît aussi l'ancien mot "COLIS" sur une déclaration déjà remplie et la convertit en "CARTON". File: `badr_login_test.py` (`fill_declaration_form`, ÉTAPE 5.4). See PROGRESS.md.

- **2026-10-29 — Checkpoint manquant pour les LTA standard (non-partiel) en Phase 1 ED.** `process_lta_folder_ed_only` relançait `create_etat_depotage` pour un LTA dont le DS était déjà validé (ligne 4 du fichier shipper déjà remplie = "DS Validé" vert dans la GUI), dupliquant le travail sur BADR. Le checkpoint ajouté le 2026-08-20 ne couvrait que les partiels. Fix: lecture directe de la ligne 4 avant `create_etat_depotage` → skip si déjà validé. Noter aussi: le mode GUI par défaut ("Tous les LTAs avec DS") ignore les cases à cocher — comportement séparé, non traité ici. File: `badr_login_test.py`. See PROGRESS.md.

- **2026-10-14 — Faux "MAWB MISMATCH" avec un fichier MAWB en double téléchargement.** Un `MAWB *.pdf` téléchargé deux fois (ex: `... [1].pdf`) faisait échouer tout le dossier: le `[1]` polluait le numéro MAWB extrait du nom de fichier, ET `glob.glob()` renvoyait l'ordre des fichiers de façon non déterministe (le même dossier passait ou échouait selon l'ordre). Fix: dédoublonnage explicite (garde la copie sans marqueur `[N]`/`(N)`) + strip du marqueur à l'extraction. File: `script_all_fuzy_match.py`. See PROGRESS.md.

- **2026-10-02 — Vérification de saisie Année/Série/Clé (ED + Préapurement DS).** Nouveau `type_and_verify_field()` relit le champ après `send_keys()` et réessaie si la valeur ne correspond pas (au lieu d'afficher "✓" sans vérifier). Appliqué dans `create_etat_depotage`, `create_etat_depotage_partial`, et le formulaire Préapurement DS (par DUM). Not applied to: la recherche "Modifier une déclaration" (`rootForm:_anneeId`) — même pattern non vérifié, hors scope du bug rapporté. File: `badr_login_test.py`. See PROGRESS.md.
- **2026-08-20 — Phase 1 ED partiels: checkpoint reprise.** `process_lta_folder_ed_only` saute désormais un partiel dont `ds_validated`/`signed_series` existe déjà (ED déjà créé + validé) → après crash, ne recrée que les partiels manquants. File: `badr_login_test.py`. See PROGRESS.md.
- **2026-08-20 — `badr_playwright/` flux complet + checkpoint.** Ajout onglet Articles → upload `Sheet <N>.xlsx` → SAUVEGARDER → VALIDER ×2 (Commerce électronique = Oui) dans `src/finalizeDeclaration.js`. Checkpoint de reprise (`utils/checkpoint.js` → `checkpoint.json`, clé `LTA::serieRaw`): les DUMs réussis sont ignorés au rerun. ⚠️ VALIDER soumet la déclaration. Non testé end-to-end. See PROGRESS.md.
- **2026-08-20 — Nouveau projet `badr_playwright/`.** Script Playwright/Node séparé: pour chaque série de DUM (lue dans `generated_excel`), ouvre la déclaration via *Modifier une déclaration* et annexe LTA + `mnN.pdf` (onglet Documents). Parseur Excel testé (6eme LTA → 16 DUMs). Non testé end-to-end vs BADR. TODO: clic final "Enregistrer" éventuel. See PROGRESS.md.
- **2026-08-19 — [PROVISOIRE/DEBUG] Stop après SAUVEGARDER, saute l'onglet Documents.** Flag `DEBUG_STOP_AFTER_SAVE = True` dans `badr_login_test.py`: après sauvegarde réussie, extrait la série → écrit dans `generated_excel` → retour accueil, sans faire l'annexation des documents. ⚠️ Remettre `False` pour le flux complet. Incertitude: la série n'est peut-être pas encore disponible juste après le SAUVEGARDER (à valider en run). Helpers `_extract_and_save_dum_series` / `_return_to_home_for_next_dum` (ÉTAPE 12/13 refactorées). File: `badr_login_test.py`. See PROGRESS.md.
- **2026-08-19 — [PROVISOIRE] Navigateur Edge → Chrome.** Bascule réversible via l'interrupteur `BROWSER` en haut de `badr_login_test.py` (défaut `"chrome"`; mettre `"edge"` ou `BROWSER=edge` dans `.env` pour revenir). Toute la logique navigateur passe par des variables `BROWSER_*`. ⚠️ À rebasculer sur Edge quand la période provisoire est finie. File: `badr_login_test.py`. See PROGRESS.md.
- **2026-08-03 — Blocage 0.01 kg adjustment skipped.** `correct_blocage_weights` never applied a subtraction of exactly 0.01 kg (loop guard `<= 0.01` broke immediately, then falsely printed success). Now uses a 0.005 epsilon so a 0.01 kg correction is applied (e.g. last DUM 125.32 → 125.31, sum 784.31). File: `badr_login_test.py`. See PROGRESS.md.
- **2026-07-27 — Missing `generated_excel*.xlsx` now flagged.** `validate_required_files_present` explicitly checks for `generated_excel*.xlsx`; if absent it's reported as a hard "Fichiers Manquants" error (and no longer silently skipped when `expected_count == 0`). File: `script_all_fuzy_match.py`, GUI banner via `gui/utils/file_utils.py`. See PROGRESS.md.
- **2026-07-24 — PDS single-DUM division on blocage LTAs.** Préapurement DS no longer divides a single-DUM *blocage* LTA into `/1` + `/2`; blocage LTAs now enter only their real DUM(s). Normal single-DUM LTAs keep /1 + /2. File: `badr_login_test.py` (ÉTAPE PDS block). See PROGRESS.md.

## Current Incomplete / Broken Areas

1. **Phase 2 CLI selection bug in `badr_login_test.py`**
   - In the Phase 2 GUI-mode branch, `selected_lta_indices == "all"` assigns `s = lta_folders` instead of `folders_to_process = lta_folders`.
   - Impact: when Phase 2 is launched with `"all"` from GUI/CLI, selected folders can be empty or inconsistent.
   - File: `badr_login_test.py`

2. **`ScriptManager.run_phase2()` has unfinished folder-filtering flow**
   - Variables `moved_folders` / `temp_dir` are declared for parity with Phase 1, but no folders are moved in `run_phase2`.
   - Logic block checks `if moved_folders:` although it is never populated.
   - Impact: confusing behavior and misleading maintenance surface; selection currently depends only on index arguments.
   - File: `gui/utils/script_manager.py`

3. **Duplicate function name `read_lta_file` in `file_utils.py`**
   - The module defines `read_lta_file(file_path)` and later redefines `read_lta_file(lta_folder_path, folder_name)`.
   - Python keeps only the second definition at runtime.
   - Impact: easy source of regressions and wrong assumptions for future edits.
   - File: `gui/utils/file_utils.py`

4. **Large monolithic automation script with mixed responsibilities**
   - `badr_login_test.py` combines browser lifecycle, parsing, ED flow, DUM flow, partial/blocage/exception logic, persistence, and CLI routing in one file.
   - Impact: high change risk and low testability.
   - File: `badr_login_test.py`

5. **No automated regression suite for core business paths**
   - Existing tests are ad-hoc (`test_distribution.py`, `test_integration.py`) and not enforcing end-to-end contracts.
   - Missing durable tests for: standard LTA, blocage, partial normal split, partial exception 2/3/4 vols, E2800124 handling.
   - Files: test scripts + core modules

6. **Runtime auto-update (`git pull --autostash`) in production paths**
   - Executed in `gui/main.py` and `script_all_fuzy_match.py`.
   - Impact: operational risk (network dependency, unpinned updates, local-state drift at runtime).
   - Files: `gui/main.py`, `script_all_fuzy_match.py`

7. **Config and credential handling still partially hardcoded**
   - Username defaults (`BK707345`) and multiple constants are embedded in code.
   - Password loaded from `.env` but mixed with GUI placeholders/legacy assumptions.
   - Impact: deployment friction and security/maintainability concerns.
   - Files: `gui/screens/phase1_ed.py`, `gui/screens/phase2_dum.py`, `badr_login_test.py`, `.env.example`

8. **Selector fragility against BADR UI changes**
   - Heavy reliance on brittle IDs/XPaths and exact labels.
   - Impact: frequent breakage after BADR frontend updates.
   - File: `badr_login_test.py`

9. **Documentation drift risk**
   - Multiple architecture docs exist (`context.md`, `FULL_APP_CONTEXT_FOR_AI.md`, now `ai-docs/*`) and can diverge.
   - Impact: onboarding confusion and stale operational guidance.

10. **Build artifacts folders exist but are empty in current snapshot**
    - `build/` and `dist/` are present but contain no tracked artifacts in this workspace state.
    - Impact: unclear release reproducibility unless `build_exe.py` is run locally.
    - Files: `build_exe.py`, `MedAfrica.spec`, folders `build/`, `dist/`

---

## Logical Next Steps

1. **Fix Phase 2 selection bug first**
   - Correct `"all"` assignment in `badr_login_test.py` Phase 2 branch.
   - Add a small guard log that prints final `folders_to_process` before execution.

2. **Clean `run_phase2` selection mechanics**
   - Either fully implement temp-folder filtering (like Phase 1) or remove dead `moved_folders` branches.
   - Keep one canonical selection strategy.

3. **Resolve `file_utils` function shadowing**
   - Rename one `read_lta_file` variant (`read_lta_file_dict` vs `read_lta_file_lines`) and update call sites.
   - Add docstrings with explicit return types.

4. **Stabilize business-rule exit semantics**
   - Keep business errors (e.g. E2800124) non-retriable.
   - Extend explicit error-type signaling for other deterministic validation failures.

5. **Add contract tests for data files**
   - Unit tests for shipper txt, LTA txt, and partial JSON read/write round-trips.

6. **Email config UI**
   - Add a small settings dialog in the GUI so operators can configure SMTP server, credentials, and recipient list without editing `config/email_config.json` manually.

7. **Delete dead screen files**
   - `gui/screens/partial_config_dialog_check_first_dum.py` and `partial_config_dialog_old_logic.py` are superseded — safe to remove.
   - Validate line-index invariants and backward compatibility.

8. **Add targeted flow tests (mocked where needed)**
   - Distribution/preapurement lot generation tests:
     - standard
     - partial split
     - exception 2/3 vols
     - multi-vol (4+) non-exception
   - Verify lot reference format and split_id sequencing.

9. **Separate automation concerns incrementally**
   - Extract modules from `badr_login_test.py`:
     - `browser_session.py`
     - `ed_flow.py`
     - `dum_flow.py`
     - `partial_logic.py`
     - `io_contracts.py`
   - Keep behavior unchanged while reducing blast radius.

10. **Harden runtime update policy**
    - Gate auto-update behind config/env flag.
    - Add explicit logging when skipped/failed, without impacting workflows.

11. **Centralize configuration**
    - Move repeated constants and defaults to `config/settings.json` + env.
    - Avoid scattered literals for credentials, URLs, and fixed IDs where possible.

12. **Keep AI docs synchronized**
    - Treat `ai-docs/PROJECT.md` + `STACK.md` as canonical engineering docs.
    - Update `PROGRESS.md` at each fix with: problem, solution, files changed.
