"use strict";

const APP_VERSION = 1;
const STORAGE_KEY = "kanji-dragon-autosave-v1";
const THEME_STORAGE_KEY = "kanji-den-theme-v1";
const JP_FONT_STORAGE_KEY = "kanji-den-jp-font-v1";
const BROWSER_BACKUP_DB_NAME = "kanji-den-browser-backup-v1";
const BROWSER_BACKUP_STORE_NAME = "deckBackup";
const BROWSER_BACKUP_KEY = "currentDeck";
const THEME_OPTIONS = ["forest", "sakura", "fire", "deep", "guardian", "black", "parchment", "berry", "mystic", "burnt", "sunshine", "ivory"];
const LEGACY_THEME_ALIASES = {
  pink: "berry",
  yellow: "sunshine",
  orange: "burnt",
  blue: "deep",
  purple: "mystic",
  vampire: "fire"
};
const JAPANESE_FONT_OPTIONS = {
  "yu-mincho": "\"Yu Mincho\", \"Hiragino Mincho ProN\", \"MS PMincho\", serif",
  "yu-gothic": "\"Yu Gothic UI\", \"Yu Gothic\", \"Meiryo\", sans-serif",
  meiryo: "\"Meiryo\", \"Yu Gothic UI\", \"Segoe UI\", sans-serif",
  "biz-udmincho": "\"BIZ UDPMincho\", \"Yu Mincho\", \"Hiragino Mincho ProN\", \"MS PMincho\", serif"
};
const DEFAULT_JP_FONT = "yu-mincho";
const REVIEW_RETENTION = 0.9;
const FSRS_DECAY = -0.5;
const FSRS_FACTOR = Math.pow(REVIEW_RETENTION, 1 / FSRS_DECAY) - 1;
const WORD_CATEGORIES = ["on", "kun", "irregular"];
const MASTERY_LEVELS = ["egg", "hatchling", "chick", "chicken", "cockatrice", "dragon"];
const COCKATRICE_ICON_PATH = "assets/icons/cockatrice-icon.png";
const LEECH_ICON_PATH = "assets/icons/leech-icon.png";
const LEECH_THRESHOLD = 5;
const LOOKUP_SCRIPT_ROOT = "lookup";
const LOOKUP_BUCKET_ROOT = `${LOOKUP_SCRIPT_ROOT}/word-buckets`;
const LOOKUP_RESULT_LIMIT = 48;
const MASTERY_LABELS = {
  egg: "egg",
  hatchling: "hatchling",
  chick: "chick",
  chicken: "chicken",
  cockatrice: "cockatrice",
  dragon: "dragon"
};
const GRADE_CONFIG = {
  again: { score: 1 },
  hard: { score: 2 },
  good: { score: 3 },
  easy: { score: 4 }
};
const MASTERY_EMOJIS = {
  egg: "\u{1F95A}",
  hatchling: "\u{1F423}",
  chick: "\u{1F425}",
  chicken: "\u{1F414}",
  cockatrice: "",
  dragon: "\u{1F432}"
};
const KANJI_SEARCH_CHAR_PATTERN = /[\u3400-\u4DBF\u4E00-\u9FFF\uF900-\uFAFF]/u;
const JAPANESE_SEARCH_CHAR_PATTERN = /[\u3040-\u30ff\u3400-\u4DBF\u4E00-\u9FFF\uF900-\uFAFF]/u;
const HAN_SEARCH_GLOBAL_PATTERN = /[\u3400-\u4DBF\u4E00-\u9FFF\uF900-\uFAFF]/gu;
const SEARCH_SPLIT_PATTERN = /[\s,、，]+/;

const dom = {
  deckSummaryText: document.querySelector("#deck-summary-text"),
  statSource: document.querySelector("#stat-source"),
  statKanji: document.querySelector("#stat-kanji"),
  statCards: document.querySelector("#stat-cards"),
  statDue: document.querySelector("#stat-due"),
  statSave: document.querySelector("#stat-save"),
  statusMessage: document.querySelector("#status-message"),
  importInput: document.querySelector("#json-import"),
  txtImportInput: document.querySelector("#txt-import"),
  exportButton: document.querySelector("#export-json"),
  openFileDeckButton: document.querySelector("#open-file-deck"),
  saveFileDeckButton: document.querySelector("#save-file-deck"),
  saveFileAsDeckButton: document.querySelector("#save-file-as-deck"),
  fileAccessNote: document.querySelector("#file-access-note"),
  bulkImportTemplateButton: document.querySelector("#bulk-import-template"),
  exportAllKanjiDataButton: document.querySelector("#export-all-kanji-data"),
  exportSelectedKanjiDataButton: document.querySelector("#export-selected-kanji-data"),
  exportFilteredKanjiDataButton: document.querySelector("#export-filtered-kanji-data"),
  bulkImportPreview: document.querySelector("#bulk-import-preview"),
  themeSelect: document.querySelector("#theme-select"),
  jpFontSelect: document.querySelector("#jp-font-select"),
  newDeckButton: document.querySelector("#new-deck"),
  demoDeckButton: document.querySelector("#load-demo"),
  restoreBackupButton: document.querySelector("#restore-backup"),
  startReviewButton: document.querySelector("#start-review"),
  reviewEmpty: document.querySelector("#review-empty"),
  reviewCard: document.querySelector("#review-card"),
  reviewKind: document.querySelector("#review-kind"),
  reviewStage: document.querySelector("#review-stage"),
  reviewDue: document.querySelector("#review-due"),
  reviewPrompt: document.querySelector("#review-prompt"),
  reviewFront: document.querySelector("#review-front"),
  reviewContext: document.querySelector("#review-context"),
  answerForm: document.querySelector("#answer-form"),
  answerInput: document.querySelector("#answer-input"),
  answerSubmitButton: document.querySelector("#answer-submit"),
  reviewFeedback: document.querySelector("#review-feedback"),
  reviewTypoWrap: document.querySelector("#review-typo-wrap"),
  reviewAdoptWrap: document.querySelector("#review-adopt-wrap"),
  reviewEasyWrap: document.querySelector("#review-easy-wrap"),
  reviewAdoptButton: document.querySelector("#review-adopt-answer"),
  reviewEasyButton: document.querySelector("#review-easy"),
  reviewTypoButton: document.querySelector("#review-typo"),
  reviewRemainingCount: document.querySelector("#review-remaining-count"),
  reviewQueueToggle: document.querySelector("#review-queue-toggle"),
  shuffleDueButton: document.querySelector("#shuffle-due"),
  refreshTrainingButton: document.querySelector("#refresh-training"),
  reviewQueuePanel: document.querySelector("#review-queue-panel"),
  reviewWrongTally: document.querySelector("#review-wrong-tally"),
  cramDragonModeButton: document.querySelector("#cram-dragon-mode"),
  startCramButton: document.querySelector("#start-cram"),
  clearCramButton: document.querySelector("#clear-cram"),
  cramRandomTenButton: document.querySelector("#cram-random-ten"),
  cramSelection: document.querySelector("#cram-selection"),
  cramEmpty: document.querySelector("#cram-empty"),
  cramCard: document.querySelector("#cram-card"),
  cramKind: document.querySelector("#cram-kind"),
  cramStage: document.querySelector("#cram-stage"),
  cramPrompt: document.querySelector("#cram-prompt"),
  cramFront: document.querySelector("#cram-front"),
  cramContext: document.querySelector("#cram-context"),
  cramForm: document.querySelector("#cram-form"),
  cramInput: document.querySelector("#cram-input"),
  cramSubmitButton: document.querySelector("#cram-submit"),
  cramFeedback: document.querySelector("#cram-feedback"),
  cramTypoWrap: document.querySelector("#cram-typo-wrap"),
  cramTypoButton: document.querySelector("#cram-typo"),
  cramAdoptWrap: document.querySelector("#cram-adopt-wrap"),
  cramAdoptButton: document.querySelector("#cram-adopt-answer"),
  cramWrongTally: document.querySelector("#cram-wrong-tally"),
  kanjiForm: document.querySelector("#kanji-form"),
  resetFormButton: document.querySelector("#reset-form"),
  kanjiId: document.querySelector("#kanji-id"),
  kanjiLookupButton: document.querySelector("#kanji-lookup"),
  kanjiLookupStatus: document.querySelector("#kanji-lookup-status"),
  kanjiChar: document.querySelector("#kanji-char"),
  kanjiMeanings: document.querySelector("#kanji-meanings"),
  kanjiOnyomi: document.querySelector("#kanji-onyomi"),
  kanjiKunyomi: document.querySelector("#kanji-kunyomi"),
  kanjiJlpt: document.querySelector("#kanji-jlpt"),
  kanjiGrade: document.querySelector("#kanji-grade"),
  kanjiMnemonic: document.querySelector("#kanji-mnemonic"),
  wordLookupSearch: document.querySelector("#word-lookup-search"),
  wordLookupButton: document.querySelector("#word-lookup-run"),
  wordLookupStatus: document.querySelector("#word-lookup-status"),
  wordLookupResults: document.querySelector("#word-lookup-results"),
  wordDraftTerm: document.querySelector("#word-draft-term"),
  wordDraftReading: document.querySelector("#word-draft-reading"),
  wordDraftDefinition: document.querySelector("#word-draft-definition"),
  wordDraftCategory: document.querySelector("#word-draft-category"),
  wordDraftAddButton: document.querySelector("#word-draft-add"),
  libraryControls: document.querySelector("#library-controls"),
  libraryJlptFilter: document.querySelector("#library-jlpt-filter"),
  libraryGradeFilter: document.querySelector("#library-grade-filter"),
  wordLists: {
    on: document.querySelector("#word-list-on"),
    kun: document.querySelector("#word-list-kun"),
    irregular: document.querySelector("#word-list-irregular")
  },
  wordSummaryLists: {
    on: document.querySelector("#word-summary-on"),
    kun: document.querySelector("#word-summary-kun"),
    irregular: document.querySelector("#word-summary-irregular")
  },
  searchInput: document.querySelector("#library-search"),
  kanjiList: document.querySelector("#kanji-list")
};

const state = {
  deck: createEmptyDeck(),
  descriptors: [],
  descriptorMap: new Map(),
  sourceLabel: "New deck",
  unsavedChanges: false,
  currentCardId: null,
  currentChallenge: null,
  lastCheck: null,
  lastWordByDescriptor: {},
  deckFileHandle: null,
  bulkImportDraft: null,
  theme: loadStoredTheme(),
  jpFont: loadStoredJapaneseFont(),
  libraryView: "grid",
  libraryGridExpanded: false,
  masteryFilter: "all",
  jlptFilter: "all",
  gradeFilter: "all",
  selectedKanjiId: null,
  libraryMultiSelect: false,
  multiSelectKanjiIds: [],
  wordLookupResults: [],
  wordLookupTotalForKanji: 0,
  lastAuthoringLookupChar: "",
  reviewQueueOpen: false,
  reviewOrderIds: [],
  reviewWrongPrompts: [],
  cramKanjiIds: [],
  cramDragonMode: false,
  lastRandomCramKanjiIds: [],
  cramQueue: [],
  cramCurrentCardId: null,
  cramCurrentChallenge: null,
  cramLastCheck: null,
  cramWrongPrompts: []
};

const lookupRuntime = {
  scriptPromises: new Map(),
  kanjiLoaded: false,
  kanjiInputTimer: null,
  wordLookupComposing: false,
  wordLookupRequestId: 0
};

let browserBackupDbPromise = null;
let deckLoadToken = 0;

function focusInputIfSectionVisible(input) {
  if (!input) {
    return;
  }

  const section = input.closest("section");
  const activeElement = document.activeElement;
  const shouldFocus = section && section.contains(activeElement);
  if (!shouldFocus) {
    return;
  }

  try {
    input.focus({ preventScroll: true });
  } catch (error) {
    input.focus();
  }
}

init();

async function init() {
  applyTheme(state.theme);
  applyJapaneseFont(state.jpFont);
  attachEvents();
  const initLoadToken = deckLoadToken;
  const restoreState = await hydrateDeckFromBrowserBackup();
  if (deckLoadToken !== initLoadToken) {
    return;
  }
  if (restoreState.restored && restoreState.backup) {
    applyBrowserBackup(restoreState.backup);
  }
  if (!restoreState.restored) {
    resetEditor();
    refreshDerivedState();
  }
  render();

  if (restoreState.restored) {
    setStatus("Restored your browser backup automatically.", "success");
  } else if (restoreState.error) {
    setStatus("The browser backup exists but could not be restored.", "error");
  }
}

function attachEvents() {
  dom.importInput.addEventListener("change", handleImportFile);
  dom.txtImportInput.addEventListener("change", handleTxtImportFile);
  dom.exportButton.addEventListener("click", exportDeck);
  dom.openFileDeckButton.addEventListener("click", openDeckFile);
  dom.saveFileDeckButton.addEventListener("click", saveDeckFile);
  dom.saveFileAsDeckButton.addEventListener("click", saveDeckFileAs);
  dom.bulkImportTemplateButton.addEventListener("click", downloadBulkImportTemplate);
  dom.exportAllKanjiDataButton.addEventListener("click", exportAllKanjiData);
  dom.exportSelectedKanjiDataButton.addEventListener("click", exportSelectedKanjiData);
  dom.exportFilteredKanjiDataButton.addEventListener("click", exportFilteredKanjiData);
  dom.themeSelect.addEventListener("change", () => {
    applyTheme(dom.themeSelect.value);
  });
  if (dom.jpFontSelect) {
    dom.jpFontSelect.addEventListener("change", () => {
      applyJapaneseFont(dom.jpFontSelect.value);
    });
  }
  dom.newDeckButton.addEventListener("click", createBlankDeck);
  dom.demoDeckButton.addEventListener("click", loadDemoDeck);
  dom.restoreBackupButton.addEventListener("click", restoreBackupDeck);
  dom.startReviewButton.addEventListener("click", startReviewSession);
  dom.reviewQueueToggle.addEventListener("click", toggleReviewQueue);
  dom.shuffleDueButton.addEventListener("click", shuffleReviewQueue);
  dom.refreshTrainingButton?.addEventListener("click", handleManualTrainingRefresh);
  dom.answerForm.addEventListener("submit", handleAnswerSubmit);
  dom.reviewEasyButton.addEventListener("click", markReviewEasy);
  dom.reviewTypoButton.addEventListener("click", markReviewTypo);
  dom.reviewAdoptButton?.addEventListener("click", adoptReviewAnswer);
  dom.startCramButton.addEventListener("click", startCramSession);
  dom.cramDragonModeButton.addEventListener("click", toggleCramDragonMode);
  dom.clearCramButton.addEventListener("click", clearCramList);
  dom.cramRandomTenButton?.addEventListener("click", addRandomKanjiToCram);
  dom.cramForm.addEventListener("submit", handleCramSubmit);
  dom.cramTypoButton.addEventListener("click", markCramTypo);
  dom.cramAdoptButton?.addEventListener("click", adoptCramAnswer);
  dom.kanjiForm.addEventListener("submit", handleKanjiSave);
  dom.resetFormButton.addEventListener("click", resetEditor);
  dom.kanjiLookupButton.addEventListener("click", () => {
    lookupKanjiForEditor({ force: true });
  });
  dom.kanjiChar.addEventListener("input", scheduleKanjiLookup);
  dom.kanjiChar.addEventListener("keydown", (event) => {
    if (event.key !== "Enter") {
      return;
    }

    event.preventDefault();
    lookupKanjiForEditor({ force: true });
  });
  dom.wordLookupButton.addEventListener("click", () => {
    refreshWordLookupForCurrentKanji();
  });
  dom.wordLookupSearch.addEventListener("input", () => {
    if (!safeString(dom.wordLookupSearch.value)) {
      clearWordLookupResults("", "hidden");
    } else {
      setLookupStatus(dom.wordLookupStatus, "", "hidden");
    }
  });
  dom.wordLookupSearch.addEventListener("compositionstart", () => {
    lookupRuntime.wordLookupComposing = true;
  });
  dom.wordLookupSearch.addEventListener("compositionend", () => {
    lookupRuntime.wordLookupComposing = false;
  });
  dom.wordLookupSearch.addEventListener("keydown", (event) => {
    if (event.isComposing || lookupRuntime.wordLookupComposing) {
      return;
    }

    if (event.key !== "Enter") {
      return;
    }

    event.preventDefault();
    refreshWordLookupForCurrentKanji();
  });
  [dom.wordDraftTerm, dom.wordDraftReading, dom.wordDraftDefinition].forEach((input) => {
    input.addEventListener("keydown", (event) => {
      if (event.key !== "Enter") {
        return;
      }

      event.preventDefault();
      handleWordDraftAdd();
    });
  });
  dom.wordDraftAddButton.addEventListener("click", handleWordDraftAdd);
  dom.wordLookupResults.addEventListener("click", handleWordLookupResultClick);
  dom.searchInput.addEventListener("input", renderLibrary);
  dom.libraryControls.addEventListener("click", handleLibraryControlsClick);
  dom.libraryJlptFilter.addEventListener("change", handleLibraryMetaFilterChange);
  dom.libraryGradeFilter.addEventListener("change", handleLibraryMetaFilterChange);
  window.addEventListener("resize", handleLibraryGridResize);
  dom.cramSelection.addEventListener("click", handleCramSelectionClick);
  dom.reviewWrongTally?.addEventListener("click", handleReviewWrongTallyClick);
  dom.cramWrongTally?.addEventListener("click", handleCramWrongTallyClick);
  dom.bulkImportPreview.addEventListener("click", handleBulkImportPreviewClick);

  document.querySelectorAll("[data-add-word]").forEach((button) => {
    button.addEventListener("click", () => {
      const category = button.dataset.addWord;
      appendWordRow(category, blankWord(), { focus: true });
    });
  });

  WORD_CATEGORIES.forEach((category) => {
    dom.wordLists[category].addEventListener("click", (event) => {
      const removeButton = event.target.closest("[data-remove-word]");
      if (!removeButton) {
        return;
      }

      const row = removeButton.closest(".word-row");
      if (row) {
        row.remove();
      }

      if (!dom.wordLists[category].children.length) {
        appendWordRow(category, blankWord());
      }

      renderAuthoringWordSummary();
    });

    dom.wordLists[category].addEventListener("input", renderAuthoringWordSummary);
    dom.wordSummaryLists[category].addEventListener("click", (event) => {
      const removeButton = event.target.closest("[data-remove-summary-word]");
      if (!removeButton) {
        return;
      }

      removeWordById(category, removeButton.dataset.removeSummaryWord || "");
    });
  });

  dom.kanjiList.addEventListener("click", (event) => {
    const toggleGridButton = event.target.closest("[data-toggle-library-grid]");
    if (toggleGridButton) {
      state.libraryGridExpanded = !state.libraryGridExpanded;
      renderLibrary();
      return;
    }

    const closeButton = event.target.closest("[data-close-details]");
    if (closeButton) {
      state.selectedKanjiId = null;
      renderLibrary();
      return;
    }

    const cramButton = event.target.closest("[data-toggle-cram]");
    if (cramButton) {
      toggleCramKanji(cramButton.dataset.toggleCram);
      return;
    }

    const multiSelectButton = event.target.closest("[data-toggle-multi-select]");
    if (multiSelectButton) {
      toggleMultiSelectKanji(multiSelectButton.dataset.toggleMultiSelect);
      return;
    }

    const actionButton = event.target.closest("[data-action]");
    if (!actionButton) {
      const row = event.target.closest("[data-select-kanji]");
      if (!row) {
        return;
      }

      const nextId = row.dataset.selectKanji || null;
      if (state.libraryMultiSelect) {
        toggleMultiSelectKanji(nextId);
        return;
      }

      state.selectedKanjiId = state.selectedKanjiId === nextId ? null : nextId;
      renderLibrary();
      return;
    }

    const kanjiId = actionButton.dataset.kanjiId;
    if (!kanjiId) {
      return;
    }

    if (actionButton.dataset.action === "edit") {
      loadKanjiIntoEditor(kanjiId);
      return;
    }

    if (actionButton.dataset.action === "delete") {
      deleteKanji(kanjiId);
    }
  });

  window.addEventListener("beforeunload", (event) => {
    if (!state.unsavedChanges) {
      return;
    }

    event.preventDefault();
    event.returnValue = "";
  });
}

function scheduleKanjiLookup() {
  window.clearTimeout(lookupRuntime.kanjiInputTimer);
  lookupRuntime.kanjiInputTimer = window.setTimeout(() => {
    lookupKanjiForEditor();
  }, 180);
}

async function lookupKanjiForEditor(options = {}) {
  const kanjiChar = extractEditorKanjiChar();
  if (!kanjiChar) {
    state.lastAuthoringLookupChar = "";
    setLookupStatus(dom.kanjiLookupStatus, "", "hidden");
    clearWordLookupResults("", "hidden");
    return;
  }

  if (!options.force && state.lastAuthoringLookupChar === kanjiChar) {
    return;
  }

  try {
    const lookup = await ensureKanjiLookupLoaded();
    const entry = lookup[kanjiChar];
    state.lastAuthoringLookupChar = kanjiChar;

    if (!entry) {
      setLookupStatus(dom.kanjiLookupStatus, `No KANJIDIC2 match was found for ${kanjiChar}. You can still fill the fields manually.`, "error");
      return;
    }

    dom.kanjiChar.value = kanjiChar;
    dom.kanjiMeanings.value = (entry.meanings || []).join("\n");
    dom.kanjiOnyomi.value = (entry.onyomi || []).join("\n");
    dom.kanjiKunyomi.value = (entry.kunyomi || []).join("\n");
    dom.kanjiJlpt.value = entry.jlpt || "";
    dom.kanjiGrade.value = entry.grade || "";
    setLookupStatus(dom.kanjiLookupStatus, "", "hidden");
  } catch (error) {
    setLookupStatus(dom.kanjiLookupStatus, "The kanji lookup data could not be loaded right now.", "error");
  }
}

async function refreshWordLookupForCurrentKanji() {
  const rawQuery = safeString(dom.wordLookupSearch.value);
  const requestId = ++lookupRuntime.wordLookupRequestId;

  if (!rawQuery) {
    clearWordLookupResults("", "hidden");
    return;
  }

  try {
    const candidates = await collectWordLookupCandidates(rawQuery);
    const queryTerms = parseLookupQuery(rawQuery);
    const filtered = candidates.filter((candidate) => matchesWordLookupQuery(candidate, queryTerms)).slice(0, LOOKUP_RESULT_LIMIT);

    if (requestId !== lookupRuntime.wordLookupRequestId) {
      return;
    }

    state.wordLookupResults = filtered;
    state.wordLookupTotalForKanji = candidates.length;
    renderWordLookupResults(true);

    if (!filtered.length) {
      setLookupStatus(dom.wordLookupStatus, `No words matched "${rawQuery}".`, "error");
      return;
    }

    setLookupStatus(dom.wordLookupStatus, "", "hidden");
  } catch (error) {
    clearWordLookupResults("The word lookup data could not be loaded right now.", "error");
  }
}

function renderWordLookupResults(resetScroll = false) {
  const previousScrollTop = dom.wordLookupResults.scrollTop;
  if (!state.wordLookupResults.length) {
    dom.wordLookupResults.innerHTML = "";
    if (resetScroll) {
      dom.wordLookupResults.scrollTop = 0;
    }
    return;
  }

  dom.wordLookupResults.innerHTML = state.wordLookupResults.map((candidate, index) => `
    <button class="word-lookup-result${isWordCandidateAdded(candidate) ? " is-added" : ""}" type="button" data-word-lookup-index="${index}">
      <span class="word-lookup-result-head">
        <span class="word-lookup-term">${escapeHtml(candidate.term)}</span>
        <span class="word-lookup-rank">${escapeHtml(formatWordLookupRanks(candidate))}</span>
      </span>
      <span class="word-lookup-reading">${escapeHtml(candidate.readings.join(", "))}</span>
      <span class="word-lookup-definition">${escapeHtml(candidate.definition || "No short gloss available")}</span>
      ${isWordCandidateAdded(candidate) ? '<span class="word-lookup-added">Added</span>' : ""}
    </button>
  `).join("");
  dom.wordLookupResults.scrollTop = resetScroll ? 0 : previousScrollTop;
}

function isWordCandidateAdded(candidate) {
  const candidateTerm = safeString(candidate && candidate.term);
  if (!candidateTerm) {
    return false;
  }

  const candidateReadings = new Set(
    (Array.isArray(candidate && candidate.readings) ? candidate.readings : [])
      .map((reading) => normalizeReading(reading))
      .filter(Boolean)
  );

  return WORD_CATEGORIES.some((category) => collectWordsFromList(dom.wordLists[category]).some((word) => {
    if (safeString(word.word) !== candidateTerm) {
      return false;
    }

    if (!candidateReadings.size) {
      return true;
    }

    return parseWordReadings(word.reading)
      .map((reading) => normalizeReading(reading))
      .some((reading) => candidateReadings.has(reading));
  }));
}

function handleWordLookupResultClick(event) {
  const result = event.target.closest("[data-word-lookup-index]");
  if (!result) {
    return;
  }

  const candidate = state.wordLookupResults[Number(result.dataset.wordLookupIndex)];
  if (!candidate) {
    return;
  }

  dom.wordDraftTerm.value = candidate.term;
  dom.wordDraftReading.value = candidate.readings.join(", ");
  dom.wordDraftDefinition.value = candidate.definition || "";
  dom.wordDraftCategory.value = suggestWordCategory(candidate.readings);
  dom.wordDraftTerm.focus();
}

function handleWordDraftAdd() {
  const category = dom.wordDraftCategory.value;
  const word = {
    id: uid("word"),
    word: safeString(dom.wordDraftTerm.value),
    reading: safeString(dom.wordDraftReading.value),
    definition: safeString(dom.wordDraftDefinition.value)
  };

  if (!word.word || !word.reading) {
    setStatus("Choose a lookup result or type at least a word and reading before adding it.", "error");
    return;
  }

  appendWordRow(category, word);
  renderAuthoringWordSummary();
  renderWordLookupResults(false);
  setStatus(`Added ${word.word} to the ${WORD_CATEGORY_LABELS[category]} word list.`, "success");
  resetWordDraft(category);
  dom.wordLookupSearch.focus();
}

function renderAuthoringWordSummary() {
  WORD_CATEGORIES.forEach((category) => {
    const container = dom.wordSummaryLists[category];
    const words = collectWordsFromList(dom.wordLists[category]);

    if (!words.length) {
      container.innerHTML = `<p class="word-summary-empty">No words yet.</p>`;
      return;
    }

    container.innerHTML = words.map((word) => `
      <div class="word-summary-item">
        <div class="word-summary-copy">
          <strong>${escapeHtml(word.word || "Untitled")}</strong>
          <span>${escapeHtml(formatWordSummaryLine(word))}</span>
        </div>
        <button class="button button-small word-summary-remove" type="button" data-remove-summary-word="${escapeHtml(word.id)}">Remove</button>
      </div>
    `).join("");
  });
}

function removeWordById(category, wordId) {
  const rows = Array.from(dom.wordLists[category].querySelectorAll(".word-row"));
  const row = rows.find((item) => item.dataset.wordId === wordId);
  if (row) {
    row.remove();
  }

  if (!dom.wordLists[category].children.length) {
    appendWordRow(category, blankWord());
  }

  renderAuthoringWordSummary();
  renderWordLookupResults(false);
}

function resetWordDraft(category = "on") {
  dom.wordDraftTerm.value = "";
  dom.wordDraftReading.value = "";
  dom.wordDraftDefinition.value = "";
  dom.wordDraftCategory.value = WORD_CATEGORIES.includes(category) ? category : "on";
}

function clearWordLookupResults(message, tone) {
  state.wordLookupResults = [];
  state.wordLookupTotalForKanji = 0;
  dom.wordLookupResults.innerHTML = "";
  setLookupStatus(dom.wordLookupStatus, message, tone);
}

function setLookupStatus(element, message, tone = "info") {
  if (!element) {
    return;
  }

  element.textContent = message;
  element.dataset.tone = safeString(message) ? tone : "hidden";
}

async function ensureKanjiLookupLoaded() {
  if (lookupRuntime.kanjiLoaded && window.KANJI_DEN_KANJI_LOOKUP) {
    return window.KANJI_DEN_KANJI_LOOKUP;
  }

  await loadLookupScript(`${LOOKUP_SCRIPT_ROOT}/kanji-lookup-data.js`, () => Boolean(window.KANJI_DEN_KANJI_LOOKUP));
  lookupRuntime.kanjiLoaded = Boolean(window.KANJI_DEN_KANJI_LOOKUP);
  return window.KANJI_DEN_KANJI_LOOKUP || {};
}

async function ensureWordLookupBucket(kanjiChar) {
  const bucketId = getWordLookupBucketId(kanjiChar);
  return ensureWordLookupBucketById(bucketId);
}

async function ensureWordLookupBucketById(bucketId) {
  await loadLookupScript(
    `${LOOKUP_BUCKET_ROOT}/${bucketId}.js`,
    () => Boolean(window.KANJI_DEN_WORD_LOOKUP_BUCKETS && window.KANJI_DEN_WORD_LOOKUP_BUCKETS[bucketId])
  );

  return window.KANJI_DEN_WORD_LOOKUP_BUCKETS && window.KANJI_DEN_WORD_LOOKUP_BUCKETS[bucketId]
    ? window.KANJI_DEN_WORD_LOOKUP_BUCKETS[bucketId]
    : {};
}

async function ensureWordLookupManifestLoaded() {
  if (Array.isArray(window.KANJI_DEN_WORD_LOOKUP_MANIFEST)) {
    return window.KANJI_DEN_WORD_LOOKUP_MANIFEST;
  }

  await loadLookupScript(
    `${LOOKUP_SCRIPT_ROOT}/word-bucket-manifest.js`,
    () => Array.isArray(window.KANJI_DEN_WORD_LOOKUP_MANIFEST)
  );

  return Array.isArray(window.KANJI_DEN_WORD_LOOKUP_MANIFEST) ? window.KANJI_DEN_WORD_LOOKUP_MANIFEST : [];
}

async function collectWordLookupCandidates(rawQuery) {
  const hanChars = extractHanCharsFromText(rawQuery);
  const tuples = [];

  if (hanChars.length) {
    const bucketMaps = await Promise.all(hanChars.map((kanjiChar) => ensureWordLookupBucket(kanjiChar)));
    bucketMaps.forEach((bucketMap, index) => {
      const hanChar = hanChars[index];
      const list = Array.isArray(bucketMap[hanChar]) ? bucketMap[hanChar] : [];
      tuples.push(...list);
    });
  } else {
    const bucketIds = await ensureWordLookupManifestLoaded();
    const bucketMaps = await Promise.all(bucketIds.map((bucketId) => ensureWordLookupBucketById(bucketId)));
    bucketMaps.forEach((bucketMap) => {
      Object.values(bucketMap).forEach((list) => {
        if (Array.isArray(list)) {
          tuples.push(...list);
        }
      });
    });
  }

    return dedupeWordLookupCandidates(tuples)
      .map(normalizeWordLookupCandidate)
      .sort(compareWordLookupCandidates);
}

function loadLookupScript(src, isReady) {
  if (isReady()) {
    return Promise.resolve();
  }

  if (lookupRuntime.scriptPromises.has(src)) {
    return lookupRuntime.scriptPromises.get(src);
  }

  const promise = new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = src;
    script.async = true;
    script.onload = () => resolve();
    script.onerror = () => reject(new Error(`Could not load ${src}`));
    document.head.appendChild(script);
  }).catch((error) => {
    lookupRuntime.scriptPromises.delete(src);
    throw error;
  });

  lookupRuntime.scriptPromises.set(src, promise);
  return promise;
}

function extractEditorKanjiChar() {
  const value = safeString(dom.kanjiChar.value);
  const match = value.match(KANJI_SEARCH_CHAR_PATTERN);
  return match ? match[0] : "";
}

function extractHanCharsFromText(value) {
  return uniqueList(Array.from(safeString(value).matchAll(HAN_SEARCH_GLOBAL_PATTERN), (match) => match[0]));
}

function getWordLookupBucketId(kanjiChar) {
  const codePoint = kanjiChar.codePointAt(0) || 0;
  return Math.floor(codePoint / 256).toString(16).padStart(2, "0");
}

function normalizeWordLookupCandidate(tuple) {
  return {
    term: safeString(tuple && tuple[0]),
    readings: Array.isArray(tuple && tuple[1]) ? tuple[1].map((reading) => safeString(reading)).filter(Boolean) : [],
    definition: safeString(tuple && tuple[2]),
    writtenRank: Number.isFinite(tuple && tuple[3]) ? tuple[3] : null,
    kanaRank: Number.isFinite(tuple && tuple[4]) ? tuple[4] : null
  };
}

function dedupeWordLookupCandidates(tuples) {
  const seen = new Set();
  const deduped = [];

  tuples.forEach((tuple) => {
    const key = JSON.stringify(tuple);
    if (seen.has(key)) {
      return;
    }

    seen.add(key);
    deduped.push(tuple);
  });

  return deduped;
}

function compareWordLookupCandidates(left, right) {
  const leftHasWritten = Number.isFinite(left.writtenRank);
  const rightHasWritten = Number.isFinite(right.writtenRank);

  if (leftHasWritten !== rightHasWritten) {
    return leftHasWritten ? -1 : 1;
  }

  if (leftHasWritten && left.writtenRank !== right.writtenRank) {
    return left.writtenRank - right.writtenRank;
  }

  const leftHasKana = Number.isFinite(left.kanaRank);
  const rightHasKana = Number.isFinite(right.kanaRank);

  if (leftHasKana !== rightHasKana) {
    return leftHasKana ? -1 : 1;
  }

  if (leftHasKana && left.kanaRank !== right.kanaRank) {
    return left.kanaRank - right.kanaRank;
  }

  return left.term.localeCompare(right.term, "ja");
}

function parseLookupQuery(value) {
  return safeString(value)
    .split(SEARCH_SPLIT_PATTERN)
    .map(buildWordLookupQueryToken)
    .filter(Boolean);
}

function matchesWordLookupQuery(candidate, queryTerms) {
  if (!queryTerms.length) {
    return true;
  }

  const japaneseHaystack = buildSearchText([
    candidate.term,
    candidate.readings.join(" ")
  ].join(" "));
  const definitionTokens = buildMeaningWordSet(candidate.definition);

  return queryTerms.every((queryTerm) => {
    if (queryTerm.type === "latin") {
      return queryTerm.words.every((word) => definitionTokens.has(word));
    }

    return queryTerm.forms.some((form) => japaneseHaystack.includes(form));
  });
}

function buildWordLookupQueryToken(value) {
  const text = safeString(value);
  if (!text) {
    return null;
  }

  if (containsJapaneseSearchChar(text)) {
    const forms = buildSearchForms(text);
    return forms.length ? { type: "japanese", forms } : null;
  }

  const words = Array.from(buildMeaningWordSet(text));
  return words.length ? { type: "latin", words } : null;
}

function containsJapaneseSearchChar(value) {
  return JAPANESE_SEARCH_CHAR_PATTERN.test(safeString(value));
}

function buildMeaningWordSet(value) {
  return new Set(
    normalizeMeaning(value)
      .split(/[\s-]+/)
      .map((part) => part.trim())
      .filter(Boolean)
  );
}

function formatWordLookupRanks(candidate) {
  if (Number.isFinite(candidate.writtenRank)) {
    return `#${candidate.writtenRank}${Number.isFinite(candidate.kanaRank) ? ` | #${candidate.kanaRank} ㋕` : ""}`;
  }

  if (Number.isFinite(candidate.kanaRank)) {
    return `#${candidate.kanaRank} ㋕`;
  }

  return "Unranked";
}

function suggestWordCategory(readings) {
  const normalizedReadings = uniqueList((Array.isArray(readings) ? readings : []).map((reading) => normalizeReading(reading)).filter(Boolean));
  const onSet = new Set(parseListFromField(dom.kanjiOnyomi.value).map((reading) => normalizeReading(reading)));
  const kunSet = new Set(parseListFromField(dom.kanjiKunyomi.value).map((reading) => normalizeReading(reading)));
  const matchesOn = normalizedReadings.some((reading) => onSet.has(reading));
  const matchesKun = normalizedReadings.some((reading) => kunSet.has(reading));

  if (matchesOn && !matchesKun) {
    return "on";
  }

  if (matchesKun && !matchesOn) {
    return "kun";
  }

  return "irregular";
}

const WORD_CATEGORY_LABELS = {
  on: "onyomi",
  kun: "kunyomi",
  irregular: "irregular"
};

function loadStoredTheme() {
  const storedTheme = localStorage.getItem(THEME_STORAGE_KEY);
  return normalizeThemeChoice(storedTheme);
}

function applyTheme(theme) {
  const nextTheme = normalizeThemeChoice(theme);
  state.theme = nextTheme;
  document.documentElement.dataset.theme = nextTheme;
  localStorage.setItem(THEME_STORAGE_KEY, nextTheme);

  if (dom.themeSelect) {
    dom.themeSelect.value = nextTheme;
  }
}

function normalizeThemeChoice(theme) {
  const normalizedTheme = LEGACY_THEME_ALIASES[theme] || theme;
  return THEME_OPTIONS.includes(normalizedTheme) ? normalizedTheme : "forest";
}

function loadStoredJapaneseFont() {
  const storedFont = localStorage.getItem(JP_FONT_STORAGE_KEY);
  return Object.prototype.hasOwnProperty.call(JAPANESE_FONT_OPTIONS, storedFont)
    ? storedFont
    : DEFAULT_JP_FONT;
}

function applyJapaneseFont(fontKey) {
  const nextFont = Object.prototype.hasOwnProperty.call(JAPANESE_FONT_OPTIONS, fontKey)
    ? fontKey
    : DEFAULT_JP_FONT;

  state.jpFont = nextFont;
  document.documentElement.style.setProperty("--jp-font-family", JAPANESE_FONT_OPTIONS[nextFont]);
  localStorage.setItem(JP_FONT_STORAGE_KEY, nextFont);

  if (dom.jpFontSelect) {
    dom.jpFontSelect.value = nextFont;
  }
}

function createEmptyDeck() {
  const now = new Date().toISOString();
  return {
    version: APP_VERSION,
    meta: {
      appName: "The Kanji Den",
      createdAt: now,
      updatedAt: now,
      desiredRetention: REVIEW_RETENTION
    },
    kanji: [],
    progress: {
      cards: {}
    }
  };
}

function createDemoDeck() {
  return normalizeDeck({
    version: APP_VERSION,
    kanji: [
      {
        id: uid("kanji"),
        char: "日",
        meanings: ["sun", "day"],
        onyomi: ["ニチ", "ジツ"],
        kunyomi: ["ひ", "-び", "-か"],
        words: {
          on: [
            { id: uid("word"), word: "日本", reading: "にほん", definition: "Japan" },
            { id: uid("word"), word: "休日", reading: "きゅうじつ", definition: "day off" }
          ],
          kun: [
            { id: uid("word"), word: "日向", reading: "ひなた", definition: "sunny place" }
          ],
          irregular: [
            { id: uid("word"), word: "今日", reading: "きょう", definition: "today" }
          ]
        }
      },
      {
        id: uid("kanji"),
        char: "生",
        meanings: ["life", "birth", "raw"],
        onyomi: ["セイ", "ショウ"],
        kunyomi: ["いきる", "うまれる", "なま"],
        words: {
          on: [
            { id: uid("word"), word: "学生", reading: "がくせい", definition: "student" },
            { id: uid("word"), word: "一生", reading: "いっしょう", definition: "a lifetime" }
          ],
          kun: [
            { id: uid("word"), word: "生まれる", reading: "うまれる", definition: "to be born" },
            { id: uid("word"), word: "生きる", reading: "いきる", definition: "to live" }
          ],
          irregular: [
            { id: uid("word"), word: "芝生", reading: "しばふ", definition: "lawn" }
          ]
        }
      }
    ]
  });
}

function normalizeDeck(rawDeck) {
  const emptyDeck = createEmptyDeck();
  const kanji = Array.isArray(rawDeck && rawDeck.kanji)
    ? rawDeck.kanji.map(normalizeKanji).filter((entry) => entry.char)
    : [];

  return {
    version: APP_VERSION,
    meta: {
      ...emptyDeck.meta,
      ...(rawDeck && typeof rawDeck.meta === "object" ? rawDeck.meta : {}),
      updatedAt: new Date().toISOString()
    },
    kanji,
    progress: {
      cards: normalizeCards(rawDeck && rawDeck.progress && rawDeck.progress.cards)
    }
  };
}

function normalizeKanji(rawKanji) {
  const words = rawKanji && typeof rawKanji.words === "object" ? rawKanji.words : {};
  return {
    id: safeString(rawKanji && rawKanji.id) || uid("kanji"),
    char: safeString(rawKanji && rawKanji.char),
    meanings: uniqueList(normalizeList(rawKanji && rawKanji.meanings)),
    onyomi: uniqueList(normalizeList(rawKanji && rawKanji.onyomi)),
    kunyomi: uniqueList(normalizeList(rawKanji && rawKanji.kunyomi)),
    jlpt: safeString(rawKanji && rawKanji.jlpt),
    grade: safeString(rawKanji && rawKanji.grade),
    mnemonic: safeString(rawKanji && rawKanji.mnemonic),
    words: {
      on: normalizeWordList(words.on),
      kun: normalizeWordList(words.kun),
      irregular: normalizeWordList(words.irregular)
    }
  };
}

function normalizeWordList(list) {
  if (!Array.isArray(list)) {
    return [];
  }

  return list
    .map((word) => ({
      id: safeString(word && word.id) || uid("word"),
      word: safeString(word && word.word),
      reading: safeString(word && word.reading),
      definition: safeString(word && word.definition)
    }))
    .filter((word) => word.word || word.reading || word.definition);
}

function normalizeCards(cards) {
  if (!cards || typeof cards !== "object") {
    return {};
  }

  const normalized = {};
  Object.entries(cards).forEach(([cardId, rawCard]) => {
    normalized[cardId] = normalizeCardRecord(rawCard, cardId);
  });
  return normalized;
}

function refreshDerivedState() {
  const descriptors = buildDescriptors(state.deck);
  const descriptorMap = new Map();
  const syncedCards = {};

  descriptors.forEach((descriptor) => {
    descriptorMap.set(descriptor.id, descriptor);
    syncedCards[descriptor.id] = normalizeCardRecord(state.deck.progress.cards[descriptor.id], descriptor.id);
  });

  state.deck.progress.cards = syncedCards;
  state.deck.meta.updatedAt = new Date().toISOString();
  state.descriptors = descriptors;
  state.descriptorMap = descriptorMap;

  if (state.currentCardId && !state.descriptorMap.has(state.currentCardId)) {
    clearActiveReview();
  }

  if (state.selectedKanjiId && !state.deck.kanji.some((entry) => entry.id === state.selectedKanjiId)) {
    state.selectedKanjiId = null;
  }

  state.multiSelectKanjiIds = state.multiSelectKanjiIds.filter((kanjiId) => state.deck.kanji.some((entry) => entry.id === kanjiId));
  state.reviewOrderIds = state.reviewOrderIds.filter((cardId) => state.descriptorMap.has(cardId));
  state.reviewWrongPrompts = state.reviewWrongPrompts.filter((item) => state.deck.kanji.some((entry) => entry.id === item.kanjiId));
  state.cramKanjiIds = state.cramKanjiIds.filter((kanjiId) => state.deck.kanji.some((entry) => entry.id === kanjiId));
  state.lastRandomCramKanjiIds = state.lastRandomCramKanjiIds.filter((kanjiId) => state.deck.kanji.some((entry) => entry.id === kanjiId));
  state.cramQueue = state.cramQueue.filter((cardId) => state.descriptorMap.has(cardId));
  state.cramWrongPrompts = state.cramWrongPrompts.filter((item) => state.deck.kanji.some((entry) => entry.id === item.kanjiId));

  if (state.cramCurrentCardId && !state.descriptorMap.has(state.cramCurrentCardId)) {
    clearCramSession();
  }
}

function normalizeCardRecord(rawCard, id) {
  const fallback = {
    id,
    state: "new",
    due: new Date().toISOString(),
    lastReviewAt: null,
    stability: 0,
    difficulty: 5.5,
    scheduledDays: 0,
    reps: 0,
    lapses: 0,
    consecutiveFailures: 0,
    successfulReviews: 0,
    masteryLevel: 0,
    history: []
  };

  if (!rawCard || typeof rawCard !== "object") {
    return fallback;
  }

  return {
    id,
    state: typeof rawCard.state === "string" ? rawCard.state : fallback.state,
    due: safeDateString(rawCard.due) || fallback.due,
    lastReviewAt: safeDateString(rawCard.lastReviewAt),
    stability: clampNumber(rawCard.stability, 0, 3650, fallback.stability),
    difficulty: clampNumber(rawCard.difficulty, 1, 10, fallback.difficulty),
    scheduledDays: clampNumber(rawCard.scheduledDays, 0, 3650, fallback.scheduledDays),
    reps: clampNumber(rawCard.reps, 0, 100000, fallback.reps),
    lapses: clampNumber(rawCard.lapses, 0, 100000, fallback.lapses),
    consecutiveFailures: clampNumber(rawCard.consecutiveFailures, 0, 100000, fallback.consecutiveFailures),
    successfulReviews: clampNumber(rawCard.successfulReviews, 0, 100000, fallback.successfulReviews),
    masteryLevel: clampNumber(rawCard.masteryLevel, 0, MASTERY_LEVELS.length - 1, inferMasteryLevel(rawCard)),
    history: Array.isArray(rawCard.history) ? rawCard.history.slice(-60) : []
  };
}

function buildDescriptors(deck) {
  const descriptors = [];

  deck.kanji.forEach((entry) => {
    const onWords = getValidWordPrompts(entry.words.on);
    const kunWords = getValidWordPrompts(entry.words.kun);
    const irregularWords = getValidWordPrompts(entry.words.irregular);

    if (entry.meanings.length) {
      descriptors.push({
        id: `${entry.id}::meaning`,
        kanjiId: entry.id,
        kind: "meaning",
        label: "Meaning",
        char: entry.char,
        meanings: entry.meanings
      });
    }

    if (entry.onyomi.length || onWords.length) {
      descriptors.push({
        id: `${entry.id}::onyomi`,
        kanjiId: entry.id,
        kind: "onyomi",
        label: "Onyomi",
        char: entry.char,
        readings: entry.onyomi,
        words: onWords
      });
    }

    if (entry.kunyomi.length || kunWords.length) {
      descriptors.push({
        id: `${entry.id}::kunyomi`,
        kanjiId: entry.id,
        kind: "kunyomi",
        label: "Kunyomi",
        char: entry.char,
        readings: entry.kunyomi,
        words: kunWords
      });
    }

    if (irregularWords.length) {
      descriptors.push({
        id: `${entry.id}::irregular`,
        kanjiId: entry.id,
        kind: "irregular",
        label: "Irregular",
        char: entry.char,
        words: irregularWords
      });
    }
  });

  return descriptors;
}

function getValidWordPrompts(words) {
  return (Array.isArray(words) ? words : []).filter((word) => word.word && word.reading);
}

function createChallenge(descriptor, options = {}) {
  if (descriptor.kind === "meaning") {
    return {
      prompt: "Type one valid meaning for this kanji.",
      front: descriptor.char,
      context: "Give one English meaning.",
      answers: descriptor.meanings,
      answerMode: "meaning",
      reveal: `Accepted meanings: ${descriptor.meanings.join(" / ")}`,
      frontClass: ""
    };
  }

  const words = Array.isArray(descriptor.words) ? descriptor.words : [];
  if (words.length) {
    const chosenWord = options.specificWord || pickWordForChallenge(descriptor, words);
    return {
      prompt: "How is this word read?",
      front: chosenWord.word,
      context: buildWordPromptContext(descriptor, chosenWord),
      definition: chosenWord.definition,
      answers: parseWordReadings(chosenWord.reading),
      wordId: chosenWord.id || "",
      answerMode: "reading",
      reveal: buildWordReveal(chosenWord),
      frontClass: "is-word"
    };
  }

  const readingLabel = descriptor.kind === "onyomi" ? "onyomi" : "kunyomi";
  return {
    prompt: `Type one valid ${readingLabel} reading for this kanji.`,
    front: descriptor.char,
    context: `Enter one ${readingLabel} reading in kana or katakana.`,
    answers: descriptor.readings || [],
    answerMode: "reading",
    reveal: `Accepted ${readingLabel} readings: ${(descriptor.readings || []).join(" / ")}`,
    frontClass: ""
  };
}

function buildWordPromptContext(descriptor, word) {
  return `${descriptor.label} word | Parent kanji: ${descriptor.char}`;
}

function pickWordForChallenge(descriptor, words) {
  const lastKey = state.lastWordByDescriptor[descriptor.id] || "";
  const availableWords = words.length > 1
    ? words.filter((word) => getWordPromptKey(word) !== lastKey)
    : words;
  const chosenWord = pickRandom(availableWords.length ? availableWords : words);
  if (chosenWord) {
    state.lastWordByDescriptor[descriptor.id] = getWordPromptKey(chosenWord);
  }
  return chosenWord;
}

function getWordPromptKey(word) {
  return word.id || `${word.word}|${word.reading}|${word.definition}`;
}

function buildWordReveal(word) {
  const readings = parseWordReadings(word.reading);
  return `Reading: ${readings.length ? readings.join(" / ") : word.reading}`;
}

function parseWordReadings(value) {
  return uniqueList(safeString(value)
    .split(/[,;、；]+/)
    .map((entry) => safeString(entry))
    .filter(Boolean));
}

function renderChallengeContext(target, challenge, revealDefinition, suffix = "") {
  const parts = [];
  if (challenge.context) {
    parts.push(`<span>${escapeHtml(challenge.context)}</span>`);
  }

  if (challenge.definition) {
    parts.push(revealDefinition
      ? `<span class="definition-revealed">Definition: ${escapeHtml(challenge.definition)}</span>`
      : `<span class="definition-tooltip" tabindex="0">definition &#128214;<span class="definition-tooltip-text">${escapeHtml(challenge.definition)}</span></span>`);
  }

  if (suffix) {
    parts.push(`<span>${escapeHtml(suffix)}</span>`);
  }

  target.innerHTML = parts.join(' <span class="context-divider">|</span> ');
}

function handleLibraryControlsClick(event) {
  const viewButton = event.target.closest("[data-library-view]");
  if (viewButton) {
    state.libraryView = viewButton.dataset.libraryView || "grid";
    renderLibrary();
    return;
  }

  const filterButton = event.target.closest("[data-mastery-filter]");
  if (filterButton) {
    state.masteryFilter = filterButton.dataset.masteryFilter || "all";
    if (state.selectedKanjiId && !getFilteredKanji().some((entry) => entry.id === state.selectedKanjiId)) {
      state.selectedKanjiId = null;
    }
    renderLibrary();
    return;
  }

  const selectModeButton = event.target.closest("[data-library-select-mode]");
  if (selectModeButton) {
    state.libraryMultiSelect = !state.libraryMultiSelect;
    if (state.libraryMultiSelect) {
      state.selectedKanjiId = null;
    }
    renderLibrary();
    return;
  }

  const selectedActionButton = event.target.closest("[data-selected-cram]");
  if (selectedActionButton) {
    if (selectedActionButton.dataset.selectedCram === "add") {
      addSelectedKanjiToCram();
      return;
    }

    if (selectedActionButton.dataset.selectedCram === "clear") {
      clearMultiSelectKanji();
    }
    return;
  }

  const bulkButton = event.target.closest("[data-bulk-cram]");
  if (!bulkButton) {
    return;
  }

  if (bulkButton.dataset.bulkCram === "add") {
    addFilteredKanjiToCram();
    return;
  }

  if (bulkButton.dataset.bulkCram === "remove") {
    removeFilteredKanjiFromCram();
  }
}

function handleLibraryMetaFilterChange() {
  state.jlptFilter = safeString(dom.libraryJlptFilter.value) || "all";
  state.gradeFilter = safeString(dom.libraryGradeFilter.value) || "all";

  if (state.selectedKanjiId && !getFilteredKanji().some((entry) => entry.id === state.selectedKanjiId)) {
    state.selectedKanjiId = null;
  }

  renderLibrary();
}

function render() {
  renderStats();
  renderFileAccessControls();
  renderBulkImportPreview();
  renderReviewPanel();
  renderReviewDueQueue();
  renderLibrary();
  renderCramPanel();
  renderReviewWrongTally();
  renderCramWrongTally();
  renderBackupButton();
}

function renderStats() {
  const cards = Object.values(state.deck.progress.cards);
  const dueCount = getDueCardIds().length;
  const saveStatus = state.unsavedChanges ? "Unsaved changes" : "Clean";

  dom.deckSummaryText.textContent = "Appearance, Load/Export, Bulk Import/Export";
  dom.statSource.textContent = state.sourceLabel;
  dom.statKanji.textContent = String(state.deck.kanji.length);
  dom.statCards.textContent = String(cards.length);
  dom.statDue.textContent = String(dueCount);
  dom.statSave.textContent = saveStatus;
}

function renderFileAccessControls() {
  const supported = isFileSystemAccessSupported();
  dom.openFileDeckButton.disabled = !supported;
  dom.saveFileDeckButton.disabled = !supported;
  dom.saveFileAsDeckButton.disabled = !supported;
  dom.saveFileDeckButton.textContent = "Save to Current Deck";
  dom.fileAccessNote.textContent = "Your progress is saved to this browser and can be lost if you accidentally clear the site data, keep a back up by periodically exporting a json file of your progress.";
}

function renderReviewPanel() {
  const activeReview = resolveActiveReview();
  if (!activeReview) {
    dom.reviewEmpty.classList.remove("is-hidden");
    dom.reviewCard.classList.add("is-hidden");
    dom.answerSubmitButton.textContent = "Enter";
    dom.reviewAdoptWrap?.classList.add("is-hidden");
    dom.reviewEasyWrap.classList.add("is-hidden");
    dom.reviewTypoWrap.classList.add("is-hidden");
    return;
  }

  const { descriptor, challenge } = activeReview;
  const card = state.deck.progress.cards[descriptor.id];
  const check = getCurrentCheck(descriptor.id);

  dom.reviewEmpty.classList.add("is-hidden");
  dom.reviewCard.classList.remove("is-hidden");
  dom.reviewKind.textContent = descriptor.label;
  dom.reviewStage.innerHTML = getMasteryIconMarkup(card.masteryLevel, true);
  dom.reviewStage.className = `badge badge-review-stage badge-${MASTERY_LEVELS[card.masteryLevel]}`;
  dom.reviewStage.title = formatMasteryLabel(card.masteryLevel);
  dom.reviewStage.setAttribute("aria-label", formatMasteryLabel(card.masteryLevel));
  dom.reviewDue.textContent = formatDue(card.due);
  dom.reviewPrompt.textContent = challenge.prompt;
  dom.reviewFront.textContent = challenge.front;
  dom.reviewFront.className = `review-front${challenge.frontClass ? ` ${challenge.frontClass}` : ""}`;
  renderChallengeContext(dom.reviewContext, challenge, Boolean(check));

  if (!check) {
    dom.answerInput.disabled = false;
    dom.answerInput.value = "";
    dom.answerSubmitButton.disabled = false;
    dom.answerSubmitButton.textContent = "Enter";
    dom.reviewAdoptWrap?.classList.add("is-hidden");
    dom.reviewEasyWrap.classList.add("is-hidden");
    dom.reviewEasyButton.classList.remove("is-active");
    dom.reviewEasyButton.setAttribute("aria-pressed", "false");
    dom.reviewTypoWrap.classList.add("is-hidden");
    dom.reviewFeedback.className = "feedback is-empty";
    dom.reviewFeedback.textContent = "";
    focusInputIfSectionVisible(dom.answerInput);
    return;
  }

  const preview = previewReview(card, check.autoGrade, check.correct);
  dom.answerInput.disabled = true;
  dom.answerSubmitButton.disabled = false;
  dom.answerSubmitButton.textContent = "Next Card";
  dom.reviewEasyWrap.classList.toggle("is-hidden", !check.correct);
  dom.reviewEasyButton.classList.toggle("is-active", check.autoGrade === "easy");
  dom.reviewEasyButton.setAttribute("aria-pressed", check.autoGrade === "easy" ? "true" : "false");
  dom.reviewTypoWrap.classList.toggle("is-hidden", check.correct);
  dom.reviewAdoptWrap?.classList.toggle("is-hidden", check.correct);

  if (check.correct) {
    dom.reviewFeedback.className = "feedback is-correct";
    dom.reviewFeedback.textContent = `Correct. ${challenge.reveal}. Next: ${formatInterval(preview.scheduledDays)}. Press Enter again.`;
  } else {
    dom.reviewFeedback.className = "feedback is-incorrect";
    dom.reviewFeedback.textContent = `Not quite. ${challenge.reveal}. Next: ${formatInterval(preview.scheduledDays)}. Press Enter again.`;
  }

  dom.answerSubmitButton.focus();
}

function handleManualTrainingRefresh() {
  renderStats();
  renderReviewDueQueue();
  if (!hasPendingReviewInteraction()) {
    renderReviewPanel();
  }
  renderBackupButton();

  const dueCount = getDueCardIds().length;
  setStatus(dueCount
    ? `Training Grounds refreshed. ${dueCount} due card${dueCount === 1 ? "" : "s"} ready.`
    : "Training Grounds refreshed. No cards are due right now.", "info");
}

function hasPendingReviewInteraction() {
  return Boolean(
    state.currentCardId ||
    state.currentChallenge ||
    state.lastCheck
  );
}

function toggleReviewQueue() {
  state.reviewQueueOpen = !state.reviewQueueOpen;
  renderReviewDueQueue();
}

function shuffleReviewQueue() {
  if (state.lastCheck) {
    setStatus("Finish the checked answer before shuffling due cards.", "info");
    return;
  }

  const dueIds = getSortedDueCardIds();
  const dueGroups = getCardGroups(dueIds);
  if (dueGroups.length < 2) {
    state.reviewOrderIds = dueIds;
    renderReviewDueQueue();
    setStatus(dueGroups.length ? "Only one kanji has due cards right now, so there is nothing to shuffle." : "No cards are due right now.", "info");
    return;
  }

  const activeId = state.currentCardId && dueIds.includes(state.currentCardId) ? state.currentCardId : null;
  const activeGroup = activeId ? dueGroups.find((group) => group.cardIds.includes(activeId)) : null;
  const groupsToShuffle = activeGroup ? dueGroups.filter((group) => group !== activeGroup) : dueGroups;
  const shuffledGroups = activeGroup
    ? [...shuffle(groupsToShuffle), withActiveCardFirst(activeGroup, activeId)]
    : shuffle(dueGroups);

  state.reviewOrderIds = shuffledGroups.flatMap((group) => group.cardIds);
  clearActiveReview();
  render();
  setStatus(activeGroup
    ? `Shuffled ${dueGroups.length} due kanji bundle${dueGroups.length === 1 ? "" : "s"} and moved the current one later.`
    : `Shuffled ${dueGroups.length} due kanji bundle${dueGroups.length === 1 ? "" : "s"}. Due dates and SRS levels were not changed.`, "info");
}

function renderReviewDueQueue() {
  dom.reviewQueueToggle.textContent = "👁️ due";
  dom.reviewQueueToggle.setAttribute("aria-label", state.reviewQueueOpen ? "Hide due queue" : "Show due queue");
  dom.reviewQueueToggle.setAttribute("aria-expanded", String(state.reviewQueueOpen));
  if (dom.refreshTrainingButton) {
    dom.refreshTrainingButton.textContent = "↻ refresh";
  }

  const dueCardCount = getDueCardIds().length;
  const dueKanjiCount = getCardGroups(getSortedDueCardIds()).length;
  if (dom.reviewRemainingCount) {
    dom.reviewRemainingCount.textContent = `${dueCardCount} review${dueCardCount === 1 ? "" : "s"} left`;
  }
  dom.shuffleDueButton.disabled = dueKanjiCount < 2;

  if (!state.reviewQueueOpen) {
    dom.reviewQueuePanel.classList.add("is-hidden");
    dom.reviewQueuePanel.innerHTML = "";
    return;
  }

  dom.reviewQueuePanel.classList.remove("is-hidden");
  const groups = getReviewQueueGroups();
  if (!groups.length) {
    dom.reviewQueuePanel.innerHTML = "<p class='due-queue-empty'>No cards are scheduled yet. Add kanji to start filling the training grounds.</p>";
    return;
  }

  dom.reviewQueuePanel.innerHTML = groups.map((group) => `
    <section class="due-queue-group">
      <h3>${escapeHtml(group.label)}</h3>
      <div class="due-kanji-list">
        ${group.items.map((item) => `<span class="due-kanji" title="${escapeHtml(item.title)}">${escapeHtml(item.char)}</span>`).join("")}
      </div>
    </section>
  `).join("");
}

function getReviewQueueGroups() {
  const items = getReviewQueueItems();
  const groups = new Map();

  items.forEach((item) => {
    const bucket = getDueBucket(item.due);
    if (!groups.has(bucket.key)) {
      groups.set(bucket.key, {
        label: bucket.label,
        sort: bucket.sort,
        items: []
      });
    }
    groups.get(bucket.key).items.push(item);
  });

  return [...groups.values()].sort((left, right) => left.sort - right.sort);
}

function getReviewQueueItems() {
  const kanjiMap = new Map(state.deck.kanji.map((entry) => [entry.id, entry]));
  const includedKanjiIds = new Set();
  const dueItems = getCardGroups(getDueCardIds())
    .map((group, index) => buildReviewQueueItem(group, kanjiMap, index))
    .filter(Boolean);

  dueItems.forEach((item) => includedKanjiIds.add(item.kanjiId));

  const futureCardIds = state.descriptors
    .map((descriptor) => state.deck.progress.cards[descriptor.id])
    .filter((card) => {
      const descriptor = card ? state.descriptorMap.get(card.id) : null;
      return card && descriptor && !isDue(card.due) && !includedKanjiIds.has(descriptor.kanjiId);
    })
    .sort(compareReviewCards)
    .map((card) => card.id);

  const futureItems = getCardGroups(futureCardIds)
    .map((group, index) => buildReviewQueueItem(group, kanjiMap, dueItems.length + index))
    .filter(Boolean);

  return [...dueItems, ...futureItems];
}

function buildReviewQueueItem(group, kanjiMap, order) {
  const kanji = kanjiMap.get(group.kanjiId);
  const cards = group.cardIds.map((cardId) => state.deck.progress.cards[cardId]).filter(Boolean);
  if (!kanji || !cards.length) {
    return null;
  }

  const earliestCard = cards.reduce((earliest, card) => compareReviewCards(card, earliest) < 0 ? card : earliest, cards[0]);
  const labels = uniqueList(group.cardIds
    .map((cardId) => state.descriptorMap.get(cardId))
    .filter(Boolean)
    .map((descriptor) => descriptor.label));

  return {
    kanjiId: group.kanjiId,
    char: kanji.char,
    title: `${kanji.char} - ${labels.join(", ")}`,
    due: earliestCard.due,
    order
  };
}

function renderLibrary() {
  syncLibraryControls();
  dom.kanjiList.innerHTML = "";

  const filteredKanji = getFilteredKanji();
  if (!filteredKanji.length) {
    const emptyState = document.createElement("div");
    emptyState.className = "empty-state";
    emptyState.innerHTML = "<p class='empty-title'>No kanji match this search.</p><p class='empty-copy'>Try a different term or adjust the hoard filters.</p>";
    dom.kanjiList.appendChild(emptyState);
    if (state.libraryMultiSelect) {
      renderMultiSelectSummary();
    }
    return;
  }

  if (!state.libraryMultiSelect && state.selectedKanjiId && !filteredKanji.some((entry) => entry.id === state.selectedKanjiId)) {
    state.selectedKanjiId = null;
  }

  if (state.libraryView === "list") {
    renderLibraryListView(filteredKanji);
    return;
  }

  renderLibraryGridView(filteredKanji);
}

function syncLibraryControls() {
  dom.libraryControls.querySelectorAll("[data-library-view]").forEach((button) => {
    const isActive = button.dataset.libraryView === state.libraryView;
    button.classList.toggle("is-active", isActive);
    button.setAttribute("aria-pressed", isActive ? "true" : "false");
  });

  dom.libraryControls.querySelectorAll("[data-mastery-filter]").forEach((button) => {
    const isActive = button.dataset.masteryFilter === state.masteryFilter;
    button.classList.toggle("is-active", isActive);
    button.setAttribute("aria-pressed", isActive ? "true" : "false");
  });

  syncLibraryMetaFilterOptions();

  const filteredCount = getFilteredKanji().length;
  dom.libraryControls.querySelectorAll("[data-bulk-cram]").forEach((button) => {
    button.disabled = filteredCount === 0;
  });
  dom.exportFilteredKanjiDataButton.disabled = filteredCount === 0;
  dom.exportAllKanjiDataButton.disabled = state.deck.kanji.length === 0;

  const selectedCount = state.multiSelectKanjiIds.length;
  const selectModeButton = dom.libraryControls.querySelector("[data-library-select-mode]");
  if (selectModeButton) {
    selectModeButton.classList.toggle("is-active", state.libraryMultiSelect);
    selectModeButton.setAttribute("aria-pressed", state.libraryMultiSelect ? "true" : "false");
    selectModeButton.textContent = state.libraryMultiSelect
      ? `Selecting (${selectedCount})`
      : `Select Kanji (${selectedCount})`;
  }

  dom.libraryControls.querySelectorAll("[data-selected-cram]").forEach((button) => {
    button.disabled = selectedCount === 0;
  });
  dom.exportSelectedKanjiDataButton.disabled = selectedCount === 0;
}

function getFilteredKanji() {
  const query = parseLibrarySearch(dom.searchInput.value);
  return state.deck.kanji.filter((entry) => {
    if (!matchesSearch(entry, query)) {
      return false;
    }

    const summary = buildKanjiSummary(entry);
    if (state.masteryFilter === "leech") {
      if (!hasLeechCardsForKanji(entry.id)) {
        return false;
      }
    } else if (state.masteryFilter !== "all" && MASTERY_LEVELS[summary.overallLevel] !== state.masteryFilter) {
      return false;
    }

    if (state.jlptFilter !== "all" && safeString(entry.jlpt) !== state.jlptFilter) {
      return false;
    }

    if (state.gradeFilter !== "all" && safeString(entry.grade) !== state.gradeFilter) {
      return false;
    }

    return true;
  }).reverse();
}

function syncLibraryMetaFilterOptions() {
  syncLibraryMetaSelect(dom.libraryJlptFilter, {
    allLabel: "All JLPT",
    selected: state.jlptFilter,
    values: getSortedMetaValues("jlpt"),
    labelForValue: (value) => /^n\d$/i.test(value) ? `JLPT ${value.toUpperCase()}` : `JLPT N${value}`
  });

  syncLibraryMetaSelect(dom.libraryGradeFilter, {
    allLabel: "All Grades",
    selected: state.gradeFilter,
    values: getSortedMetaValues("grade"),
    labelForValue: (value) => `Grade ${value}`
  });
}

function syncLibraryMetaSelect(select, options) {
  if (!select) {
    return;
  }

  const values = Array.isArray(options.values) ? options.values : [];
  const normalizedSelected = values.includes(options.selected) ? options.selected : "all";
  const nextOptions = [
    `<option value="all">${escapeHtml(options.allLabel)}</option>`,
    ...values.map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(options.labelForValue(value))}</option>`)
  ].join("");

  if (select.innerHTML !== nextOptions) {
    select.innerHTML = nextOptions;
  }

  select.value = normalizedSelected;

  if (select === dom.libraryJlptFilter) {
    state.jlptFilter = normalizedSelected;
  } else if (select === dom.libraryGradeFilter) {
    state.gradeFilter = normalizedSelected;
  }
}

function getSortedMetaValues(field) {
  return uniqueList(state.deck.kanji.map((entry) => safeString(entry[field])).filter(Boolean))
    .sort(compareMetaFilterValue);
}

function compareMetaFilterValue(left, right) {
  const leftNumber = parseNumericFilterValue(left);
  const rightNumber = parseNumericFilterValue(right);

  if (Number.isFinite(leftNumber) && Number.isFinite(rightNumber) && leftNumber !== rightNumber) {
    return leftNumber - rightNumber;
  }

  if (Number.isFinite(leftNumber) !== Number.isFinite(rightNumber)) {
    return Number.isFinite(leftNumber) ? -1 : 1;
  }

  return safeString(left).localeCompare(safeString(right), "en", { numeric: true, sensitivity: "base" });
}

function parseNumericFilterValue(value) {
  const match = safeString(value).match(/\d+/);
  return match ? Number(match[0]) : Number.NaN;
}

function buildKanjiSummary(entry) {
  const cards = getCardsForKanji(entry.id);
  const dueCount = cards.filter((card) => isDue(card.due)).length;
  const nextDue = cards.length
    ? cards
        .map((card) => card.due)
        .sort((left, right) => new Date(left).getTime() - new Date(right).getTime())[0]
    : null;
  const overallLevel = cards.length
    ? Math.round(cards.reduce((sum, card) => sum + card.masteryLevel, 0) / cards.length)
    : 0;

  return {
    cards,
    dueCount,
    nextDue,
    dueLabel: dueCount ? `${dueCount} due` : nextDue ? formatDue(nextDue) : "Not scheduled",
    overallLevel
  };
}

function isLeechCard(card) {
  return clampNumber(card && card.consecutiveFailures, 0, 100000, 0) > LEECH_THRESHOLD;
}

function hasLeechCardsForKanji(kanjiId) {
  return getCardsForKanji(kanjiId).some((card) => isLeechCard(card));
}

function renderLibraryGridView(filteredKanji) {
  const topToggleWrap = document.createElement("div");
  topToggleWrap.className = "library-grid-toggle-wrap library-grid-toggle-wrap-top";
  topToggleWrap.hidden = true;
  const topToggleButton = document.createElement("button");
  topToggleButton.type = "button";
  topToggleButton.className = "button button-small library-grid-toggle";
  topToggleButton.dataset.toggleLibraryGrid = "true";
  topToggleWrap.appendChild(topToggleButton);
  const gridShell = document.createElement("div");
  gridShell.className = "kanji-grid-shell";
  const grid = document.createElement("div");
  grid.className = "kanji-grid";
  const bottomToggleWrap = document.createElement("div");
  bottomToggleWrap.className = "library-grid-toggle-wrap library-grid-toggle-wrap-bottom";
  bottomToggleWrap.hidden = true;
  const bottomToggleButton = document.createElement("button");
  bottomToggleButton.type = "button";
  bottomToggleButton.className = "button button-small library-grid-toggle";
  bottomToggleButton.dataset.toggleLibraryGrid = "true";
  bottomToggleWrap.appendChild(bottomToggleButton);

  filteredKanji.forEach((entry) => {
    const isMultiSelected = state.multiSelectKanjiIds.includes(entry.id);
    const tile = document.createElement("button");
    tile.type = "button";
    tile.className = "kanji-tile";
    tile.dataset.selectKanji = entry.id;
    tile.setAttribute("aria-label", state.libraryMultiSelect ? `${isMultiSelected ? "Deselect" : "Select"} ${entry.char} for cram` : `Open details for ${entry.char}`);
    if (state.libraryMultiSelect) {
      tile.classList.add("is-multi-select");
      tile.setAttribute("aria-pressed", isMultiSelected ? "true" : "false");
    }
    if (state.libraryMultiSelect && isMultiSelected) {
      tile.classList.add("is-multi-selected");
    } else if (!state.libraryMultiSelect && state.selectedKanjiId === entry.id) {
      tile.classList.add("is-selected");
    }

    tile.innerHTML = `${state.libraryMultiSelect ? `<span class="multi-select-mark">${isMultiSelected ? "In" : "+"}</span>` : ""}<span class="kanji-tile-glyph">${escapeHtml(entry.char)}</span>`;
    grid.appendChild(tile);
  });

  dom.kanjiList.appendChild(topToggleWrap);
  gridShell.appendChild(grid);
  dom.kanjiList.appendChild(gridShell);
  dom.kanjiList.appendChild(bottomToggleWrap);
  if (state.libraryMultiSelect) {
    renderMultiSelectSummary();
    queueGridCollapseState(gridShell, grid, topToggleWrap, topToggleButton, bottomToggleWrap, bottomToggleButton, 20);
    return;
  }

  renderLibrarySelection(filteredKanji, "Click a kanji tile to open its details.");
  queueGridCollapseState(gridShell, grid, topToggleWrap, topToggleButton, bottomToggleWrap, bottomToggleButton, 20);
}

function renderLibraryListView(filteredKanji) {
  const tableShell = document.createElement("div");
  tableShell.className = "library-table-shell";

  const table = document.createElement("table");
  table.className = "kanji-table";
  table.innerHTML = `
    <thead>
      <tr>
        <th>Kanji</th>
        <th>Meaning</th>
        <th>On Reading</th>
        <th>Kun Reading</th>
        <th>Cards</th>
        <th>Due</th>
        <th>Stage</th>
      </tr>
    </thead>
    <tbody></tbody>
  `;

  const tbody = table.querySelector("tbody");
  filteredKanji.forEach((entry) => {
    const summary = buildKanjiSummary(entry);
    const isMultiSelected = state.multiSelectKanjiIds.includes(entry.id);
    const row = document.createElement("tr");
    row.className = "kanji-table-row";
    row.dataset.selectKanji = entry.id;
    if (state.libraryMultiSelect) {
      row.classList.add("is-multi-select");
      row.setAttribute("aria-pressed", isMultiSelected ? "true" : "false");
    }
    if (state.libraryMultiSelect && isMultiSelected) {
      row.classList.add("is-multi-selected");
    } else if (!state.libraryMultiSelect && state.selectedKanjiId === entry.id) {
      row.classList.add("is-selected");
    }

    row.innerHTML = `
      <td><span class="kanji-table-glyph">${escapeHtml(entry.char)}</span>${state.libraryMultiSelect ? `<span class="multi-select-mark multi-select-mark-table">${isMultiSelected ? "In" : "+"}</span>` : ""}</td>
      <td>${escapeHtml(entry.meanings.join(", ") || "No meanings yet")}</td>
      <td>${escapeHtml(entry.onyomi.join(", ") || "None")}</td>
      <td>${escapeHtml(entry.kunyomi.join(", ") || "None")}</td>
      <td>${summary.cards.length}</td>
      <td>${escapeHtml(summary.dueLabel)}</td>
      <td><span class="badge badge-${MASTERY_LEVELS[summary.overallLevel]}">${formatMasteryDisplayMarkup(summary.overallLevel)}</span></td>
    `;
    tbody.appendChild(row);
  });

  tableShell.appendChild(table);
  dom.kanjiList.appendChild(tableShell);
  applyTableRowLimit(tableShell, table, 5);
  if (state.libraryMultiSelect) {
    renderMultiSelectSummary();
    return;
  }

  renderLibrarySelection(filteredKanji, "Click any kanji row to open its details.");
}

function renderLibrarySelection(filteredKanji, emptyCopy) {
  const selection = document.createElement("section");
  selection.className = "library-selection";

  if (!state.selectedKanjiId) {
    selection.innerHTML = `<div class='empty-state'><p class='empty-title'>No kanji selected.</p><p class='empty-copy'>${escapeHtml(emptyCopy)}</p></div>`;
    dom.kanjiList.appendChild(selection);
    return;
  }

  const selectedEntry = filteredKanji.find((entry) => entry.id === state.selectedKanjiId);
  if (!selectedEntry) {
    selection.innerHTML = "<div class='empty-state'><p class='empty-title'>That kanji is outside the current filter.</p><p class='empty-copy'>Choose another entry or change the filter to see it again.</p></div>";
    dom.kanjiList.appendChild(selection);
    return;
  }

  const label = document.createElement("p");
  label.className = "library-selection-title";
  label.textContent = "Selected Kanji";
  selection.append(label, createKanjiDetailPanel(selectedEntry));
  dom.kanjiList.appendChild(selection);
}

function renderMultiSelectSummary() {
  const selectedEntries = getMultiSelectedKanjiEntries();
  const selection = document.createElement("section");
  selection.className = "library-selection multi-select-summary";

  const label = document.createElement("p");
  label.className = "library-selection-title";
  label.textContent = `Cram selection: ${selectedEntries.length} kanji`;

  if (!selectedEntries.length) {
    selection.innerHTML = `<p class='library-selection-title'>Cram selection: 0 kanji</p><div class='empty-state'><p class='empty-title'>Select mode is on.</p><p class='empty-copy'>Click kanji to choose an exact cram set. Search and filters can change; your selections will stay until you clear them or add them to cram.</p></div>`;
    dom.kanjiList.appendChild(selection);
    return;
  }

  const chipList = document.createElement("div");
  chipList.className = "multi-select-chips";
  selectedEntries.forEach((entry) => {
    const chip = document.createElement("button");
    chip.type = "button";
    chip.className = "multi-select-chip";
    chip.dataset.toggleMultiSelect = entry.id;
    chip.setAttribute("aria-label", `Remove ${entry.char} from cram selection`);
    chip.innerHTML = `<span class="cram-chip-glyph">${escapeHtml(entry.char)}</span><span>${escapeHtml(entry.meanings[0] || "No meaning")}</span><span aria-hidden="true">x</span>`;
    chipList.appendChild(chip);
  });

  const hint = document.createElement("p");
  hint.className = "kanji-subtle";
  hint.textContent = "Selected kanji stay selected while you change filters or search. Click a chip to remove it.";

  selection.append(label, chipList, hint);
  dom.kanjiList.appendChild(selection);
}

function createKanjiDetailPanel(entry) {
  const summary = buildKanjiSummary(entry);
  const isInCram = state.cramKanjiIds.includes(entry.id);
  const panel = document.createElement("article");
  panel.className = "library-detail-panel";

  const top = document.createElement("div");
  top.className = "library-detail-head";

  const main = document.createElement("div");
  main.className = "library-detail-main";

  const glyph = document.createElement("div");
  glyph.className = "kanji-glyph";
  glyph.textContent = entry.char;

  const meta = document.createElement("div");
  meta.className = "kanji-meta";

  const title = document.createElement("strong");
  title.textContent = entry.meanings.join(", ") || "No meanings yet";

  const subtitle = document.createElement("div");
  subtitle.className = "kanji-subtle";
  subtitle.textContent = `${entry.onyomi.length} onyomi | ${entry.kunyomi.length} kunyomi | ${countExampleWords(entry)} example words`;

  meta.append(title, subtitle);
  main.append(glyph, meta);

  const actions = document.createElement("div");
  actions.className = "library-detail-actions";
  actions.innerHTML = `
    <button class="button button-small" type="button" data-close-details>Close Details</button>
    <button class="button button-small" type="button" data-toggle-cram="${entry.id}">${isInCram ? "Remove From Cram" : "Add To Cram"}</button>
    <span class="badge badge-${MASTERY_LEVELS[summary.overallLevel]}">${formatMasteryDisplayMarkup(summary.overallLevel)}</span>
    <button class="button button-small" type="button" data-action="edit" data-kanji-id="${entry.id}">Edit</button>
    <button class="button button-small" type="button" data-action="delete" data-kanji-id="${entry.id}">Delete</button>
  `;

  top.append(main, actions);

  const tags = document.createElement("div");
  tags.className = "kanji-tags";
  tags.append(
    createTag(`Due now: ${summary.dueCount}`),
    createTag(`Cards: ${summary.cards.length}`),
    createTag(summary.dueLabel),
    createTag(`On words: ${entry.words.on.length}`),
    createTag(`Kun words: ${entry.words.kun.length}`),
    createTag(`Irregular words: ${entry.words.irregular.length}`)
  );

  if (isInCram) {
    tags.append(createTag("In cram list"));
  }

  const detailGrid = document.createElement("div");
  detailGrid.className = "kanji-detail-grid";
  detailGrid.append(
    createDetailBlock("Meanings", entry.meanings),
    createDetailBlock("Onyomi", entry.onyomi),
    createDetailBlock("Kunyomi", entry.kunyomi),
    createDetailBlock("JLPT", entry.jlpt || "None"),
    createDetailBlock("Grade", entry.grade || "None"),
    createDetailBlock("Mnemonic", entry.mnemonic || "None"),
    createWordBlock("On Words", entry.words.on),
    createWordBlock("Kun Words", entry.words.kun),
    createWordBlock("Irregular Words", entry.words.irregular)
  );

  panel.append(top, tags, detailGrid);
  return panel;
}

function createDetailBlock(label, values) {
  const block = document.createElement("div");
  block.className = "kanji-detail-block";
  const lines = Array.isArray(values) ? values.filter(Boolean) : [values];
  const list = document.createElement("div");
  list.className = "kanji-detail-lines";

  if (lines.length) {
    lines.forEach((line) => {
      const item = document.createElement("div");
      item.className = "kanji-detail-line";
      item.textContent = line;
      list.appendChild(item);
    });
  } else {
    const empty = document.createElement("div");
    empty.className = "kanji-detail-line";
    empty.textContent = "None yet";
    list.appendChild(empty);
  }

  const labelNode = document.createElement("span");
  labelNode.className = "kanji-detail-label";
  labelNode.textContent = label;
  block.append(labelNode, list);
  return block;
}

function createWordBlock(label, words) {
  const block = document.createElement("div");
  block.className = "kanji-detail-block";

  const labelNode = document.createElement("span");
  labelNode.className = "kanji-detail-label";
  labelNode.textContent = label;

  const list = document.createElement("div");
  list.className = "kanji-detail-lines";

  if (Array.isArray(words) && words.length) {
    words.forEach((word) => {
      const item = document.createElement("div");
      item.className = "kanji-detail-line";
      item.textContent = formatWordLine(word);
      list.appendChild(item);
    });
  } else {
    const empty = document.createElement("div");
    empty.className = "kanji-detail-line";
    empty.textContent = "None yet";
    list.appendChild(empty);
  }

  block.append(labelNode, list);
  return block;
}

function formatWordLine(word) {
  const parts = [];
  if (word.word) {
    parts.push(word.word);
  }
  if (word.reading) {
    parts.push(`(${word.reading})`);
  }
  if (word.definition) {
    parts.push(`- ${word.definition}`);
  }
  return parts.join(" ");
}

function formatWordSummaryLine(word) {
  const parts = [];
  if (word.reading) {
    parts.push(`(${word.reading})`);
  }
  if (word.definition) {
    parts.push(word.definition);
  }
  return parts.join(" - ") || "No reading or definition yet";
}

function applyTableRowLimit(shell, table, rowCount) {
  shell.style.maxHeight = "";
  const rows = Array.from(table.querySelectorAll("tbody tr"));
  if (rows.length <= rowCount) {
    return;
  }

  const headHeight = table.querySelector("thead") ? table.querySelector("thead").offsetHeight : 0;
  const visibleRows = rows.slice(0, rowCount);
  const totalHeight = headHeight + visibleRows.reduce((sum, row) => sum + row.offsetHeight, 0) + 2;
  shell.style.maxHeight = `${Math.ceil(totalHeight)}px`;
}

function applyGridCollapseState(shell, grid, topToggleWrap, topToggleButton, bottomToggleWrap, bottomToggleButton, visibleCount) {
  shell.classList.remove("is-collapsed", "is-expanded");
  const tiles = Array.from(grid.children);
  tiles.forEach((tile) => {
    tile.hidden = false;
  });
  if (tiles.length <= visibleCount) {
    topToggleWrap.hidden = true;
    bottomToggleWrap.hidden = true;
    return;
  }

  if (state.libraryGridExpanded) {
    topToggleWrap.hidden = false;
    topToggleWrap.classList.add("is-sticky");
    topToggleButton.textContent = "Collapse";
    topToggleButton.setAttribute("aria-expanded", "true");
    bottomToggleWrap.hidden = true;
    bottomToggleWrap.classList.remove("is-sticky");
    shell.classList.add("is-expanded");
    return;
  }

  topToggleWrap.hidden = true;
  topToggleWrap.classList.remove("is-sticky");
  bottomToggleWrap.hidden = false;
  bottomToggleWrap.classList.remove("is-sticky");
  bottomToggleButton.textContent = "Show All Kanji";
  bottomToggleButton.setAttribute("aria-expanded", "false");
  tiles.forEach((tile, index) => {
    tile.hidden = index >= visibleCount;
  });
  shell.classList.add("is-collapsed");
}

function queueGridCollapseState(shell, grid, topToggleWrap, topToggleButton, bottomToggleWrap, bottomToggleButton, visibleCount) {
  applyGridCollapseState(shell, grid, topToggleWrap, topToggleButton, bottomToggleWrap, bottomToggleButton, visibleCount);
}

let libraryGridResizeFrame = 0;
function handleLibraryGridResize() {
  if (state.libraryView !== "grid") {
    return;
  }

  const shell = dom.kanjiList.querySelector(".kanji-grid-shell");
  const grid = dom.kanjiList.querySelector(".kanji-grid");
  const topToggleWrap = dom.kanjiList.querySelector(".library-grid-toggle-wrap-top");
  const bottomToggleWrap = dom.kanjiList.querySelector(".library-grid-toggle-wrap-bottom");
  const topToggleButton = topToggleWrap?.querySelector("[data-toggle-library-grid]");
  const bottomToggleButton = bottomToggleWrap?.querySelector("[data-toggle-library-grid]");

  if (!shell || !grid || !topToggleWrap || !bottomToggleWrap || !topToggleButton || !bottomToggleButton) {
    return;
  }

  if (libraryGridResizeFrame) {
    cancelAnimationFrame(libraryGridResizeFrame);
  }

  libraryGridResizeFrame = requestAnimationFrame(() => {
    libraryGridResizeFrame = 0;
    applyGridCollapseState(shell, grid, topToggleWrap, topToggleButton, bottomToggleWrap, bottomToggleButton, 20);
  });
}

function renderBackupButton() {
  dom.restoreBackupButton.disabled = !(Boolean(localStorage.getItem(STORAGE_KEY)) || "indexedDB" in window);
}

function createTag(text) {
  const tag = document.createElement("span");
  tag.className = "tag";
  tag.textContent = text;
  return tag;
}

function handleCramSelectionClick(event) {
  const removeButton = event.target.closest("[data-remove-cram]");
  if (!removeButton) {
    return;
  }

  toggleCramKanji(removeButton.dataset.removeCram);
}

function toggleCramKanji(kanjiId) {
  if (!kanjiId) {
    return;
  }

  if (state.cramKanjiIds.includes(kanjiId)) {
    state.cramKanjiIds = state.cramKanjiIds.filter((id) => id !== kanjiId);
    clearCramSession();
    render();
    setStatus("Removed that kanji from the cram list.", "info");
    return;
  }

  state.cramKanjiIds = [...state.cramKanjiIds, kanjiId];
  clearCramSession();
  render();
  setStatus("Added that kanji to the cram list.", "info");
}

function addKanjiIdsToCram(kanjiIds, successMessage, duplicateMessage = "Those kanji were already in the cram list.") {
  const ids = uniqueList((Array.isArray(kanjiIds) ? kanjiIds : []).filter((kanjiId) => state.deck.kanji.some((entry) => entry.id === kanjiId)));
  if (!ids.length) {
    return 0;
  }

  const beforeCount = state.cramKanjiIds.length;
  state.cramKanjiIds = Array.from(new Set([...state.cramKanjiIds, ...ids]));
  const addedCount = state.cramKanjiIds.length - beforeCount;
  clearCramSession();
  render();

  if (addedCount === 0) {
    setStatus(duplicateMessage, "info");
    return 0;
  }

  setStatus(successMessage || `Added ${addedCount} kanji to the cram list.`, "success");
  return addedCount;
}

function toggleMultiSelectKanji(kanjiId) {
  if (!kanjiId) {
    return;
  }

  if (state.multiSelectKanjiIds.includes(kanjiId)) {
    state.multiSelectKanjiIds = state.multiSelectKanjiIds.filter((id) => id !== kanjiId);
  } else {
    state.multiSelectKanjiIds = [...state.multiSelectKanjiIds, kanjiId];
  }

  renderLibrary();
}

function getMultiSelectedKanjiEntries() {
  const kanjiById = new Map(state.deck.kanji.map((entry) => [entry.id, entry]));
  return state.multiSelectKanjiIds
    .map((kanjiId) => kanjiById.get(kanjiId))
    .filter(Boolean);
}

function addSelectedKanjiToCram() {
  const selectedIds = state.multiSelectKanjiIds.filter((kanjiId) => state.deck.kanji.some((entry) => entry.id === kanjiId));
  if (!selectedIds.length) {
    setStatus("Select one or more kanji first.", "info");
    return;
  }

  const uniqueSelectedIds = uniqueList(selectedIds);
  const addedCount = addKanjiIdsToCram(uniqueSelectedIds, `Added ${uniqueSelectedIds.length} selected kanji to the cram list.`, "Those selected kanji were already in the cram list.");
  state.multiSelectKanjiIds = [];
  state.libraryMultiSelect = false;

  if (addedCount === 0) {
    renderLibrary();
    return;
  }

  renderLibrary();
}

function clearMultiSelectKanji() {
  state.multiSelectKanjiIds = [];
  renderLibrary();
  setStatus("Cleared the selected kanji.", "info");
}

function addFilteredKanjiToCram() {
  const filteredIds = getFilteredKanji().map((entry) => entry.id);
  if (!filteredIds.length) {
    setStatus("No kanji match the current search and filter.", "info");
    return;
  }

  const uniqueFilteredIds = uniqueList(filteredIds);
  addKanjiIdsToCram(uniqueFilteredIds, `Added ${uniqueFilteredIds.length} filtered kanji to the cram list.`, "Those filtered kanji were already in the cram list.");
}

function removeFilteredKanjiFromCram() {
  const filteredIds = new Set(getFilteredKanji().map((entry) => entry.id));
  if (!filteredIds.size) {
    setStatus("No kanji match the current search and filter.", "info");
    return;
  }

  const beforeCount = state.cramKanjiIds.length;
  state.cramKanjiIds = state.cramKanjiIds.filter((kanjiId) => !filteredIds.has(kanjiId));
  const removedCount = beforeCount - state.cramKanjiIds.length;
  clearCramSession();
  render();

  if (removedCount === 0) {
    setStatus("None of the filtered kanji were in the cram list.", "info");
    return;
  }

  setStatus(`Removed ${removedCount} filtered kanji from the cram list.`, "info");
}

function clearCramList() {
  state.cramKanjiIds = [];
  state.lastRandomCramKanjiIds = [];
  clearCramSession();
  render();
  setStatus("Cleared the cram list.", "info");
}

function addRandomKanjiToCram() {
  const allIds = state.deck.kanji.map((entry) => entry.id);
  if (!allIds.length) {
    setStatus("Your hoard is empty, so there is nothing random to add yet.", "info");
    return;
  }

  const drawCount = Math.min(10, allIds.length);
  let eligibleIds = allIds.filter((kanjiId) => !state.lastRandomCramKanjiIds.includes(kanjiId));
  if (eligibleIds.length < drawCount) {
    eligibleIds = allIds.slice();
  }

  const drawnIds = shuffle(eligibleIds).slice(0, drawCount);
  state.lastRandomCramKanjiIds = drawnIds.slice();
  state.cramKanjiIds = drawnIds.slice();
  clearCramSession();
  renderLibrary();
  startCramSession();
}

function renderCramPanel() {
  renderCramSelection();

  dom.startCramButton.disabled = !state.cramKanjiIds.length;
  dom.clearCramButton.disabled = !state.cramKanjiIds.length;
  if (dom.cramRandomTenButton) {
    dom.cramRandomTenButton.disabled = state.deck.kanji.length === 0;
  }

  const activeCram = resolveCramReview();
  if (!activeCram) {
    dom.cramCard.classList.add("is-hidden");
    dom.cramEmpty.classList.remove("is-hidden");
    dom.cramSubmitButton.textContent = "Enter";
    dom.cramTypoWrap?.classList.add("is-hidden");
    dom.cramAdoptWrap?.classList.add("is-hidden");
    dom.cramFeedback.className = "feedback";
    dom.cramFeedback.textContent = "";
    syncCramModeToggle();
    return;
  }

  const { descriptor, challenge } = activeCram;
  const card = state.deck.progress.cards[descriptor.id];
  const check = getCurrentCramCheck(descriptor.id);

  dom.cramEmpty.classList.add("is-hidden");
  dom.cramCard.classList.remove("is-hidden");
  dom.cramKind.textContent = descriptor.label;
  dom.cramStage.innerHTML = formatMasteryDisplayMarkup(card.masteryLevel);
  dom.cramStage.className = `badge badge-${MASTERY_LEVELS[card.masteryLevel]}`;
  dom.cramPrompt.textContent = challenge.prompt;
  dom.cramFront.textContent = challenge.front;
  dom.cramFront.className = `review-front${challenge.frontClass ? ` ${challenge.frontClass}` : ""}`;
  renderChallengeContext(
    dom.cramContext,
    challenge,
    Boolean(check),
    `${state.cramQueue.length + 1} prompt${state.cramQueue.length === 0 ? "" : "s"} remaining in this session.`
  );

  if (!check) {
    dom.cramInput.disabled = false;
    dom.cramInput.value = "";
    dom.cramSubmitButton.disabled = false;
    dom.cramSubmitButton.textContent = "Enter";
    dom.cramTypoWrap?.classList.add("is-hidden");
    dom.cramAdoptWrap?.classList.add("is-hidden");
    dom.cramFeedback.className = "feedback";
    dom.cramFeedback.textContent = "Practice mode only. Type your answer and press Enter.";
    focusInputIfSectionVisible(dom.cramInput);
    syncCramModeToggle();
    return;
  }

  dom.cramInput.disabled = true;
  dom.cramSubmitButton.disabled = false;
  dom.cramSubmitButton.textContent = "Next Card";
  dom.cramTypoWrap?.classList.toggle("is-hidden", check.correct);
  dom.cramAdoptWrap?.classList.toggle("is-hidden", check.correct);

  if (check.correct) {
    dom.cramFeedback.className = "feedback is-correct";
    dom.cramFeedback.textContent = check.applyToSrs
      ? `Correct. ${challenge.reveal}. This will count as a real review when you move on. Press Enter again.`
      : `Correct. ${challenge.reveal}. Practice only, SRS unchanged. Press Enter again.`;
  } else {
    dom.cramFeedback.className = "feedback is-incorrect";
    dom.cramFeedback.textContent = `Not quite. ${challenge.reveal}. Practice only, SRS unchanged. Press Enter again.`;
  }

  dom.cramSubmitButton.focus();
  syncCramModeToggle();
}

function renderCramSelection() {
  dom.cramSelection.innerHTML = "";
  if (!state.cramKanjiIds.length) {
    const chip = document.createElement("span");
    chip.className = "tag";
    chip.textContent = "Cram list is empty";
    dom.cramSelection.appendChild(chip);
    return;
  }

  state.cramKanjiIds.forEach((kanjiId) => {
    const entry = state.deck.kanji.find((kanji) => kanji.id === kanjiId);
    if (!entry) {
      return;
    }

    const chip = document.createElement("div");
    chip.className = "cram-chip";
    chip.innerHTML = `
      <span class="cram-chip-glyph">${escapeHtml(entry.char)}</span>
      <span>${escapeHtml(entry.meanings[0] || "No meaning")}</span>
      <button class="cram-chip-remove" type="button" data-remove-cram="${entry.id}" aria-label="Remove ${escapeHtml(entry.char)} from cram">x</button>
    `;
    dom.cramSelection.appendChild(chip);
  });
}

function renderReviewWrongTally() {
  renderWrongTally(dom.reviewWrongTally, state.reviewWrongPrompts, "review");
}

function renderCramWrongTally() {
  renderWrongTally(dom.cramWrongTally, state.cramWrongPrompts, "cram");
}

function renderWrongTally(container, items, scope) {
  if (!container) {
    return;
  }

  if (!items.length) {
    container.classList.add("is-hidden");
    container.innerHTML = "";
    return;
  }

  const totalMisses = items.reduce((sum, item) => sum + item.misses, 0);
  const label = scope === "review" ? "Review" : "Cram";
  container.classList.remove("is-hidden");
  container.innerHTML = `
    <div class="session-tally-head">
      <div>
        <p class="session-tally-title">${escapeHtml(label)} misses: ${totalMisses}</p>
        <p class="session-tally-copy">Keep an eye on the cards that tripped you up.</p>
      </div>
      <div class="session-tally-actions">
        <button class="button button-small" type="button" data-wrong-tally-action="add-to-cram" data-wrong-tally-scope="${escapeHtml(scope)}">Add To Cram</button>
        <button class="button button-small" type="button" data-wrong-tally-action="clear" data-wrong-tally-scope="${escapeHtml(scope)}">Clear</button>
      </div>
    </div>
    <div class="session-tally-list">
      ${items.map((item) => `
        <span class="session-tally-chip" title="${escapeHtml(`${item.label} - ${item.prompt}`)}">
          <span class="session-tally-chip-front">${escapeHtml(item.front)}</span>
          <span class="session-tally-chip-label">${escapeHtml(item.label)}</span>
          <span class="session-tally-chip-count">x${item.misses}</span>
        </span>
      `).join("")}
    </div>
  `;
}

function handleReviewWrongTallyClick(event) {
  handleWrongTallyClick(event, "review");
}

function handleCramWrongTallyClick(event) {
  handleWrongTallyClick(event, "cram");
}

function handleWrongTallyClick(event, scope) {
  const actionButton = event.target.closest("[data-wrong-tally-action]");
  if (!actionButton) {
    return;
  }

  const action = actionButton.dataset.wrongTallyAction;
  if (action === "clear") {
    clearWrongPromptTally(scope);
    return;
  }

  if (action === "add-to-cram") {
    addWrongPromptTallyToCram(scope);
  }
}

function clearWrongPromptTally(scope) {
  if (scope === "review") {
    state.reviewWrongPrompts = [];
    renderReviewWrongTally();
  } else {
    state.cramWrongPrompts = [];
    renderCramWrongTally();
  }

  setStatus(`Cleared the ${scope} miss tally.`, "info");
}

function addWrongPromptTallyToCram(scope) {
  const items = scope === "review" ? state.reviewWrongPrompts : state.cramWrongPrompts;
  const kanjiIds = uniqueList(items.map((item) => item.kanjiId).filter(Boolean));
  if (!kanjiIds.length) {
    setStatus(`There are no ${scope} misses to add to cram.`, "info");
    return;
  }

  addKanjiIdsToCram(kanjiIds, `Added ${kanjiIds.length} ${scope} miss kanji to the cram list.`);
}

function rememberWrongPrompt(scope, descriptor, challenge) {
  const collection = scope === "review" ? state.reviewWrongPrompts : state.cramWrongPrompts;
  const key = getChallengeTallyKey(descriptor, challenge);
  const existing = collection.find((item) => item.key === key);
  if (existing) {
    existing.misses += 1;
    existing.lastMissAt = new Date().toISOString();
  } else {
    collection.push({
      key,
      kanjiId: descriptor.kanjiId,
      cardId: descriptor.id,
      front: challenge.front,
      label: descriptor.label,
      prompt: challenge.prompt,
      misses: 1,
      lastMissAt: new Date().toISOString()
    });
  }

  if (scope === "review") {
    renderReviewWrongTally();
  } else {
    renderCramWrongTally();
  }
}

function getChallengeTallyKey(descriptor, challenge) {
  if (challenge.wordId) {
    return `${descriptor.id}::${challenge.wordId}`;
  }
  return descriptor.id;
}

function startCramSession() {
  if (!state.cramKanjiIds.length) {
    setStatus("Add at least one kanji to the cram list first.", "info");
    return;
  }

  state.cramQueue = buildCramQueue();
  state.cramCurrentCardId = null;
  state.cramCurrentChallenge = null;
  state.cramLastCheck = null;

  const active = resolveCramReview();
  renderCramPanel();

  if (!active) {
    setStatus("Those kanji do not have any available prompts yet.", "info");
    return;
  }

  setStatus(
    state.cramDragonMode
      ? "Dragon mode cram started. This practice will not change your SRS levels."
      : "Cram session started. This practice will not change your SRS levels.",
    "info"
  );
}

function buildCramQueue() {
  const queue = state.descriptors
    .filter((descriptor) => state.cramKanjiIds.includes(descriptor.kanjiId))
    .flatMap((descriptor) => buildCramSeedsForDescriptor(descriptor));
  return shuffle(queue);
}

function buildCramSeedsForDescriptor(descriptor) {
  if (!state.cramDragonMode || descriptor.kind === "meaning") {
    return [{ descriptorId: descriptor.id, wordKey: "" }];
  }

  const words = Array.isArray(descriptor.words) ? descriptor.words : [];
  if (!words.length) {
    return [{ descriptorId: descriptor.id, wordKey: "" }];
  }

  return words.map((word) => ({
    descriptorId: descriptor.id,
    wordKey: getWordPromptKey(word)
  }));
}

function resolveCramReview() {
  if (state.cramCurrentCardId && state.cramCurrentChallenge && state.descriptorMap.has(state.cramCurrentCardId)) {
    return {
      descriptor: state.descriptorMap.get(state.cramCurrentCardId),
      challenge: state.cramCurrentChallenge
    };
  }

  const nextSeed = state.cramQueue.shift() || null;
  if (!nextSeed) {
    clearCramSession();
    return null;
  }

  const descriptorId = typeof nextSeed === "string" ? nextSeed : nextSeed.descriptorId;
  const descriptor = state.descriptorMap.get(descriptorId);
  if (!descriptor) {
    return resolveCramReview();
  }

  state.cramCurrentCardId = descriptorId;
  state.cramCurrentChallenge = createChallenge(descriptor, {
    specificWord: resolveCramSeedWord(descriptor, nextSeed)
  });
  state.cramLastCheck = null;

  return {
    descriptor,
    challenge: state.cramCurrentChallenge
  };
}

function resolveCramSeedWord(descriptor, seed) {
  const wordKey = seed && typeof seed === "object" ? safeString(seed.wordKey) : "";
  if (!wordKey) {
    return null;
  }

  const words = Array.isArray(descriptor && descriptor.words) ? descriptor.words : [];
  return words.find((word) => getWordPromptKey(word) === wordKey) || null;
}

function clearCramSession() {
  state.cramQueue = [];
  state.cramCurrentCardId = null;
  state.cramCurrentChallenge = null;
  state.cramLastCheck = null;
}

function toggleCramDragonMode() {
  state.cramDragonMode = !state.cramDragonMode;
  clearCramSession();
  renderCramPanel();
  setStatus(
    state.cramDragonMode
      ? "Dragon mode is on. Cram will review every added word."
      : "Dragon mode is off. Cram will pick one random word per reading type.",
    "info"
  );
}

function syncCramModeToggle() {
  if (!dom.cramDragonModeButton) {
    return;
  }

  dom.cramDragonModeButton.classList.toggle("is-active", state.cramDragonMode);
  dom.cramDragonModeButton.setAttribute("aria-pressed", state.cramDragonMode ? "true" : "false");
}

function getCurrentCramCheck(cardId) {
  if (!state.cramLastCheck || state.cramLastCheck.cardId !== cardId) {
    return null;
  }
  return state.cramLastCheck;
}

function handleCramSubmit(event) {
  event.preventDefault();
  const activeCram = resolveCramReview();

  if (!activeCram) {
    setStatus("No cram prompt is active. Start a cram session from the selected cram list.", "info");
    return;
  }

  if (getCurrentCramCheck(activeCram.descriptor.id)) {
    finalizeCramReview();
    return;
  }

  const correct = isAnswerCorrect(activeCram.challenge, dom.cramInput.value);
  state.cramLastCheck = {
    cardId: activeCram.descriptor.id,
    correct,
    autoGrade: "good",
    applyToSrs: false
  };
  if (!correct) {
    rememberWrongPrompt("cram", activeCram.descriptor, activeCram.challenge);
  }
  renderCramPanel();
}

function markCramTypo() {
  const activeCram = resolveCramReview();
  if (!activeCram) {
    return;
  }

  const check = getCurrentCramCheck(activeCram.descriptor.id);
  if (!check || check.correct) {
    return;
  }

  state.cramLastCheck = {
    ...check,
    correct: true,
    autoGrade: "good",
    applyToSrs: false
  };
  renderCramPanel();
  setStatus("Typo forgiven for this cram prompt. SRS was unchanged either way.", "success");
}

function adoptCramAnswer() {
  const activeCram = resolveCramReview();
  if (!activeCram) {
    return;
  }

  const check = getCurrentCramCheck(activeCram.descriptor.id);
  if (!check || check.correct) {
    return;
  }

  const addedAnswers = adoptAnswerToChallenge(activeCram.descriptor, activeCram.challenge, dom.cramInput.value);
  if (!addedAnswers.length) {
    setStatus("That answer is already accepted for this card.", "info");
    return;
  }

  touchAutosave();
  state.cramLastCheck = {
    ...check,
    correct: true,
    autoGrade: "good",
    applyToSrs: true
  };
  renderLibrary();
  renderCramPanel();
  setStatus("Added that answer to the card. This cram result will count as a real correct review when you move on.", "success");
}

function finalizeCramReview() {
  const activeCram = resolveCramReview();
  if (!activeCram) {
    return;
  }

  const check = getCurrentCramCheck(activeCram.descriptor.id);
  const label = activeCram.descriptor.label.toLowerCase();
  if (check && check.correct && check.applyToSrs) {
    state.deck.progress.cards[activeCram.descriptor.id] = applyReview(
      state.deck.progress.cards[activeCram.descriptor.id],
      check.autoGrade || "good",
      true
    );
    touchAutosave();
  }

  state.cramCurrentCardId = null;
  state.cramCurrentChallenge = null;
  state.cramLastCheck = null;
  if (check && check.correct && check.applyToSrs) {
    render();
  } else {
    renderCramPanel();
  }

  if (state.cramQueue.length || state.cramCurrentCardId) {
    setStatus(
      check && check.correct && check.applyToSrs
        ? `Finished one cram ${label} prompt and counted it as a real review. ${state.cramQueue.length + (state.cramCurrentCardId ? 1 : 0)} prompt${state.cramQueue.length + (state.cramCurrentCardId ? 1 : 0) === 1 ? "" : "s"} left.`
        : `Finished one cram ${label} prompt. ${state.cramQueue.length + (state.cramCurrentCardId ? 1 : 0)} prompt${state.cramQueue.length + (state.cramCurrentCardId ? 1 : 0) === 1 ? "" : "s"} left.`,
      "info"
    );
  } else {
    setStatus(
      check && check.correct && check.applyToSrs
        ? "Cram session finished. Your taught answer was saved and counted as a real review."
        : "Cram session finished. Your SRS levels were left unchanged.",
      "success"
    );
  }
}

function resolveActiveReview() {
  if (state.currentCardId && state.currentChallenge && state.descriptorMap.has(state.currentCardId)) {
    return {
      descriptor: state.descriptorMap.get(state.currentCardId),
      challenge: state.currentChallenge
    };
  }

  const nextId = getDueCardIds()[0] || null;
  if (!nextId) {
    clearActiveReview();
    return null;
  }

  const descriptor = state.descriptorMap.get(nextId);
  state.currentCardId = nextId;
  state.currentChallenge = createChallenge(descriptor);
  state.lastCheck = null;

  return {
    descriptor,
    challenge: state.currentChallenge
  };
}

function clearActiveReview() {
  state.currentCardId = null;
  state.currentChallenge = null;
  state.lastCheck = null;
}

function clearReviewOrder() {
  state.reviewOrderIds = [];
}

function getCurrentCheck(cardId) {
  if (!state.lastCheck || state.lastCheck.cardId !== cardId) {
    return null;
  }
  return state.lastCheck;
}

function getSortedDueCardIds() {
  return state.descriptors
    .map((descriptor) => state.deck.progress.cards[descriptor.id])
    .filter((card) => card && isDue(card.due))
    .sort(compareReviewCards)
    .map((card) => card.id);
}

function getDueCardIds() {
  const sortedDueIds = getSortedDueCardIds();
  if (!state.reviewOrderIds.length) {
    return sortedDueIds;
  }

  const dueSet = new Set(sortedDueIds);
  const orderedIds = state.reviewOrderIds.filter((cardId) => dueSet.has(cardId));
  const orderedSet = new Set(orderedIds);
  const remainingIds = sortedDueIds.filter((cardId) => !orderedSet.has(cardId));
  return [...orderedIds, ...remainingIds];
}

function getCardGroups(cardIds) {
  const groups = [];
  const groupMap = new Map();

  cardIds.forEach((cardId) => {
    const descriptor = state.descriptorMap.get(cardId);
    if (!descriptor) {
      return;
    }

    if (!groupMap.has(descriptor.kanjiId)) {
      const group = {
        kanjiId: descriptor.kanjiId,
        cardIds: []
      };
      groupMap.set(descriptor.kanjiId, group);
      groups.push(group);
    }

    groupMap.get(descriptor.kanjiId).cardIds.push(cardId);
  });

  return groups;
}

function withActiveCardFirst(group, activeId) {
  if (!activeId || !group.cardIds.includes(activeId)) {
    return group;
  }

  return {
    ...group,
    cardIds: [activeId, ...group.cardIds.filter((cardId) => cardId !== activeId)]
  };
}

function compareReviewCards(left, right) {
  const dueDelta = new Date(left.due).getTime() - new Date(right.due).getTime();
  if (dueDelta !== 0) {
    return dueDelta;
  }

  if (left.masteryLevel !== right.masteryLevel) {
    return left.masteryLevel - right.masteryLevel;
  }

  return left.stability - right.stability;
}

function getCardsForKanji(kanjiId) {
  return state.descriptors
    .filter((descriptor) => descriptor.kanjiId === kanjiId)
    .map((descriptor) => state.deck.progress.cards[descriptor.id]);
}

function startReviewSession() {
  const nextId = getDueCardIds()[0] || null;

  if (!nextId) {
    clearActiveReview();
    renderReviewPanel();
    setStatus("Nothing is due right now. New cards are due immediately, and reviewed cards return when their schedules come back around.", "info");
    return;
  }

  const descriptor = state.descriptorMap.get(nextId);
  state.currentCardId = nextId;
  state.currentChallenge = createChallenge(descriptor);
  state.lastCheck = null;
  renderReviewPanel();
  setStatus("Review session started.", "info");
}

function handleAnswerSubmit(event) {
  event.preventDefault();
  const activeReview = resolveActiveReview();

  if (!activeReview) {
    setStatus("No due cards are available yet. Add more kanji or wait for scheduled reviews.", "info");
    return;
  }

  if (getCurrentCheck(activeReview.descriptor.id)) {
    finalizeCurrentReview();
    return;
  }

  const correct = isAnswerCorrect(activeReview.challenge, dom.answerInput.value);
  state.lastCheck = {
    cardId: activeReview.descriptor.id,
    correct,
    autoGrade: correct ? "good" : "hard"
  };
  if (!correct) {
    rememberWrongPrompt("review", activeReview.descriptor, activeReview.challenge);
  }
  renderReviewPanel();
}

function markReviewTypo() {
  const activeReview = resolveActiveReview();
  if (!activeReview) {
    return;
  }

  const check = getCurrentCheck(activeReview.descriptor.id);
  if (!check || check.correct) {
    return;
  }

  state.lastCheck = {
    ...check,
    correct: true,
    autoGrade: "good"
  };
  renderReviewPanel();
  setStatus("Typo forgiven. This review will count as correct when you move to the next card.", "success");
}

function adoptReviewAnswer() {
  const activeReview = resolveActiveReview();
  if (!activeReview) {
    return;
  }

  const check = getCurrentCheck(activeReview.descriptor.id);
  if (!check || check.correct) {
    return;
  }

  const addedAnswers = adoptAnswerToChallenge(activeReview.descriptor, activeReview.challenge, dom.answerInput.value);
  if (!addedAnswers.length) {
    setStatus("That answer is already accepted for this card.", "info");
    return;
  }

  touchAutosave();
  state.lastCheck = {
    ...check,
    correct: true,
    autoGrade: "good"
  };
  renderLibrary();
  renderReviewPanel();
  setStatus("Added that answer to the card. This review will count as correct when you move to the next card.", "success");
}

function markReviewEasy() {
  const activeReview = resolveActiveReview();
  if (!activeReview) {
    return;
  }

  const check = getCurrentCheck(activeReview.descriptor.id);
  if (!check || !check.correct || check.autoGrade === "easy") {
    return;
  }

  state.lastCheck = {
    ...check,
    autoGrade: "easy"
  };
  renderReviewPanel();
  setStatus("Marked as easy. This review will count as easy and push the next interval farther out.", "success");
}

function finalizeCurrentReview() {
  const activeReview = resolveActiveReview();
  if (!activeReview) {
    return;
  }

  const check = getCurrentCheck(activeReview.descriptor.id);
  if (!check) {
    return;
  }

  const updatedCard = applyReview(state.deck.progress.cards[activeReview.descriptor.id], check.autoGrade, check.correct);
  state.deck.progress.cards[activeReview.descriptor.id] = updatedCard;

  const label = activeReview.descriptor.label.toLowerCase();
  clearActiveReview();
  markDirty(`Reviewed ${label}. Next in ${formatInterval(updatedCard.scheduledDays)}.`);
  render();
}

function isAnswerCorrect(challenge, rawAnswer) {
  const submitted = tokenizeAnswers(rawAnswer, challenge.answerMode);
  const accepted = challenge.answers.map((answer) => normalizeAnswer(answer, challenge.answerMode));
  return submitted.some((value) => accepted.includes(value));
}

function tokenizeAnswers(value, mode) {
  return safeString(value)
    .split(/[\n,;\/|、；]+/)
    .map((chunk) => normalizeAnswer(chunk, mode))
    .filter(Boolean);
}

function normalizeAnswer(value, mode) {
  return mode === "reading" ? normalizeReading(value) : normalizeMeaning(value);
}

function normalizeMeaning(value) {
  return safeString(value)
    .toLowerCase()
    .replace(/\b(a|an|the)\b/g, " ")
    .replace(/[^\p{L}\p{N}\s-]/gu, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeReading(value) {
  return toHiragana(safeString(value))
    .toLowerCase()
    .replace(/[・･\-\s]/g, "")
    .trim();
}

function toHiragana(value) {
  return value.replace(/[\u30a1-\u30f6]/g, (char) => String.fromCharCode(char.charCodeAt(0) - 0x60));
}

function toKatakana(value) {
  return value.replace(/[\u3041-\u3096]/g, (char) => String.fromCharCode(char.charCodeAt(0) + 0x60));
}

function extractRawAnswerTokens(value) {
  return safeString(value)
    .split(/[\n,;\/|ã€ï¼›]+/)
    .map((chunk) => safeString(chunk))
    .filter(Boolean);
}

function adoptAnswerToChallenge(descriptor, challenge, rawAnswer) {
  if (!descriptor || !challenge) {
    return [];
  }

  if (challenge.answerMode === "meaning") {
    const addedMeanings = [];
    extractRawAnswerTokens(rawAnswer).forEach((token) => {
      if (pushUniqueAcceptedAnswer(descriptor.meanings, token, "meaning")) {
        addedMeanings.push(token);
      }
    });
    refreshChallengeAcceptedAnswers(descriptor, challenge);
    return addedMeanings;
  }

  const storedReadings = extractRawAnswerTokens(rawAnswer)
    .map((token) => formatStoredReadingToken(token, descriptor, challenge))
    .filter(Boolean);

  if (challenge.wordId) {
    const word = (Array.isArray(descriptor.words) ? descriptor.words : []).find((item) => item.id === challenge.wordId);
    if (!word) {
      return [];
    }

    const readings = parseWordReadings(word.reading);
    const addedReadings = [];
    storedReadings.forEach((reading) => {
      if (pushUniqueAcceptedAnswer(readings, reading, "reading")) {
        addedReadings.push(reading);
      }
    });
    if (addedReadings.length) {
      word.reading = readings.join(", ");
    }
    refreshChallengeAcceptedAnswers(descriptor, challenge);
    return addedReadings;
  }

  const targetList = Array.isArray(descriptor.readings) ? descriptor.readings : [];
  const addedReadings = [];
  storedReadings.forEach((reading) => {
    if (pushUniqueAcceptedAnswer(targetList, reading, "reading")) {
      addedReadings.push(reading);
    }
  });
  refreshChallengeAcceptedAnswers(descriptor, challenge);
  return addedReadings;
}

function pushUniqueAcceptedAnswer(list, value, mode) {
  const normalizedValue = normalizeAnswer(value, mode);
  if (!normalizedValue) {
    return false;
  }

  if (list.some((item) => normalizeAnswer(item, mode) === normalizedValue)) {
    return false;
  }

  list.push(value);
  return true;
}

function formatStoredReadingToken(value, descriptor, challenge) {
  const normalized = normalizeReading(value);
  if (!normalized) {
    return "";
  }

  if (challenge.wordId) {
    return normalized;
  }

  return descriptor.kind === "onyomi" ? toKatakana(normalized) : normalized;
}

function refreshChallengeAcceptedAnswers(descriptor, challenge) {
  if (challenge.answerMode === "meaning") {
    challenge.answers = descriptor.meanings;
    challenge.reveal = `Accepted meanings: ${descriptor.meanings.join(" / ")}`;
    return;
  }

  if (challenge.wordId) {
    const word = (Array.isArray(descriptor.words) ? descriptor.words : []).find((item) => item.id === challenge.wordId);
    if (!word) {
      return;
    }
    challenge.answers = parseWordReadings(word.reading);
    challenge.reveal = buildWordReveal(word);
    return;
  }

  const readingLabel = descriptor.kind === "onyomi" ? "onyomi" : "kunyomi";
  challenge.answers = descriptor.readings || [];
  challenge.reveal = `Accepted ${readingLabel} readings: ${(descriptor.readings || []).join(" / ")}`;
}

function previewReview(card, grade, wasCorrect) {
  return applyReview(card, grade, wasCorrect, true);
}

function applyReview(card, grade, wasCorrect, previewOnly = false) {
  const score = GRADE_CONFIG[grade].score;
  const now = new Date();
  const lastReview = card.lastReviewAt ? new Date(card.lastReviewAt) : null;
  const elapsedDays = lastReview ? Math.max(0, (now.getTime() - lastReview.getTime()) / 86400000) : 0;
  const retrievability = card.stability > 0 ? forgettingCurve(elapsedDays, card.stability) : 0;

  let stability = card.stability;
  let difficulty = card.difficulty;
  let stateName = card.state;
  let lapses = card.lapses;
  let consecutiveFailures = clampNumber(card.consecutiveFailures, 0, 100000, 0);
  let successfulReviews = card.successfulReviews;

  if (card.reps === 0) {
    if (wasCorrect) {
      stability = initialStability(score);
      difficulty = initialDifficulty(score);
      stateName = stability >= 1 ? "review" : "learning";
      consecutiveFailures = 0;
      successfulReviews += 1;
    } else {
      stability = initialStability(1);
      difficulty = initialDifficulty(score);
      stateName = "learning";
      lapses += 1;
      consecutiveFailures += 1;
    }
  } else if (!wasCorrect) {
    stability = updateStabilityOnFailure(card.stability, card.difficulty);
    difficulty = updateDifficulty(card.difficulty, score);
    stateName = "relearning";
    lapses += 1;
    consecutiveFailures += 1;
    successfulReviews = Math.max(0, successfulReviews - 1);
  } else {
    stability = updateStabilityOnSuccess(card.stability, card.difficulty, retrievability, score);
    difficulty = updateDifficulty(card.difficulty, score);
    stateName = stability >= 1 ? "review" : "learning";
    consecutiveFailures = 0;
    successfulReviews += 1;
  }

  const scheduledDays = intervalFromStability(stability);
  const masteryLevel = inferMasteryLevel({
    reps: card.reps + 1,
    successfulReviews,
    scheduledDays
  });
  const due = new Date(now.getTime() + scheduledDays * 86400000).toISOString();

  const nextCard = {
    ...card,
    state: stateName,
    due,
    lastReviewAt: now.toISOString(),
    stability,
    difficulty,
    scheduledDays,
    reps: card.reps + 1,
    lapses,
    consecutiveFailures,
    successfulReviews,
    masteryLevel
  };

  if (previewOnly) {
    return nextCard;
  }

  return {
    ...nextCard,
    history: [
      ...(Array.isArray(card.history) ? card.history.slice(-59) : []),
      {
        at: now.toISOString(),
        grade,
        correct: wasCorrect,
        scheduledDays,
        elapsedDays,
        stability,
        difficulty
      }
    ]
  };
}

function forgettingCurve(elapsedDays, stability) {
  if (!stability) {
    return 0;
  }
  return Math.pow(1 + (FSRS_FACTOR * elapsedDays) / stability, FSRS_DECAY);
}

function initialStability(score) {
  if (score === 1) {
    return 0.02;
  }
  if (score === 2) {
    return 0.2;
  }
  if (score === 3) {
    return 1;
  }
  return 3;
}

function initialDifficulty(score) {
  return clampNumber(7.6 - (score - 1) * 1.15, 1, 10, 5.5);
}

function updateDifficulty(currentDifficulty, score) {
  const deltaByScore = { 1: 0.9, 2: 0.25, 3: -0.1, 4: -0.35 };
  const shifted = currentDifficulty + deltaByScore[score];
  const reverted = shifted * 0.78 + initialDifficulty(score) * 0.22;
  return clampNumber(reverted, 1, 10, 5.5);
}

function updateStabilityOnSuccess(currentStability, difficulty, retrievability, score) {
  const difficultyBonus = 1 + (11 - difficulty) * 0.08;
  const retrievalBonus = 1 + (1 - retrievability) * 1.4;
  const scaleBonus = 1 + Math.log10(currentStability + 1) * 0.24;
  const gradeBonus = score === 4 ? 1.18 : score === 2 ? 0.82 : 1;
  const nextStability = currentStability * difficultyBonus * retrievalBonus * scaleBonus * gradeBonus;
  return Math.max(currentStability + 0.15, nextStability);
}

function updateStabilityOnFailure(currentStability, difficulty) {
  const relearn = 0.03 * Math.pow(currentStability + 1, 0.3) / (0.45 + difficulty * 0.08);
  return clampNumber(relearn, 0.01, 0.2, 0.03);
}

function intervalFromStability(stability) {
  const interval = (stability / FSRS_FACTOR) * (Math.pow(REVIEW_RETENTION, 1 / FSRS_DECAY) - 1);
  return clampNumber(interval, 0.01, 3650, 0.01);
}

function inferMasteryLevel(card) {
  const reps = clampNumber(card && card.reps, 0, 100000, 0);
  const successes = clampNumber(card && card.successfulReviews, 0, 100000, 0);
  const interval = clampNumber(card && card.scheduledDays, 0, 3650, 0);

  if (reps === 0) {
    return 0;
  }
  if (successes >= 8 || interval >= 60) {
    return 5;
  }
  if (successes >= 6 || interval >= 21) {
    return 4;
  }
  if (successes >= 4 || interval >= 7) {
    return 3;
  }
  if (successes >= 2 || interval >= 2) {
    return 2;
  }
  return 1;
}

function formatMasteryLabel(level) {
  const key = typeof level === "number" ? MASTERY_LEVELS[level] : level;
  return MASTERY_LABELS[key] || key || "";
}

function formatMasteryEmoji(level) {
  const key = typeof level === "number" ? MASTERY_LEVELS[level] : level;
  return MASTERY_EMOJIS[key] || "";
}

function getMasteryIconMarkup(level, iconOnly = false) {
  const key = typeof level === "number" ? MASTERY_LEVELS[level] : level;
  if (!key) {
    return "";
  }
  if (key === "cockatrice") {
    return `<img class="mastery-inline-icon${iconOnly ? " mastery-inline-icon-only" : ""}" src="${COCKATRICE_ICON_PATH}" alt="" aria-hidden="true">`;
  }
  const emoji = formatMasteryEmoji(key);
  if (!emoji) {
    return "";
  }
  return `<span class="mastery-emoji${iconOnly ? " mastery-inline-icon-only" : ""}" aria-hidden="true">${escapeHtml(emoji)}</span>`;
}

function formatMasteryDisplayMarkup(level) {
  const label = formatMasteryLabel(level);
  const icon = getMasteryIconMarkup(level);
  return icon ? `${escapeHtml(label)} ${icon}` : escapeHtml(label);
}

function formatDue(dateString) {
  const due = new Date(dateString);
  if (Number.isNaN(due.getTime()) || isDue(dateString)) {
    return "Due now";
  }

  return `Due ${due.toLocaleString([], { month: "short", day: "numeric", hour: "numeric", minute: "2-digit" })}`;
}

function getDueBucket(dateString) {
  const due = new Date(dateString);
  if (Number.isNaN(due.getTime()) || due.getTime() <= Date.now()) {
    return { key: "due-now", label: "due now", sort: 0 };
  }

  const today = startOfLocalDay(new Date());
  const dueDay = startOfLocalDay(due);
  const dayDelta = Math.round((dueDay.getTime() - today.getTime()) / 86400000);

  if (dayDelta <= 0) {
    return { key: "later-today", label: "later today", sort: 1 };
  }

  if (dayDelta === 1) {
    return { key: "tomorrow", label: "tomorrow", sort: 2 };
  }

  if (dayDelta <= 30) {
    return { key: `${dayDelta}-days`, label: `${dayDelta} days`, sort: dayDelta + 1 };
  }

  return {
    key: due.toISOString().slice(0, 10),
    label: due.toLocaleDateString([], { month: "short", day: "numeric", year: "numeric" }).toLowerCase(),
    sort: dayDelta + 1
  };
}

function startOfLocalDay(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function formatInterval(days) {
  if (days < 1 / 24) {
    return `${Math.max(1, Math.round(days * 24 * 60))}m`;
  }
  if (days < 1) {
    return `${Math.max(1, Math.round(days * 24))}h`;
  }
  if (days < 30) {
    return `${roundToOne(days)}d`;
  }
  if (days < 365) {
    return `${roundToOne(days / 30)}mo`;
  }
  return `${roundToOne(days / 365)}y`;
}

function roundToOne(value) {
  return Math.round(value * 10) / 10;
}

function isDue(dateString) {
  const due = new Date(dateString);
  return !Number.isNaN(due.getTime()) && due.getTime() <= Date.now();
}

function handleKanjiSave(event) {
  event.preventDefault();
  const kanji = collectKanjiFromForm();

  if (!kanji.char) {
    setStatus("Add a kanji character before saving.", "error");
    dom.kanjiChar.focus();
    return;
  }

  if (!kanji.meanings.length && !kanji.onyomi.length && !kanji.kunyomi.length && !countExampleWords(kanji)) {
    setStatus("Add at least one meaning, reading, or example word so the app has something to quiz.", "error");
    return;
  }

  const existingIndex = state.deck.kanji.findIndex((entry) => entry.id === kanji.id);
  if (existingIndex >= 0) {
    state.deck.kanji[existingIndex] = kanji;
  } else {
    state.deck.kanji.push(kanji);
  }

  refreshDerivedState();
  resetEditor();
  markDirty(`Saved ${kanji.char}. ${getCardsForKanji(kanji.id).length} review card${getCardsForKanji(kanji.id).length === 1 ? "" : "s"} are ready.`);
  render();
}

function collectKanjiFromForm() {
  return normalizeKanji({
    id: safeString(dom.kanjiId.value) || uid("kanji"),
    char: safeString(dom.kanjiChar.value),
    meanings: parseListFromField(dom.kanjiMeanings.value),
    onyomi: parseListFromField(dom.kanjiOnyomi.value),
    kunyomi: parseListFromField(dom.kanjiKunyomi.value),
    jlpt: safeString(dom.kanjiJlpt.value),
    grade: safeString(dom.kanjiGrade.value),
    mnemonic: safeString(dom.kanjiMnemonic.value),
    words: {
      on: collectWordsFromList(dom.wordLists.on),
      kun: collectWordsFromList(dom.wordLists.kun),
      irregular: collectWordsFromList(dom.wordLists.irregular)
    }
  });
}

function collectWordsFromList(container) {
  return Array.from(container.querySelectorAll(".word-row"))
    .map((row) => ({
      id: row.dataset.wordId || uid("word"),
      word: safeString(row.querySelector("[data-field='word']").value),
      reading: safeString(row.querySelector("[data-field='reading']").value),
      definition: safeString(row.querySelector("[data-field='definition']").value)
    }))
    .filter((word) => word.word || word.reading || word.definition);
}

function resetEditor() {
  dom.kanjiForm.reset();
  dom.kanjiId.value = "";
  state.lastAuthoringLookupChar = "";
  resetWordDraft();
  WORD_CATEGORIES.forEach((category) => {
    dom.wordLists[category].innerHTML = "";
    appendWordRow(category, blankWord());
  });
  renderAuthoringWordSummary();
  setLookupStatus(dom.kanjiLookupStatus, "", "hidden");
  clearWordLookupResults("", "hidden");
}

function blankWord() {
  return {
    id: uid("word"),
    word: "",
    reading: "",
    definition: ""
  };
}

function appendWordRow(category, word, options = {}) {
  const row = createWordRow(word);
  dom.wordLists[category].appendChild(row);
  renderAuthoringWordSummary();

  if (options.focus) {
    requestAnimationFrame(() => {
      const input = row.querySelector("[data-field='word']");
      if (!input) {
        return;
      }

      input.focus();
      input.scrollIntoView({ behavior: "smooth", block: "nearest" });
    });
  }

  return row;
}

function createWordRow(word) {
  const row = document.createElement("div");
  row.className = "word-row";
  row.dataset.wordId = word.id || uid("word");
  row.innerHTML = `
    <input class="text-input" data-field="word" type="text" placeholder="Word" value="${escapeHtml(word.word || "")}">
    <input class="text-input" data-field="reading" type="text" placeholder="Reading, alternate reading" value="${escapeHtml(word.reading || "")}">
    <input class="text-input" data-field="definition" type="text" placeholder="Definition" value="${escapeHtml(word.definition || "")}">
    <button class="button button-small" type="button" data-remove-word>Remove</button>
  `;
  return row;
}

function loadKanjiIntoEditor(kanjiId) {
  const kanji = state.deck.kanji.find((entry) => entry.id === kanjiId);
  if (!kanji) {
    return;
  }

  dom.kanjiId.value = kanji.id;
  dom.kanjiChar.value = kanji.char;
  dom.kanjiMeanings.value = kanji.meanings.join("\n");
  dom.kanjiOnyomi.value = kanji.onyomi.join("\n");
  dom.kanjiKunyomi.value = kanji.kunyomi.join("\n");
  dom.kanjiJlpt.value = kanji.jlpt || "";
  dom.kanjiGrade.value = kanji.grade || "";
  dom.kanjiMnemonic.value = kanji.mnemonic || "";
  state.lastAuthoringLookupChar = kanji.char;

  WORD_CATEGORIES.forEach((category) => {
    dom.wordLists[category].innerHTML = "";
    const words = kanji.words[category].length ? kanji.words[category] : [blankWord()];
    words.forEach((word) => appendWordRow(category, word));
  });

  renderAuthoringWordSummary();
  resetWordDraft();
  clearWordLookupResults("", "hidden");
  dom.kanjiChar.focus();
  setStatus(`Editing ${kanji.char}.`, "info");
}

function deleteKanji(kanjiId) {
  const kanji = state.deck.kanji.find((entry) => entry.id === kanjiId);
  if (!kanji) {
    return;
  }

  if (!window.confirm(`Delete ${kanji.char} and its review progress?`)) {
    return;
  }

  state.deck.kanji = state.deck.kanji.filter((entry) => entry.id !== kanjiId);
  refreshDerivedState();

  if (dom.kanjiId.value === kanjiId) {
    resetEditor();
  }

  markDirty(`Deleted ${kanji.char}.`);
  render();
}

function countExampleWords(kanji) {
  return WORD_CATEGORIES.reduce((sum, category) => sum + kanji.words[category].length, 0);
}

function matchesSearch(kanji, query) {
  if (!query || !query.textTerms.length && !query.kanjiChars.length) {
    return true;
  }

  if (query.isBareKanjiBatch) {
    return query.kanjiChars.includes(kanji.char);
  }

  const haystack = [
    kanji.char,
    kanji.jlpt,
    kanji.grade,
    kanji.mnemonic,
    ...kanji.meanings,
    ...kanji.onyomi,
    ...kanji.kunyomi,
    ...WORD_CATEGORIES.flatMap((category) => kanji.words[category].flatMap((word) => [word.word, word.reading, word.definition]))
  ]
    .map((value) => buildSearchText(value))
    .join(" ");

  return query.textTerms.every((term) => haystack.includes(term));
}

function parseLibrarySearch(value) {
  const raw = safeString(value);
  if (!raw) {
    return {
      textTerms: [],
      kanjiChars: [],
      isBareKanjiBatch: false
    };
  }

  const textTerms = raw
    .split(SEARCH_SPLIT_PATTERN)
    .map((part) => buildSearchText(part))
    .filter(Boolean);
  const condensed = raw.replace(SEARCH_SPLIT_PATTERN, "");
  const isBareKanjiBatch = condensed.length > 1 && Array.from(condensed).every(isKanjiSearchChar);

  return {
    textTerms,
    kanjiChars: isBareKanjiBatch ? Array.from(new Set(Array.from(condensed))) : [],
    isBareKanjiBatch
  };
}

function isKanjiSearchChar(char) {
  return KANJI_SEARCH_CHAR_PATTERN.test(char);
}

function buildSearchText(value) {
  const text = safeString(value);
  return buildSearchForms(text)
    .filter(Boolean)
    .join(" ");
}

function buildSearchForms(value) {
  const text = safeString(value);
  const forms = uniqueList([
    text.toLowerCase(),
    normalizeMeaning(text),
    normalizeReading(text)
  ].filter(Boolean));
  return forms;
}

function isFileSystemAccessSupported() {
  return typeof window.showOpenFilePicker === "function"
    && typeof window.showSaveFilePicker === "function";
}

function getDeckFilePickerOptions() {
  return {
    types: [
      {
        description: "Kanji Den JSON deck",
        accept: {
          "application/json": [".json"]
        }
      }
    ],
    excludeAcceptAllOption: false
  };
}

async function openDeckFile() {
  if (!isFileSystemAccessSupported()) {
    setStatus("Direct file save is not supported in this browser. Try desktop Chrome or Edge, or use Load JSON and Export JSON.", "error");
    return;
  }

  try {
    const [handle] = await window.showOpenFilePicker({
      ...getDeckFilePickerOptions(),
      multiple: false
    });
    const file = await handle.getFile();
    const text = await file.text();
    const parsedDeck = parseImportedDeckPayload(JSON.parse(text), file.name);
    if (!parsedDeck) {
      throw new Error("Imported JSON did not contain a recognizable deck payload.");
    }

    deckLoadToken += 1;
    applyLoadedDeck(parsedDeck.deck, {
      sourceLabel: file.name,
      unsavedChanges: false,
      deckFileHandle: handle
    });
    persistAutosave();
    render();
    setStatus(`Opened ${file.name}. Save Deck File can now write back to this JSON file.`, "success");
  } catch (error) {
    if (error && error.name === "AbortError") {
      return;
    }
    console.error("Deck open failed:", error);
    setStatus("That deck file could not be opened. Check that it is a valid Kanji Den JSON deck.", "error");
  }
}

async function saveDeckFile() {
  if (!isFileSystemAccessSupported()) {
    setStatus("Direct file save is not supported in this browser. Use Export JSON instead.", "error");
    return;
  }

  if (!state.deckFileHandle) {
    await saveDeckFileAs();
    return;
  }

  try {
    const canWrite = await requestFileWritePermission(state.deckFileHandle);
    if (!canWrite) {
      setStatus("The browser did not grant permission to write that deck file.", "error");
      return;
    }

    await writeDeckToFileHandle(state.deckFileHandle);
    state.sourceLabel = state.deckFileHandle.name || state.sourceLabel;
    state.unsavedChanges = false;
    persistAutosave();
    render();
    setStatus(`Saved ${state.sourceLabel}.`, "success");
  } catch (error) {
    if (error && error.name === "AbortError") {
      return;
    }
    setStatus("The deck file could not be saved. Try Save Deck As or Export JSON.", "error");
  }
}

async function saveDeckFileAs() {
  if (!isFileSystemAccessSupported()) {
    setStatus("Direct file save is not supported in this browser. Use Export JSON instead.", "error");
    return;
  }

  try {
    const suggestedName = state.sourceLabel && state.sourceLabel.toLowerCase().endsWith(".json")
      ? state.sourceLabel
      : `kanji-den-${new Date().toISOString().slice(0, 10)}.json`;
    const handle = await window.showSaveFilePicker({
      ...getDeckFilePickerOptions(),
      suggestedName
    });

    state.deckFileHandle = handle;
    await writeDeckToFileHandle(handle);
    state.sourceLabel = handle.name || suggestedName;
    state.unsavedChanges = false;
    persistAutosave();
    render();
    setStatus(`Saved ${state.sourceLabel}. Future Save Deck File clicks will write to this file.`, "success");
  } catch (error) {
    if (error && error.name === "AbortError") {
      return;
    }
    setStatus("The deck file could not be saved. Use Export JSON as a fallback.", "error");
  }
}

async function requestFileWritePermission(handle) {
  const options = { mode: "readwrite" };
  if (typeof handle.queryPermission === "function" && await handle.queryPermission(options) === "granted") {
    return true;
  }
  return typeof handle.requestPermission === "function" && await handle.requestPermission(options) === "granted";
}

async function writeDeckToFileHandle(handle) {
  refreshDerivedState();
  const payload = JSON.stringify(state.deck, null, 2);
  const writable = await handle.createWritable();
  await writable.write(payload);
  await writable.close();
}

function exportDeck() {
  refreshDerivedState();
  const payload = JSON.stringify(state.deck, null, 2);
  const fileName = `kanji-den-${new Date().toISOString().slice(0, 10)}.json`;
  downloadBrowserFile(fileName, payload, "application/json");

  state.sourceLabel = fileName;
  state.unsavedChanges = false;
  persistAutosave();
  render();
  setStatus(`Exported ${fileName}.`, "success");
}

function downloadBulkImportTemplate() {
  const payload = [
    "# The Kanji Den bulk import template",
    "# Use one KANJI block per entry.",
    "# Word lines use: word | reading | definition",
    "",
    "KANJI:",
    "MEANINGS:",
    "ON:",
    "KUN:",
    "JLPT:",
    "GRADE:",
    "MNEMONIC:",
    "ON WORDS:",
    "",
    "KUN WORDS:",
    "",
    "IRREGULAR WORDS:",
    ""
  ].join("\n");

  const fileName = "kanji-den-bulk-import-template.txt";
  downloadBrowserFile(fileName, payload, "text/plain;charset=utf-8");

  setStatus(`Downloaded ${fileName}.`, "success");
}

function exportAllKanjiData() {
  exportKanjiDataAsBulkTxt(state.deck.kanji, "all");
}

function exportSelectedKanjiData() {
  exportKanjiDataAsBulkTxt(getMultiSelectedKanjiEntries(), "selected");
}

function exportFilteredKanjiData() {
  exportKanjiDataAsBulkTxt(getFilteredKanji(), "filtered");
}

function exportKanjiDataAsBulkTxt(entries, scope) {
  const kanjiEntries = normalizeKanjiExportEntries(entries);
  if (!kanjiEntries.length) {
    const scopeMessages = {
      all: "There is no kanji data to export yet.",
      selected: "Select one or more kanji in the hoard first.",
      filtered: "No kanji match the current hoard filters."
    };
    setStatus(scopeMessages[scope] || "There is no kanji data to export yet.", "error");
    return;
  }

  const fileName = `kanji-den-${scope}-kanji-data-${new Date().toISOString().slice(0, 10)}.txt`;
  const payload = buildBulkKanjiExportPayload(kanjiEntries);
  downloadBrowserFile(fileName, payload, "text/plain;charset=utf-8");
  setStatus(`Exported ${kanjiEntries.length} kanji to ${fileName}. Only kanji and word data were included.`, "success");
}

function normalizeKanjiExportEntries(entries) {
  return (Array.isArray(entries) ? entries : [])
    .map((entry) => normalizeKanji(entry))
    .filter((entry) => entry.char);
}

function buildBulkKanjiExportPayload(entries) {
  return entries.map((entry) => formatBulkKanjiExportBlock(entry)).join("\n\n");
}

function formatBulkKanjiExportBlock(entry) {
  const lines = [
    `KANJI: ${safeString(entry.char)}`,
    `MEANINGS: ${formatBulkListField(entry.meanings)}`,
    `ON: ${formatBulkListField(entry.onyomi)}`,
    `KUN: ${formatBulkListField(entry.kunyomi)}`,
    `JLPT: ${safeString(entry.jlpt)}`,
    `GRADE: ${safeString(entry.grade)}`
  ];

  appendBulkTextField(lines, "MNEMONIC", entry.mnemonic);
  appendBulkWordSection(lines, "ON WORDS", entry.words && entry.words.on);
  appendBulkWordSection(lines, "KUN WORDS", entry.words && entry.words.kun);
  appendBulkWordSection(lines, "IRREGULAR WORDS", entry.words && entry.words.irregular);
  return lines.join("\n");
}

function formatBulkListField(values) {
  return normalizeList(values).join(", ");
}

function appendBulkTextField(lines, label, value) {
  const parts = safeString(value)
    .split(/\r?\n/)
    .map((part) => safeString(part))
    .filter(Boolean);

  if (!parts.length) {
    lines.push(`${label}:`);
    return;
  }

  parts.forEach((part) => {
    lines.push(`${label}: ${part}`);
  });
}

function appendBulkWordSection(lines, label, words) {
  lines.push(`${label}:`);
  normalizeWordList(words).forEach((word) => {
    lines.push(formatBulkWordLine(word));
  });
}

function formatBulkWordLine(word) {
  return [
    safeString(word && word.word),
    safeString(word && word.reading),
    safeString(word && word.definition)
  ].join(" | ");
}

function downloadBrowserFile(fileName, payload, mimeType) {
  const blob = new Blob([payload], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(url);
}

function handleImportFile(event) {
  const [file] = event.target.files || [];
  if (!file) {
    return;
  }

  const reader = new FileReader();
  reader.onload = () => {
    try {
      const parsedDeck = parseImportedDeckPayload(JSON.parse(String(reader.result)), file.name);
      if (!parsedDeck) {
        throw new Error("Imported JSON did not contain a recognizable deck payload.");
      }

      deckLoadToken += 1;
      applyLoadedDeck(parsedDeck.deck, {
        sourceLabel: file.name,
        unsavedChanges: false,
        deckFileHandle: null
      });
      persistAutosave();
      render();
      setStatus(`Loaded ${file.name}.`, "success");
    } catch (error) {
      console.error("Deck import failed:", error);
      setStatus("That JSON file could not be loaded. Check that it is a valid Kanji Dragon deck.", "error");
    }
  };
  reader.readAsText(file);
  event.target.value = "";
}

function handleTxtImportFile(event) {
  const [file] = event.target.files || [];
  if (!file) {
    return;
  }

  const reader = new FileReader();
  reader.onload = () => {
    try {
      state.bulkImportDraft = createBulkImportDraft(String(reader.result), file.name);
      renderBulkImportPreview();

      const validCount = state.bulkImportDraft.entries.length;
      if (!validCount) {
        setStatus("No valid kanji blocks were found in that TXT file. Check the preview notes in Deck Tools.", "error");
        return;
      }

      setStatus(`Preview ready for ${validCount} TXT kanji block${validCount === 1 ? "" : "s"}. Choose how to import it in Deck Tools.`, "info");
    } catch (error) {
      state.bulkImportDraft = null;
      renderBulkImportPreview();
      setStatus("That TXT file could not be parsed.", "error");
    }
  };
  reader.readAsText(file);
  event.target.value = "";
}

function renderBulkImportPreview() {
  const draft = state.bulkImportDraft;
  if (!draft) {
    dom.bulkImportPreview.classList.add("is-hidden");
    dom.bulkImportPreview.innerHTML = "";
    return;
  }

  const errorItems = draft.errors.slice(0, 6).map((error) => `<li>${escapeHtml(error)}</li>`).join("");
  const extraErrorCount = Math.max(0, draft.errors.length - 6);

  dom.bulkImportPreview.classList.remove("is-hidden");
  dom.bulkImportPreview.innerHTML = `
    <div class="bulk-import-head">
      <div>
        <p class="bulk-import-title">TXT Import Preview</p>
        <p class="bulk-import-file">${escapeHtml(draft.fileName)}</p>
      </div>
      <button class="button button-small" type="button" data-bulk-import="cancel">Cancel</button>
    </div>
    <div class="bulk-import-counts">
      <span><strong>${draft.entries.length}</strong> valid</span>
      <span><strong>${draft.newEntries.length}</strong> new</span>
      <span><strong>${draft.existingEntries.length}</strong> existing</span>
      <span><strong>${draft.errors.length}</strong> notes</span>
    </div>
    <p class="bulk-import-note">Existing kanji are merged, not replaced. SRS progress stays intact.</p>
    ${draft.errors.length ? `<ul class="bulk-import-errors">${errorItems}${extraErrorCount ? `<li>And ${extraErrorCount} more note${extraErrorCount === 1 ? "" : "s"}.</li>` : ""}</ul>` : ""}
    <div class="bulk-import-actions">
      <button class="button button-small" type="button" data-bulk-import="new" ${draft.newEntries.length ? "" : "disabled"}>Add New Only</button>
      <button class="button button-small" type="button" data-bulk-import="update" ${draft.existingEntries.length ? "" : "disabled"}>Update Existing</button>
      <button class="button button-small button-primary" type="button" data-bulk-import="all" ${draft.entries.length ? "" : "disabled"}>Add New + Update Existing</button>
    </div>
  `;
}

function handleBulkImportPreviewClick(event) {
  const button = event.target.closest("[data-bulk-import]");
  if (!button) {
    return;
  }

  const action = button.dataset.bulkImport;
  if (action === "cancel") {
    state.bulkImportDraft = null;
    renderBulkImportPreview();
    setStatus("TXT import canceled.", "info");
    return;
  }

  applyBulkImport(action);
}

function applyBulkImport(action) {
  const draft = state.bulkImportDraft;
  if (!draft) {
    return;
  }

  const includeNew = action === "new" || action === "all";
  const includeExisting = action === "update" || action === "all";
  const existingByChar = new Map(state.deck.kanji.map((entry) => [entry.char, entry]));
  let added = 0;
  let updated = 0;

  draft.entries.forEach((entry) => {
    const existing = existingByChar.get(entry.char);
    if (existing) {
      if (!includeExisting) {
        return;
      }

      const merged = mergeKanjiData(existing, entry);
      const index = state.deck.kanji.findIndex((kanji) => kanji.id === existing.id);
      if (index >= 0) {
        state.deck.kanji[index] = merged;
        existingByChar.set(merged.char, merged);
        updated += 1;
      }
      return;
    }

    if (!includeNew) {
      return;
    }

    const newEntry = normalizeKanji({
      ...entry,
      id: uid("kanji")
    });
    state.deck.kanji.push(newEntry);
    existingByChar.set(newEntry.char, newEntry);
    added += 1;
  });

  if (!added && !updated) {
    setStatus("TXT import did not change the deck with that option.", "info");
    return;
  }

  state.bulkImportDraft = null;
  clearActiveReview();
  clearReviewOrder();
  clearCramSession();
  refreshDerivedState();
  resetEditor();
  markDirty(`TXT import complete. Added ${added} new kanji and updated ${updated} existing kanji.`);
  render();
}

function createBulkImportDraft(text, fileName) {
  const parsed = parseTxtKanjiBlocks(text);
  const entries = mergeDuplicateImportEntries(parsed.entries, parsed.errors);
  const existingChars = new Set(state.deck.kanji.map((entry) => entry.char));
  return {
    fileName,
    entries,
    newEntries: entries.filter((entry) => !existingChars.has(entry.char)),
    existingEntries: entries.filter((entry) => existingChars.has(entry.char)),
    errors: parsed.errors
  };
}

function parseTxtKanjiBlocks(text) {
  const entries = [];
  const errors = [];
  const lines = String(text).replace(/^\uFEFF/, "").split(/\r?\n/);
  let current = null;
  let currentWordSection = null;

  const finishCurrent = (lineNumber) => {
    if (!current) {
      return;
    }

    const normalized = normalizeKanji(current);
    if (!normalized.char) {
      errors.push(`Line ${current.lineNumber}: Missing kanji after KANJI:.`);
    } else if (!normalized.meanings.length && !normalized.onyomi.length && !normalized.kunyomi.length && !countExampleWords(normalized)) {
      errors.push(`Line ${current.lineNumber}: ${normalized.char} was skipped because it has no meanings, readings, or words.`);
    } else {
      entries.push(normalized);
    }

    current = null;
    currentWordSection = null;
  };

  lines.forEach((rawLine, index) => {
    const lineNumber = index + 1;
    const line = rawLine.trim();
    if (!line || line.startsWith("#")) {
      return;
    }

    const header = line.match(/^([^:]+):\s*(.*)$/);
    if (header) {
      const label = normalizeTxtLabel(header[1]);
      const value = safeString(header[2]);
      const type = getTxtHeaderType(label);

      if (type && type.kind === "kanji") {
        finishCurrent(lineNumber - 1);
        current = {
          id: uid("kanji"),
          char: value,
          meanings: [],
          onyomi: [],
          kunyomi: [],
          jlpt: "",
          grade: "",
          mnemonic: "",
          words: { on: [], kun: [], irregular: [] },
          lineNumber
        };
        currentWordSection = null;
        return;
      }

      if (!current) {
        errors.push(`Line ${lineNumber}: ${label}: appears before any KANJI: block.`);
        return;
      }

      if (type && type.kind === "list") {
        current[type.field] = uniqueList([...current[type.field], ...parseListFromField(value)]);
        currentWordSection = null;
        return;
      }

      if (type && type.kind === "text") {
        current[type.field] = mergeTxtTextField(type.field, current[type.field], value);
        currentWordSection = null;
        return;
      }

      if (type && type.kind === "words") {
        currentWordSection = type.category;
        if (value) {
          parseTxtWordChunk(value, current, currentWordSection, lineNumber, errors);
        }
        return;
      }

      errors.push(`Line ${lineNumber}: Unknown TXT label "${label}".`);
      return;
    }

    if (current && currentWordSection) {
      parseTxtWordLine(line, current, currentWordSection, lineNumber, errors);
      return;
    }

    errors.push(`Line ${lineNumber}: Could not understand "${line}".`);
  });

  finishCurrent(lines.length);
  return { entries, errors };
}

function normalizeTxtLabel(label) {
  return safeString(label).toUpperCase().replace(/\s+/g, " ");
}

function getTxtHeaderType(label) {
  const labels = {
    KANJI: { kind: "kanji" },
    MEANING: { kind: "list", field: "meanings" },
    MEANINGS: { kind: "list", field: "meanings" },
    JLPT: { kind: "text", field: "jlpt" },
    "JLPT LEVEL": { kind: "text", field: "jlpt" },
    GRADE: { kind: "text", field: "grade" },
    "GRADE LEVEL": { kind: "text", field: "grade" },
    MNEMONIC: { kind: "text", field: "mnemonic" },
    MNEMONICS: { kind: "text", field: "mnemonic" },
    ON: { kind: "list", field: "onyomi" },
    ONYOMI: { kind: "list", field: "onyomi" },
    "ON READINGS": { kind: "list", field: "onyomi" },
    "ONYOMI READINGS": { kind: "list", field: "onyomi" },
    KUN: { kind: "list", field: "kunyomi" },
    KUNYOMI: { kind: "list", field: "kunyomi" },
    "KUN READINGS": { kind: "list", field: "kunyomi" },
    "KUNYOMI READINGS": { kind: "list", field: "kunyomi" },
    "ON WORDS": { kind: "words", category: "on" },
    "ONYOMI WORDS": { kind: "words", category: "on" },
    "KUN WORDS": { kind: "words", category: "kun" },
    "KUNYOMI WORDS": { kind: "words", category: "kun" },
    "IRREGULAR WORDS": { kind: "words", category: "irregular" },
    "IRREGULAR READING WORDS": { kind: "words", category: "irregular" }
  };
  return labels[label] || null;
}

function parseTxtWordChunk(value, current, category, lineNumber, errors) {
  value.split(";").map((entry) => safeString(entry)).filter(Boolean).forEach((entry) => {
    parseTxtWordLine(entry, current, category, lineNumber, errors);
  });
}

function parseTxtWordLine(line, current, category, lineNumber, errors) {
  const delimiter = line.includes("|") ? "|" : line.includes("=") ? "=" : null;
  if (!delimiter) {
    errors.push(`Line ${lineNumber}: Word lines should use "word | reading | definition".`);
    return;
  }

  const parts = line.split(delimiter).map((part) => safeString(part));
  const word = parts[0];
  const reading = parts[1];
  const definition = parts.slice(2).join(` ${delimiter} `).trim();

  if (!word || !reading) {
    errors.push(`Line ${lineNumber}: Word lines need at least a word and reading.`);
    return;
  }

  current.words[category].push({
    id: uid("word"),
    word,
    reading,
    definition
  });
}

function mergeDuplicateImportEntries(entries, errors) {
  const merged = [];
  const byChar = new Map();

  entries.forEach((entry) => {
    if (!byChar.has(entry.char)) {
      byChar.set(entry.char, entry);
      merged.push(entry);
      return;
    }

    const combined = mergeKanjiData(byChar.get(entry.char), entry);
    byChar.set(entry.char, combined);
    const index = merged.findIndex((item) => item.char === entry.char);
    if (index >= 0) {
      merged[index] = combined;
    }
    errors.push(`Merged duplicate TXT blocks for ${entry.char}.`);
  });

  return merged;
}

function mergeKanjiData(existing, incoming) {
  return normalizeKanji({
    id: existing.id,
    char: existing.char || incoming.char,
    meanings: uniqueList([...existing.meanings, ...incoming.meanings]),
    onyomi: uniqueList([...existing.onyomi, ...incoming.onyomi]),
    kunyomi: uniqueList([...existing.kunyomi, ...incoming.kunyomi]),
    jlpt: chooseUpdatedScalar(existing.jlpt, incoming.jlpt),
    grade: chooseUpdatedScalar(existing.grade, incoming.grade),
    mnemonic: mergeMnemonicText(existing.mnemonic, incoming.mnemonic),
    words: {
      on: mergeWordLists(existing.words.on, incoming.words.on),
      kun: mergeWordLists(existing.words.kun, incoming.words.kun),
      irregular: mergeWordLists(existing.words.irregular, incoming.words.irregular)
    }
  });
}

function mergeMnemonicText(existingText, incomingText) {
  const parts = [safeString(existingText), safeString(incomingText)].filter(Boolean);
  return uniqueList(parts).join("\n");
}

function mergeTxtTextField(fieldName, existingValue, incomingValue) {
  if (fieldName === "mnemonic") {
    return mergeMnemonicText(existingValue, incomingValue);
  }

  return chooseUpdatedScalar(existingValue, incomingValue);
}

function chooseUpdatedScalar(existingValue, incomingValue) {
  return safeString(incomingValue) || safeString(existingValue);
}

function mergeWordLists(existingWords, incomingWords) {
  const merged = normalizeWordList(existingWords).map((word) => ({ ...word }));
  const indexes = new Map();

  merged.forEach((word, index) => {
    indexes.set(getWordMergeKey(word), index);
  });

  normalizeWordList(incomingWords).forEach((word) => {
    const key = getWordMergeKey(word);
    const existingIndex = indexes.get(key);
    if (existingIndex === undefined) {
      indexes.set(key, merged.length);
      merged.push({
        ...word,
        id: word.id || uid("word")
      });
      return;
    }

    if (!merged[existingIndex].definition && word.definition) {
      merged[existingIndex] = {
        ...merged[existingIndex],
        definition: word.definition
      };
    }
  });

  return merged;
}

function getWordMergeKey(word) {
  return `${safeString(word.word).toLowerCase()}|${normalizeReading(word.reading)}`;
}

function createBlankDeck() {
  if (state.unsavedChanges && !window.confirm("Start a fresh blank deck and discard unsaved changes?")) {
    return;
  }

  state.deck = createEmptyDeck();
  state.sourceLabel = "New deck";
  state.unsavedChanges = false;
  state.deckFileHandle = null;
  state.bulkImportDraft = null;
  clearActiveReview();
  clearReviewOrder();
  clearCramSession();
  refreshDerivedState();
  resetEditor();
  persistAutosave();
  render();
  setStatus("Started a new blank deck.", "info");
}

function loadDemoDeck() {
  if (state.unsavedChanges && !window.confirm("Load the demo deck and replace unsaved changes?")) {
    return;
  }

  state.deck = createDemoDeck();
  state.sourceLabel = "Demo deck";
  state.unsavedChanges = true;
  state.deckFileHandle = null;
  state.bulkImportDraft = null;
  clearActiveReview();
  clearReviewOrder();
  clearCramSession();
  refreshDerivedState();
  resetEditor();
  persistAutosave();
  render();
  setStatus("Loaded the demo deck. Export it to keep a portable copy.", "info");
}

async function restoreBackupDeck() {
  const restoreState = await hydrateDeckFromBrowserBackup();
  if (restoreState.missing) {
    setStatus("No browser backup was found for this app in this browser.", "error");
    return;
  }

  if (restoreState.restored && restoreState.backup) {
    deckLoadToken += 1;
    applyBrowserBackup(restoreState.backup);
    render();
    setStatus("Restored the browser backup. Export JSON to make it portable again.", "success");
  } else {
    setStatus("The browser backup exists but could not be restored.", "error");
  }
}

async function hydrateDeckFromBrowserBackup() {
  const indexedBackup = parseBrowserBackupPayload(await readBrowserBackupRecord());
  if (indexedBackup) {
    return {
      restored: true,
      missing: false,
      error: false,
      backup: indexedBackup
    };
  }

  const backup = localStorage.getItem(STORAGE_KEY);
  if (!backup) {
    return {
      restored: false,
      missing: true,
      error: false,
      backup: null
    };
  }

  try {
    const parsedBackup = parseBrowserBackupPayload(JSON.parse(backup));
    if (!parsedBackup) {
      throw new Error("Browser backup payload was empty.");
    }

    return {
      restored: true,
      missing: false,
      error: false,
      backup: parsedBackup
    };
  } catch (error) {
    return {
      restored: false,
      missing: false,
      error: true,
      backup: null
    };
  }
}

function applyBrowserBackup(backup) {
  applyLoadedDeck(backup.deck, {
    sourceLabel: backup.sourceLabel,
    unsavedChanges: true,
    deckFileHandle: null
  });
  void persistIndexedBrowserBackup();
}

function parseBrowserBackupPayload(rawPayload) {
  return parseImportedDeckPayload(rawPayload, "Browser backup");
}

function parseImportedDeckPayload(rawPayload, fallbackSourceLabel) {
  if (!rawPayload || typeof rawPayload !== "object") {
    return null;
  }

  const deckPayload = extractDeckPayload(rawPayload);
  if (!deckPayload) {
    return null;
  }

  return {
    deck: normalizeDeck(deckPayload),
    sourceLabel: safeString(rawPayload.sourceLabel) || fallbackSourceLabel || "Imported deck"
  };
}

function extractDeckPayload(rawPayload) {
  const candidates = [
    rawPayload,
    rawPayload && rawPayload.deck,
    rawPayload && rawPayload.payload,
    rawPayload && rawPayload.payload && rawPayload.payload.deck,
    rawPayload && rawPayload.backup,
    rawPayload && rawPayload.backup && rawPayload.backup.deck,
    rawPayload && rawPayload.data,
    rawPayload && rawPayload.data && rawPayload.data.deck,
    rawPayload && rawPayload.state,
    rawPayload && rawPayload.state && rawPayload.state.deck
  ];

  return candidates.find(looksLikeDeckPayload) || null;
}

function looksLikeDeckPayload(candidate) {
  return Boolean(candidate)
    && typeof candidate === "object"
    && (
      Array.isArray(candidate.kanji)
      || (candidate.progress && typeof candidate.progress === "object")
      || (candidate.meta && typeof candidate.meta === "object")
    );
}

function applyLoadedDeck(deck, options = {}) {
  state.deck = normalizeDeck(deck);
  state.sourceLabel = options.sourceLabel || "New deck";
  state.unsavedChanges = Boolean(options.unsavedChanges);
  state.deckFileHandle = options.deckFileHandle || null;
  state.bulkImportDraft = null;
  clearActiveReview();
  clearReviewOrder();
  clearCramSession();
  refreshDerivedState();
  resetEditor();
}

function persistAutosave() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state.deck));
  } catch (error) {
    // Keep going so larger decks can still land in IndexedDB.
  }

  void persistIndexedBrowserBackup();
}

async function persistIndexedBrowserBackup() {
  await writeBrowserBackupRecord({
    version: APP_VERSION,
    sourceLabel: state.sourceLabel,
    unsavedChanges: state.unsavedChanges,
    savedAt: new Date().toISOString(),
    deck: state.deck
  });
}

async function readBrowserBackupRecord() {
  const db = await getBrowserBackupDb();
  if (!db) {
    return null;
  }

  return new Promise((resolve) => {
    const transaction = db.transaction(BROWSER_BACKUP_STORE_NAME, "readonly");
    const request = transaction.objectStore(BROWSER_BACKUP_STORE_NAME).get(BROWSER_BACKUP_KEY);
    request.onsuccess = () => resolve(request.result ?? null);
    request.onerror = () => resolve(null);
    transaction.onabort = () => resolve(null);
  });
}

async function writeBrowserBackupRecord(payload) {
  const db = await getBrowserBackupDb();
  if (!db) {
    return false;
  }

  return new Promise((resolve) => {
    const transaction = db.transaction(BROWSER_BACKUP_STORE_NAME, "readwrite");
    transaction.objectStore(BROWSER_BACKUP_STORE_NAME).put(payload, BROWSER_BACKUP_KEY);
    transaction.oncomplete = () => resolve(true);
    transaction.onerror = () => resolve(false);
    transaction.onabort = () => resolve(false);
  });
}

async function getBrowserBackupDb() {
  if (!("indexedDB" in window)) {
    return null;
  }

  if (!browserBackupDbPromise) {
    browserBackupDbPromise = new Promise((resolve) => {
      try {
        const request = window.indexedDB.open(BROWSER_BACKUP_DB_NAME, 1);
        request.onupgradeneeded = () => {
          const db = request.result;
          if (!db.objectStoreNames.contains(BROWSER_BACKUP_STORE_NAME)) {
            db.createObjectStore(BROWSER_BACKUP_STORE_NAME);
          }
        };
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => resolve(null);
        request.onblocked = () => resolve(null);
      } catch (error) {
        resolve(null);
      }
    });
  }

  return browserBackupDbPromise;
}

function markDirty(message) {
  touchAutosave();
  setStatus(message, "success");
}

function touchAutosave() {
  state.unsavedChanges = true;
  persistAutosave();
}

function setStatus(message, tone) {
  dom.statusMessage.textContent = message;
  document.body.dataset.statusTone = tone || "info";
}

function parseListFromField(value) {
  const text = safeString(value);
  if (!text) {
    return [];
  }

  const separator = text.includes("\n") ? /\r?\n/ : /[,;、]+/;
  return uniqueList(text.split(separator).map((entry) => safeString(entry)).filter(Boolean));
}

function normalizeList(value) {
  if (Array.isArray(value)) {
    return value.map((entry) => safeString(entry)).filter(Boolean);
  }
  if (typeof value === "string") {
    return parseListFromField(value);
  }
  return [];
}

function uniqueList(items) {
  const seen = new Set();
  return items.filter((item) => {
    if (seen.has(item)) {
      return false;
    }
    seen.add(item);
    return true;
  });
}

function safeString(value) {
  return typeof value === "string" ? value.trim() : "";
}

function safeDateString(value) {
  if (typeof value !== "string") {
    return null;
  }
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? null : date.toISOString();
}

function clampNumber(value, min, max, fallback) {
  const number = Number(value);
  if (!Number.isFinite(number)) {
    return fallback;
  }
  return Math.min(max, Math.max(min, number));
}

function uid(prefix) {
  const stamp = Date.now().toString(36);
  const random = Math.random().toString(36).slice(2, 10);
  return `${prefix}-${stamp}-${random}`;
}

function pickRandom(items) {
  if (!items.length) {
    return null;
  }
  const index = Math.floor(Math.random() * items.length);
  return items[index];
}

function shuffle(items) {
  const copy = [...items];
  for (let index = copy.length - 1; index > 0; index -= 1) {
    const swapIndex = Math.floor(Math.random() * (index + 1));
    [copy[index], copy[swapIndex]] = [copy[swapIndex], copy[index]];
  }
  return copy;
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
