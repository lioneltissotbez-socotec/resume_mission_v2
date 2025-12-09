// Dashboard Missions – Synthèse Liciel (version PRO fiches missions)
// ----------------------------------------------------
// - Scan de dossiers LICIEL via File System Access API
// - Lecture des XML Table_General_Bien / _conclusions
// - Photo principale : /photos/presentation.jpg
// - Étiquettes DPE : /images/DPE_2020_Etiquette_Energie.jpg & _CO2.jpg
// - Filtres avancés (DO, proprio, opérateur, mission, conclusion)
// - Export CSV numéros + JSON complet réutilisable
// - Encodage corrigé (UTF-8 / Windows-1252)
// - Fiche mission détaillée (blocs 1/2/3/4)
// ----------------------------------------------------

let rootDirHandle = null;
let allMissions = [];
let filteredMissions = [];
let isScanning = false;
let currentFicheMission = null;

// Types de missions indexés sur LiColonne_Mission_Missions_programmees
const MISSION_TYPES = [
  "Amiante (DTA)",                 // 00
  "Amiante (Vente)",               // 01
  "Amiante (Travaux)",             // 02
  "Amiante (Démolition)",          // 03
  "Diagnostic Termites",           // 04
  "Diagnostic Parasites",          // 05
  "Métrage (Carrez)",              // 06
  "CREP",                          // 07
  "Assainissement",                // 08
  "Piscine",                       // 09
  "Gaz",                           // 10
  "Électricité",                   // 11
  "Diagnostic Technique Global (DTG)", // 12
  "DPE",                           // 13
  "Prêt à taux zéro",              // 14
  "ERP / ESRIS",                   // 15
  "État d’Habitabilité",           // 16
  "État des lieux",                // 17
  "Plomb dans l’eau",              // 18
  "Ascenseur",                     // 19
  "Radon",                         // 20
  "Diagnostic Incendie",           // 21
  "Accessibilité Handicapé",       // 22
  "Mesurage (Boutin)",             // 23
  "Amiante (DAPP)",                // 24
  "DRIPP",                         // 25
  "Performance Numérique",         // 26
  "Infiltrométrie",                // 27
  "Amiante (Avant Travaux)",       // 28
  "Gestion Déchets / PEMD",        // 29
  "Plomb (Après Travaux)",         // 30
  "Amiante (Contrôle périodique)", // 31
  "Empoussièrement",               // 32
  "Module Interne",                // 33
  "Home Inspection",               // 34
  "Home Inspection 4PT",           // 35
  "Wind Mitigation",               // 36
  "Plomb (Avant Travaux)",         // 37
  "Amiante (HAP)",                 // 38
  "[Non utilisé]",                 // 39
  "DPEG"                           // 40
];

// Champs de Table_Z_Conclusions.xml → diagnostics lisibles
const CONCLUSION_FIELDS = [
  { tag: "LiColonne_Variable_resume_conclusion_termites",        label: "Termites" },
  { tag: "LiColonne_Variable_resume_conclusion_autres_parasites",label: "Parasites" },
  { tag: "LiColonne_Variable_resume_conclusion_amiante",         label: "Amiante" },
  { tag: "LiColonne_Variable_resume_conclusion_carrez",          label: "Loi Carrez" },
  { tag: "LiColonne_Variable_resume_conclusion_crep",            label: "Plomb (CREP)" },
  { tag: "LiColonne_Variable_resume_conclusion_dpe",             label: "DPE" },
  { tag: "LiColonne_Variable_resume_conclusion_gaz",             label: "Gaz" },
  { tag: "LiColonne_Variable_resume_conclusion_elec",            label: "Électricité" },
  { tag: "LiColonne_Variable_resume_conclusion_ernt",            label: "ERP / ESRIS" },
  { tag: "LiColonne_Variable_resume_conclusion_Assainissement",  label: "Assainissement" },
  { tag: "LiColonne_Variable_resume_conclusion_PTZ",             label: "PTZ" },
  { tag: "LiColonne_Variable_resume_conclusion_Piscine",         label: "Piscine" },
  { tag: "LiColonne_Variable_resume_conclusion_SRU",             label: "SRU" },
  { tag: "LiColonne_Variable_resume_conclusion_RADON",           label: "Radon" },
  { tag: "LiColonne_Variable_resume_conclusion_Ascenseur",       label: "Ascenseur" },
  { tag: "LiColonne_Variable_resume_conclusion_Robien",          label: "Robien" },
  { tag: "LiColonne_Variable_resume_conclusion_EDL",             label: "État des lieux" },
  { tag: "LiColonne_Variable_resume_conclusion_Incendie",        label: "Incendie" },
  { tag: "LiColonne_Variable_resume_conclusion_Handicape",       label: "Accessibilité handicapé" },
  { tag: "LiColonne_Variable_resume_conclusion_Infiltrometrie",  label: "Infiltrométrie" },
  { tag: "LiColonne_Variable_resume_conclusion_Pb_eau",          label: "Plomb dans l'eau" }
];


// Helpers DOM
const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => Array.from(document.querySelectorAll(sel));

// Petit debounce pour filtrage temps réel
function debounce(fn, delay = 150) {
  let timer;
  return (...args) => {
    clearTimeout(timer);
    timer = setTimeout(() => fn(...args), delay);
  };
}

document.addEventListener("DOMContentLoaded", () => {
  initUI();
});

function initUI() {
  const btnPickRoot = $("#btnPickRoot");
  const btnScan = $("#btnScan");
  const scanModeRadios = $$("input[name='scanMode']");

  const prefixBlock = $("#prefixBlock");
  const listBlock = $("#listBlock");

  const btnApplyFilters = $("#btnApplyFilters");
  const btnResetFilters = $("#btnResetFilters");

  const btnExportCSV = $("#btnExportCSV");
  const btnCopyClipboard = $("#btnCopyClipboard");
  const btnExportJSON = $("#btnExportJSON");
  const btnImportJSON = $("#btnImportJSON");
  const jsonFileInput = $("#jsonFileInput");

  const btnCloseModal = $("#btnCloseModal");
  const btnClosePhoto = $("#btnClosePhoto");
  const btnExportFiche = $("#btnExportFiche");

  // 👉 on RELAXE la condition : on regarde seulement si showDirectoryPicker existe
  const hasFsAccess = typeof window.showDirectoryPicker === "function";

  if (!hasFsAccess) {
    alert(
      "⚠️ Votre navigateur ne supporte pas la File System Access API (showDirectoryPicker). Utilisez Chrome / Edge récents."
    );
    btnPickRoot.disabled = true;
    $("#rootInfo").textContent = "La sélection de dossiers nécessite un navigateur compatible (Chrome/Edge).";
  }

  // Gestion changement de mode de scan
  scanModeRadios.forEach((radio) => {
    radio.addEventListener("change", () => {
      const mode = getScanMode();
      prefixBlock.classList.toggle("hidden", mode !== "prefix");
      listBlock.classList.toggle("hidden", mode !== "list");
    });
  });

  btnPickRoot.addEventListener("click", onPickRoot);
  btnScan.addEventListener("click", onScan);

  // Filtres : bouton + auto-apply
  const debouncedApply = debounce(applyFilters, 200);
  btnApplyFilters.addEventListener("click", applyFilters);
  btnResetFilters.addEventListener("click", resetFilters);

  ["#filterDO", "#filterProp", "#filterOp", "#filterType"].forEach((sel) => {
    const el = $(sel);
    el.addEventListener("change", debouncedApply);
  });
  $("#filterConclusion").addEventListener("input", debouncedApply);
  $("#filterConclusion").addEventListener("keyup", (e) => {
    if (e.key === "Enter") applyFilters();
  });

  // Export / JSON
  btnExportCSV.addEventListener("click", exportFilteredAsCSV);
  btnCopyClipboard.addEventListener("click", copyFilteredToClipboard);
  btnExportJSON.addEventListener("click", exportAllAsJSON);
  btnImportJSON.addEventListener("click", () => jsonFileInput.click());
  jsonFileInput.addEventListener("change", onImportJSON);

  // Modales
  btnCloseModal.addEventListener("click", closeConclusionModal);
  $("#modalOverlay").addEventListener("click", (e) => {
    if (e.target.id === "modalOverlay") closeConclusionModal();
  });

  btnClosePhoto.addEventListener("click", closePhotoModal);
  $("#photoOverlay").addEventListener("click", (e) => {
    if (e.target.id === "photoOverlay") closePhotoModal();
  });

  // Export fiche mission (une seule mission)
  btnExportFiche.addEventListener("click", () => {
    if (!currentFicheMission) {
      alert("Aucune fiche mission active à exporter.");
      return;
    }
    exportSingleMissionAsJSON(currentFicheMission);
  });

  updateProgress(0, 0, "En attente…");
  updateStats();
  updateExportButtonsState();
}

function getScanMode() {
  const selected = document.querySelector("input[name='scanMode']:checked");
  return selected ? selected.value : "all";
}

async function onPickRoot() {
  if (!window.showDirectoryPicker) {
    alert(
      "Sélection impossible : votre navigateur ne supporte pas showDirectoryPicker (File System Access API)."
    );
    return;
  }

  try {
    rootDirHandle = await window.showDirectoryPicker();
    $("#rootInfo").textContent = "Dossier racine : " + rootDirHandle.name;
    $("#btnScan").disabled = false;
  } catch (err) {
    console.warn("Sélection de dossier annulée :", err);
  }
}

async function onScan() {
  if (!rootDirHandle) {
    alert("Veuillez d'abord choisir un dossier racine.");
    return;
  }
  if (isScanning) return;

  const mode = getScanMode();
  const prefix = $("#inputPrefix").value.trim();
  const listText = $("#inputDossierList").value.trim();

  let listItems = [];
  if (listText) {
    listItems = listText
      .split(/[\s,;]+/)
      .map((s) => s.trim())
      .filter((s) => !!s);
  }

  if (mode === "prefix" && !prefix) {
    alert("Veuillez saisir un préfixe de dossier.");
    return;
  }
  if (mode === "list" && listItems.length === 0) {
    alert("Veuillez coller au moins un numéro de dossier.");
    return;
  }

  // Reset data
  allMissions = [];
  filteredMissions = [];
  renderTable();
  updateStats();
  updateExportButtonsState();

  setScanningState(true);
  $("#progressText").textContent = "Scan des sous-dossiers en cours…";
  updateProgress(0, 0, "Préparation du scan…");

  try {
    const allEntries = [];
    for await (const [name, handle] of rootDirHandle.entries()) {
      if (handle.kind === "directory") {
        allEntries.push({ name, handle });
      }
    }

    const candidates = allEntries.filter(({ name }) => {
      if (mode === "all") return true;
      if (mode === "prefix") {
        return name.startsWith(prefix);
      }
      if (mode === "list") {
        return listItems.some((item) => name.startsWith(item));
      }
      return true;
    });

    if (candidates.length === 0) {
      alert("Aucun dossier ne correspond aux critères.");
      updateProgress(0, 0, "Aucun dossier trouvé.");
      return;
    }

    if (candidates.length > 100) {
      const proceed = confirm(
        `Vous allez scanner ${candidates.length} dossiers. Voulez-vous continuer ?`
      );
      if (!proceed) {
        updateProgress(0, 0, "Scan annulé par l'utilisateur.");
        return;
      }
    }

    let processed = 0;
    const total = candidates.length;

    for (const { name, handle } of candidates) {
      try {
        const mission = await processMissionFolder(handle, name);
        if (mission) {
          allMissions.push(mission);
        }
      } catch (err) {
        console.error("Erreur lors du traitement du dossier", name, err);
      }
      processed++;
      updateProgress(processed, total, `Scan : ${processed} / ${total} dossiers…`);
    }

    filteredMissions = [...allMissions];
    populateFilterOptions();
    renderTable();
    updateStats();
    updateExportButtonsState();
    updateProgress(total, total, `Scan terminé : ${allMissions.length} missions valides.`);

    if (allMissions.length > 0) {
      $("#filtersSection").classList.remove("hidden-block");
    }
  } catch (err) {
    console.error("Erreur globale de scan :", err);
    alert("Erreur lors du scan des dossiers. Voir la console pour le détail.");
    updateProgress(0, 0, "Erreur lors du scan.");
  } finally {
    setScanningState(false);
  }
}

function setScanningState(scanning) {
  isScanning = scanning;
  $("#btnScan").disabled = scanning || !rootDirHandle;
  $("#btnPickRoot").disabled = scanning;
}

// --------------------------------------------------------
// Lecture générique d'un XML avec gestion d'encodage
// --------------------------------------------------------
async function readXmlFile(dirHandle, fileName) {
  try {
    const fileHandle = await dirHandle.getFileHandle(fileName);
    const file = await fileHandle.getFile();
    const buffer = await file.arrayBuffer();

    // Tentative UTF-8 par défaut
    let textUtf8 = new TextDecoder("utf-8", { fatal: false }).decode(buffer);
    let text = textUtf8;

    // Encodage déclaré ?
    const m = textUtf8.match(/encoding="([^"]+)"/i);
    const declared = m && m[1] ? m[1].toLowerCase() : null;

    if (declared && declared !== "utf-8" && declared !== "utf8") {
      try {
        text = new TextDecoder(declared).decode(buffer);
      } catch (e) {
        // Fallback classique LICIEL
        try {
          text = new TextDecoder("windows-1252").decode(buffer);
        } catch (e2) {
          // On garde UTF-8
        }
      }
    } else {
      // Heuristique : si beaucoup de �, on tente windows-1252
      const badUtf8 = (textUtf8.match(/�/g) || []).length;
      if (badUtf8 >= 3) {
        try {
          const text1252 = new TextDecoder("windows-1252").decode(buffer);
          const bad1252 = (text1252.match(/�/g) || []).length;
          if (bad1252 < badUtf8) {
            text = text1252;
          }
        } catch (e) {
          // On garde UTF-8
        }
      }
    }

    const parser = new DOMParser();
    const doc = parser.parseFromString(text, "application/xml");
    if (doc.querySelector("parsererror")) {
      console.error("Erreur de parsing XML pour", fileName);
      return null;
    }
    return doc;
  } catch (err) {
    console.warn("Impossible de lire le fichier XML", fileName, err);
    return null;
  }
}

function getXmlValue(xmlDoc, tagName) {
  const el = xmlDoc.querySelector(tagName);
  return el ? (el.textContent || "").trim() : "";
}

function decodeMissions(bits) {
  if (!bits) return [];
  const result = [];
  for (let i = 0; i < bits.length && i < .length; i++) {
    if (bits[i] === "1") {
      result.push([i]);
    }
  }
  return result;
}
// Analyse du texte DPE pour extraire conso, classes et N° ADEME
function parseDpeInfo(dpeText) {
  if (!dpeText) return null;
  const txt = dpeText.replace(/\s+/g, " ");

  const consoMatch = txt.match(/Consommation.*?:\s*([0-9]+)\s*kWh/i);
  const classeEnerMatch = txt.match(/Consommation.*?\(Classe\s*([A-G])\)/i);

  const co2Match = txt.match(/émissions.*?:\s*([0-9]+)\s*kg/i);
  const classeCo2Match = txt.match(/émissions.*?\(Classe\s*([A-G])\)/i);

  const ademeMatch = txt.match(/N[°º]\s*ADEME\s*:\s*([A-Z0-9]+)/i);

  return {
    conso: consoMatch ? parseInt(consoMatch[1], 10) : null,
    classeEner: classeEnerMatch ? classeEnerMatch[1].toUpperCase() : null,
    co2: co2Match ? parseInt(co2Match[1], 10) : null,
    classeCO2: classeCo2Match ? classeCo2Match[1].toUpperCase() : null,
    ademe: ademeMatch ? ademeMatch[1] : null
  };
}

// --------------------------------------------------------
// Découpe du bloc de conclusion en conclusions par type
// --------------------------------------------------------
function buildConclusionsList(rawText, missionsEffectuees) {
  if (!rawText) return [];
  let txt = rawText.replace(/\s+/g, " ").trim();

  const setEffectuees = new Set(missionsEffectuees || []);

  const found = [];
  .forEach((label) => {
    const idx = txt.indexOf(label);
    if (idx !== -1) {
      found.push({ label, idx });
    }
  });

  if (!found.length) return [];

  found.sort((a, b) => a.idx - b.idx);

  const results = [];
  for (let i = 0; i < found.length; i++) {
    const { label, idx } = found[i];
    const start = idx + label.length;
    const end = i + 1 < found.length ? found[i + 1].idx : txt.length;

    let chunk = txt.slice(start, end).trim();
    // Nettoyage des indices/numéros en début de chaîne (0 0, 6 6, etc.)
    chunk = chunk.replace(/^[0-9\s:()\-]+/, "").trim();

    if (!chunk) continue;
    if (setEffectuees.size && !setEffectuees.has(label)) continue;

    results.push({
      type: label,
      text: chunk
    });
  }

  return results;
}

// --------------------------------------------------------
// Traitement d'un dossier de mission
// --------------------------------------------------------
async function processMissionFolder(folderHandle, folderName) {
  let xmlDir;
  try {
    xmlDir = await folderHandle.getDirectoryHandle("XML");
  } catch (err) {
    console.warn(`Dossier XML manquant dans ${folderName}`);
    return null;
  }

  const bienXml = await readXmlFile(xmlDir, "Table_General_Bien.xml");
  if (!bienXml) {
    console.warn(`Table_General_Bien.xml manquant ou invalide dans ${folderName}`);
    return null;
  }

  const mission = {};

  mission.numDossier = getXmlValue(bienXml, "LiColonne_Mission_Num_Dossier") || folderName;

  mission.donneurOrdre = {
    nom: getXmlValue(bienXml, "LiColonne_DOrdre_Nom"),
    entete: getXmlValue(bienXml, "LiColonne_DOrdre_Entete"),
    adresse: getXmlValue(bienXml, "LiColonne_DOrdre_Adresse1"),
    departement: getXmlValue(bienXml, "LiColonne_DOrdre_Departement"),
    commune: getXmlValue(bienXml, "LiColonne_DOrdre_Commune")
  };

  mission.proprietaire = {
    entete: getXmlValue(bienXml, "LiColonne_Prop_Entete"),
    nom: getXmlValue(bienXml, "LiColonne_Prop_Nom"),
    adresse: getXmlValue(bienXml, "LiColonne_Prop_Adresse1"),
    departement: getXmlValue(bienXml, "LiColonne_Prop_Departement"),
    commune: getXmlValue(bienXml, "LiColonne_Prop_Commune")
  };

  mission.immeuble = {
    adresse: getXmlValue(bienXml, "LiColonne_Immeuble_Adresse1"),
    departement: getXmlValue(bienXml, "LiColonne_Immeuble_Departement"),
    commune: getXmlValue(bienXml, "LiColonne_Immeuble_Commune"),
    lot: getXmlValue(bienXml, "LiColonne_Immeuble_Lot"),
    natureBien: getXmlValue(bienXml, "LiColonne_Immeuble_Nature_bien"),
    typeBien: getXmlValue(bienXml, "LiColonne_Immeuble_Type_bien"),
    typeDossier: getXmlValue(bienXml, "LiColonne_Immeuble_Type_Dossier"),
    description: getXmlValue(bienXml, "LiColonne_Immeuble_Description")
  };

  const missionsProgrammes = getXmlValue(bienXml, "LiColonne_Mission_Missions_programmees");
  mission.mission = {
    dateVisite: getXmlValue(bienXml, "LiColonne_Mission_Date_Visite"),
    dateRapport: getXmlValue(bienXml, "LiColonne_Mission_Date_Rapport"),
    missionsProgrammes,
    missionsEffectuees: decodeMissions(missionsProgrammes)
  };

  mission.operateur = {
    nomFamille: getXmlValue(bienXml, "LiColonne_Gen_Nom_operateur_UniquementNomFamille"),
    prenom: getXmlValue(bienXml, "LiColonne_Gen_Nom_operateur_UniquementPreNom"),
    certifSociete: getXmlValue(bienXml, "LiColonne_Gen_certif_societe"),
    numCertif: getXmlValue(bienXml, "LiColonne_Gen_num_certif")
  };

  // -------------------------------------------------
  // 1) Conclusions structurées : Table_Z_Conclusions.xml
  // -------------------------------------------------
  mission.conclusionsByDiag = {};
  let allConclusionsText = "";

  const zConclXml = await readXmlFile(xmlDir, "Table_Z_Conclusions.xml");
  if (zConclXml) {
    for (const field of CONCLUSION_FIELDS) {
      const txt = getXmlValue(zConclXml, field.tag);
      if (txt) {
        mission.conclusionsByDiag[field.label] = txt;
        allConclusionsText += " " + txt;
      }
    }

    // DPE : extraction des valeurs pour filtres / lien ADEME
    const dpeText = getXmlValue(zConclXml, "LiColonne_Variable_resume_conclusion_dpe");
    if (dpeText) {
      if (!mission.conclusionsByDiag["DPE"]) {
        mission.conclusionsByDiag["DPE"] = dpeText;
      }
      mission.dpe = {
        energieUrl: null,
        co2Url: null,
        ...parseDpeInfo(dpeText)
      };
    } else {
      mission.dpe = {
        energieUrl: null,
        co2Url: null,
        conso: null,
        classeEner: null,
        co2: null,
        classeCO2: null,
        ademe: null
      };
    }
  } else {
    // Fallback : pas de Table_Z_Conclusions → rien de structuré
    mission.dpe = {
      energieUrl: null,
      co2Url: null,
      conso: null,
      classeEner: null,
      co2: null,
      classeCO2: null,
      ademe: null
    };
  }

  mission.conclusionRaw = allConclusionsText.trim();
  mission.conclusion = mission.conclusionRaw;

  // ----------------------------------
  // 2) Photo présentation & étiquettes DPE (fichiers fixes)
  // ----------------------------------
  mission.photoUrl = null;
  mission.photoPath = null;

  // Photo principale : /photos/presentation.jpg
  try {
    const photoFile = await getFileFromRelativePath(folderHandle, "photos/presentation.jpg");
    const blobUrl = URL.createObjectURL(photoFile);
    mission.photoUrl = blobUrl;
    mission.photoPath = "photos/presentation.jpg";
  } catch (err) {
    // pas de photo => laissé à null
  }

  // Initialisation dpe si pas déjà fait plus haut
  if (!mission.dpe) {
    mission.dpe = {
      energieUrl: null,
      co2Url: null,
      conso: null,
      classeEner: null,
      co2: null,
      classeCO2: null,
      ademe: null
    };
  }

  // Étiquettes DPE uniquement si DPE fait partie des missions effectuées
  if (mission.mission.missionsEffectuees.includes("DPE")) {
    try {
      const fileE = await getFileFromRelativePath(
        folderHandle,
        "images/DPE_2020_Etiquette_Energie.jpg"
      );
      mission.dpe.energieUrl = URL.createObjectURL(fileE);
    } catch (err) {
      // Pas d'étiquette énergie
    }

    try {
      const fileC = await getFileFromRelativePath(
        folderHandle,
        "images/DPE_2020_Etiquette_CO2.jpg"
      );
      mission.dpe.co2Url = URL.createObjectURL(fileC);
    } catch (err) {
      // Pas d'étiquette CO2
    }
  }

  // Champs normalisés pour filtre texte
  mission._norm = {
    conclusion: allConclusionsText.toLowerCase()
  };

  return mission;
}


  // ----------------------------------
  // Photo présentation & DPE (fichiers fixes)
  // ----------------------------------
  mission.photoUrl = null;
  mission.photoPath = null;
  mission.dpe = {
    energieUrl: null,
    co2Url: null
  };

  // Photo principale : /photos/presentation.jpg
  try {
    const photoFile = await getFileFromRelativePath(folderHandle, "photos/presentation.jpg");
    const blobUrl = URL.createObjectURL(photoFile);
    mission.photoUrl = blobUrl;
    mission.photoPath = "photos/presentation.jpg";
  } catch (err) {
    // Pas de photo => laissé à null
    // console.warn("Photo de présentation introuvable dans", folderName);
  }

  // Étiquettes DPE uniquement si DPE fait partie des missions effectuées
  if (mission.mission.missionsEffectuees.includes("DPE")) {
    try {
      const fileE = await getFileFromRelativePath(
        folderHandle,
        "images/DPE_2020_Etiquette_Energie.jpg"
      );
      mission.dpe.energieUrl = URL.createObjectURL(fileE);
    } catch (err) {
      // Pas d'étiquette énergie
    }

    try {
      const fileC = await getFileFromRelativePath(
        folderHandle,
        "images/DPE_2020_Etiquette_CO2.jpg"
      );
      mission.dpe.co2Url = URL.createObjectURL(fileC);
    } catch (err) {
      // Pas d'étiquette CO2
    }
  }

  // Champs normalisés pour filtres texte
  mission._norm = {
    conclusion: (conclusionRaw || "").toLowerCase()
  };

  return mission;
}

// Parcours d'un chemin relatif
async function getFileFromRelativePath(rootHandle, relPath) {
  const parts = relPath.split("/").filter((p) => !!p && p !== ".");

  let current = rootHandle;
  for (let i = 0; i < parts.length; i++) {
    const part = parts[i];
    const isLast = i === parts.length - 1;

    if (isLast) {
      const fileHandle = await current.getFileHandle(part);
      return fileHandle.getFile();
    } else {
      current = await current.getDirectoryHandle(part);
    }
  }
  throw new Error("Chemin vide");
}

// ----------------------------------------
// Filtres & rendu
// ----------------------------------------
function updateProgress(done, total, text) {
  const bar = $("#progressFill");
  const label = $("#progressText");
  let percent = 0;
  if (total > 0) {
    percent = Math.round((done / total) * 100);
  }
  bar.style.width = percent + "%";
  if (label) label.textContent = text || "";
}

function populateFilterOptions() {
  const doSelect = $("#filterDO");
  const propSelect = $("#filterProp");
  const opSelect = $("#filterOp");
  const typeSelect = $("#filterType");

  doSelect.innerHTML = "";
  propSelect.innerHTML = "";
  opSelect.innerHTML = "";
  typeSelect.innerHTML = "";

  const DOset = new Set();
  const PropSet = new Set();
  const OpSet = new Set();
  const TypeSet = new Set();

  for (const m of allMissions) {
    const doLabel = formatDonneurOrdre(m);
    if (doLabel) DOset.add(doLabel);

    const propLabel = formatProprietaire(m);
    if (propLabel) PropSet.add(propLabel);

    const opLabel = formatOperateur(m);
    if (opLabel) OpSet.add(opLabel);

    (m.mission.missionsEffectuees || []).forEach((t) => TypeSet.add(t));
  }

  [...DOset].sort().forEach((val) => {
    const opt = document.createElement("option");
    opt.value = val;
    opt.textContent = val;
    doSelect.appendChild(opt);
  });

  [...PropSet].sort().forEach((val) => {
    const opt = document.createElement("option");
    opt.value = val;
    opt.textContent = val;
    propSelect.appendChild(opt);
  });

  [...OpSet].sort().forEach((val) => {
    const opt = document.createElement("option");
    opt.value = val;
    opt.textContent = val;
    opSelect.appendChild(opt);
  });

  [...TypeSet].sort().forEach((val) => {
    const opt = document.createElement("option");
    opt.value = val;
    opt.textContent = val;
    typeSelect.appendChild(opt);
  });
}

function applyFilters() {
  const doValues = getSelectedOptions("#filterDO");
  const propValues = getSelectedOptions("#filterProp");
  const opValues = getSelectedOptions("#filterOp");
  const typeValues = getSelectedOptions("#filterType");
  const conclText = ($("#filterConclusion").value || "").trim().toLowerCase();

  filteredMissions = allMissions.filter((m) => {
    if (doValues.length) {
      const label = formatDonneurOrdre(m);
      if (!doValues.includes(label)) return false;
    }

    if (propValues.length) {
      const label = formatProprietaire(m);
      if (!propValues.includes(label)) return false;
    }

    if (opValues.length) {
      const label = formatOperateur(m);
      if (!opValues.includes(label)) return false;
    }

    if (typeValues.length) {
      const missionTypes = m.mission.missionsEffectuees || [];
      const ok = missionTypes.some((t) => typeValues.includes(t));
      if (!ok) return false;
    }

    if (conclText) {
      const c = m._norm?.conclusion ?? (m.conclusion || "").toLowerCase();
      if (!c.includes(conclText)) return false;
    }
  // Filtres DPE (optionnels si les champs existent dans le HTML)
  const dpeClassSelect = $("#filterDpeClass");
  const dpeCo2ClassSelect = $("#filterDpeCo2Class");
  const dpeConsoMinInput = $("#filterDpeConsoMin");
  const dpeConsoMaxInput = $("#filterDpeConsoMax");
  const dpeCo2MinInput = $("#filterDpeCo2Min");
  const dpeCo2MaxInput = $("#filterDpeCo2Max");

  const dpeClass = dpeClassSelect ? dpeClassSelect.value : "";
  const dpeCo2Class = dpeCo2ClassSelect ? dpeCo2ClassSelect.value : "";

  const consoMin = dpeConsoMinInput && dpeConsoMinInput.value
    ? parseFloat(dpeConsoMinInput.value)
    : null;
  const consoMax = dpeConsoMaxInput && dpeConsoMaxInput.value
    ? parseFloat(dpeConsoMaxInput.value)
    : null;
  const co2Min = dpeCo2MinInput && dpeCo2MinInput.value
    ? parseFloat(dpeCo2MinInput.value)
    : null;
  const co2Max = dpeCo2MaxInput && dpeCo2MaxInput.value
    ? parseFloat(dpeCo2MaxInput.value)
    : null;
    // Filtres DPE par lettre / seuils
    if (dpeClass) {
      const c = m.dpe && m.dpe.classeEner;
      if (!c || c !== dpeClass) return false;
    }

    if (dpeCo2Class) {
      const c2 = m.dpe && m.dpe.classeCO2;
      if (!c2 || c2 !== dpeCo2Class) return false;
    }

    if (consoMin != null) {
      const v = m.dpe && m.dpe.conso;
      if (v == null || v < consoMin) return false;
    }
    if (consoMax != null) {
      const v = m.dpe && m.dpe.conso;
      if (v == null || v > consoMax) return false;
    }

    if (co2Min != null) {
      const v = m.dpe && m.dpe.co2;
      if (v == null || v < co2Min) return false;
    }
    if (co2Max != null) {
      const v = m.dpe && m.dpe.co2;
      if (v == null || v > co2Max) return false;
    }

    return true;
  });

  renderTable();
  updateStats();
  updateExportButtonsState();
}

function resetFilters() {
  ["#filterDO", "#filterProp", "#filterOp", "#filterType"].forEach((sel) => {
    const el = $(sel);
    if (!el) return;
    Array.from(el.options).forEach((opt) => (opt.selected = false));
  });
  $("#filterConclusion").value = "";
  filteredMissions = [...allMissions];
  renderTable();
  updateStats();
  updateExportButtonsState();
}

function getSelectedOptions(selector) {
  const select = $(selector);
  if (!select) return [];
  return Array.from(select.selectedOptions).map((o) => o.value);
}

function formatDonneurOrdre(m) {
  const d = m.donneurOrdre || {};
  const parts = [];
  if (d.entete) parts.push(d.entete);
  if (d.nom) parts.push(d.nom);
  return parts.join(" ").trim();
}

function formatProprietaire(m) {
  const p = m.proprietaire || {};
  const parts = [];
  if (p.entete) parts.push(p.entete);
  if (p.nom) parts.push(p.nom);
  return parts.join(" ").trim();
}

function formatOperateur(m) {
  const o = m.operateur || {};
  const parts = [];
  if (o.nomFamille || o.prenom) {
    parts.push([o.nomFamille, o.prenom].filter(Boolean).join(" "));
  }
  if (o.certifSociete) {
    parts.push("(" + o.certifSociete + ")");
  }
  return parts.join(" ").trim();
}

function renderTable() {
  const tbody = $("#resultsTable tbody");
  tbody.innerHTML = "";

  for (const m of filteredMissions) {
    const tr = document.createElement("tr");

    // 1) N° dossier
    const tdNum = document.createElement("td");
    tdNum.textContent = m.numDossier || "";
    tr.appendChild(tdNum);

    // 2) Donneur d'ordre
    const tdDO = document.createElement("td");
    tdDO.textContent = formatDonneurOrdre(m);
    tr.appendChild(tdDO);

    // 3) Propriétaire
    const tdProp = document.createElement("td");
    tdProp.textContent = formatProprietaire(m);
    tr.appendChild(tdProp);

    // 4) Adresse immeuble
    const tdAdr = document.createElement("td");
    const im = m.immeuble || {};
    tdAdr.textContent = [im.adresse, im.departement, im.commune]
      .filter(Boolean)
      .join(" ");
    tr.appendChild(tdAdr);

    // 5) Type / nature bien
    const tdType = document.createElement("td");
    tdType.textContent = [im.typeBien, im.natureBien, im.typeDossier]
      .filter(Boolean)
      .join(" / ");
    tr.appendChild(tdType);

    // 6) Dates
    const tdDates = document.createElement("td");
    const lines = [];
    if (m.mission.dateVisite) lines.push("Visite : " + m.mission.dateVisite);
    if (m.mission.dateRapport) lines.push("Rapport : " + m.mission.dateRapport);
    tdDates.textContent = lines.join("\n");
    tr.appendChild(tdDates);

    // 7) Diagnostiqueur
    const tdOp = document.createElement("td");
    tdOp.textContent = formatOperateur(m);
    if (m.operateur.numCertif) {
      const small = document.createElement("div");
      small.style.fontSize = "11px";
      small.style.color = "#6b7280";
      small.textContent = "Certif : " + m.operateur.numCertif;
      tdOp.appendChild(document.createElement("br"));
      tdOp.appendChild(small);
    }
    tr.appendChild(tdOp);

    // 8) Missions effectuées
    const tdMissions = document.createElement("td");
    (m.mission.missionsEffectuees || []).forEach((type) => {
      const span = document.createElement("span");
      span.className = "tag";
      span.textContent = type;
      tdMissions.appendChild(span);
    });
    tr.appendChild(tdMissions);

    // 9) Fiche mission (bouton)
    const tdFiche = document.createElement("td");
    const btnFiche = document.createElement("button");
    btnFiche.className = "btn-link";
    btnFiche.textContent = "Ouvrir la fiche";
    btnFiche.addEventListener("click", () => openFicheModal(m));
    tdFiche.appendChild(btnFiche);
    tr.appendChild(tdFiche);

    // 10) Photo (vignette)
    const tdPhoto = document.createElement("td");
    if (m.photoUrl) {
      const img = document.createElement("img");
      img.src = m.photoUrl;
      img.alt = "Photo présentation";
      img.className = "photo-thumb";
      img.addEventListener("click", () => openPhotoModal(m.photoUrl));
      tdPhoto.appendChild(img);
    } else if (m.photoPath) {
      const span = document.createElement("span");
      span.className = "photo-placeholder";
      span.textContent = "Chemin : " + m.photoPath;
      tdPhoto.appendChild(span);
    } else {
      const span = document.createElement("span");
      span.className = "photo-placeholder";
      span.textContent = "Aucune photo";
      tdPhoto.appendChild(span);
    }
    tr.appendChild(tdPhoto);

    tbody.appendChild(tr);
  }
}

function updateStats() {
  const statsText = $("#statsText");
  const total = allMissions.length;
  const filt = filteredMissions.length;
  statsText.textContent = `Missions : ${filt} affichées / ${total} scannées.`;
}

function updateExportButtonsState() {
  const hasData = filteredMissions.length > 0;
  $("#btnExportCSV").disabled = !hasData;
  $("#btnCopyClipboard").disabled = !hasData;
  $("#btnExportJSON").disabled = !allMissions.length;
}

// ----------------------------------------
// Fiche mission (modale détaillée)
// ----------------------------------------
function openFicheModal(mission) {
  currentFicheMission = mission;

  const overlay = $("#modalOverlay");
  const title = $("#ficheTitle");
  const subtitle = $("#ficheSubtitle");

  title.textContent = `Mission ${mission.numDossier || ""}`;

  const im = mission.immeuble || {};
  const adrParts = [im.adresse, im.commune, im.departement].filter(Boolean);
  subtitle.textContent = adrParts.join(" • ");

  // Bloc 1 : Admin
  $("#ficheDO").textContent = buildBlocTexte(mission.donneurOrdre, [
    ["Entité", "entete"],
    ["Nom", "nom"],
    ["Adresse", "adresse"],
    ["Commune", "commune"],
    ["Département", "departement"]
  ]);

  $("#ficheProp").textContent = buildBlocTexte(mission.proprietaire, [
    ["Entité", "entete"],
    ["Nom", "nom"],
    ["Adresse", "adresse"],
    ["Commune", "commune"],
    ["Département", "departement"]
  ]);

  $("#ficheImmeuble").textContent = buildBlocTexte(mission.immeuble, [
    ["Adresse", "adresse"],
    ["Commune", "commune"],
    ["Département", "departement"],
    ["Lot", "lot"],
    ["Nature du bien", "natureBien"],
    ["Type de bien", "typeBien"],
    ["Type de dossier", "typeDossier"],
    ["Description", "description"]
  ]);

  const m = mission.mission || {};
  const missionsList = (m.missionsEffectuees || []).join(", ");
  $("#ficheMission").textContent =
    [
      m.dateVisite ? `Date de visite : ${m.dateVisite}` : "",
      m.dateRapport ? `Date de rapport : ${m.dateRapport}` : "",
      missionsList ? `Diagnostics réalisés : ${missionsList}` : ""
    ]
      .filter(Boolean)
      .join("\n");

  // Bloc 2 : Photo
  const photoContainer = $("#fichePhotoContainer");
  photoContainer.innerHTML = "";
  if (mission.photoUrl) {
    const img = document.createElement("img");
    img.src = mission.photoUrl;
    img.alt = "Photo de présentation";
    img.addEventListener("click", () => openPhotoModal(mission.photoUrl));
    photoContainer.appendChild(img);
  } else {
    const span = document.createElement("div");
    span.className = "fiche-photo-placeholder";
    span.textContent = "Aucune photo trouvée pour cette mission.";
    photoContainer.appendChild(span);
  }

    // Bloc 2 : DPE (visuels + valeurs + lien ADEME)
  const dpeContainer = $("#ficheDPEContainer");
  dpeContainer.innerHTML = "";

  const hasE = mission.dpe && mission.dpe.energieUrl;
  const hasC = mission.dpe && mission.dpe.co2Url;
  const hasData =
    mission.dpe &&
    (mission.dpe.conso != null ||
      mission.dpe.classeEner ||
      mission.dpe.co2 != null ||
      mission.dpe.classeCO2 ||
      mission.dpe.ademe);

  if (!hasE && !hasC && !hasData) {
    const span = document.createElement("div");
    span.className = "fiche-dpe-placeholder";
    span.textContent = "DPE non réalisé ou données DPE indisponibles.";
    dpeContainer.appendChild(span);
  } else {
    // Lien ADEME si dispo
    if (mission.dpe.ademe) {
      const link = document.createElement("a");
      link.href =
        "https://observatoire-dpe-audit.ademe.fr/afficher-dpe/" +
        mission.dpe.ademe;
      link.target = "_blank";
      link.rel = "noopener noreferrer";
      link.textContent = "Voir le DPE sur le site officiel ADEME";
      link.className = "dpe-link";
      dpeContainer.appendChild(link);
    }

    // Valeurs DPE
    if (hasData) {
      const info = document.createElement("div");
      info.className = "dpe-info";

      const lines = [];
      if (mission.dpe.classeEner) {
        lines.push("Classe Énergie : " + mission.dpe.classeEner);
      }
      if (mission.dpe.conso != null) {
        lines.push("Consommation : " + mission.dpe.conso + " kWh/m²/an");
      }
      if (mission.dpe.classeCO2) {
        lines.push("Classe CO₂ : " + mission.dpe.classeCO2);
      }
      if (mission.dpe.co2 != null) {
        lines.push("Émissions : " + mission.dpe.co2 + " kgCO₂/m²/an");
      }

      info.textContent = lines.join(" • ");
      dpeContainer.appendChild(info);
    }

    // Étiquettes DPE
    const visuels = document.createElement("div");
    visuels.className = "dpe-visuels";

    if (hasE) {
      const imgE = document.createElement("img");
      imgE.src = mission.dpe.energieUrl;
      imgE.alt = "Étiquette énergie DPE";
      visuels.appendChild(imgE);
    }
    if (hasC) {
      const imgC = document.createElement("img");
      imgC.src = mission.dpe.co2Url;
      imgC.alt = "Étiquette CO₂ DPE";
      visuels.appendChild(imgC);
    }

    if (hasE || hasC) {
      dpeContainer.appendChild(visuels);
    }
  }


   // Bloc 3 : cartes diagnostics (tous ceux non vides)
  const diagsContainer = $("#ficheDiagsContainer");
  diagsContainer.innerHTML = "";

  const entries = Object.entries(mission.conclusionsByDiag || {});
  if (!entries.length) {
    const span = document.createElement("div");
    span.textContent = "Aucune conclusion structurée trouvée pour cette mission.";
    diagsContainer.appendChild(span);
  } else {
    entries.forEach(([label, text]) => {
      const card = document.createElement("div");
      card.className = "diag-card";

      const header = document.createElement("div");
      header.className = "diag-header";

      const headerLeft = document.createElement("div");
      headerLeft.className = "diag-header-left";

      const titleSpan = document.createElement("span");
      titleSpan.className = "diag-title";
      titleSpan.textContent = label;

      headerLeft.appendChild(titleSpan);
      header.appendChild(headerLeft);

      const body = document.createElement("div");
      body.className = "diag-body";
      body.textContent = text;

      card.appendChild(header);
      card.appendChild(body);

      diagsContainer.appendChild(card);
    });
  }

  // (➜ Bloc 4 récapitulatif global supprimé, on ne fait plus rien ici)


// ----------------------------------------
// Exports & JSON
// ----------------------------------------
function exportFilteredAsCSV() {
  if (!filteredMissions.length) {
    alert("Aucune mission filtrée à exporter.");
    return;
  }
  const lines = ["num_dossier"];
  filteredMissions.forEach((m) => {
    const num = (m.numDossier || "").toString().replace(/"/g, '""');
    lines.push(`"${num}"`);
  });
  const csvContent = lines.join("\r\n");
  downloadTextFile(csvContent, "missions_filtrees.csv", "text/csv");
}

async function copyFilteredToClipboard() {
  if (!filteredMissions.length) {
    alert("Aucune mission filtrée à copier.");
    return;
  }
  const text = filteredMissions.map((m) => m.numDossier || "").join("\n");
  try {
    await navigator.clipboard.writeText(text);
    alert("Liste des numéros copiée dans le presse-papier.");
  } catch (err) {
    console.error("Erreur lors de la copie dans le presse-papier", err);
    alert("Impossible de copier dans le presse-papier dans ce navigateur.");
  }
}

function exportAllAsJSON() {
  if (!allMissions.length) {
    alert("Aucune mission à exporter.");
    return;
  }
  const payload = {
    version: 1,
    exportedAt: new Date().toISOString(),
    missions: allMissions.map((m) => ({
      ...m,
      photoUrl: null,
      dpe: { energieUrl: null, co2Url: null }
    }))
  };
  const jsonStr = JSON.stringify(payload, null, 2);
  downloadTextFile(jsonStr, "missions_export.json", "application/json");
}

// Export d'une seule mission (fiche)
function exportSingleMissionAsJSON(mission) {
  const payload = {
    version: 1,
    exportedAt: new Date().toISOString(),
    mission: {
      ...mission,
      photoUrl: null,
      dpe: { energieUrl: null, co2Url: null }
    }
  };
  const fileName = `mission_${mission.numDossier || "fiche"}.json`;
  const jsonStr = JSON.stringify(payload, null, 2);
  downloadTextFile(jsonStr, fileName, "application/json");
}

function downloadTextFile(content, fileName, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function onImportJSON(event) {
  const file = event.target.files && event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const text = e.target.result;
      const data = JSON.parse(text);
      if (Array.isArray(data)) {
        allMissions = data;
      } else if (data && Array.isArray(data.missions)) {
        allMissions = data.missions;
      } else if (data && data.mission) {
        // Import d'une fiche seule -> on l'ajoute
        allMissions.push(data.mission);
      } else {
        throw new Error("Format JSON inattendu");
      }

      // On nettoie les éventuels blobUrl
      allMissions.forEach((m) => {
        if (m.photoUrl) m.photoUrl = null;
        if (m.dpe) {
          m.dpe.energieUrl = null;
          m.dpe.co2Url = null;
        }
      });

      filteredMissions = [...allMissions];
      populateFilterOptions();
      renderTable();
      updateStats();
      updateExportButtonsState();
      updateProgress(allMissions.length, allMissions.length, "Base JSON chargée (sans rescanner les dossiers).");
      if (allMissions.length > 0) {
        $("#filtersSection").classList.remove("hidden-block");
      }
    } catch (err) {
      console.error("Erreur de lecture du JSON", err);
      alert("Erreur lors de la lecture du fichier JSON.");
    } finally {
      event.target.value = "";
    }
  };
  reader.readAsText(file, "utf-8");
}
