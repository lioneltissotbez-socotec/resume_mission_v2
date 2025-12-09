// Dashboard Missions – Synthèse Liciel (version PRO fiches missions)
// ----------------------------------------------------
// - Scan de dossiers LICIEL via File System Access API
// - Lecture des XML Table_General_Bien / Table_Z_Conclusions
// - Photo principale : /photos/presentation.jpg
// - Étiquettes DPE : /images/DPE_2020_Etiquette_Energie.jpg & CO2.jpg
// - Filtres avancés DO / proprio / opérateur / type / conclusion / DPE
// - Export CSV + JSON
// - Encodage Windows-1252 / UTF-8 auto-détecté
// - Fiches missions (admin + DPE + conclusions structurées)
// - Bloc récapitulatif global supprimé (ta demande)
// ----------------------------------------------------

let rootDirHandle = null;
let allMissions = [];
let filteredMissions = [];
let isScanning = false;
let currentFicheMission = null;

// Types de missions indexés sur LiColonne_Mission_Missions_programmees
const MISSION_TYPES = [
  "Amiante (DTA)", "Amiante (Vente)", "Amiante (Travaux)", "Amiante (Démolition)",
  "Diagnostic Termites", "Diagnostic Parasites", "Métrage (Carrez)", "CREP",
  "Assainissement", "Piscine", "Gaz", "Électricité", "Diagnostic Technique Global (DTG)",
  "DPE", "Prêt à taux zéro", "ERP / ESRIS", "État d’Habitabilité", "État des lieux",
  "Plomb dans l’eau", "Ascenseur", "Radon", "Diagnostic Incendie",
  "Accessibilité Handicapé", "Mesurage (Boutin)", "Amiante (DAPP)", "DRIPP",
  "Performance Numérique", "Infiltrométrie", "Amiante (Avant Travaux)",
  "Gestion Déchets / PEMD", "Plomb (Après Travaux)", "Amiante (Contrôle périodique)",
  "Empoussièrement", "Module Interne", "Home Inspection", "Home Inspection 4PT",
  "Wind Mitigation", "Plomb (Avant Travaux)", "Amiante (HAP)", "[Non utilisé]", "DPEG"
];

// Mapping des champs de Table_Z_Conclusions vers diagnostics lisibles
const CONCLUSION_FIELDS = [
  { tag: "LiColonne_Variable_resume_conclusion_termites",        label: "Termites" },
  { tag: "LiColonne_Variable_resume_conclusion_autres_parasites",label: "Parasites" },
  { tag: "LiColonne_Variable_resume_conclusion_amiante",         label: "Amiante" },
  { tag: "LiColonne_Variable_resume_conclusion_carrez",          label: "Carrez" },
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

// Helpers
const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => Array.from(document.querySelectorAll(sel));

function debounce(fn, delay = 150) {
  let timer;
  return (...a) => { clearTimeout(timer); timer = setTimeout(() => fn(...a), delay); };
}

// -----------------------------------------------
// INITIALISATION UI
// -----------------------------------------------
document.addEventListener("DOMContentLoaded", () => initUI());

function initUI() {
  const btnPickRoot = $("#btnPickRoot");
  const btnScan = $("#btnScan");

  const hasFsAccess = typeof window.showDirectoryPicker === "function";
  if (!hasFsAccess) {
    alert("⚠️ Ce navigateur ne supporte pas showDirectoryPicker (nécessaire). Utilisez Chrome/Edge récents.");
    btnPickRoot.disabled = true;
  }

  btnPickRoot.addEventListener("click", onPickRoot);
  btnScan.addEventListener("click", onScan);

  // Filtres
  const debouncedApply = debounce(applyFilters, 200);
  ["#filterDO", "#filterProp", "#filterOp", "#filterType"].forEach(sel => {
    $(sel).addEventListener("change", debouncedApply);
  });
  $("#filterConclusion").addEventListener("input", debouncedApply);

  // Export
  $("#btnExportCSV").addEventListener("click", exportFilteredAsCSV);
  $("#btnCopyClipboard").addEventListener("click", copyFilteredToClipboard);
  $("#btnExportJSON").addEventListener("click", exportAllAsJSON);
  $("#btnImportJSON").addEventListener("click", () => $("#jsonFileInput").click());
  $("#jsonFileInput").addEventListener("change", onImportJSON);

  // Modales
  $("#btnCloseModal").addEventListener("click", closeConclusionModal);
  $("#modalOverlay").addEventListener("click", e => { if (e.target.id === "modalOverlay") closeConclusionModal(); });

  $("#btnClosePhoto").addEventListener("click", closePhotoModal);
  $("#photoOverlay").addEventListener("click", e => { if (e.target.id === "photoOverlay") closePhotoModal(); });

  $("#btnExportFiche").addEventListener("click", () => {
    if (!currentFicheMission) return alert("Aucune fiche mission active.");
    exportSingleMissionAsJSON(currentFicheMission);
  });

  updateProgress(0, 0, "En attente…");
  updateStats();
  updateExportButtonsState();
}

// -----------------------------------------------
// PICK DOSSIER RACINE
// -----------------------------------------------
async function onPickRoot() {
  try {
    rootDirHandle = await window.showDirectoryPicker();
    $("#rootInfo").textContent = "Dossier racine : " + rootDirHandle.name;
    $("#btnScan").disabled = false;
  } catch (e) {
    console.warn("Sélection annulée :", e);
  }
}

// -----------------------------------------------
// SCAN GLOBAL
// -----------------------------------------------
async function onScan() {
  if (!rootDirHandle) return alert("Choisissez d'abord un dossier racine.");
  if (isScanning) return;

  const mode = document.querySelector("input[name='scanMode']:checked")?.value || "all";
  const prefix = $("#inputPrefix").value.trim();
  const list = $("#inputDossierList").value.trim().split(/[\s,;]+/).filter(s => s);

  if (mode === "prefix" && !prefix) return alert("Saisir un préfixe.");
  if (mode === "list" && list.length === 0) return alert("Coller au moins un n° de dossier.");

  allMissions = [];
  filteredMissions = [];
  renderTable();
  updateStats();
  updateExportButtonsState();

  setScanningState(true);
  updateProgress(0, 0, "Lecture des dossiers…");

  try {
    const entries = [];
    for await (const [name, handle] of rootDirHandle.entries()) {
      if (handle.kind === "directory") entries.push({ name, handle });
    }

    const candidates = entries.filter(e => {
      if (mode === "all") return true;
      if (mode === "prefix") return e.name.startsWith(prefix);
      if (mode === "list") return list.some(pref => e.name.startsWith(pref));
      return true;
    });

    if (candidates.length === 0) {
      updateProgress(0, 0, "Aucun dossier trouvé.");
      return;
    }

    let i = 0;
    for (const { name, handle } of candidates) {
      try {
        const mission = await processMissionFolder(handle, name);
        if (mission) allMissions.push(mission);
      } catch (err) {
        console.error("Erreur dossier", name, err);
      }
      i++;
      updateProgress(i, candidates.length, `Scan : ${i}/${candidates.length}`);
    }

    filteredMissions = [...allMissions];
    populateFilterOptions();
    renderTable();
    updateStats();
    updateExportButtonsState();
    updateProgress(i, i, `Scan terminé : ${allMissions.length} missions.`);
    if (allMissions.length) $("#filtersSection").classList.remove("hidden-block");

  } catch (err) {
    console.error(err);
    alert("Erreur lors du scan.");
  }

  setScanningState(false);
}

function setScanningState(state) {
  isScanning = state;
  $("#btnScan").disabled = state;
  $("#btnPickRoot").disabled = state;
}

// -----------------------------------------------
// LECTURE XML + DÉCODAGE ENCODAGE
// -----------------------------------------------
async function readXmlFile(dirHandle, fileName) {
  try {
    const h = await dirHandle.getFileHandle(fileName);
    const f = await h.getFile();
    const buffer = await f.arrayBuffer();

    let textUtf8 = new TextDecoder("utf-8", { fatal: false }).decode(buffer);
    let text = textUtf8;

    const m = textUtf8.match(/encoding="([^"]+)"/i);
    const declared = m ? m[1].toLowerCase() : null;

    if (declared && declared !== "utf-8") {
      try { text = new TextDecoder(declared).decode(buffer); }
      catch { text = new TextDecoder("windows-1252").decode(buffer); }
    } else {
      const bad = (textUtf8.match(/�/g) || []).length;
      if (bad >= 3) {
        try {
          const t1252 = new TextDecoder("windows-1252").decode(buffer);
          if ((t1252.match(/�/g) || []).length < bad) text = t1252;
        } catch {}
      }
    }

    const doc = new DOMParser().parseFromString(text, "application/xml");
    if (doc.querySelector("parsererror")) return null;
    return doc;

  } catch (err) {
    console.warn("XML manquant:", fileName, err);
    return null;
  }
}

// -----------------------------------------------
// PARSE DPE (valeurs + classes + ADEME)
// -----------------------------------------------
function parseDpeInfo(str) {
  if (!str) return { conso:null, classeEner:null, co2:null, classeCO2:null, ademe:null };

  const t = str.replace(/\s+/g," ");
  const conso = t.match(/([0-9]+)\s*kWh/);
  const classeEner = t.match(/Classe\s*([A-G])/i);
  const co2 = t.match(/([0-9]+)\s*kg/i);
  const classeCO2 = t.match(/émissions.*classe\s*([A-G])/i);
  const ademe = t.match(/ADEME[:\s]*([A-Z0-9]+)/i);

  return {
    conso: conso ? parseInt(conso[1],10) : null,
    classeEner: classeEner ? classeEner[1] : null,
    co2: co2 ? parseInt(co2[1],10) : null,
    classeCO2: classeCO2 ? classeCO2[1] : null,
    ademe: ademe ? ademe[1] : null
  };
}

// -----------------------------------------------
// GET FILE BY RELATIVE PATH
// -----------------------------------------------
async function getFileFromRelativePath(root, rel) {
  const parts = rel.split("/").filter(p => p);
  let dir = root;
  for (let i=0;i<parts.length;i++) {
    const p = parts[i];
    const last = i === parts.length - 1;
    if (last) {
      const fh = await dir.getFileHandle(p);
      return fh.getFile();
    } else {
      dir = await dir.getDirectoryHandle(p);
    }
  }
}
// --------------------------------------------------------
// Traitement d'un dossier de mission
// --------------------------------------------------------
async function processMissionFolder(folderHandle, folderName) {

  // ---------- 1) Dossier XML obligatoire ----------
  let xmlDir;
  try {
    xmlDir = await folderHandle.getDirectoryHandle("XML");
  } catch (e) {
    console.warn("Pas de dossier XML dans", folderName);
    return null;
  }

  // ---------- 2) Lecture Table_General_Bien ----------
  const bienXml = await readXmlFile(xmlDir, "Table_General_Bien.xml");
  if (!bienXml) {
    console.warn("Table_General_Bien.xml manquant dans", folderName);
    return null;
  }

  const mission = {};

  // ----------- A) Informations administratives ----------
  mission.numDossier = getXmlValue(bienXml, "LiColonne_Mission_Num_Dossier") || folderName;

  mission.donneurOrdre = {
    entete: getXmlValue(bienXml, "LiColonne_DOrdre_Entete"),
    nom: getXmlValue(bienXml, "LiColonne_DOrdre_Nom"),
    adresse: getXmlValue(bienXml, "LiColonne_DOrdre_Adresse1"),
    commune: getXmlValue(bienXml, "LiColonne_DOrdre_Commune"),
    departement: getXmlValue(bienXml, "LiColonne_DOrdre_Departement")
  };

  mission.proprietaire = {
    entete: getXmlValue(bienXml, "LiColonne_Prop_Entete"),
    nom: getXmlValue(bienXml, "LiColonne_Prop_Nom"),
    adresse: getXmlValue(bienXml, "LiColonne_Prop_Adresse1"),
    commune: getXmlValue(bienXml, "LiColonne_Prop_Commune"),
    departement: getXmlValue(bienXml, "LiColonne_Prop_Departement")
  };

  mission.immeuble = {
    adresse: getXmlValue(bienXml, "LiColonne_Immeuble_Adresse1"),
    commune: getXmlValue(bienXml, "LiColonne_Immeuble_Commune"),
    departement: getXmlValue(bienXml, "LiColonne_Immeuble_Departement"),
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

  // ---------- 3) Lecture Table_Z_Conclusions (NOUVELLE BASE UNIQUE) ----------
  const conclXml = await readXmlFile(xmlDir, "Table_Z_Conclusions.xml");

  mission.conclusionsList = [];
  mission.dpeInfo = null;

  if (conclXml) {
    CONCLUSION_FIELDS.forEach(({ tag, label }) => {
      const val = getXmlValue(conclXml, tag);
      if (val && val.trim() !== "") {

        // Extraction DPE si c'est le bon champ
        if (label === "DPE") {
          const parsed = parseDpeInfo(val);
          mission.dpeInfo = parsed;
        }

        mission.conclusionsList.push({
          type: label,
          text: val.trim()
        });
      }
    });
  }

  // ---------- 4) Photo présentation ----------
  mission.photoUrl = null;
  try {
    const fp = await getFileFromRelativePath(folderHandle, "photos/presentation.jpg");
    mission.photoUrl = URL.createObjectURL(fp);
  } catch (e) {}

  // ---------- 5) Images DPE si mission contient DPE ----------
  mission.dpe = { energieUrl: null, co2Url: null, ademeLink: null };

  if (mission.mission.missionsEffectuees.includes("DPE")) {

    // Énergie
    try {
      const f = await getFileFromRelativePath(folderHandle, "images/DPE_2020_Etiquette_Energie.jpg");
      mission.dpe.energieUrl = URL.createObjectURL(f);
    } catch (e) {}

    // CO2
    try {
      const f = await getFileFromRelativePath(folderHandle, "images/DPE_2020_Etiquette_CO2.jpg");
      mission.dpe.co2Url = URL.createObjectURL(f);
    } catch (e) {}

    // Lien ADEME
    if (mission.dpeInfo && mission.dpeInfo.ademe) {
      mission.dpe.ademeLink = "https://observatoire-dpe-audit.ademe.fr/afficher-dpe/" + mission.dpeInfo.ademe;
    }
  }

  // ---------- 6) Normalisation pour filtres avancés ----------
  mission._norm = {
    conclusion: mission.conclusionsList.map(c => c.text.toLowerCase()).join(" "),
    dpeClasse: mission.dpeInfo?.classeEner || "",
    dpeConso: mission.dpeInfo?.conso || null,
    dpeCO2: mission.dpeInfo?.co2 || null
  };

  return mission;
}
// --------------------------------------------------------
// Rendu tableau principal
// --------------------------------------------------------
function renderTable() {
  const tbody = $("#resultsTable tbody");
  tbody.innerHTML = "";

  for (const m of filteredMissions) {
    const tr = document.createElement("tr");

    // 1) Num dossier
    tr.appendChild(tdText(m.numDossier));

    // 2) Donneur d'ordre
    tr.appendChild(tdText(formatDonneurOrdre(m)));

    // 3) Propriétaire
    tr.appendChild(tdText(formatProprietaire(m)));

    // 4) Adresse immeuble
    const im = m.immeuble || {};
    tr.appendChild(tdText([im.adresse, im.commune, im.departement].filter(Boolean).join(" ")));

    // 5) Type/nature bien
    tr.appendChild(
      tdText([im.typeBien, im.natureBien, im.typeDossier].filter(Boolean).join(" / "))
    );

    // 6) Dates
    const lines = [];
    if (m.mission.dateVisite) lines.push("Visite : " + m.mission.dateVisite);
    if (m.mission.dateRapport) lines.push("Rapport : " + m.mission.dateRapport);
    tr.appendChild(tdText(lines.join("\n")));

    // 7) Opérateur
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

    // 8) Missions effectuées → tags
    const tdTags = document.createElement("td");
    (m.mission.missionsEffectuees || []).forEach((type) => {
      const span = document.createElement("span");
      span.className = "tag";
      span.textContent = type;
      tdTags.appendChild(span);
    });
    tr.appendChild(tdTags);

    // 9) Bouton fiche
    const tdFiche = document.createElement("td");
    const btn = document.createElement("button");
    btn.className = "btn-link";
    btn.textContent = "Ouvrir";
    btn.addEventListener("click", () => openFicheModal(m));
    tdFiche.appendChild(btn);
    tr.appendChild(tdFiche);

    // 10) Photo présentation
    const tdPhoto = document.createElement("td");
    if (m.photoUrl) {
      const img = document.createElement("img");
      img.src = m.photoUrl;
      img.className = "photo-thumb";
      img.addEventListener("click", () => openPhotoModal(m.photoUrl));
      tdPhoto.appendChild(img);
    } else {
      tdPhoto.textContent = "—";
    }
    tr.appendChild(tdPhoto);

    tbody.appendChild(tr);
  }
}

function tdText(txt) {
  const td = document.createElement("td");
  td.textContent = txt || "";
  return td;
}

// --------------------------------------------------------
// FICHE MISSION (modale)
// --------------------------------------------------------
function openFicheModal(m) {
  currentFicheMission = m;

  $("#ficheTitle").textContent = "Mission " + (m.numDossier || "");
  $("#ficheSubtitle").textContent =
    [m.immeuble.adresse, m.immeuble.commune, m.immeuble.departement]
      .filter(Boolean)
      .join(" • ");

  // ------- Bloc ADMIN -------
  $("#ficheDO").textContent = buildBlocTexte(m.donneurOrdre, [
    ["Entité", "entete"],
    ["Nom", "nom"],
    ["Adresse", "adresse"],
    ["Commune", "commune"],
    ["Département", "departement"]
  ]);

  $("#ficheProp").textContent = buildBlocTexte(m.proprietaire, [
    ["Entité", "entete"],
    ["Nom", "nom"],
    ["Adresse", "adresse"],
    ["Commune", "commune"],
    ["Département", "departement"]
  ]);

  $("#ficheImmeuble").textContent = buildBlocTexte(m.immeuble, [
    ["Adresse", "adresse"],
    ["Commune", "commune"],
    ["Département", "departement"],
    ["Lot", "lot"],
    ["Nature du bien", "natureBien"],
    ["Type de bien", "typeBien"],
    ["Type de dossier", "typeDossier"],
    ["Description", "description"]
  ]);

  $("#ficheMission").textContent = [
    m.mission.dateVisite ? "Date visite : " + m.mission.dateVisite : "",
    m.mission.dateRapport ? "Date rapport : " + m.mission.dateRapport : "",
    m.mission.missionsEffectuees.length
      ? "Diagnostics réalisés : " + m.mission.missionsEffectuees.join(", ")
      : ""
  ]
    .filter(Boolean)
    .join("\n");

  // ------- Bloc PHOTO -------
  const p = $("#fichePhotoContainer");
  p.innerHTML = "";
  if (m.photoUrl) {
    const img = document.createElement("img");
    img.src = m.photoUrl;
    img.addEventListener("click", () => openPhotoModal(m.photoUrl));
    p.appendChild(img);
  } else {
    p.textContent = "Aucune photo disponible";
  }

  // ------- Bloc DPE -------
  const d = $("#ficheDPEContainer");
  d.innerHTML = "";

  if (m.dpe.energieUrl || m.dpe.co2Url || m.dpeInfo) {
    if (m.dpe.energieUrl) {
      const imgE = document.createElement("img");
      imgE.src = m.dpe.energieUrl;
      imgE.className = "dpe-img";
      d.appendChild(imgE);
    }

    if (m.dpe.co2Url) {
      const imgC = document.createElement("img");
      imgC.src = m.dpe.co2Url;
      imgC.className = "dpe-img";
      d.appendChild(imgC);
    }

    if (m.dpeInfo) {
      const box = document.createElement("div");
      box.className = "dpe-info-box";

      box.innerHTML = `
        <div><strong>Consommation :</strong> ${m.dpeInfo.conso || "—"} kWh/m²/an (Classe ${m.dpeInfo.classeEner || "?"})</div>
        <div><strong>Émissions CO₂ :</strong> ${m.dpeInfo.co2 || "—"} kg/m²/an (Classe ${m.dpeInfo.classeCO2 || "?"})</div>
      `;

      if (m.dpe.ademeLink) {
        const link = document.createElement("a");
        link.href = m.dpe.ademeLink;
        link.target = "_blank";
        link.textContent = "Voir le DPE sur le site ADEME";
        link.className = "btn-ademe";
        box.appendChild(link);
      }

      d.appendChild(box);
    }
  } else {
    d.textContent = "DPE non réalisé.";
  }

  // ------- Bloc DIAGNOSTICS (pas de toggle, juste affichage) -------
  const container = $("#ficheDiagsContainer");
  container.innerHTML = "";

  m.conclusionsList.forEach((c) => {
    const card = document.createElement("div");
    card.className = "diag-card-open";

    const title = document.createElement("div");
    title.className = "diag-title-open";
    title.textContent = c.type;

    const content = document.createElement("div");
    content.className = "diag-content-open";
    content.textContent = c.text;

    card.appendChild(title);
    card.appendChild(content);
    container.appendChild(card);
  });

  $("#modalOverlay").classList.remove("hidden");
}

function closeConclusionModal() {
  $("#modalOverlay").classList.add("hidden");
  currentFicheMission = null;
}

// --------------------------------------------------------
// FILTRES AVANCÉS (ajout DPE)
// --------------------------------------------------------
function applyFilters() {
  const doValues = getSelectedOptions("#filterDO");
  const propValues = getSelectedOptions("#filterProp");
  const opValues = getSelectedOptions("#filterOp");
  const typeValues = getSelectedOptions("#filterType");

  const conclText = ($("#filterConclusion").value || "").trim().toLowerCase();

  // Filtres DPE — ces champs doivent être ajoutés dans ton HTML
  const dpeClass = ($("#filterDpeClass")?.value || "").trim().toUpperCase();
  const consoMin = parseInt($("#filterDpeConsoMin")?.value || "");
  const consoMax = parseInt($("#filterDpeConsoMax")?.value || "");
  const co2Min = parseInt($("#filterDpeCO2Min")?.value || "");
  const co2Max = parseInt($("#filterDpeCO2Max")?.value || "");

  filteredMissions = allMissions.filter((m) => {
    if (doValues.length && !doValues.includes(formatDonneurOrdre(m))) return false;
    if (propValues.length && !propValues.includes(formatProprietaire(m))) return false;
    if (opValues.length && !opValues.includes(formatOperateur(m))) return false;

    if (typeValues.length) {
      const eff = m.mission.missionsEffectuees || [];
      if (!eff.some((t) => typeValues.includes(t))) return false;
    }

    if (conclText) {
      const txt = m._norm.conclusion || "";
      if (!txt.includes(conclText)) return false;
    }

    // --- Filtres DPE ---
    const dpe = m._norm;

    if (dpeClass && dpe.dpeClasse !== dpeClass) return false;

    if (!isNaN(consoMin) && dpe.dpeConso !== null && dpe.dpeConso < consoMin) return false;
    if (!isNaN(consoMax) && dpe.dpeConso !== null && dpe.dpeConso > consoMax) return false;

    if (!isNaN(co2Min) && dpe.dpeCO2 !== null && dpe.dpeCO2 < co2Min) return false;
    if (!isNaN(co2Max) && dpe.dpeCO2 !== null && dpe.dpeCO2 > co2Max) return false;

    return true;
  });

  renderTable();
  updateStats();
  updateExportButtonsState();
}

// --------------------------------------------------------
// PHOTO MODALE
// --------------------------------------------------------
function openPhotoModal(url) {
  $("#modalPhoto").src = url;
  $("#photoOverlay").classList.remove("hidden");
}

function closePhotoModal() {
  $("#photoOverlay").classList.add("hidden");
  $("#modalPhoto").src = "";
}

// --------------------------------------------------------
// EXPORTS
// --------------------------------------------------------
function exportFilteredAsCSV() {
  if (!filteredMissions.length) {
    alert("Aucune mission filtrée.");
    return;
  }
  const lines = ["num_dossier"];
  filteredMissions.forEach((m) => lines.push(`"${m.numDossier}"`));
  downloadTextFile(lines.join("\r\n"), "missions_filtrees.csv", "text/csv");
}

async function copyFilteredToClipboard() {
  const text = filteredMissions.map((m) => m.numDossier).join("\n");
  await navigator.clipboard.writeText(text);
  alert("Numéros copiés.");
}

function exportAllAsJSON() {
  const payload = {
    version: 2,
    exportedAt: new Date().toISOString(),
    missions: allMissions.map((m) => ({
      ...m,
      photoUrl: null,
      dpe: { energieUrl: null, co2Url: null }
    }))
  };
  downloadTextFile(JSON.stringify(payload, null, 2), "missions_export.json", "application/json");
}

function exportSingleMissionAsJSON(m) {
  const payload = {
    version: 2,
    exportedAt: new Date().toISOString(),
    mission: {
      ...m,
      photoUrl: null,
      dpe: { energieUrl: null, co2Url: null }
    }
  };
  downloadTextFile(
    JSON.stringify(payload, null, 2),
    `mission_${m.numDossier}.json`,
    "application/json"
  );
}
