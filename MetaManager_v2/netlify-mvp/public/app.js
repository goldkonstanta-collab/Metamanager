const tabKp = document.getElementById("tab-kp");
const tabContract = document.getElementById("tab-contract");
const kpSection = document.getElementById("kp-section");
const contractSection = document.getElementById("contract-section");
const statusEl = document.getElementById("status");

const kpForm = document.getElementById("kp-form");
const contractForm = document.getElementById("contract-form");

const smrTypeSelect = kpForm.elements.smrType;
const noSmrFields = document.getElementById("no-smr-fields");
const smrFields = document.getElementById("smr-fields");
const pirCheckbox = kpForm.elements.includePir;
const pirFields = document.getElementById("pir-fields");
const workAddressCheckbox = contractForm.elements.includeWorkAddress;
const workAddressWrap = document.getElementById("work-address-wrap");
const wellsDepthInput = kpForm.elements.wellsDepth;
const wellsPricePerMeterInput = kpForm.elements.wellsPricePerMeter;
const wellsPriceInput = kpForm.elements.wellsPrice;

const contractInnInput = contractForm.elements.customerInn;
const contractBikInput = contractForm.elements.customerBik;
const contractKpFileInput = contractForm.elements.kpFile;
const contractInnStatus = document.getElementById("contract-inn-status");
const contractBikStatus = document.getElementById("contract-bik-status");
const contractKpStatus = document.getElementById("contract-kp-status");
const kpUploadDrop = document.getElementById("kp-upload-drop");

const telegramChatIdInput = document.getElementById("telegram-chat-id");
const telegramSaveBtn = document.getElementById("telegram-save");
const telegramClearBtn = document.getElementById("telegram-clear");
const telegramStatusEl = document.getElementById("telegram-status");

const TELEGRAM_CHAT_ID_KEY = "metaManager.telegramChatId";

function getStoredTelegramChatId() {
  try {
    return (localStorage.getItem(TELEGRAM_CHAT_ID_KEY) || "").trim();
  } catch (e) {
    return "";
  }
}

function setTelegramStatusText(text, tone) {
  if (!telegramStatusEl) return;
  telegramStatusEl.textContent = text || "";
  telegramStatusEl.dataset.tone = tone || "";
}

function initTelegramChatId() {
  if (!telegramChatIdInput) return;
  const stored = getStoredTelegramChatId();
  if (stored) {
    telegramChatIdInput.value = stored;
    setTelegramStatusText(`Сохранено: ${stored}`, "ok");
  } else {
    setTelegramStatusText("Ключ не задан — документы уйдут в общий чат бота.");
  }

  telegramSaveBtn?.addEventListener("click", () => {
    const value = (telegramChatIdInput.value || "").trim();
    if (!value) {
      setTelegramStatusText("Введите chat ID перед сохранением.", "err");
      return;
    }
    if (!/^-?\d+$/.test(value)) {
      setTelegramStatusText("Chat ID должен состоять только из цифр.", "err");
      return;
    }
    try {
      localStorage.setItem(TELEGRAM_CHAT_ID_KEY, value);
      setTelegramStatusText(`Сохранено: ${value}`, "ok");
    } catch (e) {
      setTelegramStatusText("Не удалось сохранить в этом браузере.", "err");
    }
  });

  telegramClearBtn?.addEventListener("click", () => {
    try {
      localStorage.removeItem(TELEGRAM_CHAT_ID_KEY);
    } catch (e) {}
    telegramChatIdInput.value = "";
    setTelegramStatusText("Ключ сброшен.");
  });
}

function currentTelegramChatId() {
  const typed = (telegramChatIdInput?.value || "").trim();
  return typed || getStoredTelegramChatId();
}

let innLookupTimer = null;
let bikLookupTimer = null;
let innLookupToken = 0;
let bikLookupToken = 0;
let directBackendBase = "";

async function initBackendBase() {
  try {
    const response = await fetch("/api/backend-url", { method: "GET" });
    if (!response.ok) {
      return;
    }
    const data = await response.json();
    const candidate = typeof data?.backendUrl === "string" ? data.backendUrl.trim() : "";
    if (candidate) {
      directBackendBase = candidate.replace(/\/+$/, "");
    }
  } catch (e) {
    // fallback to Netlify proxy routes
  }
}

initBackendBase();
initTelegramChatId();

function setStatus(obj) {
  statusEl.textContent = typeof obj === "string" ? obj : JSON.stringify(obj, null, 2);
}

function showKp() {
  tabKp.classList.add("active");
  tabContract.classList.remove("active");
  kpSection.classList.remove("hidden");
  contractSection.classList.add("hidden");
}

function showContract() {
  tabContract.classList.add("active");
  tabKp.classList.remove("active");
  contractSection.classList.remove("hidden");
  kpSection.classList.add("hidden");
}

tabKp.addEventListener("click", showKp);
tabContract.addEventListener("click", showContract);

function syncSmrFields() {
  const isSmr = smrTypeSelect.value === "с смр";
  smrFields.classList.toggle("hidden", !isSmr);
  noSmrFields.classList.toggle("hidden", isSmr);
}

smrTypeSelect.addEventListener("change", syncSmrFields);
syncSmrFields();

pirCheckbox.addEventListener("change", () => {
  pirFields.classList.toggle("hidden", !pirCheckbox.checked);
});

workAddressCheckbox.addEventListener("change", () => {
  workAddressWrap.classList.toggle("hidden", !workAddressCheckbox.checked);
});

function recalcWellPrice() {
  const depth = Number(wellsDepthInput.value || 0);
  const pricePerMeter = Number(wellsPricePerMeterInput.value || 0);
  if (!depth || !pricePerMeter) {
    return;
  }
  wellsPriceInput.value = String(Math.round(depth * pricePerMeter));
}

wellsDepthInput.addEventListener("input", recalcWellPrice);
wellsPricePerMeterInput.addEventListener("input", recalcWellPrice);

function httpErrorMessage(data, fallback) {
  if (!data) {
    return fallback;
  }
  if (typeof data.detail === "string") {
    return data.detail;
  }
  if (Array.isArray(data.detail) && data.detail.length) {
    const first = data.detail[0];
    if (first && typeof first.msg === "string") {
      return first.msg;
    }
  }
  return data.error || data.message || fallback;
}

function resolveApiUrl(directPath, proxyPath) {
  if (directBackendBase) {
    return `${directBackendBase}${directPath}`;
  }
  return proxyPath;
}

function updateKpFileStatus() {
  const file = contractKpFileInput.files && contractKpFileInput.files[0];
  if (!file) {
    contractKpStatus.textContent = "";
    return;
  }
  contractKpStatus.textContent = `Файл КП: ${file.name}`;
}

function applyDroppedKpFile(fileList) {
  if (!fileList || !fileList.length) {
    return;
  }
  const file = fileList[0];
  if (!file.name.toLowerCase().endsWith(".docx")) {
    contractKpStatus.textContent = "Ошибка: загрузите файл Word .docx";
    return;
  }
  try {
    const dt = new DataTransfer();
    dt.items.add(file);
    contractKpFileInput.files = dt.files;
    updateKpFileStatus();
  } catch (e) {
    contractKpStatus.textContent = "Перетаскивание не поддерживается. Нажмите поле выбора файла.";
  }
}

contractKpFileInput.addEventListener("change", updateKpFileStatus);

if (kpUploadDrop) {
  kpUploadDrop.addEventListener("click", () => contractKpFileInput.click());
  kpUploadDrop.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      contractKpFileInput.click();
    }
  });
  kpUploadDrop.addEventListener("dragover", (e) => {
    e.preventDefault();
    kpUploadDrop.classList.add("active");
  });
  kpUploadDrop.addEventListener("dragleave", () => {
    kpUploadDrop.classList.remove("active");
  });
  kpUploadDrop.addEventListener("drop", (e) => {
    e.preventDefault();
    kpUploadDrop.classList.remove("active");
    applyDroppedKpFile(e.dataTransfer.files);
  });
}

async function getJSON(url) {
  const response = await fetch(url, { method: "GET" });
  let data = null;
  try {
    data = await response.json();
  } catch (e) {
    data = null;
  }
  if (!response.ok) {
    throw new Error(httpErrorMessage(data, `Request failed (${response.status})`));
  }
  return data;
}

function applyCompanyFields(company) {
  if (!company) {
    return;
  }
  const mapping = [
    ["customerFullname", company.customerFullname],
    ["customerShortname", company.customerShortname],
    ["customerAddress", company.customerAddress],
    ["customerOgrn", company.customerOgrn],
    ["customerInn", company.customerInn],
    ["customerKpp", company.customerKpp],
    ["customerPhone", company.customerPhone],
    ["customerEmail", company.customerEmail],
    ["customerDirectorTitle", company.customerDirectorTitle],
    ["customerDirectorName", company.customerDirectorName],
    ["customerBasis", company.customerBasis]
  ];
  for (const [name, value] of mapping) {
    if (value === undefined || value === null) {
      continue;
    }
    const el = contractForm.elements[name];
    if (el) {
      el.value = String(value);
    }
  }
}

function scheduleCompanyLookup() {
  if (innLookupTimer) {
    clearTimeout(innLookupTimer);
  }
  innLookupTimer = setTimeout(() => {
    innLookupTimer = null;
    maybeLookupCompanyByInn();
  }, 450);
}

async function maybeLookupCompanyByInn() {
  const raw = contractInnInput.value.trim();
  const digits = raw.replace(/\D/g, "");
  contractInnInput.value = digits;

  if (digits.length !== 10 && digits.length !== 12) {
    contractInnStatus.textContent = "";
    return;
  }

  const token = ++innLookupToken;
  contractInnStatus.textContent = "Загрузка данных по ИНН...";

  try {
    const result = await getJSON(
      resolveApiUrl(
        `/lookup/company?inn=${encodeURIComponent(digits)}`,
        `/api/lookup/company?inn=${encodeURIComponent(digits)}`
      )
    );
    if (token !== innLookupToken) {
      return;
    }
    applyCompanyFields(result.company);
    contractInnStatus.textContent = result.company?.customerShortname
      ? `Данные загружены: ${result.company.customerShortname}`
      : "Данные загружены";
  } catch (err) {
    if (token !== innLookupToken) {
      return;
    }
    contractInnStatus.textContent = `Ошибка: ${err.message}`;
  }
}

function scheduleBankLookup() {
  if (bikLookupTimer) {
    clearTimeout(bikLookupTimer);
  }
  bikLookupTimer = setTimeout(() => {
    bikLookupTimer = null;
    maybeLookupBankByBik();
  }, 350);
}

async function maybeLookupBankByBik() {
  const raw = contractBikInput.value.trim();
  const digits = raw.replace(/\D/g, "");
  contractBikInput.value = digits;

  if (digits.length !== 9) {
    contractBikStatus.textContent = "";
    return;
  }

  const token = ++bikLookupToken;
  contractBikStatus.textContent = "Поиск банка по БИК...";

  try {
    const result = await getJSON(
      resolveApiUrl(
        `/lookup/bank?bic=${encodeURIComponent(digits)}`,
        `/api/lookup/bank?bic=${encodeURIComponent(digits)}`
      )
    );
    if (token !== bikLookupToken) {
      return;
    }
    if (result.bank?.customerBank) {
      contractForm.elements.customerBank.value = String(result.bank.customerBank);
    }
    if (result.bank?.customerKs && !contractForm.elements.customerKs.value.trim()) {
      contractForm.elements.customerKs.value = String(result.bank.customerKs);
    }
    contractBikStatus.textContent = result.bank?.customerBank
      ? `✓ ${String(result.bank.customerBank).slice(0, 48)}`
      : "Банк найден";
  } catch (err) {
    if (token !== bikLookupToken) {
      return;
    }
    contractBikStatus.textContent = `Ошибка: ${err.message}`;
  }
}

contractInnInput.addEventListener("input", scheduleCompanyLookup);
contractBikInput.addEventListener("input", scheduleBankLookup);

async function postJSON(url, payload) {
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });
  let data = null;
  try {
    data = await response.json();
  } catch (e) {
    data = null;
  }
  if (!response.ok) {
    throw new Error(httpErrorMessage(data, `Request failed (${response.status})`));
  }
  return data;
}

async function postFormData(url, formData) {
  const response = await fetch(url, {
    method: "POST",
    body: formData
  });
  let data = null;
  try {
    data = await response.json();
  } catch (e) {
    data = null;
  }
  if (!response.ok) {
    throw new Error(httpErrorMessage(data, `Request failed (${response.status})`));
  }
  return data;
}

kpForm.addEventListener("submit", async (e) => {
  e.preventDefault();
  const payload = {
    kpName: kpForm.elements.kpName.value.trim(),
    kpTitle: kpForm.elements.kpTitle.value.trim(),
    kpManagerName: kpForm.elements.kpManagerName.value.trim(),
    branch: kpForm.elements.branch.value,
    volume: kpForm.elements.volume.value,
    smrType: kpForm.elements.smrType.value,
    wellsCount: Number(kpForm.elements.wellsCount.value || 1),
    includeWells: kpForm.elements.includeWells.checked,
    includePump: kpForm.elements.includePump.checked,
    includeBmz: kpForm.elements.includeBmz.checked,
    wellsCountSmr: Number(kpForm.elements.wellsCountSmr.value || 1),
    wellsDesign: kpForm.elements.wellsDesign.value.trim(),
    wellsDepth: kpForm.elements.wellsDepth.value.trim(),
    wellsPricePerMeter: kpForm.elements.wellsPricePerMeter.value.trim(),
    wellsPrice: kpForm.elements.wellsPrice.value.trim(),
    pumpPrice: kpForm.elements.pumpPrice.value.trim(),
    bmzSize: kpForm.elements.bmzSize.value.trim(),
    bmzPrice: kpForm.elements.bmzPrice.value.trim(),
    includePir: kpForm.elements.includePir.checked,
    pirCount: Number(kpForm.elements.pirCount.value || 1),
    pirPrice: Number(kpForm.elements.pirPrice.value || 0),
    telegramChatId: currentTelegramChatId()
  };

  setStatus("Отправка КП...");
  try {
    const result = await postJSON(resolveApiUrl("/generate/kp", "/api/kp"), payload);
    setStatus(result);
  } catch (err) {
    setStatus({ error: err.message });
  }
});

contractForm.addEventListener("submit", async (e) => {
  e.preventDefault();
  const payload = {
    contractNumber: contractForm.elements.contractNumber.value.trim(),
    customerFullname: contractForm.elements.customerFullname.value.trim(),
    customerShortname: contractForm.elements.customerShortname.value.trim(),
    customerAddress: contractForm.elements.customerAddress.value.trim(),
    customerOgrn: contractForm.elements.customerOgrn.value.trim(),
    customerInn: contractForm.elements.customerInn.value.trim(),
    customerKpp: contractForm.elements.customerKpp.value.trim(),
    customerBik: contractForm.elements.customerBik.value.trim(),
    customerBank: contractForm.elements.customerBank.value.trim(),
    customerRs: contractForm.elements.customerRs.value.trim(),
    customerKs: contractForm.elements.customerKs.value.trim(),
    customerPhone: contractForm.elements.customerPhone.value.trim(),
    customerEmail: contractForm.elements.customerEmail.value.trim(),
    advancePercent: contractForm.elements.advancePercent.value.trim(),
    customerDirectorTitle: contractForm.elements.customerDirectorTitle.value.trim(),
    customerDirectorName: contractForm.elements.customerDirectorName.value.trim(),
    customerBasis: contractForm.elements.customerBasis.value.trim(),
    includeWorkAddress: contractForm.elements.includeWorkAddress.checked,
    workAddress: contractForm.elements.workAddress.value.trim(),
    telegramChatId: currentTelegramChatId()
  };

  setStatus("Отправка договора...");
  try {
    const kpFile = contractKpFileInput.files && contractKpFileInput.files[0];
    const endpoint = resolveApiUrl("/generate/contract", "/api/contract");
    let result = null;

    if (kpFile) {
      const formData = new FormData();
      for (const [key, value] of Object.entries(payload)) {
        formData.append(key, value);
      }
      formData.append("kpFile", kpFile, kpFile.name);
      result = await postFormData(endpoint, formData);
    } else {
      result = await postJSON(endpoint, payload);
    }
    setStatus(result);
  } catch (err) {
    setStatus({ error: err.message });
  }
});
