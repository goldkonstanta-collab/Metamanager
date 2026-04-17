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
const contractInnStatus = document.getElementById("contract-inn-status");
const contractBikStatus = document.getElementById("contract-bik-status");

let innLookupTimer = null;
let bikLookupTimer = null;
let innLookupToken = 0;
let bikLookupToken = 0;

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
    const result = await getJSON(`/api/lookup/company?inn=${encodeURIComponent(digits)}`);
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
    const result = await getJSON(`/api/lookup/bank?bic=${encodeURIComponent(digits)}`);
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

kpForm.addEventListener("submit", async (e) => {
  e.preventDefault();
  const payload = {
    kpName: kpForm.elements.kpName.value.trim(),
    kpTitle: kpForm.elements.kpTitle.value.trim(),
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
    pirPrice: Number(kpForm.elements.pirPrice.value || 0)
  };

  setStatus("Отправка КП...");
  try {
    const result = await postJSON("/api/kp", payload);
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
    workAddress: contractForm.elements.workAddress.value.trim()
  };

  setStatus("Отправка договора...");
  try {
    const result = await postJSON("/api/contract", payload);
    setStatus(result);
  } catch (err) {
    setStatus({ error: err.message });
  }
});
