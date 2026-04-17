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
    const message =
      (data && (data.error || data.detail || data.message)) ||
      `Request failed (${response.status})`;
    throw new Error(message);
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
