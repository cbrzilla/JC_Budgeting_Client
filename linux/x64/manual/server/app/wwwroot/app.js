const currency = new Intl.NumberFormat(undefined, {
  style: "currency",
  currency: "USD"
});

const shortDate = new Intl.DateTimeFormat(undefined, {
  month: "2-digit",
  day: "2-digit",
  year: "numeric"
});

const statusBanner = document.getElementById("statusBanner");
const refreshButton = document.getElementById("refreshButton");
const periodSelect = document.getElementById("periodSelect");
const budgetItemsHost = document.getElementById("budgetItemsHost");
const summaryBreakdownChart = document.getElementById("summaryBreakdownChart");
const pageSections = Array.from(document.querySelectorAll("[data-page]"));
const navButtons = Array.from(document.querySelectorAll("[data-nav-page]"));

const state = {
  periods: [],
  selectedPeriodIndex: null,
  activePage: "summary",
  lastKnownChangeToken: "",
  suppressChangePromptUntil: 0,
  changePromptOpen: false
};

const CHANGE_TOKEN_POLL_INTERVAL_MS = 10000;

window.addEventListener("error", event => {
  setStatus(`Page error: ${event.message || "Unknown script error"}`, true);
});

window.addEventListener("unhandledrejection", event => {
  const message = event?.reason?.message || event?.reason || "Unknown async error";
  setStatus(`Load error: ${message}`, true);
});

function formatMoney(value) {
  return currency.format(Number(value || 0));
}

function formatDate(value) {
  if (!value) {
    return "N/A";
  }

  return shortDate.format(new Date(value));
}

function setStatus(text, isError = false) {
  statusBanner.textContent = text;
  statusBanner.classList.toggle("error", isError);
}

function setActivePage(pageName) {
  state.activePage = pageName;

  for (const section of pageSections) {
    section.classList.toggle("active", section.dataset.page === pageName);
  }

  for (const button of navButtons) {
    button.classList.toggle("active", button.dataset.navPage === pageName);
  }
}

function renderCards(targetId, items, kind, mapItem) {
  const host = document.getElementById(targetId);
  host.innerHTML = "";

  for (const item of items) {
    const data = mapItem(item);
    const card = document.createElement("article");
    card.className = `list-card ${kind} ${data.cardClass || ""}`.trim();
    card.innerHTML = `
      <div class="list-title">${data.title}</div>
      <div class="list-value ${data.valueClass || ""}">${data.value}</div>
      <div class="list-detail">${data.detail}</div>
      ${data.meta ? `<div class="list-meta">${data.meta}</div>` : ""}
    `;
    host.appendChild(card);
  }
}

async function fetchJson(url, options) {
  const response = await fetch(url, {
    cache: "no-store",
    ...options
  });

  if (!response.ok) {
    const error = await response.json().catch(() => ({ error: `Request failed: ${response.status}` }));
    throw new Error(error.error || `Request failed: ${response.status}`);
  }

  if (response.status === 204) {
    return null;
  }

  return response.json();
}

function getSelectedPeriod() {
  return state.periods.find(period => period.index === state.selectedPeriodIndex) || null;
}

function updatePeriodSummary(period) {
  return period;
}

function populatePeriodSelect() {
  periodSelect.innerHTML = "";

  for (const period of state.periods) {
    const option = document.createElement("option");
    option.value = String(period.index);
    option.textContent = period.label;
    option.selected = period.index === state.selectedPeriodIndex;
    periodSelect.appendChild(option);
  }

  periodSelect.disabled = state.periods.length === 0;
  updatePeriodSummary(getSelectedPeriod());
}

function getBudgetCardClass(item) {
  if (item.isSavingsLinked) {
    return "budget-item-savings";
  }

  if (item.sectionName === "Debt Payments") {
    return "budget-item-debt";
  }

  if (item.sectionName === "Savings Contributions") {
    return "budget-item-positive";
  }

  return "budget-item-default";
}

function getBaseBudget(item) {
  return Number(item.scheduledAmount || 0);
}

function getAdditionalAmount(item) {
  return Number(item.additional || 0);
}

function getEditableAdditionalAmount(item) {
  if (item.sectionName === "Savings" || item.sectionName === "Income") {
    return Number(item.effectiveAmount || 0) - Number(item.scheduledAmount || 0);
  }

  return getAdditionalAmount(item);
}

function getTotalBudget(item) {
  return getBaseBudget(item) + getAdditionalAmount(item);
}

function getMobileSpentAmount(item) {
  return Number(item.actualAmount || 0) + getAdditionalAmount(item);
}

function getMobileAvailableAmount(item) {
  if (item.sectionName === "Income") {
    return null;
  }

  if (item.sectionName === "Savings") {
    return Number(item.linkedSavingsAvailable ?? 0);
  }

  if (item.sectionName === "Debts" || item.sectionName === "Debt Payments") {
    return Number(item.effectiveAmount || 0);
  }

  if (item.isSavingsLinked) {
    return Number(item.linkedSavingsAvailable || 0);
  }

  const autoBudget = getBaseBudget(item);
  return autoBudget !== 0 ? autoBudget - getMobileSpentAmount(item) : 0;
}

function buildBudgetGroups(items) {
  const sections = new Map();

  for (const item of items) {
    const sectionName = item.sectionName || "Other";
    const groupName = item.groupName || "Items";
    if (!sections.has(sectionName)) {
      sections.set(sectionName, new Map());
    }

    const groups = sections.get(sectionName);
    if (!groups.has(groupName)) {
      groups.set(groupName, []);
    }

    groups.get(groupName).push(item);
  }

  return sections;
}

function getBudgetSectionDisplayName(sectionName) {
  return sectionName === "Savings" ? "Savings Contributions" : sectionName;
}

function getAdjustButtonClasses(item) {
  const isIncome = item.sectionName === "Income";
  return {
    subtract: isIncome ? "danger" : "success",
    add: isIncome ? "success" : "danger"
  };
}

function getPrimaryMetric(item) {
  if (item.sectionName === "Income") {
    return {
      label: "Amount",
      value: Number(item.effectiveAmount || 0)
    };
  }

  if (item.sectionName === "Savings") {
    return {
      label: "Contribution",
      value: Number(item.effectiveAmount || 0)
    };
  }

  if (item.sectionName === "Debts" || item.sectionName === "Debt Payments") {
    return {
      label: "Paid",
      value: Number(item.effectiveAmount || 0)
    };
  }

  return {
    label: "Spent",
    value: getMobileSpentAmount(item)
  };
}

function getOutgoingBreakdownValue(item) {
  if (item.sectionName === "Income") {
    return 0;
  }

  if (item.sectionName === "Savings") {
    return Number(item.effectiveAmount || 0) + getAdditionalAmount(item);
  }

  if (item.sectionName === "Debts" || item.sectionName === "Debt Payments") {
    return Number(item.effectiveAmount || 0);
  }

  return getMobileSpentAmount(item);
}

function renderSummaryBreakdown(items) {
  summaryBreakdownChart.innerHTML = "";

  if (!items || items.length === 0) {
    summaryBreakdownChart.innerHTML = `<div class="summary-breakdown-empty">No outgoing items were available for this period.</div>`;
    return;
  }

  const grouped = new Map();
  let savingsContributionTotal = 0;

  for (const item of items) {
    if (item.sectionName === "Income") {
      continue;
    }

    const value = getOutgoingBreakdownValue(item);
    if (value <= 0) {
      continue;
    }

    if (item.sectionName === "Savings") {
      savingsContributionTotal += value;
      continue;
    }

    grouped.set(item.label, (grouped.get(item.label) || 0) + value);
  }

  if (savingsContributionTotal > 0) {
    grouped.set("Savings Contributions", savingsContributionTotal);
  }

  const rows = Array.from(grouped.entries())
    .map(([label, value]) => ({ label, value }))
    .sort((a, b) => b.value - a.value);

  if (rows.length === 0) {
    summaryBreakdownChart.innerHTML = `<div class="summary-breakdown-empty">No outgoing items with non-zero values were available for this period.</div>`;
    return;
  }

  const maxValue = Math.max(...rows.map(row => row.value), 1);

  for (const row of rows) {
    const article = document.createElement("article");
    article.className = "summary-breakdown-row";
    const widthPercent = Math.max(4, (row.value / maxValue) * 100);
    article.innerHTML = `
      <div class="summary-breakdown-topline">
        <div class="summary-breakdown-label">${row.label}</div>
        <div class="summary-breakdown-value">${formatMoney(row.value)}</div>
      </div>
      <div class="summary-breakdown-track">
        <div class="summary-breakdown-fill" style="width:${widthPercent}%;"></div>
      </div>
    `;
    summaryBreakdownChart.appendChild(article);
  }
}

function sumAmount(items, selector) {
  return items.reduce((total, item) => total + Number(selector(item) || 0), 0);
}

function getGroupBudgetAmount(item) {
  if (item.sectionName === "Income") {
    return Number(item.effectiveAmount || 0);
  }

  if (item.sectionName === "Savings") {
    return Number(item.effectiveAmount || 0);
  }

  if (item.sectionName === "Debts" || item.sectionName === "Debt Payments") {
    return Number(item.effectiveAmount || 0);
  }

  if (item.isSavingsLinked) {
    return Number(item.linkedSavingsAvailable || 0);
  }

  return Number(item.effectiveAmount || 0);
}

function buildSectionSummary(items) {
  return {
    budgeted: sumAmount(items, item => getGroupBudgetAmount(item)),
    actual: sumAmount(items, item => getPrimaryMetric(item).value),
    difference: sumAmount(items, item => getGroupBudgetAmount(item) - getPrimaryMetric(item).value)
  };
}

function hasNegativeSavingsAvailability(item) {
  if (!item) {
    return false;
  }

  if (item.sectionName === "Savings") {
    return Number(item.linkedSavingsAvailable ?? 0) < 0;
  }

  if (!item.isSavingsLinked) {
    return false;
  }

  return Number(getMobileAvailableAmount(item) || 0) < 0;
}

function wireCollapsibleToggle(button, content) {
  button.addEventListener("click", () => {
    const expanded = button.getAttribute("aria-expanded") === "true";
    button.setAttribute("aria-expanded", expanded ? "false" : "true");
    content.hidden = expanded;
  });
}

function renderBudgetItems(items) {
  budgetItemsHost.innerHTML = "";

  if (!items || items.length === 0) {
    budgetItemsHost.innerHTML = `<div class="empty-state">No budget items were available for this period.</div>`;
    return;
  }

  const sections = buildBudgetGroups(items);
  for (const [sectionName, groups] of sections.entries()) {
    const section = document.createElement("section");
    section.className = "budget-section";
    const displaySectionName = getBudgetSectionDisplayName(sectionName);

    const sectionItems = Array.from(groups.values()).flat();
    const sectionSummary = buildSectionSummary(sectionItems);
    const sectionHasSavingsWarning = sectionItems.some(hasNegativeSavingsAvailability);

      const sectionHeader = document.createElement("button");
      sectionHeader.type = "button";
      sectionHeader.className = "budget-section-header budget-section-toggle";
      sectionHeader.setAttribute("aria-expanded", "false");
      sectionHeader.innerHTML = `
        <div class="budget-section-title-row section-title-row">
          <span class="budget-disclosure" aria-hidden="true"></span>
          <div class="budget-section-title-wrap">
            ${sectionHasSavingsWarning ? `<span class="budget-warning-indicator" title="One or more savings-linked items in this section have a negative available balance.">⚠</span>` : ""}
            <div class="budget-section-title">${displaySectionName}</div>
          </div>
          <div class="budget-section-count">${sectionItems.length} item${sectionItems.length === 1 ? "" : "s"}</div>
        </div>
      </div>
    `;
    section.appendChild(sectionHeader);

    const sectionContent = document.createElement("div");
    sectionContent.className = "budget-section-content";
    sectionContent.hidden = true;
    wireCollapsibleToggle(sectionHeader, sectionContent);

    for (const [groupName, groupItems] of groups.entries()) {
      const group = document.createElement("div");
      group.className = "budget-group";

      const groupSummary = buildSectionSummary(groupItems);
      const groupHasSavingsWarning = groupItems.some(hasNegativeSavingsAvailability);

      const groupTitle = document.createElement("button");
      groupTitle.type = "button";
      groupTitle.className = "budget-group-header budget-group-toggle";
      groupTitle.setAttribute("aria-expanded", "false");
      groupTitle.innerHTML = `
        <div class="budget-group-title-row">
          <span class="budget-disclosure small" aria-hidden="true"></span>
          <div class="budget-group-title-wrap">
            ${groupHasSavingsWarning ? `<span class="budget-warning-indicator" title="One or more savings-linked items in this group have a negative available balance.">⚠</span>` : ""}
            <div class="budget-group-title">${groupName}</div>
          </div>
        </div>
        <div class="budget-group-meta">
          <span>${groupItems.length} item${groupItems.length === 1 ? "" : "s"}</span>
          <span>${formatMoney(groupSummary.actual)}</span>
        </div>
      `;
      group.appendChild(groupTitle);

      const groupContent = document.createElement("div");
      groupContent.className = "budget-group-content";
      groupContent.hidden = true;
      wireCollapsibleToggle(groupTitle, groupContent);

      for (const item of groupItems) {
        const card = document.createElement("article");
        card.className = `budget-item-card ${getBudgetCardClass(item)}`;
        const buttonClasses = getAdjustButtonClasses(item);
        const primaryMetric = getPrimaryMetric(item);
        const availableAmount = getMobileAvailableAmount(item);
        const showAvailable = availableAmount !== null;
        const availableLabel = item.sectionName === "Debts" || item.sectionName === "Debt Payments"
          ? "Planned"
          : "Available";
        const showSavingsWarning = hasNegativeSavingsAvailability(item);

        card.innerHTML = `
          <div class="budget-item-layout">
            <div class="budget-item-left">
              <div class="budget-item-title-wrap">
                ${showSavingsWarning ? `<span class="budget-warning-indicator" title="This savings-linked item has a negative available balance.">⚠</span>` : ""}
                <div class="budget-item-title">${item.label}</div>
              </div>
              <div class="budget-adjust-row">
                <button class="mini-button compact ${buttonClasses.subtract}" type="button" data-adjust-button="subtract">-</button>
                <button class="mini-button compact ${buttonClasses.add}" type="button" data-adjust-button="add">+</button>
              </div>
            </div>
            <div class="budget-item-right">
              <div class="budget-inline-metric align-right">
                <span class="budget-inline-label">${primaryMetric.label}:</span>
                <span class="budget-inline-value">${formatMoney(primaryMetric.value)}</span>
              </div>
              ${showAvailable
                ? `
                  <div class="budget-inline-metric align-right">
                    <span class="budget-inline-label">${availableLabel}:</span>
                    <span class="budget-inline-value ${Number(availableAmount) < 0 ? "negative" : ""}">${formatMoney(availableAmount)}</span>
                  </div>
                `
                : `<div class="budget-inline-metric align-right placeholder"></div>`}
            </div>
          </div>
          <div class="budget-item-footer budget-item-footer-hidden">
          </div>
        `;

        const addButton = card.querySelector("[data-adjust-button='add']");
        const subtractButton = card.querySelector("[data-adjust-button='subtract']");
        addButton.addEventListener("click", () => adjustAdditional(item, 1));
        subtractButton.addEventListener("click", () => adjustAdditional(item, -1));

        groupContent.appendChild(card);
      }

      group.appendChild(groupContent);
      sectionContent.appendChild(group);
    }

    section.appendChild(sectionContent);
    budgetItemsHost.appendChild(section);
  }
}

async function adjustAdditional(item, direction) {
  const verb = direction > 0 ? "add to" : "subtract from";
  const input = window.prompt(`How much would you like to ${verb} ${item.label}?`, "0.00");
  if (input === null) {
    return;
  }

  const trimmed = input.trim();
  const normalized = trimmed.replace(/[$,]/g, "");
  const delta = Number(normalized);
  if (!Number.isFinite(delta)) {
    setStatus("Please enter a valid dollar amount.", true);
    return;
  }

  const currentValue = getEditableAdditionalAmount(item);
  const additional = currentValue + (Math.abs(delta) * direction);

  setStatus(`Saving additional for ${item.label}...`);

  const response = await fetchJson("/api/budget/additional", {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      periodIndex: state.selectedPeriodIndex ?? item.periodIndex ?? 0,
      sourceKey: item.sourceKey,
      itemLabel: item.label,
      additional
    })
  });

  state.suppressChangePromptUntil = Date.now() + 8000;
  await loadSummary();
  setStatus(response?.message || `Saved additional for ${item.label}.`);
}

async function loadBudgetItems(updateStatus = true) {
  if (state.selectedPeriodIndex === null) {
    renderBudgetItems([]);
    return;
  }

  if (updateStatus) {
    setStatus("Loading period budget...");
  }

  const budgetItems = await fetchJson(`/api/budget/periods/${state.selectedPeriodIndex}/items`);
  renderBudgetItems(budgetItems.items || []);
  renderSummaryBreakdown(budgetItems.items || []);

  if (updateStatus) {
    const count = (budgetItems.items || []).length;
    setStatus(`Loaded ${count} budget items for ${formatDate(budgetItems.periodStart)}.`);
  }
}

async function loadSummary() {
  const [summary, periods, accounts, savings, debts, transactions, changeToken] = await Promise.all([
    fetchJson("/api/overview/summary"),
    fetchJson("/api/periods"),
    fetchJson("/api/accounts"),
    fetchJson("/api/savings"),
    fetchJson("/api/debts"),
    fetchJson("/api/transactions?limit=40"),
    fetchJson("/api/change-token")
  ]);

  const totalAvailableSavings = (savings || []).reduce(
    (total, item) => total + Number(item.currentBalance || 0),
    0
  );

  document.getElementById("incomeTotal").textContent = formatMoney(summary.incomeTotal);
  document.getElementById("plannedOutflow").textContent = formatMoney(summary.plannedOutflow);
  document.getElementById("savingsTotal").textContent = formatMoney(totalAvailableSavings);
  document.getElementById("debtBalance").textContent = formatMoney(summary.currentDebtBalance);

  state.periods = periods || [];

  const desiredPeriodIndex = state.selectedPeriodIndex;
  const defaultCurrent = state.periods.find(period => period.isCurrent)?.index ?? 0;
  const selectedStillExists = state.periods.some(period => period.index === desiredPeriodIndex);
  state.selectedPeriodIndex = selectedStillExists ? desiredPeriodIndex : defaultCurrent;

  populatePeriodSelect();

  document.getElementById("accountCount").textContent = String(accounts.length);
  document.getElementById("savingsCount").textContent = String(savings.length);
  document.getElementById("debtsCount").textContent = String(debts.length);
  document.getElementById("transactionsCount").textContent = String(transactions.length);

  renderCards("accountsList", accounts, "account", item => {
    const belowZero = Number(item.currentBalance) < 0;
    const belowSafety = !belowZero && Number(item.currentBalance) < Number(item.safetyNet || 0);
    return {
      title: item.description || "Account",
      value: formatMoney(item.currentBalance),
      valueClass: belowZero ? "negative" : "",
      detail: item.accountType || "Account",
      meta: `Safety net: ${formatMoney(item.safetyNet || 0)}`,
      cardClass: belowZero ? "negative-card" : belowSafety ? "warning-card" : ""
    };
  });

  renderCards("savingsList", savings, "savings", item => ({
    title: item.description || "Savings",
    value: formatMoney(item.currentBalance),
    valueClass: Number(item.currentBalance) < 0 ? "negative" : "",
    detail: item.category || "Savings",
    meta: item.isHidden ? "Hidden" : "",
    cardClass: Number(item.currentBalance) < 0 ? "negative-card" : "positive-card"
  }));

  renderCards("debtsList", debts, "debt", item => ({
    title: item.description || "Debt",
    value: formatMoney(item.currentBalance),
    valueClass: Number(item.currentBalance) < 0 ? "negative" : "",
    detail: item.category || "Debt",
    meta: item.expectedPayoffDate ? `Expected payoff: ${formatDate(item.expectedPayoffDate)}` : "Expected payoff: Not projected",
    cardClass: "negative-card"
  }));

  const transactionsHost = document.getElementById("transactionsList");
  transactionsHost.innerHTML = "";
  for (const transaction of transactions) {
    const card = document.createElement("article");
    card.className = "transaction-card";
    const assignmentChips = (transaction.assignments || []).map(assignment => {
      const reviewClass = assignment.needsReview ? " review" : "";
      return `<span class="assignment-chip${reviewClass}">${assignment.itemLabel}: ${formatMoney(assignment.amount)}</span>`;
    }).join("");
    const sourceText = transaction.sourceName ? ` | ${transaction.sourceName}` : "";

    card.innerHTML = `
      <div class="transaction-topline">
        <div class="transaction-title">${transaction.description || "Transaction"}</div>
        <div class="transaction-amount ${Number(transaction.amount) < 0 ? "negative" : ""}">${formatMoney(transaction.amount)}</div>
      </div>
      <div class="transaction-meta">${formatDate(transaction.transactionDate)}${sourceText}</div>
      ${(transaction.assignments || []).length > 0
        ? `<div class="assignment-chip-row">${assignmentChips}</div>`
        : `<div class="transaction-meta">Unassigned</div>`}
    `;
    transactionsHost.appendChild(card);
  }

  await loadBudgetItems(false);
  state.lastKnownChangeToken = changeToken?.token || state.lastKnownChangeToken;
  setStatus(`Loaded ${accounts.length} accounts, ${savings.length} savings items, ${debts.length} debts, and ${(transactions || []).length} recent transactions.`);
}

async function pollForExternalChanges() {
  if (state.changePromptOpen || document.visibilityState === "hidden") {
    return;
  }

  try {
    const response = await fetchJson("/api/change-token");
    const token = response?.token || "";
    if (!token) {
      return;
    }

    if (!state.lastKnownChangeToken) {
      state.lastKnownChangeToken = token;
      return;
    }

    if (token === state.lastKnownChangeToken) {
      return;
    }

    if (Date.now() < state.suppressChangePromptUntil) {
      state.lastKnownChangeToken = token;
      return;
    }

    state.changePromptOpen = true;
    const shouldRefresh = window.confirm("Another Device has Made Changes to the Budget.\n\nRefresh now to load the latest values?");
    state.changePromptOpen = false;

    if (shouldRefresh) {
      await loadSummary();
    } else {
      state.lastKnownChangeToken = token;
      setStatus("A newer budget change is available. Refresh when you're ready.");
    }
  } catch {
    // Ignore transient polling failures so the page stays responsive.
  }
}

periodSelect.addEventListener("change", async event => {
  state.selectedPeriodIndex = Number(event.target.value);
  updatePeriodSummary(getSelectedPeriod());

  try {
    await loadBudgetItems(true);
  } catch (error) {
    setStatus(error.message || "Failed to load period budget.", true);
  }
});

for (const button of navButtons) {
  button.addEventListener("click", () => {
    setActivePage(button.dataset.navPage || "summary");
    window.scrollTo({ top: 0, behavior: "smooth" });
  });
}

refreshButton.addEventListener("click", async () => {
  try {
    setStatus("Refreshing mobile budget view...");
    await loadSummary();
  } catch (error) {
    setStatus(error.message || "Failed to refresh summary.", true);
  }
});

loadSummary().catch(error => {
  setStatus(error.message || "Failed to load summary.", true);
});

window.setInterval(() => {
  pollForExternalChanges();
}, CHANGE_TOKEN_POLL_INTERVAL_MS);
